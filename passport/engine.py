# -*- coding: utf-8 -*-
"""
Движок заполнения паспорта изделия.

Шаблон — обычный .docx, в котором переменные места помечены метками вида {{КЛЮЧ}}.
Метки ищутся во всём документе: основной текст, таблицы, надписи (text box),
колонтитулы. Метка может быть разбита Word'ом на несколько фрагментов (runs) —
это обрабатывается. Форматирование берётся у первого символа метки.

Таблица «Технические данные и характеристики» находится автоматически — это
первая таблица, идущая сразу за абзацем с таким заголовком (пустые абзацы между
ними допускаются). Строки таблицы перезаписываются, лишние удаляются,
недостающие добавляются копией последней строки (с её оформлением).

Изображение заменяется у рисунка, в «Замещающем тексте» (Alt Text) которого
стоит метка, например {{ИЗОБРАЖЕНИЕ_МАКЕТА}}. Новая картинка вписывается
в рамку старой с сохранением пропорций.
"""
import copy
import re
from datetime import date
from pathlib import Path

import docx
from docx.oxml.ns import qn

PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
TECH_TABLE_HEADING = "технические данные и характеристики"

W_P, W_R, W_T = qn("w:p"), qn("w:r"), qn("w:t")
W_TBL, W_TR, W_TC = qn("w:tbl"), qn("w:tr"), qn("w:tc")
W_BR, W_TAB = qn("w:br"), qn("w:tab")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

MONTHS_GEN = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
              "августа", "сентября", "октября", "ноября", "декабря"]


# ---------------------------------------------------------------- форматирование
def months_word(n: int) -> str:
    """1 месяц, 2 месяца, 5 месяцев, 21 месяц, 36 месяцев."""
    n = abs(int(n))
    if n % 10 == 1 and n % 100 != 11:
        return "месяц"
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return "месяца"
    return "месяцев"


def format_warranty(months) -> str:
    m = int(months)
    return f"{m} {months_word(m)}"


def format_sale_date(d: date) -> str:
    """«20» октября 2026 г."""
    return f"«{d.day:02d}» {MONTHS_GEN[d.month - 1]} {d.year} г."


def parse_date(text: str) -> date:
    text = text.strip()
    m = re.fullmatch(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})", text)
    if not m:
        raise ValueError("Дата должна быть в формате ДД.ММ.ГГГГ")
    d, mo, y = (int(x) for x in m.groups())
    if y < 100:
        y += 2000
    return date(y, mo, d)


# ---------------------------------------------------------------- обход частей
def _xml_parts(document):
    """Основной документ + колонтитулы + сноски — всё, где может быть текст."""
    seen = set()
    for part in document.part.package.iter_parts():
        ct = getattr(part, "content_type", "")
        if not ct.endswith("+xml"):
            continue
        if not any(k in ct for k in ("document.main", "header", "footer",
                                    "footnotes", "endnotes")):
            continue
        el = getattr(part, "element", None)
        if el is None:
            el = getattr(part, "_element", None)
        if el is not None and id(el) not in seen:
            seen.add(id(el))
            yield part, el


def _own_text_nodes(p):
    """w:t абзаца, без текста вложенных надписей (у них свои абзацы)."""
    out = []
    for t in p.iter(W_T):
        anc = t.getparent()
        while anc is not None and anc is not p and anc.tag != W_P:
            anc = anc.getparent()
        if anc is p:
            out.append(t)
    return out


def paragraph_text(p) -> str:
    return "".join(t.text or "" for t in _own_text_nodes(p))


def _set_t(t, text):
    t.text = text
    if text != text.strip() or "  " in text:
        t.set(XML_SPACE, "preserve")


def _expand_newlines(t):
    """Текст с \\n внутри w:t → w:t, w:br, w:t … внутри того же run."""
    if t.text is None or "\n" not in t.text:
        return
    parts = t.text.split("\n")
    _set_t(t, parts[0])
    anchor = t
    for piece in parts[1:]:
        br = t.makeelement(W_BR, {})
        anchor.addnext(br)
        nt = t.makeelement(W_T, {})
        _set_t(nt, piece)
        br.addnext(nt)
        anchor = nt


def replace_in_paragraph(p, old: str, new: str) -> int:
    """Заменяет все вхождения old на new в абзаце, даже если old разбит на runs.
    Возвращает число замен."""
    nodes = _own_text_nodes(p)
    if not nodes:
        return 0
    full = "".join(t.text or "" for t in nodes)
    if old not in full:
        return 0
    # карта: позиция символа -> (узел, смещение)
    spans, pos = [], 0
    for t in nodes:
        s = t.text or ""
        spans.append((pos, pos + len(s), t))
        pos += len(s)

    hits = [m.start() for m in re.finditer(re.escape(old), full)]
    for start in reversed(hits):
        end = start + len(old)
        first = True
        for a, b, t in spans:
            if b <= start or a >= end:
                continue
            s = t.text or ""
            lo, hi = max(start, a) - a, min(end, b) - a
            if first:
                _set_t(t, s[:lo] + new + s[hi:])
                first = False
            else:
                _set_t(t, s[:lo] + s[hi:])
        # пересчитать карту после изменения
        spans, pos = [], 0
        for t in nodes:
            s = t.text or ""
            spans.append((pos, pos + len(s), t))
            pos += len(s)
    for t in nodes:
        _expand_newlines(t)
    return len(hits)


def find_placeholders(document) -> list:
    """Список меток шаблона в порядке появления (текст + Alt Text рисунков)."""
    keys = []
    for _, root in _xml_parts(document):
        for p in root.iter(W_P):
            for m in PLACEHOLDER_RE.finditer(paragraph_text(p)):
                if m.group(1) not in keys:
                    keys.append(m.group(1))
        for el in root.iter(qn("wp:docPr")):
            for attr in ("descr", "title", "name"):
                for m in PLACEHOLDER_RE.finditer(el.get(attr) or ""):
                    if m.group(1) not in keys:
                        keys.append(m.group(1))
    return keys


def image_placeholders(document) -> list:
    keys = []
    for _, root in _xml_parts(document):
        for el in root.iter(qn("wp:docPr")):
            for attr in ("descr", "title", "name"):
                for m in PLACEHOLDER_RE.finditer(el.get(attr) or ""):
                    if m.group(1) not in keys:
                        keys.append(m.group(1))
    return keys


def fill_placeholders(document, values: dict) -> dict:
    """Подставляет значения. Возвращает {ключ: число замен}."""
    counts = {k: 0 for k in values}
    for _, root in _xml_parts(document):
        for p in list(root.iter(W_P)):
            txt = paragraph_text(p)
            if "{{" not in txt:
                continue
            m0 = PLACEHOLDER_RE.fullmatch(txt.strip())
            if m0 and m0.group(1) in values and "\n" in str(values[m0.group(1)]):
                # метка одна в абзаце, значение многострочное → каждая строка отдельным
                # абзацем с тем же оформлением (интервалы как в образце)
                lines = str(values[m0.group(1)]).split("\n")
                proto = copy.deepcopy(p)
                replace_in_paragraph(p, m0.group(0), lines[0])
                last = p
                for line in lines[1:]:
                    np_ = copy.deepcopy(proto)
                    replace_in_paragraph(np_, m0.group(0), line)
                    last.addnext(np_)
                    last = np_
                counts[m0.group(1)] += 1
                continue
            for m in list(PLACEHOLDER_RE.finditer(txt)):
                key = m.group(1)
                if key in values:
                    counts[key] += replace_in_paragraph(p, m.group(0), str(values[key]))
    return counts


# ---------------------------------------------------------------- таблица техданных
def _body_children(document):
    return list(document.element.body)


def find_tech_table(document, heading=TECH_TABLE_HEADING):
    armed, gap = False, 0
    for el in _body_children(document):
        if el.tag == W_P:
            txt = paragraph_text(el).strip().lower()
            if heading in txt:
                armed, gap = True, 0
            elif armed:
                if txt == "":
                    gap += 1
                    if gap > 3:
                        armed = False
                else:
                    armed = False
        elif el.tag == W_TBL and armed:
            return el
    return None


def _cell_text(tc) -> str:
    paras = [paragraph_text(p) for p in tc.findall(W_P)]
    return "\n".join(paras).strip()


def read_tech_table(document) -> list:
    tbl = find_tech_table(document)
    if tbl is None:
        return []
    rows = []
    for tr in tbl.findall(W_TR):
        tcs = tr.findall(W_TC)
        if not tcs:
            continue
        name = _cell_text(tcs[0])
        value = _cell_text(tcs[-1]) if len(tcs) > 1 else ""
        rows.append([name, value])
    return rows


def _set_cell_text(tc, text: str):
    if _cell_text(tc) == text.strip():
        return
    paras = tc.findall(W_P)
    if not paras:
        p = tc.makeelement(W_P, {})
        tc.append(p)
        paras = [p]
    p = paras[0]
    for extra in paras[1:]:
        tc.remove(extra)
    runs = [r for r in p if r.tag == W_R]
    if runs:
        r = runs[0]
        for other in runs[1:]:
            p.remove(other)
        for ch in list(r):
            if ch.tag in (W_T, W_TAB, W_BR, qn("w:cr")):
                r.remove(ch)
    else:
        r = p.makeelement(W_R, {})
        ppr = p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:rPr")) is not None:
            r.append(copy.deepcopy(ppr.find(qn("w:rPr"))))
        p.append(r)
    t = r.makeelement(W_T, {})
    _set_t(t, text)
    r.append(t)
    _expand_newlines(t)


def write_tech_table(document, rows: list) -> bool:
    tbl = find_tech_table(document)
    if tbl is None:
        return False
    trs = tbl.findall(W_TR)
    if not trs:
        return False
    proto = trs[-1]
    # добавить недостающие
    while len(trs) < len(rows):
        new_tr = copy.deepcopy(proto)
        trs[-1].addnext(new_tr)
        trs.append(new_tr)
    # удалить лишние (с конца)
    while len(trs) > max(len(rows), 1):
        tbl.remove(trs.pop())
    for tr, (name, value) in zip(trs, rows):
        tcs = tr.findall(W_TC)
        if not tcs:
            continue
        _set_cell_text(tcs[0], name)
        if len(tcs) > 1:
            _set_cell_text(tcs[-1], value)
    return True


# ---------------------------------------------------------------- изображение
EMU_PER_TWIP = 635


def _image_size(path):
    from PIL import Image, ImageOps
    with Image.open(path) as im:
        im = ImageOps.exif_transpose(im)
        return im.size


def _put_picture(document, para, path, box_cx, box_cy):
    """Вставляет картинку в абзац-рамку: вписывает в поле box_cx×box_cy с сохранением
    пропорций и центрирует по обеим осям (по горизонтали — выравнивание абзаца,
    по вертикали — равные интервалы до и после)."""
    holder = next(para.iter(qn("wp:inline")), None)
    if holder is None:
        holder = next(para.iter(qn("wp:anchor")))
    w, h = _image_size(path)
    scale = min(box_cx / w, box_cy / h)
    cx, cy = max(1, int(w * scale)), max(1, int(h * scale))
    rid, _ = document.part.get_or_add_image(path)
    next(holder.iter(qn("a:blip"))).set(qn("r:embed"), rid)
    holder.find(qn("wp:extent")).set("cx", str(cx))
    holder.find(qn("wp:extent")).set("cy", str(cy))
    for ext in holder.iter(qn("a:ext")):
        if ext.get("cx") is not None:
            ext.set("cx", str(cx))
            ext.set("cy", str(cy))
    for src in holder.iter(qn("a:srcRect")):          # обрезку старой картинки сбрасываем
        for a in list(src.attrib):
            del src.attrib[a]
    # pPr правим методами python-docx — они соблюдают порядок элементов схемы Word
    ppr = para.get_or_add_pPr()
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    ppr.get_or_add_jc().val = WD_ALIGN_PARAGRAPH.CENTER
    ppr._remove_ind()
    sp = ppr.get_or_add_spacing()
    for a in list(sp.attrib):
        del sp.attrib[a]
    pad = int((box_cy - cy) / 2 / EMU_PER_TWIP)
    sp.set(qn("w:before"), str(pad))
    sp.set(qn("w:after"), str(pad))
    sp.set(qn("w:line"), "240")
    sp.set(qn("w:lineRule"), "auto")


def _empty_line(document):
    """Пустая строка обычного текста (стиль «Обычный»)."""
    p = document.element.body.makeelement(qn("w:p"), {})
    return p


def replace_image(document, key: str, image_paths) -> bool:
    """Заменяет рисунок с меткой key в «Замещающем тексте». image_paths — один путь
    или список (сторона 1, сторона 2 …): рисунки встают один под другим, между ними
    одна пустая строка. Поле (рамка) = размер рисунка в шаблоне."""
    if isinstance(image_paths, (str, Path)):
        image_paths = [image_paths]
    paths = [str(p) for p in image_paths if p]
    if not paths:
        return False
    tag = "{{" + key + "}}"
    done = False
    for docpr in list(document.element.iter(qn("wp:docPr"))):
        if not any(tag in (docpr.get(a) or "") for a in ("descr", "title", "name")):
            continue
        holder = docpr.getparent()
        extent = holder.find(qn("wp:extent"))
        box_cx, box_cy = int(extent.get("cx")), int(extent.get("cy"))
        para = holder
        while para is not None and para.tag != W_P:
            para = para.getparent()
        proto = copy.deepcopy(para)
        _put_picture(document, para, paths[0], box_cx, box_cy)
        chain = [para]
        prev = para.getprevious()
        if prev is not None and prev.tag == W_P and paragraph_text(prev).strip():
            chain.insert(0, prev)           # заголовок «Приложение 1…» не отрывается от макета
        last = para
        max_id = max([int(e.get("id", 0)) for e in document.element.iter(qn("wp:docPr"))] + [0])
        for path in paths[1:]:
            gap = _empty_line(document)
            last.addnext(gap)
            new_p = copy.deepcopy(proto)
            gap.addnext(new_p)
            max_id += 1
            for e in list(new_p.iter(qn("wp:docPr"))) + list(new_p.iter(qn("pic:cNvPr"))):
                e.set("id", str(max_id))
            _put_picture(document, new_p, path, box_cx, box_cy)
            chain += [gap, new_p]
            last = new_p
        # «Не отрывать от следующего» — макеты не разрываются между страницами
        for el in chain[:-1]:
            el.get_or_add_pPr().keepNext_val = True
        done = True
    return done


def clear_image_tags(document, keys):
    for docpr in document.element.iter(qn("wp:docPr")):
        for a in ("descr", "title"):
            v = docpr.get(a)
            if v:
                for k in keys:
                    v = v.replace("{{" + k + "}}", "")
                docpr.set(a, v.strip())


# ---------------------------------------------------------------- всё вместе
# ---------------------------------------------------------------- макет: строка-замена
def _drawing_run(el):
    while el is not None and el.tag != W_R:
        el = el.getparent()
    return el


def maket_fallback(document, key, has_image):
    """Рядом с картинкой макета в шаблоне есть скрытая строка (её текст — в title рисунка).
    Есть изображение → строка удаляется; нет → картинка шаблона удаляется, строка видна."""
    tag = "{{" + key + "}}"
    for dp in list(document.element.iter(qn("wp:docPr"))):
        if tag not in (dp.get("descr") or "") or not dp.get("title"):
            continue
        text = dp.get("title")
        para = dp
        while para is not None and para.tag != W_P:
            para = para.getparent()
        hidden = [r for r in para.findall(W_R)
                  if r.find(qn("w:rPr") + "/" + qn("w:vanish")) is not None
                  and "".join(t.text or "" for t in r.findall(W_T)) == text]
        if has_image:
            for r in hidden:
                para.remove(r)
            dp.set("title", "")
        else:
            run = _drawing_run(dp)
            if run is not None and run.getparent() is not None:
                run.getparent().remove(run)
            for r in hidden:
                r.find(qn("w:rPr")).remove(r.find(qn("w:rPr") + "/" + qn("w:vanish")))


# ---------------------------------------------------------------- печать и подписи
STAMP_DEFAULTS = {
    "печать_диаметр_мм": 38.0,          # реальный диаметр оттиска
    "подпись_ширина_мм": 40.0,
    "поворот_град": 5.0,                # ± к горизонту
    "сдвиг_мм": 8.0,                    # максимальный сдвиг центра
    "сдвиг_подписи_по_вертикали_мм": 3.0,
}
STAMP_RE = re.compile(r"\[\[([^\[\]]+)\]\]")


def _ink_image(path, angle, tmpdir):
    """Обрезает по краю оттиска, поворачивает (прозрачный фон) → (png, w_px, h_px, w0, h0)."""
    from PIL import Image
    im = Image.open(path).convert("RGBA")
    a = im.getchannel("A")
    if a.getextrema()[0] == 255:                 # фон непрозрачный (белый) — делаем прозрачным
        import numpy as np
        arr = np.asarray(im).astype(float)
        ink = 1 - arr[..., :3].min(axis=2) / 255.0
        alpha = np.clip(ink * 1.6, 0, 1)
        rgb = np.where(alpha[..., None] > 0,
                       255 - (255 - arr[..., :3]) / np.maximum(alpha[..., None], 1e-3), 255)
        im = Image.fromarray(np.dstack([np.clip(rgb, 0, 255), alpha * 255]).astype("uint8"), "RGBA")
        a = im.getchannel("A")
    bbox = a.point(lambda v: 255 if v > 40 else 0).getbbox()
    if bbox:
        im = im.crop(bbox)
    w0, h0 = im.size
    rot = im.rotate(angle, resample=Image.BICUBIC, expand=True)
    out = Path(tmpdir) / f"ink_{abs(hash((str(path), angle)))}.png"
    rot.save(out)
    return out, rot.size[0], rot.size[1], w0, h0


def apply_stamps(document, enabled, library, settings=None, rng=None):
    """Места [[ПЕЧАТЬ]], [[ПОДПИСЬ_…]] в шаблоне. Выключено или нет файла — место удаляется.
    Каждый оттиск: свой поворот (±поворот_град) и сдвиг (до сдвиг_мм), поэтому в соседних
    документах печати и подписи не совпадают."""
    import math
    import random
    import tempfile
    cfg = dict(STAMP_DEFAULTS, **(settings or {}))
    rng = rng or random.SystemRandom()
    tmpdir = tempfile.mkdtemp(prefix="pasp_")
    placed = []
    for dp in list(document.element.iter(qn("wp:docPr"))):
        m = STAMP_RE.fullmatch((dp.get("descr") or "").strip())
        if not m:
            continue
        key = m.group(1)
        run = _drawing_run(dp)
        files = (library or {}).get(key) or []
        if not enabled or not files:
            if run is not None and run.getparent() is not None:
                run.getparent().remove(run)
            continue
        holder = dp.getparent()
        ext = holder.find(qn("wp:extent"))
        ph = holder.find(qn("wp:positionH") + "/" + qn("wp:posOffset"))
        pv = holder.find(qn("wp:positionV") + "/" + qn("wp:posOffset"))
        cx0, cy0 = int(ext.get("cx")), int(ext.get("cy"))
        center_x = int(ph.text) + cx0 / 2
        center_y = int(pv.text) + cy0 / 2
        angle = rng.uniform(-cfg["поворот_град"], cfg["поворот_град"])
        png, wpx, hpx, w0, h0 = _ink_image(rng.choice(files), angle, tmpdir)
        if key.startswith("ПЕЧАТЬ"):
            px_per_mm = (w0 + h0) / 2 / cfg["печать_диаметр_мм"]
            r = cfg["сдвиг_мм"] * math.sqrt(rng.random())
            t = rng.uniform(0, 2 * math.pi)
            dx, dy = r * math.cos(t), r * math.sin(t)
        else:
            px_per_mm = w0 / cfg["подпись_ширина_мм"]
            dx = rng.uniform(-cfg["сдвиг_мм"], cfg["сдвиг_мм"])
            v = min(cfg["сдвиг_мм"], cfg["сдвиг_подписи_по_вертикали_мм"])
            dy = rng.uniform(-v, v)
        cx, cy = int(wpx / px_per_mm * 36000), int(hpx / px_per_mm * 36000)
        rid, _ = document.part.get_or_add_image(str(png))
        next(holder.iter(qn("a:blip"))).set(qn("r:embed"), rid)
        ext.set("cx", str(cx))
        ext.set("cy", str(cy))
        for e in holder.iter(qn("a:ext")):
            if e.get("cx") is not None:
                e.set("cx", str(cx))
                e.set("cy", str(cy))
        ph.text = str(int(center_x - cx / 2 + dx * 36000))
        pv.text = str(int(center_y - cy / 2 + dy * 36000))
        dp.set("descr", key.replace("_", " ").capitalize())
        placed.append((key, round(angle, 1), round(dx, 1), round(dy, 1)))
    return placed


def stamp_places(document):
    return sorted({m.group(1) for dp in document.element.iter(qn("wp:docPr"))
                   for m in [STAMP_RE.fullmatch((dp.get("descr") or "").strip())] if m})


def generate(template_path, out_path, values: dict, table_rows=None, images=None,
             stamps=None):
    """Создаёт паспорт. stamps = {"включено": bool, "файлы": {ключ: [png…]}, "настройки": {…}}.
    Возвращает отчёт (dict)."""
    document = docx.Document(template_path)
    report = {"не найдены в шаблоне": [], "остались незаполненными": [],
              "таблица": None, "изображения": {}, "печати": []}
    counts = fill_placeholders(document, values)
    report["не найдены в шаблоне"] = [k for k, c in counts.items() if c == 0
                                      and k not in (images or {})]
    if table_rows is not None:
        report["таблица"] = write_tech_table(document, table_rows)
    img_keys = image_placeholders(document)
    for key in img_keys:
        path = (images or {}).get(key)
        has = bool(path) and (not isinstance(path, (list, tuple)) or any(path))
        maket_fallback(document, key, has)
        if has:
            report["изображения"][key] = replace_image(document, key, path)
    clear_image_tags(document, img_keys)
    st = stamps or {}
    report["печати"] = apply_stamps(document, st.get("включено", False), st.get("файлы"),
                                    st.get("настройки"))
    left = [k for k in find_placeholders(document)]
    report["остались незаполненными"] = left
    document.save(out_path)
    return report
