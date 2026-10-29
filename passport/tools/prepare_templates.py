# -*- coding: utf-8 -*-
"""Разовая подготовка шаблонов из исходных паспортов (.docx, сконвертированных из .doc):
расставляет метки {{...}} и нормализует XML для совместимости со всеми версиями Word.
Запуск: python prepare_templates.py dyn.docx stat.docx out_dir
"""
import json
import sys
from pathlib import Path

import docx
from docx.oxml.ns import qn
from docx.shared import Mm

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import engine  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def normalize(document):
    """jc=start/end и w:start/w:end (ISO-стиль LibreOffice) → left/right."""
    for part in document.part.package.iter_parts():
        el = getattr(part, "_element", None)
        if el is None:
            continue
        for jc in el.iter(W + "jc"):
            v = jc.get(W + "val")
            if v == "start":
                jc.set(W + "val", "left")
            elif v == "end":
                jc.set(W + "val", "right")
        for ind in el.iter(W + "ind"):
            for a, b in (("start", "left"), ("end", "right"),
                         ("startChars", "leftChars"), ("endChars", "rightChars")):
                if ind.get(W + a) is not None:
                    ind.set(W + b, ind.get(W + a))
                    del ind.attrib[W + a]
        # w:start/w:end как стороны рамок и полей ячеек (НЕ w:start в нумерации!)
        sides = {W + "pBdr", W + "tcBorders", W + "tblBorders", W + "tcMar",
                 W + "tblCellMar", W + "tblPrExBorders"}
        for tag, new_tag in (("start", "left"), ("end", "right")):
            for e in list(el.iter(W + tag)):
                if e.getparent() is not None and e.getparent().tag in sides:
                    e.tag = W + new_tag


def P(document, i):
    return document.paragraphs[i]._p


def must(n, what):
    if not n:
        raise RuntimeError("Не найдено: " + what)


def rep(p, old, new):
    must(engine.replace_in_paragraph(p, old, new), old)


PAGE_W, MARGIN_L, MARGIN_R = 210.0, 20.0025, 12.506
GOST = "GOST 2.304 type A"
KEEP_FONTS = {"RussianRail G Pro", "Symbol", "Wingdings", "Akrobat Bold", "OpenSymbol"}


def cover(document, first, last, axis_mm, width_mm):
    """Абзацы титула first..last → один абзац с меткой {{НАИМЕНОВАНИЕ_ТИТУЛ}}.
    При заполнении каждая строка титула станет отдельным абзацем с тем же оформлением,
    что в образце (кегль, интервалы). Строки центрируются по оси axis_mm от левого края
    листа — ось снята с образца; ширина поля width_mm."""
    p = P(document, first)
    runs = p.findall(qn("w:r"))
    keep = runs[0]
    for r in runs[1:]:
        p.remove(r)
    for ch in list(keep):
        if ch.tag != qn("w:rPr"):
            keep.remove(ch)
    t = keep.makeelement(qn("w:t"), {})
    t.text = "{{НАИМЕНОВАНИЕ_ТИТУЛ}}"
    keep.append(t)
    para = document.paragraphs[first]
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = para.paragraph_format
    pf.left_indent = Mm(axis_mm - width_mm / 2 - MARGIN_L)
    pf.right_indent = Mm((PAGE_W - MARGIN_R) - (axis_mm + width_mm / 2))
    pf.first_line_indent = Mm(0)
    for i in range(last, first, -1):
        el = P(document, i)
        el.getparent().remove(el)


def _style_fonts(document):
    """styleId → шрифт стиля (с учётом basedOn) до изменений."""
    st = document.styles.element
    styles = {s.get(qn("w:styleId")): s for s in st.findall(qn("w:style"))}
    def font(sid, depth=0):
        s = styles.get(sid)
        if s is None or depth > 10:
            return None
        rf = s.find(qn("w:rPr") + "/" + qn("w:rFonts"))
        if rf is not None and rf.get(qn("w:ascii")):
            return rf.get(qn("w:ascii")).split(";")[0]
        b = s.find(qn("w:basedOn"))
        return font(b.get(qn("w:val")), depth + 1) if b is not None else None
    return {sid: font(sid) for sid in styles}


def pin_cover_fonts(document):
    """Титул (всё до первого разрыва страницы) не должен «поехать», когда стили
    переводятся на ГОСТ: фиксируем явно шрифт стиля у знаков абзаца и у фрагментов
    без своего шрифта — высота строк и положение надписей останутся как в образце."""
    from lxml import etree
    fonts = _style_fonts(document)
    roots = [document.element.body]
    # колонтитулы первой страницы (титула) — тоже фиксируем, от них зависит отступ сверху
    sp = document.sections[0]._sectPr
    for ref in sp:
        if ref.tag in (qn("w:headerReference"), qn("w:footerReference")) and ref.get(qn("w:type")) == "first":
            roots.append(document.part.related_parts[ref.get(qn("r:id"))]._element)
    for root in roots:
        _pin(root, fonts, stop_at_page_break=root is document.element.body)


def _pin(root, fonts, stop_at_page_break):
    from lxml import etree
    for el in root:
        if el.tag != qn("w:p"):
            continue
        ps = el.find(qn("w:pPr") + "/" + qn("w:pStyle"))
        f = fonts.get(ps.get(qn("w:val")) if ps is not None else "Normal") or fonts.get("Normal")
        if f:
            ppr = el.get_or_add_pPr()
            rpr = ppr.find(qn("w:rPr"))
            if rpr is None:
                rpr = ppr.makeelement(qn("w:rPr"), {})
                ppr.append(rpr)
            targets = [rpr] + [r.get_or_add_rPr() for r in el.findall(qn("w:r"))]
            for t in targets:
                rf = t.find(qn("w:rFonts"))
                if rf is None:
                    rf = t.makeelement(qn("w:rFonts"), {})
                    t.insert(0, rf)
                for a in ("ascii", "hAnsi", "cs", "eastAsia"):
                    if not rf.get(qn("w:" + a)):
                        rf.set(qn("w:" + a), f)
        if stop_at_page_break and 'w:type="page"' in etree.tostring(el, encoding="unicode"):
            break


def fonts_to_gost(document):
    """Весь текст, кроме титула/обложек (RussianRail) и значков списков, — GOST 2.304 type A.
    Заодно убирает «хвосты» LibreOffice в именах шрифтов («GOST 2.304 type A;Calibri»)."""
    n = 0
    for part in document.part.package.iter_parts():
        el = getattr(part, "_element", None)
        if el is None:
            continue
        name = str(part.partname)
        normal = None
        if "styles" in name:
            normal = next((x for x in el.findall(qn("w:style")) if x.get(qn("w:styleId")) == "Normal"), None)
        for rf in el.iter(W + "rFonts"):
            if normal is not None and any(a is normal for a in rf.iterancestors()):
                continue    # «Обычный» не трогаем: от него зависит раскладка титула
            for a in list(rf.attrib):
                v = rf.get(a)
                if v and ";" in v:
                    rf.set(a, v.split(";")[0])
            if not any(k in name for k in ("document", "styles", "header", "footer", "numbering")):
                continue
            cur = rf.get(W + "ascii") or rf.get(W + "hAnsi") or ""
            if cur in KEEP_FONTS:
                continue
            for a in ("ascii", "hAnsi", "cs", "eastAsia"):
                rf.set(W + a, GOST)
            for a in ("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"):
                if rf.get(W + a) is not None:
                    del rf.attrib[W + a]
            n += 1
        for f in el.iter(W + "font"):
            v = f.get(W + "name")
            if v and ";" in v:
                f.set(W + "name", v.split(";")[0])
    # текст без явного шрифта (берёт «Обычный») — назначаем ГОСТ явно, кроме титула
    from lxml import etree
    roots = [(document.element.body, True)]
    sp = document.sections[0]._sectPr
    for ref in sp:
        if ref.tag in (qn("w:headerReference"), qn("w:footerReference")) and ref.get(qn("w:type")) != "first":
            roots.append((document.part.related_parts[ref.get(qn("r:id"))]._element, False))
    for root, has_cover in roots:
        in_cover = has_cover
        for el in root:
            if in_cover and el.tag == qn("w:sdt"):      # оглавление — уже не титул
                in_cover = False
            if in_cover:
                if 'w:type="page"' in etree.tostring(el, encoding="unicode"):
                    in_cover = False
                continue
            for r in el.iter(qn("w:r")):
                if not any((t.text or "").strip() for t in r.findall(qn("w:t"))):
                    continue
                rpr = r.get_or_add_rPr()
                rf = rpr.find(qn("w:rFonts"))
                if rf is None:
                    rf = rpr.makeelement(qn("w:rFonts"), {})
                    rpr.insert(0, rf)
                if (rf.get(W + "ascii") or "") in KEEP_FONTS:
                    continue
                for a in ("ascii", "hAnsi", "cs", "eastAsia"):
                    rf.set(W + a, GOST)
                n += 1
    return n


def _is_empty_par(el):
    from lxml import etree
    if el.tag != qn("w:p"):
        return False
    x = etree.tostring(el, encoding="unicode")
    return not engine.paragraph_text(el).strip() and "<w:drawing" not in x and \
        "<w:pict" not in x and 'w:type="page"' not in x and "<w:sectPr" not in x


def on_new_page(el):
    """Удаляет пустые абзацы перед элементом и ставит «С новой страницы»."""
    prev = el.getprevious()
    while prev is not None and _is_empty_par(prev):
        nxt = prev.getprevious()
        prev.getparent().remove(prev)
        prev = nxt
    ppr = el.get_or_add_pPr()
    ppr.pageBreakBefore_val = True
    sp = ppr.find(qn("w:spacing"))
    if sp is not None and sp.get(qn("w:before")):
        sp.set(qn("w:before"), "0")


def sheet_on_new_page(document, heading):
    el = next(p._p for p in document.paragraphs if p.text.strip().startswith(heading))
    on_new_page(el)


def back_cover_on_new_page(document):
    """Задняя обложка (последний абзац с плавающим рисунком) — всегда последней
    страницей. Помечается «Задняя обложка» в замещающем тексте — по этой метке
    утилита вставляет пустые страницы перед ней."""
    body = [e for e in document.element.body if e.tag == qn("w:p")]
    el = next(e for e in reversed(body) if e.find(".//" + qn("wp:anchor")) is not None)
    for dp in el.iter(qn("wp:docPr")):
        dp.set("descr", "Задняя обложка")
    on_new_page(el)


def keep_table_on_page(document, heading):
    """Заголовок с новой страницы + вся следующая за ним таблица на одной странице:
    строки не разрываются, каждая «не отрывается от следующей»."""
    head = next(p._p for p in document.paragraphs if p.text.strip().startswith(heading))
    on_new_page(head)
    head.get_or_add_pPr().keepNext_val = True
    el = head.getnext()
    while el is not None and el.tag != qn("w:tbl"):
        if el.tag == qn("w:p"):
            el.get_or_add_pPr().keepNext_val = True
        el = el.getnext()
    rows = el.findall(qn("w:tr"))
    for k, tr in enumerate(rows):
        trpr = tr.find(qn("w:trPr"))
        if trpr is None:
            trpr = tr.makeelement(qn("w:trPr"), {})
            tr.insert(0 if tr.find(qn("w:tblPrEx")) is None else 1, trpr)
        if trpr.find(qn("w:cantSplit")) is None:
            trpr.insert(0, trpr.makeelement(qn("w:cantSplit"), {}))
        if k < len(rows) - 1:
            for p in tr.iter(qn("w:p")):
                p.get_or_add_pPr().keepNext_val = True


def talon(table, label_to_value):
    for row in table.rows:
        label = row.cells[0].text.strip()
        for key, (new_label, value) in label_to_value.items():
            if label.startswith(key):
                if new_label:
                    engine._set_cell_text(row.cells[0]._tc, new_label)
                engine._set_cell_text(row.cells[-1]._tc, value)


def prepare_dynamic(src, out):
    d = docx.Document(src)
    normalize(d)
    # порядок: сначала то, что ниже по документу (индексы абзацев не сдвигаются)
    rep(P(d, 202), "«____» ___________202   г.", "{{ДАТА_ПРОДАЖИ}}")
    rep(P(d, 196), "Гарантийный талон № ", "Гарантийный талон № {{СЕРИЙНЫЙ_НОМЕР}}")
    rep(P(d, 171), "Динамическое информационное табло. Одностороннее. Уличное. 1450x970 мм",
        "{{НАИМЕНОВАНИЕ}}")
    rep(P(d, 171), "ASD-TAB P2,5 1450x970 ID", "{{АРТИКУЛ}}")
    rep(P(d, 147), "____ месяцев", "{{ГАРАНТИЯ}}")
    talon(d.tables[3], {
        "Название оборудования": (None, "{{НАИМЕНОВАНИЕ}}"),
        "Артикул": (None, "{{АРТИКУЛ}}"),
        "Серийный номер": (None, "№ {{СЕРИЙНЫЙ_НОМЕР}}"),
        "Срок гарантийной": (None, "{{ГАРАНТИЯ}}"),
    })
    keep_table_on_page(d, "Лист тестирования экрана")
    sheet_on_new_page(d, "Лист выдачи изделия")
    sheet_on_new_page(d, "Гарантийный талон")
    back_cover_on_new_page(d)
    # титул не трогаем: «Обычный» стиль и знаки абзацев титула остаются как в образце
    fonts_to_gost(d)
    cover(d, 9, 12, axis_mm=96.4, width_mm=124)
    d.save(out)
    return {
        "тип": "Динамика МЦД",
        "значения": {
            "НАИМЕНОВАНИЕ": "Динамическое информационное табло. Одностороннее. Уличное.",
            "НАИМЕНОВАНИЕ_ТИТУЛ": "Динамическое информационное\nтабло. Одностороннее.\n"
                                  "Уличное.\nASD-TAB P2,5 1450x970 ID",
            "АРТИКУЛ": "ASD-TAB P2,5 1450x970 ID",
            "ГАРАНТИЯ": "36",
            "СЕРИЙНЫЙ_НОМЕР": "",
        },
        "последний_серийный": "0240925001",
    }


def prepare_static(src, out):
    d = docx.Document(src)
    normalize(d)
    rep(P(d, 125), "« 26 »_Декабря__2025г.", "{{ДАТА_ПРОДАЖИ}}")
    rep(P(d, 96), "Статический навигационный указатель. Световой. Односторонний. 2000х400",
        "{{НАИМЕНОВАНИЕ}}")
    rep(P(d, 96), "ASD-L 2000х400", "{{АРТИКУЛ}}")
    rep(P(d, 72), "Приложение 1 ", "Приложение 1. Макет № {{НОМЕР_МАКЕТА}}")
    # картинка макета в Приложении 1 — метка в «Замещающем тексте»
    img = next(P(d, 73).iter(qn("wp:docPr")))
    img.set("descr", "{{ИЗОБРАЖЕНИЕ_МАКЕТА}}")
    rep(P(d, 62), "36 месяцев со дня ввода", "{{ГАРАНТИЯ}} со дня ввода")
    talon(d.tables[1], {
        "Название изделия": (None, "{{НАИМЕНОВАНИЕ}}"),
        "Артикул": (None, "{{АРТИКУЛ}}"),
        "Срок гарантийной": (None, "{{ГАРАНТИЯ}}"),
        "Серийный номер": ("Номер макета:", "{{НОМЕР_МАКЕТА}}"),
    })
    sheet_on_new_page(d, "Приложение 1")
    sheet_on_new_page(d, "Лист выдачи изделия")
    sheet_on_new_page(d, "Гарантийный талон")
    back_cover_on_new_page(d)
    # титул не трогаем: «Обычный» стиль и знаки абзацев титула остаются как в образце
    fonts_to_gost(d)
    cover(d, 9, 12, axis_mm=92.8, width_mm=150)
    d.save(out)
    return {
        "тип": "Статика МЦД",
        "значения": {
            "НАИМЕНОВАНИЕ": "Статический навигационный указатель. Световой. Односторонний.",
            "НАИМЕНОВАНИЕ_ТИТУЛ": "Статический навигационный\nуказатель. Световой.\n"
                                  "Односторонний.\nASD-L 2000х400",
            "АРТИКУЛ": "ASD-L 2000х400",
            "ГАРАНТИЯ": "36",
            "НОМЕР_МАКЕТА": "",
        },
    }


# ======================================================================= круглый указатель
MAKET_FALLBACK = "Нет актуальной информации. Ожидается"


def add_maket_fallback(par_el):
    """Скрытая строка «Нет актуальной информации. Ожидается» рядом с картинкой макета:
    если макет не выбран, утилита покажет её вместо картинки шаблона."""
    r = par_el.makeelement(qn("w:r"), {})
    rpr = r.get_or_add_rPr()
    van = rpr.makeelement(qn("w:vanish"), {})
    rpr.append(van)
    t = r.makeelement(qn("w:t"), {})
    t.text = MAKET_FALLBACK
    r.append(t)
    par_el.append(r)
    for dp in par_el.iter(qn("wp:docPr")):
        dp.set("title", MAKET_FALLBACK)


def prepare_round(src, out, maket_h_mm=80):
    d = docx.Document(src)
    normalize(d)
    rep(P(d, 133), "« 26 »_Декабря__2025г.", "{{ДАТА_ПРОДАЖИ}}")
    rep(P(d, 104), "ASD-L D=1200", "{{АРТИКУЛ}}")
    rep(P(d, 103), "Статический навигационный указатель. Световой. Односторонний. D=1200",
        "{{НАИМЕНОВАНИЕ}}")
    # «Нет актуальной информации» → поле макета (картинка) + скрытая строка-замена
    par = d.paragraphs[77]
    for r in list(par._p.findall(qn("w:r"))):
        par._p.remove(r)
    ph = placeholder_png("maket", 177.4, maket_h_mm)
    par.add_run().add_picture(str(ph), width=Mm(177.4), height=Mm(maket_h_mm))
    for dp in par._p.iter(qn("wp:docPr")):
        dp.set("descr", "{{ИЗОБРАЖЕНИЕ_МАКЕТА}}")
    add_maket_fallback(par._p)
    rep(P(d, 75), "Приложение 1 ", "Приложение 1. Макет № {{НОМЕР_МАКЕТА}}")
    rep(P(d, 65), "36 месяцев со дня ввода", "{{ГАРАНТИЯ}} со дня ввода")
    t = d.tables[1]
    talon(t, {
        "Название изделия": (None, "{{НАИМЕНОВАНИЕ}}"),
        "Артикул": (None, "{{АРТИКУЛ}}"),
        "Срок гарантийной": (None, "{{ГАРАНТИЯ}}"),
    })
    import copy as _copy
    new_tr = _copy.deepcopy(t.rows[-1]._tr)          # строка «Номер макета» как в статике
    t.rows[-1]._tr.addnext(new_tr)
    tcs = new_tr.findall(qn("w:tc"))
    engine._set_cell_text(tcs[0], "Номер макета:")
    engine._set_cell_text(tcs[-1], "{{НОМЕР_МАКЕТА}}")
    sheet_on_new_page(d, "Приложение 1")
    sheet_on_new_page(d, "Лист выдачи изделия")
    sheet_on_new_page(d, "Гарантийный талон")
    back_cover_on_new_page(d)
    fonts_to_gost(d)
    cover(d, 9, 12, axis_mm=92.8, width_mm=150)
    d.save(out)
    return {
        "тип": "Статика МЦД круглый",
        "значения": {
            "НАИМЕНОВАНИЕ": "Статический навигационный указатель. Световой. Односторонний.",
            "НАИМЕНОВАНИЕ_ТИТУЛ": "Статический навигационный\nуказатель. Световой.\n"
                                  "Односторонний.\nASD-L D=1200",
            "АРТИКУЛ": "ASD-L D=1200",
            "ГАРАНТИЯ": "36",
            "НОМЕР_МАКЕТА": "",
        },
    }


# ======================================================================= печать и подписи
TMP = Path("/tmp/claude-0/prep")


def placeholder_png(kind, w_mm, h_mm):
    """Картинка-заглушка для шаблона: видно в Word, где будет печать/подпись/макет."""
    from PIL import Image, ImageDraw, ImageFont
    TMP.mkdir(parents=True, exist_ok=True)
    k = 12  # px на мм
    W, H = int(w_mm * k), int(h_mm * k)
    im = Image.new("RGBA", (W, H), (0, 0, 0, 0) if kind != "maket" else (245, 247, 250, 255))
    d = ImageDraw.Draw(im)
    col = (120, 140, 190, 255)
    try:
        f = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", int(4 * k))
    except Exception:
        f = ImageFont.load_default()
    if kind == "stamp":
        for a in range(0, 360, 12):
            d.arc([6, 6, W - 6, H - 6], a, a + 7, fill=col, width=5)
        d.text((W / 2, H / 2), "М.П.", fill=col, font=f, anchor="mm")
    elif kind == "sign":
        for x in range(0, W, 30):
            d.line([(x, H - 4), (x + 16, H - 4)], fill=col, width=4)
        d.text((W / 2, H / 2), "подпись", fill=col, font=f, anchor="mm")
    elif kind == "date":
        for x in range(0, W, 24):
            d.line([(x, H - 3), (x + 12, H - 3)], fill=col, width=3)
        d.text((W / 2, H / 2), "дата", fill=col, font=f, anchor="mm")
    else:
        d.rectangle([2, 2, W - 3, H - 3], outline=col, width=4)
        d.text((W / 2, H / 2), "Макет", fill=col, font=f, anchor="mm")
    path = TMP / f"ph_{kind}_{w_mm:.0f}x{h_mm:.0f}.png"
    im.save(path)
    return path


ANCHOR_XML = (
    '<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{z}" behindDoc="0" '
    'locked="0" layoutInCell="1" allowOverlap="1">'
    '<wp:simplePos x="0" y="0"/>'
    '<wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH>'
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>{y}</wp:posOffset></wp:positionV>'
    '<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
    '<wp:docPr id="{id}" name="{name}" descr="{descr}"/><wp:cNvGraphicFramePr/>'
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="{id}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
    '<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
    '</wp:anchor>')
EMU = 36000  # EMU в мм


def add_anchor(document, par_el, png, w_mm, h_mm, x_mm, y_mm, descr, name):
    """Плавающий рисунок «перед текстом»: x — от левого края листа, y — от верха абзаца."""
    from lxml import etree
    rid, _ = document.part.get_or_add_image(str(png))
    ids = [int(e.get("id", 0)) for e in document.element.iter(qn("wp:docPr"))]
    nid = max(ids + [0]) + 1
    xml = ANCHOR_XML.format(z=251660000 + nid, x=int(x_mm * EMU), y=int(y_mm * EMU),
                            cx=int(w_mm * EMU), cy=int(h_mm * EMU), id=nid, name=name,
                            descr=descr, rid=rid)
    anchor = etree.fromstring(xml)
    r = par_el.makeelement(qn("w:r"), {})
    drawing = r.makeelement(qn("w:drawing"), {})
    drawing.append(anchor)
    r.append(drawing)
    par_el.insert(1 if par_el.find(qn("w:pPr")) is not None else 0, r)
    return r


def render_pdf(path):
    import subprocess
    out = TMP / "render"
    out.mkdir(parents=True, exist_ok=True)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(out), str(path)],
                   capture_output=True)
    return out / (Path(path).stem + ".pdf")


def words_bbox(pdf):
    """[(страница, x0, y0, x1, y1, слово)] в мм."""
    import re
    import subprocess
    html = TMP / "bbox.html"
    subprocess.run(["pdftotext", "-bbox", str(pdf), str(html)], capture_output=True)
    x = html.read_text("utf-8")
    out, page = [], -1
    for m in re.finditer(r'<page |<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', x):
        if m.group(0).startswith("<page"):
            page += 1
            continue
        k = 25.4 / 72
        out.append((page, float(m.group(1)) * k, float(m.group(2)) * k, float(m.group(3)) * k,
                    float(m.group(4)) * k, m.group(5)))
    return out


def find_red(pdf, page, dpi=150):
    """Верх-лево красного маркера на странице (мм)."""
    import subprocess
    from PIL import Image
    import numpy as np
    base = TMP / "cal"
    subprocess.run(["pdftoppm", "-r", str(dpi), "-png", "-f", str(page + 1), "-l", str(page + 1),
                    str(pdf), str(base)], capture_output=True)
    png = sorted(TMP.glob("cal-*.png"))[-1]
    a = np.asarray(Image.open(png).convert("RGB")).astype(int)
    png.unlink()
    m = (a[..., 0] > 200) & (a[..., 1] < 60) & (a[..., 2] < 60)
    ys, xs = np.where(m)
    if not len(ys):
        return None
    k = 25.4 / dpi
    return xs.min() * k, ys.min() * k


def stamp_places(path):
    """Ставит в шаблон места для печати и подписей:
       • лист выдачи — печать поверх строк «подпись» и «дата», подпись ответственного
         за качество на строке «подпись»;
       • гарантийный талон — печать и подпись продавца по центру линии
         «Фамилия и подпись продавца».
    Координаты снимаются с отрисованного шаблона (калибровка красным маркером)."""
    from lxml import etree
    path = Path(path)
    d = docx.Document(path)
    paras = d.paragraphs
    # абзацы-якоря
    q = next(i for i, p in enumerate(paras) if "подтверждаю" in p.text)
    lines = [i for i in range(q + 1, len(paras)) if set(paras[i].text.strip()) == {"_"}][:2]
    sel = next(i for i, p in enumerate(paras) if p.text.strip().startswith("Фамилия и подпись продавца"))
    anchors = {"sheet": paras[lines[0]]._p, "talon": paras[sel]._p}
    # 1) калибровка: где верх абзаца-якоря на странице
    TMP.mkdir(parents=True, exist_ok=True)
    red = TMP / "red.png"
    from PIL import Image
    Image.new("RGB", (20, 20), (255, 0, 0)).save(red)
    marks = {k: add_anchor(d, el, red, 1.5, 1.5, 5, 0, "cal", "cal") for k, el in anchors.items()}
    tmp_docx = TMP / ("cal_" + path.name)
    d.save(tmp_docx)
    pdf = render_pdf(tmp_docx)
    words = words_bbox(pdf)
    sheet_page = next(w[0] for w in words if w[5] == "дата")       # подпись под строкой «дата»
    talon_page = max(w[0] for w in words if w[5] == "продавца")
    origin = {"sheet": find_red(pdf, sheet_page), "talon": find_red(pdf, talon_page)}
    for r in marks.values():
        r.getparent().remove(r)
    # 2) целевые точки по тексту
    sw = [w for w in words if w[0] == sheet_page and set(w[5]) == {"_"}]
    sw = sorted(sw, key=lambda w: w[2])[-2:]            # две нижние линии: подпись, дата
    sx = sum((w[1] + w[3]) / 2 for w in sw) / 2
    s_line1 = sw[0][4]
    sy = (sw[0][4] + sw[1][4]) / 2
    tw = [w for w in words if w[0] == talon_page and set(w[5]) == {"_"}]
    prod = next(w for w in words if w[0] == talon_page and w[5] == "продавца")
    tl = min(tw, key=lambda w: abs(w[4] - prod[4]) + abs(w[1] - prod[3]))
    tx, ty = (tl[1] + tl[3]) / 2, tl[4]
    # 3) места: печать Ø38 (поле 38×38), подпись 40×34 мм
    S, SW, SH = 38.0, 40.0, 34.0
    stamp_ph = placeholder_png("stamp", S, S)
    sign_ph = placeholder_png("sign", SW, SH)
    oy_s, oy_t = origin["sheet"][1], origin["talon"][1]
    add_anchor(d, anchors["sheet"], stamp_ph, S, S, sx - S / 2, sy - S / 2 - oy_s,
               "[[ПЕЧАТЬ]]", "Печать")
    add_anchor(d, anchors["sheet"], sign_ph, SW, SH, sx - SW / 2, s_line1 - SH * 0.62 - oy_s,
               "[[ПОДПИСЬ_ДУДИН]]", "Подпись Дудин")
    add_anchor(d, anchors["talon"], stamp_ph, S, S, tx - S / 2, ty - S / 2 - oy_t,
               "[[ПЕЧАТЬ]]", "Печать")
    add_anchor(d, anchors["talon"], sign_ph, SW, SH, tx - SW / 2, ty - SH * 0.62 - oy_t,
               "[[ПОДПИСЬ_ЗАЛИВАКИН]]", "Подпись Заливакин")
    d.save(path)
    return {"лист": (round(sx, 1), round(sy, 1), sheet_page + 1), "талон": (round(tx, 1), round(ty, 1), talon_page + 1)}


DATE_W, DATE_H = 32.0, 8.0        # место рукописной даты, мм (низ места — линия письма)


def date_places(path):
    """Места для рукописной даты подписания (опция «Печать, подпись и дата»):
       • лист выдачи — на строке «дата» под «Качество изделия подтверждаю … Дудин Е.В.»;
       • гарантийный талон — на месте даты продажи (под «Дата продажи:»), по правому краю.
    Повторный запуск заменяет прежние места. Калибровка красным маркером, как у печати."""
    path = Path(path)
    d = docx.Document(path)
    for dp in list(d.element.iter(qn("wp:docPr"))):           # прежние места даты — убрать
        if (dp.get("descr") or "").startswith("[[ДАТА_"):
            r = dp.getparent().getparent().getparent()
            r.getparent().remove(r)
    paras = d.paragraphs
    q = next(i for i, p in enumerate(paras) if "подтверждаю" in p.text)
    lines = [i for i in range(q + 1, len(paras)) if set(paras[i].text.strip()) == {"_"}][:2]
    sale = next(i for i, p in enumerate(paras) if "{{ДАТА_ПРОДАЖИ}}" in p.text)
    anchors = {"sheet": paras[lines[-1]]._p, "talon": paras[sale]._p}
    TMP.mkdir(parents=True, exist_ok=True)
    red = TMP / "red.png"
    from PIL import Image
    Image.new("RGB", (20, 20), (255, 0, 0)).save(red)
    marks = {k: add_anchor(d, el, red, 1.5, 1.5, 5, 0, "cal", "cal") for k, el in anchors.items()}
    tmp_docx = TMP / ("cald_" + path.name)
    d.save(tmp_docx)
    pdf = render_pdf(tmp_docx)
    words = words_bbox(pdf)
    sheet_page = next(w[0] for w in words if w[5] == "дата")
    talon_page = max(w[0] for w in words if w[5] == "продавца")
    origin = {"sheet": find_red(pdf, sheet_page), "talon": find_red(pdf, talon_page)}
    for r in marks.values():
        r.getparent().remove(r)
    # строка «дата»: самая нижняя линия из подчёркиваний на листе выдачи
    sw = sorted([w for w in words if w[0] == sheet_page and set(w[5]) == {"_"}], key=lambda w: w[2])
    line = sw[-1]
    sx, sy = (line[1] + line[3]) / 2, line[4] - 0.9          # линия письма чуть выше подчёркивания
    # талон: метка {{ДАТА_ПРОДАЖИ}} — правый край и низ строки
    tw = next(w for w in words if w[0] == talon_page and "ДАТА_ПРОДАЖИ" in w[5])
    tx, ty = tw[3] - DATE_W / 2, tw[4]
    ph = placeholder_png("date", DATE_W, DATE_H)
    add_anchor(d, anchors["sheet"], ph, DATE_W, DATE_H, sx - DATE_W / 2,
               sy - DATE_H - origin["sheet"][1], "[[ДАТА_ДУДИН]]", "Дата Дудин")
    add_anchor(d, anchors["talon"], ph, DATE_W, DATE_H, tx - DATE_W / 2,
               ty - DATE_H - origin["talon"][1], "[[ДАТА_ЗАЛИВАКИН]]", "Дата Заливакин")
    d.save(path)
    return {"лист": (round(sx, 1), round(sy, 1), sheet_page + 1),
            "талон": (round(tx, 1), round(ty, 1), talon_page + 1)}


def replace_people(document, pairs):
    n = 0
    for _, root in engine._xml_parts(document):
        for p in list(root.iter(qn("w:p"))):
            for old, new in pairs:
                n += engine.replace_in_paragraph(p, old, new)
    return n


def issue_sheet_edits(document):
    """Лист выдачи: убрать строку «Сборщики: …» вместе с её линией для подписи;
    «Ответственный за качество:» → «Специалист ОТК:»; «Качество экрана» → «Качество изделия»."""
    body = list(document.element.body)
    texts = [engine.paragraph_text(e).strip() if e.tag == qn("w:p") else None for e in body]
    i = next(k for k, t in enumerate(texts) if t and t.startswith("Сборщики"))
    j = next(k for k in range(i, len(body)) if texts[k] and texts[k].startswith("Ответственный за качество"))
    removed = 0
    for e in body[i:j]:          # «Сборщики», пустые строки и линия подписи под ними
        assert texts[body.index(e)] is not None
        e.getparent().remove(e)
        removed += 1
    n = 0
    for p in document.element.body.iter(qn("w:p")):
        n += engine.replace_in_paragraph(p, "Ответственный за качество:", "Специалист ОТК:")
        n += engine.replace_in_paragraph(p, "Качество экрана подтверждаю", "Качество изделия подтверждаю")
    return removed, n


COVER_MARGIN_MM = 4.0


def fit_cover_zone(path, y_from=150, y_to=222):
    """Титул — только на тёмно-синем поле. Отрисовывает обложку без текста, находит
    по картинке левую линию и «квадратики» справа в полосе строк наименования и
    ограничивает абзац {{НАИМЕНОВАНИЕ_ТИТУЛ}} этой зоной (с отступом COVER_MARGIN_MM)."""
    import numpy as np
    import subprocess
    from lxml import etree
    from PIL import Image
    path = Path(path)
    d = docx.Document(path)
    blank = docx.Document(path)
    body = list(blank.element.body)
    stop = next(i for i, e in enumerate(body) if e.tag == qn("w:sdt") or
                'w:type="page"' in etree.tostring(e, encoding="unicode"))
    for el in body[:stop]:
        for t in el.iter(qn("w:t")):
            t.text = ""
    tmp = TMP / ("blank_" + path.name)
    blank.save(tmp)
    pdf = render_pdf(tmp)
    subprocess.run(["pdftoppm", "-r", "100", "-png", "-f", "1", "-l", "1", str(pdf), str(TMP / "zone")],
                   capture_output=True)
    png = sorted(TMP.glob("zone-*.png"))[-1]
    im = np.asarray(Image.open(png).convert("RGB")).astype(int)
    png.unlink()
    H, W, _ = im.shape
    k = 25.4 / 100
    bg = np.median(im[int(H * 0.5):int(H * 0.55), int(W * 0.15):int(W * 0.3)].reshape(-1, 3), axis=0)
    mask = np.abs(im - bg).sum(axis=2) > 60
    band = mask[int(y_from / k):int(y_to / k)]
    cols = np.where(band.any(axis=0))[0] * k
    L = max(c for c in cols if c < 0.25 * W * k) + COVER_MARGIN_MM
    R = min(c for c in cols if c > 0.40 * W * k) - COVER_MARGIN_MM
    para = next(p for p in d.paragraphs if "{{НАИМЕНОВАНИЕ_ТИТУЛ}}" in p.text)
    pf = para.paragraph_format
    pf.left_indent = Mm(L - MARGIN_L)
    pf.right_indent = Mm((PAGE_W - MARGIN_R) - R)
    pf.first_line_indent = Mm(0)
    d.save(path)
    return L, R


def refresh_toc(path):
    """Номера страниц в «Содержании» шаблона — по его фактической вёрстке."""
    import book
    d = docx.Document(path)
    pdf = render_pdf(path)
    n = book.update_toc(d, book.pdf_pages_text(pdf), book.page_offset(d))
    d.save(path)
    return n


def post_process(path):
    d = docx.Document(path)
    n = replace_people(d, [("Бокланов И.А.", "Дудин Е.В."), ("Бокланов И. А.", "Дудин Е.В.")])
    print("   лист выдачи: удалено абзацев", *issue_sheet_edits(d))
    # строка-замена для макета (если в шаблоне есть картинка макета без неё)
    for dp in list(d.element.iter(qn("wp:docPr"))):
        if dp.get("descr") == "{{ИЗОБРАЖЕНИЕ_МАКЕТА}}" and not dp.get("title"):
            par = dp
            while par.tag != qn("w:p"):
                par = par.getparent()
            add_maket_fallback(par)
    d.save(path)
    places = stamp_places(path)
    places["даты"] = date_places(path)
    zone = fit_cover_zone(path)
    toc = refresh_toc(path)
    print(f"   титул: свободное поле {zone[0]:.1f}–{zone[1]:.1f} мм; оглавление: исправлено пунктов {toc}")
    return n, places


if __name__ == "__main__":
    dyn, stat, rnd, out_dir = sys.argv[1], sys.argv[2], sys.argv[3], Path(sys.argv[4])
    out_dir.mkdir(parents=True, exist_ok=True)
    jobs = ((prepare_dynamic, dyn, "Паспорт_Динамика_МЦД"),
            (prepare_static, stat, "Паспорт_Статика_МЦД"),
            (prepare_round, rnd, "Паспорт_Статика_МЦД_круглый"))
    for fn, src, name in jobs:
        out = out_dir / f"{name}.docx"
        meta = fn(src, out)
        n, places = post_process(out)
        (out_dir / f"{name}.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), "utf-8")
        d = docx.Document(out)
        print(name, engine.find_placeholders(d), "строк в таблице:", len(engine.read_tech_table(d)),
              "| Бокланов→Дудин:", n, "| места печати:", places)


# ===================================================================== V2.2 — вёрстка
def _find_par(document, text, start=0):
    """Индекс абзаца в теле документа, текст которого начинается с text."""
    body = list(document.element.body)
    for i in range(start, len(body)):
        if body[i].tag != qn("w:p"):
            continue
        if engine.paragraph_text(body[i]).strip().startswith(text):
            return i
    raise LookupError(text)


def _sdt_index(document):
    body = list(document.element.body)
    return next(i for i, e in enumerate(body) if e.tag == qn("w:sdt"))


def _run_size(par, pt):
    for r in par.iter(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            rPr = r.makeelement(qn("w:rPr"), {})
            r.insert(0, rPr)
        for tag in ("w:sz", "w:szCs"):    # порядок внутри rPr важен для Word
            engine.insert_ordered(rPr, tag).set(qn("w:val"), str(int(pt * 2)))


def cover_like(path, reference):
    """Титул приводится к образцу (паспорт статики): те же кегли «Технический паспорт»
    и наименования изделия, столько же пустых абзацев сверху, без пустых между ними.
    Первый абзац не трогаем — в нём рисунок обложки."""
    import copy
    d, ref = docx.Document(path), docx.Document(reference)

    def block(doc):
        i_t = _find_par(doc, "Технический паспорт")
        i_n = _find_par(doc, "{{НАИМЕНОВАНИЕ_ТИТУЛ}}")
        return i_t, i_n

    rt, rn = block(ref)
    rbody = list(ref.element.body)
    size_title = rbody[rt].find(qn("w:r") + "/" + qn("w:rPr") + "/" + qn("w:sz")).get(qn("w:val"))
    size_name = rbody[rn].find(qn("w:r") + "/" + qn("w:rPr") + "/" + qn("w:sz")).get(qn("w:val"))
    gap_ref = rn - rt - 1                       # пустых абзацев между ними в образце
    top_ref = rt                                # сколько абзацев до «Технический паспорт»

    ref_text = engine.paragraph_text(rbody[rt])
    lead = len(ref_text) - len(ref_text.lstrip(" "))     # отступ пробелами — как в образце

    it, inm = block(d)
    body = list(d.element.body)
    for el in body[it + 1:inm]:                 # пустые между заголовком и наименованием
        el.getparent().remove(el)
    _run_size(body[it], int(size_title) / 2)
    _run_size(body[inm], int(size_name) / 2)
    done = False                                # «Технический паспорт» встаёт как в образце
    for t in body[it].iter(qn("w:t")):
        if done:
            break
        if t.text and t.text.strip():
            t.text = " " * lead + t.text.lstrip(" ")
            done = True
        else:
            t.text = ""                         # отдельный прогон из одних пробелов
    d.save(path)

    d = docx.Document(path)
    it, inm = block(d)
    body = list(d.element.body)
    while it < top_ref:                         # пустые абзацы сверху — как в образце
        body[1].addprevious(copy.deepcopy(body[1]))
        it += 1
    while it > top_ref and it > 1:
        body[1].getparent().remove(body[1])
        body = list(d.element.body)
        it -= 1
    d.save(path)
    return {"кегль_заголовка": int(size_title) / 2, "кегль_наименования": int(size_name) / 2,
            "пустых_сверху": top_ref, "пустых_между": gap_ref}


def toc_like(path, reference):
    """Страница содержания — как в образце (паспорт динамики): те же отступы и
    заголовок «Содержание» перед оглавлением."""
    import copy
    d, ref = docx.Document(path), docx.Document(reference)
    rbody = list(ref.element.body)
    r_from = _find_par(ref, "{{НАИМЕНОВАНИЕ_ТИТУЛ}}") + 1
    r_to = _sdt_index(ref)
    # пустые абзацы-распорки не переносим: «Содержание» и так начинает новую страницу,
    # а лишняя высота на титуле даёт пустую страницу перед оглавлением
    block = [copy.deepcopy(e) for e in rbody[r_from:r_to] if engine.paragraph_text(e).strip()]
    body = list(d.element.body)
    i_from = _find_par(d, "{{НАИМЕНОВАНИЕ_ТИТУЛ}}") + 1
    i_to = _sdt_index(d)
    sdt = body[i_to]
    for el in body[i_from:i_to]:
        el.getparent().remove(el)
    for e in block:
        sdt.addprevious(e)
    d.save(path)
    return len(block)


def keep_block_together(document, heading, stop_re=r"^\d+\.\d+\."):
    """Раздел не разрывается между страницами: заголовок и его пункты держатся вместе,
    поэтому «2.5. Комплектность» целиком уходит на следующую страницу, а не висит внизу."""
    import re as _re
    body = list(document.element.body)
    i = _find_par(document, heading)
    j = i + 1
    while j < len(body) and body[j].tag == qn("w:p"):
        t = engine.paragraph_text(body[j]).strip()
        if t and _re.match(stop_re, t):
            break
        j += 1
    for el in body[i:j - 1]:                     # у последнего абзаца блока keepNext не нужен
        if el.tag != qn("w:p"):
            continue
        pPr = el.find(qn("w:pPr"))
        if pPr is None:
            pPr = el.makeelement(qn("w:pPr"), {})
            el.insert(0, pPr)
        if pPr.find(qn("w:keepNext")) is None:
            # порядок элементов внутри w:pPr важен: keepNext идёт сразу за pStyle,
            # иначе Word откажется открывать файл (LibreOffice такое прощает)
            at = 1 if pPr.find(qn("w:pStyle")) is not None else 0
            pPr.insert(at, pPr.makeelement(qn("w:keepNext"), {}))
    return j - i


def picture_wrap(document, text_anchor, mode="wrapSquare", side="left"):
    """Плавающий рисунок перестаёт лежать поверх текста: текст обтекает его сбоку.
    text_anchor — начало текста абзаца, в котором закреплён рисунок."""
    from lxml import etree
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    i = _find_par(document, text_anchor)
    el = list(document.element.body)[i]
    n = 0
    for a in el.iter(qn("wp:anchor")):
        for ch in list(a):
            if ch.tag.startswith(WP + "wrap"):
                a.remove(ch)
        w = a.makeelement(WP + mode, {})
        if mode == "wrapSquare":
            w.set("wrapText", side)
        docPr = a.find(WP + "docPr")      # обтекание стоит строго перед docPr
        if docPr is not None:
            docPr.addprevious(w)
        else:
            a.append(w)
        a.set("distL", "114300")
        a.set("distR", "114300")
        a.set("behindDoc", "0")
        n += 1
    return n


def fit_cover_heading(path, margin_mm=3.0):
    """«Технический паспорт» на титуле не должен заходить на «квадратики»: если заходит,
    надпись центрируется по свободному тёмно-синему полю в своей строке."""
    import numpy as np
    import subprocess
    from PIL import Image
    path = Path(path)
    TMP.mkdir(parents=True, exist_ok=True)

    def blanked(keep_heading):
        d = docx.Document(path)
        body = list(d.element.body)
        stop = _sdt_index(d)
        head = _find_par(d, "Технический паспорт")
        for i, el in enumerate(body[:stop]):
            if keep_heading and i == head:
                continue
            for t in el.iter(qn("w:t")):
                t.text = ""
        f = TMP / (("head_" if keep_heading else "zone_") + path.name)
        d.save(f)
        return render_pdf(f)

    words = [w for w in words_bbox(blanked(True)) if w[0] == 0 and w[5].strip()]
    if not words:
        return None
    x0 = min(w[1] for w in words); x1 = max(w[3] for w in words)
    y0 = min(w[2] for w in words); y1 = max(w[4] for w in words)
    pdf = blanked(False)
    subprocess.run(["pdftoppm", "-r", "100", "-png", "-f", "1", "-l", "1", str(pdf), str(TMP / "hzone")],
                   capture_output=True)
    png = sorted(TMP.glob("hzone-*.png"))[-1]
    im = np.asarray(Image.open(png).convert("RGB")).astype(int)
    png.unlink()
    H, W, _ = im.shape
    k = 25.4 / 100
    bg = np.median(im[int(H * 0.5):int(H * 0.55), int(W * 0.15):int(W * 0.3)].reshape(-1, 3), axis=0)
    mask = np.abs(im - bg).sum(axis=2) > 60
    band = mask[int(y0 / k):max(int(y1 / k), int(y0 / k) + 1)]
    cols = np.where(band.any(axis=0))[0] * k
    L = max([c for c in cols if c < 0.25 * W * k] or [0]) + margin_mm
    R = min([c for c in cols if c > 0.40 * W * k] or [PAGE_W]) - margin_mm
    if x1 <= R:
        return {"надпись": (round(x0, 1), round(x1, 1)), "поле": (round(L, 1), round(R, 1)),
                "поправка": "не нужна"}
    d = docx.Document(path)
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    par = d.paragraphs[[i for i, p in enumerate(d.paragraphs)
                        if p.text.strip().startswith("Технический паспорт")][0]]
    done = False
    for t in par._p.iter(qn("w:t")):
        if done:
            break
        if t.text and t.text.strip():
            t.text = t.text.strip()
            done = True
        else:
            t.text = ""
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = par.paragraph_format
    pf.left_indent = Mm(L - MARGIN_L)
    pf.right_indent = Mm((PAGE_W - MARGIN_R) - R)
    pf.first_line_indent = Mm(0)
    d.save(path)
    return {"надпись": (round(x0, 1), round(x1, 1)), "поле": (round(L, 1), round(R, 1)),
            "поправка": "по центру поля"}


def normalize_order(path):
    """Порядок дочерних элементов внутри w:pPr, w:rPr и wp:anchor — строго по схеме OOXML.
    Word не открывает файл с нарушенным порядком (LibreOffice открывает), а в образцах
    после конвертации из .doc порядок местами нарушен."""
    import docx_check as chk
    d = docx.Document(path)
    fixed = 0
    for parent_tag, order in chk.ORDER.items():
        slot = chk.SLOT.get(parent_tag, {})
        for el in d.element.iter(parent_tag):
            kids = list(el)
            keys, prev = [], -1.0
            for ch in kids:
                name = ch.tag.split("}")[-1]
                if name in slot:
                    prev = float(slot[name])
                elif name in order:
                    prev = float(order.index(name))
                else:
                    prev = prev + 0.001          # чужие элементы остаются на своём месте
                keys.append(prev)
            new = [c for _, c in sorted(zip(keys, kids), key=lambda kc: kc[0])]
            if new != kids:
                for c in new:
                    el.append(c)
                fixed += 1
    d.save(path)
    return fixed
