# -*- coding: utf-8 -*-
"""
Чтение сохранённого проекта Калькулятора экранов.

Формат файла не зашит жёстко: читатель ищет в файле списки изделий по смыслу
полей (ширина/высота/тираж/тип/префикс/артикул/номер макета…) — русские и
английские названия, с единицами или без. Поддерживается:
  • JSON (любое расширение, если внутри JSON) — вложенность любая;
  • CSV / XLSX — таблица, первая строка — заголовки (например, выгрузка журнала).

Результат: {"проект": "название", "изделия": [dict, ...]} — у каждого изделия
нормализованные ключи: префикс, тип («Динамика МЦД»/«Статика МЦД»), ширина,
высота, тираж, сторон, шаг, артикул, наименование, номер_макета,
серийный_номер, исполнение, освещение, масса, мощность_Вт, _источник.
"""
import csv
import json
import re
from pathlib import Path

# синонимы полей (сравниваются без регистра, пробелов, знаков и единиц)
SYN = {
    "ширина": ["ширина", "width", "w", "ширинамм", "шир", "widthmm", "b"],
    "диаметр": ["диаметр", "diameter", "d", "ø", "диаметрмм", "dmm"],
    "форма": ["форма", "shape", "формаизделия"],
    "высота": ["высота", "height", "h", "высотамм", "выс", "heightmm"],
    "тираж": ["тираж", "количество", "колво", "кол", "qty", "quantity", "count", "шт",
              "колвошт", "количествошт", "amount"],
    "тип": ["тип", "type", "вид", "категория", "kind", "category", "типизделия", "producttype"],
    "сторон": ["сторон", "стороны", "колвосторон", "sides", "сторонность", "двусторонний",
               "двустороннее", "doublesided", "twosided"],
    "префикс": ["префикс", "prefix", "позиция", "поз", "код", "code", "метка", "label", "id"],
    "артикул": ["артикул", "article", "sku", "screenarticle"],
    "наименование": ["наименование", "название", "name", "title", "изделие", "product"],
    "номер_макета": ["номермакета", "макет", "макетномер", "maket", "layout", "layoutno",
                     "nмакета", "макета"],
    "серийный_номер": ["серийныйномер", "серийный", "serial", "serialnumber", "sn", "сн",
                       "партномер", "partnumber"],
    "шаг": ["шаг", "шагпикселя", "pitch", "p", "pixelpitch"],
    "исполнение": ["исполнение", "среда", "установка", "environment", "outdoor", "indoor"],
    "освещение": ["освещение", "подсветка", "lighting", "backlight", "световой"],
    "масса": ["масса", "вес", "weight", "массакг", "веськг", "mass"],
    "мощность_Вт": ["мощность", "power", "мощностьвт", "потребляемаямощность", "powerw"],
}
PROJECT_KEYS = ["названиепроекта", "проект", "project", "projectname", "объект", "object",
                "заказ", "order", "станция"]


def _norm(k) -> str:
    k = str(k).lower().replace("ё", "е")
    k = re.sub(r"[,(].*$", "", k) if not k.startswith("(") else k   # «Ширина, мм» → «ширина»
    return re.sub(r"[^a-zа-я0-9]", "", k)


def _field_of(key):
    n = _norm(key)
    for field, names in SYN.items():
        if n in names:
            return field
    for field, names in SYN.items():            # «ширина_изделия», «heightMm»
        if any(len(s) > 3 and n.startswith(s) for s in names):
            return field
    return None


def _detect_type(text) -> str:
    t = str(text).lower()
    if any(s in t for s in ("динам", "dynamic", "экран", "табло", "screen", "led", "tab")):
        return "Динамика МЦД"
    if any(s in t for s in ("стат", "static", "указат", "sign", "лайтбокс", "световой короб",
                            "as-l", "asd-l")):
        return "Статика МЦД"
    return ""


def _sides(v) -> int:
    if isinstance(v, bool):
        return 2 if v else 1
    s = str(v).lower()
    if any(x in s for x in ("дву", "2", "double", "two", "true", "да")):
        return 2
    return 1


def normalize_item(raw: dict, hint_type: str = "") -> dict:
    it = {}
    extra = {}
    for k, v in raw.items():
        if isinstance(v, (dict, list)):
            continue
        f = _field_of(k)
        if f and f not in it and v not in (None, ""):
            it[f] = v
        else:
            extra[k] = v
    # размеры вида «2000х400» в одном поле
    if ("ширина" not in it or "высота" not in it):
        for k, v in raw.items():
            m = re.search(r"(\d{2,5})\s*[xх×*]\s*(\d{2,5})", str(v))
            if m and _norm(k) in ("размер", "размеры", "size", "габарит", "габариты", "wxh"):
                it.setdefault("ширина", m.group(1))
                it.setdefault("высота", m.group(2))
    # круглые: диаметр отдельным полем или «D=1200» / «Ø1200» в размере
    if "диаметр" not in it:
        for k, v in raw.items():
            m = re.search(r"(?:\bD\s*=?\s*|Ø\s*)(\d{2,5})", str(v))
            if m and (_norm(k) in ("размер", "размеры", "size", "габарит", "габариты")
                      or _field_of(k) in ("артикул", "наименование")):
                it["диаметр"] = m.group(1)
                break
    round_ = "диаметр" in it or any(x in str(it.get("форма", "")).lower()
                                    for x in ("круг", "round", "circle"))
    if "диаметр" in it:
        m = re.search(r"\d+", str(it["диаметр"]))
        if m:
            it.setdefault("ширина", m.group(0))
            it.setdefault("высота", m.group(0))
    if round_:
        it["форма"] = "круг"
    if "сторон" in it:
        it["сторон"] = _sides(it["сторон"])
    tp = _detect_type(it.get("тип", "")) or _detect_type(it.get("артикул", "")) \
        or _detect_type(it.get("наименование", "")) or hint_type
    if tp == "Статика МЦД" and round_:
        tp = "Статика МЦД круглый"
    it["тип"] = tp
    if "шаг" in it:
        m = re.search(r"\d+(?:[.,]\d+)?", str(it["шаг"]))
        it["шаг"] = m.group(0).replace(",", ".") if m else ""
    for f in ("ширина", "высота", "тираж"):
        if f in it:
            m = re.search(r"\d+(?:[.,]\d+)?", str(it[f]))
            it[f] = int(float(m.group(0).replace(",", "."))) if m else 0
    it.setdefault("тираж", 1)
    it.setdefault("сторон", 1)
    it["_прочее"] = extra
    return it


def _is_item(d) -> bool:
    if not isinstance(d, dict):
        return False
    fields = {_field_of(k) for k in d}
    if "ширина" in fields and "высота" in fields or "диаметр" in fields:
        return True
    return any(re.search(r"\d{2,5}\s*[xх×]\s*\d{2,5}|(?:\bD\s*=?\s*|Ø\s*)\d{2,5}", str(v))
               for v in d.values()) and \
        bool(fields & {"тираж", "тип", "артикул", "префикс"})


def _walk(node, path, found, project):
    if isinstance(node, dict):
        for k, v in node.items():
            if isinstance(v, str) and _norm(k) in PROJECT_KEYS and not project.get("name"):
                project["name"] = v.strip()
            if isinstance(v, dict) and _norm(k) in PROJECT_KEYS:
                for kk, vv in v.items():
                    if isinstance(vv, str) and _norm(kk) in ["название", "name", "наименование"] + PROJECT_KEYS \
                            and not project.get("name"):
                        project["name"] = vv.strip()
            _walk(v, path + [str(k)], found, project)
    elif isinstance(node, list):
        dicts = [x for x in node if isinstance(x, dict)]
        items = [x for x in dicts if _is_item(x)]
        if items and len(items) >= max(1, len(dicts) // 2):
            hint = _detect_type(" ".join(path[-2:]))
            for x in items:
                found.append((x, hint, "/".join(path)))
        for i, x in enumerate(node):
            if not (isinstance(x, dict) and _is_item(x)):
                _walk(x, path + [str(i)], found, project)


def _read_table(path: Path):
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        rows = []
        for ws in wb.worksheets:
            data = [list(r) for r in ws.iter_rows(values_only=True)]
            # строка заголовков — первая, где есть ширина и высота
            for hi, r in enumerate(data[:30]):
                fields = {_field_of(c) for c in r if c is not None}
                if {"ширина", "высота"} <= fields or "тираж" in fields and "артикул" in fields:
                    head = [str(c) if c is not None else "" for c in r]
                    for rr in data[hi + 1:]:
                        if any(c not in (None, "") for c in rr):
                            rows.append(dict(zip(head, rr)))
                    break
        return rows
    text = path.read_text("utf-8-sig", errors="ignore")
    dialect = csv.Sniffer().sniff(text[:4000], delimiters=";,\t")
    return list(csv.DictReader(text.splitlines(), dialect=dialect))


# ---------------------------------------------------------------- формат Конструктора
def _num(v):
    m = re.search(r"\d+(?:[.,]\d+)?", str(v if v is not None else ""))
    return float(m.group(0).replace(",", ".")) if m else None


ARTICLE_RE = re.compile(r"(?:AS|ASD)-[A-Z0-9][^\s]*(?:\s+\S+)*$")


def _article_tail(text):
    """«Динамическое информационное табло. Двустороннее. Внутреннее. ASD-TAB P2,5 II 1780x580 OD»
    → «ASD-TAB P2,5 II 1780x580 OD». В Конструкторе поле «Артикул экрана» хранит наименование
    вместе с артикулом; в паспорте это разные поля."""
    m = ARTICLE_RE.search(str(text or "").strip())
    return m.group(0).strip() if m else str(text or "").strip()


def _strip_article(name, article):
    """В Конструкторе наименование = «… Односторонний ASD-L 1508x357» — артикул в конце
    дублируется; в паспорте наименование и артикул — отдельные поля."""
    name, article = str(name or "").strip(), str(article or "").strip()
    if article and name.endswith(article):
        name = name[: -len(article)].rstrip(" ,;")      # точку в конце наименования сохраняем
    return name


def _passport_block(raw, it):
    ps = raw.get("passport") if isinstance(raw.get("passport"), dict) else {}
    if ps.get("weight_kg") not in (None, ""):
        it["масса"] = ps["weight_kg"]
    if ps.get("psu_total_w") not in (None, "", 0):
        it["мощность_Вт"] = ps["psu_total_w"]
    if ps.get("psu"):
        it["блоки_питания"] = ps["psu"]
    if ps.get("calc"):
        it["_расчёт"] = ps["calc"]
    if ps.get("error"):
        it["_ошибка_расчёта"] = ps["error"]


def load_ledproj(data, path):
    """Точное чтение файла проекта Конструктора LED-экранов (format = LED_SCREEN_PROJECT)."""
    items = []
    for i, s in enumerate(data.get("screens") or []):
        if not isinstance(s, dict):
            continue
        mode = str(s.get("mode", ""))
        it = {"префикс": str(s.get("spec_position") or i + 1),
              "тип": "Динамика МЦД" if mode == "МЦД" else f"Динамика {mode or 'РЖД'} (нет шаблона)",
              "ширина": int(_num(s.get("width")) or 0), "высота": int(_num(s.get("height")) or 0),
              "сторон": int(_num(s.get("sides")) or 1),
              "тираж": int(_num(s.get("quantity")) or 1),
              "шаг": (str(_num(s.get("pixel_pitch")) or "").rstrip("0").rstrip(".") or ""),
              "исполнение": str(s.get("environment") or "Уличное"),
              "артикул": "",
              "_источник": "screens"}
        full = str(s.get("screen_article") or "").strip()
        it["артикул"] = str(s.get("module_article") or "").strip() or _article_tail(full)
        nm = _strip_article(full, it["артикул"])
        if nm:
            it["наименование"] = nm
        if _num(s.get("module_total")):                 # LED-модулей на сторону — как в Конструкторе
            it["модулей_на_сторону"] = int(_num(s.get("module_total")))
        serials = [str(x.get("part_number") or "").strip() for x in (s.get("instances") or [])
                   if isinstance(x, dict) and str(x.get("part_number") or "").strip()]
        if not serials and str(s.get("part_number") or "").strip():
            serials = [str(s.get("part_number")).strip()]
        if serials:
            it["серийный_номер"] = ", ".join(serials) if len(serials) > 1 else serials[0]
        _passport_block(s, it)
        items.append(it)
    for i, s in enumerate(data.get("static_signs") or []):
        if not isinstance(s, dict):
            continue
        mode = str(s.get("mode", ""))
        rnd = bool(s.get("round"))
        w = _num(s.get("width")) or 0
        h = _num(s.get("height")) or 0
        d = _num(s.get("diameter")) or (w if rnd else None)
        if mode == "МЦД" and s.get("face_type", "С подсветкой") == "С подсветкой":
            typ = "Статика МЦД круглый" if rnd else "Статика МЦД"
        else:
            typ = f"Статика {mode or '?'} {s.get('face_type', '')} (нет шаблона)".replace("  ", " ")
        it = {"префикс": str((s.get("journal_order") if s.get("journal_order") is not None else i) + 1),
              "тип": typ,
              "ширина": int(d if rnd else w), "высота": int(d if rnd else h),
              "сторон": int(_num(s.get("sides")) or 1),
              "тираж": int(_num(s.get("quantity")) or 1),
              "артикул": str(s.get("article") or "").strip(),
              "освещение": "Световой" if s.get("face_type", "С подсветкой") == "С подсветкой" else "",
              "_источник": "static_signs"}
        if rnd:
            it["форма"] = "круг"
        if str(s.get("layout_no") or "").strip():
            it["номер_макета"] = str(s["layout_no"]).strip()
        nm = _strip_article(s.get("name"), s.get("article"))
        if nm and nm != str(s.get("name") or "").strip():
            pass    # наименование было «авто» (с артикулом) — строим по правилам паспорта
        elif nm:
            it["наименование"] = nm
        _passport_block(s, it)
        items.append(it)
    items.sort(key=lambda x: (x["_источник"] != "screens",))
    return {"проект": str(data.get("project_name") or Path(path).stem), "изделия": items,
            "файл": str(path), "формат": f"Конструктор LED-экранов, версия файла {data.get('version')}",
            "программа": data.get("app") or ""}


def load_project(path) -> dict:
    path = Path(path)
    project = {}
    found = []
    data = None
    if path.suffix.lower() not in (".csv", ".xlsx", ".xlsm"):
        try:
            data = json.loads(path.read_text("utf-8-sig"))
        except Exception:
            data = None
    if isinstance(data, dict) and data.get("format") == "LED_SCREEN_PROJECT":
        return load_ledproj(data, path)
    if data is not None:
        _walk(data, [], found, project)
    else:
        try:
            rows = _read_table(path)
        except Exception as e:
            raise ValueError(f"Не удалось прочитать файл проекта: {e}")
        found = [(r, "", "таблица") for r in rows]
    items = []
    for raw, hint, src in found:
        it = normalize_item(raw, hint)
        if not it.get("ширина") or not it.get("высота"):
            continue
        it["_источник"] = src
        items.append(it)
    if not items:
        raise ValueError("В файле не найдено ни одного изделия (нет полей ширины и высоты).\n"
                         "Пришлите пример файла — формат будет добавлен.")
    for i, it in enumerate(items, 1):
        it.setdefault("префикс", str(i))
    return {"проект": project.get("name") or path.stem, "изделия": items, "файл": str(path)}
