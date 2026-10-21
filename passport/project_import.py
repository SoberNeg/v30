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
    "артикул": ["артикул", "article", "sku", "партномер", "partnumber"],
    "наименование": ["наименование", "название", "name", "title", "изделие", "product"],
    "номер_макета": ["номермакета", "макет", "макетномер", "maket", "layout", "layoutno",
                     "nмакета", "макета"],
    "серийный_номер": ["серийныйномер", "серийный", "serial", "serialnumber", "sn", "сн"],
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
