# -*- coding: utf-8 -*-
"""
Расчёт параметров паспорта по данным изделия из проекта.

Все коэффициенты и форматы строк лежат в «Шаблоны/правила_расчёта.json» —
их можно править без изменения программы. Значения по умолчанию подобраны так,
чтобы образцы из шаблонов (ASD-TAB P2,5 1450x970 ID и ASD-L 2000х400) давали
ровно те же цифры, что стоят в паспортах. Это ОРИЕНТИРОВОЧНЫЕ правила до сверки
с формулами Калькулятора экранов.

Если в проекте у изделия уже есть значение (масса, мощность, артикул…),
берётся значение из проекта — оно точнее.
"""
import json
import math
import re
from pathlib import Path

DEFAULT_RULES = {
    "Динамика МЦД": {
        "модуль_мм": [320, 160],
        "мин_поле_мм": 50,
        "глубина_корпуса_мм": 175,
        "шаг_по_умолчанию": 2.5,
        "мощность_модуля_макс_Вт": 30,
        "доля_средней_мощности": 0.5,
        "масса_кг_на_м2": 56.88,
        "коэф_массы_двусторонний": 1.8,
        "исполнение_по_умолчанию": "Уличное",
        "титул_кегль_pt": 18,
        "титул_ширина_мм": 120,
        "титул_без_размера": False,
        "стороны": ["Одностороннее", "Двустороннее"],
        "наименование": "Динамическое информационное табло. {стороны}. {исполнение}. {W}x{H} мм",
        "артикул": "ASD-TAB P{шаг} {W}x{H} ID{II}",
        "таблица": {
            "габаритные размеры": "{W}х{H}х{D} мм",
            "разрешение": "{px_w}х{px_h} пкс",
            "шаг пикселя": "{шаг} мм",
            "- максимальная": "{P_max_кВт} кВт",
            "- средняя": "{P_ср_кВт} кВт",
            "масса": "{масса} кг",
        },
    },
    "Статика МЦД": {
        "мощность_Вт_на_м2_стороны": 187.5,
        "блоки_питания_Вт": [40, 60, 100, 150],
        "масса_кг_на_м2": 25,
        "коэф_массы_двусторонний": 1.6,
        "освещение_по_умолчанию": "Световой",
        "титул_кегль_pt": 24,
        "титул_ширина_мм": 147,
        "титул_без_размера": True,
        "стороны": ["Односторонний", "Двусторонний"],
        "наименование": "Статический навигационный указатель. {освещение}. {стороны}. {W}х{H}",
        "артикул": "ASD-L{II} {W}х{H}",
        "таблица": {
            "размер": "{W}х{H}",
            "кол-во сторон": "{сторон}",
            "потребляемая мощность": "{P_Вт}",
            "масса": "{масса}",
        },
    },
}


DEFAULT_RULES["Статика МЦД круглый"] = {
    "форма": "круг",
    "мощность_Вт_на_м2_стороны": 88.4,       # подобрано по образцу ASD-L D=1200: 100 Вт
    "блоки_питания_Вт": [40, 60, 100, 150],
    "масса_кг_на_м2": 22.1,                  # по образцу ASD-L D=1200: 25 кг
    "коэф_массы_двусторонний": 1.6,
    "освещение_по_умолчанию": "Световой",
    "титул_кегль_pt": 24,
    "титул_ширина_мм": 147,
    "титул_без_размера": True,
    "стороны": ["Односторонний", "Двусторонний"],
    "наименование": "Статический навигационный указатель. {освещение}. {стороны}. D={D}",
    "артикул": "ASD-L{II} D={D}",
    "таблица": {
        "размер": "{D}х{D}",
        "кол-во сторон": "{сторон}",
        "потребляемая мощность": "{P_Вт}",
        "масса": "{масса}",
    },
}


def load_rules(path: Path) -> dict:
    rules = json.loads(json.dumps(DEFAULT_RULES))
    if path.exists():
        try:
            user = json.loads(path.read_text("utf-8"))
            for t, r in user.items():
                rules.setdefault(t, {}).update(r)
        except Exception:
            pass
    else:
        try:
            path.write_text(json.dumps(DEFAULT_RULES, ensure_ascii=False, indent=2), "utf-8")
        except Exception:
            pass
    return rules


def num(x, digits=2) -> str:
    """0.6 → «0,6», 2.5 → «2,5», 80.0 → «80»."""
    s = f"{float(x):.{digits}f}"
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s.replace(".", ",")


def _int(v, default=0):
    try:
        return int(round(float(str(v).replace(",", "."))))
    except Exception:
        return default


def _float(v, default=0.0):
    try:
        return float(str(v).replace(",", "."))
    except Exception:
        return default


FONT_FILE = None          # путь к RussianRail G Pro — задаёт программа (папка «Шрифты»)
_font_cache = {}


def _measure_mm(text, pt):
    """Ширина строки в мм реальным шрифтом титула; без шрифта — оценка."""
    if FONT_FILE and Path(FONT_FILE).exists():
        try:
            from PIL import ImageFont
            key = (FONT_FILE, pt)
            if key not in _font_cache:
                _font_cache[key] = ImageFont.truetype(str(FONT_FILE), int(pt * 10))
            return _font_cache[key].getlength(text) / 10 * 25.4 / 72
        except Exception:
            pass
    return len(text) * pt * 0.24     # ≈ средняя ширина знака RussianRail


def wrap_mm(text, pt, max_mm):
    words, lines, cur = text.split(), [], ""
    for w in words:
        cand = (cur + " " + w).strip()
        if cur and _measure_mm(cand, pt) > max_mm:
            lines.append(cur)
            cur = w
        else:
            cur = cand
    if cur:
        lines.append(cur)
    return lines


SIZE_TAIL = re.compile(r"\s*(?:\d{2,5}\s*[xх×]\s*\d{2,5}|(?:D\s*=\s*|Ø\s*)\d{2,5})(\s*мм)?\s*$")


def title_lines(name: str, article: str, typ: str = "", rules: dict = None) -> str:
    """Строки титула как в образце: наименование, перенесённое по ширине поля титула
    (реальным шрифтом), последней строкой — артикул. Для статики размер в наименовании
    на титул не выносится (он есть в артикуле)."""
    r = (rules or DEFAULT_RULES).get(typ, {})
    pt = r.get("титул_кегль_pt", 18)
    width = r.get("титул_ширина_мм", 120)
    if r.get("титул_без_размера"):
        name = SIZE_TAIL.sub("", name).rstrip()
    lines = wrap_mm(name, pt, width)
    if article:
        lines.append(article)
    return "\n".join(lines)


def variables(item: dict, rules: dict) -> dict:
    """Все расчётные величины изделия."""
    t = item.get("тип", "")
    r = rules.get(t, {})
    W, H = _int(item.get("ширина")), _int(item.get("высота"))
    sides = max(1, min(2, _int(item.get("сторон"), 1)))
    area = W * H / 1e6
    if r.get("форма") == "круг":
        H = W
        area = math.pi * W * W / 4 / 1e6
    v = {"W": W, "H": H, "D": W, "сторон": sides, "площадь_м2": num(area, 3),
         "стороны": r.get("стороны", ["", ""])[sides - 1],
         "II": " II" if sides == 2 else "", "тираж": _int(item.get("тираж"), 1)}

    if t == "Динамика МЦД":
        mw, mh = r["модуль_мм"]
        pitch = _float(item.get("шаг") or r["шаг_по_умолчанию"], r["шаг_по_умолчанию"])
        margin = r["мин_поле_мм"]
        cols = max(0, int((W - 2 * margin) // mw))
        rows = max(0, int((H - 2 * margin) // mh))
        modules = cols * rows
        p_max = modules * r["мощность_модуля_макс_Вт"] * sides / 1000
        mass = area * r["масса_кг_на_м2"] * (r["коэф_массы_двусторонний"] if sides == 2 else 1)
        v.update({
            "D": r["глубина_корпуса_мм"], "шаг": num(pitch),
            "модулей_по_ширине": cols, "модулей_по_высоте": rows, "модулей": modules,
            "поле_гор_мм": num((W - cols * mw) / 2, 1), "поле_верт_мм": num((H - rows * mh) / 2, 1),
            "px_w": int(round(cols * mw / pitch)), "px_h": int(round(rows * mh / pitch)),
            "P_max_кВт": num(p_max), "P_ср_кВт": num(p_max * r["доля_средней_мощности"]),
            "масса": num(mass, 0),
            "исполнение": item.get("исполнение") or r["исполнение_по_умолчанию"],
        })
    elif t.startswith("Статика МЦД"):
        load = area * r["мощность_Вт_на_м2_стороны"] * sides
        psus = sorted(r["блоки_питания_Вт"])
        if load <= psus[-1]:
            power = next(p for p in psus if p >= load - 1e-6)
            n_psu = 1
        else:
            n_psu = math.ceil(load / psus[-1] - 1e-9)
            power = n_psu * psus[-1]
        mass = area * r["масса_кг_на_м2"] * (r["коэф_массы_двусторонний"] if sides == 2 else 1)
        v.update({"P_Вт": num(power, 0), "нагрузка_Вт": num(load, 0), "блоков_питания": n_psu,
                  "масса": num(mass, 0),
                  "освещение": item.get("освещение") or r["освещение_по_умолчанию"]})
    # значения, заданные в проекте, главнее расчёта
    if item.get("масса"):
        v["масса"] = num(_float(item["масса"]), 1)
    if item.get("мощность_Вт"):
        v["P_Вт"] = num(_float(item["мощность_Вт"]), 0)
        v["P_max_кВт"] = num(_float(item["мощность_Вт"]) / 1000)
    return v


def _fmt(pattern, v):
    class D(dict):
        def __missing__(self, k):
            return "{" + k + "}"
    return pattern.format_map(D(v))


def compute(item: dict, rules: dict, template_table: list):
    """→ (значения меток, строки таблицы [[параметр, значение, источник]], переменные)."""
    t = item.get("тип", "")
    r = rules.get(t, {})
    v = variables(item, rules)
    name = item.get("наименование") or _fmt(r.get("наименование", ""), v)
    article = item.get("артикул") or _fmt(r.get("артикул", ""), v)
    values = {
        "НАИМЕНОВАНИЕ": name,
        "АРТИКУЛ": article,
        "НАИМЕНОВАНИЕ_ТИТУЛ": title_lines(name, article, t, rules),
    }
    if item.get("номер_макета"):
        values["НОМЕР_МАКЕТА"] = str(item["номер_макета"])
    if item.get("серийный_номер"):
        values["СЕРИЙНЫЙ_НОМЕР"] = str(item["серийный_номер"])

    from_project = {"масса"} if item.get("масса") else set()
    if item.get("мощность_Вт"):
        from_project |= {"P_Вт", "P_max_кВт"}
    rows = []
    for name_, value in template_table:
        key = name_.strip().lower()
        pat = next((p for k, p in r.get("таблица", {}).items() if key.startswith(k.lower())), None)
        if pat is None:
            rows.append([name_, value, "шаблон"])
            continue
        new = _fmt(pat, v)
        src = "проект" if any("{" + f + "}" in pat for f in from_project) else "расчёт"
        rows.append([name_, new, src])
    return values, rows, v
