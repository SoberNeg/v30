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
        # мощность — как в Конструкторе: 1 блок питания на 4 LED-модуля (все стороны),
        # «максимальная» = сумма мощностей блоков питания, «средняя» = доля от неё
        "модулей_на_блок_питания": 4,
        "блок_питания_Вт": 200,
        "доля_средней_мощности": 0.5,
        "масса_кг_на_м2": 56.88,
        "коэф_массы_двусторонний": 1.8,
        "исполнение_по_умолчанию": "Уличное",
        "титул_кегль_pt": 18,
        "титул_ширина_мм": 120,
        "титул_без_размера": False,
        "стороны": ["Одностороннее", "Двустороннее"],
        "наименование": "Динамическое информационное табло. {стороны}. {исполнение}.",
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
        # мощность — как в Конструкторе: LED по 1,5 Вт × 1,2 (запас), блоки питания
        # 40/60/100/150 Вт — наименьшее число блоков, затем наименьшая суммарная мощность;
        # в паспорт идёт сумма мощностей блоков питания
        "LED_на_м2_одностороннего": 60,
        "шаг_торцевых_LED_мм": 90,
        "короткая_сторона_без_торцевых_LED_мм": 400,
        "шаг_LED_по_окружности_мм": 65,
        "мощность_LED_Вт": 1.5,
        "запас_мощности": 1.2,
        "блоки_питания_Вт": [40, 60, 100, 150],
        "масса_кг_на_м2": 25,
        "коэф_массы_двусторонний": 1.6,
        "освещение_по_умолчанию": "Световой",
        "титул_кегль_pt": 24,
        "титул_ширина_мм": 147,
        "титул_без_размера": True,
        "стороны": ["Односторонний", "Двусторонний"],
        "наименование": "Статический навигационный указатель. {освещение}. {стороны}.",
        "артикул": "ASD-L{II} {W}х{H}",
        "таблица": {
            "размер": "{W}х{H}",
            "тип установки": "{установка}",
            "кол-во сторон": "{сторон}",
            "потребляемая мощность": "{P_Вт}",
            "масса": "{масса}",
        },
    },
}


DEFAULT_RULES["Статика МЦД круглый"] = {
    "форма": "круг",
    # мощность — как в Конструкторе: LED по 1,5 Вт × 1,2 (запас), блоки питания
    # 40/60/100/150 Вт — наименьшее число блоков, затем наименьшая суммарная мощность;
    # в паспорт идёт сумма мощностей блоков питания
    "LED_на_м2_одностороннего": 60,
    "шаг_торцевых_LED_мм": 90,
    "короткая_сторона_без_торцевых_LED_мм": 400,
    "шаг_LED_по_окружности_мм": 65,
    "мощность_LED_Вт": 1.5,
    "запас_мощности": 1.2,
    "блоки_питания_Вт": [40, 60, 100, 150],
    "масса_кг_на_м2": 22.1,                  # по образцу ASD-L D=1200: 25 кг
    "коэф_массы_двусторонний": 1.6,
    "освещение_по_умолчанию": "Световой",
    "титул_кегль_pt": 24,
    "титул_ширина_мм": 147,
    "титул_без_размера": True,
    "стороны": ["Односторонний", "Двусторонний"],
    "наименование": "Статический навигационный указатель. {освещение}. {стороны}.",
    "артикул": "ASD-L{II} D={D}",
    "таблица": {
        "размер": "{D}х{D}",
        "тип установки": "{установка}",
        "кол-во сторон": "{сторон}",
        "потребляемая мощность": "{P_Вт}",
        "масса": "{масса}",
    },
}


RULES_VERSION = 3
OBSOLETE_KEYS = ("мощность_Вт_на_м2_стороны", "мощность_модуля_макс_Вт")   # до V1.9


def _migrate(user: dict) -> dict:
    """Правила из прежних версий: наименование с размером → без размера (если не меняли
    по-своему), новые строки таблицы (например, «тип установки») добавляются."""
    for t, r in user.items():
        if not isinstance(r, dict) or t not in DEFAULT_RULES:
            continue
        d = DEFAULT_RULES[t]
        if isinstance(r.get("наименование"), str) and re.search(r"\{(W|H|D)\}", r["наименование"]):
            r["наименование"] = d["наименование"]
        if isinstance(r.get("таблица"), dict):
            r["таблица"] = {**d.get("таблица", {}), **r["таблица"]}
        for k in OBSOLETE_KEYS:                  # прежняя оценка мощности по площади
            r.pop(k, None)
    return user


def load_rules(path: Path) -> dict:
    """Правила по умолчанию + правки пользователя из файла. Файл старой версии
    сохраняется копией (…v1.json) и переписывается в новой."""
    rules = json.loads(json.dumps(DEFAULT_RULES))
    user = {}
    if path.exists():
        try:
            user = json.loads(path.read_text("utf-8"))
        except Exception:
            user = {}
        if user.get("_версия", 1) < RULES_VERSION:
            try:
                path.with_name(path.stem + f".v{user.get('_версия', 1)}.json").write_text(
                    path.read_text("utf-8"), "utf-8")
            except Exception:
                pass
            user = _migrate(user)
            user["_версия"] = RULES_VERSION
            try:
                path.write_text(json.dumps(user, ensure_ascii=False, indent=2), "utf-8")
            except Exception:
                pass
    else:
        try:
            path.write_text(json.dumps({"_версия": RULES_VERSION, **DEFAULT_RULES},
                                       ensure_ascii=False, indent=2), "utf-8")
        except Exception:
            pass
    for t, r in user.items():
        if isinstance(r, dict):
            base = rules.setdefault(t, {})
            tbl = {**base.get("таблица", {}), **r.get("таблица", {})}
            base.update(r)
            if tbl:
                base["таблица"] = tbl
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


def _name_without_size(name, r):
    if r.get("титул_без_размера", True):
        name = SIZE_TAIL.sub("", name).rstrip()
    return name


def title_lines(name: str, article: str, typ: str = "", rules: dict = None,
                width=None, pt=None, max_name_lines=3) -> str:
    """Строки титула: наименование, перенесённое по ширине поля титула (только тёмно-синяя
    зона обложки; ширина и кегль — из шаблона), последней строкой — артикул. Если с кеглем
    шаблона наименование не укладывается в max_name_lines строк, кегль уменьшается
    (до 80 %), чтобы не рвать фразы вроде «Статический навигационный»."""
    r = (rules or DEFAULT_RULES).get(typ, {})
    pt = pt or r.get("титул_кегль_pt", 18)
    width = width or r.get("титул_ширина_мм", 120)
    name = _name_without_size(name, r)
    def good(ls):     # не больше max строк и нет «висячих» слов вроде «Статический»
        return len(ls) <= max_name_lines and not any(
            len(l.split()) == 1 and not l.endswith(".") for l in ls[:-1])
    lines = wrap_mm(name, pt, width)
    if not good(lines):
        p, fallback = pt, None
        while p > pt * 0.8:
            p -= 0.5
            cand = wrap_mm(name, p, width)
            if good(cand):
                lines = cand
                break
            if fallback is None and len(cand) <= max_name_lines:
                fallback = cand
        else:
            lines = fallback or lines
    if article:
        lines.append(article)
    return "\n".join(lines)


TITLE_SAFETY = 0.97       # запас: Word и LibreOffice считают ширину чуть иначе, чем PIL


def title_fit(text: str, width, pt) -> float:
    """Наибольший кегль ≤ pt, при котором каждая строка титула помещается в поле
    (с небольшим запасом, чтобы строка не переносилась в Word)."""
    p = float(pt)
    width = width * TITLE_SAFETY
    for line in [l for l in str(text).splitlines() if l.strip()]:
        w = _measure_mm(line, pt)
        if w > width:
            p = min(p, pt * width / w)
    return max(1.0, int(p * 2) / 2)


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
    v = {"W": W, "H": H, "D": W, "сторон": sides,
         "установка": (r.get("установка") or ["Настенный", "Подвесной"])[sides - 1], "площадь_м2": num(area, 3),
         "стороны": r.get("стороны", ["", ""])[sides - 1],
         "II": " II" if sides == 2 else "", "тираж": _int(item.get("тираж"), 1)}

    if t == "Динамика МЦД":
        mw, mh = r["модуль_мм"]
        pitch = _float(item.get("шаг") or r["шаг_по_умолчанию"], r["шаг_по_умолчанию"])
        margin = r["мин_поле_мм"]
        cols = max(0, int((W - 2 * margin) // mw))
        rows = max(0, int((H - 2 * margin) // mh))
        modules = cols * rows
        if _int(item.get("модулей_на_сторону")) > 0:         # как посчитал Конструктор
            modules = _int(item["модулей_на_сторону"])
        n_psu = math.ceil(modules * sides / r["модулей_на_блок_питания"] - 1e-9)
        p_max = n_psu * r["блок_питания_Вт"] / 1000
        mass = area * r["масса_кг_на_м2"] * (r["коэф_массы_двусторонний"] if sides == 2 else 1)
        v.update({
            "D": r["глубина_корпуса_мм"], "шаг": num(pitch),
            "модулей_по_ширине": cols, "модулей_по_высоте": rows, "модулей": modules,
            "поле_гор_мм": num((W - cols * mw) / 2, 1), "поле_верт_мм": num((H - rows * mh) / 2, 1),
            "px_w": int(round(cols * mw / pitch)), "px_h": int(round(rows * mh / pitch)),
            "P_max_кВт": num(p_max), "P_ср_кВт": num(p_max * r["доля_средней_мощности"]),
            "блоков_питания": n_psu, "блоки_питания": f'{n_psu}×{r["блок_питания_Вт"]} Вт' if n_psu else "",
            "масса": num(mass, 0),
            "исполнение": item.get("исполнение") or r["исполнение_по_умолчанию"],
        })
    elif t.startswith("Статика МЦД"):
        leds = static_leds(W, H, sides, r.get("форма") == "круг", r)
        load = leds * r["мощность_LED_Вт"] * r["запас_мощности"]
        psu = psu_combo(load, r["блоки_питания_Вт"])
        power, n_psu = sum(psu), len(psu)
        mass = area * r["масса_кг_на_м2"] * (r["коэф_массы_двусторонний"] if sides == 2 else 1)
        v.update({"P_Вт": num(power, 0), "нагрузка_Вт": num(load, 1), "блоков_питания": n_psu,
                  "LED_шт": leds, "блоки_питания": psu_text(psu),
                  "масса": num(mass, 0),
                  "освещение": item.get("освещение") or r["освещение_по_умолчанию"]})
    # значения, заданные в проекте, главнее расчёта
    if item.get("масса") not in (None, ""):
        v["масса"] = num(_float(item["масса"]), 2)          # как столбец «Вес, кг» журнала
    if item.get("мощность_Вт") not in (None, ""):           # суммарная мощность блоков питания
        pw = _float(item["мощность_Вт"])
        v["P_Вт"] = num(pw, 0)
        if item.get("блоки_питания"):
            v["блоки_питания"] = " + ".join(f'{num(_float(b.get("qty")), 0)}×{num(_float(b.get("w")), 0)} Вт'
                                            for b in item["блоки_питания"] if isinstance(b, dict))
        v["P_max_кВт"] = num(pw / 1000)
        if "доля_средней_мощности" in r:
            v["P_ср_кВт"] = num(pw / 1000 * r["доля_средней_мощности"])
    return v


def static_leds(W, H, sides, rnd, r):
    """Количество LED статики МЦД — те же правила, что в Конструкторе (V30.68/V30.69/V20.94)."""
    if sides == 1:
        face = (math.pi * W * W / 4 if rnd else W * H) / 1e6
        return max(1, math.ceil(face * r["LED_на_м2_одностороннего"] - 1e-9))
    if rnd:                                       # двусторонний круг — LED по окружности
        return math.ceil(math.pi * W / r["шаг_LED_по_окружности_мм"])
    long_, short = max(W, H), min(W, H)
    if long_ <= 0:
        return 0
    step = r["шаг_торцевых_LED_мм"]
    n = math.ceil(long_ / step - 1e-9) * 2
    if short > r["короткая_сторона_без_торцевых_LED_мм"]:
        n += math.ceil(short / step - 1e-9) * 2
    return n


def psu_combo(required, allowed):
    """Подбор блоков питания как в Конструкторе: сначала наименьшее число блоков,
    затем наименьшая суммарная мощность (при равенстве — блоки крупнее)."""
    import itertools
    allowed = tuple(sorted({int(a) for a in allowed}))
    if required <= 1e-9 or not allowed:
        return []
    n0 = max(1, math.ceil(required / allowed[-1]))
    for n in range(n0, n0 + len(allowed) + 1):
        cand = [(sum(c), c) for c in itertools.combinations_with_replacement(allowed, n)
                if sum(c) + 1e-9 >= required]
        if cand:
            best = min(cand, key=lambda x: (x[0], tuple(-p for p in sorted(x[1], reverse=True))))[1]
            return sorted(best, reverse=True)
    return []


def psu_text(psu):
    """[100, 60, 60] → «1×100 Вт + 2×60 Вт»."""
    out = []
    for w in sorted(set(psu), reverse=True):
        out.append(f"{psu.count(w)}×{w} Вт")
    return " + ".join(out)


def _fmt(pattern, v):
    class D(dict):
        def __missing__(self, k):
            return "{" + k + "}"
    return pattern.format_map(D(v))


def compute(item: dict, rules: dict, template_table: list, title_box=None):
    """→ (значения меток, строки таблицы [[параметр, значение, источник]], переменные)."""
    t = item.get("тип", "")
    r = rules.get(t, {})
    v = variables(item, rules)
    name = item.get("наименование") or _fmt(r.get("наименование", ""), v)
    if r.get("наименование_без_размера", True):
        name = SIZE_TAIL.sub("", name).rstrip()
    article = item.get("артикул") or _fmt(r.get("артикул", ""), v)
    values = {
        "НАИМЕНОВАНИЕ": name,
        "АРТИКУЛ": article,
        "НАИМЕНОВАНИЕ_ТИТУЛ": title_lines(name, article, t, rules, *(title_box or (None, None))),
    }
    if item.get("номер_макета"):
        values["НОМЕР_МАКЕТА"] = str(item["номер_макета"])
    if item.get("серийный_номер"):
        values["СЕРИЙНЫЙ_НОМЕР"] = str(item["серийный_номер"])

    from_project = {"масса"} if item.get("масса") not in (None, "") else set()
    if item.get("мощность_Вт") not in (None, ""):
        from_project |= {"P_Вт", "P_max_кВт", "P_ср_кВт"}
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
