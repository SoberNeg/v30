"""CI: заполняет паспорта по демо-проекту для всех изделий (без печатей) и печатает отчёт."""
import sys
from pathlib import Path
P = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(P))
import docx, engine, calc, project_import, book, docx_check  # noqa: E402

calc.FONT_FILE = str(P / "Шрифты" / "RussianRail_G_Pro_Regular.otf")   # замер титула — как в программе
rules = calc.load_rules(Path(sys.argv[2]) / "правила.json")
proj = project_import.load_project(P / "Демо" / "Демо_проект_Одинцово.json")
out = Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
tpl = {t: str(P / "Шаблоны" / f"Паспорт_{t.replace(' ', '_')}.docx")
       for t in ("Динамика МЦД", "Статика МЦД", "Статика МЦД круглый")}
lines = [f"Проект: {proj['проект']}, изделий {len(proj['изделия'])}; PDF-сборщики: {book.find_renderers() or 'нет'}"]
ok = True
for it in proj["изделия"]:
    t = it["тип"]; d = docx.Document(tpl[t]); table = engine.read_tech_table(d)
    values, rows, v = calc.compute(it, rules, table, engine.title_box(d))
    w_mm, pt = engine.title_box(d)
    fit = calc.title_fit(values["НАИМЕНОВАНИЕ_ТИТУЛ"], w_mm, pt)
    if fit < pt:
        values["НАИМЕНОВАНИЕ_ТИТУЛ"] = {"text": values["НАИМЕНОВАНИЕ_ТИТУЛ"], "pt": fit}
    keys = engine.find_placeholders(d)
    values["ГАРАНТИЯ"] = engine.format_warranty(36)
    values["ДАТА_ПРОДАЖИ"] = engine.format_sale_date(engine.parse_date("21.10.2026"))
    if "СЕРИЙНЫЙ_НОМЕР" in keys:
        values.setdefault("СЕРИЙНЫЙ_НОМЕР", "21102026010")
    imgs = {}
    if it.get("номер_макета") and "ИЗОБРАЖЕНИЕ_МАКЕТА" in keys:
        imgs["ИЗОБРАЖЕНИЕ_МАКЕТА"] = [str(f) for f in sorted((P / "Демо" / "Макеты").glob(it["номер_макета"] + "*.png"))][:it["сторон"]]
    fn = out / f'{it["префикс"]}.docx'
    rep = engine.generate(tpl[t], fn, values, [r[:2] for r in rows] if table else None, imgs, {"включено": False})
    good = not rep["остались незаполненными"] and rep["таблица"] and all(rep["изображения"].values())
    ok &= bool(good)
    xml = docx_check.check(fn)                 # Word не откроет файл с нарушенным порядком XML
    ok &= not xml
    lines.append(f'{"OK " if good and not xml else "ERR"} {it["префикс"]} {t} | {values["АРТИКУЛ"]} | строк {len(rows)} | '
                 f'макет {rep["изображения"] or "-"} | метки без значения {rep["остались незаполненными"] or "нет"} | {fn.stat().st_size} б'
                 + (f' | ПОРЯДОК XML: {len(xml)} ошибок — {xml[0][:120]}' if xml else ''))
print("\n".join(lines))
sys.exit(0 if ok else 1)
