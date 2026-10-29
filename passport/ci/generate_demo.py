"""CI: заполняет паспорта по демо-проекту для всех изделий (без печатей) и печатает отчёт."""
import sys
from pathlib import Path
P = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(P))
import docx, engine, calc, project_import, book  # noqa: E402
try:
    import pdf_maket  # V2.x: макеты в PDF
except Exception:
    pdf_maket = None

rules = calc.load_rules(Path(sys.argv[2]) / "правила.json")
proj = project_import.load_project(P / "Демо" / "Демо_проект_Одинцово.json")
out = Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
tpl = {t: str(P / "Шаблоны" / f"Паспорт_{t.replace(' ', '_')}.docx")
       for t in ("Динамика МЦД", "Статика МЦД", "Статика МЦД круглый")}
lines = [f"Проект: {proj['проект']}, изделий {len(proj['изделия'])}; PDF-сборщики: {book.find_renderers() or 'нет'}"]
ok = True
for it in proj["изделия"]:
    t = it["тип"]; d = docx.Document(tpl[t]); table = engine.read_tech_table(d)
    values, rows, v = calc.compute(it, rules, table)
    keys = engine.find_placeholders(d)
    values["ГАРАНТИЯ"] = engine.format_warranty(36)
    values["ДАТА_ПРОДАЖИ"] = engine.format_sale_date(engine.parse_date("21.10.2026"))
    if "СЕРИЙНЫЙ_НОМЕР" in keys:
        values.setdefault("СЕРИЙНЫЙ_НОМЕР", "21102026010")
    imgs = {}
    if it.get("номер_макета") and "ИЗОБРАЖЕНИЕ_МАКЕТА" in keys:
        mk = P / "Демо" / "Макеты"
        pdf = mk / (it["номер_макета"] + ".pdf")
        if pdf_maket is not None and pdf.exists():          # PDF-макет → PNG через pypdfium2
            imgs["ИЗОБРАЖЕНИЕ_МАКЕТА"] = [str(pdf_maket.render_png(pdf_maket.make_spec(str(pdf), 1), out / "pdf"))]
            lines.append(f'PDF-макет {pdf.name}: отрисован → {Path(imgs["ИЗОБРАЖЕНИЕ_МАКЕТА"][0]).name}')
        else:
            imgs["ИЗОБРАЖЕНИЕ_МАКЕТА"] = [str(f) for f in sorted(mk.glob(it["номер_макета"] + "*.png"))][:it["сторон"]]
    fn = out / f'{it["префикс"]}.docx'
    rep = engine.generate(tpl[t], fn, values, [r[:2] for r in rows] if table else None, imgs, {"включено": False})
    good = not rep["остались незаполненными"] and rep["таблица"] and all(rep["изображения"].values())
    ok &= bool(good)
    lines.append(f'{"OK " if good else "ERR"} {it["префикс"]} {t} | {values["АРТИКУЛ"]} | строк {len(rows)} | '
                 f'макет {rep["изображения"] or "-"} | метки без значения {rep["остались незаполненными"] or "нет"} | {fn.stat().st_size} б')
print("\n".join(lines))
sys.exit(0 if ok else 1)
