"""Автотест интерфейса для GitHub Actions (Windows).

Запускает программу из исходника, выбирает МЦД в стартовом окне, добавляет
тестовые изделия, делает скриншоты вкладок и пишет отчёт о строках таблицы
«Комплектующие» (теги заливки). Использование:
    python ci/ui_screenshots.py <исходник.py> <light|dark> <папка_для_скриншотов>
"""
import importlib.util
import json
import os
import sys
import time
import traceback
from pathlib import Path

src, theme, out_dir = sys.argv[1], sys.argv[2], Path(sys.argv[3])
out_dir.mkdir(parents=True, exist_ok=True)
os.environ["LED_THEME"] = theme
report = {"theme": theme, "steps": [], "errors": []}

import tkinter.messagebox as mb  # noqa: E402
for _n in ("showinfo", "showwarning", "showerror"):
    setattr(mb, _n, lambda *a, **k: report["steps"].append(f"messagebox {a[:2]}") or "ok")
for _n in ("askyesno", "askokcancel", "askyesnocancel"):
    setattr(mb, _n, lambda *a, **k: False)

from PIL import ImageGrab  # noqa: E402

spec = importlib.util.spec_from_file_location("led_app", src)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

# программа может заменить messagebox.* своими окнами (V40.03) — запоминаем их
# для скриншота и снова ставим заглушки, чтобы тест не ждал нажатия кнопок
REAL_SHOWINFO = mb.showinfo
for _n in ("showinfo", "showwarning", "showerror"):
    setattr(mb, _n, lambda *a, **k: report["steps"].append(f"messagebox {a[:2]}") or "ok")
for _n in ("askyesno", "askokcancel", "askyesnocancel"):
    setattr(mb, _n, lambda *a, **k: False)

app = mod.ScreenDesigner()


def shot(name, widget=None):
    try:
        app.update()
        time.sleep(0.4)
        app.update()
        if widget is not None:
            x, y = widget.winfo_rootx(), widget.winfo_rooty()
            bbox = (x - 8, y - 32, x + widget.winfo_width() + 8, y + widget.winfo_height() + 8)
            img = ImageGrab.grab(bbox=bbox, all_screens=True)
        else:
            img = ImageGrab.grab(all_screens=True)
        img.save(out_dir / f"{theme}_{name}.png")
        report["steps"].append(f"shot {name}")
    except Exception:
        report["errors"].append(f"shot {name}: " + traceback.format_exc())


def step(fn, label):
    try:
        fn()
        report["steps"].append(f"ok {label}")
    except Exception:
        report["errors"].append(f"{label}: " + traceback.format_exc())


def chooser():
    wins = [w for w in app.winfo_children() if w.winfo_class() == "Toplevel" and w.title() == "Выбор проекта"]
    report["chooser_visible"] = bool(wins and wins[0].winfo_viewable())
    shot("01_startup_chooser")
    if wins:
        buttons = []
        stack = [wins[0]]
        while stack:
            w = stack.pop(0)
            stack.extend(w.winfo_children())
            if w.winfo_class() == "Button":
                buttons.append(w)
        logo = [b for b in buttons if str(b.cget("image"))]          # [РЖД, МЦД] — кнопки-логотипы
        mcd = [b for b in buttons if b.cget("text") == "МЦД"]
        (mcd or logo[1:2] or buttons[1:2])[0].invoke()
        report["brand"] = getattr(app, "project_brand_context", None)


def tab(text):
    for t in app.notebook.tabs():
        if app.notebook.tab(t, "text") == text:
            app.notebook.select(t)
            return
    raise RuntimeError("tab not found: " + text)


def add_items():
    # статические указатели МЦД двух размеров
    tab("Статические указатели")
    for w_, h_ in ((600, 400), (900, 300)):
        app._static_draft_mode = True
        app.static_width_mm, app.static_height_mm = w_, h_
        app._add_static_sign_to_project()
    # динамика МЦД
    tab("LED-экраны")
    app._dynamic_draft_mode = True
    app.width_mm, app.height_mm = 1780, 580
    app._add_to_project()
    report["static_signs"] = len(getattr(app, "saved_static_signs", []))
    report["screens"] = len(getattr(app, "saved_screens", []))


def components_report(label):
    tree = app.components_tree
    rows = []
    for parent in tree.get_children(""):
        rows.append({"group": tree.item(parent, "text"), "tags": list(tree.item(parent, "tags") or ())})
        for iid in tree.get_children(parent):
            vals = tree.item(iid, "values")
            rows.append({"src": tree.item(iid, "text"), "name": vals[0] if vals else "",
                         "tags": list(tree.item(iid, "tags") or ())})
    report[f"components_{label}"] = rows


def sequence():
    step(chooser, "chooser")
    app.after(1500, part2)


def part2():
    step(add_items, "add items")
    step(lambda: tab("LED-экраны"), "tab led")
    shot("02_led_screens", app)
    step(lambda: tab("Статические указатели"), "tab static")
    shot("03_static", app)
    step(lambda: tab("Журнал проекта"), "tab journal")
    shot("04_journal", app)
    step(lambda: tab("Комплектующие"), "tab components")
    step(lambda: (app.component_scope_var.set("Весь проект"), app._refresh_components_table()), "scope whole")
    step(lambda: components_report("whole"), "report whole")
    shot("05_components_whole", app)
    if hasattr(mod, "_v3090_set_sort"):
        step(lambda: mod._v3090_set_sort(app, "name"), "sort by name")
        step(lambda: components_report("whole_sorted"), "report sorted")
        shot("05b_components_sorted", app)
    step(lambda: (app.component_scope_var.set("Статика"), app._refresh_components_table()), "scope static")
    step(lambda: components_report("static"), "report static")
    shot("06_components_static", app)
    step(lambda: tab("Справочник"), "tab reference")
    shot("07_reference", app)
    def close_box():
        shot("09_message_box")
        for w in app.winfo_children():
            if w.winfo_class() == "Toplevel" and w.title() == "Проверка темы":
                w.destroy()
    if getattr(REAL_SHOWINFO, "__module__", "") != "tkinter.messagebox" or hasattr(mod, "_v4003_show"):
        app.after(900, close_box)
        step(lambda: REAL_SHOWINFO("Проверка темы", "Окно сообщения программы в текущей теме.", parent=app), "message box")
    if theme == "dark":
        # обратное переключение: всё должно вернуться к светлому виду
        step(lambda: (app.apply_theme("light"), tab("Комплектующие")), "switch to light")
        shot("08_after_switch_to_light", app)
    app.after(500, finish)


def finish():
    (out_dir / f"{theme}_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    app.destroy()


app.after(2500, sequence)
app.mainloop()
