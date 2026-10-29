"""Перенос исправлений Windows-сборки на новую версию автора.
usage: python3 port_fixes.py <src.py> <out.py> <VER e.g. 30.89>"""
import difflib, re, subprocess, sys, shutil
from pathlib import Path
src, out, ver = sys.argv[1], Path(sys.argv[2]), sys.argv[3]
S = Path(__file__).parent
shutil.copy(src, out)
r = subprocess.run(["patch", "-p0", "--no-backup-if-mismatch", "-r", "-", str(out)], stdin=open(S/"my_fixes.patch"), capture_output=True, text=True)
print(r.stdout.strip())
s = out.read_text(encoding="utf-8")
s = re.sub(r'APP_TITLE = "Конструктор LED-экранов V[0-9.]+"', f'APP_TITLE = "Конструктор LED-экранов V{ver}.1"', s, count=1)
s = s.replace("# V30.74 THEME LIGHT/DARK + COMPONENTS STATIC FILL FIX\n# V30.73 WINDOWS STARTUP WINDOW FIX (РЖД/МЦД после mainloop)\n",
    f"# V{ver}.1 = V{ver} + перенесённые исправления сборки Windows:\n#   стартовое окно РЖД/МЦД после mainloop; тема «Светлая / Тёмная»;\n#   заливка статических строк в «Комплектующих»\n", 1)
orig = Path(src).read_text(encoding="utf-8").splitlines(keepends=True); work = s.splitlines(keepends=True)
res = []
for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, orig, work, autojunk=False).get_opcodes():
    seg = work[j1:j2]
    if tag in ("insert", "replace"):
        seg = [re.sub(r"\bV30\.7[34]\b", f"V{ver}.1", l).replace("_v3074_", "_uitheme_").replace("_V3074_", "_UITHEME_") for l in seg]
    res.extend(seg)
s = "".join(res); out.write_text(s, encoding="utf-8")
checks = {f'APP_TITLE = "Конструктор LED-экранов V{ver}.1"': 1, "self.after(150, self._run_startup_choice)": 1,
          "def _log_startup_dialog_error": 1, "_components_restripe_static()": 1, '"static_all":True': 1,
          'static_gray_1", background="#ececec"': 1, "ScreenDesigner._build_ui = _uitheme_build_ui": 1,
          "_v3074_state": 0, "startup_action = self._show_startup_choice()\n        except Exception as exc:\n            # V20.53: не выбираем МЦД молча при сбое стартового окна.\n            # Если": 0}
ok = True
for k, n in checks.items():
    c = s.count(k); ok &= (c == n); print(("OK " if c == n else "FAIL"), c, k.splitlines()[0][:70])
i = s.index(f"# V{ver}.1 — ТЕМА ПРОГРАММЫ"); j = s.index('if __name__ == "__main__":')
tail = s[s.index("ScreenDesigner.apply_theme"):j]
ok &= i < j and tail.count("\n") <= 4 and "ScreenDesigner._build_ui =" not in s[i + 10:j].split("ScreenDesigner._build_ui = _uitheme_build_ui")[1]
removed = [l for l in difflib.unified_diff(orig, s.splitlines(keepends=True), n=0) if l.startswith("-") and not l.startswith("---")]
print("removed lines:", len(removed)); [print("   ", l.rstrip()[:110]) for l in removed]
compile(s, str(out), "exec"); print("SYNTAX_OK"); print("ALL_OK" if ok else "CHECK_FAILED")

# ---- совместимость с V30.90+ («Комплектующие»: сортировка по столбцам, плашки статуса/УПД)
s = out.read_text(encoding="utf-8")
old = '''    try:
        cls = w.winfo_class()
    except Exception:
        return
    if cls in _UITHEME_TK_CLASSES:'''
new = '''    try:
        cls = w.winfo_class()
        # плашки-надписи поверх ячеек таблицы (статус, УПД) сохраняют свои цвета
        if cls == "Label" and w.master is not None and w.master.winfo_class() == "Treeview":
            return
    except Exception:
        return
    if cls in _UITHEME_TK_CLASSES:'''
assert s.count(old) == 1, "apply_widget anchor"
s = s.replace(old, new)
if "def _v3090_apply_sort(self):" in s:
    compat = f'''

# V{ver}.1 — после сортировки по столбцу (V30.90) серая градация статических строк
# раздаётся заново в новом порядке строк.
_uitheme_orig_v3090_apply_sort = _v3090_apply_sort


def _v3090_apply_sort(self):
    _uitheme_orig_v3090_apply_sort(self)
    try:
        self._components_restripe_static()
    except Exception:
        pass
'''
    anchor = '\n\nif __name__ == "__main__":'
    assert s.count(anchor) == 1
    s = s.replace(anchor, compat.rstrip() + "\n" + anchor)
    print("V30.90 sort compat: added")
compile(s, str(out), "exec"); out.write_text(s, encoding="utf-8"); print("compat OK")

# ---- совместимость с V30.94+ (центрирование вкладок хранится в стилях ttk, свои у каждой темы)
s = out.read_text(encoding="utf-8")
if "def _v3094_center_notebook(notebook):" in s:
    compat = f'''

# V{ver}.1 — после смены темы заново центрируем вкладки (V30.94): отступ хранится
# в стиле ttk, а у светлой и тёмной темы стили раздельные.
_uitheme_orig_apply_theme_v3094 = _uitheme_apply_theme


def _uitheme_apply_theme(app, theme, save=False):
    _uitheme_orig_apply_theme_v3094(app, theme, save)
    try:
        for nb in _v3094_collect_notebooks(app):
            nb.after(80, lambda nb=nb: _v3094_center_notebook(nb))
    except Exception:
        pass
'''
    anchor = '\n\nif __name__ == "__main__":'
    assert s.count(anchor) == 1
    s = s.replace(anchor, compat.rstrip() + "\n" + anchor)
    compile(s, str(out), "exec"); out.write_text(s, encoding="utf-8"); print("V30.94 tabs compat: added")

# ---- тёмная тема: «бледный» текст (светлее основного в светлой теме) остаётся бледным,
#      т.е. ближе к тёмному фону; иначе пары «с расчётом / без расчёта» (V40.25) сливаются
s = out.read_text(encoding="utf-8")
old = '''            out = _uitheme_hex(*_uitheme_colorsys.hls_to_rgb(h, nl, ns))
    _uitheme_color_cache[key] = out'''
new = '''            out = _uitheme_hex(*_uitheme_colorsys.hls_to_rgb(h, nl, ns))
        elif l < 0.93:                      # бледный текст (ближе к белому фону) -> ближе к тёмному фону
            nl = max(0.42, 1.12 - l)
            out = _uitheme_hex(*_uitheme_colorsys.hls_to_rgb(h, nl, s))
    _uitheme_color_cache[key] = out'''
assert s.count(old) == 1, "dark fg anchor"
s = s.replace(old, new)
compile(s, str(out), "exec"); out.write_text(s, encoding="utf-8"); print("pale-text rule: added")

# ---- вес DIN-рейки (решение автора 29.10.2026): 20 г — это отрезок 10 см → 200 г/м.п.
s = out.read_text(encoding="utf-8")
if "'DIN-рейка': '20'," in s:
    s = s.replace("'DIN-рейка': '20',", "'DIN-рейка': '200 гр/м.п.',   # 20 г = отрезок 10 см", 1)
    assert s.count('"DIN-рейка":20,') == 1
    s = s.replace('"DIN-рейка":20,', '"DIN-рейка":200,', 1)   # МЦД: количество в м.п. → г/м.п.
    fix = f'''

# V{ver}.1 — вес DIN-рейки 200 г/м.п. (раньше «20» без единицы = 2 г на отрезок 0,1 м.п.).
# Вес кабинета РЖД (fixed_grams) считался по старому весу: +0,1 м.п. × (200 − 20) г = +18 г.
for _cab_size in list(CABINET_WEIGHT_DATA_V19_71):
    if COMPONENT_RATES.get(_cab_size, {{}}).get("DIN-рейка"):
        CABINET_WEIGHT_DATA_V19_71[_cab_size]["fixed_grams"] += COMPONENT_RATES[_cab_size]["DIN-рейка"] * (200.0 - 20.0)
'''
    anchor = "\n# Миграция старых имён комплектующих в пользовательских .ledproj."
    assert s.count(anchor) == 1
    s = s.replace(anchor, fix.rstrip() + "\n" + anchor, 1)
    compile(s, str(out), "exec"); out.write_text(s, encoding="utf-8"); print("DIN weight 200 г/м.п.: applied")
