# -*- coding: utf-8 -*-
"""Список изделий проекта: строки, по которым паспорт уже скачивали, подкрашены
полупрозрачным зелёным градиентом (слева насыщеннее, к правому краю исчезает).
Поддерживает светлую/тёмную тему и прокрутку колёсиком и тачпадом
(Windows, macOS, Linux)."""
import re
import sys
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

from PIL import Image, ImageTk

ROW_H = 26
HEAD_H = 26
GREEN = (46, 160, 67)          # цвет градиента
ALPHA_FROM, ALPHA_TO = 0.42, 0.04

PALETTES = {
    "light": {"bg": "#f3f4f6", "panel": "#ffffff", "fg": "#1d1d1f", "hint": "#666666",
              "row": "#ffffff", "row_alt": "#f8f9fb", "sel": "#cde1fa", "sel_outline": "#5b8fd6",
              "head_bg": "#eceff3", "head_fg": "#333333", "grid": "#e6e8ec", "head_sep": "#c8ccd2",
              "done_fg": "#1e6b30", "field": "#ffffff", "border": "#c8ccd2", "accent": "#2f6fd6",
              "empty": "#888888", "button": "#e9ebee", "button_active": "#dde1e6",
              "err": "#c0392b",
              "src": {"расчёт": "#1f5fbf", "проект": "#1e7b34", "шаблон": "#777777", "вручную": "#b35c00"}},
    "dark": {"bg": "#1f2124", "panel": "#2a2d31", "fg": "#e8eaed", "hint": "#9aa0a6",
             "row": "#2a2d31", "row_alt": "#30343a", "sel": "#2f4a72", "sel_outline": "#6c9de6",
             "head_bg": "#33373d", "head_fg": "#e8eaed", "grid": "#3a3e44", "head_sep": "#4a4f56",
             "done_fg": "#7ddc8f", "field": "#25282c", "border": "#4a4f56", "accent": "#5b8fd6",
             "empty": "#9aa0a6", "button": "#383c42", "button_active": "#454a51",
             "err": "#ff8a7a",
             "src": {"расчёт": "#7fb2ff", "проект": "#6fd88a", "шаблон": "#9aa0a6", "вручную": "#ffb35c"}},
}


def _rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _blend(base, color, a):
    return tuple(int(b + (c - b) * a) for b, c in zip(base, color))


def natural_key(v):
    """«Д2» < «Д10», «М-0457» < «М-1001»: числа внутри строки сравниваются как числа."""
    parts = re.split(r"(\d+)", str(v if v is not None else "").strip().lower().replace("ё", "е"))
    return [(0, int(p), "") if p.isdigit() else (1, 0, p) for p in parts if p != ""]


class GradientList(ttk.Frame):
    def __init__(self, master, columns, on_select=None, stretch=None, sort_keys=None, on_sort=None):
        """columns: [(ключ, заголовок, ширина, anchor)], stretch — ключ колонки,
        которая забирает свободную ширину. Номер строки (index) везде — номер в исходном
        списке rows (= номер изделия в проекте); порядок показа — self.order.
        sort_keys: {ключ колонки: функция(row) → ключ сравнения}; остальные колонки —
        естественная сортировка. on_sort(key, direction) — сохранить выбор."""
        super().__init__(master)
        self.columns = columns
        self.stretch = stretch
        self.on_select = on_select
        self.sort_keys = sort_keys or {}
        self.on_sort = on_sort
        self.sort_col, self.sort_dir = None, 0        # 1 — по возрастанию, −1 — по убыванию
        self.rows = []
        self.order = []
        self.selected = None
        self._img_cache = {}
        self._acc = 0.0
        self.pal = PALETTES["light"]
        self.font = tkfont.nametofont("TkDefaultFont")
        self.bold = self.font.copy()
        self.bold.configure(weight="bold")

        self.head = tk.Canvas(self, height=HEAD_H, highlightthickness=0)
        self.body = tk.Canvas(self, highlightthickness=0, takefocus=1, height=ROW_H * 5,
                              yscrollincrement=1)
        self.sb = ttk.Scrollbar(self, orient="vertical", command=self.body.yview)
        self.body.configure(yscrollcommand=self.sb.set)
        self.head.grid(row=0, column=0, sticky="ew")
        self.body.grid(row=1, column=0, sticky="nsew")
        self.sb.grid(row=0, column=1, rowspan=2, sticky="ns")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.empty_text = ""
        self.set_palette(self.pal)

        self.body.bind("<Configure>", lambda e: self.redraw())
        self.body.bind("<Button-1>", self._click)
        self.head.bind("<Button-1>", self._head_click)
        self.head.configure(cursor="hand2")
        self.body.bind("<Up>", lambda e: self._step(-1))
        self.body.bind("<Down>", lambda e: self._step(1))
        # Прокрутка: пока указатель над списком, колёсико/тачпад прокручивают именно его
        # (на Windows события колёсика иначе уходят виджету с фокусом).
        for w in (self.body, self.head):
            w.bind("<Enter>", self._grab_wheel)
            w.bind("<Leave>", self._release_wheel)

    # ------------------------------------------------------------- тема
    def set_palette(self, pal):
        self.pal = pal
        self._img_cache.clear()
        self.head.configure(bg=pal["head_bg"])
        self.body.configure(bg=pal["row"])
        self.redraw()

    # ------------------------------------------------------------- прокрутка
    def _grab_wheel(self, _=None):
        root = self.winfo_toplevel()
        root.bind_all("<MouseWheel>", self._wheel)
        root.bind_all("<Shift-MouseWheel>", lambda e: "break")
        root.bind_all("<Button-4>", lambda e: self._scroll_px(-ROW_H * 1.5))
        root.bind_all("<Button-5>", lambda e: self._scroll_px(ROW_H * 1.5))
        try:                                   # Tk 8.7 / 9: точная прокрутка тачпадом
            root.bind_all("<TouchpadScroll>", self._touchpad)
        except tk.TclError:
            pass

    def _release_wheel(self, _=None):
        root = self.winfo_toplevel()
        for ev in ("<MouseWheel>", "<Shift-MouseWheel>", "<Button-4>", "<Button-5>"):
            root.unbind_all(ev)
        try:
            root.unbind_all("<TouchpadScroll>")
        except tk.TclError:
            pass

    def _scroll_px(self, px):
        self._acc += px
        n = int(self._acc)
        if n:
            self._acc -= n
            self.body.yview_scroll(n, "units")       # 1 единица = 1 пиксель
        return "break"

    def _wheel(self, e):
        d = e.delta
        if sys.platform == "darwin" and abs(d) < 120:
            px = -d * 4                      # Tk 8.6 на Mac: тачпад даёт ±1…±10 за событие
        else:
            px = -d / 120.0 * ROW_H * 1.5    # Windows/Tk 9: 120 на «щелчок», тачпад — меньше
        return self._scroll_px(px)

    def _touchpad(self, e):
        d = int(e.delta)
        dy = d & 0xFFFF
        if dy >= 0x8000:
            dy -= 0x10000
        return self._scroll_px(-dy)

    # ------------------------------------------------------------- данные
    def set_rows(self, rows, keep_selection=True):
        self.rows = rows
        if not keep_selection or self.selected is None or self.selected >= len(rows):
            self.selected = None
        self._resort()
        self.redraw()

    def update_row(self, index, row):
        self.rows[index] = row
        self._resort()
        self.redraw()

    def select(self, index, notify=True):
        """index — номер в исходном списке (номер изделия)."""
        if not self.rows:
            return
        index = max(0, min(len(self.rows) - 1, index))
        self.selected = index
        self.redraw()
        self._see(index)
        if notify and self.on_select:
            self.on_select(index)

    # ------------------------------------------------------------- сортировка
    def set_sort(self, col, direction, notify=False):
        keys = [c[0] for c in self.columns]
        if col not in keys or direction not in (1, -1):
            col, direction = None, 0
        self.sort_col, self.sort_dir = col, direction
        self._resort()
        self.redraw()
        if self.selected is not None:
            self._see(self.selected)
        if notify and self.on_sort:
            self.on_sort(col, direction)

    def _resort(self):
        idx = list(range(len(self.rows)))
        if self.sort_col and self.sort_dir:
            f = self.sort_keys.get(self.sort_col) or (lambda r, k=self.sort_col: natural_key(r.get(k, "")))
            idx.sort(key=lambda i: (f(self.rows[i]), i), reverse=self.sort_dir < 0)
        self.order = idx

    def _head_click(self, e):
        x, w = 0, max(self.body.winfo_width(), 200)
        for (k, *_), cw in zip(self.columns, self._widths(w)):
            if x <= e.x < x + cw:
                if self.sort_col != k:
                    self.set_sort(k, 1, notify=True)
                elif self.sort_dir == 1:
                    self.set_sort(k, -1, notify=True)
                else:
                    self.set_sort(None, 0, notify=True)       # третий щелчок — порядок проекта
                return
            x += cw

    # ------------------------------------------------------------- рисование
    def _widths(self, total):
        fixed = sum(w for k, _, w, _ in self.columns if k != self.stretch)
        out = []
        for k, _, w, _ in self.columns:
            out.append(max(w, total - fixed) if k == self.stretch else w)
        return out

    def _row_image(self, width, selected):
        key = (width, selected)
        if key not in self._img_cache:
            base = _rgb(self.pal["sel"] if selected else self.pal["row"])
            img = Image.new("RGB", (max(width, 1), ROW_H))
            px = img.load()
            for x in range(img.width):
                a = ALPHA_FROM + (ALPHA_TO - ALPHA_FROM) * (x / max(1, img.width - 1))
                c = _blend(base, GREEN, a)
                for y in range(ROW_H):
                    px[x, y] = c
            self._img_cache[key] = ImageTk.PhotoImage(img)
        return self._img_cache[key]

    def redraw(self):
        p = self.pal
        w = max(self.body.winfo_width(), 200)
        widths = self._widths(w)
        self.head.delete("all")
        x = 0
        for (k, title, _, anchor), cw in zip(self.columns, widths):
            active = k == self.sort_col and self.sort_dir
            text = self._clip(title, cw - (26 if active else 10), self.bold)
            self.head.create_text(x + 6, HEAD_H // 2, text=text, anchor="w", font=self.bold,
                                  fill=p["accent"] if active else p["head_fg"])
            if active:                     # стрелка сортировки рисуется, а не символом шрифта
                ax = x + 6 + self.bold.measure(text) + 9
                cy = HEAD_H // 2
                pts = ((ax - 5, cy + 3, ax + 5, cy + 3, ax, cy - 4) if self.sort_dir > 0
                       else (ax - 5, cy - 3, ax + 5, cy - 3, ax, cy + 4))
                self.head.create_polygon(*pts, fill=p["accent"], outline="")
            x += cw
            self.head.create_line(x, 4, x, HEAD_H - 4, fill=p["head_sep"])
        b = self.body
        b.delete("all")
        if not self.rows:
            b.create_text(w // 2, 40, text=self.empty_text, fill=p["empty"], font=self.font,
                          width=w - 40, justify="center")
            b.configure(scrollregion=(0, 0, w, 80))
            return
        for pos, i in enumerate(self.order):
            row = self.rows[i]
            y = pos * ROW_H
            sel = i == self.selected
            if row.get("_done"):
                b.create_image(0, y, image=self._row_image(w, sel), anchor="nw")
            else:
                bg = p["sel"] if sel else (p["row"] if pos % 2 == 0 else p["row_alt"])
                b.create_rectangle(0, y, w, y + ROW_H, fill=bg, width=0)
            if sel:
                b.create_rectangle(1, y + 1, w - 2, y + ROW_H - 1, outline=p["sel_outline"])
            x = 0
            for (k, _, _, anchor), cw in zip(self.columns, widths):
                txt = str(row.get(k, ""))
                done = k == "_скачан" and row.get("_done")
                f = self.bold if done else self.font
                tx = x + 6 if anchor == "w" else x + cw // 2
                b.create_text(tx, y + ROW_H // 2, text=self._clip(txt, cw - 12, f), anchor=anchor,
                              font=f, fill=p["done_fg"] if done else p["fg"])
                x += cw
            b.create_line(0, y + ROW_H - 1, w, y + ROW_H - 1, fill=p["grid"])
        b.configure(scrollregion=(0, 0, w, len(self.rows) * ROW_H))

    def _clip(self, txt, width, f):
        if f.measure(txt) <= width:
            return txt
        while txt and f.measure(txt + "…") > width:
            txt = txt[:-1]
        return txt + "…"

    # ------------------------------------------------------------- события
    def _click(self, e):
        self.body.focus_set()
        pos = int(self.body.canvasy(e.y) // ROW_H)
        if 0 <= pos < len(self.order):
            self.select(self.order[pos])

    def _step(self, d):
        """↑/↓ — по видимому (отсортированному) порядку."""
        if not self.rows:
            return
        if self.selected is None or self.selected not in self.order:
            pos = 0 if d > 0 else len(self.order) - 1
        else:
            pos = max(0, min(len(self.order) - 1, self.order.index(self.selected) + d))
        self.select(self.order[pos])

    def _see(self, i):
        total = len(self.rows) * ROW_H
        if not total or i not in self.order:
            return
        top, bottom = (f * total for f in self.body.yview())
        y = self.order.index(i) * ROW_H
        if y < top:
            self.body.yview_moveto(y / total)
        elif y + ROW_H > bottom:
            self.body.yview_moveto((y + ROW_H - (bottom - top)) / total)


class ScrollFrame(ttk.Frame):
    """Рамка с вертикальной прокруткой: полоса появляется, только когда содержимое
    не помещается (маленький экран). Колёсико и тачпад работают, пока указатель внутри."""

    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0, yscrollincrement=1)
        self.sb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.sb.set)
        self.body = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window(0, 0, window=self.body, anchor="nw")
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self._acc = 0.0
        self._sb_shown = False
        self.body.bind("<Configure>", self._sync)
        self.canvas.bind("<Configure>", self._sync)
        for w in (self, self.canvas):
            w.bind("<Enter>", self._grab, add="+")
        self.bind("<Leave>", self._release)

    def set_palette(self, pal):
        self.canvas.configure(bg=pal["bg"])

    def fit_width(self):
        """Запросить у окна ширину по содержимому (как обычная рамка без прокрутки)."""
        self.body.update_idletasks()
        self.canvas.configure(width=self.body.winfo_reqwidth())
        self._sync()

    def _sync(self, _=None):
        cw = self.canvas.winfo_width()
        self.canvas.itemconfigure(self._win, width=max(cw, 1))
        bh = self.body.winfo_reqheight()
        self.canvas.configure(scrollregion=(0, 0, cw, bh))
        need = bh > self.canvas.winfo_height() + 1
        if need != self._sb_shown:
            self._sb_shown = need
            if need:
                self.sb.grid(row=0, column=1, sticky="ns")
            else:
                self.sb.grid_remove()
                self.canvas.yview_moveto(0)

    # прокрутка — те же правила, что и у списка изделий
    def _grab(self, _=None):
        root = self.winfo_toplevel()
        root.bind_all("<MouseWheel>", self._wheel)
        root.bind_all("<Button-4>", lambda e: self._scroll(-ROW_H * 1.5))
        root.bind_all("<Button-5>", lambda e: self._scroll(ROW_H * 1.5))
        try:
            root.bind_all("<TouchpadScroll>", self._touchpad)
        except tk.TclError:
            pass

    def _release(self, e=None):
        # <Leave> приходит и при переходе на дочерний виджет — проверяем, где указатель
        if e is not None:
            x, y = self.winfo_pointerxy()
            w = self.winfo_containing(x, y)
            while w is not None:
                if w is self:
                    return
                w = getattr(w, "master", None)
        root = self.winfo_toplevel()
        for ev in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            root.unbind_all(ev)
        try:
            root.unbind_all("<TouchpadScroll>")
        except tk.TclError:
            pass

    def _scroll(self, px):
        if not self._sb_shown:
            return "break"
        self._acc += px
        n = int(self._acc)
        if n:
            self._acc -= n
            self.canvas.yview_scroll(n, "units")
        return "break"

    def _wheel(self, e):
        d = e.delta
        if sys.platform == "darwin" and abs(d) < 120:
            return self._scroll(-d * 4)
        return self._scroll(-d / 120.0 * ROW_H * 1.5)

    def _touchpad(self, e):
        dy = int(e.delta) & 0xFFFF
        if dy >= 0x8000:
            dy -= 0x10000
        return self._scroll(-dy)


# ===================================================================== буфер обмена
# Ctrl+C / V / X / A (на Mac ⌘) при любой раскладке. Tk при русской раскладке присылает
# не «c», а «Cyrillic_es», и стандартные привязки <<Copy>> и т.п. не срабатывают.
# Клавишу узнаём по символу русской раскладки, по физическому коду клавиши (keycode)
# или по управляющему символу; обработчик стоит в теге «all» — после стандартных привязок
# поля, поэтому при английской раскладке ничего не повторяет.
INPUT_CLASSES = ("Entry", "TEntry", "Spinbox", "TSpinbox", "TCombobox", "Text")
SINGLE_LINE = ("Entry", "TEntry", "Spinbox", "TSpinbox", "TCombobox")
ACTIONS = {"c": "<<Copy>>", "v": "<<Paste>>", "x": "<<Cut>>", "a": "<<SelectAll>>",
           "z": "<<Undo>>", "y": "<<Redo>>"}
CYRILLIC = {"cyrillic_es": "c", "cyrillic_em": "v", "cyrillic_che": "x", "cyrillic_ef": "a",
            "cyrillic_ya": "z", "cyrillic_en": "y",
            "с": "c", "м": "v", "ч": "x", "ф": "a", "я": "z", "н": "y"}
KEYCODES = {
    "win": {67: "c", 86: "v", 88: "x", 65: "a", 90: "z", 89: "y"},        # Windows VK
    "x11": {54: "c", 55: "v", 53: "x", 38: "a", 52: "z", 29: "y"},        # Linux X11 (evdev)
    "mac": {8: "c", 9: "v", 7: "x", 0: "a", 6: "z", 16: "y"},             # Apple kVK_ANSI_*
}
CONTROL_CHARS = {"\x03": "c", "\x16": "v", "\x18": "x", "\x01": "a", "\x1a": "z", "\x19": "y"}


def _platform():
    if sys.platform.startswith("win"):
        return "win"
    if sys.platform == "darwin":
        return "mac"
    return "x11"


def key_letter(keysym, keycode, char="", platform=None):
    """Какая латинская клавиша нажата вместе с Ctrl/⌘ — независимо от раскладки."""
    platform = platform or _platform()
    ks = (keysym or "").lower()
    if len(ks) == 1 and "a" <= ks <= "z":
        return ks
    if ks in CYRILLIC:
        return CYRILLIC[ks]
    if char in CONTROL_CHARS:
        return CONTROL_CHARS[char]
    kc = int(keycode or 0)
    if platform == "mac":
        if kc > 0xFFFF:                       # Tk 8.6.11+: виртуальный код в битах 16–23,
            low = chr(kc & 0xFFFF).lower()    # символ — в младших 16
            return CYRILLIC.get(low) or KEYCODES["mac"].get((kc >> 16) & 0x7F)
        return KEYCODES["mac"].get(kc)
    return KEYCODES[platform].get(kc)


def _editable(w):
    try:
        if w.winfo_class().startswith("T") and w.winfo_class() != "Text":
            return not (w.instate(["readonly"]) or w.instate(["disabled"]))
        return str(w.cget("state")) not in ("readonly", "disabled")
    except Exception:
        return True


def select_all(w):
    try:
        if w.winfo_class() == "Text":
            w.tag_add("sel", "1.0", "end-1c")
            w.mark_set("insert", "end-1c")
        else:
            w.selection_range(0, "end")
            w.icursor("end")
    except Exception:
        pass
    return "break"


def _paste_single_line(e):
    """Вставка в однострочное поле: переводы строк и табуляции → пробелы."""
    w = e.widget
    if not _editable(w):
        return "break"
    try:
        txt = w.clipboard_get()
    except tk.TclError:
        return "break"
    txt = re.sub(r"[ \t]*[\r\n]+[ \t]*", " ", txt).replace("\t", " ").strip()
    try:
        w.delete("sel.first", "sel.last")
    except tk.TclError:
        pass
    w.insert("insert", txt)
    try:
        w.xview_moveto(1.0) if w.index("insert") >= w.index("end") else None
    except Exception:
        pass
    return "break"


def _on_shortcut(e):
    w = e.widget
    if not isinstance(w, tk.Misc) or w.winfo_class() not in INPUT_CLASSES:
        return None
    ks = e.keysym or ""
    latin = len(ks) == 1 and ks.lower() in ACTIONS
    letter = key_letter(ks, e.keycode, getattr(e, "char", ""))
    if letter not in ACTIONS:
        return None
    if letter == "a":                      # «выделить всё» — всегда сами (в X11 Ctrl+A = начало строки)
        return select_all(w)
    if latin:                              # латиница: стандартная привязка уже сработала
        return None
    ev = ACTIONS[letter]
    if ev in ("<<Undo>>", "<<Redo>>") and w.winfo_class() != "Text":
        return None
    if ev in ("<<Paste>>", "<<Cut>>") and not _editable(w):
        return "break"
    try:
        w.event_generate(ev)
    except tk.TclError:
        pass
    return "break"


def _context_menu(e):
    w = e.widget
    if not isinstance(w, tk.Misc) or w.winfo_class() not in INPUT_CLASSES:
        return None
    try:
        w.focus_set()
    except Exception:
        pass
    mod = "⌘" if _platform() == "mac" else "Ctrl+"
    m = tk.Menu(w, tearoff=0)
    edit = _editable(w)
    if edit:
        m.add_command(label="Вырезать", accelerator=mod + "X", command=lambda: w.event_generate("<<Cut>>"))
    m.add_command(label="Копировать", accelerator=mod + "C", command=lambda: w.event_generate("<<Copy>>"))
    if edit:
        m.add_command(label="Вставить", accelerator=mod + "V", command=lambda: w.event_generate("<<Paste>>"))
    m.add_separator()
    m.add_command(label="Выделить всё", accelerator=mod + "A", command=lambda: select_all(w))
    try:
        m.tk_popup(e.x_root, e.y_root)
    finally:
        m.grab_release()
    return "break"


def install_clipboard(root):
    """Подключить один раз для всего приложения (все поля, в том числе в диалогах)."""
    mods = ["Command"] if _platform() == "mac" else ["Control"]   # на Mac Ctrl+A — начало строки
    for mod in mods:
        try:
            root.bind_all(f"<{mod}-KeyPress>", _on_shortcut, add="+")
        except tk.TclError:
            pass
    for cls in SINGLE_LINE:
        root.bind_class(cls, "<<Paste>>", _paste_single_line)
    buttons = ["<Button-3>"] if _platform() != "mac" else ["<Button-2>", "<Button-3>", "<Control-Button-1>"]
    for b in buttons:
        root.bind_all(b, _context_menu, add="+")
