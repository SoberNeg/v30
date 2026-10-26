# -*- coding: utf-8 -*-
"""Список изделий проекта: строки, по которым паспорт уже скачивали, подкрашены
полупрозрачным зелёным градиентом (слева насыщеннее, к правому краю исчезает).
Поддерживает светлую/тёмную тему и прокрутку колёсиком и тачпадом
(Windows, macOS, Linux)."""
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
              "src": {"расчёт": "#1f5fbf", "проект": "#1e7b34", "шаблон": "#777777", "вручную": "#b35c00"}},
    "dark": {"bg": "#1f2124", "panel": "#2a2d31", "fg": "#e8eaed", "hint": "#9aa0a6",
             "row": "#2a2d31", "row_alt": "#30343a", "sel": "#2f4a72", "sel_outline": "#6c9de6",
             "head_bg": "#33373d", "head_fg": "#e8eaed", "grid": "#3a3e44", "head_sep": "#4a4f56",
             "done_fg": "#7ddc8f", "field": "#25282c", "border": "#4a4f56", "accent": "#5b8fd6",
             "empty": "#9aa0a6", "button": "#383c42", "button_active": "#454a51",
             "src": {"расчёт": "#7fb2ff", "проект": "#6fd88a", "шаблон": "#9aa0a6", "вручную": "#ffb35c"}},
}


def _rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _blend(base, color, a):
    return tuple(int(b + (c - b) * a) for b, c in zip(base, color))


class GradientList(ttk.Frame):
    def __init__(self, master, columns, on_select=None, stretch=None):
        """columns: [(ключ, заголовок, ширина, anchor)], stretch — ключ колонки,
        которая забирает свободную ширину."""
        super().__init__(master)
        self.columns = columns
        self.stretch = stretch
        self.on_select = on_select
        self.rows = []
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
        self.redraw()

    def update_row(self, index, row):
        self.rows[index] = row
        self.redraw()

    def select(self, index, notify=True):
        if not self.rows:
            return
        index = max(0, min(len(self.rows) - 1, index))
        self.selected = index
        self.redraw()
        self._see(index)
        if notify and self.on_select:
            self.on_select(index)

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
            self.head.create_text(x + 6, HEAD_H // 2, text=title, anchor="w", font=self.bold,
                                  fill=p["head_fg"])
            x += cw
            self.head.create_line(x, 4, x, HEAD_H - 4, fill=p["head_sep"])
        b = self.body
        b.delete("all")
        if not self.rows:
            b.create_text(w // 2, 40, text=self.empty_text, fill=p["empty"], font=self.font,
                          width=w - 40, justify="center")
            b.configure(scrollregion=(0, 0, w, 80))
            return
        for i, row in enumerate(self.rows):
            y = i * ROW_H
            sel = i == self.selected
            if row.get("_done"):
                b.create_image(0, y, image=self._row_image(w, sel), anchor="nw")
            else:
                bg = p["sel"] if sel else (p["row"] if i % 2 == 0 else p["row_alt"])
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
        i = int(self.body.canvasy(e.y) // ROW_H)
        if 0 <= i < len(self.rows):
            self.select(i)

    def _step(self, d):
        if self.rows:
            self.select((self.selected if self.selected is not None else -1) + d)

    def _see(self, i):
        total = len(self.rows) * ROW_H
        if not total:
            return
        top, bottom = (f * total for f in self.body.yview())
        y = i * ROW_H
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
