# -*- coding: utf-8 -*-
"""Список изделий проекта: строки, по которым паспорт уже скачивали, подкрашены
полупрозрачным зелёным градиентом (слева насыщеннее, к правому краю исчезает)."""
import sys
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

from PIL import Image, ImageTk

ROW_H = 26
HEAD_H = 26
WHITE = (255, 255, 255)
SEL = (205, 225, 250)          # выделенная строка
GREEN = (46, 160, 67)          # цвет градиента
ALPHA_FROM, ALPHA_TO = 0.42, 0.04


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
        self.font = tkfont.nametofont("TkDefaultFont")
        self.bold = self.font.copy()
        self.bold.configure(weight="bold")

        self.head = tk.Canvas(self, height=HEAD_H, highlightthickness=0, bg="#eceff3")
        self.body = tk.Canvas(self, highlightthickness=0, bg="white", takefocus=1, height=ROW_H * 5)
        self.sb = ttk.Scrollbar(self, orient="vertical", command=self.body.yview)
        self.body.configure(yscrollcommand=self.sb.set)
        self.head.grid(row=0, column=0, sticky="ew")
        self.body.grid(row=1, column=0, sticky="nsew")
        self.sb.grid(row=0, column=1, rowspan=2, sticky="ns")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.empty_text = ""

        self.body.bind("<Configure>", lambda e: self.redraw())
        self.body.bind("<Button-1>", self._click)
        self.body.bind("<Up>", lambda e: self._step(-1))
        self.body.bind("<Down>", lambda e: self._step(1))
        if sys.platform.startswith("linux"):
            self.body.bind("<Button-4>", lambda e: self.body.yview_scroll(-1, "units"))
            self.body.bind("<Button-5>", lambda e: self.body.yview_scroll(1, "units"))
        else:
            self.body.bind("<MouseWheel>", self._wheel)

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
            base = SEL if selected else WHITE
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
        w = max(self.body.winfo_width(), 200)
        widths = self._widths(w)
        # шапка
        self.head.delete("all")
        x = 0
        for (k, title, _, anchor), cw in zip(self.columns, widths):
            self.head.create_text(x + 6, HEAD_H // 2, text=title, anchor="w", font=self.bold,
                                  fill="#333")
            x += cw
            self.head.create_line(x, 4, x, HEAD_H - 4, fill="#c8ccd2")
        # строки
        b = self.body
        b.delete("all")
        if not self.rows:
            b.create_text(w // 2, 40, text=self.empty_text, fill="#888", font=self.font,
                          width=w - 40, justify="center")
            b.configure(scrollregion=(0, 0, w, 80))
            return
        for i, row in enumerate(self.rows):
            y = i * ROW_H
            sel = i == self.selected
            if row.get("_done"):
                b.create_image(0, y, image=self._row_image(w, sel), anchor="nw")
            else:
                bg = "#%02x%02x%02x" % (SEL if sel else (WHITE if i % 2 == 0 else (248, 249, 251)))
                b.create_rectangle(0, y, w, y + ROW_H, fill=bg, width=0)
            if sel:
                b.create_rectangle(1, y + 1, w - 2, y + ROW_H - 1, outline="#5b8fd6")
            x = 0
            for (k, _, _, anchor), cw in zip(self.columns, widths):
                txt = str(row.get(k, ""))
                f = self.bold if k == "_скачан" and row.get("_done") else self.font
                color = "#1e6b30" if k == "_скачан" and row.get("_done") else "#1d1d1f"
                tx = x + 6 if anchor == "w" else x + cw // 2
                b.create_text(tx, y + ROW_H // 2, text=self._clip(txt, cw - 12, f), anchor=anchor,
                              font=f, fill=color)
                x += cw
            b.create_line(0, y + ROW_H - 1, w, y + ROW_H - 1, fill="#e6e8ec")
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

    def _wheel(self, e):
        delta = e.delta if sys.platform == "darwin" else e.delta // 120
        self.body.yview_scroll(-delta, "units")

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
