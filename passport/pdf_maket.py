# -*- coding: utf-8 -*-
"""Макеты в PDF: предпросмотр, обрезка мышью, рендер выбранной области в PNG 300 dpi.

Рамка обрезки хранится в пунктах PDF (1/72 дюйма), начало координат — левый верхний
угол страницы так, как она видна на экране: (x0, y0, x1, y1). None — весь лист.
Описание макета в поле окна («спецификация») — строка:
    <путь к PDF>::стр=<N>[::рамка=x0,y0,x1,y1]
для обычных картинок — просто путь к файлу."""
import hashlib
import re
from pathlib import Path

try:
    import pypdfium2 as pdfium
except Exception:                              # нет пакета — PNG/JPG работают как раньше
    pdfium = None

SPEC_RE = re.compile(r"^(?P<path>.+?)::стр=(?P<page>\d+)(?:::рамка=(?P<box>[-\d.,]+))?$")
DPI = 300
MAX_PX = 6000
MM = 25.4 / 72


def available():
    return pdfium is not None


def is_pdf(path):
    return str(path).lower().endswith(".pdf") or bool(SPEC_RE.match(str(path)))


def make_spec(path, page=1, box=None):
    s = f"{path}::стр={int(page)}"
    if box:
        s += "::рамка=" + ",".join(f"{v:.2f}" for v in box)
    return s


def parse_spec(spec):
    """→ (путь, страница с 1, рамка или None); для картинки — (путь, None, None)."""
    m = SPEC_RE.match(str(spec).strip())
    if not m:
        p = str(spec).strip()
        return (p, 1, None) if p.lower().endswith(".pdf") else (p, None, None)
    box = None
    if m.group("box"):
        try:
            box = tuple(float(x) for x in m.group("box").split(","))[:4]
            if len(box) != 4:
                box = None
        except ValueError:
            box = None
    return m.group("path"), int(m.group("page")), box


def describe(spec):
    """Текст для поля окна: «М-0457.pdf, стр. 1, обрезка» / путь картинки."""
    path, page, box = parse_spec(spec)
    if page is None:
        return Path(path).name
    return f"{Path(path).name}, стр. {page}, " + ("обрезка" if box else "весь лист")


def _doc(path):
    if pdfium is None:
        raise RuntimeError("Для макетов в PDF нужен пакет pypdfium2.\n"
                           "Установите: pip install pypdfium2 (Запуск_Windows.bat ставит его сам).")
    return pdfium.PdfDocument(str(path))


def page_count(path):
    doc = _doc(path)
    try:
        return len(doc)
    finally:
        doc.close()


def page_size(path, page=1):
    """Размер страницы в пунктах (как видна на экране, с учётом поворота)."""
    doc = _doc(path)
    try:
        w, h = doc[page - 1].get_size()          # PDFium уже учитывает поворот страницы
        return float(w), float(h)
    finally:
        doc.close()


def _render(path, page, box, scale):
    doc = _doc(path)
    try:
        pg = doc[page - 1]
        w, h = pg.get_size()
        crop = (0, 0, 0, 0)
        if box:
            x0, y0, x1, y1 = box
            crop = (max(0.0, x0), max(0.0, h - y1), max(0.0, w - x1), max(0.0, y0))  # л, н, п, в
        bmp = pg.render(scale=scale, crop=crop, rotation=0, may_draw_forms=True)
        return bmp.to_pil().convert("RGB")
    finally:
        doc.close()


def preview(path, page, max_w, max_h):
    """Вся страница под размер окна → (PIL-картинка, пикселей на пункт)."""
    w, h = page_size(path, page)
    scale = max(0.05, min(max_w / w, max_h / h))
    return _render(path, page, None, scale), scale


def content_box(path, page, margin_pt=0.0):
    """Рамка по содержимому: всё, что не белое (белые поля обрезаются)."""
    from PIL import ImageChops, Image
    w, h = page_size(path, page)
    scale = min(150 / 72, 3000 / max(w, h))
    im = _render(path, page, None, scale)
    diff = ImageChops.difference(im, Image.new("RGB", im.size, (255, 255, 255))).convert("L")
    bbox = diff.point(lambda v: 255 if v > 12 else 0).getbbox()
    if not bbox:
        return None
    x0, y0, x1, y1 = (v / scale for v in bbox)
    return (max(0.0, x0 - margin_pt), max(0.0, y0 - margin_pt),
            min(w, x1 + margin_pt), min(h, y1 + margin_pt))


def render_png(spec, out_dir):
    """PNG выбранной области (300 dpi, длинная сторона ≤ 6000 px) в out_dir.
    Повторный вызов с той же обрезкой берёт готовый файл."""
    path, page, box = parse_spec(spec)
    src = Path(path)
    if not src.exists():
        raise FileNotFoundError(path)
    w, h = page_size(path, page)
    x0, y0, x1, y1 = box or (0, 0, w, h)
    bw, bh = max(1.0, x1 - x0), max(1.0, y1 - y0)
    scale = min(DPI / 72, MAX_PX / max(bw, bh))
    sig = hashlib.md5(f"{src.resolve()}|{src.stat().st_mtime_ns}|{page}|{box}|{scale:.4f}".encode()).hexdigest()[:8]
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"{src.stem}_стр{page}_{sig}.png"
    if not out.exists():
        im = _render(path, page, box, scale)
        im.save(out, dpi=(round(scale * 72), round(scale * 72)))
    return str(out)


# ===================================================================== окно обрезки
def crop_dialog(master, path, page=1, box=None, title="Обрезка макета"):
    """Окно предпросмотра PDF с рамкой обрезки. → (страница, рамка|None) или None (отмена)."""
    import tkinter as tk
    from tkinter import ttk, messagebox
    from PIL import Image, ImageTk, ImageEnhance

    try:
        n_pages = page_count(path)
    except Exception as e:
        messagebox.showerror(title, str(e), parent=master)
        return None
    page = max(1, min(n_pages, int(page or 1)))
    top = tk.Toplevel(master)
    top.title(f"{title} — {Path(path).name}")
    top.transient(master)
    sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
    cw, ch = min(900, sw - 160), min(640, sh - 260)
    st = {"page": page, "box": box, "img": None, "dim": None, "scale": 1.0, "tk": None,
          "drag": None, "result": None, "size": (1, 1)}

    bar = ttk.Frame(top, padding=(10, 8, 10, 4))
    bar.pack(fill="x")
    ttk.Label(bar, text="Страница:").pack(side="left")
    prev_b = ttk.Button(bar, text="‹", width=3, command=lambda: go(-1))
    prev_b.pack(side="left", padx=(6, 2))
    page_lbl = ttk.Label(bar, width=10, anchor="center")
    page_lbl.pack(side="left")
    next_b = ttk.Button(bar, text="›", width=3, command=lambda: go(1))
    next_b.pack(side="left", padx=(2, 12))
    ttk.Button(bar, text="Весь лист", command=lambda: set_box(None)).pack(side="left")
    ttk.Button(bar, text="По содержимому", command=lambda: auto()).pack(side="left", padx=6)
    info = ttk.Label(bar, anchor="e")
    info.pack(side="right")

    cv = tk.Canvas(top, width=cw, height=ch, highlightthickness=0, bg="#5a5d63", cursor="crosshair")
    cv.pack(fill="both", expand=True, padx=10)
    ttk.Label(top, text="Нарисуйте рамку мышью; тяните за углы и стороны или передвиньте её целиком.",
              style="Hint.TLabel").pack(anchor="w", padx=12, pady=(4, 0))
    bb = ttk.Frame(top, padding=10)
    bb.pack(fill="x")
    ttk.Button(bb, text="Отмена", command=lambda: close(None)).pack(side="right")
    ttk.Button(bb, text="Готово", command=lambda: close(True)).pack(side="right", padx=8)

    def load():
        try:
            img, scale = preview(path, st["page"], cw - 20, ch - 20)
        except Exception as e:
            messagebox.showerror(title, str(e), parent=top)
            return
        st["img"], st["scale"] = img, scale
        st["dim"] = ImageEnhance.Brightness(img).enhance(0.45)
        st["size"] = page_size(path, st["page"])
        page_lbl.configure(text=f"{st['page']} из {n_pages}")
        prev_b.state(["!disabled"] if st["page"] > 1 else ["disabled"])
        next_b.state(["!disabled"] if st["page"] < n_pages else ["disabled"])
        draw()

    def origin():
        img = st["img"]
        return (max(0, (cv.winfo_width() - img.width) // 2), max(0, (cv.winfo_height() - img.height) // 2))

    def box_px():
        w, h = st["size"]
        x0, y0, x1, y1 = st["box"] or (0, 0, w, h)
        s = st["scale"]
        ox, oy = origin()
        return ox + x0 * s, oy + y0 * s, ox + x1 * s, oy + y1 * s

    def draw():
        if st["img"] is None:
            return
        img = st["img"]
        if st["box"]:
            comp = st["dim"].copy()
            ox, oy = origin()
            x0, y0, x1, y1 = (int(round(v)) for v in box_px())
            reg = (max(0, x0 - ox), max(0, y0 - oy), min(img.width, x1 - ox), min(img.height, y1 - oy))
            if reg[2] > reg[0] and reg[3] > reg[1]:
                comp.paste(img.crop(reg), reg[:2])
        else:
            comp = img
        st["tk"] = ImageTk.PhotoImage(comp)
        cv.delete("all")
        ox, oy = origin()
        cv.create_image(ox, oy, image=st["tk"], anchor="nw")
        x0, y0, x1, y1 = box_px()
        if st["box"]:
            cv.create_rectangle(x0, y0, x1, y1, outline="#ffffff", width=2)
            cv.create_rectangle(x0, y0, x1, y1, outline="#2f6fd6", width=1, dash=(4, 3))
            for hx, hy in handles():
                cv.create_rectangle(hx - 4, hy - 4, hx + 4, hy + 4, fill="#2f6fd6", outline="#ffffff")
        w, h = st["size"]
        bx = st["box"] or (0, 0, w, h)
        info.configure(text=f"{'Обрезка' if st['box'] else 'Весь лист'}: "
                            f"{(bx[2] - bx[0]) * MM:.0f}×{(bx[3] - bx[1]) * MM:.0f} мм")

    def handles():
        x0, y0, x1, y1 = box_px()
        xm, ym = (x0 + x1) / 2, (y0 + y1) / 2
        return [(x0, y0), (xm, y0), (x1, y0), (x1, ym), (x1, y1), (xm, y1), (x0, y1), (x0, ym)]

    def to_pt(x, y):
        ox, oy = origin()
        w, h = st["size"]
        s = st["scale"]
        return min(w, max(0.0, (x - ox) / s)), min(h, max(0.0, (y - oy) / s))

    def hit(x, y):
        if not st["box"]:
            return None
        names = ["nw", "n", "ne", "e", "se", "s", "sw", "w"]
        for nm, (hx, hy) in zip(names, handles()):
            if abs(x - hx) <= 7 and abs(y - hy) <= 7:
                return nm
        x0, y0, x1, y1 = box_px()
        if x0 < x < x1 and y0 < y < y1:
            return "move"
        return None

    def press(e):
        mode = hit(e.x, e.y)
        px, py = to_pt(e.x, e.y)
        if mode is None:
            st["box"] = (px, py, px, py)
            mode = "se"
        st["drag"] = (mode, px, py, st["box"])

    def motion(e):
        if not st["drag"]:
            try:
                cv.configure(cursor={"move": "fleur", None: "crosshair"}.get(hit(e.x, e.y), "hand2"))
            except tk.TclError:
                pass
            return
        mode, sx, sy, b0 = st["drag"]
        px, py = to_pt(e.x, e.y)
        w, h = st["size"]
        x0, y0, x1, y1 = b0
        if mode == "move":
            dx = min(max(px - sx, -x0), w - x1)
            dy = min(max(py - sy, -y0), h - y1)
            st["box"] = (x0 + dx, y0 + dy, x1 + dx, y1 + dy)
        else:
            if "w" in mode:
                x0 = px
            if "e" in mode:
                x1 = px
            if mode.startswith("n"):
                y0 = py
            if mode.startswith("s"):
                y1 = py
            st["box"] = (x0, y0, x1, y1)
        draw()

    def release(_e):
        st["drag"] = None
        if st["box"]:
            x0, y0, x1, y1 = st["box"]
            x0, x1 = sorted((x0, x1))
            y0, y1 = sorted((y0, y1))
            w, h = st["size"]
            if x1 - x0 < 3 or y1 - y0 < 3:            # щелчок без рамки — весь лист
                st["box"] = None
            elif x0 <= 0.5 and y0 <= 0.5 and x1 >= w - 0.5 and y1 >= h - 0.5:
                st["box"] = None
            else:
                st["box"] = (x0, y0, x1, y1)
        draw()

    def set_box(b):
        st["box"] = b
        draw()

    def auto():
        try:
            set_box(content_box(path, st["page"]))
        except Exception as e:
            messagebox.showerror(title, str(e), parent=top)

    def go(d):
        p = st["page"] + d
        if 1 <= p <= n_pages:
            old = st["size"]
            st["page"] = p
            load()
            if st["box"] and st["size"] != old:       # другой размер листа — рамку сбрасываем
                set_box(None)

    def close(ok):
        if ok:
            st["result"] = (st["page"], tuple(round(v, 2) for v in st["box"]) if st["box"] else None)
        top.destroy()

    cv.bind("<ButtonPress-1>", press)
    cv.bind("<B1-Motion>", motion)
    cv.bind("<Motion>", motion)
    cv.bind("<ButtonRelease-1>", release)
    cv.bind("<Configure>", lambda e: draw())
    top.bind("<Escape>", lambda e: close(None))
    top.bind("<Return>", lambda e: close(True))
    top.protocol("WM_DELETE_WINDOW", lambda: close(None))
    top.after(10, load)
    top.update_idletasks()
    x = master.winfo_rootx() + max(0, (master.winfo_width() - top.winfo_reqwidth()) // 2)
    y = master.winfo_rooty() + max(0, (master.winfo_height() - top.winfo_reqheight()) // 3)
    top.geometry(f"+{x}+{y}")
    try:
        top.grab_set()
    except tk.TclError:
        pass
    top.focus_set()
    master.wait_window(top)
    return st["result"]
