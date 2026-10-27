# -*- coding: utf-8 -*-
"""Утилита заполнения паспортов изделий (Динамика МЦД / Статика МЦД).
V1.1 — работа с сохранёнными проектами Калькулятора экранов."""
import csv
import json
import os
import re
import shutil
import subprocess
import sys
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import docx

import book
import calc
import engine
import fonts
import pdf_maket
import project_import
from widgets import PALETTES, GradientList, ScrollFrame, install_clipboard

APP_NAME = "Утилита заполнения паспортов"
APP_VERSION = "V2.1"

# ------------------------------------------------------------------ пути
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
    BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
else:
    APP_DIR = Path(__file__).resolve().parent
    BUNDLE_DIR = APP_DIR



def _unmojibake(name):
    """«╨¿╨░╨▒╨╗╨╛╨╜╤ï» → «Шаблоны»: так Проводник Windows (и 7-Zip) распаковывают
    русские имена из архивов без флага UTF-8."""
    if all(ord(c) < 128 for c in name):
        return None
    for enc in ("cp866", "cp437", "cp1251"):
        try:
            cand = name.encode(enc).decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
        if cand != name and re.search("[А-Яа-яЁё]", cand):
            return cand
    return None


def repair_names(root: Path):
    """Возвращает русские имена файлам и папкам программы, если их испортила распаковка."""
    fixed = []
    for dirpath, dirnames, filenames in os.walk(root, topdown=False):
        for name in filenames + dirnames:
            good = _unmojibake(name)
            if good:
                src, dst = os.path.join(dirpath, name), os.path.join(dirpath, good)
                if not os.path.exists(dst):
                    try:
                        os.rename(src, dst)
                        fixed.append(good)
                    except OSError:
                        pass
    return fixed


REPAIRED = []
if not (APP_DIR / "Шаблоны").is_dir():
    REPAIRED = repair_names(APP_DIR)

TEMPLATES_DIR = APP_DIR / "Шаблоны"
FONTS_DIR = APP_DIR / "Шрифты"
STAMPS_DIR = APP_DIR / "Печати и подписи"
RULES_FILE = TEMPLATES_DIR / "правила_расчёта.json"
SETTINGS_FILE = APP_DIR / "настройки.json"
BLANK_DATE = "«___» ____________ 20___ г."        # паспорт без даты — место для записи от руки
JOURNAL_FILE = APP_DIR / "Журнал_паспортов.csv"
JOURNAL_COLUMNS = ["Создан", "Проект", "Префикс", "Тип", "Наименование", "Артикул",
                   "Серийный номер", "Номер макета", "Тираж", "Гарантия, мес",
                   "Дата продажи", "Печать, подпись и дата", "Файл"]

FIELD_DEFS = {
    "НАИМЕНОВАНИЕ": ("Наименование изделия", "entry"),
    "НАИМЕНОВАНИЕ_ТИТУЛ": ("Наименование\nна титульном листе\n(по строкам)", "text"),
    "АРТИКУЛ": ("Артикул", "entry"),
    "СЕРИЙНЫЙ_НОМЕР": ("Серийный номер\n(первый)", "serial"),
    "НОМЕР_МАКЕТА": ("Номера макетов", "makets"),
    "ГАРАНТИЯ": ("Гарантийный срок,\nмесяцев", "months"),
    "ДАТА_ПРОДАЖИ": ("Дата продажи", "date"),
    "ИЗОБРАЖЕНИЕ_МАКЕТА": ("Изображение макета\n(Приложение 1)", "image"),
}
FIELD_ORDER = list(FIELD_DEFS)
SRC_COLORS = {"расчёт": "#1f5fbf", "проект": "#1e7b34", "шаблон": "#777777", "вручную": "#b35c00"}


def ensure_templates():
    for name, target in (("Шаблоны", TEMPLATES_DIR), ("Шрифты", FONTS_DIR),
                         ("Печати и подписи", STAMPS_DIR)):
        src = BUNDLE_DIR / name
        target.mkdir(exist_ok=True)
        if src.exists() and src.resolve() != target.resolve():
            for f in src.iterdir():
                dst = target / f.name
                if not dst.exists():
                    shutil.copy2(f, dst)


def stamp_library():
    """Файлы из папки «Печати и подписи»:
         Печать*.png                 → [[ПЕЧАТЬ]]
         Подпись_<Фамилия>*.png      → [[ПОДПИСЬ_<ФАМИЛИЯ>]]
       Несколько файлов на одно место (Подпись_Заливакин_1.png, _2.png) — варианты,
       для каждого документа выбирается случайный."""
    lib, cfg = {}, {}
    if STAMPS_DIR.is_dir():
        for f in sorted(STAMPS_DIR.iterdir()):
            if f.suffix.lower() not in (".png", ".gif", ".tif", ".tiff", ".webp"):
                continue
            stem = f.stem.replace(" ", "_")
            m = re.match(r"(?i)(печать|подпись)(?:_([^_\d]+))?", stem)
            if not m:
                continue
            key = "ПЕЧАТЬ" if m.group(1).lower() == "печать" else "ПОДПИСЬ_" + (m.group(2) or "").upper()
            lib.setdefault(key, []).append(str(f))
        conf = STAMPS_DIR / "настройки.json"
        try:
            cfg = json.loads(conf.read_text("utf-8"))
        except Exception:
            cfg = {}
        full = {**engine.STAMP_DEFAULTS, **engine.DATE_DEFAULTS, **cfg}
        if full != cfg:                       # новые параметры (рукописная дата) — дописать в файл
            try:
                conf.write_text(json.dumps(full, ensure_ascii=False, indent=2), "utf-8")
            except Exception:
                pass
        cfg = full
    return lib, cfg


def handwriting_font():
    """Рукописный шрифт для даты: «Шрифты/Caveat*.ttf» (OFL) или любой файл с «рукопис» в имени."""
    if FONTS_DIR.is_dir():
        for f in sorted(FONTS_DIR.iterdir()):
            n = f.name.lower()
            if f.suffix.lower() in (".ttf", ".otf") and ("caveat" in n or "рукопис" in n):
                return str(f)
    return None


def own_digits():
    """Свои цифры: «Печати и подписи/Цифры_Дудин/0.png … 9.png, точка.png»."""
    out = {}
    for place, folder in (("ДАТА_ДУДИН", "Цифры_Дудин"), ("ДАТА_ЗАЛИВАКИН", "Цифры_Заливакин")):
        d = STAMPS_DIR / folder
        if d.is_dir() and all((d / f"{i}.png").exists() for i in range(10)):
            out[place] = str(d)
    return out


def load_types():
    types = []
    for f in sorted(TEMPLATES_DIR.glob("*.docx")):
        if f.name.startswith("~$"):
            continue
        meta = {"тип": f.stem.replace("Паспорт_", "").replace("_", " "), "значения": {}}
        side = f.with_suffix(".json")
        if side.exists():
            try:
                meta.update(json.loads(side.read_text("utf-8")))
            except Exception:
                pass
        meta["шаблон"] = str(f)
        types.append(meta)
    return types


def safe_name(s):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", str(s)).strip(" ._")
    return re.sub(r"\s+", " ", s)


def open_path(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as e:
        messagebox.showerror(APP_NAME, f"Не удалось открыть:\n{path}\n\n{e}")


def docx_to_pdf(docx_path: Path) -> Path:
    pdf = docx_path.with_suffix(".pdf")
    errors = []
    if sys.platform.startswith("win"):
        try:
            import win32com.client
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            try:
                d = word.Documents.Open(str(docx_path), ReadOnly=True)
                d.SaveAs2(str(pdf), FileFormat=17)
                d.Close(False)
            finally:
                word.Quit()
            return pdf
        except Exception as e:
            errors.append(f"MS Word: {e}")
    if sys.platform == "darwin":
        # Word для Mac через AppleScript
        word_app = Path("/Applications/Microsoft Word.app")
        if word_app.exists():
            script = (f'tell application "Microsoft Word"\n'
                      f'open POSIX file "{docx_path}"\n'
                      f'save as active document file name "{pdf}" file format format PDF\n'
                      f'close active document saving no\nend tell')
            r = subprocess.run(["osascript", "-e", script], capture_output=True, timeout=180)
            if pdf.exists():
                return pdf
            errors.append("MS Word: " + r.stderr.decode(errors="ignore")[:200])
    candidates = [shutil.which("soffice"), shutil.which("libreoffice"),
                  r"C:\Program Files\LibreOffice\program\soffice.exe",
                  r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                  "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    for exe in candidates:
        if exe and Path(exe).exists():
            r = subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir",
                                str(docx_path.parent), str(docx_path)],
                               capture_output=True, timeout=180)
            if pdf.exists():
                return pdf
            errors.append(f"LibreOffice: {r.stderr.decode(errors='ignore')[:200]}")
    raise RuntimeError("PDF не создан. Нужен установленный MS Word или LibreOffice.\n"
                       + "\n".join(errors))


# ------------------------------------------------------------------ серийные номера
def serial_numbers(text) -> list:
    """«0240925002–0240925005» / «0240925002-0240925005» / «…001, …007» → список строк."""
    text = str(text or "").replace("№", "")
    text = re.sub(r"\s*([–—-])\s*", r"\1", text)
    out = []
    for part in re.split(r"[,;\s]+(?![–—-])", text.strip()):
        m = re.fullmatch(r"\s*(\d+)\s*[–—-]\s*(\d+)\s*", part)
        if m:
            a, b = m.groups()
            w = len(a)
            out += [str(n).zfill(w) for n in range(int(a), int(b) + 1)]
        elif part.strip().isdigit():
            out.append(part.strip())
    return out


def serial_range(first: str, qty: int, sep="–") -> str:
    first = first.replace("№", "").strip()
    if qty <= 1 or not first.isdigit():
        return first
    last = str(int(first) + qty - 1).zfill(len(first))
    return f"{first}{sep}{last}"


def next_serial(current: str, issued: list) -> str:
    nums = [s for s in issued + [current] if s and s.isdigit()]
    if not nums:
        return current
    width = max(len(s) for s in nums)
    return str(max(int(s) for s in nums) + 1).zfill(width)


def times(n):
    n = int(n)
    if n % 10 in (2, 3, 4) and not 12 <= n % 100 <= 14:
        return f"{n} раза"
    return f"{n} раз"


def plural(n, one, few, many):
    """1 макет, 2 макета, 5 макетов."""
    n = int(n)
    if 11 <= n % 100 <= 14:
        return f"{n} {many}"
    if n % 10 == 1:
        return f"{n} {one}"
    if n % 10 in (2, 3, 4):
        return f"{n} {few}"
    return f"{n} {many}"


# ------------------------------------------------------------------ журнал
def journal_rows():
    if not JOURNAL_FILE.exists():
        return []
    with open(JOURNAL_FILE, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


def journal_add(row):
    rows = journal_rows()
    header_ok = False
    if JOURNAL_FILE.exists():
        with open(JOURNAL_FILE, encoding="utf-8-sig") as f:
            header_ok = f.readline().strip().split(";") == JOURNAL_COLUMNS
    mode = "a" if header_ok else "w"
    with open(JOURNAL_FILE, mode, encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=JOURNAL_COLUMNS, delimiter=";", extrasaction="ignore")
        if not header_ok:
            w.writeheader()
            for r in rows:                    # перенос старых записей в новый формат
                if "Печать и подпись" in r and not r.get("Печать, подпись и дата"):
                    r["Печать, подпись и дата"] = r["Печать и подпись"]
                w.writerow(r)
        w.writerow(row)


# ------------------------------------------------------------------ счётчик скачиваний
class DownloadStore:
    """Хранится рядом с файлом проекта: «<файл проекта>.паспорта.json»."""

    def __init__(self, project_file):
        self.path = Path(str(project_file) + ".паспорта.json")
        try:
            self.data = json.loads(self.path.read_text("utf-8"))
        except Exception:
            self.data = {"изделия": {}}

    @staticmethod
    def key(item):
        return f'{item.get("префикс", "")}|{item.get("тип", "")}|{item.get("ширина")}x{item.get("высота")}'

    def get(self, item):
        return self.data["изделия"].get(self.key(item), {})

    def register(self, item, file, serial=""):
        rec = self.data["изделия"].setdefault(self.key(item), {"скачан": 0, "файлы": []})
        rec["скачан"] = rec.get("скачан", 0) + 1
        rec["последний"] = datetime.now().strftime("%d.%m.%Y %H:%M")
        if str(file) not in rec["файлы"]:
            rec["файлы"].append(str(file))
        if serial:
            rec["серийный_номер"] = serial
        try:
            self.path.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), "utf-8")
        except Exception as e:
            messagebox.showwarning(APP_NAME, f"Не удалось сохранить счётчик скачиваний:\n{e}")
        return rec

    def set_makets(self, item, makets):
        """Макеты изделия: [{"номер", "тираж", "картинки": [путь/PDF-спец, …]}, …]."""
        rec = self.data["изделия"].setdefault(self.key(item), {"скачан": 0, "файлы": []})
        rec["макеты"] = list(makets)
        try:
            self.path.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), "utf-8")
        except Exception:
            pass

    def all_serials(self):
        out = []
        for rec in self.data["изделия"].values():
            out += serial_numbers(rec.get("серийный_номер", ""))
        return out


class MaketRows:
    """Номера макетов изделия: у каждого свой номер и своя часть тиража.
    Сумма по строкам равна тиражу изделия, на каждый макет создаётся свой паспорт.
    Интерфейс как у StringVar: значение — строки «номер|тираж»."""

    def __init__(self, cell, on_change=None, on_number=None):
        self.cell = cell
        self.on_change = on_change or (lambda: None)
        self.on_number = on_number or (lambda i: None)
        self.rows = []                     # [{"num": StringVar, "qty": IntVar, "w": [виджеты]}]
        self.total = 1
        self.err = PALETTES["light"]["err"]
        cell.columnconfigure(1, weight=1)
        self.hint = ttk.Label(cell, style="Hint.TLabel", wraplength=300, justify="left")
        self.hint.grid(row=99, column=0, columnspan=6, sticky="w", pady=(2, 0))
        cell.bind("<Configure>", lambda e: self.hint.configure(wraplength=max(200, e.width - 10)))
        self._add(notify=False)

    # --------------------------------------------------------- строки
    def _add(self, after=None, notify=True):
        i = len(self.rows) if after is None else after + 1
        num, qty = tk.StringVar(), tk.IntVar(value=1)
        row = {"num": num, "qty": qty, "w": []}
        cap = ttk.Label(self.cell, style="Hint.TLabel")
        ent = ttk.Entry(self.cell, textvariable=num, width=12)
        qcap = ttk.Label(self.cell, text="Тираж:", style="Hint.TLabel")
        spn = ttk.Spinbox(self.cell, from_=1, to=9999, width=4, textvariable=qty)
        plus = ttk.Button(self.cell, text="+", width=2, command=lambda r=row: self.add_after(r))
        minus = ttk.Button(self.cell, text="−", width=2, command=lambda r=row: self.remove(r))
        row["w"] = [cap, ent, qcap, spn, plus, minus]
        row["entry"] = ent
        self.rows.insert(i, row)
        qty.trace_add("write", lambda *a: self._changed())
        num.trace_add("write", lambda *a: self.on_change())
        idx = lambda r=row: self.rows.index(r)
        ent.bind("<FocusOut>", lambda e, r=row: self.on_number(idx()))
        ent.bind("<Return>", lambda e, r=row: self.on_number(idx()))
        self._layout()
        if notify:
            self._changed()
        return row

    def add_after(self, row):
        """«+»: новая строка забирает 1 шт. у этой строки (или у любой, где больше одной)."""
        donor = row if row["qty"].get() > 1 else next((r for r in self.rows if r["qty"].get() > 1), None)
        if donor is None:
            messagebox.showinfo(APP_NAME, f"Тираж изделия — {self.total} шт.: "
                                          f"больше макетов не получится, в каждой строке уже по 1 шт.")
            return
        donor["qty"].set(donor["qty"].get() - 1)
        new = self._add(after=self.rows.index(row))
        new["qty"].set(1)
        new["entry"].focus_set()

    def remove(self, row):
        if len(self.rows) < 2:
            return
        i = self.rows.index(row)
        back = self.rows[i - 1] if i else self.rows[1]
        back["qty"].set(back["qty"].get() + row["qty"].get())
        for w in row["w"]:
            w.destroy()
        self.rows.pop(i)
        self._layout()
        self._changed()

    def _layout(self):
        one = len(self.rows) == 1
        for i, r in enumerate(self.rows):
            cap, ent, qcap, spn, plus, minus = r["w"]
            cap.configure(text="" if one else f"{i + 1}.")      # подпись поля — «Номера макетов»
            cap.grid(row=i, column=0, sticky="e", padx=(0, 4), pady=1)
            ent.grid(row=i, column=1, sticky="ew", pady=1)
            qcap.grid(row=i, column=2, padx=(8, 3))
            spn.grid(row=i, column=3)
            plus.grid(row=i, column=4, padx=(5, 0))
            minus.grid(row=i, column=5, padx=(2, 0))
            if one:
                minus.grid_remove()

    # --------------------------------------------------------- тираж
    def set_palette(self, pal):
        self.err = pal["err"]
        self._changed()

    def set_total(self, total):
        """Тираж изделия поменялся: разница уходит в последнюю строку."""
        self.total = max(1, int(total or 1))
        d = self.total - sum(r["qty"].get() for r in self.rows)
        if d:
            last = self.rows[-1]
            last["qty"].set(max(1, last["qty"].get() + d))
        self._changed()

    def distributed(self):
        return sum(max(1, r["qty"].get()) for r in self.rows)

    def ok(self):
        return self.distributed() == self.total

    def _changed(self):
        for r in self.rows:
            try:
                if r["qty"].get() < 1:
                    r["qty"].set(1)
            except tk.TclError:
                r["qty"].set(1)
        n = self.distributed()
        if len(self.rows) == 1:
            self.hint.configure(text="«+» — ещё один макет: тираж разделится, на каждый "
                                     "будет свой паспорт", foreground="")
        elif n == self.total:
            self.hint.configure(text=f"{plural(len(self.rows), 'макет', 'макета', 'макетов')}, "
                                     f"всего {n} шт. — на каждый свой паспорт", foreground="")
        else:
            self.hint.configure(text=f"распределено {n} из {self.total} шт. — исправьте тираж",
                                foreground=self.err)
        self.on_change()

    # --------------------------------------------------------- значения
    def items(self):
        return [{"номер": r["num"].get().strip(), "тираж": max(1, r["qty"].get())} for r in self.rows]

    def numbers(self):
        return [r["num"].get().strip() for r in self.rows]

    def count(self):
        return len(self.rows)

    def get(self):
        return "\n".join(f'{x["номер"]}|{x["тираж"]}' for x in self.items())

    def set(self, value):
        parts = [l for l in str(value or "").splitlines() if l.strip() != ""] or [""]
        while len(self.rows) > len(parts):
            r = self.rows.pop()
            for w in r["w"]:
                w.destroy()
        while len(self.rows) < len(parts):
            self._add(notify=False)
        for r, line in zip(self.rows, parts):
            num, _, q = line.partition("|")
            r["num"].set(num.strip())
            try:
                r["qty"].set(max(1, int(q)))
            except ValueError:
                r["qty"].set(self.total if len(parts) == 1 else 1)
        self._layout()
        self._changed()


class MaketImages:
    """Поля картинок Приложения 1: свой набор на каждый макет, в наборе — по числу сторон.
    Значение (get/set) — JSON: [[спец стороны 1, спец стороны 2], …] по макетам."""

    def __init__(self, cell, app, key):
        self.cell = cell
        self.app = app
        self.key = key
        self.sets = []            # [{"vars":[StringVar×2], "disp":[StringVar×2], "w":[строки]}]
        self.sides = 1
        self.titles = [""]
        cell.columnconfigure(1, weight=1)
        self._ensure(1)

    # --------------------------------------------------------- построение
    def _ensure(self, n):
        while len(self.sets) < n:
            self._add_set()
        while len(self.sets) > n:
            st = self.sets.pop()
            for row in st["w"]:
                for w in row:
                    w.destroy()

    def _add_set(self):
        m = len(self.sets)
        st = {"vars": [], "disp": [], "w": [], "crop": []}
        for side in (1, 2):
            v, d = tk.StringVar(), tk.StringVar()
            cap = ttk.Label(self.cell, style="Hint.TLabel")
            ent = ttk.Entry(self.cell, textvariable=d, state="readonly", width=12)
            pick = ttk.Button(self.cell, text="…", width=3,
                              command=lambda mm=m, ss=side: self.app._choose_image(self, mm, ss))
            crop = ttk.Button(self.cell, text="Обрезать",
                              command=lambda mm=m, ss=side: self.app._crop_maket(self, mm, ss))
            clr = ttk.Button(self.cell, text="×", width=3,
                             command=lambda mm=m, ss=side: self.app._clear_maket(self, mm, ss))
            crop.state(["disabled"])
            v.trace_add("write", lambda *a, v=v, d=d, c=crop: self._sync(v, d, c))
            st["vars"].append(v)
            st["disp"].append(d)
            st["crop"].append(crop)
            st["w"].append([cap, ent, pick, crop, clr])
        self.sets.append(st)

    @staticmethod
    def _sync(v, d, crop):
        spec = v.get().strip()
        d.set(pdf_maket.describe(spec) if spec else "")
        try:
            crop.state(["!disabled"] if pdf_maket.is_pdf(spec) else ["disabled"])
        except tk.TclError:
            pass

    def show(self, sides, titles=None):
        """sides — сторон у изделия (1 или 2); titles — подписи макетов («М-0457» …)."""
        self.sides = 2 if int(sides or 1) >= 2 else 1
        self.titles = list(titles or [""])
        self._ensure(max(1, len(self.titles)))
        r = 0
        for m, st in enumerate(self.sets):
            name = (self.titles[m] if m < len(self.titles) else "") or f"Макет {m + 1}"
            for side in (1, 2):
                cap, ent, pick, crop, clr = st["w"][side - 1]
                if side > self.sides:
                    for w in (cap, ent, pick, crop, clr):
                        w.grid_remove()
                    st["vars"][side - 1].set("")
                    continue
                one = len(self.sets) == 1
                cap.configure(text=(f"Сторона {side}" if one else f"{name} ст. {side}")
                              if self.sides == 2 else ("Файл" if one else name))
                cap.grid(row=r, column=0, sticky="w", padx=(0, 6))
                ent.grid(row=r, column=1, sticky="ew", pady=1)
                pick.grid(row=r, column=2, padx=(4, 0))
                crop.grid(row=r, column=3, padx=(2, 0))
                clr.grid(row=r, column=4, padx=(2, 0))
                r += 1

    # --------------------------------------------------------- значения
    def var(self, maket, side):
        self._ensure(max(len(self.sets), maket + 1))
        return self.sets[maket]["vars"][side - 1]

    def specs(self):
        """[[сторона 1, сторона 2], …] — по макетам, с учётом числа сторон."""
        out = []
        for st in self.sets:
            out.append([st["vars"][0].get().strip(),
                        st["vars"][1].get().strip() if self.sides == 2 else ""])
        return out

    def set_specs(self, specs):
        self._ensure(max(1, len(specs)))
        for st, pair in zip(self.sets, list(specs) + [["", ""]] * len(self.sets)):
            pair = (list(pair) + ["", ""])[:2]
            st["vars"][0].set(pair[0])
            st["vars"][1].set(pair[1])

    def get(self):
        return json.dumps(self.specs(), ensure_ascii=False)

    def set(self, value):
        if isinstance(value, str) and value.strip().startswith("["):
            try:
                self.set_specs(json.loads(value))
                return
            except Exception:
                pass
        parts = [x for x in str(value or "").splitlines() if x.strip()] + ["", ""]
        self.set_specs([[parts[0], parts[1]]])


# ------------------------------------------------------------------ окно
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} — {APP_VERSION}")
        # окно по размеру экрана (ноутбуки 1366×768 и меньше — разворачиваем)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(1280, sw - 40), min(860, sh - 80)
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 3)}")
        self.minsize(min(1050, sw - 40), min(640, sh - 80))
        if sh < 900 and sys.platform.startswith("win"):
            try:
                self.state("zoomed")
            except tk.TclError:
                pass
        self.style = ttk.Style(self)

        ensure_templates()
        self.rules = calc.load_rules(RULES_FILE)
        self.bundled_fonts = fonts.bundled(FONTS_DIR)
        calc.FONT_FILE = self.bundled_fonts.get("RussianRail G Pro")
        self.settings = self._load_settings()
        self.theme = tk.StringVar(value=self.settings.get("тема", "light"))
        self._apply_theme_style()
        self.types = load_types()
        if not self.types:
            messagebox.showerror(APP_NAME, f"Не найдены шаблоны паспортов (.docx) в папке:\n{TEMPLATES_DIR}\n\n"
                                           "Распакуйте архив целиком и запускайте программу из распакованной "
                                           "папки (не из окна архива).")
            self.destroy()
            raise SystemExit(1)
        self.state_by_type = self.settings.get("состояние", {})
        self.type_var = tk.StringVar(value=self.settings.get("тип", self.types[0]["тип"]))
        if self.type_var.get() not in [t["тип"] for t in self.types]:
            self.type_var.set(self.types[0]["тип"])
        self.current_type = None
        self.field_widgets = {}
        self.out_dir = tk.StringVar(value=self.settings.get("папка", str(Path.home() / "Паспорта")))
        self.make_pdf = tk.BooleanVar(value=self.settings.get("pdf", False))
        self.print_ready = tk.BooleanVar(value=self.settings.get("типография", True))
        self.use_stamps = tk.BooleanVar(value=False)     # по умолчанию всегда выключено
        self.open_after = tk.BooleanVar(value=self.settings.get("открыть", True))
        self.project_name = tk.StringVar()
        self.qty = tk.StringVar(value="1")
        self.status = tk.StringVar(value="Готово")
        self.project = None
        self.store = None
        self.current_item = None

        self._build()
        install_clipboard(self)          # Ctrl+C/V/X/A при любой раскладке, меню правой кнопкой
        self._switch_type()
        self._apply_theme()
        self.after(400, self._check_fonts)
        if REPAIRED:
            self.status.set(f"Исправлены имена файлов после распаковки архива: {len(REPAIRED)}.")
        last = self.settings.get("проект")
        if last and Path(last).exists():
            self._open_project(last, quiet=True)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------------------------------------------------- темы
    def _apply_theme_style(self):
        """Светлая: на Windows — стандартный вид; на Mac/Linux — светлая палитра.
        Тёмная — своя палитра на теме clam (одинаково на всех системах)."""
        name = self.theme.get() if self.theme.get() in PALETTES else "light"
        p = PALETTES[name]
        st = self.style
        native = name == "light" and sys.platform.startswith("win") and "vista" in st.theme_names()
        st.theme_use("vista" if native else "clam")
        if not native:
            st.configure(".", background=p["bg"], foreground=p["fg"], fieldbackground=p["field"],
                         bordercolor=p["border"], lightcolor=p["bg"], darkcolor=p["bg"],
                         troughcolor=p["panel"], selectbackground=p["accent"],
                         selectforeground="#ffffff", insertcolor=p["fg"], arrowcolor=p["fg"])
            st.map(".", background=[("disabled", p["bg"])], foreground=[("disabled", p["hint"])])
            st.configure("TButton", background=p["button"], foreground=p["fg"], padding=(8, 3))
            st.map("TButton", background=[("pressed", p["button_active"]), ("active", p["button_active"])])
            st.configure("TLabelframe", background=p["bg"], bordercolor=p["border"])
            st.configure("TLabelframe.Label", background=p["bg"], foreground=p["fg"])
            for w in ("TCheckbutton", "TRadiobutton"):
                st.configure(w, background=p["bg"], foreground=p["fg"], indicatorbackground=p["field"],
                             indicatorforeground=p["fg"])
                st.map(w, background=[("active", p["bg"])],
                       indicatorbackground=[("selected", p["accent"]), ("active", p["field"])])
            st.configure("TEntry", fieldbackground=p["field"], foreground=p["fg"])
            st.configure("TSpinbox", fieldbackground=p["field"], foreground=p["fg"], arrowcolor=p["fg"])
            st.configure("Treeview", background=p["field"], fieldbackground=p["field"], foreground=p["fg"])
            st.map("Treeview", background=[("selected", p["accent"])], foreground=[("selected", "#ffffff")])
            st.configure("Treeview.Heading", background=p["head_bg"], foreground=p["head_fg"])
            st.map("Treeview.Heading", background=[("active", p["button_active"])])
            st.configure("TScrollbar", background=p["button"], troughcolor=p["bg"], arrowcolor=p["fg"])
            st.configure("TPanedwindow", background=p["bg"])
            st.configure("TFrame", background=p["bg"])
            st.configure("TLabel", background=p["bg"], foreground=p["fg"])
        st.configure("Big.TButton", font=("Segoe UI", 11, "bold"), padding=(18, 8))
        st.configure("Hint.TLabel", foreground=p["hint"])
        st.configure("Name.TLabel", foreground=p["fg"], font=("Segoe UI", 10, "underline"))
        st.configure("Proj.TLabel", font=("Segoe UI", 11, "bold"))
        st.configure("Treeview", rowheight=24)
        return p

    def _apply_theme(self):
        p = self._apply_theme_style()
        self.configure(bg=p["bg"])
        self.option_add("*Toplevel.background", p["bg"])
        if hasattr(self, "items_list"):
            self.items_list.set_palette(p)
        if hasattr(self, "fields_scroll"):
            self.fields_scroll.set_palette(p)
        if hasattr(self, "tree"):
            for src, col in p["src"].items():
                self.tree.tag_configure(src, foreground=col)
        self._theme_text_widgets()
        self._save_settings_soft()

    def _theme_text_widgets(self):
        p = PALETTES.get(self.theme.get(), PALETTES["light"])
        for kind, w in getattr(self, "field_widgets", {}).values():
            if kind == "makets":
                w.set_palette(p)
            if kind == "text":
                w.configure(bg=p["field"], fg=p["fg"], insertbackground=p["fg"],
                            highlightbackground=p["border"], selectbackground=p["accent"])

    def _save_settings_soft(self):
        if getattr(self, "current_type", None):
            try:
                self._save_settings()
            except Exception:
                pass

    # ---------------------------------------------------------- шрифты
    def _check_fonts(self):
        miss = fonts.missing()
        if not miss:
            return
        lines = "\n".join(f"• {f} — {fonts.REQUIRED[f]}" for f in miss)
        have = {f: p for f, p in self.bundled_fonts.items() if f in miss}
        if have and messagebox.askyesno(
                APP_NAME, f"В системе не установлены шрифты паспорта:\n{lines}\n\n"
                          f"Без них Word подставит другие шрифты: титул и страницы «поедут».\n\n"
                          f"Установить из папки «Шрифты» сейчас?"):
            err = fonts.install(have)
            if err:
                messagebox.showerror(APP_NAME, "Не удалось установить:\n" + "\n".join(err))
            else:
                messagebox.showinfo(APP_NAME, "Шрифты установлены. Если Word открыт — перезапустите его.")
        elif not have:
            messagebox.showwarning(APP_NAME, f"Не установлены шрифты:\n{lines}\n\n"
                                             f"Положите файлы шрифтов в папку:\n{FONTS_DIR}")

    # ---------------------------------------------------------- настройки
    def _load_settings(self):
        try:
            return json.loads(SETTINGS_FILE.read_text("utf-8"))
        except Exception:
            return {}

    def _save_settings(self):
        if self.current_type and not self.current_item:
            self.state_by_type[self.current_type["тип"]] = self._collect_state()
        data = {"тип": self.type_var.get(), "папка": self.out_dir.get(),
                "pdf": self.make_pdf.get(), "открыть": self.open_after.get(),
                "типография": self.print_ready.get(),
                "тема": self.theme.get(),
                "сортировка": [self.items_list.sort_col, self.items_list.sort_dir]
                if hasattr(self, "items_list") else [None, 0],
                "шаблоны": self.settings.get("шаблоны", {}),
                "проект": self.project["файл"] if self.project else "",
                "состояние": self.state_by_type}
        try:
            SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")
        except Exception:
            pass

    def _on_close(self):
        self._save_settings()
        self.destroy()

    # ---------------------------------------------------------- разметка
    def _build(self):
        top = ttk.LabelFrame(self, text="Тип изделия", padding=(10, 6))
        top.pack(fill="x", padx=10, pady=(10, 4))
        self.type_buttons = []
        for t in self.types:
            b = ttk.Radiobutton(top, text=t["тип"], value=t["тип"], variable=self.type_var,
                                command=self._on_type_radio)
            b.pack(side="left", padx=(0, 18))
        th = ttk.Frame(top)
        th.pack(side="right", padx=(12, 0))
        ttk.Label(th, text="Тема:").pack(side="left")
        ttk.Radiobutton(th, text="Светлая", value="light", variable=self.theme,
                        command=self._apply_theme).pack(side="left", padx=(4, 0))
        ttk.Radiobutton(th, text="Тёмная", value="dark", variable=self.theme,
                        command=self._apply_theme).pack(side="left", padx=(4, 0))
        ttk.Button(top, text="Открыть шаблон в Word", command=self._open_template).pack(side="right")
        ttk.Button(top, text="Другой шаблон…", command=self._choose_template).pack(side="right", padx=6)
        self.tpl_label = ttk.Label(top, style="Hint.TLabel")
        self.tpl_label.pack(side="right", padx=10)

        ttk.Label(self, textvariable=self.status, anchor="w", relief="sunken",
                  padding=(8, 2)).pack(fill="x", side="bottom")
        bottom = ttk.LabelFrame(self, text="Сохранение", padding=(10, 8))
        bottom.pack(fill="x", side="bottom", padx=10, pady=4)
        self._bottom = bottom
        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=10, pady=4)

        # ---- проект
        proj = ttk.LabelFrame(paned, text="Проект из Калькулятора экранов", padding=(10, 6))
        bar = ttk.Frame(proj)
        bar.pack(fill="x")
        ttk.Button(bar, text="Открыть проект…", command=self._choose_project).pack(side="left")
        ttk.Button(bar, text="Обновить", command=self._reload_project).pack(side="left", padx=6)
        ttk.Button(bar, text="Закрыть проект", command=self._close_project).pack(side="left")
        ttk.Button(bar, text="Правила расчёта", command=lambda: open_path(RULES_FILE)).pack(side="right")
        self.proj_label = ttk.Label(bar, text="Проект не открыт", style="Proj.TLabel")
        self.proj_label.pack(side="left", padx=16)
        cols = [("№", "№", 40, "center"), ("префикс", "Префикс", 80, "w"), ("тип", "Тип", 115, "w"),
                ("_артикул", "Артикул", 240, "w"), ("_размер", "Размер, мм", 118, "center"),
                ("_сторон", "Сторон", 65, "center"), ("тираж", "Тираж, шт", 92, "center"),
                ("_идент", "Макет / Серийный №", 210, "w"), ("_скачан", "Скачан", 95, "center")]
        num = lambda k: (lambda r: calc._float(r.get(k), 0))
        sort_keys = {"№": num("№"), "_сторон": num("_сторон"), "тираж": num("тираж"),
                     "_скачан": num("_n"), "_размер": lambda r: r.get("_wh", (0, 0))}
        self.items_list = GradientList(proj, cols, on_select=self._on_item, stretch="_артикул",
                                       sort_keys=sort_keys, on_sort=lambda *a: self._save_settings_soft())
        srt = self.settings.get("сортировка") or [None, 0]
        if isinstance(srt, list) and len(srt) == 2:
            self.items_list.set_sort(srt[0], srt[1])
        self.items_list.empty_text = ("Откройте сохранённый файл проекта Калькулятора экранов — "
                                      "изделия появятся списком. Щелчок по строке заполнит паспорт.")
        self.calc_label = ttk.Label(proj, style="Hint.TLabel")
        self.calc_label.pack(fill="x", side="bottom", pady=(4, 0))
        self.items_list.pack(fill="both", expand=True, pady=(6, 0))
        paned.add(proj, weight=2)

        # ---- поля + таблица
        mid = ttk.Frame(paned)
        # minsize — поля паспорта не сжимаются до нечитаемого на узком окне,
        # место уступает таблица техданных (её колонки растягиваются сами)
        mid.columnconfigure(0, weight=6, minsize=500)
        mid.columnconfigure(1, weight=5, minsize=330)
        mid.rowconfigure(0, weight=1)
        paned.add(mid, weight=3)

        ff = ttk.LabelFrame(mid, text="Данные паспорта", padding=(10, 6, 2, 6))
        ff.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.fields_scroll = ScrollFrame(ff)
        self.fields_scroll.pack(fill="both", expand=True)
        self.fields_frame = self.fields_scroll.body

        tbl = ttk.LabelFrame(mid, text="Технические данные и характеристики", padding=10)
        tbl.grid(row=0, column=1, sticky="nsew")
        tbl.rowconfigure(0, weight=1)
        tbl.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(tbl, columns=("p", "v", "s"), show="headings", selectmode="browse")
        self.tree.heading("p", text="Параметр")
        self.tree.heading("v", text="Значение")
        self.tree.heading("s", text="Источник")
        # узкие исходные ширины: на широком окне колонки растянутся сами,
        # а на узком таблица не отнимает место у полей паспорта
        self.tree.column("p", width=150, minwidth=90, anchor="w")
        self.tree.column("v", width=110, minwidth=70, anchor="w")
        self.tree.column("s", width=70, minwidth=55, anchor="center")
        for src, col in SRC_COLORS.items():
            self.tree.tag_configure(src, foreground=col)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind("<Double-1>", self._edit_cell)
        bt = ttk.Frame(tbl)
        bt.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(bt, text="+ Строка", command=self._row_add).pack(side="left")
        ttk.Button(bt, text="− Удалить", command=self._row_del).pack(side="left", padx=4)
        ttk.Button(bt, text="Вверх", command=lambda: self._row_move(-1)).pack(side="left")
        ttk.Button(bt, text="Вниз", command=lambda: self._row_move(1)).pack(side="left", padx=4)
        ttk.Button(bt, text="Сбросить из шаблона", command=self._table_from_template).pack(side="right")
        ttk.Label(tbl, text="Источник: расчёт — по правилам, проект — из файла проекта, "
                            "шаблон — как в шаблоне, вручную — исправлено вами. "
                            "Двойной щелчок — изменить.", style="Hint.TLabel",
                  wraplength=560).grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        bottom = self._bottom
        ttk.Label(bottom, text="Папка для паспортов:").grid(row=0, column=0, sticky="w")
        ttk.Entry(bottom, textvariable=self.out_dir).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(bottom, text="Обзор…", command=self._choose_out).grid(row=0, column=2)
        bottom.columnconfigure(1, weight=1)
        self.name_override = ""
        nf = ttk.Frame(bottom)
        nf.grid(row=1, column=0, columnspan=3, sticky="ew")
        self.name_caption = ttk.Label(nf, text="Имя файла:")
        self.name_caption.grid(row=0, column=0, sticky="w", pady=(6, 0))
        self.name_preview = ttk.Label(nf, style="Name.TLabel", cursor="xterm")
        self.name_preview.grid(row=0, column=1, sticky="w", padx=6, pady=(6, 0))
        self.name_preview.bind("<Double-Button-1>", self._edit_name)
        self.name_hint = ttk.Label(nf, text="(двойной щелчок — изменить)", style="Hint.TLabel")
        self.name_hint.grid(row=0, column=2, sticky="w", pady=(6, 0))
        nf.columnconfigure(3, weight=1)
        ttk.Button(nf, text="Сохранить данные…", command=self._save_data).grid(
            row=0, column=5, sticky="e", pady=(6, 0))
        ttk.Button(nf, text="Открыть данные…", command=self._load_data).grid(
            row=0, column=4, sticky="e", padx=8, pady=(6, 0))
        opts = ttk.Frame(bottom)
        opts.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Checkbutton(opts, text="Для типографии (PDF, кратно 4 стр.)",
                        variable=self.print_ready).pack(side="left")
        ttk.Checkbutton(opts, text="PDF", variable=self.make_pdf).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(opts, text="Печать, подпись и дата", variable=self.use_stamps).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(opts, text="Открыть папку",
                        variable=self.open_after).pack(side="left", padx=12)
        ttk.Button(opts, text="Скачать паспорт", style="Big.TButton",
                   command=self._generate).pack(side="right")
        ttk.Button(opts, text="Журнал", command=self._show_journal).pack(side="right", padx=8)

        self.project_name.trace_add("write", lambda *a: self._update_name_preview())
        self.qty.trace_add("write", lambda *a: self._qty_changed())

    # ---------------------------------------------------------- тип / шаблон
    def _type_meta(self, name=None):
        name = name or self.type_var.get()
        return next(t for t in self.types if t["тип"] == name)

    def _template_path(self, meta=None):
        meta = meta or self.current_type
        return self.settings.get("шаблоны", {}).get(meta["тип"], meta["шаблон"])

    def _on_type_radio(self):
        # ручной выбор типа — выходим из режима «изделие проекта»
        was_item = self.current_item is not None
        self._deselect_item()
        self._switch_type(save=not was_item)

    def _switch_type(self, restore=True, save=True):
        if save and self.current_type is not None and not self.current_item:
            self.state_by_type[self.current_type["тип"]] = self._collect_state()
        self.current_type = self._type_meta()
        tpl = self._template_path()
        try:
            d = docx.Document(tpl)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Не удалось открыть шаблон:\n{tpl}\n\n{e}")
            return
        self.tpl_keys = engine.find_placeholders(d)
        self.tpl_img_keys = engine.image_placeholders(d)
        self.tpl_table = engine.read_tech_table(d)
        self.tpl_label.configure(text="Шаблон: " + Path(tpl).name)
        self.title_box = engine.title_box(d)
        self._build_fields()
        if restore:
            state = self.state_by_type.get(self.current_type["тип"])
            self._apply_state(state or {"значения": self.current_type.get("значения", {}),
                                        "таблица": self.tpl_table})
        self.status.set(f"Шаблон: меток {len(self.tpl_keys)}, строк в таблице техданных "
                        f"{len(self.tpl_table)}." if self.tpl_table else
                        "В шаблоне не найдена таблица «Технические данные и характеристики».")

    def _choose_template(self):
        f = filedialog.askopenfilename(title="Шаблон паспорта", filetypes=[("Word", "*.docx")])
        if not f:
            return
        self.settings.setdefault("шаблоны", {})[self.current_type["тип"]] = f
        if not self.current_item:
            self.state_by_type[self.current_type["тип"]] = self._collect_state()
        self.current_type = None
        self._switch_type()
        if self.current_item is not None:
            self._on_item(self.items_list.selected)

    def _open_template(self):
        open_path(self._template_path())

    # ---------------------------------------------------------- проект
    def _choose_project(self):
        f = filedialog.askopenfilename(
            title="Файл проекта Калькулятора экранов",
            filetypes=[("Проект", "*.json *.ledproj *.proj *.csv *.xlsx"), ("Все файлы", "*.*")])
        if f:
            self._open_project(f)

    def _reload_project(self):
        if self.project:
            sel = self.items_list.selected
            self._open_project(self.project["файл"], quiet=True)
            if sel is not None:
                self.items_list.select(sel)

    def _close_project(self):
        self.project = None
        self.store = None
        self._deselect_item()
        self.items_list.set_rows([], keep_selection=False)
        self.proj_label.configure(text="Проект не открыт")
        self.project_name.set("")
        self.calc_label.configure(text="")

    def _open_project(self, path, quiet=False):
        try:
            p = project_import.load_project(path)
        except Exception as e:
            if not quiet:
                messagebox.showerror(APP_NAME, str(e))
            return
        known = [t["тип"] for t in self.types]
        unknown = [it for it in p["изделия"] if it["тип"] not in known]
        self.project = p
        self.store = DownloadStore(path)
        self.project_name.set(p["проект"])
        self.proj_label.configure(text=f'{p["проект"]}  ·  изделий: {len(p["изделия"])}  ·  '
                                       f'{Path(path).name}')
        self._deselect_item()
        self._refresh_list(keep=False)
        msg = f'Проект открыт: {p["проект"]}, изделий {len(p["изделия"])}.'
        if unknown:
            msg += (f" У {len(unknown)} изделий не удалось определить тип — при выборе "
                    f"будет использован тип, отмеченный вверху.")
        self.status.set(msg)

    def _item_row(self, i, it):
        rec = self.store.get(it) if self.store else {}
        saved = [m for m in (rec.get("макеты") or []) if isinstance(m, dict)]
        if saved and any(m.get("номер") for m in saved):
            ident = ", ".join(f'{m.get("номер") or "Б/Н"}×{m.get("тираж", 1)}' for m in saved) \
                if len(saved) > 1 else str(saved[0].get("номер") or "")
        else:
            ident = it.get("номер_макета") or rec.get("серийный_номер") or it.get("серийный_номер") or ""
        if it.get("тип") == "Динамика МЦД" and ident and "–" not in ident:
            ident = serial_range(ident, int(it.get("тираж") or 1))
        typ = it.get("тип") or "?"
        try:
            _, _, v = calc.compute(it if it.get("тип") else {**it, "тип": self.type_var.get()},
                                   self.rules, [])
            art = it.get("артикул") or calc._fmt(self.rules.get(it.get("тип") or self.type_var.get(),
                                                                 {}).get("артикул", ""), v)
        except Exception:
            art = it.get("артикул", "")
        n = rec.get("скачан", 0)
        return {"№": i + 1, "префикс": it.get("префикс", ""), "тип": typ, "_артикул": art,
                "_размер": (f'Ø{it.get("ширина")}' if it.get("форма") == "круг"
                            else f'{it.get("ширина")}×{it.get("высота")}'),
                "_сторон": it.get("сторон", 1), "тираж": it.get("тираж", 1),
                "_идент": ident, "_скачан": times(n) if n else "—", "_done": n > 0,
                "_n": n, "_wh": (calc._float(it.get("ширина"), 0), calc._float(it.get("высота"), 0))}

    def _refresh_list(self, keep=True):
        rows = [self._item_row(i, it) for i, it in enumerate(self.project["изделия"])] \
            if self.project else []
        self.items_list.set_rows(rows, keep_selection=keep)

    def _deselect_item(self):
        self.name_override = ""
        self.current_item = None
        self.items_list.selected = None
        self.items_list.redraw()
        self.calc_label.configure(text="")
        self._update_name_preview()

    def _on_item(self, index):
        self.name_override = ""
        if index is None or not self.project:
            return
        it = dict(self.project["изделия"][index])
        if not it.get("тип"):
            it["тип"] = self.type_var.get()
        if it["тип"] not in [t["тип"] for t in self.types]:
            messagebox.showwarning(APP_NAME, f'Для типа «{it["тип"]}» нет шаблона паспорта.')
            return
        if not self.current_item and self.current_type:
            self.state_by_type[self.current_type["тип"]] = self._collect_state()
        self.current_item = None
        if it["тип"] != self.type_var.get() or self.current_type is None:
            self.type_var.set(it["тип"])
        self._switch_type(restore=True, save=False)   # гарантия и дата — последние введённые
        self.current_item = it
        values, rows, v = calc.compute(it, self.rules, self.tpl_table, self.title_box)
        rec = self.store.get(it)
        if rec.get("серийный_номер"):
            values["СЕРИЙНЫЙ_НОМЕР"] = serial_numbers(rec["серийный_номер"])[0]
        elif "СЕРИЙНЫЙ_НОМЕР" in self.field_widgets and not it.get("серийный_номер"):
            values["СЕРИЙНЫЙ_НОМЕР"] = next_serial("", self._issued_serials())  # предложение
        for k, val in values.items():
            self._set(k, val)
        self._set_table(rows)
        self.qty.set(str(it.get("тираж", 1)))
        self._restore_makets(it)
        self.project_name.set(self.project["проект"])
        if it["тип"] == "Динамика МЦД":
            info = (f'Расчёт {it.get("префикс")}: модулей {v["модулей_по_ширине"]}×{v["модулей_по_высоте"]} '
                    f'= {v["модулей"]} шт, поля {v["поле_гор_мм"]} / {v["поле_верт_мм"]} мм, '
                    f'разрешение {v["px_w"]}×{v["px_h"]}, блоки питания {v["блоки_питания"]} → '
                    f'P макс {v["P_max_кВт"]} кВт, P ср {v["P_ср_кВт"]} кВт, масса {v["масса"]} кг')
        else:
            info = (f'Расчёт {it.get("префикс")}: LED {v["LED_шт"]} шт, мощность LED +20% {v["нагрузка_Вт"]} Вт → '
                    f'блоки питания {v["блоки_питания"]} = {v["P_Вт"]} Вт, масса {v["масса"]} кг')
        if it.get("масса") not in (None, "") and it.get("_расчёт"):
            psu = ", ".join(f'{calc.num(p_["qty"], 0)}×{p_["w"]} Вт' for p_ in it.get("блоки_питания") or [])
            info = (f'{it.get("префикс")}: по расчёту {it["_расчёт"]} — масса {calc.num(it["масса"], 2)} кг'
                    + (f', блоки питания {psu}' if psu else '')
                    + (f', суммарно {calc.num(it["мощность_Вт"], 0)} Вт' if it.get("мощность_Вт") else ''))
        if rec.get("скачан"):
            info += f'   ·   скачан {times(rec["скачан"])}, последний раз {rec.get("последний", "")}'
        if self.project.get("формат") and it.get("масса") in (None, ""):
            info = ("⚠ Вес ориентировочный (мощность — по правилам Конструктора): в файле проекта нет "
                    "расчёта Конструктора — пересохраните проект в V40.04.2 или новее.   ·   " + info)
        self.calc_label.configure(text=info)
        self._update_name_preview()
        for kind, w in self.field_widgets.values():   # Ctrl+Z не возвращает титул прошлого изделия
            if kind == "text":
                w.edit_reset()
        self.status.set(f'Изделие {it.get("префикс")} загружено. Проверьте данные и нажмите «Скачать паспорт».')

    # ---------------------------------------------------------- поля
    def _build_fields(self):
        for w in self.fields_frame.winfo_children():
            w.destroy()
        for tr in getattr(self, "_qty_traces", []):
            try:
                self.qty.trace_remove("write", tr)
            except Exception:
                pass
        self._qty_traces = []
        self.field_widgets = {}
        self.field_entries = {}
        keys = [k for k in FIELD_ORDER if k in self.tpl_keys] + \
               [k for k in self.tpl_keys if k not in FIELD_DEFS]
        f = self.fields_frame
        f.columnconfigure(1, weight=1)
        f.columnconfigure(2, minsize=8)          # отступ от полосы прокрутки
        r = 0
        # общие поля: проект и тираж (для имени файла и журнала)
        ttk.Label(f, text="Проект:").grid(row=r, column=0, sticky="e", padx=(0, 8), pady=4)
        qf = ttk.Frame(f)
        qf.grid(row=r, column=1, sticky="ew", pady=4)
        qf.columnconfigure(0, weight=1)
        ttk.Entry(qf, textvariable=self.project_name).grid(row=0, column=0, sticky="ew")
        ttk.Label(qf, text="Тираж, шт:").grid(row=0, column=1, padx=(10, 4))
        ttk.Spinbox(qf, from_=1, to=9999, textvariable=self.qty, width=6).grid(row=0, column=2)
        r += 1
        for key in keys:
            label, kind = FIELD_DEFS.get(key, (key.replace("_", " ").capitalize(), "entry"))
            if key in self.tpl_img_keys:
                kind = "image"
            ttk.Label(f, text=label + ":", justify="right").grid(
                row=r, column=0, sticky="ne", padx=(0, 8), pady=3)
            cell = ttk.Frame(f)
            cell.grid(row=r, column=1, sticky="ew", pady=3)
            cell.columnconfigure(0, weight=1)
            if kind == "text":
                w = tk.Text(cell, height=4, width=34, wrap="none", font="TkDefaultFont",
                            relief="solid", borderwidth=1, undo=True, autoseparators=True, maxundo=-1)
                w.grid(row=0, column=0, sticky="ew")
                ttk.Button(cell, text="Из наименования", command=self._title_from_name).grid(
                    row=0, column=1, sticky="n", padx=(6, 0))
                self.field_widgets[key] = ("text", w)
            elif kind == "months":
                v = tk.StringVar(value="36")
                ttk.Spinbox(cell, from_=1, to=240, textvariable=v, width=6).grid(row=0, column=0, sticky="w")
                lab = ttk.Label(cell, style="Hint.TLabel")
                lab.grid(row=0, column=1, sticky="w", padx=8)
                cell.columnconfigure(0, weight=0)
                cell.columnconfigure(1, weight=1)
                v.trace_add("write", lambda *a, v=v, lab=lab: self._preview_months(v, lab))
                self._preview_months(v, lab)
                self.field_widgets[key] = ("var", v)
            elif kind == "date":
                v = tk.StringVar(value=date.today().strftime("%d.%m.%Y"))
                self.field_entries[key] = ttk.Entry(cell, textvariable=v, width=12)
                self.field_entries[key].grid(row=0, column=0, sticky="w")
                ttk.Button(cell, text="×", width=3, command=lambda v=v: v.set("")).grid(
                    row=0, column=1, padx=(4, 0))
                ttk.Button(cell, text="Сегодня",
                           command=lambda v=v: v.set(date.today().strftime("%d.%m.%Y"))).grid(
                    row=0, column=2, padx=6)
                lab = ttk.Label(cell, style="Hint.TLabel")
                lab.grid(row=0, column=3, sticky="w")
                cell.columnconfigure(0, weight=0)      # при нехватке места сжимается подсказка,
                cell.columnconfigure(3, weight=1)      # а не поле даты
                v.trace_add("write", lambda *a, v=v, lab=lab: self._preview_date(v, lab))
                self._preview_date(v, lab)
                self.field_widgets[key] = ("var", v)
            elif kind == "serial":
                v = tk.StringVar()
                ttk.Entry(cell, textvariable=v).grid(row=0, column=0, sticky="ew")
                ttk.Button(cell, text="Следующий №", command=lambda v=v: self._next_serial(v)).grid(
                    row=0, column=1, padx=(6, 0))
                lab = ttk.Label(cell, style="Hint.TLabel")
                lab.grid(row=1, column=0, columnspan=2, sticky="w")
                upd = lambda *a, v=v, lab=lab: lab.configure(text=self._serial_hint(v.get()))
                v.trace_add("write", upd)
                self._qty_traces.append(self.qty.trace_add("write", upd))
                self.field_widgets[key] = ("var", v)
            elif kind == "makets":
                mk = MaketRows(cell, on_change=self._on_makets_change,
                               on_number=self._maket_number_changed)
                self.field_widgets[key] = ("makets", mk)
            elif kind == "image":
                imgs = MaketImages(cell, self, key)
                tip = ttk.Label(cell, text="?", style="Hint.TLabel", cursor="question_arrow")
                tip.grid(row=0, column=5, padx=(4, 0))
                tip.bind("<Button-1>", lambda e: messagebox.showinfo(
                    APP_NAME, "Макет — картинка (PNG, JPG…) или PDF. У PDF выбирается страница и "
                              "обрезка («Обрезать…»): в паспорт идёт выбранная область, 300 dpi.\n"
                              "Макет вписывается в поле пропорционально и встаёт по центру.\n"
                              "Полей столько, сколько сторон у изделия, и свой набор на каждый "
                              "номер макета. Сторона 2 встанет под первой через одну строку.\n"
                              "Файлы ищутся сами рядом с проектом и в папке «Макеты» по номеру макета.\n"
                              "Пусто — в паспорте будет «Нет актуальной информации. Ожидается»."))
                self.field_widgets[key] = ("image", imgs)
            else:
                v = tk.StringVar()
                ttk.Entry(cell, textvariable=v).grid(row=0, column=0, sticky="ew")
                self.field_widgets[key] = ("var", v)
            if key in ("АРТИКУЛ", "СЕРИЙНЫЙ_НОМЕР"):
                self.field_widgets[key][1].trace_add("write", lambda *a: self._update_name_preview())
            r += 1
        self._theme_text_widgets()
        self._refresh_makets()
        self.fields_scroll.fit_width()

    def _qty(self):
        try:
            return max(1, int(self.qty.get()))
        except ValueError:
            return 1

    def _qty_changed(self):
        """Тираж изделия поменяли руками — разница уходит в последнюю строку макетов."""
        mk = self._makets()
        if mk is not None:
            mk.set_total(self._qty())
        self._update_name_preview()

    def _restore_makets(self, it):
        """Номера макетов, деление тиража и картинки: из «<проект>.паспорта.json»,
        иначе один макет из проекта и автопоиск файлов по его номеру."""
        mk, imgs = self._makets(), self._images()
        saved = (self.store.get(it).get("макеты") if self.store else None) or []
        if saved and isinstance(saved[0], str):          # старый формат: список картинок
            saved = [{"номер": str(it.get("номер_макета") or ""), "тираж": it.get("тираж", 1),
                      "картинки": list(saved)}]
        total = max(1, int(it.get("тираж") or 1))
        if not saved or sum(int(x.get("тираж") or 1) for x in saved) != total:
            saved = [{"номер": str(it.get("номер_макета") or ""), "тираж": total, "картинки": None}]
        if mk is not None:
            mk.total = total
            mk.set("\n".join(f'{x.get("номер", "")}|{int(x.get("тираж") or 1)}' for x in saved))
        if imgs is not None:
            specs = []
            for x in saved:
                pics = x.get("картинки")
                if pics is None:                          # автопоиск по номеру макета
                    pics = self._find_maket_images(it, x.get("номер", ""))
                specs.append((list(pics) + ["", ""])[:2])
            imgs.set_specs(specs)
        self._refresh_makets()

    def _serial_hint(self, first):
        first = first.replace("№", "").strip()
        q = self._qty()
        if not first:
            return "«Следующий №» — продолжит нумерацию по журналу"
        if q > 1 and first.isdigit():
            return f"В паспорте: {serial_range(first, q)}  ({q} шт)"
        return ""

    def _preview_months(self, v, lab):
        try:
            lab.configure(text="= " + engine.format_warranty(int(v.get())))
        except Exception:
            lab.configure(text="введите число")

    def _preview_date(self, v, lab):
        if not v.get().strip():
            lab.configure(text="без даты — место для записи от руки")
            return
        try:
            lab.configure(text="= " + engine.format_sale_date(engine.parse_date(v.get())))
        except Exception:
            lab.configure(text="формат ДД.ММ.ГГГГ")

    def _focus_field(self, key):
        w = getattr(self, "field_entries", {}).get(key)
        if w is not None:
            try:
                w.focus_set()
                w.select_range(0, "end")
            except Exception:
                pass

    def _get(self, key):
        kind, w = self.field_widgets[key]
        if kind == "text":
            return w.get("1.0", "end").rstrip("\n")
        return w.get()

    def _get_or(self, key, default=""):
        return self._get(key) if key in self.field_widgets else default

    def _set(self, key, value):
        if key not in self.field_widgets:
            return
        kind, w = self.field_widgets[key]
        if kind == "text":
            w.edit_separator()                     # «Из наименования» можно отменить Ctrl+Z
            w.delete("1.0", "end")
            w.insert("1.0", value or "")
            w.edit_separator()
        else:
            w.set(value or "")

    def _title_from_name(self):
        self._set("НАИМЕНОВАНИЕ_ТИТУЛ", calc.title_lines(self._get_or("НАИМЕНОВАНИЕ"),
                                                          self._get_or("АРТИКУЛ"),
                                                          self.current_type["тип"], self.rules,
                                                          *self.title_box))

    def _issued_serials(self):
        issued = []
        for r in journal_rows():
            issued += serial_numbers(r.get("Серийный номер", ""))
        if self.store:
            issued += self.store.all_serials()
        issued.append(str(self.current_type.get("последний_серийный", "")))
        return issued

    def _next_serial(self, v):
        v.set(next_serial("", self._issued_serials()))

    def _find_maket_images(self, item, num=None):
        """Ищет файлы макета рядом с проектом (и в папке «Макеты») по номеру макета:
        «М-0457.png», «М-0457_1.jpg», «М-0457 сторона 2.png», «М-0457.pdf» …"""
        num = str(num if num is not None else item.get("номер_макета") or "").strip()
        if not num or not self.project:
            return []
        norm = lambda t: re.sub(r"[\s_\-–№]+", "", t.lower())
        key = norm(num)
        base = Path(self.project["файл"]).resolve().parent
        found = []
        exts = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff") \
            + ((".pdf",) if pdf_maket.available() else ())
        for d in (base, base / "Макеты", base / "макеты", base / "Maket", base / "Makets"):
            if d.is_dir():
                for f in sorted(d.iterdir()):
                    if f.suffix.lower() in exts and key in norm(f.stem) and f not in found:
                        found.append(f)
        sides = max(1, int(item.get("сторон") or 1))
        specs = []
        for f in found:
            if f.suffix.lower() != ".pdf":
                specs.append(str(f))
                continue
            try:
                n = pdf_maket.page_count(f)
            except Exception:
                continue
            for page in range(1, min(n, sides - len(specs)) + 1):   # стр. 1 → сторона 1, стр. 2 → 2
                specs.append(pdf_maket.make_spec(f, page))
            if len(specs) >= sides:
                break
        return specs[:sides]

    def _choose_image(self, imgs, maket, side):
        f = filedialog.askopenfilename(
            title="Макет — картинка или PDF",
            filetypes=[("Макеты", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.pdf"),
                       ("PDF", "*.pdf"), ("Изображения", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff")])
        if not f:
            return
        if f.lower().endswith(".pdf"):
            self._crop_maket(imgs, maket, side, path=f)
            return
        imgs.var(maket, side).set(f)
        self._remember_makets()

    def _crop_maket(self, imgs, maket, side, path=None):
        """Окно предпросмотра PDF: страница и рамка обрезки."""
        if not pdf_maket.available():
            messagebox.showerror(APP_NAME, "Для макетов в PDF нужен пакет pypdfium2.\n"
                                           "Установите: pip install pypdfium2\n"
                                           "(«Запуск_Windows.bat» и «Запуск_Mac.command» ставят его сами).")
            return
        cur = imgs.var(maket, side).get().strip()
        page, box = (side, None) if path else pdf_maket.parse_spec(cur)[1:]
        path = path or pdf_maket.parse_spec(cur)[0]
        if not path or not Path(path).exists():
            messagebox.showerror(APP_NAME, f"Файл не найден:\n{path}")
            return
        try:
            n = pdf_maket.page_count(path)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Не удалось открыть PDF:\n{path}\n\n{e}")
            return
        titles = self._maket_titles()
        name = titles[maket] if maket < len(titles) else f"Макет {maket + 1}"
        res = pdf_maket.crop_dialog(self, path, min(page or 1, n), box,
                                    title=f"{name}, сторона {side}" if self._sides() == 2 else name)
        if not res:
            return
        page, box = res
        imgs.var(maket, side).set(pdf_maket.make_spec(path, page, box))
        # двусторонний макет одним PDF: вторая сторона — следующая страница
        if side == 1 and imgs.sides == 2 and not imgs.var(maket, 2).get().strip() and n >= page + 1:
            same = pdf_maket.page_size(path, page) == pdf_maket.page_size(path, page + 1)
            imgs.var(maket, 2).set(pdf_maket.make_spec(path, page + 1, box if same else None))
            self.status.set(f"Сторона 2 — страница {page + 1} того же PDF"
                            + (" с той же обрезкой." if same and box else ".")
                            + " Изменить — «Обрезать…» у стороны 2.")
        self._remember_makets()

    def _clear_maket(self, imgs, maket, side):
        imgs.var(maket, side).set("")
        self._remember_makets()

    def _remember_makets(self):
        """Номера макетов, деление тиража и выбранные картинки (с обрезкой PDF)
        запоминаются в «<проект>.паспорта.json»."""
        if not (self.store and self.current_item):
            return
        imgs, mk = self._images(), self._makets()
        if imgs is None:
            return
        specs = imgs.specs()
        rows = mk.items() if mk is not None else [{"номер": "", "тираж": self._qty()}]
        data = []
        for i, row in enumerate(rows):
            data.append({**row, "картинки": specs[i] if i < len(specs) else ["", ""]})
        self.store.set_makets(self.current_item, data)

    def _maket_cache_dir(self):
        if self.project:
            f = Path(self.project["файл"])
            return f.parent / f"{f.stem}.макеты"
        return APP_DIR / "Макеты из PDF"

    # ---------------------------------------------------------- таблица
    def _table_rows(self):
        return [list(self.tree.item(i, "values"))[:2] for i in self.tree.get_children()]

    def _table_rows_full(self):
        return [list(self.tree.item(i, "values")) for i in self.tree.get_children()]

    def _set_table(self, rows):
        self.tree.delete(*self.tree.get_children())
        for row in rows:
            p, v = row[0], row[1]
            src = row[2] if len(row) > 2 else "шаблон"
            self.tree.insert("", "end", values=(p, v, src), tags=(src,))
        self._refresh_makets()

    def _sides(self):
        if self.current_item and self.current_item.get("сторон"):
            try:
                return int(self.current_item["сторон"])
            except (TypeError, ValueError):
                pass
        for i in self.tree.get_children():
            p, v = self.tree.item(i, "values")[:2]
            if str(p).strip().lower().startswith("кол-во сторон"):
                m = re.search(r"\d", str(v))
                if m:
                    return int(m.group(0))
        return 1

    def _makets(self):
        """Виджет номеров макетов (у динамики его нет)."""
        w = self.field_widgets.get("НОМЕР_МАКЕТА")
        return w[1] if w and w[0] == "makets" else None

    def _images(self):
        for kind, w in self.field_widgets.values():
            if kind == "image":
                return w
        return None

    def _maket_titles(self):
        mk = self._makets()
        if mk is None:
            return [""]
        return [n or f"Макет {i + 1}" for i, n in enumerate(mk.numbers())]

    def _refresh_makets(self):
        """Поля картинок: по набору на каждый макет, в наборе — по числу сторон."""
        imgs = self._images()
        if imgs is not None:
            imgs.show(self._sides(), self._maket_titles())
        self.fields_scroll.fit_width()

    def _on_makets_change(self):
        self._refresh_makets()
        self._update_name_preview()

    def _maket_number_changed(self, index):
        """Номер макета ввели вручную — подставим картинки этого макета, если поля пустые."""
        imgs, mk = self._images(), self._makets()
        if imgs is None or mk is None or index >= mk.count():
            return
        num = mk.numbers()[index]
        if not num or any(imgs.specs()[index] if index < len(imgs.specs()) else []):
            return
        found = self._find_maket_images(self.current_item or {}, num)
        for side, spec in enumerate(found, start=1):
            imgs.var(index, side).set(spec)
        if found:
            self._remember_makets()

    def _table_from_template(self):
        if messagebox.askyesno(APP_NAME, "Заменить таблицу значениями из шаблона?"):
            self._set_table(self.tpl_table)

    def _row_add(self):
        sel = self.tree.selection()
        idx = self.tree.index(sel[0]) + 1 if sel else "end"
        iid = self.tree.insert("", idx, values=("Новый параметр", "", "вручную"), tags=("вручную",))
        self.tree.selection_set(iid)
        self.tree.see(iid)
        self._edit_cell(None, iid=iid, col="#1")

    def _row_del(self):
        for i in self.tree.selection():
            self.tree.delete(i)

    def _row_move(self, d):
        sel = self.tree.selection()
        if sel:
            self.tree.move(sel[0], "", max(0, self.tree.index(sel[0]) + d))

    def _edit_cell(self, event, iid=None, col=None):
        if event is not None:
            if self.tree.identify_region(event.x, event.y) != "cell":
                return
            iid = self.tree.identify_row(event.y)
            col = self.tree.identify_column(event.x)
        if not iid or col == "#3":
            return
        x, y, w, h = self.tree.bbox(iid, col)
        ci = int(col[1:]) - 1
        vals = list(self.tree.item(iid, "values"))
        e = ttk.Entry(self.tree)
        e.insert(0, vals[ci])
        e.select_range(0, "end")
        e.place(x=x, y=y, width=w, height=h)
        e.focus_set()

        def commit(_=None):
            if e.get() != vals[ci]:
                vals[ci] = e.get()
                vals[2] = "вручную"
                self.tree.item(iid, values=vals, tags=("вручную",))
                self._refresh_makets()
            e.destroy()

        e.bind("<Return>", commit)
        e.bind("<FocusOut>", commit)
        e.bind("<Escape>", lambda _: e.destroy())

    # ---------------------------------------------------------- состояние
    def _collect_state(self):
        return {"значения": {k: self._get(k) for k in self.field_widgets},
                "таблица": self._table_rows_full(),
                "проект": self.project_name.get(), "тираж": self.qty.get()}

    def _apply_state(self, st):
        vals = st.get("значения", {})
        for k in self.field_widgets:
            if k in vals:
                self._set(k, vals[k])
        if "ДАТА_ПРОДАЖИ" in self.field_widgets and not vals.get("ДАТА_ПРОДАЖИ"):
            self._set("ДАТА_ПРОДАЖИ", date.today().strftime("%d.%m.%Y"))
        self._set_table(st.get("таблица") or self.tpl_table)
        if not self.project:
            self.project_name.set(st.get("проект", ""))
        self.qty.set(st.get("тираж", "1") or "1")

    def _save_data(self):
        f = filedialog.asksaveasfilename(title="Сохранить данные паспорта", defaultextension=".json",
                                         filetypes=[("Данные паспорта", "*.json")],
                                         initialfile=safe_name(self._get_or("АРТИКУЛ", "паспорт")) + ".json")
        if f:
            data = {"тип": self.current_type["тип"], **self._collect_state()}
            Path(f).write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")
            self.status.set("Данные сохранены: " + f)

    def _load_data(self):
        f = filedialog.askopenfilename(title="Открыть данные паспорта",
                                       filetypes=[("Данные паспорта", "*.json")])
        if not f:
            return
        try:
            data = json.loads(Path(f).read_text("utf-8"))
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Не удалось прочитать файл:\n{e}")
            return
        self._deselect_item()
        if data.get("тип") in [t["тип"] for t in self.types] and data["тип"] != self.type_var.get():
            self.type_var.set(data["тип"])
            self._switch_type()
        self._apply_state(data)
        self.project_name.set(data.get("проект", self.project_name.get()))
        self.status.set("Данные загружены: " + f)

    # ---------------------------------------------------------- имя файла
    def _file_base(self, maket=None, qty=None):
        """«N шт. Артикул НомерМакета Проект» (статика), «N шт. Артикул СерийныйНомер Проект» (динамика).
        maket / qty — для конкретной строки макетов; без них берётся первая строка."""
        art = self._get_or("АРТИКУЛ").strip() or self.current_type["тип"]
        mk = self._makets()
        if qty is None:
            qty = (mk.items()[0]["тираж"] if mk is not None and mk.count() else self._qty())
        if "СЕРИЙНЫЙ_НОМЕР" in self.field_widgets:
            ident = serial_range(self._get("СЕРИЙНЫЙ_НОМЕР"), qty, sep="-")
        elif maket is not None:
            ident = str(maket).strip()
        else:
            ident = (mk.numbers()[0].strip() if mk is not None and mk.count() else "")
        parts = [f"{qty} шт.", art, ident, self.project_name.get().strip()]
        return safe_name(" ".join(p for p in parts if p))

    def _out_base(self, maket=None, qty=None, index=0):
        """Имя файла; при ручной правке и нескольких макетах к имени добавляется номер макета."""
        if getattr(self, "name_override", ""):
            name = safe_name(self.name_override)
            if index and maket:
                name = safe_name(f"{name} {maket}")
            elif index:
                name = safe_name(f"{name} ({index + 1})")
            return name
        return self._file_base(maket, qty)

    def _update_name_preview(self):
        try:
            manual = bool(getattr(self, "name_override", ""))
            mk = self._makets()
            rows = mk.items() if mk is not None else []
            self.name_caption.configure(text="Имя файла (изменено вручную):" if manual else
                                             ("Имена файлов:" if len(rows) > 1 else "Имя файла:"))
            first = self._out_base(rows[0]["номер"] if rows else None,
                                   rows[0]["тираж"] if rows else None) + ".docx"
            if len(rows) > 1:
                first += f"   и ещё {len(rows) - 1}"
            self.name_preview.configure(text=first)
        except Exception:
            pass

    def _reset_name(self):
        self.name_override = ""
        self._update_name_preview()

    def _edit_name(self, _=None):
        """Двойной щелчок по имени файла — правка. Пустое поле — снова автоматическое имя."""
        if getattr(self, "_name_entry", None):
            return
        e = ttk.Entry(self.name_preview.master)
        e.insert(0, self._out_base())
        e.select_range(0, "end")
        info = self.name_preview.grid_info()
        e.grid(row=info["row"], column=info["column"], columnspan=3,   # имя + подсказка + пусто
               sticky="ew", padx=6, pady=(6, 0))
        self.name_preview.grid_remove()
        self.name_hint.grid_remove()
        e.focus_set()
        self._name_entry = e

        def done(save):
            if save:
                txt = e.get().strip()
                if txt.lower().endswith(".docx"):
                    txt = txt[:-5]
                self.name_override = "" if (not txt or txt == self._file_base()) else txt
            e.destroy()
            self._name_entry = None
            self.name_preview.grid()
            self.name_hint.grid()
            self._update_name_preview()

        e.bind("<Return>", lambda _e: done(True))
        e.bind("<KP_Enter>", lambda _e: done(True))
        e.bind("<FocusOut>", lambda _e: done(True))
        e.bind("<Escape>", lambda _e: done(False))

    # ---------------------------------------------------------- скачать паспорт
    def _maket_paths(self, specs):
        """Пути картинок макета; PDF превращается в PNG 300 dpi. None — ошибка (уже показана)."""
        paths = []
        for x in [str(v).strip() for v in specs if str(v).strip()]:
            src = pdf_maket.parse_spec(x)[0]
            if not Path(src).exists():
                messagebox.showerror(APP_NAME, f"Файл макета не найден:\n{src}")
                return None
            if pdf_maket.is_pdf(x):
                try:
                    self.status.set("Готовлю макет из PDF…")
                    self.update_idletasks()
                    x = pdf_maket.render_png(x, self._maket_cache_dir())
                except Exception as e:
                    messagebox.showerror(APP_NAME, f"Не удалось подготовить макет из PDF:\n{src}\n\n{e}")
                    return None
            paths.append(x)
        return paths

    def _generate(self):
        values, missing = {}, []
        no_date = False
        q = self._qty()
        img_key = next((k for k, (kind, _) in self.field_widgets.items() if kind == "image"), None)
        for key in self.field_widgets:
            kind = self.field_widgets[key][0]
            if kind in ("image", "makets"):          # разбираются по строкам макетов, ниже
                continue
            raw = self._get(key).strip()
            if not raw and key == "ДАТА_ПРОДАЖИ":
                no_date = True
                values[key] = BLANK_DATE
                continue
            if not raw:
                missing.append(FIELD_DEFS.get(key, (key,))[0].replace("\n", " "))
                continue
            if key == "ГАРАНТИЯ":
                try:
                    values[key] = engine.format_warranty(int(raw))
                except ValueError:
                    messagebox.showerror(APP_NAME, "Гарантийный срок — целое число месяцев.")
                    return
            elif key == "ДАТА_ПРОДАЖИ":
                try:
                    values[key] = engine.format_sale_date(engine.parse_date(raw))
                except ValueError as e:
                    messagebox.showerror(APP_NAME, f"Дата продажи: {e}")
                    return
            elif key == "НАИМЕНОВАНИЕ_ТИТУЛ":
                text = "\n".join(l.strip() for l in raw.splitlines() if l.strip())
                values[key] = text
                width, pt = getattr(self, "title_box", (None, None))
                if width and pt:        # титул только на тёмно-синем поле: кегль под ширину поля
                    fit = calc.title_fit(text, width, pt)
                    if fit < pt:
                        values[key] = {"text": text, "pt": fit}
            elif key == "СЕРИЙНЫЙ_НОМЕР":
                values[key] = serial_range(raw, q)
            else:
                values[key] = raw
        if missing:
            messagebox.showwarning(APP_NAME, "Заполните поля:\n• " + "\n• ".join(missing))
            return
        if no_date and not messagebox.askyesno(
                APP_NAME, "Дата продажи не указана. Создать паспорт без даты?\n"
                          "В паспорте останется место для записи от руки."):
            self._focus_field("ДАТА_ПРОДАЖИ")
            return

        # ---- строки макетов: на каждую — свой паспорт
        mk, imgs = self._makets(), self._images()
        if mk is not None:
            if not mk.ok():
                messagebox.showwarning(APP_NAME, f"Распределено {mk.distributed()} шт. из {q}.\n"
                                                 f"Сумма тиражей по макетам должна совпадать "
                                                 f"с тиражом изделия.")
                return
            rows = mk.items()
            nums = [x["номер"] for x in rows]
            dups = {n for n in nums if n and nums.count(n) > 1}
            if dups and not messagebox.askyesno(
                    APP_NAME, "Номер макета повторяется: " + ", ".join(sorted(dups)) +
                              ".\nПаспорты получат одинаковые имена файлов. Продолжить?"):
                return
            empty = [i + 1 for i, n in enumerate(nums) if not n]
            if empty and not messagebox.askyesno(
                    APP_NAME, ("Номер макета не указан — в паспорте будет «Б/Н».\nПродолжить?"
                               if len(rows) == 1 else
                               f"Номер макета не указан в строках: {', '.join(map(str, empty))} — "
                               f"в паспорте будет «Б/Н».\nПродолжить?")):
                return
        else:
            rows = [{"номер": "", "тираж": q}]
        specs = imgs.specs() if imgs is not None else [["", ""]]
        sides = self._sides()
        jobs = []
        for i, row in enumerate(rows):
            pics = specs[i] if i < len(specs) else ["", ""]
            paths = self._maket_paths(pics)
            if paths is None:
                return
            name = row["номер"] or ("" if mk is None else "Б/Н")
            if img_key:
                if paths and sides == 2 and len(paths) == 1 and not messagebox.askyesno(
                        APP_NAME, f"{'Макет ' + (row['номер'] or str(i + 1)) + ': и' if len(rows) > 1 else 'И'}"
                                  f"зделие двустороннее, а выбрано изображение только одной стороны.\n"
                                  f"Продолжить?"):
                    return
                if not paths and not messagebox.askyesno(
                        APP_NAME, f"{'Макет ' + (row['номер'] or str(i + 1)) + ': изображение' if len(rows) > 1 else 'Изображение'}"
                                  f" макета не выбрано — в Приложении 1 будет написано "
                                  f"«Нет актуальной информации. Ожидается».\nПродолжить?"):
                    return
            jobs.append({"номер": name, "тираж": row["тираж"],
                         "картинки": {img_key: paths} if img_key else {},
                         "имя": self._out_base(row["номер"], row["тираж"], i)})

        serial = values.get("СЕРИЙНЫЙ_НОМЕР", "")
        if serial:
            mine = set(serial_numbers(serial))
            own = set(serial_numbers(self.store.get(self.current_item).get("серийный_номер", ""))) \
                if (self.store and self.current_item) else set()
            clash = []
            for r in journal_rows():
                nums_r = set(serial_numbers(r.get("Серийный номер", "")))
                if nums_r & mine and not nums_r <= own:
                    clash.append(r)
            if clash and not messagebox.askyesno(
                    APP_NAME, f"Серийные номера {serial} пересекаются с уже выданными:\n"
                              f"{clash[-1].get('Серийный номер')} — {clash[-1].get('Создан', '')}, "
                              f"{Path(clash[-1].get('Файл', '')).name}\n\nВсё равно создать паспорт?"):
                return

        # рукописная дата подписания (опция «Печать, подпись и дата»): та же дата продажи,
        # в талоне печатная дата тогда не ставится — на её месте дата от руки
        hand_date = ""
        if self.use_stamps.get() and not no_date and "ДАТА_ПРОДАЖИ" in self.field_widgets:
            hand_date = engine.parse_date(self._get("ДАТА_ПРОДАЖИ")).strftime("%d.%m.%Y")
            places = engine.stamp_places(docx.Document(self._template_path()))
            lib_now = stamp_library()[0]
            if "ДАТА_ЗАЛИВАКИН" in places and lib_now.get("ПОДПИСЬ_ЗАЛИВАКИН") and handwriting_font():
                values["ДАТА_ПРОДАЖИ"] = ""

        out_dir = Path(self.out_dir.get().strip() or APP_DIR)
        out_dir.mkdir(parents=True, exist_ok=True)
        for job in jobs:                              # файлы: заменить / копия / отмена
            out = out_dir / f'{job["имя"]}.docx'
            if out.exists():
                ans = messagebox.askyesnocancel(
                    APP_NAME, f"Файл уже есть:\n{out.name}\n\nДа — заменить\nНет — сохранить копию\n"
                              f"Отмена — не скачивать")
                if ans is None:
                    return
                if ans is False:
                    n = 2
                    while out.exists():
                        out = out_dir / f'{job["имя"]} ({n}).docx'
                        n += 1
            job["файл"] = out

        made, notes = [], []
        for i, job in enumerate(jobs, start=1):
            self.status.set(f"Создаю паспорт {i} из {len(jobs)}…" if len(jobs) > 1 else "Создаю паспорт…")
            self.update_idletasks()
            vals = dict(values)
            if mk is not None:
                vals["НОМЕР_МАКЕТА"] = job["номер"]
                if "СЕРИЙНЫЙ_НОМЕР" in vals:
                    vals["СЕРИЙНЫЙ_НОМЕР"] = serial_range(self._get("СЕРИЙНЫЙ_НОМЕР"), job["тираж"])
            res = self._make_passport(vals, job, hand_date)
            if res is None:
                return
            made.append(res)
            notes += [f'{job["файл"].name}: {n}' if len(jobs) > 1 else n for n in res["заметки"]]

        count_note = ""
        if self.store and self.current_item is not None:
            for res in made:
                rec = self.store.register(self.current_item, res["файл"], serial)
            self._remember_makets()
            idx = self.items_list.selected
            self._refresh_list(keep=True)
            if idx is not None:
                self.items_list.select(idx, notify=False)
                self._on_item_info_refresh(idx)
            count_note = f"\nСкачан {times(rec['скачан'])}."
        self._save_settings()
        self.status.set(("Паспорта созданы: " if len(made) > 1 else "Паспорт создан: ")
                        + ", ".join(r["файл"].name for r in made))
        if len(made) == 1:
            msg = f'Паспорт создан:\n{made[0]["файл"].name}{made[0]["pdf"]}{count_note}'
        else:
            msg = (f"Создано паспортов: {len(made)}\n" +
                   "\n".join(f'• {r["файл"].name}{r["pdf"]}' for r in made) + count_note)
        if notes:
            messagebox.showwarning(APP_NAME, msg + "\n\nВнимание:\n• " + "\n• ".join(notes))
        else:
            messagebox.showinfo(APP_NAME, msg)
        if self.open_after.get():
            open_path(out_dir)

    def _make_passport(self, values, job, hand_date):
        """Один паспорт: Word, печати и подписи, PDF и оглавление, запись в журнал.
        → {"файл", "pdf", "заметки"} или None при ошибке."""
        out = job["файл"]
        try:
            lib, cfg = stamp_library()
            stamps = {"включено": self.use_stamps.get(), "файлы": lib, "настройки": cfg,
                      "дата": True, "дата_текст": hand_date, "шрифт_даты": handwriting_font(),
                      "цифры": own_digits()}
            rep = engine.generate(self._template_path(), out, values,
                                  self._table_rows() if self.tpl_table else None,
                                  job["картинки"], stamps)
        except PermissionError:
            messagebox.showerror(APP_NAME, f"Файл занят (открыт в Word?):\n{out}")
            self.status.set("Ошибка")
            return None
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Ошибка при создании паспорта:\n{e}")
            self.status.set("Ошибка")
            return None
        notes = []
        if rep["остались незаполненными"]:
            notes.append("Остались метки без значения: " + ", ".join(rep["остались незаполненными"]))
        if rep["таблица"] is False:
            notes.append("Таблица техданных в шаблоне не найдена.")
        for k, ok in rep["изображения"].items():
            if not ok:
                notes.append(f"Рисунок с меткой {{{{{k}}}}} не найден.")
        if self.current_item and self.project and self.project.get("формат") \
                and self.current_item.get("масса") in (None, ""):
            notes.append("Вес в паспорте ориентировочный (мощность посчитана по правилам Конструктора): "
                         "в файле проекта нет расчёта Конструктора. Пересохраните проект "
                         "в Конструкторе V40.04.2 или новее.")
        if self.use_stamps.get():
            places = engine.stamp_places(docx.Document(self._template_path()))
            done = {k for k, *_ in rep["печати"]}
            if "ПЕЧАТЬ" in places and "ПЕЧАТЬ" not in done:
                notes.append("Нет файла печати в папке «Печати и подписи» — печать не поставлена.")
            skipped = [k.split("_", 1)[-1].capitalize() for k in places
                       if k.startswith("ПОДПИСЬ") and k not in done]
            if hand_date and not rep.get("даты"):
                notes.append("Дата от руки не поставлена: нет рукописного шрифта в папке «Шрифты» "
                             "(Caveat) или файла подписи.")
            if skipped:   # подписи без файла — без всплывающего окна, только в строке состояния
                self.after(50, lambda: self.status.set(
                    self.status.get() + "   ·   без подписи (нет файла): " + ", ".join(skipped)))
        pdf_note = ""
        if self.print_ready.get():
            try:
                pr = book.prepare_for_print(out, log=lambda m: (self.status.set(m), self.update_idletasks()))
                pdf_note = (f"\nPDF: {pr['pdf'].name}\nСтраниц: {pr['страниц']}"
                            f"{' (добавлено пустых: ' + str(pr['пустых']) + ')' if pr['пустых'] else ''}"
                            f", оглавление обновлено, сборка: {pr['сборщик']}")
                if not pr["кратно"]:
                    notes.append(f"Страниц {pr['страниц']} — не кратно 4, проверьте PDF.")
            except Exception as e:
                notes.append(f"Подготовка к печати не выполнена: {e}")
        else:
            # без «Для типографии»: номера страниц в «Содержании» всё равно пересчитываем
            self.status.set("Пересчитываю оглавление…")
            self.update_idletasks()
            try:
                book.refresh_toc(out, keep_pdf=self.make_pdf.get())
                if self.make_pdf.get():
                    pdf_note = "\nPDF: " + out.with_suffix(".pdf").name
            except Exception as e:
                if self.make_pdf.get():
                    notes.append(str(e))
        journal_add({
            "Создан": datetime.now().strftime("%d.%m.%Y %H:%M"),
            "Проект": self.project_name.get().strip(),
            "Префикс": self.current_item.get("префикс", "") if self.current_item else "",
            "Тип": self.current_type["тип"],
            "Наименование": values.get("НАИМЕНОВАНИЕ", ""),
            "Артикул": values.get("АРТИКУЛ", ""),
            "Серийный номер": values.get("СЕРИЙНЫЙ_НОМЕР", ""),
            "Номер макета": values.get("НОМЕР_МАКЕТА", ""),
            "Тираж": job["тираж"],
            "Гарантия, мес": self._get_or("ГАРАНТИЯ"),
            "Дата продажи": self._get_or("ДАТА_ПРОДАЖИ"),
            "Печать, подпись и дата": ("да" + (", дата" if rep.get("даты") else ""))
            if self.use_stamps.get() and rep["печати"] else "",
            "Файл": str(out),
        })
        return {"файл": out, "pdf": pdf_note, "заметки": notes}

    def _on_item_info_refresh(self, idx):
        rec = self.store.get(self.current_item)
        txt = self.calc_label.cget("text").split("   ·   скачан")[0]
        self.calc_label.configure(text=f'{txt}   ·   скачан {times(rec["скачан"])}, '
                                       f'последний раз {rec.get("последний", "")}')

    def _choose_out(self):
        d = filedialog.askdirectory(title="Папка для паспортов", initialdir=self.out_dir.get())
        if d:
            self.out_dir.set(d)

    def _show_journal(self):
        rows = journal_rows()
        win = tk.Toplevel(self)
        win.title("Журнал выданных паспортов")
        win.geometry("1250x440")
        cols = JOURNAL_COLUMNS
        tv = ttk.Treeview(win, columns=cols, show="headings")
        widths = [110, 140, 60, 100, 260, 170, 170, 90, 55, 70, 85, 90, 300]
        for c, w in zip(cols, widths):
            tv.heading(c, text=c)
            tv.column(c, width=w, anchor="w")
        for r in reversed(rows):
            tv.insert("", "end", values=[r.get(c, "") for c in cols])
        tv.pack(fill="both", expand=True, padx=8, pady=8)
        tv.bind("<Double-1>", lambda e: (lambda s: open_path(tv.item(s[0], "values")[-1])
                                         if s else None)(tv.selection()))
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(bar, text=f"Записей: {len(rows)}. Двойной щелчок — открыть паспорт.",
                  style="Hint.TLabel").pack(side="left")
        ttk.Button(bar, text="Открыть журнал в Excel",
                   command=lambda: open_path(JOURNAL_FILE) if JOURNAL_FILE.exists() else None).pack(side="right")


def _startup_error(exc):
    """Ошибка при запуске: текст — в «ошибка_запуска.txt» рядом с программой и на экран."""
    import traceback
    text = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    log = APP_DIR / "ошибка_запуска.txt"
    try:
        log.write_text(f"{APP_NAME} {APP_VERSION}\n{datetime.now():%d.%m.%Y %H:%M}\n"
                       f"Python {sys.version}\n{sys.platform}\nПапка: {APP_DIR}\n\n{text}", "utf-8")
    except Exception:
        pass
    print(text, file=sys.stderr)
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_NAME, f"Программа не запустилась:\n\n{exc}\n\n"
                                       f"Подробности сохранены в файле:\n{log}")
        root.destroy()
    except Exception:
        pass


if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception as e:      # noqa: BLE001
        _startup_error(e)
        sys.exit(1)
