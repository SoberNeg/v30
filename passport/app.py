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
import project_import
from widgets import PALETTES, GradientList, ScrollFrame

APP_NAME = "Утилита заполнения паспортов"
APP_VERSION = "V1.8"

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
JOURNAL_FILE = APP_DIR / "Журнал_паспортов.csv"
JOURNAL_COLUMNS = ["Создан", "Проект", "Префикс", "Тип", "Наименование", "Артикул",
                   "Серийный номер", "Номер макета", "Тираж", "Гарантия, мес",
                   "Дата продажи", "Печать и подпись", "Файл"]

FIELD_DEFS = {
    "НАИМЕНОВАНИЕ": ("Наименование изделия", "entry"),
    "НАИМЕНОВАНИЕ_ТИТУЛ": ("Наименование на титульном листе\n(по строкам)", "text"),
    "АРТИКУЛ": ("Артикул", "entry"),
    "СЕРИЙНЫЙ_НОМЕР": ("Серийный номер (первый)", "serial"),
    "НОМЕР_МАКЕТА": ("Номер макета", "entry"),
    "ГАРАНТИЯ": ("Гарантийный срок, месяцев", "months"),
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
            try:
                conf.write_text(json.dumps(engine.STAMP_DEFAULTS, ensure_ascii=False, indent=2), "utf-8")
            except Exception:
                pass
    return lib, cfg


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

    def all_serials(self):
        out = []
        for rec in self.data["изделия"].values():
            out += serial_numbers(rec.get("серийный_номер", ""))
        return out


class ImagePair:
    """Поля картинок макета по числу сторон изделия (1 или 2) с интерфейсом StringVar:
    значение — пути через \\n."""

    def __init__(self):
        self.v1, self.v2 = tk.StringVar(), tk.StringVar()
        self.rows = {1: [], 2: []}
        self.sides = 2

    def show(self, sides):
        """Сколько сторон — столько полей; поле второй стороны скрывается и очищается."""
        sides = 2 if int(sides or 1) >= 2 else 1
        self.sides = sides
        for w in self.rows[2]:
            (w.grid if sides == 2 else w.grid_remove)()
        if sides == 1:
            self.v2.set("")

    def get(self):
        return "\n".join(x for x in (self.v1.get().strip(), self.v2.get().strip()) if x)

    def set(self, value):
        parts = [x for x in str(value or "").splitlines() if x.strip()] + ["", ""]
        self.v1.set(parts[0])
        self.v2.set(parts[1])


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
                ("_артикул", "Артикул", 240, "w"), ("_размер", "Размер, мм", 105, "center"),
                ("_сторон", "Сторон", 65, "center"), ("тираж", "Тираж, шт", 80, "center"),
                ("_идент", "Макет / Серийный №", 210, "w"), ("_скачан", "Скачан", 95, "center")]
        self.items_list = GradientList(proj, cols, on_select=self._on_item, stretch="_артикул")
        self.items_list.empty_text = ("Откройте сохранённый файл проекта Калькулятора экранов — "
                                      "изделия появятся списком. Щелчок по строке заполнит паспорт.")
        self.calc_label = ttk.Label(proj, style="Hint.TLabel")
        self.calc_label.pack(fill="x", side="bottom", pady=(4, 0))
        self.items_list.pack(fill="both", expand=True, pady=(6, 0))
        paned.add(proj, weight=2)

        # ---- поля + таблица
        mid = ttk.Frame(paned)
        mid.columnconfigure(0, weight=5)
        mid.columnconfigure(1, weight=6)
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
        self.tree.column("p", width=300, anchor="w")
        self.tree.column("v", width=190, anchor="w")
        self.tree.column("s", width=80, anchor="center")
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
        nf.columnconfigure(1, weight=1)
        self.name_caption = ttk.Label(nf, text="Имя файла:")
        self.name_caption.grid(row=0, column=0, sticky="w", pady=(6, 0))
        self.name_preview = ttk.Label(nf, style="Name.TLabel", cursor="xterm")
        self.name_preview.grid(row=0, column=1, sticky="w", padx=6, pady=(6, 0))
        self.name_preview.bind("<Double-Button-1>", self._edit_name)
        ttk.Label(nf, text="двойной щелчок — изменить", style="Hint.TLabel").grid(
            row=0, column=2, sticky="e", pady=(6, 0))
        opts = ttk.Frame(bottom)
        opts.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Checkbutton(opts, text="Для типографии (PDF, кратно 4 стр.)",
                        variable=self.print_ready).pack(side="left")
        ttk.Checkbutton(opts, text="PDF", variable=self.make_pdf).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(opts, text="Печать и подпись", variable=self.use_stamps).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(opts, text="Открыть папку",
                        variable=self.open_after).pack(side="left", padx=12)
        ttk.Button(opts, text="Скачать паспорт", style="Big.TButton",
                   command=self._generate).pack(side="right")
        ttk.Button(opts, text="Журнал", command=self._show_journal).pack(side="right", padx=8)
        ttk.Button(opts, text="Сохранить данные…", command=self._save_data).pack(side="right")
        ttk.Button(opts, text="Открыть данные…", command=self._load_data).pack(side="right", padx=8)
        for ch in opts.pack_slaves():           # кнопки справа не должны исчезать на узком окне
            if isinstance(ch, ttk.Button):
                ch.pack_configure(ipadx=0)

        for v in (self.project_name, self.qty):
            v.trace_add("write", lambda *a: self._update_name_preview())

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
                "_идент": ident, "_скачан": times(n) if n else "—", "_done": n > 0}

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
        if "НОМЕР_МАКЕТА" in self.field_widgets and not it.get("номер_макета"):
            self._set("НОМЕР_МАКЕТА", "")
        for k in self.tpl_img_keys:                 # картинки прошлого изделия не переносим
            if k in self.field_widgets:
                self._set(k, "\n".join(self._find_maket_images(it)))
        self._set_table(rows)
        self.qty.set(str(it.get("тираж", 1)))
        self.project_name.set(self.project["проект"])
        if it["тип"] == "Динамика МЦД":
            info = (f'Расчёт {it.get("префикс")}: модулей {v["модулей_по_ширине"]}×{v["модулей_по_высоте"]} '
                    f'= {v["модулей"]} шт, поля {v["поле_гор_мм"]} / {v["поле_верт_мм"]} мм, '
                    f'разрешение {v["px_w"]}×{v["px_h"]}, P макс {v["P_max_кВт"]} кВт, масса {v["масса"]} кг')
        else:
            info = (f'Расчёт {it.get("префикс")}: площадь {v["площадь_м2"]} м², нагрузка {v["нагрузка_Вт"]} Вт → '
                    f'блоков питания {v["блоков_питания"]}, мощность {v["P_Вт"]} Вт, масса {v["масса"]} кг')
        if it.get("масса") not in (None, "") and it.get("_расчёт"):
            psu = ", ".join(f'{calc.num(p_["qty"], 0)}×{p_["w"]} Вт' for p_ in it.get("блоки_питания") or [])
            info = (f'{it.get("префикс")}: по расчёту {it["_расчёт"]} — масса {calc.num(it["масса"], 2)} кг'
                    + (f', блоки питания {psu}' if psu else '')
                    + (f', суммарно {calc.num(it["мощность_Вт"], 0)} Вт' if it.get("мощность_Вт") else ''))
        if rec.get("скачан"):
            info += f'   ·   скачан {times(rec["скачан"])}, последний раз {rec.get("последний", "")}'
        if self.project.get("формат") and it.get("масса") in (None, ""):
            info = ("⚠ Вес и мощность ориентировочные: в файле проекта нет расчёта Конструктора — "
                    "пересохраните проект в Конструкторе V40.04.2 или новее.   ·   " + info)
        self.calc_label.configure(text=info)
        self._update_name_preview()
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
                w = tk.Text(cell, height=4, width=34, wrap="none", font=("Segoe UI", 10),
                            relief="solid", borderwidth=1)
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
                ttk.Entry(cell, textvariable=v, width=12).grid(row=0, column=0, sticky="w")
                ttk.Button(cell, text="Сегодня",
                           command=lambda v=v: v.set(date.today().strftime("%d.%m.%Y"))).grid(
                    row=0, column=1, padx=6)
                lab = ttk.Label(cell, style="Hint.TLabel")
                lab.grid(row=0, column=2, sticky="w")
                cell.columnconfigure(0, weight=0)      # при нехватке места сжимается подсказка,
                cell.columnconfigure(2, weight=1)      # а не поле даты
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
            elif kind == "image":
                pair = ImagePair()
                for i, (cap, var) in enumerate((("Сторона 1", pair.v1), ("Сторона 2", pair.v2))):
                    ws = [ttk.Label(cell, text=cap, style="Hint.TLabel"),
                          ttk.Entry(cell, textvariable=var),
                          ttk.Button(cell, text="…", width=3,
                                     command=lambda var=var: self._choose_image(var)),
                          ttk.Button(cell, text="×", width=3, command=lambda var=var: var.set(""))]
                    ws[0].grid(row=i, column=0, sticky="w")
                    ws[1].grid(row=i, column=1, sticky="ew", pady=1)
                    ws[2].grid(row=i, column=2, padx=(4, 0))
                    ws[3].grid(row=i, column=3, padx=(2, 0))
                    pair.rows[i + 1] = ws
                cell.columnconfigure(0, weight=0)
                cell.columnconfigure(1, weight=1)
                tip = ttk.Label(cell, text="?", style="Hint.TLabel", cursor="question_arrow")
                tip.grid(row=0, column=4, padx=(4, 0))
                tip.bind("<Button-1>", lambda e: messagebox.showinfo(
                    APP_NAME, "Макет вписывается в поле пропорционально и встаёт по центру.\n"
                              "Полей столько, сколько сторон у изделия. Сторона 2 встанет под первой "
                              "через одну строку.\n"
                              "Файлы ищутся сами рядом с проектом и в папке «Макеты» по номеру макета.\n"
                              "Пусто — в паспорте будет «Нет актуальной информации. Ожидается»."))
                self.field_widgets[key] = ("image", pair)
            else:
                v = tk.StringVar()
                ttk.Entry(cell, textvariable=v).grid(row=0, column=0, sticky="ew")
                self.field_widgets[key] = ("var", v)
            if key in ("АРТИКУЛ", "НОМЕР_МАКЕТА", "СЕРИЙНЫЙ_НОМЕР"):
                self.field_widgets[key][1].trace_add("write", lambda *a: self._update_name_preview())
            r += 1
        self._theme_text_widgets()
        self._refresh_image_sides()
        self.fields_scroll.fit_width()

    def _qty(self):
        try:
            return max(1, int(self.qty.get()))
        except ValueError:
            return 1

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
        try:
            lab.configure(text="= " + engine.format_sale_date(engine.parse_date(v.get())))
        except Exception:
            lab.configure(text="формат ДД.ММ.ГГГГ")

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
            w.delete("1.0", "end")
            w.insert("1.0", value or "")
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

    def _find_maket_images(self, item):
        """Ищет файлы макета рядом с проектом (и в папке «Макеты») по номеру макета:
        «М-0457.png», «М-0457_1.jpg», «М-0457 сторона 2.png» …"""
        num = str(item.get("номер_макета") or "").strip()
        if not num or not self.project:
            return []
        norm = lambda t: re.sub(r"[\s_\-–№]+", "", t.lower())
        key = norm(num)
        base = Path(self.project["файл"]).resolve().parent
        found = []
        for d in (base, base / "Макеты", base / "макеты", base / "Maket", base / "Makets"):
            if d.is_dir():
                for f in sorted(d.iterdir()):
                    if f.suffix.lower() in (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff") \
                            and key in norm(f.stem) and f not in found:
                        found.append(f)
        return [str(f) for f in found[: max(1, int(item.get("сторон") or 1))]]

    def _choose_image(self, v):
        f = filedialog.askopenfilename(title="Изображение макета",
                                       filetypes=[("Изображения", "*.png *.jpg *.jpeg *.bmp *.gif")])
        if f:
            v.set(f)

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
        self._refresh_image_sides()

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

    def _refresh_image_sides(self):
        for kind, w in self.field_widgets.values():
            if kind == "image":
                w.show(self._sides())

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
                self._refresh_image_sides()
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
    def _file_base(self, values=None):
        """Статика: Артикул_НомерМакета_Тираж_Проект; динамика: Артикул_СерийныйНомер_Тираж_Проект."""
        art = self._get_or("АРТИКУЛ").strip() or self.current_type["тип"]
        q = self._qty()
        if "СЕРИЙНЫЙ_НОМЕР" in self.field_widgets:
            ident = serial_range(self._get("СЕРИЙНЫЙ_НОМЕР"), q, sep="-")
        else:
            ident = self._get_or("НОМЕР_МАКЕТА").strip()
        parts = [art, ident, f"{q}шт", self.project_name.get().strip()]
        return safe_name(" ".join(p for p in parts if p))

    def _out_base(self):
        return safe_name(self.name_override) if getattr(self, "name_override", "") else self._file_base()

    def _update_name_preview(self):
        try:
            manual = bool(getattr(self, "name_override", ""))
            self.name_caption.configure(text="Имя файла (изменено вручную):" if manual else "Имя файла:")
            self.name_preview.configure(text=self._out_base() + ".docx")
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
        e.grid(row=info["row"], column=info["column"], columnspan=info.get("columnspan", 1),
               sticky="ew", pady=(6, 0))
        self.name_preview.grid_remove()
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
            self._update_name_preview()

        e.bind("<Return>", lambda _e: done(True))
        e.bind("<KP_Enter>", lambda _e: done(True))
        e.bind("<FocusOut>", lambda _e: done(True))
        e.bind("<Escape>", lambda _e: done(False))

    # ---------------------------------------------------------- скачать паспорт
    def _generate(self):
        values, images, missing = {}, {}, []
        q = self._qty()
        for key in self.field_widgets:
            kind = self.field_widgets[key][0]
            raw = self._get(key).strip()
            if kind == "image":
                paths = [x.strip() for x in raw.splitlines() if x.strip()]
                for x in paths:
                    if not Path(x).exists():
                        messagebox.showerror(APP_NAME, f"Файл изображения не найден:\n{x}")
                        return
                images[key] = paths
                continue
            if not raw and key == "НОМЕР_МАКЕТА":
                if not messagebox.askyesno(APP_NAME, "Номер макета не указан — в паспорте будет «Б/Н».\n"
                                                     "Продолжить?"):
                    return
                values[key] = "Б/Н"
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
        sides = int((self.current_item or {}).get("сторон") or 1)
        for key, path in images.items():
            if path and sides == 2 and len(path) == 1 and not messagebox.askyesno(
                    APP_NAME, "Изделие двустороннее, а выбрано изображение только одной стороны.\n"
                              "Продолжить?"):
                return
            if not path and not messagebox.askyesno(
                    APP_NAME, "Изображение макета не выбрано — в Приложении 1 будет написано "
                              "«Нет актуальной информации. Ожидается».\nПродолжить?"):
                return

        serial = values.get("СЕРИЙНЫЙ_НОМЕР", "")
        if serial:
            mine = set(serial_numbers(serial))
            own = set(serial_numbers(self.store.get(self.current_item).get("серийный_номер", ""))) \
                if (self.store and self.current_item) else set()
            clash = []
            for r in journal_rows():
                nums = set(serial_numbers(r.get("Серийный номер", "")))
                if nums & mine and not nums <= own:
                    clash.append(r)
            if clash and not messagebox.askyesno(
                    APP_NAME, f"Серийные номера {serial} пересекаются с уже выданными:\n"
                              f"{clash[-1].get('Серийный номер')} — {clash[-1].get('Создан', '')}, "
                              f"{Path(clash[-1].get('Файл', '')).name}\n\nВсё равно создать паспорт?"):
                return

        out_dir = Path(self.out_dir.get().strip() or APP_DIR)
        out_dir.mkdir(parents=True, exist_ok=True)
        base = self._out_base()
        out = out_dir / f"{base}.docx"
        if out.exists():
            ans = messagebox.askyesnocancel(
                APP_NAME, f"Файл уже есть:\n{out.name}\n\nДа — заменить\nНет — сохранить копию\n"
                          f"Отмена — не скачивать")
            if ans is None:
                return
            if ans is False:
                n = 2
                while out.exists():
                    out = out_dir / f"{base} ({n}).docx"
                    n += 1

        self.status.set("Создаю паспорт…")
        self.update_idletasks()
        try:
            lib, cfg = stamp_library()
            stamps = {"включено": self.use_stamps.get(), "файлы": lib, "настройки": cfg}
            rep = engine.generate(self._template_path(), out, values,
                                  self._table_rows() if self.tpl_table else None, images, stamps)
        except PermissionError:
            messagebox.showerror(APP_NAME, f"Файл занят (открыт в Word?):\n{out}")
            self.status.set("Ошибка")
            return
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Ошибка при создании паспорта:\n{e}")
            self.status.set("Ошибка")
            return
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
            notes.append("Вес и мощность в паспорте ориентировочные: в файле проекта нет расчёта "
                         "Конструктора. Пересохраните проект в Конструкторе V40.04.2 или новее.")
        if self.use_stamps.get():
            places = engine.stamp_places(docx.Document(self._template_path()))
            done = {k for k, *_ in rep["печати"]}
            if "ПЕЧАТЬ" in places and "ПЕЧАТЬ" not in done:
                notes.append("Нет файла печати в папке «Печати и подписи» — печать не поставлена.")
            skipped = [k.split("_", 1)[-1].capitalize() for k in places
                       if k.startswith("ПОДПИСЬ") and k not in done]
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
                n = book.refresh_toc(out, keep_pdf=self.make_pdf.get())
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
            "Серийный номер": serial,
            "Номер макета": values.get("НОМЕР_МАКЕТА", ""),
            "Тираж": q,
            "Гарантия, мес": self._get_or("ГАРАНТИЯ"),
            "Дата продажи": self._get_or("ДАТА_ПРОДАЖИ"),
            "Печать и подпись": "да" if self.use_stamps.get() and rep["печати"] else "",
            "Файл": str(out),
        })
        count_note = ""
        if self.store and self.current_item is not None:
            rec = self.store.register(self.current_item, out, serial)
            idx = self.items_list.selected
            self._refresh_list(keep=True)
            if idx is not None:
                self.items_list.select(idx, notify=False)
                self._on_item_info_refresh(idx)
            count_note = f"\nСкачан {times(rec['скачан'])}."
        self._save_settings()
        self.status.set("Паспорт создан: " + str(out))
        msg = f"Паспорт создан:\n{out.name}{pdf_note}{count_note}"
        if notes:
            messagebox.showwarning(APP_NAME, msg + "\n\nВнимание:\n• " + "\n• ".join(notes))
        else:
            messagebox.showinfo(APP_NAME, msg)
        if self.open_after.get():
            open_path(out_dir)

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
