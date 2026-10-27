"""CI (Windows): проверка запуска прототипа так, как это делает пользователь.
  1. копирует папку passport/ в путь с русскими буквами;
  2. режим garbled: портит русские имена так, как их распаковывает Проводник
     из архива без флага UTF-8 (cp866), — программа должна сама их исправить;
  3. запускает Запуск_Windows.bat, ждёт окно «Утилита заполнения паспортов … V2.0»,
     снимает экран, закрывает. Код возврата 0 — окно появилось, ошибок запуска нет."""
import ctypes, os, shutil, subprocess, sys, time
from ctypes import wintypes
from pathlib import Path

mode, out = sys.argv[1], Path(sys.argv[2])
src = Path(__file__).resolve().parents[1]
dst = Path(os.environ.get("RUNNER_TEMP", "C:/t")) / f"Тест {mode}" / "Паспорта_прототип_V2.0"
if dst.parent.exists():
    shutil.rmtree(dst.parent)
shutil.copytree(src, dst, ignore=shutil.ignore_patterns("ci", "__pycache__"))
if mode == "garbled":
    for dp, dn, fn in os.walk(dst, topdown=False):
        for n in fn + dn:
            g = n.encode("utf-8").decode("cp866", errors="replace")
            if g != n:
                os.rename(os.path.join(dp, n), os.path.join(dp, g))
bat = next(p for p in dst.iterdir() if p.suffix.lower() == ".bat")
print("Папка:", dst, "\nЗапуск:", bat.name, "\nДо запуска:", sorted(p.name for p in dst.iterdir()))
proc = subprocess.Popen(["cmd.exe", "/c", str(bat)], cwd=str(dst), creationflags=subprocess.CREATE_NEW_CONSOLE)

user32 = ctypes.windll.user32
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)


def titles():
    found = []
    def cb(h, _):
        if user32.IsWindowVisible(h):
            n = user32.GetWindowTextLengthW(h)
            if n:
                b = ctypes.create_unicode_buffer(n + 1)
                user32.GetWindowTextW(h, b, n + 1)
                found.append(b.value)
        return True
    user32.EnumWindows(EnumWindowsProc(cb), 0)
    return found


ok, seen = False, []
for _ in range(240):                     # до 4 минут: первый запуск ставит пакеты
    seen = titles()
    if any(t.startswith("Утилита заполнения паспортов") and "V2.0" in t for t in seen):
        ok = True
        break
    if (dst / "ошибка_запуска.txt").exists():
        break
    time.sleep(1)
time.sleep(4)
out.mkdir(parents=True, exist_ok=True)
try:
    from PIL import ImageGrab
    ImageGrab.grab().save(out / f"screen_{mode}.png")
except Exception as e:
    print("снимок экрана не сделан:", e)
err = dst / "ошибка_запуска.txt"
print("Окна:", [t for t in seen if t][:15])
print("После запуска:", sorted(p.name for p in dst.iterdir()))
if err.exists():
    print("ОШИБКА ЗАПУСКА:\n", err.read_text("utf-8"))
subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True)
print("РЕЗУЛЬТАТ:", "OK — окно программы открылось" if ok and not err.exists() else "FAIL")
sys.exit(0 if ok and not err.exists() else 1)
