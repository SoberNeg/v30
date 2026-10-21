# -*- coding: utf-8 -*-
"""Шрифты паспорта: RussianRail G Pro (титул) и GOST 2.304 type A (весь остальной текст).
Файлы лежат в папке «Шрифты» рядом с программой. Проверка, установлены ли они в системе
(Word/LibreOffice берут шрифты только из системы), и установка для текущего пользователя."""
import os
import shutil
import subprocess
import sys
from pathlib import Path

REQUIRED = {"RussianRail G Pro": "титульный лист", "GOST 2.304 type A": "основной текст и колонтитулы"}
EXT = (".ttf", ".otf", ".ttc")


def _family(path):
    try:
        from PIL import ImageFont
        return ImageFont.truetype(str(path), 12).getname()[0]
    except Exception:
        return None


def system_font_dirs():
    home = Path.home()
    if sys.platform.startswith("win"):
        return [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
                Path(os.environ.get("LOCALAPPDATA", home / "AppData/Local")) / "Microsoft/Windows/Fonts"]
    if sys.platform == "darwin":
        return [home / "Library/Fonts", Path("/Library/Fonts"), Path("/System/Library/Fonts")]
    return [home / ".fonts", home / ".local/share/fonts", Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts")]


def _candidates(d):
    keys = ("rail", "gost", "гост")
    try:
        for f in d.rglob("*"):
            if f.suffix.lower() in EXT and any(k in f.name.lower() for k in keys):
                yield f
    except Exception:
        return


def installed():
    found = {}
    for d in system_font_dirs():
        if d.is_dir():
            for f in _candidates(d):
                fam = _family(f)
                if fam in REQUIRED:
                    found[fam] = f
    return found


def bundled(fonts_dir: Path):
    out = {}
    if fonts_dir.is_dir():
        for f in fonts_dir.iterdir():
            if f.suffix.lower() in EXT:
                fam = _family(f)
                if fam in REQUIRED:
                    out[fam] = f
    return out


def missing():
    have = installed()
    return [f for f in REQUIRED if f not in have]


def install(files) -> list:
    """Устанавливает шрифты для текущего пользователя. Возвращает список ошибок."""
    errors = []
    for fam, src in files.items():
        try:
            if sys.platform.startswith("win"):
                import winreg
                dst_dir = Path(os.environ["LOCALAPPDATA"]) / "Microsoft/Windows/Fonts"
                dst_dir.mkdir(parents=True, exist_ok=True)
                dst = dst_dir / Path(src).name
                shutil.copy2(src, dst)
                kind = "OpenType" if Path(src).suffix.lower() == ".otf" else "TrueType"
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                    r"Software\Microsoft\Windows NT\CurrentVersion\Fonts", 0,
                                    winreg.KEY_SET_VALUE) as k:
                    winreg.SetValueEx(k, f"{fam} ({kind})", 0, winreg.REG_SZ, str(dst))
            elif sys.platform == "darwin":
                dst_dir = Path.home() / "Library/Fonts"
                dst_dir.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst_dir / Path(src).name)
            else:
                dst_dir = Path.home() / ".local/share/fonts"
                dst_dir.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst_dir / Path(src).name)
                subprocess.run(["fc-cache", "-f"], capture_output=True)
        except Exception as e:
            errors.append(f"{fam}: {e}")
    return errors
