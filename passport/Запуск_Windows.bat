@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Утилита заполнения паспортов

rem --- 1. Python 3.9+ (py-лаунчер или python в PATH; заглушку Microsoft Store пропускаем)
set "PY="
py -3 -c "import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>&1 && set "PY=py -3"
if not defined PY python -c "import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>&1 && set "PY=python"
if not defined PY goto nopython

rem --- 2. tkinter (окно программы)
%PY% -c "import tkinter" >nul 2>&1 || goto notk

rem --- 3. пакеты (ставятся один раз; pypdfium2 — макеты в PDF, без него работают PNG/JPG)
%PY% -c "import docx, PIL, pypdf" >nul 2>&1
if errorlevel 1 (
  echo Первый запуск: устанавливаю пакеты python-docx, pillow, pypdf, pypdfium2, pywin32...
  %PY% -m pip install --user --disable-pip-version-check -q python-docx pillow pypdf pypdfium2 pywin32
  %PY% -c "import docx, PIL, pypdf" >nul 2>&1 || goto nopkg
)
%PY% -c "import pypdfium2" >nul 2>&1
if errorlevel 1 (
  echo Устанавливаю pypdfium2 для макетов в PDF...
  %PY% -m pip install --user --disable-pip-version-check -q pypdfium2
)

rem --- 4. запуск
%PY% -X utf8 app.py
if errorlevel 1 (
  echo.
  echo Программа завершилась с ошибкой. Подробности: ошибка_запуска.txt
  pause
)
exit /b

:nopython
echo.
echo Не найден Python 3.9 или новее.
echo Установите Python с https://www.python.org/downloads/windows/
echo При установке отметьте галочку "Add python.exe to PATH", затем запустите этот файл снова.
start "" https://www.python.org/downloads/windows/
pause
exit /b 1

:notk
echo.
echo В установленном Python нет модуля tkinter (окна).
echo Переустановите Python с python.org, компонент "tcl/tk and IDLE" должен быть включён.
pause
exit /b 1

:nopkg
echo.
echo Не удалось установить пакеты (нет интернета или доступа).
echo Выполните вручную:  %PY% -m pip install --user python-docx pillow pypdf pypdfium2 pywin32
pause
exit /b 1
