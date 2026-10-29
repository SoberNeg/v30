"""CI (Windows): открывает готовые паспорта в самом Word.
Word не открывает .docx с нарушенным порядком элементов XML, а LibreOffice открывает,
поэтому проверка на Linux такую поломку не ловит — ловит эта.
Аргумент: папка с .docx. Код возврата 0 — все файлы открылись."""
import glob
import os
import sys

import win32com.client

folder = sys.argv[1]
files = sorted(glob.glob(os.path.join(folder, "*.docx")))
if not files:
    print("нет файлов для проверки:", folder)
    sys.exit(1)

try:
    word = win32com.client.DispatchEx("Word.Application")
except Exception as e:                          # noqa: BLE001
    print("Word на этой машине не установлен — проверка пропущена:", e)
    print("РЕЗУЛЬТАТ: ПРОПУЩЕНО — нет Word")
    sys.exit(0)
word.Visible = False
word.DisplayAlerts = 0
bad = []
for f in files:
    path = os.path.abspath(f)
    try:
        doc = word.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False,
                                  ConfirmConversions=False)
        pages = doc.ComputeStatistics(2)        # wdStatisticPages
        words = doc.ComputeStatistics(0)        # wdStatisticWords
        doc.Close(False)
        print(f"OK  {os.path.basename(f)} — Word открыл, страниц {pages}, слов {words}")
    except Exception as e:                      # noqa: BLE001
        bad.append(os.path.basename(f))
        print(f"ERR {os.path.basename(f)} — Word не открыл файл: {e}")
try:
    word.Quit()
except Exception:                               # noqa: BLE001
    pass
print("РЕЗУЛЬТАТ:", "OK — все паспорта открываются в Word" if not bad
      else "FAIL — Word не открыл: " + ", ".join(bad))
sys.exit(1 if bad else 0)
