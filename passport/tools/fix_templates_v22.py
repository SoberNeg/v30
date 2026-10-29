"""V2.2 — правки вёрстки шаблонов по замечаниям (запускается на готовых .docx):
  • титул динамики — как в статике (кегли «Технический паспорт» и наименования, отступы);
  • «Технический паспорт» не заходит на «квадратики» обложки;
  • страница «Содержание» во всех паспортах — как в динамике (заголовок + новая страница);
  • «2.5. Комплектность» не разрывается и целиком уходит на следующую страницу;
  • текст не наезжает на рисунок 2.2 «Маркировка разъема Ethernet»;
  • «Спецификация экрана:» → «Спецификация изделия:».
Запуск:  python3 tools/fix_templates_v22.py Шаблоны
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import engine  # noqa: E402
import prepare_templates as pt  # noqa: E402


def main(folder):
    folder = Path(folder)
    dyn = folder / "Паспорт_Динамика_МЦД.docx"
    others = [folder / "Паспорт_Статика_МЦД.docx", folder / "Паспорт_Статика_МЦД_круглый.docx"]

    print("титул динамики по образцу статики:", pt.cover_like(dyn, others[0]))
    for f in [dyn] + others:
        print(f"{f.name}: блок содержания — абзацев {pt.toc_like(f, dyn)}")
    for f in [dyn] + others:
        print(f"{f.name}: «Технический паспорт» —", pt.fit_cover_heading(f))

    d = docx.Document(dyn)
    print("динамика: абзацев в блоке «2.5. Комплектность»",
          pt.keep_block_together(d, "2.5. Комплектность"))
    print("динамика: рисунок 2.2 — обтекание текстом,", pt.picture_wrap(d, "Оранжевый-белый"))
    n = 0
    for p in d.element.body.iter(qn("w:p")):
        n += engine.replace_in_paragraph(p, "Спецификация экрана:", "Спецификация изделия:")
    print("динамика: «Спецификация изделия» —", n)
    d.save(dyn)

    for f in [dyn] + others:
        zone = pt.fit_cover_zone(f)
        print(f"{f.name}: поле наименования {zone[0]:.1f}–{zone[1]:.1f} мм")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "Шаблоны")
