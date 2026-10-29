"""Проверка порядка элементов в .docx — Word открывает файл, только если дочерние
элементы идут в порядке схемы OOXML (LibreOffice это прощает).
Запуск:  python3 tools/check_docx.py файл.docx [ещё файлы]
"""
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"

ORDER = {
    W + "pPr": """pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr
        suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct
        topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind
        contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment
        textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange""".split(),
    W + "rPr": """ins del moveFrom moveTo rStyle rFonts b bCs i iCs caps smallCaps strike dstrike
        outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern
        position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang
        eastAsianLayout specVanish oMath rPrChange""".split(),
    WP + "anchor": """simplePos positionH positionV extent effectExtent wrapNone wrapSquare
        wrapTight wrapThrough wrapTopAndBottom docPr cNvGraphicFramePr graphic sizeRelH
        sizeRelV""".split(),
}
# у обтекания один «слот»: любой wrap* стоит на месте wrapNone
SLOT = {WP + "anchor": {"wrapNone": 5, "wrapSquare": 5, "wrapTight": 5, "wrapThrough": 5,
                        "wrapTopAndBottom": 5}}


def check_xml(data, where):
    bad = []
    root = etree.fromstring(data)
    for parent_tag, order in ORDER.items():
        pos = {n: i for i, n in enumerate(order)}
        pos.update({k: v for k, v in SLOT.get(parent_tag, {}).items()})
        for el in root.iter(parent_tag):
            seen = -1
            for ch in el:
                name = ch.tag.split("}")[-1]
                if name not in pos:                 # чужие пространства имён пропускаем
                    continue
                i = pos[name]
                if i < seen:
                    kids = [c.tag.split("}")[-1] for c in el]
                    bad.append(f"{where}: <{parent_tag.split('}')[-1]}> — <{name}> не на месте: {kids}")
                    break
                seen = i
    return bad


def check(path):
    bad = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name.endswith(".xml") and (name.startswith("word/") or name == "word/document.xml"):
                bad += check_xml(z.read(name), f"{Path(path).name}/{name}")
    return bad


if __name__ == "__main__":
    problems = []
    for f in sys.argv[1:]:
        problems += check(f)
    if problems:
        print("\n".join(problems))
        print("ОШИБОК ПОРЯДКА:", len(problems))
        sys.exit(1)
    print("порядок элементов в норме:", ", ".join(Path(f).name for f in sys.argv[1:]))
