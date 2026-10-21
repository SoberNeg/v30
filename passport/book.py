# -*- coding: utf-8 -*-
"""
Подготовка паспорта к печати в типографии (брошюра).

1. Собирает PDF (MS Word → LibreOffice).
2. Проставляет в оглавлении фактические номера страниц (по тексту PDF).
3. Добивает документ пустыми страницами (с колонтитулом и номером, как в образце)
   перед задней обложкой, чтобы общее число страниц было кратно 4: обложка —
   первая страница, задняя обложка — последняя.
4. Пересобирает PDF и проверяет результат.
"""
import re
import shutil
import subprocess
import sys
from pathlib import Path

import docx
from docx.oxml.ns import qn
from lxml import etree

import engine

BACK_COVER_TAG = "Задняя обложка"
MULTIPLE = 4


# ------------------------------------------------------------------ PDF
def find_renderers():
    out = []
    if sys.platform.startswith("win"):
        try:
            import win32com.client  # noqa: F401
            out.append("word")
        except Exception:
            pass
    if sys.platform == "darwin" and Path("/Applications/Microsoft Word.app").exists():
        out.append("word-mac")
    for exe in (shutil.which("soffice"), shutil.which("libreoffice"),
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if exe and Path(exe).exists():
            out.append(exe)
            break
    return out


def render_pdf(docx_path: Path, renderer: str) -> Path:
    docx_path = Path(docx_path).resolve()
    pdf = docx_path.with_suffix(".pdf")
    if pdf.exists():
        pdf.unlink()
    if renderer == "word":
        import win32com.client
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        try:
            d = word.Documents.Open(str(docx_path), ReadOnly=True)
            d.SaveAs2(str(pdf), FileFormat=17)
            d.Close(False)
        finally:
            word.Quit()
    elif renderer == "word-mac":
        script = (f'tell application "Microsoft Word"\nopen POSIX file "{docx_path}"\n'
                  f'save as active document file name "{pdf}" file format format PDF\n'
                  f'close active document saving no\nend tell')
        subprocess.run(["osascript", "-e", script], capture_output=True, timeout=300)
    else:
        subprocess.run([renderer, "--headless", "--convert-to", "pdf", "--outdir",
                        str(docx_path.parent), str(docx_path)], capture_output=True, timeout=300)
    if not pdf.exists():
        raise RuntimeError(f"PDF не создан ({renderer}).")
    return pdf


def pdf_pages_text(pdf: Path):
    from pypdf import PdfReader
    return [(p.extract_text() or "") for p in PdfReader(str(pdf)).pages]


# ------------------------------------------------------------------ оглавление
def _norm(t):
    return re.sub(r"[^0-9a-zа-я]", "", t.lower().replace("ё", "е"))


def _toc_entries(document):
    """[(абзац, текст пункта, w:t с номером страницы)] из оглавления."""
    out = []
    for p in document.element.body.iter(qn("w:p")):
        ps = p.find(qn("w:pPr") + "/" + qn("w:pStyle"))
        in_toc = ps is not None and re.match(r"(?i)(toc|оглавление|contents)", ps.get(qn("w:val")) or "")
        anc = p.getparent()
        while anc is not None and not in_toc:
            if anc.tag == qn("w:sdt") and b"Table of Contents" in etree.tostring(anc)[:600]:
                in_toc = True
            anc = anc.getparent()
        if not in_toc:
            continue
        # текст до последней табуляции = пункт, w:t после неё = номер
        items, seen_tab = [], False
        title, num_t = "", None
        for el in p.iter(qn("w:t"), qn("w:tab")):
            if el.tag == qn("w:tab"):
                if el.getparent().tag == qn("w:tabs"):     # это позиция табуляции в pPr
                    continue
                seen_tab = True
                num_t = None
                continue
            if not seen_tab:
                title += el.text or ""
            elif (el.text or "").strip().isdigit():
                num_t = el
        if title.strip() and num_t is not None:
            out.append((p, title.strip(), num_t))
    return out


def update_toc(document, pages_text, page_offset) -> int:
    """Ставит в оглавлении фактические номера. Возвращает число изменённых пунктов."""
    entries = _toc_entries(document)
    if not entries:
        return 0
    lines = [[_norm(l) for l in t.splitlines() if l.strip()] for t in pages_text]
    # страница оглавления — где встречаются пункты с номерами
    toc_page = 0
    for i, ls in enumerate(lines):
        if sum(1 for _, title, _ in entries if any(l.startswith(_norm(title)) for l in ls)) >= \
                max(2, len(entries) // 2):
            toc_page = i
            break
    changed, start = 0, toc_page + 1
    for _, title, num_t in entries:
        e = _norm(title)
        e2 = re.sub(r"^\d+", "", e)
        found = None
        for strict in (True, False):
            for i in range(start, len(lines)):
                for l in lines[i]:
                    if (strict and l.startswith(e)) or \
                            (not strict and len(e2) >= 6 and re.sub(r"^\d+", "", l).startswith(e2)):
                        found = i
                        break
                if found is not None:
                    break
            if found is not None:
                break
        if found is None:
            continue
        start = found
        printed = str(found + 1 + page_offset)
        if num_t.text != printed:
            num_t.text = printed
            changed += 1
    return changed


def page_offset(document) -> int:
    """Номер, печатаемый на странице N PDF = N + offset (у паспортов титул = 0)."""
    sp = document.sections[0]._sectPr
    pn = sp.find(qn("w:pgNumType"))
    if pn is not None and pn.get(qn("w:start")) is not None:
        return int(pn.get(qn("w:start"))) - 1
    return 0


# ------------------------------------------------------------------ пустые страницы
def _back_cover(document):
    for dp in document.element.body.iter(qn("wp:docPr")):
        if (dp.get("descr") or "") == BACK_COVER_TAG:
            p = dp
            while p is not None and p.tag != qn("w:p"):
                p = p.getparent()
            return p
    paras = [e for e in document.element.body if e.tag == qn("w:p")]
    return next((e for e in reversed(paras) if e.find(".//" + qn("wp:anchor")) is not None), None)


def add_blank_pages(document, k) -> bool:
    back = _back_cover(document)
    if back is None:
        return False
    for _ in range(k):
        p = back.makeelement(qn("w:p"), {})
        back.addprevious(p)
        p.get_or_add_pPr().pageBreakBefore_val = True
    return True


# ------------------------------------------------------------------ всё вместе
def prepare_for_print(docx_path, renderer=None, log=print):
    """Возвращает отчёт: pdf, страниц, добавлено пустых, исправлено пунктов оглавления."""
    docx_path = Path(docx_path)
    renderers = [renderer] if renderer else find_renderers()
    if not renderers:
        raise RuntimeError("Для подготовки к печати нужен MS Word или LibreOffice — "
                           "без них нельзя посчитать страницы.")
    r = renderers[0]
    log("Собираю PDF…")
    pdf = render_pdf(docx_path, r)
    texts = pdf_pages_text(pdf)
    document = docx.Document(str(docx_path))
    off = page_offset(document)
    toc_fixed = update_toc(document, texts, off)
    pages = len(texts)
    k = (-pages) % MULTIPLE
    if k:
        if not add_blank_pages(document, k):
            raise RuntimeError("Не найдена задняя обложка — не могу добавить пустые страницы.")
    if toc_fixed or k:
        document.save(str(docx_path))
        log(f"Пересобираю PDF (оглавление: {toc_fixed}, пустых страниц: {k})…")
        pdf = render_pdf(docx_path, r)
        texts = pdf_pages_text(pdf)
        # проверка оглавления после вставки (номера не должны поменяться)
        document = docx.Document(str(docx_path))
        if update_toc(document, texts, off):
            document.save(str(docx_path))
            pdf = render_pdf(docx_path, r)
            texts = pdf_pages_text(pdf)
    total = len(texts)
    return {"pdf": pdf, "страниц": total, "пустых": k, "оглавление": toc_fixed,
            "кратно": total % MULTIPLE == 0, "сборщик": "MS Word" if r.startswith("word") else "LibreOffice"}
