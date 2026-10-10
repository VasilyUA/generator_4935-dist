"""Спільне читання .docx для БР і ПБД: абзаци тіла документа по порядку,
текст абзацу і жирні фрази в ньому - у тому самому форматі item, що й
тексти БН у data.json (рядок або [текст, [жирні фрази]])."""
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

_SKIPPED_RUN_ANCESTORS = {qn("w:del"), qn("w:moveFrom")}


def iter_block_items(doc):
    """Абзаци й таблиці тіла документа в порядку появи (копія
    _iter_block_items з generator_br_and_report_for_money/content/report_document_reader.py)."""
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def body_paragraphs(doc):
    """Лише абзаци верхнього рівня (таблиці пропускаються - у БР це порожні
    рамки 1x1, у ПБД розділ 2 таблиць не має)."""
    return [block for block in iter_block_items(doc) if isinstance(block, Paragraph)]


def _iter_runs(paragraph):
    """Усі runs абзацу, включно з тими, що всередині гіперпосилань/вставок, але
    без видаленого тексту відстежуваних змін."""
    p = paragraph._p
    for r in p.iter(qn("w:r")):
        parent = r.getparent()
        skipped = False
        while parent is not None and parent is not p:
            if parent.tag in _SKIPPED_RUN_ANCESTORS:
                skipped = True
                break
            parent = parent.getparent()
        if not skipped:
            yield Run(r, paragraph)


def _clean(text):
    # Розрив рядка (w:br) у джерелах - лише верстка, у БН це звичайний пробіл.
    return text.replace("\r", " ").replace("\n", " ")


def paragraph_text(paragraph):
    return "".join(_clean(run.text) for run in _iter_runs(paragraph)).strip()


def paragraph_item(paragraph):
    """Абзац -> item: "текст" або ["текст", ["жирна фраза", ...]]. Сусідні
    жирні runs зливаються в одну фразу."""
    segments = []
    for run in _iter_runs(paragraph):
        text = _clean(run.text)
        if not text:
            continue
        bold = bool(run.bold)
        if segments and segments[-1][1] == bold:
            segments[-1][0] += text
        else:
            segments.append([text, bold])
    full = "".join(text for text, _ in segments).strip()
    if not full:
        return ""
    phrases = []
    for text, bold in segments:
        phrase = text.strip()
        if bold and phrase and phrase not in phrases:
            phrases.append(phrase)
    return [full, phrases] if phrases else full


def item_text(item):
    """Текст item будь-якої форми (рядок, [текст, фрази], {"text": ...})."""
    if isinstance(item, str):
        return item
    if isinstance(item, list) and item:
        return str(item[0])
    if isinstance(item, dict):
        return str(item.get("text", ""))
    return ""


def with_text(item, text):
    """Той самий item з іншим текстом (жирні фрази, яких більше немає в тексті,
    відкидаються)."""
    if isinstance(item, list) and len(item) > 1:
        phrases = [phrase for phrase in item[1] if phrase and phrase in text]
        return [text, phrases] if phrases else text
    if isinstance(item, dict):
        return {**item, "text": text}
    return text


def merge_items(first, second, separator=" "):
    """Два абзаци -> один item (жирні фрази обох)."""
    text = f"{item_text(first)}{separator}{item_text(second)}".strip()
    phrases = []
    for item in (first, second):
        if isinstance(item, list) and len(item) > 1:
            phrases.extend(phrase for phrase in item[1] if phrase not in phrases)
    return [text, phrases] if phrases else text
