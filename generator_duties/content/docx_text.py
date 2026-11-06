"""Спільне читання .docx (зразок, Положення): блоки тіла по порядку і абзац
як ParagraphSpec - текст + жирність + вирівнювання, без чужих стилів і
нумерації (у зразку їх десятки різних, скопійованих з різних джерел)."""
import re
from dataclasses import dataclass

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

_APOSTROPHES_RE = re.compile(r"['ʼ`´‘]")
_SPACES_RE = re.compile(r"\s+")


@dataclass
class ParagraphSpec:
    text: str
    bold: bool = False
    center: bool = False
    listed: bool = False


def iter_block_items(doc):
    """Абзаци й таблиці тіла документа в порядку появи (копія з
    generator_bn_and_map/content/docx_blocks.py)."""
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def clean_text(text):
    """Апостроф - завжди ’ (як у зразку), розриви рядків і кілька пробілів -
    один пробіл."""
    text = _APOSTROPHES_RE.sub("’", (text or "").replace("\xa0", " "))
    return _SPACES_RE.sub(" ", text).strip()


def paragraph_spec(paragraph):
    text = clean_text(paragraph.text)
    text_runs = [run for run in paragraph.runs if run.text.strip()]
    bold = bool(text_runs) and all(run.bold for run in text_runs)
    center = paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
    p_pr = paragraph._p.pPr
    listed = p_pr is not None and p_pr.numPr is not None
    return ParagraphSpec(text, bold, center, listed)


def read_paragraph_specs(doc):
    """Непорожні абзаци верхнього рівня документа (таблиці пропускаються)."""
    specs = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            spec = paragraph_spec(block)
            if spec.text:
                specs.append(spec)
    return specs
