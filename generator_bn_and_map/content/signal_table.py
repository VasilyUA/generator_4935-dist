"""Таблиця сигналів бойового управління батальйону (.docx): рядки
"№ | Словникова група | Цифрова група | Зміст сигналу", поділені рядками-
заголовками розділів ("1. ЗАГАЛЬНІ СИГНАЛИ"), груп ("3.2. Для БнГ") і рубежів
("На першому оборонному рубежі").

Сигнали з таблиці вважаються підтвердженими джерелом, коли трапляються в
тексті БН; сигнали БН, яких у таблиці немає (чи які в ній інші), дописуються
у виправлену копію таблиці - generators/generate_signals.py."""
import re
from dataclasses import dataclass, field

from docx import Document
from docx.oxml.ns import qn

_SECTION_RE = re.compile(r"^\d+\.\s*\D")
_GROUP_RE = re.compile(r"^\d+\.\d+\.?\s*\S")
_LINE_RE = re.compile(r"^на\s+\S+.*рубеж", re.IGNORECASE)


def signal_key(text):
    """"НАЗВА-123", "НАЗВА 123", ‘‘НАЗВА – 123’’ -> "НАЗВА123"."""
    return re.sub(r"[\W_]", "", text or "").upper()


def split_signal(name):
    """"НАЗВА-123" -> ("НАЗВА", "123")."""
    m = re.match(r"^\s*(?P<word>.*?\D)\s*[-–]?\s*(?P<number>\d+)\s*$", name or "")
    if not m:
        return (name or "").strip(), ""
    return m.group("word").strip(" -–"), m.group("number")


@dataclass
class SignalRow:
    word: str
    number: str
    meaning: str
    section: str = ""
    group: str = ""
    line: str = ""
    index: int = 0

    @property
    def key(self):
        return signal_key(f"{self.word}{self.number}")

    @property
    def name(self):
        return f"{self.word}-{self.number}"


@dataclass
class SignalTable:
    path: str
    table_index: int = 0
    rows: list = field(default_factory=list)
    warnings: list = field(default_factory=list)

    @property
    def keys(self):
        return {row.key for row in self.rows}

    def find(self, key):
        return next((row for row in self.rows if row.key == key), None)


def cell_texts(tr):
    """Тексти комірок рядка таблиці (об'єднана комірка - один текст)."""
    texts = []
    for tc in tr.iter(qn("w:tc")):
        parts = []
        for paragraph in tc.iter(qn("w:p")):
            parts.append("".join(t.text or "" for t in paragraph.iter(qn("w:t"))))
        texts.append(" ".join(p.strip() for p in parts if p.strip()))
    return texts


def _is_signal_table(table):
    header = " ".join(" ".join(cell_texts(tr)) for tr in table._tbl.tr_lst[:3]).lower()
    return "словников" in header or "цифров" in header


def parse_rows(table):
    """python-docx Table -> [SignalRow] (рядки-заголовки задають розділ/групу)."""
    rows, section, group, line = [], "", "", ""
    for index, tr in enumerate(table._tbl.tr_lst):
        cells = cell_texts(tr)
        filled = [c for c in cells if c]
        if len(cells) >= 4 and re.fullmatch(r"\d+", cells[2] or "") and cells[1]:
            rows.append(SignalRow(cells[1].strip(), cells[2].strip(), cells[3].strip(), section, group, line, index))
        elif len(filled) == 1:
            label = filled[0]
            if _GROUP_RE.match(label):
                group, line = label, ""
            elif _SECTION_RE.match(label):
                section, group, line = label, "", ""
            elif _LINE_RE.match(label):
                line = label
    return rows


def read_signal_table(path):
    doc = Document(path)
    result = SignalTable(path)
    for index, table in enumerate(doc.tables):
        if _is_signal_table(table):
            result.table_index = index
            result.rows = parse_rows(table)
            break
    else:
        result.warnings.append("У таблиці сигналів не знайдено таблиці з колонками «Словникова група / Цифрова група».")
    return result
