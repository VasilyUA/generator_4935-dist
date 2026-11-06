"""Позатермінове донесення з переліком перейменованих позицій:
`Вогнева позиція «СТАРА» перейменовано на «НОВА»` -> [Rename(ВП, СТАРА, НОВА)].

Перейменування застосовуються до згадок позицій у статичному тексті БН
(content/position_refresh.py), тож старі назви з попереднього наказу
замінюються новими без пошуку за координатами."""
import re
from dataclasses import dataclass, field

from content.docx_blocks import item_text, read_paragraph_items

_OPEN = "«\"“„"
_CLOSE = "»\"”“"
_RENAME_RE = re.compile(
    rf"(?P<type>[^{_OPEN}\n]*?)\s*[{_OPEN}]\s*(?P<old>[^{_CLOSE}\n]+?)\s*[{_CLOSE}]\s*"
    rf"перейменован\w*\s+на\s*[{_OPEN}]\s*(?P<new>[^{_CLOSE}\n]+?)\s*[{_CLOSE}]",
    re.IGNORECASE,
)
_HEADER_RE = re.compile(r"№\s*(?P<number>\d+)")
_DATE_RE = re.compile(r"(?<!\d)(\d{2}\.\d{2}\.\d{4})(?!\d)")

# Повна назва типу позиції -> скорочення, як у ПБД і тексті БН.
_TYPE_WORDS = (
    ("вогнева засідка", "ВЗ"),
    ("вогнева позиція", "ВП"),
    ("позиція відділення", "ПВ"),
    ("спостережна позиція", "СП"),
    ("спостережний пост", "СП"),
)
_TYPE_ABBREVIATIONS = ("СПАР", "КСП", "ЗКП", "ВОП", "РОП", "ХАБ", "ВП", "ПВ", "СП", "ВЗ", "ТЗ", "ПУ")


@dataclass(frozen=True)
class Rename:
    type: str
    old: str
    new: str


@dataclass
class RenameReport:
    path: str
    number: str = ""
    date: str = ""
    renames: list = field(default_factory=list)

    @property
    def label(self):
        parts = [f"№{self.number}" if self.number else "", f"від {self.date}" if self.date else ""]
        return " ".join(p for p in parts if p) or self.path


def position_type(text):
    """"Вогнева позиція" / "ВОП" / "Позиція відділення" -> скорочення типу або ""."""
    lower = (text or "").strip().lower()
    for words, abbreviation in _TYPE_WORDS:
        if words in lower:
            return abbreviation
    upper = (text or "").upper()
    for abbreviation in _TYPE_ABBREVIATIONS:
        if re.search(rf"(?<![А-ЯІЇЄҐA-Z]){abbreviation}(?![А-ЯІЇЄҐA-Z])", upper):
            return abbreviation
    return ""


def parse_renames(texts):
    renames = []
    for text in texts:
        for m in _RENAME_RE.finditer(text or ""):
            old, new = m.group("old").strip(), m.group("new").strip()
            if old and new and old.upper() != new.upper():
                renames.append(Rename(position_type(m.group("type")), old.upper(), new.upper()))
    return renames


def read_renames(path):
    texts = [item_text(item) for item in read_paragraph_items(path)]
    report = RenameReport(path, renames=parse_renames(texts))
    for text in texts:
        if "ДОНЕСЕННЯ" in text.upper():
            number = _HEADER_RE.search(text)
            date = _DATE_RE.findall(text)
            report.number = number.group("number") if number else ""
            report.date = date[-1] if date else ""
            break
    return report
