"""Статут внутрішньої служби ЗСУ (збережена сторінка zakon.rada.gov.ua/laws/show/548-14)
→ підрозділи: заголовок (абзац class="rvps7") і його статті (class="rvps2").
Примітки про зміни ("{Стаття … із змінами …}") і виноски (rvps8) пропускаються.

Кожен підрозділ розкладається на розділи таблиці обов'язків: абзаци до
рядка "… зобов’язаний:" - загальні положення (відповідає за…, підпорядковується…),
пункти після нього - обов'язки, після "… має право:" - права. Додаткові
переліки "Крім того, командир окремої роти зобов’язаний:" не беруться."""
import re
from dataclasses import dataclass, field

from bs4 import BeautifulSoup

from constants import SECTION_DUTIES, SECTION_GENERAL, SECTION_RIGHTS
from content.docx_text import clean_text

_ARTICLE_RE = re.compile(r"^(\d+(?:\s*-\s*\d+)?)\s*\.\s*")
_PARENTHESES_RE = re.compile(r"\([^)]*\)")


@dataclass
class StatuteSection:
    heading: str
    key: str
    articles: list = field(default_factory=list)
    sections: dict = field(default_factory=dict)

    @property
    def source(self):
        return f"Статут ВС, ст. {', '.join(self.articles)}" if self.articles else "Статут ВС"

    @property
    def title(self):
        """Заголовок без уточнень у дужках: "Командир батальйону"."""
        return clean_text(_PARENTHESES_RE.sub(" ", self.heading)).replace(" ,", ",").strip(" ,")

    def items(self, section):
        return self.sections.get(section, [])


def heading_key(text):
    """"Командир батальйону (корабля 3 рангу)*" → "командир батальйону"."""
    text = _PARENTHESES_RE.sub(" ", clean_text(text).replace("*", ""))
    return clean_text(text).replace(" ,", ",").strip(" ,").lower()


def _paragraph_class(p):
    return " ".join(p.get("class", []))


def parse_statute_html(html):
    """→ [StatuteSection] у порядку документа."""
    soup = BeautifulSoup(html, "html.parser")
    parsed = []  # [(StatuteSection, [(текст, чи це початок статті)])]
    for p in soup.find_all("p"):
        cls = _paragraph_class(p)
        text = clean_text(p.get_text(" ", strip=True))
        if not text:
            continue
        if "rvps7" in cls:
            parsed.append((StatuteSection(text.replace("*", "").strip(), heading_key(text)), []))
            continue
        if not parsed or "rvps8" in cls or text.startswith("{"):
            continue
        section, paragraphs = parsed[-1]
        match = _ARTICLE_RE.match(text)
        if match:
            section.articles.append(match.group(1).replace(" ", ""))
            text = text[match.end():].strip()
        paragraphs.append((text.replace("*", "").strip(), bool(match)))
    for section, paragraphs in parsed:
        section.sections = _split_sections(paragraphs)
    return [section for section, _ in parsed]


def _split_sections(paragraphs):
    result = {SECTION_GENERAL: [], SECTION_DUTIES: [], SECTION_RIGHTS: []}
    state = SECTION_GENERAL
    seen_duties = False
    for text, is_article in paragraphs:
        lowered = text.lower().replace("'", "’")
        if text.endswith(":") and ("зобов’язаний" in lowered or "зобов’язана" in lowered):
            if seen_duties and lowered.startswith("крім того"):
                state = None
            else:
                state = SECTION_DUTIES
                seen_duties = True
            continue
        if text.endswith(":") and "має право" in lowered:
            state = SECTION_RIGHTS
            continue
        if is_article and state != SECTION_GENERAL:
            # Окрема стаття після переліку (напр. обмеження) - вже не пункт переліку.
            state = SECTION_GENERAL
        if state is not None:
            result[state].append(text)
    return {name: items for name, items in result.items() if items}


def read_statute(path):
    with open(path, "rb") as f:
        raw = f.read()
    for encoding in ("utf-8", "cp1251"):
        try:
            return parse_statute_html(raw.decode(encoding))
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Не вдалось прочитати {path}: невідоме кодування.")


def find_section(sections, heading_start):
    """Підрозділ, заголовок якого (без дужок) збігається з heading_start або
    починається з нього; з кількох - найкоротший (найзагальніший)."""
    key = heading_key(heading_start)
    exact = [s for s in sections if s.key == key]
    if exact:
        return exact[0]
    starting = [s for s in sections if s.key.startswith(key)]
    return min(starting, key=lambda s: len(s.key)) if starting else None
