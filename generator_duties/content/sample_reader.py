"""Отриманий зразок (Положення + функціональні обов'язки іншого батальйону)
→ витяг зі Статуту (як є) і блоки посад: назва + пункти за розділами.

Блок посади в зразку починається рядком "Загальні положення"; назва посади
стоїть перед ним - заголовок Heading 2 або центрований жирний абзац (деякі
заголовки в зразку порожні - тоді назва береться з рядка "<Посада> зобов’язаний:").
Рядок "ВОС …" пропускається: ВОС береться зі свого штату. Блоки закінчуються
підписом "Командир … військової частини …" перед листом ознайомлення.

Пункти, що трапляються в більшості блоків (стандартна відповідальність тощо),
стають спільними - рядками "Для всіх" у таблиці обов'язків."""
import re
from dataclasses import dataclass, field

from docx import Document
from docx.table import Table

from constants import SAMPLE_COMMON_SHARE, SECTION_DUTIES, SECTION_GENERAL, SECTION_RESPONSIBILITY, SECTION_RIGHTS
from content.docx_text import iter_block_items, paragraph_spec
from content.positions import compare_key

_BULLET_RE = re.compile(r"^[-–—•·*]+\s*")
_VOS_RE = re.compile(r"^ВОС\b")
# "Начальник польової лазні … ВОС 123456А" - ВОС у тому самому абзаці, що й назва.
_TRAILING_VOS_RE = re.compile(r"\s+ВОС\b.*$")
_MARKERS = (
    ("зобов’язаний", SECTION_DUTIES), ("зобов’язана", SECTION_DUTIES), ("повинен", SECTION_DUTIES),
    ("повинна", SECTION_DUTIES), ("функції", SECTION_DUTIES), ("має право", SECTION_RIGHTS),
    ("несе відповідальність", SECTION_RESPONSIBILITY), ("відповідає за", SECTION_RESPONSIBILITY),
)
_EXTRACT_START = "витяг"
_REGULATION_START = "положення"
_REGULATION_ABOUT = "про батальйон "


@dataclass
class SampleBlock:
    title: str
    sections: dict = field(default_factory=dict)

    @property
    def key(self):
        return compare_key(self.title)


@dataclass
class Sample:
    path: str
    extract: list = field(default_factory=list)
    blocks: list = field(default_factory=list)
    common: dict = field(default_factory=dict)
    # Тип батальйону зразка ("про батальйон <тип>" у його Положенні) - у
    # текстах зразка "батальйону <тип>" стає просто "батальйону": зразок
    # описує інший батальйон.
    battalion_type: str = ""


def _marker(text):
    """Рядок-заголовок розділу блоку → (розділ, назва посади з нього)."""
    lowered = text.lower()
    if lowered.startswith("загальні положення"):
        return SECTION_GENERAL, ""
    if not text.endswith(":"):
        return None, ""
    for word, section in _MARKERS:
        index = lowered.find(word)
        if index != -1:
            return section, text[:index].strip()
    return None, ""


def _is_title(paragraph, spec):
    return (paragraph.style is not None and paragraph.style.name == "Heading 2") or (spec.center and spec.bold)


def _is_signature(text):
    return text.startswith("Командир") and "військової частини" in text and len(text) < 160


def parse_sample(doc, path=""):
    sample = Sample(path)
    state = "start"
    pending_title = []
    current = None
    section = None
    for block in iter_block_items(doc):
        if isinstance(block, Table):
            if current is not None:
                break
            continue
        spec = paragraph_spec(block)
        text = spec.text
        if not text:
            continue
        lowered = text.lower()
        if state == "start":
            if lowered.startswith(_EXTRACT_START) and spec.bold:
                state = "extract"
            else:
                continue
        if state == "extract":
            if lowered == _REGULATION_START:
                state = "regulation_title"
                continue
            sample.extract.append(spec)
            continue
        if state == "regulation_title":
            state = "blocks"
            if lowered.startswith(_REGULATION_ABOUT):
                sample.battalion_type = text[len(_REGULATION_ABOUT):].strip()
                continue
        if _VOS_RE.match(text):
            continue
        if _is_signature(text) and current is not None:
            break
        marker, title_from_marker = _marker(text)
        if marker == SECTION_GENERAL:
            current = SampleBlock(" ".join(pending_title))
            sample.blocks.append(current)
            pending_title = []
            section = SECTION_GENERAL
            continue
        if _is_title(block, spec):
            pending_title.append(_TRAILING_VOS_RE.sub("", text))
            continue
        pending_title = []
        if current is None:
            continue
        if marker is not None:
            section = marker
            if not current.title and title_from_marker:
                current.title = title_from_marker
            continue
        item = _BULLET_RE.sub("", text).strip()
        if item:
            current.sections.setdefault(section, []).append(item)
    sample.blocks = [b for b in sample.blocks if b.title]
    for block in sample.blocks:
        _split_unmarked_duties(block)
        _drop_battalion_type(block, sample.battalion_type)
    sample.common = _extract_common(sample.blocks)
    return sample


def _drop_battalion_type(block, battalion_type):
    """"Командир батальйону матеріального забезпечення підпорядковується …" →
    "Командир батальйону підпорядковується …"."""
    if not battalion_type:
        return
    pattern = re.compile(rf"\b(батальйон\w*)\s+{re.escape(battalion_type)}", re.IGNORECASE)
    for section, items in block.sections.items():
        block.sections[section] = [pattern.sub(r"\1", item) for item in items]


def _split_unmarked_duties(block):
    """Блок без рядка "… зобов’язаний:" (у зразку таке трапляється): речення
    з великої літери на початку - загальні положення, решта (пункти переліку,
    з малої літери) - обов'язки."""
    general = block.sections.get(SECTION_GENERAL, [])
    if SECTION_DUTIES in block.sections or not general:
        return
    first_item = next((i for i, text in enumerate(general) if text[:1].islower()), None)
    if first_item is None:
        return
    block.sections[SECTION_GENERAL] = general[:first_item]
    block.sections[SECTION_DUTIES] = general[first_item:]
    if not block.sections[SECTION_GENERAL]:
        del block.sections[SECTION_GENERAL]


def _extract_common(blocks):
    """Пункти, що є щонайменше в SAMPLE_COMMON_SHARE блоків, → спільні (і
    прибираються з самих блоків)."""
    if not blocks:
        return {}
    common = {}
    threshold = max(2, SAMPLE_COMMON_SHARE * len(blocks))
    for section in (SECTION_GENERAL, SECTION_DUTIES, SECTION_RIGHTS, SECTION_RESPONSIBILITY):
        counts = {}
        first_text = {}
        for block in blocks:
            for item in dict.fromkeys(compare_key(i) for i in block.sections.get(section, [])):
                counts[item] = counts.get(item, 0) + 1
            for i in block.sections.get(section, []):
                first_text.setdefault(compare_key(i), i)
        shared = [key for key, count in counts.items() if count >= threshold]
        if not shared:
            continue
        common[section] = [first_text[key] for key in shared]
        for block in blocks:
            items = [i for i in block.sections.get(section, []) if compare_key(i) not in shared]
            if items:
                block.sections[section] = items
            else:
                block.sections.pop(section, None)
    return common


def read_sample(path):
    return parse_sample(Document(path), path)


def find_block(blocks, position_key):
    """Блок зразка для посади: назва блоку дорівнює назві посади або є її
    початком (цілими словами); з кількох - найдовший."""
    best = None
    for block in blocks:
        key = block.key
        if position_key == key or position_key.startswith(key + " "):
            if best is None or len(key) > len(best.key):
                best = block
    return best
