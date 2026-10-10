"""Бойове розпорядження старшого командира (БР) -> номер, дата і блоки
тексту, які БН переносить дослівно (оцінка противника, вогневі завдання
старшого начальника, бойове завдання, сусіди, готовність)."""
import re
from dataclasses import dataclass, field

from docx import Document

from content.docx_blocks import body_paragraphs, item_text, merge_items, paragraph_item, paragraph_text, with_text

# Номер розділу "N." перед заголовком необов'язковий; заголовки - великими
# літерами, тож пошук чутливий до регістру (звичайний текст не спрацює).
_NUMBER_PREFIX = r"^\s*(?:\d+\s*[.)]\s*)?"
_SECTION_ANCHORS = {
    "enemy": re.compile(_NUMBER_PREFIX + r"(?:СТИСЛІ\s+)?ВИСНОВКИ\s+З\s+ОЦІН\w*\s+ПРОТИВНИКА"),
    "fire_tasks": re.compile(_NUMBER_PREFIX + r"ОБ\W?ЄКТИ,?\s+ЯКІ\s+УРАЖАЮТЬСЯ"),
    "mission": re.compile(_NUMBER_PREFIX + r"БОЙОВЕ\s+ЗАВДАННЯ"),
    "neighbours": re.compile(_NUMBER_PREFIX + r"СУСІДИ\s+ТА\s+РОЗМЕЖУВАЛЬНІ\s+ЛІНІЇ"),
    "readiness": re.compile(_NUMBER_PREFIX + r"ЧАС\s+ГОТОВНОСТІ"),
}
_SECTION_TITLES = {
    "enemy": "ВИСНОВКИ З ОЦІНЮВАННЯ ПРОТИВНИКА",
    "fire_tasks": "ОБ’ЄКТИ, ЯКІ УРАЖАЮТЬСЯ…",
    "mission": "БОЙОВЕ ЗАВДАННЯ",
    "neighbours": "СУСІДИ ТА РОЗМЕЖУВАЛЬНІ ЛІНІЇ",
    "readiness": "ЧАС ГОТОВНОСТІ",
}
_MAX_HEADING_LENGTH = 150
_MISSION_STOP_RE = re.compile(r"^З\s+початком\s+наступальних\s+дій\s+противника", re.I)
_SIGNATURE_RE = re.compile(r"^(Командир|Начальник\s+штабу|Тимчасово\s+виконуючий)\b")
_NUMBER_RE = re.compile(r"РОЗПОРЯДЖЕННЯ.*?№\s*(\d+)")
_DATE_TIME_RE = re.compile(r"(\d{1,2}[.:]\d{2})\s+(\d{2}\.\d{2}\.\d{4})")
_NEIGHBOUR_RE = re.compile(r"^(праворуч|ліворуч)\b\s*:?\s*", re.I)
_HEADER_SCAN_LIMIT = 25

# Етапи бою в розділі "БОЙОВЕ ЗАВДАННЯ" після "З початком наступальних дій
# противника": кожен етап закінчується рядком "Запасні маршрути відходу…",
# усередині - розмежувальні лінії, порядок відходу, маршрути.
_STAGE_END_RE = re.compile(r"^Запасні\s+маршрути\s+відходу", re.I)
_BOUNDARIES_HEADER_RE = re.compile(r"^Розмежувальні\s+лінії\s*:?\s*$", re.I)
_WITHDRAWAL_HEADER_RE = re.compile(r"^Порядок\s+відходу\b.*:\s*$", re.I)
_ROUTES_HEADER_RE = re.compile(r"^Маршрути?\s+відходу\b.*:\s*$", re.I)
_KSP_LINE_RE = re.compile(r"^(?:Запасне\s+)?КСП\b")


@dataclass
class BrOrder:
    path: str
    number: str = ""
    date: str = ""
    time: str = ""
    blocks: dict = field(default_factory=dict)
    warnings: list = field(default_factory=list)


def _parse_header(texts, order):
    number_index = None
    for i, text in enumerate(texts[:_HEADER_SCAN_LIMIT]):
        m = _NUMBER_RE.search(text)
        if m:
            order.number = m.group(1)
            number_index = i
            break
    start = number_index or 0
    for text in texts[start:_HEADER_SCAN_LIMIT]:
        m = _DATE_TIME_RE.search(text)
        if m:
            order.time = m.group(1).replace(":", ".")
            order.date = m.group(2)
            break
    if not order.number:
        order.warnings.append("У шапці БР не знайдено номер (\"РОЗПОРЯДЖЕННЯ … № N\").")
    if not order.date:
        order.warnings.append("У шапці БР не знайдено час і дату (\"ГГ.ХХ ДД.ММ.РРРР\").")


def read_br_header(path):
    """Лише номер/дата/час - для вибору найновішого БР без розбору всього тексту."""
    order = BrOrder(path=path)
    texts = [paragraph_text(p) for p in body_paragraphs(Document(path))[:_HEADER_SCAN_LIMIT]]
    _parse_header(texts, order)
    return order


def _find_anchors(texts):
    anchors = {}
    for i, text in enumerate(texts):
        if len(text) > _MAX_HEADING_LENGTH:
            continue
        for name, pattern in _SECTION_ANCHORS.items():
            if name not in anchors and pattern.match(text):
                anchors[name] = i
                break
    return anchors


def _signature_index(anchors, texts):
    """Підпис шукається лише після ОСТАННЬОГО розділу - всередині тексту БР
    абзац теж може починатися словом "Командир"."""
    last = max(anchors.values(), default=-1)
    return next((i for i in range(last + 1, len(texts)) if _SIGNATURE_RE.match(texts[i])), len(texts))


def _section_end(start, anchors, signature_index):
    return min([i for i in anchors.values() if i > start] + [signature_index])


def _merge_neighbours(items):
    """"праворуч: …"/"ліворуч: …" -> "Сусід праворуч: …". Абзац, що закінчується
    на ":", зливається з наступним абзацом-продовженням (у БР опис сусіда буває
    розбитий на два абзаци)."""
    merged = []
    for item in items:
        text = item_text(item)
        if merged and not _NEIGHBOUR_RE.match(text) and item_text(merged[-1]).rstrip().endswith(":"):
            merged[-1] = merge_items(merged[-1], item)
        else:
            merged.append(item)
    result = []
    for item in merged:
        text = item_text(item)
        m = _NEIGHBOUR_RE.match(text)
        if m:
            text = f"Сусід {m.group(1).lower()}: {text[m.end():]}"
        result.append(with_text(item, text))
    return result


def read_br(path):
    order = BrOrder(path=path)
    paragraphs = body_paragraphs(Document(path))
    texts = [paragraph_text(p) for p in paragraphs]
    _parse_header(texts, order)

    anchors = _find_anchors(texts)
    for name, title in _SECTION_TITLES.items():
        if name not in anchors:
            order.warnings.append(f"У БР не знайдено розділ «{title}».")

    signature_index = _signature_index(anchors, texts)

    def items_between(start, end):
        return [item for item in (paragraph_item(p) for p in paragraphs[start + 1:end]) if item_text(item)]

    for name in ("enemy", "fire_tasks", "neighbours", "readiness"):
        if name in anchors:
            start = anchors[name]
            order.blocks[f"br_{name}"] = items_between(start, _section_end(start, anchors, signature_index))

    if "br_neighbours" in order.blocks:
        order.blocks["br_neighbours"] = _merge_neighbours(order.blocks["br_neighbours"])

    if "mission" in anchors:
        start = anchors["mission"]
        mission = items_between(start, _section_end(start, anchors, signature_index))
        stop = next((i for i, item in enumerate(mission) if _MISSION_STOP_RE.match(item_text(item))), None)
        if stop is None:
            order.warnings.append("У розділі «БОЙОВЕ ЗАВДАННЯ» БР не знайдено «З початком наступальних дій противника» - перенесено розділ повністю.")
            order.blocks["br_mission"] = mission
        else:
            order.blocks["br_mission"] = mission[:stop]
            _add_stage_blocks(order.blocks, mission[stop:])
    return order


def _split_stages(items):
    stages, current = [], []
    for item in items:
        current.append(item)
        if _STAGE_END_RE.match(item_text(item)):
            stages.append(current)
            current = []
    if current:
        stages.append(current)
    return stages


def _stage_parts(stage):
    parts = {"all": list(stage), "actions": [], "right": [], "left": [], "withdrawal": [], "routes": []}
    mode = "actions"
    for item in stage:
        text = item_text(item)
        if _BOUNDARIES_HEADER_RE.match(text):
            mode = "boundaries"
            continue
        if _WITHDRAWAL_HEADER_RE.match(text):
            mode = "withdrawal"
            continue
        if _ROUTES_HEADER_RE.match(text):
            mode = "routes"
            continue
        if mode == "boundaries":
            side = _NEIGHBOUR_RE.match(text)
            if side:
                parts["right" if side.group(1).lower() == "праворуч" else "left"].append(item)
                continue
            mode = "actions"
        parts[mode].append(item)
    return parts


def _add_stage_blocks(blocks, items):
    """"br_stage:N:all|actions|right|left|withdrawal|routes" для кожного етапу і
    "br_ksp" - рядки КСП/запасного КСП наприкінці розділу."""
    ksp = []
    while items and _KSP_LINE_RE.match(item_text(items[-1])):
        ksp.insert(0, items[-1])
        items = items[:-1]
    if ksp:
        blocks["br_ksp"] = ksp
    for number, stage in enumerate(_split_stages(items), start=1):
        for part, part_items in _stage_parts(stage).items():
            if part_items:
                blocks[f"br_stage:{number}:{part}"] = part_items
