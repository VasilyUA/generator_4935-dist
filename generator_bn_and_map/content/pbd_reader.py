"""Підсумкове бойове донесення (ПБД) -> шапка (КСП, дата й час, видання
карти), розділ 2 "ПОЛОЖЕННЯ ТА СТАН ПІДРОЗДІЛІВ" як блоки тексту для БН,
позиції (точки карти) і рубежі/райони (лінії карти).

Розділ 2 ПБД - один абзац на рядок (так його генерує final_combat_report),
підблоки відокремлені рядками-якорями на кшталт "Положення підрозділів:".
Назви позицій беруться з лапок, координати - MGRS у дужках."""
import re
from dataclasses import dataclass, field

from docx import Document

from content.docx_blocks import body_paragraphs, item_text, paragraph_item
from utils.mgrs import find_mgrs

_SECTION_START_RE = re.compile(r"^\s*(?:2\s*\.\s*)?ПОЛОЖЕННЯ\s+ТА\s+СТАН\s+ПІДРОЗДІЛІВ")
_SECTION_END_RE = re.compile(r"^\s*(?:3\s*\.\s*)?ХІД\s+ВЕДЕННЯ\s+БОЙОВИХ\s+ДІЙ")
_HEADER_RE = re.compile(
    r"КСП\s*[–—-]\s*(?P<place>[^\d]+?)\s+(?P<time>\d{1,2}[.:]\d{2})\s+(?P<date>\d{2}\.\d{2}\.\d{4})"
)
_EDITION_RE = re.compile(r"видання\s+(\d{4})")

_ANCHORS = (
    ("boundaries", re.compile(r"^Розмежувальні\s+лінії\s*:?$", re.I)),
    ("battle_order", re.compile(r"^Бойовий\s+порядок\s*:?$", re.I)),
    ("positions", re.compile(r"^Положення\s+підрозділів\s*:?$", re.I)),
    ("reserves", re.compile(r"^Резерви\s*:?$", re.I)),
    ("observation", re.compile(r"^Спостережні\s+позиції\b.*:$", re.I)),
)
# Великі підгрупи всередині "Положення підрозділів" - лишаються частиною
# блоку positions, але завершують перелік позицій попереднього РОП.
_MAJOR_GROUP_RE = re.compile(
    r"^(Окремі\s+позиції|Основні\s+вогневі\s+позиції|Основні\s+пункти\s+управління|Логістичні\s+хаби)\b.*:$", re.I
)
_UAV_CONTROL_RE = re.compile(r"^Основні\s+пункти\s+управління\b", re.I)
_GROUND_ROBOTICS_RE = re.compile(r"^НРК\b", re.I)
_MAX_GROUP_HEADER_LENGTH = 90

_QUOTE_OPEN = "«\"“„”'‘’"
_QUOTE_CLOSE = "»\"”“'’‘"
_NAME_RE = re.compile(rf"[{_QUOTE_OPEN}]{{1,2}}\s*(?P<name>[^{_QUOTE_OPEN}{_QUOTE_CLOSE}]+?)\s*[{_QUOTE_CLOSE}]{{1,2}}")
POSITION_TYPES = ("КСП", "ВОП", "РОП", "ВП", "ПВ", "СП", "ВЗ", "ТЗ", "ХАБ", "ПУ", "СПАР")
_POSITION_RE = re.compile(r"^[-–—•\s]*(?P<type>" + "|".join(sorted(POSITION_TYPES, key=len, reverse=True)) + r")(?![\wА-яІіЇїЄєҐґ'’])")
_COMPANY_RE = re.compile(rf"(?:оборону|утримує|утримання)\s+РОП\s*[{_QUOTE_OPEN}]{{1,2}}\s*(?P<rop>[^{_QUOTE_OPEN}{_QUOTE_CLOSE}]+?)\s*[{_QUOTE_CLOSE}]")
_FRONT_EDGE_RE = re.compile(r"з\s+переднім\s+краєм", re.I)
_SIDE_RE = re.compile(r"^(праворуч|ліворуч)\b", re.I)
_LABELED_POINT_RE = re.compile(r"^(?P<label>[^:(]{1,25}?):")
_KSP_LINE_RE = re.compile(r"^КСП\b")
_STAFFING_RE = re.compile(r"^Укомплектованість\b", re.I)
# "Командир …" (а не "Командиру …") - рядок посади в підписі донесення.
_SIGNATURE_RE = re.compile(r"^Командир(?![а-яіїєґ])")


@dataclass
class PbdPoint:
    type: str
    name: str
    mgrs: str
    group: str = ""
    rop: str = ""
    vop: str = ""


@dataclass
class PbdLine:
    kind: str
    mgrs: list
    ref: str = ""


@dataclass
class PbdReport:
    path: str
    report_date: str = ""
    report_time: str = ""
    ksp_place: str = ""
    map_edition: str = ""
    ksp_mgrs: str = ""
    reserve_areas: dict = field(default_factory=dict)
    blocks: dict = field(default_factory=dict)
    points: list = field(default_factory=list)
    lines: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


def _valid_mgrs(text):
    return [m.normalized for m in find_mgrs(text) if m.valid]


def _split_front_edge(text):
    """Опис району "…(MGRS)…, з переднім краєм по рубежу: …(MGRS)…" ->
    (координати району, координати переднього краю)."""
    m = _FRONT_EDGE_RE.search(text)
    if not m:
        return _valid_mgrs(text), []
    return _valid_mgrs(text[:m.start()]), _valid_mgrs(text[m.end():])


def _is_group_header(text):
    return text.endswith(":") and len(text) <= _MAX_GROUP_HEADER_LENGTH and not find_mgrs(text)


def _anchor_name(text):
    for name, pattern in _ANCHORS:
        if pattern.match(text):
            return name
    return None


def _parse_header(texts, report):
    for text in texts[:30]:
        m = _HEADER_RE.search(text)
        if m:
            report.ksp_place = m.group("place").strip()
            report.report_time = m.group("time").replace(":", ".")
            report.report_date = m.group("date")
            edition = _EDITION_RE.search(text[m.end():])
            if edition:
                report.map_edition = edition.group(1)
            return
    report.warnings.append("У шапці ПБД не знайдено \"КСП – <н.п.> ГГ.ХХ ДД.ММ.РРРР\".")


def _section_items(paragraphs):
    """Абзаци розділу 2 (без заголовка) як items."""
    items, inside = [], False
    for p in paragraphs:
        item = paragraph_item(p)
        text = item_text(item)
        if not inside:
            inside = bool(_SECTION_START_RE.match(text))
            continue
        if _SECTION_END_RE.match(text):
            break
        if text:
            items.append(item)
    return items, inside


class _Collector:
    """Один прохід по абзацах розділу 2: розкладає їх по блоках і збирає
    точки/лінії з контекстом (якір, група, РОП, ВОП)."""

    def __init__(self, report):
        self.report = report
        self.blocks = report.blocks
        self.anchor = "intro"
        self.group = ""
        self.rop = ""
        self.rop_part = ""  # "head" | "positions"
        self.vop = ""
        self.uav = False  # всередині "Основні пункти управління БпС"
        self.battle_order_done = False
        self.seen = {}

    def add_block(self, name, item):
        self.blocks.setdefault(name, []).append(item)

    def add_point(self, type_, name, text, group=None):
        mgrs = _valid_mgrs(text)
        label = f"{type_} «{name}»" if name else type_
        if not mgrs:
            self.report.warnings.append(f"ПБД: {label} - координати не розпізнано, точку пропущено.")
            return
        key = (type_, name)
        if key in self.seen:
            if self.seen[key].mgrs != mgrs[0]:
                self.report.warnings.append(f"ПБД: {label} згадано двічі з різними координатами - взято перші.")
            return
        point = PbdPoint(type_, name, mgrs[0], group if group is not None else self.group, self.rop, self.vop)
        self.seen[key] = point
        self.report.points.append(point)

    def add_line(self, kind, mgrs, ref=""):
        if len(mgrs) >= 2:
            self.report.lines.append(PbdLine(kind, mgrs, ref))

    def feed(self, item):
        text = item_text(item)
        anchor = _anchor_name(text)
        if anchor:
            self.anchor, self.group, self.rop, self.rop_part, self.vop, self.uav = anchor, "", "", "", "", False
            return
        handler = getattr(self, f"_feed_{self.anchor}")
        handler(item, text)

    def _feed_intro(self, item, text):
        self.add_block("intro", item)
        area, front = _split_front_edge(text)
        if not any(line.kind == "bro" for line in self.report.lines) and len(area) >= 3:
            self.add_line("bro", area)
            self.add_line("flot", front)
        elif _FRONT_EDGE_RE.match(text) and not any(line.kind == "flot" for line in self.report.lines):
            self.add_line("flot", _valid_mgrs(text))

    def _feed_boundaries(self, item, text):
        self.add_block("boundaries", item)
        side = _SIDE_RE.match(text)
        if side:
            ref = text.split(":", 1)[0].strip()
            self.add_line(f"boundary_{'right' if side.group(1).lower() == 'праворуч' else 'left'}", _valid_mgrs(text), ref)

    def _feed_battle_order(self, item, text):
        if _KSP_LINE_RE.match(text) and not _NAME_RE.search(text):
            mgrs = _valid_mgrs(text)
            if mgrs and not self.report.ksp_mgrs:
                self.report.ksp_mgrs = mgrs[0]
                self.add_point("КСП", "", text, group="Бойовий порядок")
            self.battle_order_done = True
            return
        if _STAFFING_RE.match(text):
            self.battle_order_done = True
            return
        if not self.battle_order_done:
            self.add_block("battle_order", item)

    def _feed_positions(self, item, text):
        self.add_block("positions", item)
        company = _COMPANY_RE.search(text)
        if company:
            self.rop, self.rop_part, self.vop, self.group = company.group("rop").strip().upper(), "head", "", ""
            self.add_block(f"rop:{self.rop}:head", item)
            area, front = _split_front_edge(text)
            self.add_line("rop", area, f"РОП «{self.rop}»")
            self.add_line("flot_rop", front, f"РОП «{self.rop}»")
            return
        if self.rop and self.rop_part == "head" and _FRONT_EDGE_RE.match(text):
            self.add_block(f"rop:{self.rop}:head", item)
            self.add_line("flot_rop", _valid_mgrs(text), f"РОП «{self.rop}»")
            return
        if _MAJOR_GROUP_RE.match(text):
            self.rop, self.rop_part, self.vop = "", "", ""
            self.group = text.rstrip(":").strip()
            self.uav = bool(_UAV_CONTROL_RE.match(text))
            return
        if _is_group_header(text):
            self.group = text.rstrip(":").strip()
            if self.uav and _GROUND_ROBOTICS_RE.match(text):
                self.uav = False
            elif self.uav:
                self.add_block("uav_points", item)
            return
        if self.rop:
            self.rop_part = "positions"
            self.add_block(f"rop:{self.rop}:positions", item)
        if self.uav:
            self.add_block("uav_points", item)
        self._collect_position(text)

    def _collect_position(self, text):
        m = _POSITION_RE.match(text)
        if not m:
            return
        type_ = m.group("type")
        name_match = _NAME_RE.search(text, m.end())
        name = name_match.group("name").strip().upper() if name_match else ""
        if type_ == "ВОП":
            self.vop = name
            area, front = _split_front_edge(text)
            self.add_line("vop", area, f"ВОП «{name}»")
            self.add_line("flot_vop", front, f"ВОП «{name}»")
            return
        if type_ == "РОП":
            area, front = _split_front_edge(text)
            self.add_line("rop", area, f"РОП «{name}»")
            return
        if not name:
            return
        self.add_point(type_, name, text)

    def _feed_reserves(self, item, text):
        self.add_block("reserves", item)
        labeled = _LABELED_POINT_RE.match(text)
        if labeled and _valid_mgrs(text):
            label = labeled.group("label").strip()
            self.add_point(label, "", text, group="Резерви")
            # "БнГ: район зосередження … координатами: лісосмуга (MGRS)." -> "лісосмуга (MGRS)"
            self.report.reserve_areas[label] = text.rsplit(":", 1)[-1].strip().rstrip(".;").strip()

    def _feed_observation(self, item, text):
        self.add_block("observation", item)
        group = self.group
        self.group = "Спостережні позиції"
        self._collect_position(text)
        self.group = group


def read_pbd(path):
    report = PbdReport(path=path)
    paragraphs = body_paragraphs(Document(path))
    texts = [item_text(paragraph_item(p)) for p in paragraphs[:30]]
    _parse_header(texts, report)

    items, found = _section_items(paragraphs)
    if not found:
        report.warnings.append("У ПБД не знайдено розділ «ПОЛОЖЕННЯ ТА СТАН ПІДРОЗДІЛІВ НАШИХ ВІЙСЬК».")
        return report
    collector = _Collector(report)
    for item in items:
        collector.feed(item)
    signature = _signature_items(paragraphs)
    if signature:
        report.blocks["signature"] = signature
    for name in ("battle_order", "positions", "reserves"):
        if not report.blocks.get(name):
            report.warnings.append(f"У розділі 2 ПБД порожній або відсутній підблок «{name}».")
    return report


def _signature_items(paragraphs):
    """Останній рядок "Командир …" і наступний непорожній (звання й ім'я) -
    підпис командира, яким підписано донесення."""
    items = [paragraph_item(p) for p in paragraphs]
    for index in range(len(items) - 1, -1, -1):
        if _SIGNATURE_RE.match(item_text(items[index])):
            following = next((item for item in items[index + 1:] if item_text(item)), None)
            result = [items[index]]
            if following is not None:
                result.append(_collapse_tabs(following))
            return result
    return []


def _collapse_tabs(item):
    text = re.sub(r"[ \t]*\t[ \t]*", "\t", item_text(item))
    if isinstance(item, list) and len(item) > 1:
        phrases = [phrase for phrase in item[1] if phrase in text]
        return [text, phrases] if phrases else text
    return text
