"""Автооновлення координат у СТАТИЧНОМУ тексті БН.

Штабний текст (способи розгрому, евакуація, завдання тиловим) згадує позиції
у вигляді `ВП «НАЗВА» за координатами (MGRS)` чи `MGRS - ВП «НАЗВА»`. Для
кожної такої згадки координати замінюються поточними з кешу ПБД (за назвою);
якщо назви вже немає, позиція могла бути перейменована: спершу шукається в
переліку перейменувань з позатермінових донесень (content/rename_reader.py),
потім - позиція з іншою назвою в межах кількох метрів від старих координат.
Згадки, які не вдалось зіставити, лишаються як є і повертаються як "застарілі"."""
import re
from dataclasses import dataclass, field

from content.data_store import distance_m, normalize_name
from utils.mgrs import find_mgrs

_TYPES = "СПАР|КСП|ВОП|РОП|ХАБ|ВП|ПВ|СП|ВЗ|ТЗ|ПУ"
_OPEN = "«\"“„‘’'"
_CLOSE = "»\"”“’‘'"
_MENTION_RE = re.compile(
    rf"(?<!\w)(?P<type>{_TYPES})(?!\w)[^{_OPEN}{_CLOSE}()\n]{{0,40}}?"
    rf"[{_OPEN}]{{1,2}}\s*(?P<name>[^{_OPEN}{_CLOSE}\n]{{2,40}}?)\s*[{_CLOSE}]{{1,2}}"
)
# "ротний опорний пункт «НАЗВА»" - згадка району без скорочення (лише для
# перейменувань, координати таких згадок не оновлюються).
_STRONGPOINT_RE = re.compile(
    rf"(?<!\w)(?P<kind>ротн|взводн)\w*\s+опорн\w*\s+пункт\w*\s*"
    rf"[{_OPEN}]{{1,2}}\s*(?P<name>[^{_OPEN}{_CLOSE}\n]{{2,40}}?)\s*[{_CLOSE}]{{1,2}}",
    re.IGNORECASE,
)
_MAX_GAP_AFTER = 140   # від назви до координат після неї
_MAX_GAP_BEFORE = 8    # "MGRS — ВП «НАЗВА»"
# "ВП «X» на проміжний рубіж за координатами (…)" - координати НОВОГО місця
# (куди позиція переміщується), а не самої позиції: якщо між назвою і
# координатами є слова руху чи цілі, координати до позиції не прив'язуються.
MOVEMENT_GAP_RE = re.compile(
    r";|(?<!\w)(?:на|до)(?!\w)|напрямк|руб[еі]ж|перемі|відх|відій|висун|здійсн|зайн|займ|прикритт|сигнал",
    re.IGNORECASE,
)


def is_position_gap(text):
    """Чи текст між назвою позиції і координатами - лише опис її місця
    ("за координатами", "в районі з центром", "мати околиці н.п. «…»")."""
    return not MOVEMENT_GAP_RE.search(text or "")
_SAME_POSITION_METERS = 30
# ВОП/РОП - райони (кілька точок), не точки: їх координати не оновлюються.
_AREA_TYPES = {"ВОП", "РОП"}
# Типи, які можуть "перейти" один в один при перейменуванні (ВП <-> ВЗ тощо).
_POSITION_FAMILY = {"ВП", "ПВ", "СП", "ВЗ"}


@dataclass
class RefreshResult:
    text: str
    refreshed: int = 0
    renamed: list = field(default_factory=list)
    reported: list = field(default_factory=list)
    unknown: list = field(default_factory=list)
    auto_mgrs: set = field(default_factory=set)


def build_rename_map(renames):
    """[Rename(тип, стара, нова), …] у хронологічному порядку -> {СТАРА: НОВА}
    з урахуванням ланцюжків (А -> Б, потім Б -> В дає А -> В)."""
    mapping = {}
    for rename in renames:
        old, new = normalize_name(rename.old), rename.new
        if not old or not new:
            continue
        for key, value in list(mapping.items()):
            if normalize_name(value) == old:
                mapping[key] = new
        mapping[old] = new
    return mapping


class PositionRefresher:
    def __init__(self, points, aliases=None, renames=(), current_names=()):
        """points - кеш pbd.points; aliases - {"КОДОВА НАЗВА": (тип, назва)} для
        позицій, які в ПБД записані без назви (напр. КСП батальйону); renames -
        перейменування з позатермінових донесень (Rename), від старих до нових;
        current_names - чинні назви (зокрема РОП/ВОП), які не перейменовуються."""
        self.points = [p for p in points if p.get("mgrs")]
        self.by_name = {}
        for point in self.points:
            key = normalize_name(point.get("name", ""))
            if key:
                self.by_name.setdefault(key, []).append(point)
        for alias, (type_, name) in (aliases or {}).items():
            target = next((p for p in self.points if p.get("type") == type_ and p.get("name", "") == name), None)
            if target is not None and normalize_name(alias):
                self.by_name.setdefault(normalize_name(alias), []).insert(0, target)
        self.renames = build_rename_map(renames)
        self.current_names = {normalize_name(p.get("name", "")) for p in points if p.get("name")}
        self.current_names |= {normalize_name(n) for n in current_names if n}
        self.current_names.discard("")

    def _reported_name(self, name):
        """Нова назва з позатермінового донесення, якщо name - стара назва."""
        return self.renames.get(normalize_name(name))

    def _find(self, type_, name):
        candidates = self.by_name.get(normalize_name(name), [])
        same_type = [p for p in candidates if p.get("type") == type_]
        return (same_type or candidates or [None])[0]

    def _nearest(self, type_, mgrs):
        best = None
        for point in self.points:
            compatible = point.get("type") == type_ or {point.get("type"), type_} <= _POSITION_FAMILY
            if not compatible:
                continue
            distance = distance_m(point["mgrs"], mgrs)
            if distance <= _SAME_POSITION_METERS and (best is None or distance < best[0]):
                best = (distance, point)
        return best[1] if best else None

    def _rename_strongpoints(self, text, replacements, result, taken):
        """"ротний опорний пункт «СТАРА»" -> нова назва з донесення."""
        for m in _STRONGPOINT_RE.finditer(text):
            name = m.group("name").strip()
            if m.start("name") in taken or normalize_name(name) in self.current_names:
                continue
            new = self._reported_name(name)
            if new:
                kind = "РОП" if m.group("kind").lower().startswith("ротн") else "ВОП"
                replacements.append((m.start("name"), m.end("name"), new))
                result.reported.append((f"{kind} «{name}»", f"{kind} «{new}»"))

    def refresh(self, text):
        result = RefreshResult(text)
        if not text or not (self.points or self.renames):
            return result
        mentions = list(_MENTION_RE.finditer(text))
        coordinates = [m for m in find_mgrs(text) if m.valid]
        replacements = []  # (start, end, new_text)
        used = set()
        for index, mention in enumerate(mentions):
            type_, name = mention.group("type"), mention.group("name").strip()
            reported = None
            if self._find(type_, name) is None and normalize_name(name) not in self.current_names:
                reported = self._reported_name(name)
                if reported:
                    replacements.append((mention.start("name"), mention.end("name"), reported))
                    result.reported.append((f"{type_} «{name}»", f"{type_} «{reported}»"))
                    name = reported
            next_start = mentions[index + 1].start() if index + 1 < len(mentions) else len(text)
            target = next((c for c in coordinates if c.start >= mention.end() and c.start - mention.end() <= _MAX_GAP_AFTER
                           and c.start < next_start and c.start not in used), None)
            if target is not None and not is_position_gap(text[mention.end():target.start]):
                target = None
            if target is None:
                target = next((c for c in reversed(coordinates) if c.end <= mention.start()
                               and mention.start() - c.end <= _MAX_GAP_BEFORE and c.start not in used), None)
            if target is None:
                continue
            used.add(target.start)
            if type_ in _AREA_TYPES:
                used.discard(target.start)
                continue
            point = self._find(type_, name)
            if point is None and reported:
                result.unknown.append(f"{type_} «{name}»")
                continue
            if point is None:
                point = self._nearest(type_, target.normalized)
                if point is None or not point.get("name"):
                    result.unknown.append(f"{type_} «{name}»")
                    continue
                result.renamed.append((f"{type_} «{name}»", f"{point.get('type')} «{point['name']}»"))
                replacements.append((mention.start("name"), mention.end("name"), point["name"]))
            replacements.append((target.start, target.end, point["mgrs"]))
            result.auto_mgrs.add(point["mgrs"])
            result.refreshed += 1
        if self.renames:
            self._rename_strongpoints(text, replacements, result, {start for start, _end, _new in replacements})
        for start, end, new in sorted(replacements, reverse=True):
            text = text[:start] + new + text[end:]
        result.text = text
        return result
