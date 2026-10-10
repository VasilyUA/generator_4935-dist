"""Автооновлення координат у СТАТИЧНОМУ тексті БН.

Штабний текст (способи розгрому, евакуація, завдання тиловим) згадує позиції
у вигляді `ВП «НАЗВА» за координатами (MGRS)` чи `MGRS - ВП «НАЗВА»`. Для
кожної такої згадки координати замінюються поточними з кешу ПБД (за назвою);
якщо назви вже немає, але в межах кількох метрів від старих координат є
позиція з іншою назвою - позиція перейменована, оновлюється й назва.
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
_MAX_GAP_AFTER = 140   # від назви до координат після неї
_MAX_GAP_BEFORE = 8    # "MGRS — ВП «НАЗВА»"
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
    unknown: list = field(default_factory=list)
    auto_mgrs: set = field(default_factory=set)


class PositionRefresher:
    def __init__(self, points, aliases=None):
        """points - кеш pbd.points; aliases - {"КОДОВА НАЗВА": (тип, назва)} для
        позицій, які в ПБД записані без назви (напр. КСП батальйону)."""
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

    def refresh(self, text):
        result = RefreshResult(text)
        if not text or not self.points:
            return result
        mentions = list(_MENTION_RE.finditer(text))
        coordinates = [m for m in find_mgrs(text) if m.valid]
        replacements = []  # (start, end, new_text)
        used = set()
        for index, mention in enumerate(mentions):
            next_start = mentions[index + 1].start() if index + 1 < len(mentions) else len(text)
            target = next((c for c in coordinates if c.start >= mention.end() and c.start - mention.end() <= _MAX_GAP_AFTER
                           and c.start < next_start and c.start not in used), None)
            if target is None:
                target = next((c for c in reversed(coordinates) if c.end <= mention.start()
                               and mention.start() - c.end <= _MAX_GAP_BEFORE and c.start not in used), None)
            if target is None:
                continue
            used.add(target.start)
            type_, name = mention.group("type"), mention.group("name").strip()
            if type_ in _AREA_TYPES:
                used.discard(target.start)
                continue
            point = self._find(type_, name)
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
        for start, end, new in sorted(replacements, reverse=True):
            text = text[:start] + new + text[end:]
        result.text = text
        return result
