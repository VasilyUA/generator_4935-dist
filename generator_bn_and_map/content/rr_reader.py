"""Розпорядження з розвідувального забезпечення (РР) -> блоки тексту для БН.

Розділ "ЗОНИ … ВЕДЕННЯ РОЗВІДКИ" поділено на етапи абзацами "Під час ведення
оборонного бою батальйоном на <…> рубежі …" (по порядку: 1, 2, 3 …). Для
кожного етапу N:
  rr_stage:N:all      - усі абзаци етапу (без самого заголовка етапу);
  rr_stage:N:zone     - зона оглядової розвідки / відповідальності (з
                        продовженням, якщо перелік винесено в наступний абзац);
  rr_stage:N:detailed - зона детальної розвідки;
  rr_stage:N:rou      - райони особливої уваги (РОУ …);
  rr_stage:N:rpr      - райони повітряної розвідки (РПР …);
  rr_stage:N:tasks    - завдання розвідки цього етапу з розділу "МЕТА, ОСНОВНІ
                        ЗАВДАННЯ РОЗВІДКИ".
Абзаци переносяться дослівно (з жирними фразами)."""
import re
from dataclasses import dataclass, field

from content.docx_blocks import item_text, merge_items, read_paragraph_items

_ZONES_HEADING_RE = re.compile(r"^(?:\d+\.\s*)?ЗОНИ\b.*РОЗВІДКИ", re.IGNORECASE)
_TASKS_HEADING_RE = re.compile(r"^(?:\d+\.\s*)?МЕТА\b.*ЗАВДАННЯ\s+РОЗВІДКИ", re.IGNORECASE)
_NUMBERED_HEADING_RE = re.compile(r"^\d+\.\s*[А-ЯІЇЄҐ][А-ЯІЇЄҐ ,()’'\-]{8,}$")
_STAGE_RE = re.compile(r"^Під\s+час\s+ведення\s+оборонного\s+бою", re.IGNORECASE)
_ZONE_RE = re.compile(r"^Зон\w*\s+(?:оглядової\s+розвідки|відповідальності)", re.IGNORECASE)
_DETAILED_RE = re.compile(r"^Зон\w*\s+детальної\s+розвідки", re.IGNORECASE)
_HEADER_RE = re.compile(r"РОЗПОРЯДЖЕННЯ.*?№\s*(?P<number>\d+)", re.IGNORECASE)
_DATE_RE = re.compile(r"(?<!\d)(\d{2}\.\d{2}\.\d{4})(?!\d)")


@dataclass
class RrOrder:
    path: str
    number: str = ""
    date: str = ""
    blocks: dict = field(default_factory=dict)
    warnings: list = field(default_factory=list)


def _is_heading(text):
    """"3. МЕТА …" або рядок великими літерами без цифр ("КОРОТКІ ДАНІ …");
    "РОУ №1 – НАЗВА (MGRS)" теж великими, але з цифрами - це не заголовок."""
    return bool(_NUMBERED_HEADING_RE.match(text)) or (text.isupper() and len(text) > 15 and not re.search(r"\d", text))


def _section(items, start_re):
    """Абзаци від заголовка start_re (не включно) до наступного заголовка розділу."""
    texts = [item_text(item).strip() for item in items]
    start = next((i for i, text in enumerate(texts) if start_re.match(text)), None)
    if start is None:
        return []
    end = next((i for i in range(start + 1, len(texts)) if _is_heading(texts[i]) and not _STAGE_RE.match(texts[i])), len(texts))
    return items[start + 1:end]


def _split_stages(items):
    stages, current = [], None
    for item in items:
        if _STAGE_RE.match(item_text(item).strip()):
            current = []
            stages.append(current)
        elif current is not None:
            current.append(item)
    return stages


def _stage_blocks(number, items):
    blocks = {f"rr_stage:{number}:all": list(items)}
    zone, detailed, rou, rpr = [], [], [], []
    index = 0
    while index < len(items):
        item = items[index]
        text = item_text(item).strip()
        if _ZONE_RE.match(text):
            # "Зону відповідальності … мати –" + перелік наступним абзацом.
            if re.search(r"[–\-:]\s*$", text) and index + 1 < len(items):
                item = merge_items(item, items[index + 1])
                index += 1
            zone.append(item)
        elif _DETAILED_RE.match(text):
            detailed.append(item)
        elif re.match(r"^РОУ\b", text):
            rou.append(item)
        elif re.match(r"^РПР\b", text):
            rpr.append(item)
        index += 1
    for part, part_items in (("zone", zone), ("detailed", detailed), ("rou", rou), ("rpr", rpr)):
        if part_items:
            blocks[f"rr_stage:{number}:{part}"] = part_items
    return blocks


def read_rr(path):
    items = read_paragraph_items(path)
    order = RrOrder(path)
    for item in items[:15]:
        text = item_text(item)
        m = _HEADER_RE.search(text)
        if m:
            order.number = m.group("number")
            dates = _DATE_RE.findall(text)
            order.date = dates[-1] if dates else ""
            break
    zones = _split_stages(_section(items, _ZONES_HEADING_RE))
    if not zones:
        order.warnings.append("У РР не знайдено розділу «ЗОНИ … ВЕДЕННЯ РОЗВІДКИ» з етапами - блоки rr_stage порожні.")
    for number, stage_items in enumerate(zones, start=1):
        order.blocks.update(_stage_blocks(number, stage_items))
    for number, stage_items in enumerate(_split_stages(_section(items, _TASKS_HEADING_RE)), start=1):
        if stage_items:
            order.blocks[f"rr_stage:{number}:tasks"] = stage_items
    return order
