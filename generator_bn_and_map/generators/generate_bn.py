"""Бойовий наказ: тексти bn.document з data.json + блоки з БР і кешу ПБД ->
output/Бойовий наказ КБ <підрозділ> <дата>.docx.

Якщо є resources/Стиль БН.docx (прототипи абзаців зі зразка) - документ
будується в ньому, тож форматування, колонтитули й нумерація збігаються зі
зразком; інакше - базове форматування python-docx."""
import os
import re
from dataclasses import dataclass, field

from constants import BN_FILE_NAME_TEMPLATE
from content.data_store import normalize_name
from content.docx_blocks import item_text
from content.position_refresh import PositionRefresher
from generators.document_builder import (
    _SIGNAL_RE, ORIGIN_MANUAL, ORIGIN_REFRESHED, ORIGIN_SOURCE, ORIGIN_STANDARD, DocumentBuilder,
    find_stale_position_names,
)
from generators.style_template import StyleTemplate
from utils.docx_utils import save_docx_safely
from utils.mgrs import find_mgrs

ORIGIN_TITLES = {
    ORIGIN_SOURCE: "з БР/ПБД",
    ORIGIN_REFRESHED: "статичний текст з автоматично підставленими даними",
    ORIGIN_STANDARD: "незмінний типовий текст",
    ORIGIN_MANUAL: "потребує перевірки штабом",
}


@dataclass
class BnResult:
    path: str
    used_blocks: dict = field(default_factory=dict)
    warnings: list = field(default_factory=list)
    coverage: dict = field(default_factory=dict)
    manual: list = field(default_factory=list)
    automatic_percent: float = 0.0

    @property
    def staff_values(self):
        return sum(len(entry.staff_values) for entry in self.manual)


def title_case(text):
    """"НАЗВА НАСЕЛЕНОГО ПУНКТУ" -> "Назва Населеного Пункту" (str.title()
    ламає слова з апострофом)."""
    return " ".join(word[:1].upper() + word[1:].lower() for word in (text or "").split(" "))


def build_context(data, br, pbd_cache, bn_date, bn_time, bn_number):
    unit = data.get("unit", {})
    bn = data.get("bn") or {}
    ksp_place = pbd_cache.get("ksp_place", "")
    context = {
        "bn_number": bn_number,
        "bn_date": bn_date.strftime("%d.%m.%Y"),
        "bn_time": bn_time,
        "bn_year": bn_date.strftime("%Y"),
        "br_number": br.number if br else "",
        "br_date": br.date if br else "",
        "battalion": unit.get("SHORT_UNIT_BATTALION", ""),
        "brigade": unit.get("SHORT_UNIT_BRIGADE", ""),
        "full_unit_but": unit.get("FULL_UNIT_BUT", ""),
        "full_military_unit": unit.get("FULL_MILITARY_UNIT", ""),
        "ksp_place": ksp_place,
        "ksp_place_title": title_case(ksp_place),
        "ksp_mgrs": pbd_cache.get("ksp_mgrs", ""),
        "zkp_mgrs": pbd_cache.get("zkp_mgrs", ""),
        "map_edition": pbd_cache.get("map_edition", ""),
        "ksp_name": bn.get("KSP_NAME", ""),
    }
    for label, area in (pbd_cache.get("reserve_areas") or {}).items():
        context[f"reserve_area_{label}"] = area
    for number, signal in br_stage_signals(br).items():
        context[f"br_signal_{number}"] = signal
    return context


def br_stage_signals(br):
    """Сигнал бригади для кожного етапу БР ("НАЗВА-123"…) - перший сигнал у
    порядку відходу (чи, якщо його немає, у діях) етапу."""
    signals = {}
    if not br:
        return signals
    for name in sorted(br.blocks):
        m = re.match(r"br_stage:(\d+):(withdrawal|actions)$", name)
        if not m or m.group(1) in signals and m.group(2) == "actions":
            continue
        for item in br.blocks[name]:
            found = _SIGNAL_RE.search(item_text(item))
            if found:
                signals[m.group(1)] = re.sub(r"\s*[-–]\s*", "-", found.group("name"))
                break
    return signals


def source_coordinates(br, pbd_cache):
    """Усі координати поточних БР і ПБД/РОП_ВОП - підтверджені джерелами."""
    texts = [item_text(item) for items in (br.blocks.values() if br else []) for item in items]
    texts += [item_text(item) for items in (pbd_cache.get("blocks") or {}).values() for item in items]
    found = {m.normalized for text in texts for m in find_mgrs(text) if m.valid}
    found |= {point.get("mgrs") for point in pbd_cache.get("points", []) if point.get("mgrs")}
    found |= {mgrs for line in pbd_cache.get("lines", []) for mgrs in line.get("mgrs", [])}
    for key in ("ksp_mgrs", "zkp_mgrs"):
        if pbd_cache.get(key):
            found.add(pbd_cache[key])
    return found


def collect_blocks(br, pbd_cache):
    blocks = dict(br.blocks) if br else {}
    for name, items in (pbd_cache.get("blocks") or {}).items():
        blocks[f"pbd:{name}"] = items
    return blocks


def known_position_names(pbd_cache, extra=()):
    names = [point.get("name", "") for point in pbd_cache.get("points", [])]
    for line in pbd_cache.get("lines", []):
        ref = line.get("ref", "")
        if "«" in ref:
            names.append(ref.split("«", 1)[1].rstrip("»"))
    return names + [name for name in extra if name]


def build_refresher(data, pbd_cache):
    ksp_name = (data.get("bn") or {}).get("KSP_NAME", "")
    aliases = {ksp_name: ("КСП", "")} if ksp_name else {}
    return PositionRefresher(pbd_cache.get("points", []), aliases)


def generate_bn(data, br, pbd_cache, bn_date, bn_time, bn_number, output_dir, style_path=None):
    spec = (data.get("bn") or {}).get("document")
    if not spec:
        raise ValueError("У resources/data.json немає bn.document - текстів бойового наказу.")

    warnings = list(br.warnings) if br else ["БР не знайдено - блоки br_* лишаться порожніми."]
    style = StyleTemplate(style_path) if style_path and os.path.isfile(style_path) else None
    if style_path and style is None:
        warnings.append(f"Немає {os.path.basename(style_path)} - документ зібрано без стилю зразка.")
    ksp_name = (data.get("bn") or {}).get("KSP_NAME", "")
    known = known_position_names(pbd_cache, [ksp_name])
    builder = DocumentBuilder(
        build_context(data, br, pbd_cache, bn_date, bn_time, bn_number), collect_blocks(br, pbd_cache),
        style=style, refresher=build_refresher(data, pbd_cache), known_names=known,
        source_mgrs=source_coordinates(br, pbd_cache),
    )
    doc = builder.build(spec)

    for name in builder.missing_blocks:
        warnings.append(f"Блок «{name}» порожній або відсутній - у документі позначено жовтим.")
    for name in builder.fallback_blocks:
        warnings.append(f"Блок «{name}» порожній - використано запасний текст з data.json.")
    for key in sorted(builder.unknown_placeholders):
        warnings.append(f"Невідомий плейсхолдер {{{key}}} - лишився в тексті як є.")
    for old, new in dict.fromkeys(builder.renamed):
        warnings.append(f"Позицію перейменовано за координатами: {old} -> {new}.")
    renamed_old = {old.split("«", 1)[-1].rstrip("»").upper() for old, _new in builder.renamed}
    stale = [name for name in dict.fromkeys(
        [mention.split("«", 1)[-1].rstrip("»").upper() for mention in builder.unknown_mentions]
        + find_stale_position_names(spec, known, normalize_name)
    ) if name not in renamed_old]
    if stale:
        warnings.append("У статичному тексті БН згадано позиції, яких немає в ПБД/РОП_ВОП: " + ", ".join(stale))

    battalion = data.get("unit", {}).get("SHORT_UNIT_BATTALION", "")
    os.makedirs(output_dir, exist_ok=True)
    path = os.path.join(output_dir, BN_FILE_NAME_TEMPLATE.format(battalion=battalion.upper(), date=bn_date.strftime("%d.%m.%Y")))
    save_docx_safely(doc, path)
    return BnResult(path, builder.used_blocks, warnings, builder.coverage(), builder.manual_paragraphs(),
                    builder.automatic_percent())


def review_lines(result):
    """Текст файлу "Потребує_ручної_перевірки.txt" для БН."""
    lines = [
        "БОЙОВИЙ НАКАЗ",
        f"Автоматично написано: {result.automatic_percent}% тексту; перевірити вручну - "
        f"{result.staff_values} штабних значень (координати й сигнали, яких немає в БР/ПБД/РОП_ВОП) "
        f"у {len(result.manual)} абзацах.",
        "Походження абзаців (частка тексту):",
    ]
    for origin, title in ORIGIN_TITLES.items():
        lines.append(f"  - {title}: {result.coverage.get(origin, 0)}%")
    if result.warnings:
        lines += ["", "Попередження:"] + [f"  - {w}" for w in result.warnings]
    if result.manual:
        lines += ["", "Штабні значення, яких немає в джерелах (перенесено з попереднього наказу) - перевірте:"]
        for entry in result.manual:
            lines.append(f"  [{entry.path}] {entry.text[:90]}…")
            lines.append(f"      значення: {'; '.join(entry.staff_values)}")
    return lines
