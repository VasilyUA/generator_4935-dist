"""Бойовий наказ: тексти bn_document з data_auto.json + блоки з БР, РР і кешу ПБД +
штабні рішення data_auto.json → staff -> output/Бойовий наказ КБ <підрозділ> <дата>.docx
(і виправлена таблиця сигналів, якщо сигнали БН з нею не збігаються).

Якщо є resources/Стиль БН.docx (прототипи абзаців зі зразка) - документ
будується в ньому, тож форматування, колонтитули й нумерація збігаються зі
зразком; інакше - базове форматування python-docx."""
import os
import re
from dataclasses import dataclass, field

from constants import BN_FILE_NAME_TEMPLATE
from content.data_store import PREVIOUS_FIELDS, normalize_name
from content.docx_blocks import item_text
from content.position_refresh import PositionRefresher
from content.signal_table import signal_key
from content.staff import apply_delta, duplicate_coordinates, invalid_coordinates, load_staff, staff_values
from generators.document_builder import (
    _SIGNAL_RE, ORIGIN_MANUAL, ORIGIN_REFRESHED, ORIGIN_SOURCE, ORIGIN_STAFF, ORIGIN_STANDARD, DocumentBuilder,
    find_stale_position_names,
)
from generators.generate_signals import find_duplicates, write_corrected_table
from generators.style_template import StyleTemplate
from utils.docx_utils import save_docx_safely
from utils.mgrs import find_mgrs

ORIGIN_TITLES = {
    ORIGIN_SOURCE: "з БР/ПБД/РР",
    ORIGIN_REFRESHED: "статичний текст з автоматично підставленими даними",
    ORIGIN_STANDARD: "незмінний типовий текст",
    ORIGIN_STAFF: "штабні рішення з data_auto.json → staff",
    ORIGIN_MANUAL: "потребує перевірки штабом",
}
SOURCE_BR, SOURCE_PBD, SOURCE_PREVIOUS = "br", "pbd", "previous"
SOURCE_TITLES = {
    SOURCE_BR: "БР", SOURCE_PBD: "ПБД / РОП_ВОП", SOURCE_PREVIOUS: "попередній наказ (data_auto.json → previous)",
}
_SOURCE_SHORT = {SOURCE_BR: "БР", SOURCE_PBD: "ПБД", SOURCE_PREVIOUS: "попередній наказ"}
# Порядок джерел для значень, що є і в БР, і в ПБД; що першим - БР чи ПБД -
# обирається при запуску, попередній наказ - завжди останній.
PRIORITY_BR_FIRST = (SOURCE_BR, SOURCE_PBD, SOURCE_PREVIOUS)
PRIORITY_PBD_FIRST = (SOURCE_PBD, SOURCE_BR, SOURCE_PREVIOUS)
PRIORITY_TITLES = {
    "ksp_mgrs": "Координати КСП",
    "zkp_mgrs": "Координати ЗКП",
    "ksp_place": "Н.п. КСП",
    "map_edition": "Видання карти",
}
# Назва н.п. - слова великими літерами (кирилиця); "КСП"/"ЗКП" - не назва.
_SETTLEMENT_RE = re.compile(r"[А-ЯІЇЄҐ][А-ЯІЇЄҐ'’\-]{2,}(?:\s+[А-ЯІЇЄҐ][А-ЯІЇЄҐ'’\-]{2,})*")


@dataclass
class BnExtras:
    """Джерела, крім БР і ПБД (усі необов'язкові)."""
    rr: object = None                              # content.rr_reader.RrOrder
    renames: list = field(default_factory=list)    # [RenameReport] від старих до нових
    signals: object = None                         # content.signal_table.SignalTable
    delta: object = None                           # content.delta_reader.DeltaExport
    icons: dict = field(default_factory=dict)      # {sidc: PNG} з KMZ-експорту Дельти


@dataclass
class BnResult:
    path: str
    used_blocks: dict = field(default_factory=dict)
    warnings: list = field(default_factory=list)
    coverage: dict = field(default_factory=dict)
    manual: list = field(default_factory=list)
    automatic_percent: float = 0.0
    staff: list = field(default_factory=list)            # абзаци зі штабними рішеннями
    staff_entries: list = field(default_factory=list)    # StaffEntry, використані в тексті
    unused_staff: list = field(default_factory=list)     # StaffEntry без плейсхолдера в тексті
    empty_staff: list = field(default_factory=list)      # StaffEntry без координат - абзац пропущено
    priority: tuple = PRIORITY_BR_FIRST                  # порядок джерел, обраний при запуску
    priority_values: dict = field(default_factory=dict)  # {ключ PRIORITY_TITLES: (значення, джерело)}
    signals: object = None                               # generators.generate_signals.SignalsResult
    signal_stats: dict = field(default_factory=dict)
    bn_date: object = None                               # datetime.date наказу

    @property
    def staff_values(self):
        return sum(len(entry.staff_values) for entry in self.manual)

    @property
    def structured_values(self):
        return sum(len(entry.structured_values) for entry in self.staff)


def title_case(text):
    """"НАЗВА НАСЕЛЕНОГО ПУНКТУ" -> "Назва Населеного Пункту" (str.title()
    ламає слова з апострофом)."""
    return " ".join(word[:1].upper() + word[1:].lower() for word in (text or "").split(" "))


def _first_mgrs(texts):
    for text in texts:
        for match in find_mgrs(text):
            if match.valid:
                return match.normalized
    return ""


def priority_order(first):
    """Відповідь на питання при запуску ("br" / "pbd") -> порядок джерел;
    попередній наказ - завжди останній."""
    return PRIORITY_PBD_FIRST if first == SOURCE_PBD else PRIORITY_BR_FIRST


def settlement_name(text):
    """"КСП A0001 - півд.-зах. НАЗВА (MGRS)." -> "НАЗВА": н.п. - слова великими
    літерами після тире (позначення "КСП"/"ЗКП" і скорочення напрямків не беруться)."""
    parts = re.split(r"\s[-–—]\s", text or "", maxsplit=1)
    names = [m.group(0) for m in _SETTLEMENT_RE.finditer(parts[-1]) if m.group(0) not in ("КСП", "ЗКП")]
    return names[-1] if names else ""


def br_priority_values(br):
    """Значення з БР: координати й н.п. з рядків КСП / "Запасне КСП" наприкінці
    розділу, видання карти - з шапки."""
    texts = [item_text(item) for item in (br.blocks.get("br_ksp", []) if br else [])]
    reserve = [text for text in texts if re.search(r"запасн|ЗКП", text, re.IGNORECASE)]
    main = [text for text in texts if text not in reserve]
    return {
        "ksp_mgrs": _first_mgrs(main),
        "zkp_mgrs": _first_mgrs(reserve),
        "ksp_place": next((name for name in map(settlement_name, main) if name), ""),
        "map_edition": getattr(br, "map_edition", "") if br else "",
    }


def priority_sources(br, pbd_cache, bn):
    """{джерело: {ключ PRIORITY_TITLES: значення}} - що є в БР, ПБД (РОП_ВОП
    для ЗКП) і попередньому наказі (data_auto.json → previous)."""
    cache, bn = pbd_cache or {}, bn or {}
    return {
        SOURCE_BR: br_priority_values(br),
        SOURCE_PBD: {key: cache.get(key, "") for key in PRIORITY_TITLES},
        SOURCE_PREVIOUS: {key: bn.get(previous, "") for key, previous in PREVIOUS_FIELDS.items()},
    }


def resolve_priority_values(br, pbd_cache, bn, order=PRIORITY_BR_FIRST):
    """Кожне значення PRIORITY_TITLES - з першого джерела за порядком order, у
    якому воно є. -> {ключ: (значення, джерело)}; ("", "") - немає ніде."""
    sources = priority_sources(br, pbd_cache, bn)
    order = tuple(s for s in order if s in sources) + tuple(s for s in PRIORITY_BR_FIRST if s not in order)
    return {key: next(((sources[s][key], s) for s in order if sources[s].get(key)), ("", "")) for key in PRIORITY_TITLES}


def priority_overview(br, pbd_cache, bn):
    """Рядки для консолі перед питанням про пріоритет: що є в кожному джерелі."""
    sources = priority_sources(br, pbd_cache, bn)
    return [
        f"{title}: " + "; ".join(f"{_SOURCE_SHORT[s]} - {sources[s].get(key) or 'немає'}" for s in PRIORITY_BR_FIRST)
        for key, title in PRIORITY_TITLES.items()
    ]


def build_context(data, br, pbd_cache, bn_date, bn_time, bn_number, priority_values=None):
    unit = data.get("unit", {})
    bn = data.get("bn") or {}
    values = priority_values or resolve_priority_values(br, pbd_cache, bn)
    ksp_place = values["ksp_place"][0]
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
        "ksp_mgrs": values["ksp_mgrs"][0],
        "zkp_mgrs": values["zkp_mgrs"][0],
        "map_edition": values["map_edition"][0],
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


def known_signal_keys(br, signal_table):
    """Сигнали, підтверджені джерелами: усі сигнали тексту БР і таблиці сигналів."""
    keys = set(signal_table.keys) if signal_table else set()
    for items in (br.blocks.values() if br else []):
        for item in items:
            keys.update(signal_key(m.group("name")) for m in _SIGNAL_RE.finditer(item_text(item)))
    return keys


def _block_texts(blocks):
    return [item_text(item) for items in (blocks or {}).values() for item in items]


def source_coordinates(br, pbd_cache, rr=None):
    """Усі координати поточних БР, РР і ПБД/РОП_ВОП - підтверджені джерелами."""
    texts = _block_texts(br.blocks if br else {}) + _block_texts(rr.blocks if rr else {})
    texts += _block_texts(pbd_cache.get("blocks"))
    found = {m.normalized for text in texts for m in find_mgrs(text) if m.valid}
    found |= {point.get("mgrs") for point in pbd_cache.get("points", []) if point.get("mgrs")}
    found |= {mgrs for line in pbd_cache.get("lines", []) for mgrs in line.get("mgrs", [])}
    for key in ("ksp_mgrs", "zkp_mgrs"):
        if pbd_cache.get(key):
            found.add(pbd_cache[key])
    return found


def collect_blocks(br, pbd_cache, rr=None):
    blocks = dict(br.blocks) if br else {}
    if rr:
        blocks.update(rr.blocks)
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


def ordered_renames(reports):
    """Перейменування з усіх донесень у хронологічному порядку (за датою, далі - номером)."""
    def key(report):
        day, month, year = (report.date.split(".") + ["", "", ""])[:3] if report.date else ("", "", "")
        return (year, month, day, int(report.number) if str(report.number).isdigit() else 0)
    return [rename for report in sorted(reports or [], key=key) for rename in report.renames]


def build_refresher(data, pbd_cache, renames=()):
    ksp_name = (data.get("bn") or {}).get("KSP_NAME", "")
    aliases = {ksp_name: ("КСП", "")} if ksp_name else {}
    return PositionRefresher(pbd_cache.get("points", []), aliases, renames=renames,
                             current_names=known_position_names(pbd_cache, [ksp_name]))


def _strip_name(mention):
    return mention.split("«", 1)[-1].rstrip("»").upper()


def generate_bn(data, br, pbd_cache, bn_date, bn_time, bn_number, output_dir, style_path=None, extras=None,
                priority=PRIORITY_BR_FIRST):
    """priority - порядок джерел для значень, що є і в БР, і в ПБД
    (PRIORITY_BR_FIRST / PRIORITY_PBD_FIRST, обирається при запуску)."""
    extras = extras or BnExtras()
    spec = (data.get("bn") or {}).get("document")
    if not spec:
        raise ValueError("Немає тексту бойового наказу (resources/data_auto.json → bn_document).")

    warnings = list(br.warnings) if br else ["БР не знайдено - блоки br_* лишаться порожніми."]
    for source in (extras.rr, extras.signals):
        warnings.extend(getattr(source, "warnings", []) or [])
    style = StyleTemplate(style_path) if style_path and os.path.isfile(style_path) else None
    if style_path and style is None:
        warnings.append(f"Немає {os.path.basename(style_path)} - документ зібрано без стилю зразка.")

    staff_entries = load_staff(data)
    warnings.extend(f"Штабні рішення: {note}" for note in apply_delta(staff_entries, extras.delta))
    warnings.extend(f"Штабні рішення: {problem}" for problem in invalid_coordinates(staff_entries))
    warnings.extend(f"Штабні рішення з однаковими координатами (описка?): {', '.join(same)}"
                    for same in duplicate_coordinates(staff_entries))

    bn = data.get("bn") or {}
    ksp_name = bn.get("KSP_NAME", "")
    known = known_position_names(pbd_cache, [ksp_name])
    values = resolve_priority_values(br, pbd_cache, bn, priority)
    builder = DocumentBuilder(
        build_context(data, br, pbd_cache, bn_date, bn_time, bn_number, values), collect_blocks(br, pbd_cache, extras.rr),
        style=style, refresher=build_refresher(data, pbd_cache, ordered_renames(extras.renames)), known_names=known,
        source_mgrs=source_coordinates(br, pbd_cache, extras.rr), staff=staff_values(staff_entries),
        known_signals=known_signal_keys(br, extras.signals),
    )
    doc = builder.build(spec)

    for key, (value, _source) in values.items():
        if not value:
            warnings.append(f"{PRIORITY_TITLES[key]}: немає ні в БР, ні в ПБД, ні в попередньому наказі "
                            f"(data_auto.json → previous.{PREVIOUS_FIELDS[key]}).")
    for name in builder.missing_blocks:
        warnings.append(f"Блок «{name}» порожній або відсутній - у документі позначено жовтим.")
    for name in builder.fallback_blocks:
        warnings.append(f"Блок «{name}» порожній - використано запасний текст з data_auto.json.")
    for key in sorted(builder.unknown_placeholders):
        warnings.append(f"Невідомий плейсхолдер {{{key}}} - лишився в тексті як є.")
    for old, new in dict.fromkeys(builder.reported_renames):
        warnings.append(f"Перейменовано за позатерміновим донесенням: {old} -> {new}.")
    for old, new in dict.fromkeys(builder.renamed):
        warnings.append(f"Позицію перейменовано за координатами: {old} -> {new}.")
    renamed_old = {_strip_name(old) for old, _new in builder.renamed + builder.reported_renames}
    stale = [name for name in dict.fromkeys(
        [_strip_name(mention) for mention in builder.unknown_mentions]
        + find_stale_position_names(spec, known, normalize_name)
    ) if name not in renamed_old]
    if stale:
        warnings.append("У статичному тексті БН згадано позиції, яких немає в ПБД/РОП_ВОП: " + ", ".join(stale))

    battalion = data.get("unit", {}).get("SHORT_UNIT_BATTALION", "")
    os.makedirs(output_dir, exist_ok=True)
    date_text = bn_date.strftime("%d.%m.%Y")
    path = os.path.join(output_dir, BN_FILE_NAME_TEMPLATE.format(battalion=battalion.upper(), date=date_text))
    save_docx_safely(doc, path)

    signals = None
    if extras.signals is not None:
        signals = write_corrected_table(extras.signals, builder.signal_mentions, output_dir,
                                        f"бойового наказу №{bn_number} від {date_text}")
    for name, meanings in find_duplicates(builder.signal_mentions):
        warnings.append(f"Сигнал {name} у БН стоїть за різними діями: " + "; ".join(meanings))
    mentions = {mention.key: mention for mention in builder.signal_mentions}
    signal_stats = {
        "total": len(mentions),
        "known": sum(1 for mention in mentions.values() if mention.known),
        "table": len(extras.signals.rows) if extras.signals else 0,
    }
    used, skipped = builder.used_staff, builder.skipped_staff
    return BnResult(
        path, builder.used_blocks, warnings, builder.coverage(), builder.manual_paragraphs(), builder.automatic_percent(),
        staff=builder.staff_paragraphs(), staff_entries=list(used.values()),
        unused_staff=[entry for entry in staff_entries if entry.placeholder not in used and entry.placeholder not in skipped],
        empty_staff=list(skipped.values()), priority=tuple(priority), priority_values=values, signals=signals,
        signal_stats=signal_stats, bn_date=bn_date,
    )


def review_lines(result):
    """Текст файлу "Потребує_ручної_перевірки.txt" для БН."""
    lines = [
        "БОЙОВИЙ НАКАЗ",
        f"Автоматично написано: {result.automatic_percent}% тексту. Решта - штабні значення: "
        f"{result.structured_values} - зі штабних рішень data_auto.json → staff (у {len(result.staff)} абзацах), "
        f"{result.staff_values} - неструктуровані, перенесені з попереднього наказу (у {len(result.manual)} абзацах).",
        "Походження абзаців (частка тексту):",
    ]
    for origin, title in ORIGIN_TITLES.items():
        lines.append(f"  - {title}: {result.coverage.get(origin, 0)}%")
    if result.priority_values:
        order = " -> ".join(_SOURCE_SHORT.get(source, source) for source in result.priority or PRIORITY_BR_FIRST)
        lines += ["", f"Дані, що є і в БР, і в ПБД (пріоритет, обраний при запуску: {order}):"]
        for key, (value, source) in result.priority_values.items():
            lines.append(f"  - {PRIORITY_TITLES.get(key, key)}: {SOURCE_TITLES.get(source, 'немає') if value else 'немає'}")
    if result.signal_stats:
        stats = result.signal_stats
        lines += ["", f"Сигнали: у тексті БН {stats['total']} різних, з них у таблиці сигналів чи БР - {stats['known']}."]
        if result.signals is not None and result.signals.path:
            lines.append(f"  Виправлена таблиця сигналів: {result.signals.path}")
            lines += [f"  - змінено ({group}): «{meaning}» {old} -> {new}" for group, meaning, old, new in result.signals.updated]
            lines += [f"  - додано: {name} - {meaning}" for name, meaning in result.signals.added]
    if result.warnings:
        lines += ["", "Попередження:"] + [f"  - {w}" for w in result.warnings]
    if result.staff_entries or result.unused_staff or result.empty_staff:
        lines += ["", "Штабні рішення (data_auto.json → staff; плейсхолдер у тексті - {staff:<група>:<ключ>}):"]
        for entry in result.staff_entries:
            lines.append(f"  - {entry.placeholder} «{entry.title}»: {len(entry.mgrs)} точок, джерело - {entry.source}")
        for entry in result.empty_staff:
            lines.append(f"  - {entry.placeholder} «{entry.title}»: координат немає - абзац не надруковано (заповніть mgrs)")
        for entry in result.unused_staff:
            lines.append(f"  - {entry.placeholder} «{entry.title}»: не згадано в тексті БН (лише на карті)")
    if result.manual:
        lines += ["", "Штабні значення, яких немає в джерелах (перенесено з попереднього наказу) - перевірте:"]
        for entry in result.manual:
            lines.append(f"  [{entry.path}] {entry.text[:90]}…")
            lines.append(f"      значення: {'; '.join(entry.staff_values)}")
    return lines
