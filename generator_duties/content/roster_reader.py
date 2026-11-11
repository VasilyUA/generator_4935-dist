"""Штат (xlsx) → слоти: посада, звання, ПІБ, ВОС. Колонки шукаються за
ключовими словами заголовка в перших рядках аркуша (як у сусідніх
проєктах), тож порядок колонок і зайві колонки не важливі."""
import glob
import os
import re
import warnings
from dataclasses import dataclass, field

import openpyxl

from constants import (
    ROSTER_EXCLUDED_PREFIXES, ROSTER_FILE_GLOBS, ROSTER_HEADER_EXCLUDE, ROSTER_HEADER_KEYWORDS,
    ROSTER_HEADER_PREFER, ROSTER_HEADER_SCAN_ROWS,
)
from content.docx_text import clean_text

# Підпис OLE2: старий .xls або xlsx, зашифрований паролем на відкриття.
_OLE2_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_VOS_RE = re.compile(r"\bвос\b")
_COMPACT_RE = re.compile(r"[\s.]+")


class RosterError(Exception):
    pass


@dataclass
class Slot:
    row: int
    position: str
    rank: str = ""
    pib: str = ""
    vos: str = ""
    subunit: str = ""
    # Підрозділ верхнього рівня (рота, батарея, взвод) - склад батальйону в
    # чернетці Положення.
    group: str = ""
    # Посада без ланцюга вищого підрозділу, але з номерами - для листа
    # ознайомлення (заповнює content.positions.group_positions).
    display: str = ""

    @property
    def vacant(self):
        return not self.pib


@dataclass
class Roster:
    path: str
    sheet: str
    slots: list = field(default_factory=list)
    headers: dict = field(default_factory=dict)


def find_roster_file(resources_dir, configured_name=""):
    if configured_name:
        path = os.path.join(resources_dir, configured_name)
        if not os.path.isfile(path):
            raise RosterError(f"Файл штату '{configured_name}' (data.json → roster.file) не знайдено в resources/.")
        return path
    candidates = []
    for pattern in ROSTER_FILE_GLOBS:
        for path in glob.glob(os.path.join(resources_dir, pattern)):
            if not os.path.basename(path).startswith(ROSTER_EXCLUDED_PREFIXES):
                candidates.append(path)
    if not candidates:
        raise RosterError("У resources/ немає файлу штату (.xlsx) - покладіть його туди (без пароля).")
    return max(candidates, key=os.path.getmtime)


def _check_readable(path):
    with open(path, "rb") as f:
        if f.read(8) == _OLE2_SIGNATURE:
            raise RosterError(
                f"'{os.path.basename(path)}' захищено паролем (або це старий формат .xls) - відкрийте його в Excel, "
                f"зніміть пароль (Файл → Відомості → Захистити книгу → Зашифрувати паролем → очистити поле) "
                f"і збережіть як .xlsx."
            )


def _cell_text(value):
    return "" if value is None else clean_text(str(value))


def _matches(field_name, header, compact):
    if field_name == "vos" and _VOS_RE.search(header):
        return True
    return any(keyword in compact for keyword in ROSTER_HEADER_KEYWORDS[field_name])


def _distinct_values(rows, column):
    return len({_cell_text(row[column]) for row in rows if column < len(row) and _cell_text(row[column])})


def _choose_column(field_name, candidates, compact, data_rows):
    excluded = ROSTER_HEADER_EXCLUDE.get(field_name, ())
    preferred = [i for i in candidates if not any(word in compact[i] for word in excluded)] or candidates
    strong = [i for i in preferred if any(word in compact[i] for word in ROSTER_HEADER_PREFER.get(field_name, ()))]
    if strong:
        return strong[0]
    if field_name == "position" and len(preferred) > 1:
        # Кілька колонок посади (коротка, повна, …) - найповніша: найбільше різних значень.
        return max(preferred, key=lambda i: (_distinct_values(data_rows, i), -i))
    return preferred[0]


def detect_header(rows, scan_rows=ROSTER_HEADER_SCAN_ROWS):
    """rows - рядки аркуша (кортежі значень); заголовок шукається в перших
    scan_rows. → (індекс рядка заголовка, {поле: індекс колонки}) або None.
    Рядок заголовка має містити посаду і ПІБ."""
    for row_index, row in enumerate(rows[:scan_rows]):
        headers = [_cell_text(value).lower() for value in row]
        compact = [_COMPACT_RE.sub("", header) for header in headers]
        columns = {}
        for field_name in ROSTER_HEADER_KEYWORDS:
            candidates = [i for i, header in enumerate(headers) if header and _matches(field_name, header, compact[i])]
            if field_name == "position":
                candidates = [i for i in candidates if not _VOS_RE.search(headers[i])]
            if candidates:
                columns[field_name] = _choose_column(field_name, candidates, compact, rows[row_index + 1:])
            if field_name == "subunit":
                # Друга колонка "підрозділ" (не повна) - рота/батарея/взвод верхнього рівня.
                short = [i for i in candidates if i != columns.get("subunit") and "повн" not in compact[i]]
                if short:
                    columns["group"] = short[0]
        if "position" in columns and "pib" in columns:
            return row_index, columns
    return None


def _pick_sheet(workbook, sheet_name):
    if sheet_name:
        for ws in workbook.worksheets:
            if ws.title.strip().lower() == sheet_name.strip().lower():
                return ws
        raise RosterError(f"У штаті немає аркуша '{sheet_name}' (data.json → roster.sheet).")
    return None


def read_roster(resources_dir, config=None):
    """config - data.json → roster: file, sheet, include_if_contains."""
    config = config or {}
    path = find_roster_file(resources_dir, config.get("file", ""))
    _check_readable(path)
    with warnings.catch_warnings():
        # "Conditional Formatting/Data Validation extension is not supported" -
        # openpyxl лише попереджає, що не збереже їх; штат ми тільки читаємо.
        warnings.simplefilter("ignore", UserWarning)
        return _read_workbook(path, config)


def _read_workbook(path, config):
    try:
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise RosterError(f"Не вдалось відкрити '{os.path.basename(path)}': {e}") from e
    try:
        chosen = _pick_sheet(workbook, config.get("sheet", ""))
        sheets = [chosen] if chosen is not None else list(workbook.worksheets)
        for ws in sheets:
            # Розмір аркуша, записаний у файлі, буває неправильним - читаємо все.
            ws.reset_dimensions()
            rows = list(ws.iter_rows(values_only=True))
            header = detect_header(rows)
            if header is None:
                continue
            header_index, columns = header
            roster = Roster(path, ws.title, headers={
                name: _cell_text(rows[header_index][index]) for name, index in columns.items()
            })
            roster.slots = _read_slots(rows, header_index, columns, config.get("include_if_contains") or [])
            return roster
    finally:
        workbook.close()
    raise RosterError(
        f"У '{os.path.basename(path)}' не знайдено рядка заголовка з колонками «Посада» і «ПІБ»/«Прізвище»."
    )


def _read_slots(rows, header_index, columns, include_if_contains):
    filters = [clean_text(value).lower() for value in include_if_contains if clean_text(value)]
    has_vos = "vos" in columns

    def value(row, name):
        index = columns.get(name)
        return _cell_text(row[index]) if index is not None and index < len(row) else ""

    header_position = value(rows[header_index], "position").lower()
    slots = []
    for row_index in range(header_index + 1, len(rows)):
        row = rows[row_index]
        position = value(row, "position")
        if not position or position.lower() == header_position:
            continue
        slot = Slot(row_index + 1, position, value(row, "rank"), value(row, "pib"), value(row, "vos"),
                    value(row, "subunit"), value(row, "group"))
        # Рядки-розділювачі ("Управління батальйону" без ВОС і без людини) - не посада.
        if has_vos and not slot.vos and not slot.pib:
            continue
        if filters and not any(f in f"{slot.subunit} {slot.position}".lower() for f in filters):
            continue
        slots.append(slot)
    return slots
