"""Звірка книг resources/addons МІЖ СОБОЮ, без Табеля - за прямою вказівкою
користувача окреме питання на самому початку index.py.

Для кожної людини й кожної книги - ВЛАСНА шкала цієї книги (_paint_timeline
лише з її записів, ті самі правила, що й у звірці з Табелем). На датах, де
очікування є в ОБОХ книгах пари, статуси звіряються тими самими правилами
еквівалентності (_statuses_match в обидва боки: ППД=НОВ=БЗВП, 10 = статуси,
що ним оплачуються, "будь-який" тощо). День повернення/виписки (arrival_date
будь-якої з двох книг) не звіряється - книги записують цю межу по-різному.

За прямою вказівкою користувача - за ВСІ періоди, а не лише за місяць Табеля:
шкала людини - від найранішого запису в будь-якій книзі до last_date
(сьогодні чи пізніший відомий кінець періоду). Люди - з Табеля
(resources/ОБЛІК.xlsx), правило звірки з addons. Звіт - рядок на людину
(без дублювання ПІБ) з переліком розбіжних періодів, а не колонка на дату."""

import os
from datetime import date, timedelta
from itertools import combinations

from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from constants import PIB_COLUMN_NAME
from content.addon_mismatch_checker import _paint_timeline, _roster_pidrozdil, _sheet_title, _statuses_match
from content.addon_reader import ANY_STATUS
from content.oblik_timesheet import normalize_name
from content.payment_mismatch_checker import (
    _BORDER,
    _display_value,
    _FONT,
    _HEADER_ALIGNMENT,
    _HEADER_FILL,
    _HEADER_FONT,
    _MISMATCH_FILL,
)

_ANY_STATUS_LABEL = "будь-який"


def _agree(record_a, record_b, date_value):
    if date_value in (record_a.get("arrival_date"), record_b.get("arrival_date")):
        return True
    status_a, status_b = record_a["status"], record_b["status"]
    return _statuses_match(status_a, status_b) or _statuses_match(status_b, status_a)


def _last_date(period_records, today=None):
    """Кінець шкали - сьогодні або пізніший відомий кінець/початок періоду."""
    known = [r["end_date"] for r in period_records if r["end_date"]] + [r["start_date"] for r in period_records]
    return max([today or date.today()] + known)


def _days(start, end):
    return [(None, start + timedelta(days=offset)) for offset in range((end - start).days + 1)]


def _conflict_periods(conflict_days):
    """[(дата, статус A, статус B, аркуш A, аркуш B)] за зростанням дати ->
    періоди: поспіль ідучі дні з тією самою парою статусів і аркушів - один
    період {"start", "end", "status_a", "status_b", "sheet_a", "sheet_b"}."""
    periods = []
    for day, status_a, status_b, sheet_a, sheet_b in conflict_days:
        key = (status_a, status_b, sheet_a, sheet_b)
        last = periods[-1] if periods else None
        if last and last["key"] == key and last["end"] + timedelta(days=1) == day:
            last["end"] = day
        else:
            periods.append({"key": key, "start": day, "end": day})
    return [
        {"start": p["start"], "end": p["end"], "status_a": p["key"][0], "status_b": p["key"][1],
         "sheet_a": p["key"][2], "sheet_b": p["key"][3]}
        for p in periods
    ]


def find_addon_conflicts(timesheet, period_records, source_files, last_date=None):
    """Пари книг (у порядку source_files), де для людини з Табеля є хоча б
    одна дата з різними статусами, за ВСІ періоди. Повертає [{"files": (книга
    A, книга B), "roster_pib_raw", "roster_pidrozdil", "periods": [{"start",
    "end", "status_a", "status_b", "sheet_a", "sheet_b"}], "days": кількість
    розбіжних днів}]."""
    person_rows = timesheet.person_rows()
    pib_col = timesheet.label_columns[PIB_COLUMN_NAME]
    last_date = last_date or _last_date(period_records)

    by_person = {}
    for record in period_records:
        normalized = normalize_name(record["pib_raw"])
        if normalized in person_rows:
            by_person.setdefault(normalized, {}).setdefault(record["source_file"], []).append(record)

    conflicts = []
    for normalized, by_file in sorted(by_person.items()):
        present = [name for name in source_files if name in by_file]
        if len(present) < 2:
            continue
        first = min(r["start_date"] for name in present for r in by_file[name])
        days = _days(first, last_date)
        timelines = {name: _paint_timeline(by_file[name], days, last_date) for name in present}
        for file_a, file_b in combinations(present, 2):
            conflict_days = []
            for day in sorted(set(timelines[file_a]) & set(timelines[file_b])):
                record_a, record_b = timelines[file_a][day], timelines[file_b][day]
                if not _agree(record_a, record_b, day):
                    conflict_days.append((day, record_a["status"], record_b["status"],
                                          record_a["source_sheet"], record_b["source_sheet"]))
            if conflict_days:
                row_idx = person_rows[normalized]
                conflicts.append({
                    "files": (file_a, file_b),
                    "roster_pib_raw": timesheet.ws.cell(row=row_idx, column=pib_col).value,
                    "roster_pidrozdil": _roster_pidrozdil(timesheet, row_idx),
                    "periods": _conflict_periods(conflict_days),
                    "days": len(conflict_days),
                })
    return conflicts


def _status_display(status):
    return _ANY_STATUS_LABEL if status == ANY_STATUS else _display_value(status)


def _period_text(period):
    dates = period["start"].strftime("%d.%m.%Y")
    if period["end"] != period["start"]:
        dates += f"-{period['end'].strftime('%d.%m.%Y')}"
    return f"{dates}: {_status_display(period['status_a'])} / {_status_display(period['status_b'])}"


def _sheets(periods, key):
    """Аркуші книги, з яких узято розбіжні статуси, без повторів, у порядку появи."""
    return ", ".join(dict.fromkeys(period[key] for period in periods))


_COLUMN_WIDTHS = (22, 34, 30, 30, 60, 8)
_WRAP_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)


def _write_pair_sheet(ws, records, file_a, file_b):
    headers = ["Підрозділ", "ПІБ", f"{file_a}: аркуші", f"{file_b}: аркуші",
               f"Розбіжності (період: {file_a} / {file_b})", "Днів"]
    for col_idx, value in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=value)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGNMENT
        cell.border = _BORDER
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 45

    for row_idx, record in enumerate(sorted(records, key=lambda r: str(r["roster_pib_raw"])), start=2):
        values = [
            _display_value(record["roster_pidrozdil"]), _display_value(record["roster_pib_raw"]),
            _sheets(record["periods"], "sheet_a"), _sheets(record["periods"], "sheet_b"),
            "\n".join(_period_text(period) for period in record["periods"]), record["days"],
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = _FONT
            cell.alignment = _WRAP_TOP
            cell.border = _BORDER
        ws.cell(row=row_idx, column=5).fill = _MISMATCH_FILL
    for col_idx, width in enumerate(_COLUMN_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


_SHEET_TITLE_MAX_LENGTH = 31
_PAIR_SEPARATOR = " - "


def _pair_title(file_a, file_b):
    """"книга A - книга B"; задовга для аркуша Excel - кожна назва скорочується
    до половини місця, щоб у назві аркуша лишились ОБИДВІ книги."""
    title = f"{file_a}{_PAIR_SEPARATOR}{file_b}"
    if len(title) <= _SHEET_TITLE_MAX_LENGTH:
        return title
    half = (_SHEET_TITLE_MAX_LENGTH - len(_PAIR_SEPARATOR)) // 2
    return f"{file_a[:half].strip()}{_PAIR_SEPARATOR}{file_b[:half].strip()}"


def write_addon_conflicts_report(conflicts, output_path):
    """Аркуш на кожну пару книг з розбіжностями ("книга A - книга B"): рядок
    на людину - аркуші обох книг, розбіжні періоди (кожен з нового рядка
    клітинки: "дд.мм.рррр-дд.мм.рррр: статус A / статус B") і кількість
    розбіжних днів. None - немає розбіжностей."""
    if not conflicts:
        return None
    by_pair = {}
    for record in conflicts:
        by_pair.setdefault(record["files"], []).append(record)

    wb = Workbook()
    wb.remove(wb.active)
    used_titles = set()
    for (file_a, file_b), records in by_pair.items():
        _write_pair_sheet(wb.create_sheet(_sheet_title(_pair_title(file_a, file_b), used_titles)), records, file_a, file_b)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    try:
        wb.save(output_path)
    except PermissionError as e:
        raise PermissionError(
            f"Не вдалось зберегти '{output_path}': файл зараз відкритий в іншій програмі "
            "(напр. Excel). Закрийте його та запустіть звірку знову."
        ) from e
    return output_path
