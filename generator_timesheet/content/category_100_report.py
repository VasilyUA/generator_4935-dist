"""Таблиця порушень правила "категорія виплати 100 поспіль"
(Перевірка_категорії_100.xlsx): хто, які завдання за аркушем БЧС файлів
information_unit і яка саме серія днів. Стилі - ті самі, що в
error_mis_statuses.xlsx (content/payment_mismatch_checker.py)."""

import math

from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from content.excel_compat import save_workbook, workbook_with_sheet
from content.oblik_timesheet import normalize_name, solid_fill
from content.payment_mismatch_checker import _BORDER, _CENTER, _display_value, _FONT, _HEADER_ALIGNMENT, _HEADER_FILL, _HEADER_FONT
from content.report_pib_column import PIB_ALIGNMENT, pib_column_width

_SHEET_TITLE = "Категорія 100"
_HEADERS = [
    "Підрозділ", "ПІБ",
    "Завдання яке виконує військовослужбовець (ФАКТИЧНО)",
    "Бойове завдання яке виконує військовослужбовець (ПО ДОКУМЕНТАХ НА ВИПЛАТИ)",
    "Період", "Днів поспіль", "Дозволено",
]
_PIB_COLUMN = 2
_TASK_COLUMNS = (3, 4)
_COLUMN_WIDTHS = {1: 20, 3: 60, 4: 60, 5: 28, 6: 12, 7: 12}
_NOTE = (
    "Якщо так має бути - вкажіть людині кількість днів в EXCLUDE_NAMES (resources/data.json). "
    "Люди з кольором в EXCLUDE_NAMES є в таблиці навіть без порушення."
)

_TASK_ALIGNMENT = Alignment(horizontal="left", vertical="center", wrap_text=True)
_HEADER_ROW_HEIGHT = 60
_LINE_HEIGHT = 20
# Скільки символів Times New Roman 14 вміщається в рядок колонки завдань: Excel
# не підбирає висоту сам, якщо рядок створено не в ньому.
_TASK_LINE_LENGTH = 44


def _row_height(tasks):
    lines = max(math.ceil(len(task) / _TASK_LINE_LENGTH) for task in tasks)
    return max(1, lines) * _LINE_HEIGHT


def _period(record):
    if record["start"] is None:
        return _display_value(None)
    return f"{record['start']:%d.%m.%Y} - {record['end']:%d.%m.%Y}"


def write_category_100_report(records, tasks, output_path):
    """Будує output_path: рядок на кожен запис (check_payment_category_100_days і
    category_100_marked_rows) у тому самому порядку. tasks - read_bchs_tasks;
    людина без завдань (напр. у файлі підрозділу немає аркуша БЧС) отримує "—".
    Клітинка з ПІБ заливається кольором людини, якщо він заданий. Під таблицею -
    підказка про винятки. Повертає output_path."""
    wb, ws = workbook_with_sheet(_SHEET_TITLE)
    for col_idx, text in enumerate(_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=text)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGNMENT
        cell.border = _BORDER
    ws.row_dimensions[1].height = _HEADER_ROW_HEIGHT
    ws.freeze_panes = "A2"

    for row, record in enumerate(records, start=2):
        person = tasks.get(normalize_name(record["pib_raw"]), {})
        actual, documents = _display_value(person.get("actual")), _display_value(person.get("documents"))
        values = (
            _display_value(record["pidrozdil"]), record["pib_raw"], actual, documents,
            _period(record), record["days"], record["limit"],
        )
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.font = _FONT
            cell.border = _BORDER
            cell.alignment = (PIB_ALIGNMENT if col_idx == _PIB_COLUMN
                              else _TASK_ALIGNMENT if col_idx in _TASK_COLUMNS else _CENTER)
        if record["color"]:
            ws.cell(row=row, column=_PIB_COLUMN).fill = solid_fill(record["color"])
        ws.row_dimensions[row].height = _row_height((actual, documents))

    note_row = len(records) + 3
    note = ws.cell(row=note_row, column=1, value=_NOTE)
    note.font = _FONT
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=len(_HEADERS))

    widths = {**_COLUMN_WIDTHS, _PIB_COLUMN: pib_column_width([record["pib_raw"] for record in records])}
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    save_workbook(wb, output_path)
    return output_path
