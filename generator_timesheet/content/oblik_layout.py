"""Вигляд аркуша "Табель" у готовому ОБЛІК: щоб таблицю було добре видно на
екрані й на аркуші A4. Значення клітинок і заливки статусів не змінюються.

На друк іде лише таблиця до останньої дати: колонки підстав справа від неї в
область друку не входять (рішення користувача)."""

from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.views import Selection

from constants import LABEL_COLUMN_NAMES, PIB_COLUMN_NAME

_FONT_NAME = "Times New Roman"
_FONT_SIZE = 13
_LABEL_WIDTHS = {"ПІДРОЗДІЛ": 9.5, "ПОСАДА": 26, "ЗВАННЯ": 11.5, PIB_COLUMN_NAME: 31}
# У клітинках дат майже завжди 2-3 символи, тож від ширини цих колонок залежить
# масштаб друку, а отже й шрифт на папері. Довше значення ("100_БЗ") повертається:
# стиснення Excel робило "100_БПШП" меншим за 3 пт, а перенесення рвало слова.
_DATE_COLUMN_WIDTH = 4.6
_LONGEST_HORIZONTAL_VALUE = 3
_ROTATED_FONT_SIZE = 10
_HEADER_HEIGHT = 92
_ZOOM = 90

_LEFT_COLUMNS = ("ПОСАДА", PIB_COLUMN_NAME)
_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ROTATED_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="top", text_rotation=90)
_LEFT_ALIGNMENT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
_CENTER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
_DATE_ALIGNMENT = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
# Висоту рядка Excel підбирає за екраном, а на папері текст трохи довший: без
# стиснення кінець повернутого значення обрізався.
_ROTATED_DATE_ALIGNMENT = Alignment(horizontal="center", vertical="center", text_rotation=90, shrink_to_fit=True)
# Без перенесення: висоту рядків підбирає сам Excel, і довгі тексти підстав, яких
# немає на друці, робили б рядки величезними.
_NOT_PRINTED_ALIGNMENT = Alignment(horizontal="left", vertical="center")
_THIN = Side(style="thin", color="FF000000")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _is_rotated(value):
    return value is not None and len(str(value)) > _LONGEST_HORIZONTAL_VALUE


def _split_grouped_columns(ws, last_column):
    """Excel зберігає однакові сусідні колонки одним записом (E..AJ). Такий запис
    розбивається на окремі, бо інакше він накладався б на нові ширини, а це
    Excel вважає пошкодженням файлу."""
    for dimension in list(ws.column_dimensions.values()):
        first, last = dimension.min, dimension.max
        if first and last and first < last and first <= last_column:
            dimension.max = first
            for col_idx in range(first + 1, min(last, ws.max_column) + 1):
                column = ws.column_dimensions[get_column_letter(col_idx)]
                column.width, column.hidden = dimension.width, dimension.hidden


def _resized_font(cell, size=_FONT_SIZE):
    font = cell.font
    return Font(name=_FONT_NAME, sz=size, b=font.b, i=font.i, u=font.u, strike=font.strike, color=font.color)


def _style_header(ws, date_columns):
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = _resized_font(cell)
        cell.border = _BORDER
        cell.alignment = _ROTATED_HEADER_ALIGNMENT if col_idx in date_columns or col_idx == 1 else _HEADER_ALIGNMENT
        if col_idx in date_columns:
            cell.number_format = "dd.MM.yyyy"
    ws.row_dimensions[1].height = _HEADER_HEIGHT


def _style_date_cell(cell):
    rotated = _is_rotated(cell.value)
    cell.font = _resized_font(cell, _ROTATED_FONT_SIZE if rotated else _FONT_SIZE)
    cell.alignment = _ROTATED_DATE_ALIGNMENT if rotated else _DATE_ALIGNMENT
    cell.border = _BORDER


def _style_rows(ws, date_columns, label_columns):
    label_by_column = {col_idx: name for name, col_idx in label_columns.items()}
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if col_idx in date_columns:
                _style_date_cell(cell)
                continue
            cell.font = _resized_font(cell)
            if col_idx in label_by_column:
                cell.alignment = _LEFT_ALIGNMENT if label_by_column[col_idx] in _LEFT_COLUMNS else _CENTER_ALIGNMENT
                cell.border = _BORDER
            else:
                cell.alignment = _NOT_PRINTED_ALIGNMENT
        # Рядок без висоти Excel підбирає сам при відкритті, точно за текстом.
        ws.row_dimensions[row_idx].height = None


def _set_up_view(ws, first_date_column):
    view = ws.sheet_view
    view.zoomScale = _ZOOM
    view.zoomScaleNormal = _ZOOM
    view.selection = [Selection(activeCell="A1", sqref="A1")]
    ws.freeze_panes = ws.cell(row=2, column=first_date_column)


def _set_up_printing(ws, last_date_column):
    ws.print_area = f"A1:{get_column_letter(last_date_column)}{ws.max_row}"
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.3, bottom=0.45, header=0.2, footer=0.2)
    ws.print_options.horizontalCentered = True
    ws.oddFooter.center.text = "Сторінка &P з &N"


def apply_oblik_layout(timesheet):
    """Оформлює аркуш timesheet.ws для роботи й друку: шрифт, ширини колонок,
    закріплення заголовка, область і параметри друку."""
    ws = timesheet.ws
    date_columns = {col_idx for col_idx, _day in timesheet.date_columns}
    label_columns = {name: timesheet.label_columns[name] for name in LABEL_COLUMN_NAMES if name in timesheet.label_columns}

    _split_grouped_columns(ws, max(date_columns))
    for name, col_idx in label_columns.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = _LABEL_WIDTHS[name]
    for col_idx in date_columns:
        ws.column_dimensions[get_column_letter(col_idx)].width = _DATE_COLUMN_WIDTH

    _style_header(ws, date_columns)
    _style_rows(ws, date_columns, label_columns)
    _set_up_view(ws, min(date_columns))
    _set_up_printing(ws, max(date_columns))
