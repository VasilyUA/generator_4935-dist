import re
import zipfile
from datetime import date, datetime

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.dimensions import ColumnDimension

import content.oblik_timesheet as ot

_DATES = [datetime(2026, 10, day) for day in range(1, 5)]
_HEADER = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", *_DATES, "ПІДСТАВИ ДЛЯ НЕ ВИПЛАТИ", "ПІДСТАВИ ДТЛ"]
_DATE_COLUMN = 5
_LAST_DATE_COLUMN = 8
_BASIS_COLUMN = 9
_FIRST = ["1 абв", "Посада перша", "сержант", "ПЕРШИЙ Перший Перший", 100, "100_БЗ", None, "ВП", "підстава", None]
_SECOND = ["2 абв", "Посада друга", "матрос", "ДРУГИЙ Другий Другий", "РВЗ", "10_МЕД", 30, "Навч", None, None]


def _workbook(tmp_path, rows=(_FIRST, _SECOND), header=_HEADER, prepare=None):
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.title = "Табель"
    ws.append(header)
    for row in rows:
        ws.append(list(row))
    if prepare:
        prepare(ws)
    path = tmp_path / "ОБЛІК.xlsx"
    wb.save(str(path))
    return str(path)


def _saved(tmp_path, **options):
    """Табель, збережений через Timesheet.save, прочитаний з файлу."""
    output = str(tmp_path / "output.xlsx")
    ot.Timesheet(_workbook(tmp_path, **options), "Табель").save(output)
    return load_workbook(output)["Табель"]


def _sheet_xml(path):
    with zipfile.ZipFile(path) as book:
        return book.read("xl/worksheets/sheet1.xml").decode("utf-8")


def test_print_area_stops_at_the_last_date(tmp_path):
    """Рішення користувача: колонки підстав на друк не йдуть."""
    ws = _saved(tmp_path)

    assert ws.print_area == "'Табель'!$A$1:$H$3"


def test_header_repeats_on_every_printed_page(tmp_path):
    assert _saved(tmp_path).print_title_rows == "$1:$1"


def test_table_prints_on_a4_landscape_one_page_wide(tmp_path):
    ws = _saved(tmp_path)

    assert ws.page_setup.orientation == "landscape"
    assert str(ws.page_setup.paperSize) == ws.PAPERSIZE_A4
    assert ws.sheet_properties.pageSetUpPr is not None and ws.sheet_properties.pageSetUpPr.fitToPage is True
    assert (ws.page_setup.fitToWidth, ws.page_setup.fitToHeight) == (1, 0)


def test_printed_pages_are_numbered(tmp_path):
    footer = _saved(tmp_path).oddFooter

    assert footer is not None and footer.center.text == "Сторінка &P з &N"


def test_names_and_header_stay_visible_while_scrolling(tmp_path):
    assert _saved(tmp_path).freeze_panes == "E2"


def test_opens_at_the_top_of_the_table(tmp_path):
    """Реальний випадок: файл відкривався прокрученим до рядка 212."""
    def scrolled(ws):
        ws.freeze_panes = "E212"

    pane = _saved(tmp_path, prepare=scrolled).sheet_view.pane

    assert pane is not None and pane.topLeftCell == "E2"


def test_zoom_is_not_the_old_70_percent(tmp_path):
    ws = _saved(tmp_path, prepare=lambda sheet: setattr(sheet.sheet_view, "zoomScale", 70))

    assert ws.sheet_view.zoomScale == 90


def test_all_values_get_the_same_readable_font(tmp_path):
    ws = _saved(tmp_path)

    assert {(cell.font.name, cell.font.sz) for cell in ws[2] if cell.value in ("сержант", 100, "ПЕРШИЙ Перший Перший")} == {("Times New Roman", 13)}


def test_header_keeps_bold_and_its_color(tmp_path):
    def styled(ws):
        ws["D1"].font = Font(name="Calibri", sz=16, b=True, color="FFFFFFFF")

    cell = _saved(tmp_path, prepare=styled)["D1"]

    assert (cell.font.name, cell.font.sz, cell.font.b, cell.font.color.rgb) == ("Times New Roman", 13, True, "FFFFFFFF")


def test_date_headers_stay_dates_after_saving(tmp_path):
    """Заголовок дати - клітинка-дата: за ним Timesheet знаходить колонки."""
    output = str(tmp_path / "output.xlsx")
    ot.Timesheet(_workbook(tmp_path), "Табель").save(output)

    assert [day for _col, day in ot.Timesheet(output, "Табель").date_columns] == [date(2026, 10, day) for day in range(1, 5)]


def test_date_headers_use_the_format_every_device_reads(tmp_path):
    assert _saved(tmp_path).cell(row=1, column=_DATE_COLUMN).number_format == "dd.MM.yyyy"


@pytest.mark.parametrize("value", [100, 30, "ВП", "СЗЧ", "170"])
def test_short_value_stays_horizontal(tmp_path, value):
    cell = _saved(tmp_path, rows=[[*_FIRST[:4], value, None, None, None, None, None]]).cell(row=2, column=_DATE_COLUMN)

    assert not cell.alignment.textRotation
    assert cell.alignment.shrink_to_fit is True
    assert cell.font.sz == 13


@pytest.mark.parametrize("value", ["100_БЗ", "100_БПШП", "10_МЕД", "Смерть", "ВПСЗ", "Навч"])
def test_long_value_is_rotated_so_it_stays_readable(tmp_path, value):
    """Стиснення Excel робило "100_БПШП" меншим за 3 пт на папері."""
    cell = _saved(tmp_path, rows=[[*_FIRST[:4], value, None, None, None, None, None]]).cell(row=2, column=_DATE_COLUMN)

    assert cell.alignment.textRotation == 90
    assert cell.font.sz == 10


def test_rotated_value_may_shrink_so_it_never_sticks_out_of_the_row(tmp_path):
    """Висоту рядка Excel підбирає за екраном, а на папері текст трохи довший:
    без стиснення кінець "10_МЕД" обрізався у 2724 клітинках із 4271."""
    cell = _saved(tmp_path).cell(row=2, column=_DATE_COLUMN + 1)

    assert cell.alignment.textRotation == 90
    assert cell.alignment.shrink_to_fit is True


def test_date_and_subdivision_headers_are_rotated(tmp_path):
    ws = _saved(tmp_path)

    assert [ws.cell(row=1, column=col).alignment.textRotation for col in (1, _DATE_COLUMN, _LAST_DATE_COLUMN)] == [90, 90, 90]


def test_other_headers_stay_horizontal_and_wrap(tmp_path):
    ws = _saved(tmp_path)

    for col in (2, 3, 4, _BASIS_COLUMN):
        assert not ws.cell(row=1, column=col).alignment.textRotation
        assert ws.cell(row=1, column=col).alignment.wrap_text is True


def test_empty_date_cell_is_not_rotated(tmp_path):
    assert not _saved(tmp_path).cell(row=2, column=_DATE_COLUMN + 2).alignment.textRotation


def test_names_and_positions_are_left_aligned(tmp_path):
    ws = _saved(tmp_path)

    assert [ws.cell(row=2, column=col).alignment.horizontal for col in (1, 2, 3, 4)] == ["center", "left", "center", "left"]


def test_long_positions_and_names_wrap(tmp_path):
    ws = _saved(tmp_path)

    assert all(ws.cell(row=2, column=col).alignment.wrap_text for col in (2, 4))


def test_every_printed_cell_has_a_border(tmp_path):
    ws = _saved(tmp_path)

    assert all(ws.cell(row=2, column=col).border.left.style == "thin" for col in range(1, _LAST_DATE_COLUMN + 1))


def test_row_heights_are_left_for_excel_to_fit(tmp_path):
    """Реальний випадок: збережена висота обрізала довгі посади, а наша оцінка
    тексту була неточна. Рядок без висоти Excel підбирає при відкритті."""
    def fixed(ws):
        ws.row_dimensions[2].height = 25.5

    output = str(tmp_path / "output.xlsx")
    ot.Timesheet(_workbook(tmp_path, prepare=fixed), "Табель").save(output)

    assert load_workbook(output)["Табель"].row_dimensions[2].height is None
    assert not re.search(r'<row r="2"[^>]*\bht=', _sheet_xml(output))


def test_header_keeps_a_height_for_rotated_dates(tmp_path):
    assert _saved(tmp_path).row_dimensions[1].height == 92


def test_long_basis_text_does_not_stretch_the_rows(tmp_path):
    """Підстав немає на друці, а з перенесенням рядок підлаштовувався б під них."""
    def wrapped(ws):
        ws.cell(row=2, column=_BASIS_COLUMN).alignment = Alignment(wrap_text=True)

    assert not _saved(tmp_path, prepare=wrapped).cell(row=2, column=_BASIS_COLUMN).alignment.wrap_text


def test_values_stay_exactly_as_they_were(tmp_path):
    ws = _saved(tmp_path)

    assert [ws.cell(row=2, column=col).value for col in range(1, 11)] == _FIRST
    assert [ws.cell(row=3, column=col).value for col in range(1, 11)] == _SECOND


def test_status_colors_stay(tmp_path):
    def painted(ws):
        ws["E2"].fill = PatternFill(start_color="FF92D050", end_color="FF92D050", fill_type="solid")

    assert _saved(tmp_path, prepare=painted)["E2"].fill.start_color.rgb == "FF92D050"


def test_label_and_date_columns_get_their_widths(tmp_path):
    ws = _saved(tmp_path)

    assert [ws.column_dimensions[letter].width for letter in "ABCD"] == [9.5, 26, 11.5, 31]
    assert {ws.column_dimensions[letter].width for letter in "EFGH"} == {4.6}


def _grouped_date_columns(ws):
    ws.column_dimensions["E"] = ColumnDimension(ws, index="E", min=5, max=8, width=11.3, customWidth=True)
    ws.column_dimensions["J"] = ColumnDimension(ws, index="J", min=10, max=16384, width=9, customWidth=True)


def test_grouped_columns_do_not_overlap(tmp_path):
    """Excel зберігає однакові колонки одним записом E..AJ: якби він лишився поруч
    з новими ширинами, записи перекрились би, а це пошкоджений файл."""
    output = str(tmp_path / "output.xlsx")
    ot.Timesheet(_workbook(tmp_path, prepare=_grouped_date_columns), "Табель").save(output)

    ranges = sorted((int(first), int(last)) for first, last in re.findall(r'<col [^>]*?min="(\d+)" max="(\d+)"', _sheet_xml(output)))
    ranges = ranges or sorted((int(first), int(last)) for last, first in re.findall(r'<col [^>]*?max="(\d+)"[^>]*?min="(\d+)"', _sheet_xml(output)))
    assert ranges
    assert all(earlier[1] < later[0] for earlier, later in zip(ranges, ranges[1:]))


def test_every_grouped_date_column_gets_the_new_width(tmp_path):
    ws = _saved(tmp_path, prepare=_grouped_date_columns)

    assert all(ws.column_dimensions[letter].width == 4.6 for letter in "EFGH")


def test_columns_after_the_dates_keep_their_width(tmp_path):
    def wide_basis(ws):
        ws.column_dimensions["I"].width = 25.4

    assert _saved(tmp_path, prepare=wide_basis).column_dimensions["I"].width == 25.4


def test_hidden_column_of_a_group_stays_hidden(tmp_path):
    def hidden(ws):
        ws.column_dimensions["B"] = ColumnDimension(ws, index="B", min=2, max=3, width=20, hidden=True, customWidth=True)

    ws = _saved(tmp_path, prepare=hidden)

    assert ws.column_dimensions["B"].hidden and ws.column_dimensions["C"].hidden


def test_table_without_a_subdivision_column_is_still_laid_out(tmp_path):
    ws = _saved(tmp_path, header=["ПОСАДА", "ЗВАННЯ", "ПІБ", *_DATES], rows=[["Посада", "сержант", "ПЕРШИЙ Перший Перший", 100]])

    assert ws.print_area == "'Табель'!$A$1:$G$2"
    assert ws.freeze_panes == "D2"


def test_saving_twice_gives_a_valid_view(tmp_path):
    """Закріплення додає службові записи виділення: повторне збереження не має їх множити."""
    output = str(tmp_path / "output.xlsx")
    sheet = ot.Timesheet(_workbook(tmp_path), "Табель")
    sheet.save(output)
    sheet.save(output)

    panes = re.findall(r'<selection [^>]*pane="(\w+)"', _sheet_xml(output))
    assert len(panes) == len(set(panes))


def test_other_sheets_are_not_touched(tmp_path):
    def with_another_sheet(ws):
        other = ws.parent.create_sheet("ТВО")
        other["A1"] = "значення"
        other["A1"].font = Font(name="Calibri", sz=11)

    output = str(tmp_path / "output.xlsx")
    ot.Timesheet(_workbook(tmp_path, prepare=with_another_sheet), "Табель").save(output)

    cell = load_workbook(output)["ТВО"]["A1"]
    assert (cell.value, cell.font.name, cell.font.sz) == ("значення", "Calibri", 11)


def test_export_for_money_project_keeps_the_print_area_on_the_renamed_sheet(tmp_path):
    """Аркуш там перейменовується: область друку зі старою назвою зламала б файл."""
    source = str(tmp_path / "output.xlsx")
    ot.Timesheet(_workbook(tmp_path), "Табель").save(source)
    target = str(tmp_path / "експорт.xlsx")

    ot.export_for_money_project(source, target, "ЖОВТЕНЬ")

    ws = load_workbook(target)["ЖОВТЕНЬ"]
    assert ws.print_area == "'ЖОВТЕНЬ'!$A$1:$H$3"
    assert ws.print_title_rows == "$1:$1"
