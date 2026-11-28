import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import pytest
from openpyxl import Workbook, load_workbook

from checker_accounting import checker


def _write_xlsx(path, headers, rows, sheet_name="Аркуш1"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)
    return path


D1 = datetime(2026, 7, 1)
D2 = datetime(2026, 7, 2)
D3 = datetime(2026, 7, 3)


# -------------------------
# _find_check_files
# -------------------------
def test_find_check_files_raises_when_dir_empty(tmp_path):
    with pytest.raises(FileNotFoundError):
        checker._find_check_files(str(tmp_path))


def test_find_check_files_returns_sorted_xlsx_and_xlsm_paths(tmp_path):
    """.xlsm (макро-файл) - той самий ZIP+XML формат, що й .xlsx, і так само
    читається openpyxl.load_workbook - має підхоплюватись поряд з .xlsx."""
    _write_xlsx(tmp_path / "ОБЛІК2.xlsx", ["ПІБ"], [["Перший Перший"]])
    _write_xlsx(tmp_path / "ОБЛІК1.xlsx", ["ПІБ"], [["Перший Перший"]])
    _write_xlsx(tmp_path / "ОБЛІК3.xlsm", ["ПІБ"], [["Перший Перший"]])
    (tmp_path / "notes.txt").write_text("не .xlsx/.xlsm - має бути проігнороване", encoding="utf-8")

    result = checker._find_check_files(str(tmp_path))
    assert [Path(p).name for p in result] == ["ОБЛІК1.xlsx", "ОБЛІК2.xlsx", "ОБЛІК3.xlsm"]


# -------------------------
# _read_check_file
# -------------------------
def test_read_check_file_detects_columns_by_name(tmp_path):
    """Колонки ПІДРОЗДІЛ/ПОСАДА/ЗВАННЯ/ПІБ і дати визначаються за назвою/типом
    заголовка, а не позицією - працює навіть із зайвою колонкою на початку і
    порожніми "технічними" колонками між іменованими й датованими."""
    path = _write_xlsx(
        tmp_path / "ОБЛІК.xlsx",
        [None, "ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", None, None, D1, D2],
        [["оф", "Підрозділ 1", "Стрілець", "сержант", "Перший Перший", None, None, 100, 30]],
    )

    people, date_columns = checker._read_check_file(str(path))

    assert date_columns == {D1.date(), D2.date()}
    assert len(people) == 1
    person = people[0]
    assert person["pib_raw"] == "Перший Перший"
    assert person["label"] == {"ПІДРОЗДІЛ": "Підрозділ 1", "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Перший Перший"}
    assert person["days"] == {D1.date(): 100, D2.date(): 30}


def test_read_check_file_raises_without_pib_column(tmp_path):
    path = _write_xlsx(tmp_path / "ОБЛІК.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА"], [["Підрозділ 1", "Стрілець"]])

    with pytest.raises(ValueError, match="ПІБ"):
        checker._read_check_file(str(path))


def test_read_check_file_skips_rows_with_empty_pib(tmp_path):
    path = _write_xlsx(
        tmp_path / "ОБЛІК.xlsx",
        ["ПІБ", D1],
        [["Перший Перший", 100], [None, 30], ["", 30]],
    )

    people, _ = checker._read_check_file(str(path))

    assert [p["pib_raw"] for p in people] == ["Перший Перший"]


# -------------------------
# _extract_style_templates
# -------------------------
def test_extract_style_templates_raises_without_label_or_date_column(tmp_path):
    path = _write_xlsx(tmp_path / "ОБЛІК.xlsx", ["X", "Y"], [["a", "b"]])

    with pytest.raises(ValueError):
        checker._extract_style_templates(str(path))


def test_extract_style_templates_reads_widths_and_heights(tmp_path):
    path = _write_xlsx(tmp_path / "ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 100]])

    styles = checker._extract_style_templates(str(path))

    assert styles["header_label"].value == "ПІБ"
    assert styles["header_date"].value == D1
    assert styles["body_label"].value == "Перший Перший"
    assert styles["body_date"].value == 100


# -------------------------
# _normalize_name / _normalize_for_compare / _values_match
# -------------------------
def test_normalize_name_strips_spaces_and_upper_cases():
    assert checker._normalize_name("  Перший   Перший  ") == "ПЕРШИЙ ПЕРШИЙ"


def test_normalize_name_non_string_returns_empty():
    assert checker._normalize_name(None) == ""


def test_normalize_for_compare_treats_none_as_empty_string():
    assert checker._normalize_for_compare(None) == ""


def test_normalize_for_compare_strips_and_uppercases_strings():
    assert checker._normalize_for_compare(" вп ") == "ВП"


def test_normalize_for_compare_leaves_numbers_untouched():
    assert checker._normalize_for_compare(100) == 100


def test_values_match_true_for_identical_values():
    assert checker._values_match([100, 100, 100]) is True


def test_values_match_true_treats_none_and_blank_string_as_equal():
    assert checker._values_match([None, ""]) is True


def test_values_match_false_for_different_values():
    assert checker._values_match([100, 30]) is False


# -------------------------
# _format_day_value / _row_day_value_summary - колонка "Підсумок"
# -------------------------
def test_format_day_value_integer_float_drops_trailing_zero():
    assert checker._format_day_value(100.0) == "100"


def test_format_day_value_non_integer_float_kept_as_is():
    assert checker._format_day_value(100.5) == "100.5"


def test_format_day_value_string_normalized_case_and_spacing():
    assert checker._format_day_value("  вп  ") == "ВП"


def test_format_day_value_int_passthrough():
    assert checker._format_day_value(30) == "30"


def test_row_day_value_summary_counts_each_distinct_value():
    """ДИНАМІЧНИЙ перелік - будь-яке значення, що трапляється (не лише
    100/30/70/170), отримує свій запис - підтверджено користувачем."""
    day_values = [100, 100, 30, 30, 30, 70, 70, 170]

    assert checker._row_day_value_summary(day_values) == "100 - 2 днів; 30 - 3 днів; 70 - 2 днів; 170 - 1 днів"


def test_row_day_value_summary_includes_text_statuses():
    day_values = ["ВП", "ВП", 100, "ВД"]

    assert checker._row_day_value_summary(day_values) == "ВП - 2 днів; 100 - 1 днів; ВД - 1 днів"


def test_row_day_value_summary_skips_empty_cells():
    day_values = [100, None, "", 100]

    assert checker._row_day_value_summary(day_values) == "100 - 2 днів"


def test_row_day_value_summary_merges_case_variants_of_same_status():
    day_values = ["Вп", "ВП", "вп"]

    assert checker._row_day_value_summary(day_values) == "ВП - 3 днів"


def test_row_day_value_summary_order_follows_first_appearance():
    day_values = [30, 100, 30, 70]

    assert checker._row_day_value_summary(day_values) == "30 - 2 днів; 100 - 1 днів; 70 - 1 днів"


def test_row_day_value_summary_empty_when_all_cells_blank():
    assert checker._row_day_value_summary([None, "", None]) == ""


# -------------------------
# check_accounting - наскрізні сценарії
# -------------------------
def test_check_accounting_returns_none_when_no_files(tmp_path, capsys):
    result = checker.check_accounting(check_dir=str(tmp_path), output_path=str(tmp_path / "out.xlsx"))

    assert result is None
    assert not (tmp_path / "out.xlsx").exists()


def test_check_accounting_merges_pib_and_colors_matches_green(tmp_path):
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1, D2], [
        ["Підрозділ 1", "Стрілець", "сержант", "Перший Перший", 100, 100],
    ])
    _write_xlsx(check_dir / "ОБЛІК2.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1, D2], [
        ["Підрозділ 1", "Стрілець", "сержант", "Перший Перший", 100, 100],
    ])
    output_path = tmp_path / "RESULT_accounting.xlsx"

    result = checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    assert result == str(output_path)
    assert output_path.exists()

    wb = load_workbook(str(output_path))
    ws = wb.active
    assert [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))] == ["Файл", "ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1, D2, "Підсумок"]
    assert ws.cell(row=2, column=1).value == "ОБЛІК1.xlsx"
    assert ws.cell(row=3, column=1).value == "ОБЛІК2.xlsx"
    assert ws.cell(row=2, column=2).value == "Підрозділ 1"
    assert ws.cell(row=2, column=3).value == "Стрілець"
    assert ws.cell(row=2, column=4).value == "сержант"
    assert ws.cell(row=2, column=5).value == "Перший Перший"
    # ПІДРОЗДІЛ/ПОСАДА/ЗВАННЯ/ПІБ - усі об'єднані на 2 рядки цієї людини, як і ПІБ.
    merged_ranges = {str(r) for r in ws.merged_cells.ranges}
    assert merged_ranges == {"B2:B3", "C2:C3", "D2:D3", "E2:E3"}

    # ПІДРОЗДІЛ/ПОСАДА/ЗВАННЯ - об'єднані клітинки: стиль зберігається лише на
    # "якірному" (верхньому) рядку - саме так Excel показує колір усього діапазону.
    for col in (2, 3, 4):
        assert ws.cell(row=2, column=col).fill.fgColor.rgb == checker.GREEN_FILL.fgColor.rgb
    # Дати - НЕ об'єднані: колір застосовано до кожного рядка окремо.
    for col in (6, 7):
        assert ws.cell(row=2, column=col).fill.fgColor.rgb == checker.GREEN_FILL.fgColor.rgb
        assert ws.cell(row=3, column=col).fill.fgColor.rgb == checker.GREEN_FILL.fgColor.rgb


def test_check_accounting_merges_and_colors_mismatched_label_columns_red(tmp_path):
    """ПІДРОЗДІЛ/ПОСАДА/ЗВАННЯ теж об'єднуються на N рядків людини (як ПІБ), але
    порівняння й забарвлення - за ФАКТИЧНИМИ значеннями кожного файлу: розбіжність
    (тут - ПОСАДА) не зникає через те, що об'єднана клітинка показує лише канонічне
    (з першого файлу) значення."""
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1], [
        ["Підрозділ 1", "Стрілець", "сержант", "Перший Перший", 100],
    ])
    _write_xlsx(check_dir / "ОБЛІК2.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1], [
        ["Підрозділ 1", "Навідник", "сержант", "Перший Перший", 100],
    ])
    output_path = tmp_path / "RESULT_accounting.xlsx"

    checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb.active
    assert ws.cell(row=2, column=2).fill.fgColor.rgb == checker.GREEN_FILL.fgColor.rgb  # ПІДРОЗДІЛ - збігається
    assert ws.cell(row=2, column=3).fill.fgColor.rgb == checker.RED_FILL.fgColor.rgb    # ПОСАДА - розбіжність
    assert ws.cell(row=2, column=4).fill.fgColor.rgb == checker.GREEN_FILL.fgColor.rgb  # ЗВАННЯ - збігається
    # Канонічне (з файлу 1) значення все одно показане в об'єднаній клітинці.
    assert ws.cell(row=2, column=3).value == "Стрілець"


def test_check_accounting_colors_mismatch_red(tmp_path):
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ПІБ", D1], [["Перший Перший", 100]])
    _write_xlsx(check_dir / "ОБЛІК2.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])
    output_path = tmp_path / "RESULT_accounting.xlsx"

    checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb.active
    date_col = 6  # Файл, ПІДРОЗДІЛ, ПОСАДА, ЗВАННЯ, ПІБ, дата
    assert ws.cell(row=2, column=date_col).fill.fgColor.rgb == checker.RED_FILL.fgColor.rgb
    assert ws.cell(row=3, column=date_col).fill.fgColor.rgb == checker.RED_FILL.fgColor.rgb


def test_check_accounting_summary_column_reflects_each_row_own_days(tmp_path):
    """Колонка "Підсумок" (ОСТАННЯ) - ОКРЕМИЙ підсумок для КОЖНОГО рядка (файлу),
    за ЙОГО ВЛАСНИМИ днями - підтверджено користувачем: "в кінці таблички
    після кінця місяця" + динамічний перелік значень (не лише 100/30/70/170)."""
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ПІБ", D1, D2, D3], [["Перший Перший", 100, 100, "ВП"]])
    output_path = tmp_path / "RESULT_accounting.xlsx"

    checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb.active
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    assert header[-1] == "Підсумок"
    summary_col = len(header)
    assert ws.cell(row=2, column=summary_col).value == "100 - 2 днів; ВП - 1 днів"


def test_check_accounting_summary_column_colored_green_when_rows_agree(tmp_path):
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ПІБ", D1, D2], [["Перший Перший", 100, 30]])
    _write_xlsx(check_dir / "ОБЛІК2.xlsx", ["ПІБ", D1, D2], [["Перший Перший", 100, 30]])
    output_path = tmp_path / "RESULT_accounting.xlsx"

    checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb.active
    summary_col = ws.max_column
    assert ws.cell(row=2, column=summary_col).value == "100 - 1 днів; 30 - 1 днів"
    assert ws.cell(row=2, column=summary_col).fill.fgColor.rgb == checker.GREEN_FILL.fgColor.rgb
    assert ws.cell(row=3, column=summary_col).fill.fgColor.rgb == checker.GREEN_FILL.fgColor.rgb


def test_check_accounting_summary_column_colored_red_when_rows_disagree(tmp_path):
    """Якщо самі дати розходяться між файлами - розходиться й похідний підсумок,
    тож колонка "Підсумок" ТЕЖ фарбується червоною (та сама логіка порівняння,
    що й для дат - жодного окремого винятку для цієї колонки)."""
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ПІБ", D1], [["Перший Перший", 100]])
    _write_xlsx(check_dir / "ОБЛІК2.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])
    output_path = tmp_path / "RESULT_accounting.xlsx"

    checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb.active
    summary_col = ws.max_column
    assert ws.cell(row=2, column=summary_col).fill.fgColor.rgb == checker.RED_FILL.fgColor.rgb
    assert ws.cell(row=3, column=summary_col).fill.fgColor.rgb == checker.RED_FILL.fgColor.rgb


def test_check_accounting_person_missing_from_one_file_is_flagged_red(tmp_path):
    """Людина, відсутня в одному з файлів, все одно отримує рядок (порожній) для
    цього файлу - і порожнє значення проти заповненого вважається розбіжністю."""
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ПІБ", D1], [["Перший Перший", 100]])
    _write_xlsx(check_dir / "ОБЛІК2.xlsx", ["ПІБ", D1], [])
    output_path = tmp_path / "RESULT_accounting.xlsx"

    checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb.active
    assert ws.cell(row=2, column=1).value == "ОБЛІК1.xlsx"
    assert ws.cell(row=3, column=1).value == "ОБЛІК2.xlsx"
    assert ws.cell(row=2, column=5).value == "Перший Перший"  # canonical ПІБ (з файлу 1)
    assert ws.cell(row=2, column=6).value == 100
    assert ws.cell(row=3, column=6).value is None
    assert ws.cell(row=2, column=6).fill.fgColor.rgb == checker.RED_FILL.fgColor.rgb
    assert ws.cell(row=3, column=6).fill.fgColor.rgb == checker.RED_FILL.fgColor.rgb


def test_check_accounting_canonical_pib_taken_from_first_file_where_person_exists(tmp_path):
    """Якщо людина вперше з'являється НЕ в першому файлі, канонічний ПІБ (для
    об'єднаної клітинки) все одно береться з того файлу, де вона фактично є."""
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ПІБ", D1], [["Другий Другий", 100]])
    _write_xlsx(check_dir / "ОБЛІК2.xlsx", ["ПІБ", D1], [["Другий Другий", 100], ["Третій Третій", 30]])
    output_path = tmp_path / "RESULT_accounting.xlsx"

    checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb.active
    # Третій Третій - лише у файлі 2, тож його група рядків - останні два. ПІБ -
    # об'єднана клітинка: значення зберігається лише у ВЕРХНЬОМУ (row=4) рядку.
    assert ws.cell(row=4, column=1).value == "ОБЛІК1.xlsx"
    assert ws.cell(row=4, column=5).value == "Третій Третій"
    assert ws.cell(row=5, column=1).value == "ОБЛІК2.xlsx"
    assert str(next(r for r in ws.merged_cells.ranges if r.min_row == 4 and r.min_col == 5)) == "E4:E5"


def test_check_accounting_canonical_label_skips_file_missing_that_column(tmp_path):
    """Файл 1 має ЗВАННЯ/ПІБ, але взагалі не має колонки ПОСАДА/ПІДРОЗДІЛ у заголовку
    (не порожнє значення - колонки НЕМАЄ). Канонічне ПОСАДА/ПІДРОЗДІЛ має братись з
    файлу 2 (де ця колонка є), а не лишатись порожнім через те, що людину вже знайдено
    у файлі 1."""
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ЗВАННЯ", "ПІБ", D1], [["сержант", "Перший Перший", 100]])
    _write_xlsx(check_dir / "ОБЛІК2.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1], [
        ["Підрозділ 1", "Стрілець", "сержант", "Перший Перший", 100],
    ])
    output_path = tmp_path / "RESULT_accounting.xlsx"

    checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb.active
    assert ws.cell(row=2, column=2).value == "Підрозділ 1"       # ПІДРОЗДІЛ - з файлу 2
    assert ws.cell(row=2, column=3).value == "Стрілець"   # ПОСАДА - з файлу 2
    assert ws.cell(row=2, column=4).value == "сержант"    # ЗВАННЯ - з файлу 1 (є в обох)


def test_check_accounting_unions_dates_across_files_with_different_date_sets(tmp_path):
    """Кожен файл може мати ІНШИЙ набір дат - результат містить об'єднання (union)
    усіх дат, відсортоване хронологічно; відсутня в конкретному файлі дата -
    порожнє (і, відповідно, розбіжне) значення для нього."""
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ПІБ", D1, D3], [["Перший Перший", 100, 100]])
    _write_xlsx(check_dir / "ОБЛІК2.xlsx", ["ПІБ", D2], [["Перший Перший", 30]])
    output_path = tmp_path / "RESULT_accounting.xlsx"

    checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb.active
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    assert header[5:] == [D1, D2, D3, "Підсумок"]


def test_check_accounting_includes_xlsm_files_alongside_xlsx(tmp_path):
    """Доданий .xlsm (напр. ще один ОБЛІК3.xlsm) враховується так само, як .xlsx -
    людина з нього отримує власний рядок групи, а не губиться."""
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ПІБ", D1], [["Перший Перший", 100]])
    _write_xlsx(check_dir / "ОБЛІК2.xlsx", ["ПІБ", D1], [["Перший Перший", 100]])
    _write_xlsx(check_dir / "ОБЛІК3.xlsm", ["ПІБ", D1], [["Перший Перший", 30]])
    output_path = tmp_path / "RESULT_accounting.xlsx"

    checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb.active
    assert [ws.cell(row=r, column=1).value for r in (2, 3, 4)] == ["ОБЛІК1.xlsx", "ОБЛІК2.xlsx", "ОБЛІК3.xlsm"]
    assert [ws.cell(row=r, column=6).value for r in (2, 3, 4)] == [100, 100, 30]
    # ОБЛІК3.xlsm розходиться зі значенням двох інших файлів - розбіжність помітна.
    assert ws.cell(row=2, column=6).fill.fgColor.rgb == checker.RED_FILL.fgColor.rgb


def test_check_accounting_single_file_treated_as_trivially_matching(tmp_path):
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ПІБ", D1], [["Перший Перший", 100]])
    output_path = tmp_path / "RESULT_accounting.xlsx"

    result = checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    assert result == str(output_path)
    wb = load_workbook(str(output_path))
    ws = wb.active
    assert ws.cell(row=2, column=6).fill.fgColor.rgb == checker.GREEN_FILL.fgColor.rgb
    assert list(ws.merged_cells.ranges) == []


def test_check_accounting_file_paths_bypasses_directory_globbing(tmp_path):
    """file_paths= звіряє РІВНО передані файли (у цьому порядку), НЕ скануючи
    check_dir взагалі - потрібно для resources/changes/, де в одній директорії
    лежить кілька незалежних пар prev/actual за різні місяці і звичайний глоб
    усієї директорії змішав би їх в одну звірку."""
    check_dir = tmp_path / "changes"
    check_dir.mkdir()
    prev_path = _write_xlsx(check_dir / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 100]])
    actual_path = _write_xlsx(check_dir / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])
    # Файли ІНШОЇ пари (інший місяць) в ТІЙ САМІЙ директорії - НЕ повинні потрапити
    # в результат, якщо передано явний file_paths.
    _write_xlsx(check_dir / "prev_05_ОБЛІК.xlsx", ["ПІБ", D1], [["Другий Другий", 100]])
    _write_xlsx(check_dir / "actual_05_ОБЛІК.xlsx", ["ПІБ", D1], [["Другий Другий", 100]])
    output_path = tmp_path / "changes_06_ОБЛІК.xlsx"

    result = checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path), file_paths=[str(prev_path), str(actual_path)])

    assert result == str(output_path)
    wb = load_workbook(str(output_path))
    ws = wb.active
    assert [ws.cell(row=r, column=1).value for r in (2, 3)] == ["prev_06_ОБЛІК.xlsx", "actual_06_ОБЛІК.xlsx"]
    # ПІБ - об'єднана клітинка (2 рядки цієї людини): значення після перезбереження
    # лишається лише на "якірному" (верхньому) рядку цього діапазону.
    assert ws.cell(row=2, column=5).value == "Перший Перший"
    all_pib_values = {ws.cell(row=r, column=5).value for r in range(2, ws.max_row + 1)}
    assert "Другий Другий" not in all_pib_values


def test_check_accounting_creates_output_directory_if_missing(tmp_path):
    check_dir = tmp_path / "check"
    check_dir.mkdir()
    _write_xlsx(check_dir / "ОБЛІК1.xlsx", ["ПІБ", D1], [["Перший Перший", 100]])
    output_path = tmp_path / "nested" / "deeper" / "RESULT_accounting.xlsx"

    result = checker.check_accounting(check_dir=str(check_dir), output_path=str(output_path))

    assert result == str(output_path)
    assert output_path.exists()
