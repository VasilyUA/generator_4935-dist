"""Книга налаштувань resources/НАЛАШТУВАННЯ.xlsx (utils/project_settings.py,
utils/settings_migration.py, sync/stage_folder.py) - на синтетичних книгах у
tmp_path, без справжніх resources/."""
import importlib.util
import json
import os
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

from utils import project_settings as ps
from utils import settings_book as sb
from utils import settings_migration as sm

_PROJECT_DIR = Path(__file__).resolve().parent.parent
_DAY = "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY"
_WEEK = "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK"
_SAVE = "NUMBER_OF_DOCUMENTS_BRS_SAVE"
_TODAY = "27.11.2026"


def _sample_module():
    spec = importlib.util.spec_from_file_location("money_generate_sample", _PROJECT_DIR / "resources_example" / "generate_sample.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _without_row_numbers(settings):
    def strip(value):
        if isinstance(value, list) and value and isinstance(value[0], dict):
            return [{key: item for key, item in row.items() if key != "_row"} for row in value]
        return value
    return {
        key: ({name: strip(rows) for name, rows in value.items()} if key == "numbers" else strip(value))
        for key, value in settings.items()
    }


@pytest.fixture(scope="module")
def sample():
    return _sample_module().settings_example(datetime(2026, 11, 27))


def _write(tmp_path, settings):
    path = tmp_path / sb.SETTINGS_FILE_NAME
    ps.write_settings(str(path), settings)
    return path


def _headers(sheet):
    return {cell.value: cell.column for cell in sheet[sb.HEADER_ROW] if cell.value}


def _locate(workbook, table, header):
    """(аркуш, {заголовок: колонка}, рядки даних) таблиці table книги: на
    аркуші з кількома таблицями - від смуги "■ table" до наступної смуги;
    рядок заголовків - перший рядок таблиці, де є заголовок header."""
    sheet_name, section, _legacy = ps._where(table)
    sheet, _label, first_row, last_row = sb.locate_table(workbook, sheet_name, section)
    last_row = last_row or sheet.max_row
    for row in range(first_row, last_row + 1):
        headers = {cell.value: cell.column for cell in sheet[row] if isinstance(cell.value, str)}
        if header in headers:
            return sheet, headers, range(row + 1, last_row + 1)
    raise AssertionError(f"у таблиці {table} немає заголовка {header}")


def _edit_row(path, table, match, header, value):
    """Записує value в колонку header першого рядка таблиці table, у якому
    колонки з match мають ці значення. Повертає номер рядка (для повідомлень
    про помилку)."""
    workbook = sb.load_editable(str(path))
    sheet, headers, rows = _locate(workbook, table, next(iter(match)))
    for row in rows:
        if all(sb.cell_text(sheet.cell(row=row, column=headers[name]).value) == str(wanted) for name, wanted in match.items()):
            sheet.cell(row=row, column=headers[header]).value = value
            workbook.save(path)
            return row
    raise AssertionError(f"у таблиці {table} немає рядка {match}")


# --- зразок -----------------------------------------------------------------


def test_generate_sample_creates_readable_resources_once(tmp_path, capsys):
    module = _sample_module()
    target = tmp_path / "resources"

    module.main(str(target))

    assert sorted(path.name for path in target.iterdir()) == sorted(["НАЛАШТУВАННЯ.xlsx", "ОБЛІК.xlsx"])
    settings = ps.read_settings(str(target / "НАЛАШТУВАННЯ.xlsx"))
    assert _without_row_numbers(settings) == module.settings_example()
    positions = [row[1] for row in load_workbook(target / "ОБЛІК.xlsx").worksheets[0].iter_rows(min_row=2, values_only=True)]
    # Без командира й начальника штабу БР не згенерувати - зразок має працювати з коробки.
    assert settings["general"]["HIGHER_COMMANDER_TITLE"] in positions
    assert settings["general"]["COMMANDER_TITLE"] in positions

    settings_bytes = (target / "НАЛАШТУВАННЯ.xlsx").read_bytes()
    module.main(str(target))
    assert (target / "НАЛАШТУВАННЯ.xlsx").read_bytes() == settings_bytes
    assert capsys.readouterr().out.count("вже існує") == 2


def test_sample_subdivisions_cover_every_subdivision_of_the_sample_roster(sample):
    built = ps.build_constants(sample, _TODAY)
    roster_keys = {ps.section_key(person[0]) for person in _sample_module().PEOPLE}
    section_keys = {part for key in built.PARAGRAPH_MAP for part in key.split(", ")}

    assert roster_keys <= section_keys


# --- побудова структур constants.py -------------------------------------------


def test_built_categories_have_the_shape_constants_used_to_have(sample):
    built = ps.build_constants(sample, _TODAY)
    categories = built.MONEY_REPORT_CATEGORIES

    assert list(categories)[:4] == [30, 50, 100, "РТГр_100"]
    assert categories[30]["general"] == [{"start": "01.10.2026", "end": _TODAY, "lines": ["ЖБД (Справа №0т)"]}]
    assert categories[100]["general"] == built.MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR + built.MONEY_REPORT_GENERAL_REFERENCES_BN
    assert categories["РТГр_100"]["РТГр"]["exclude_general"] == "all"
    assert categories["NOT_PAID"]["column_basis_from_tracking"] == "ПІДСТАВИ ДЛЯ НЕ ВИПЛАТИ"
    assert categories["100_БПШП"]["БПШП|100_БПШП|100_ШПБП"]["column_basis_from_tracking"] == ("ПІДСТАВИ ШПБП", "ПІДСТАВИ БПШП")
    assert categories[10]["ЖИТТЄДІЯЛЬНІСТЬ"] == {
        "grounds": [], "use_brs": False, "use_brs_from_selected_folder": False, "exclude_general": [],
        "default": True, "generate_br": False, "include_to_report": True, "required_basis": False,
        "column_basis_from_tracking": "ПІДСТАВИ 10К",
    }
    assert list(built.COMMANDER_MONEY_REPORT_CATEGORIES[100]) == ["general", "ОБОРОНА"]


def test_references_keep_ids_and_open_ends_run_until_today(sample):
    built = ps.build_constants(sample, _TODAY)

    assert built.MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR == [
        {"start": "01.10.2026", "end": "14.11.2026", "lines": ["ЖБД 0 бат 00 обрбр №000дск/1 від 01.10.2026"], "id": "JBD_1"},
        {"start": "15.11.2026", "end": _TODAY, "lines": ["ЖБД 0 бат 00 обрбр №000дск/2 від 15.11.2026"], "id": "JBD_2"},
    ]
    assert built.MONEY_REPORT_CHANGES_ORDER_REFERENCES == [
        {"start": "01.10.2026", "end": "31.10.2026", "lines": ["№0000 від 01.11.2026"], "id": "C_1"},
    ]


def test_only_higher_hq_rows_are_in_br_hight_unit(sample):
    built = ps.build_constants(sample, _TODAY)

    assert list(built.BR_HIGHT_UNIT) == ["Бригада загальні", "ЖИТТЄДІЯЛЬНІСТЬ", "РТГр"]
    assert len(built.BR_HIGHT_UNIT["ЖИТТЄДІЯЛЬНІСТЬ"]) == 2


def test_texts_subdivisions_and_task_variants(sample):
    built = ps.build_constants(sample, _TODAY)

    assert built.SEQUENCE == ["1 рмп", "2 рмп", "мп"]
    assert list(built.PARAGRAPH_MAP) == ["1рмп", "2рмп", "мп", "упр, штаб"]
    assert "{coordinates}" in built.PARAGRAPH_MAP["упр, штаб"]
    assert [len(variant) for variant in built.TASK_ORDER_VARIANTS] == [2, 2]
    assert built.MONEY_REPORT_STATIC["REPORT_TITLE"] == "РАПОРТ"
    assert set(built.MONEY_REPORT_STATIC["SECTIONS"]) >= {str(point) for point in built.MONEY_REPORT_CATEGORIES}
    assert "РТГр_100" not in built.MONEY_REPORT_STATIC["COMMANDER_SECTIONS"]  # порожня клітинка - тексту немає


@pytest.mark.parametrize("name, key", [("1 рмп", "1рмп"), ("Упр , Штаб", "упр, штаб"), ("мп", "мп")])
def test_section_key(name, key):
    assert ps.section_key(name) == key


def test_empty_task_variant_row_keeps_the_rotation_slot():
    rows = [
        {"variant": 2, "title": "Б", "text": "завдання Б", "_row": 2},
        {"variant": 1, "title": "А", "text": "завдання А", "_row": 3},
        {"variant": 3, "title": "", "text": "", "_row": 4},
    ]

    assert ps.build_task_order_variants(rows) == [[("А", "завдання А")], [("Б", "завдання Б")], []]


# --- номери документів ---------------------------------------------------------


def test_number_dicts_follow_the_chosen_month(sample):
    day, week, save = ps.build_number_dicts(sample["numbers"], 2, 2027)

    assert list(day)[:2] == ["31.01.2027", "01.02.2027"]
    assert list(day)[-1] == "28.02.2027"  # 29-31 у лютому 2027 немає
    assert day["01.02.2027"] == {"бат": "101", "брг": "1001", "general_br_bat": "100 від 01.11.2026"}
    assert len(week) == len(save) == 28


def test_previous_month_row_of_january_is_the_last_of_december():
    assert ps.date_key_for_day(sb.PREVIOUS_MONTH_DAY, 1, 2027) == "31.12.2026"
    assert ps.date_key_for_day(31, 4, 2027) is None


def test_day_labels_read_the_day_from_the_key_text():
    keys = ["31.10.2026", "01.11.2026", "31.11.2026", "05.12.2026", "ключ"]

    assert ps.day_labels(keys) == {"31.10.2026": sb.PREVIOUS_MONTH_DAY, "01.11.2026": 1, "31.11.2026": 31}


def test_day_labels_previous_row_is_only_the_last_day_of_the_previous_month():
    assert ps.day_labels(["30.10.2026", "01.11.2026", "02.11.2026"]) == {"01.11.2026": 1, "02.11.2026": 2}
    # Нічия (по ключу на місяць) - основний пізніший місяць.
    assert ps.day_labels(["31.10.2026", "01.11.2026"]) == {"31.10.2026": sb.PREVIOUS_MONTH_DAY, "01.11.2026": 1}


def test_write_number_fields_updates_only_the_given_fields(tmp_path, sample):
    path = _write(tmp_path, sample)
    day, _week, save = ps.build_number_dicts(sample["numbers"], 11, 2026)
    day["01.11.2026"]["брг"] = "7777"
    save["02.11.2026"]["бз_бат"] = ""

    save["01.12.2026"] = {"бз_бат": "9", "бз_брг": "9"}  # документ чужого місяця з папки

    missing = ps.write_number_fields(str(path), [
        (_DAY, "брг", day, ["31.10.2026", "01.11.2026"]),
        (_SAVE, "бз_бат", save, ["01.11.2026", "02.11.2026", "01.12.2026", "03.11.2026"]),
    ])

    assert missing == []  # 01.12.2026 - чужий місяць, рядка для нього й не шукаємо
    numbers = ps.read_settings(str(path))["numbers"]
    assert numbers[_DAY][1]["брг"] == "7777" and numbers[_DAY][1]["бат"] == "101"
    assert numbers[_SAVE][1]["бз_бат"] == "" and numbers[_SAVE][1]["бз_брг"] == "2001"


def test_numbers_workbook_round_trip(tmp_path):
    path = tmp_path / "НОМЕРИ ДОКУМЕНТІВ.xlsx"
    numbers = {
        _DAY: {"31.10.2026": {"бат": "1", "брг": "", "general_br_bat": ""}, "01.11.2026": {"бат": "2", "брг": "3", "general_br_bat": "4 від 01.11.2026"}},
        _WEEK: {"01.11.2026": {"бат": "4", "посилання_брг": "", "посилання_бат": ""}},
        _SAVE: {},
    }
    ps.write_numbers_workbook(str(path), numbers)

    loaded = ps.read_numbers_workbook(str(path), {name: list(value) + ["02.11.2026"] for name, value in numbers.items()})

    assert loaded == numbers


# --- помилки в книзі --------------------------------------------------------------


def test_unknown_group_lists_the_known_ones(tmp_path, sample):
    path = _write(tmp_path, sample)
    row = _edit_row(path, ps.POINTS_SHEET, {"Пункт": 30, "Категорія": "ЖИТТЄДІЯЛЬНІСТЬ"}, "Підстави", "ЖИТТЄ")

    with pytest.raises(sb.SettingsError) as error:
        ps.build_constants(ps.read_settings(str(path)), _TODAY)

    assert str(error.value).startswith(
        f'НАЛАШТУВАННЯ.xlsx, аркуш "Пункти рапорту", таблиця "Рапорт на особовий склад", рядок {row}, колонка "Підстави": '
        'немає групи підстав "ЖИТТЄ"'
    )
    assert "(є: ЖБД, БН, Бригада загальні" in str(error.value)


def test_group_names_are_matched_without_regard_to_case(tmp_path, sample):
    path = _write(tmp_path, sample)
    _edit_row(path, ps.POINTS_SHEET, {"Пункт": 100, "Категорія": ps.POINT_ROW_MARKER}, "Підстави", "жбд, бн")

    built = ps.build_constants(ps.read_settings(str(path)), _TODAY)

    assert len(built.MONEY_REPORT_CATEGORIES[100]["general"]) == 3


def test_duplicate_category_of_a_point_is_an_error(tmp_path, sample):
    path = _write(tmp_path, sample)
    row = _edit_row(path, ps.POINTS_SHEET, {"Пункт": 30, "Категорія": "РТГр_30"}, "Категорія", "ЖИТТЄДІЯЛЬНІСТЬ")

    with pytest.raises(sb.SettingsError, match=f"рядок {row}, колонка \"Категорія\": категорія 'ЖИТТЄДІЯЛЬНІСТЬ' уже є в пункті 30"):
        ps.build_constants(ps.read_settings(str(path)), _TODAY)


def test_reserved_category_name_is_an_error(tmp_path, sample):
    path = _write(tmp_path, sample)
    _edit_row(path, ps.POINTS_SHEET, {"Пункт": 30, "Категорія": "ЖИТТЄДІЯЛЬНІСТЬ"}, "Категорія", "general")

    with pytest.raises(sb.SettingsError, match="службова назва"):
        ps.build_constants(ps.read_settings(str(path)), _TODAY)


def test_duplicate_day_is_an_error(tmp_path, sample):
    path = _write(tmp_path, sample)
    first = _edit_row(path, ps.NUMBERS_SHEET, {"День": 1}, "День", 1)
    second = _edit_row(path, ps.NUMBERS_SHEET, {"День": 2}, "День", 1)

    with pytest.raises(sb.SettingsError, match=f'"Номери документів", рядок {second}, колонка "День": 1 уже є в рядку {first}'):
        ps.read_settings(str(path))


def test_duplicate_point_is_an_error(tmp_path, sample):
    path = _write(tmp_path, sample)
    _edit_row(path, ps.POINTS_SHEET, {"Пункт": 50, "Категорія": ps.POINT_ROW_MARKER}, "Пункт", 30)

    with pytest.raises(sb.SettingsError, match="пункт 30 уже є вище"):
        ps.build_constants(ps.read_settings(str(path)), _TODAY)


def test_task_variant_needs_both_title_and_text(tmp_path, sample):
    path = _write(tmp_path, sample)
    _edit_row(path, ps.TASK_VARIANTS_SHEET, {"Підрозділ": "1 рмп (оборона РОП «Приклад-1»)"}, "Завдання", None)

    with pytest.raises(sb.SettingsError, match="потрібні і підрозділ, і завдання"):
        ps.build_constants(ps.read_settings(str(path)), _TODAY)


def test_points_may_be_written_as_numbers_or_text(tmp_path, sample):
    path = _write(tmp_path, sample)
    sheet, headers, rows = _locate(load_workbook(path), ps.POINTS_SHEET, "Пункт")
    assert sheet.cell(row=rows[0], column=headers["Пункт"]).value == 30  # число, як у клітинці дати ОБЛІК.xlsx
    _edit_row(path, ps.POINTS_SHEET, {"Пункт": 30, "Категорія": ps.POINT_ROW_MARKER}, "Пункт", "30")

    assert list(ps.build_constants(ps.read_settings(str(path)), _TODAY).MONEY_REPORT_CATEGORIES)[0] == 30


# --- вигляд книги: аркуші пунктів, підстав і номерів -------------------------------------


def test_points_sheet_is_one_block_per_point(tmp_path, sample):
    """Пункт - рядок "весь пункт" з його загальними підставами, а під ним - його категорії."""
    sheet, headers, data_rows = _locate(load_workbook(_write(tmp_path, sample)), ps.POINTS_SHEET, "Пункт")
    rows = [
        (sheet.cell(row=row, column=headers["Пункт"]).value, sheet.cell(row=row, column=headers["Категорія"]).value,
         sheet.cell(row=row, column=headers["Підстави"]).value)
        for row in data_rows[:6]
    ]

    assert rows == [
        (30, ps.POINT_ROW_MARKER, "Бригада загальні"),
        (30, "ЖИТТЄДІЯЛЬНІСТЬ", "ЖИТТЄДІЯЛЬНІСТЬ"),
        (30, "РТГр_30", "РТГр"),
        (50, ps.POINT_ROW_MARKER, None),
        (50, "ПУ", None),
        (100, ps.POINT_ROW_MARKER, "ЖБД, БН"),
    ]


def test_category_rows_may_leave_the_point_empty(tmp_path, sample):
    path = _write(tmp_path, sample)
    _edit_row(path, ps.POINTS_SHEET, {"Пункт": 30, "Категорія": "РТГр_30"}, "Пункт", None)
    _edit_row(path, ps.POINTS_SHEET, {"Пункт": 50, "Категорія": "ПУ"}, "Пункт", None)

    settings = ps.read_settings(str(path))

    assert _without_row_numbers(settings)["categories"] == sample["categories"]


def test_empty_category_is_a_point_row(tmp_path, sample):
    path = _write(tmp_path, sample)
    _edit_row(path, ps.POINTS_SHEET, {"Пункт": 30, "Категорія": ps.POINT_ROW_MARKER}, "Категорія", None)

    assert _without_row_numbers(ps.read_settings(str(path)))["points"] == sample["points"]


def test_category_flags_in_a_point_row_are_an_error(tmp_path, sample):
    """Прапорець у рядку "весь пункт" нічого б не змінив - краще сказати, ніж мовчки ігнорувати."""
    path = _write(tmp_path, sample)
    row = _edit_row(path, ps.POINTS_SHEET, {"Пункт": 30, "Категорія": ps.POINT_ROW_MARKER}, "Включати в рапорт", "Ні")

    with pytest.raises(sb.SettingsError, match=f'рядок {row}, колонка "Включати в рапорт": у рядку "весь пункт" ця колонка не діє'):
        ps.read_settings(str(path))


def test_empty_category_flags_take_their_defaults(tmp_path, sample):
    path = _write(tmp_path, sample)
    for header in ("Додати ЗАВДАННЯ", "Підстава обов'язкова", "За замовчуванням", "Генерувати БР", "Включати в рапорт"):
        _edit_row(path, ps.POINTS_SHEET, {"Пункт": 50, "Категорія": "ПУ"}, header, None)

    category = next(row for row in ps.read_settings(str(path))["categories"] if row["category"] == "ПУ")

    assert {key: category[key] for key in ("use_brs", "required_basis", "default", "generate_br", "include_to_report")} == {
        "use_brs": False, "required_basis": True, "default": False, "generate_br": True, "include_to_report": True,
    }


def test_references_are_split_into_tables_and_read_back_as_one_list(tmp_path, sample):
    workbook = load_workbook(_write(tmp_path, sample))

    def groups(table, column=None):
        sheet, headers, rows = _locate(workbook, table, "Початок")
        if column is None:
            return sum(1 for row in rows if sheet.cell(row=row, column=headers["Початок"]).value)
        return [sheet.cell(row=row, column=headers[column]).value for row in rows
                if sheet.cell(row=row, column=headers[column]).value]

    assert groups(ps.LOG_WAR_SHEET) == 2 and groups(ps.BN_SHEET) == 1
    assert groups(ps.HIGHER_REFERENCES_SHEET, "Група") == ["Бригада загальні", "ЖИТТЄДІЯЛЬНІСТЬ", "ЖИТТЄДІЯЛЬНІСТЬ", "РТГр"]
    assert groups(ps.OTHER_REFERENCES_SHEET, "Група") == ["РТГр_100", "ЗРАДН", "РТГр_170", "70_РТГр"]


def test_reference_rows_of_one_group_are_written_together():
    rows = [
        {"group": "А", "higher_hq": False, "start": "01.11.2026"},
        {"group": "Б", "higher_hq": False, "start": "02.11.2026"},
        {"group": "а", "higher_hq": False, "start": "03.11.2026"},
        {"group": "ЖБД", "higher_hq": True, "start": "04.11.2026"},
    ]

    by_sheet = ps.reference_sheet_rows(rows)

    assert [row["start"] for row in by_sheet[ps.OTHER_REFERENCES_SHEET]] == ["01.11.2026", "03.11.2026", "02.11.2026"]
    assert by_sheet[ps.HIGHER_REFERENCES_SHEET] == [rows[3]]  # ЖБД вищого штабу - не на аркуші ЖБД
    assert by_sheet[ps.LOG_WAR_SHEET] == []


def test_numbers_of_three_kinds_share_one_row_per_day(tmp_path, sample):
    sheet = load_workbook(_write(tmp_path, sample))[ps.NUMBERS_SHEET]
    headers = _headers(sheet)
    first_day = sb.FIRST_DATA_ROW + 1  # над ним - рядок "попередній"

    assert sheet.cell(row=sb.FIRST_DATA_ROW, column=1).value == sb.PREVIOUS_MONTH_DAY
    assert [sheet.cell(row=first_day, column=headers[header]).value
            for header in ("День", "БР бат №", "ЗАВДАННЯ бат №", "БЗ бат №")] == [1, 101, 100, 104]


def test_stage_numbers_workbook_of_the_former_layout_is_still_read(tmp_path):
    """Книга номерів етапу, записана до спільного аркуша (три окремі аркуші), читається й так."""
    path = tmp_path / "НОМЕРИ ДОКУМЕНТІВ.xlsx"
    workbook = sb.new_workbook()
    for name, sheet_name in ps.LEGACY_NUMBER_SHEETS.items():
        columns = (ps.DAY_COLUMN, *(column for _field, column in ps.NUMBER_TABLES[name]))
        rows = [{"day": 1, **{column.key: str(index + 10) for index, (_field, column) in enumerate(ps.NUMBER_TABLES[name])}}]
        sb.write_table(workbook, sheet_name, columns, rows)
    sb.save_workbook(workbook, str(path))

    loaded = ps.read_numbers_workbook(str(path), {_DAY: ["01.11.2026"], _SAVE: ["01.11.2026"]})

    assert loaded == {
        _DAY: {"01.11.2026": {"бат": "10", "брг": "11", "general_br_bat": "12"}},
        _SAVE: {"01.11.2026": {"бз_бат": "10", "бз_брг": "11"}},
    }


def test_not_paid_exceptions_come_from_their_sheet(tmp_path, sample):
    path = _write(tmp_path, sample)
    workbook = sb.load_editable(str(path))
    sheet, headers, rows = _locate(workbook, ps.NOT_PAID_EXCEPTIONS_SHEET, "Статус невиплати")
    row = next(row for row in rows if sheet.cell(row=row, column=headers["Статус невиплати"]).value is None)
    sheet.cell(row=row, column=headers["Статус невиплати"]).value = "Задув."
    sheet.cell(row=row, column=headers["Не враховувати, якщо в місяці є"]).value = "РОЗП, Повер"
    sb.save_workbook(workbook, str(path))

    built = ps.build_constants(ps.read_settings(str(path)), _TODAY)

    assert built.NOT_PAID_IGNORED_WHEN_CONFIRMED_BY == {"СЗЧ": {"РОЗП"}, "Задув.": {"РОЗП", "Повер"}}


def _write_former_layout(path, settings, hidden=()):
    """Книга колишнього вигляду: кожна таблиця - на власному аркуші (назва
    аркуша - колишня назва таблиці), без смуг "■"; аркуші hidden - приховані."""
    workbook = sb.new_workbook()
    ps._write_numbers_sheet(workbook, settings["numbers"])
    for table, section in ps._sections(settings).items():
        sheet_name = ps.LEGACY_SHEETS.get(table, table)
        if section.settings is not None:
            sheet = sb.write_settings_sheet(workbook, sheet_name, section.settings, section.values or {})
        else:
            sheet = sb.write_table(workbook, sheet_name, section.columns, section.rows, band_by=section.band_by)
        if sheet_name in hidden:
            sheet.sheet_state = "hidden"
    sb.save_workbook(workbook, str(path))


def test_workbook_without_not_paid_exceptions_gets_the_table(tmp_path, sample, capsys):
    """Книга, створена до появи таблиці "Винятки невиплати" (колись - словник у
    constants.py): діють початкові винятки, а книга переписується вже з цією
    таблицею."""
    path = tmp_path / sb.SETTINGS_FILE_NAME
    _write_former_layout(path, dict(sample, not_paid_exceptions=None))
    workbook = sb.load_editable(str(path))
    del workbook[ps.NOT_PAID_EXCEPTIONS_SHEET]
    sb.save_workbook(workbook, str(path))
    assert ps.read_settings(str(path))["not_paid_exceptions"] is None

    settings = ps.load_settings(str(path), None, str(tmp_path))

    assert "переведено на новий вигляд" in capsys.readouterr().out
    assert ps.build_constants(settings, _TODAY).NOT_PAID_IGNORED_WHEN_CONFIRMED_BY == {"СЗЧ": {"РОЗП"}}
    with sb.open_for_reading(str(path)) as workbook:
        assert sb.has_table(workbook, ps.POINTS_BOOK_SHEET, ps.NOT_PAID_EXCEPTIONS_SHEET)


def test_former_layout_is_read_and_rewritten_once_into_tables_on_few_sheets(tmp_path, sample, capsys):
    """Книга, де кожна таблиця мала власний аркуш (і рідкісні були приховані),
    читається так само, а load_settings один раз переписує її в теперішній
    вигляд - ті самі дані, копія колишньої лишається поруч."""
    path = tmp_path / sb.SETTINGS_FILE_NAME
    _write_former_layout(path, sample, hidden={ps.SECTION_TEXTS_SHEET, "Тексти рапорту", ps.TASK_TEXTS_SHEET})
    former = ps.read_settings(str(path))
    assert ps._layout_view(former) == ps._layout_view(sample)

    settings = ps.load_settings(str(path), None, str(tmp_path))

    assert "переведено на новий вигляд" in capsys.readouterr().out
    assert ps._layout_view(settings) == ps._layout_view(sample)
    assert vars(ps.build_constants(settings, _TODAY)) == vars(ps.build_constants(former, _TODAY))
    workbook = load_workbook(path)
    assert workbook.sheetnames == [sb.INSTRUCTIONS_SHEET, ps.NUMBERS_SHEET, *(sheet for sheet, *_rest in ps.BOOK_SHEETS)]
    backup = tmp_path / "НАЛАШТУВАННЯ (до оновлення).xlsx"
    assert len(load_workbook(backup).sheetnames) > 10  # колишня книга - аркуш на кожну таблицю

    ps.load_settings(str(path), None, str(tmp_path))  # уже в теперішньому вигляді - більше не переписується

    assert "новий вигляд" not in capsys.readouterr().out


def test_backup_of_an_earlier_upgrade_is_never_overwritten(tmp_path, sample):
    path = tmp_path / sb.SETTINGS_FILE_NAME
    _write_former_layout(path, sample)
    earlier = tmp_path / "НАЛАШТУВАННЯ (до оновлення).xlsx"
    earlier.write_bytes(b"earlier backup")

    ps.load_settings(str(path), None, str(tmp_path))

    assert earlier.read_bytes() == b"earlier backup"
    assert len(load_workbook(tmp_path / "НАЛАШТУВАННЯ (до оновлення 2).xlsx").sheetnames) > 7


def test_unreadable_rewrite_brings_the_former_workbook_back(tmp_path, sample, monkeypatch, capsys):
    path = tmp_path / sb.SETTINGS_FILE_NAME
    _write_former_layout(path, sample)
    former_bytes = path.read_bytes()
    real_read = ps.read_settings
    calls = []

    def read_settings(target):
        calls.append(target)
        if len(calls) == 2:  # книга, щойно переписана в новий вигляд
            raise sb.SettingsError("зіпсовано")
        return real_read(target)

    monkeypatch.setattr(ps, "read_settings", read_settings)

    settings = ps.load_settings(str(path), None, str(tmp_path))

    assert "колишню повернуто" in capsys.readouterr().out
    assert path.read_bytes() == former_bytes
    assert ps._layout_view(settings) == ps._layout_view(sample)


def test_few_sheets_with_several_tables_and_instructions_say_where_to_change_what(tmp_path, sample):
    workbook = load_workbook(_write(tmp_path, sample))
    instructions = workbook[sb.INSTRUCTIONS_SHEET]
    texts = [cell.value for row in instructions.iter_rows() for cell in row if cell.value]
    links = {cell.value: cell.hyperlink.location for row in instructions.iter_rows() for cell in row if cell.hyperlink}

    assert all(sheet.sheet_state == "visible" for sheet in workbook.worksheets)
    assert workbook.sheetnames == [sb.INSTRUCTIONS_SHEET, ps.NUMBERS_SHEET, *(sheet for sheet, *_rest in ps.BOOK_SHEETS)]
    assert len(workbook.sheetnames) <= 7
    assert sb.TASKS_HEADING in texts
    # Кожна таблиця - смуга "■ назва" на своєму аркуші, і посилання з
    # "Інструкції" веде саме на неї.
    for sheet_name, table in ps.BOOK_TABLES:
        if table is None:
            continue
        sheet = workbook[sheet_name]
        band_row = next(row for row in range(1, sheet.max_row + 1) if sheet.cell(row=row, column=1).value == sb.section_band_text(table))
        assert links[f"{sheet_name} > {table}"] == f"'{sheet_name}'!A{band_row}"
    assert ps.build_constants(ps.read_settings(str(tmp_path / sb.SETTINGS_FILE_NAME)), _TODAY).MONEY_REPORT_STATIC["REPORT_TITLE"]


def test_tables_of_one_sheet_share_its_columns_sensibly(tmp_path, sample):
    """Таблиці одного аркуша мають спільні колонки аркуша: колонки дат таблиць
    підстав - вузькі, бо в усіх таблицях на тих самих місцях."""
    workbook = load_workbook(_write(tmp_path, sample))
    sheet = workbook[ps.REFERENCES_BOOK_SHEET]

    for table in (ps.LOG_WAR_SHEET, ps.HIGHER_REFERENCES_SHEET, ps.CHANGES_SHEET):
        _sheet, headers, _rows = _locate(workbook, table, "Початок")
        assert (headers["Початок"], headers["Кінець"]) == (1, 2)
    assert sheet.column_dimensions["A"].width < 20 and sheet.column_dimensions["C"].width >= 70


# --- статуси, перейменовані в табелі ----------------------------------------------


_STATUS_ROWS = [("СЗЧ", "СЗЧ"), ("Самовільно", "СЗЧ"), ("100 БПШП", "100_БПШП"), ("30", "30")]
_STATUS_COLUMNS = (sb.Column("status", "Статус"), sb.Column("meaning", "Що означає"), sb.Column("code", "Код"))


def _timesheet_statuses_workbook(path, rows):
    """Книга на кшталт книги налаштувань generator_timesheet давнішого вигляду:
    аркуш "Статуси" (назва статусу й незмінний код)."""
    workbook = sb.new_workbook()
    sb.write_table(workbook, "Статуси", _STATUS_COLUMNS, [{"status": status, "code": code} for status, code in rows])
    sb.save_workbook(workbook, str(path))


def test_renamed_timesheet_statuses_are_read_from_its_workbook(tmp_path):
    from utils import status_names

    path = tmp_path / "НАЛАШТУВАННЯ.xlsx"
    _timesheet_statuses_workbook(path, _STATUS_ROWS)

    assert status_names.read_status_codes(str(path)) == {"САМОВІЛЬНО": "СЗЧ", "100 БПШП": "100_БПШП"}


def test_timesheet_statuses_table_is_found_on_any_sheet(tmp_path):
    """Теперішня книга табеля: таблиця "■ Статуси" - одна з кількох таблиць
    аркуша (і не першого)."""
    from utils import status_names

    path = tmp_path / "НАЛАШТУВАННЯ.xlsx"
    workbook = sb.new_workbook()
    other = sb.Section("Інше", columns=(sb.Column("name", "ПІБ"),), rows=[{"name": "Статус"}])
    sb.write_sections(workbook, "Люди", [other], title="Люди")
    sb.write_sections(workbook, "Статуси", [
        sb.Section("Статуси", columns=_STATUS_COLUMNS, rows=[{"status": status, "code": code} for status, code in _STATUS_ROWS]),
        sb.Section("Групи статусів", columns=(sb.Column("status", "Статус"), sb.Column("code", "Код")),
                   rows=[{"status": "Втік", "code": "Повер"}]),
    ], title="Статуси")
    sb.save_workbook(workbook, str(path))

    assert status_names.read_status_codes(str(path)) == {"САМОВІЛЬНО": "СЗЧ", "100 БПШП": "100_БПШП"}


def test_oblik_values_with_a_new_status_name_become_codes(monkeypatch):
    """ОБЛІК.xlsx, збережений табелем після перейменування статусу, - у
    налаштуваннях і коді цього проєкту статус і далі за кодом."""
    import pandas as pd
    from utils import excel_reader, status_names

    monkeypatch.setattr(status_names, "_CODES", {"САМОВІЛЬНО": "СЗЧ", "100 БПШП": "100_БПШП"})
    frame = pd.DataFrame({"ПІБ": ["А", "Б"], "01.11.2026": [" самовільно", 100], "02.11.2026": ["100 БПШП", "ВП"]})

    rows = excel_reader.get_data(frame, list(frame.columns))["rows"]

    assert rows == [
        {"ПІБ": "А", "01.11.2026": "СЗЧ", "02.11.2026": "100_БПШП"},
        {"ПІБ": "Б", "01.11.2026": 100, "02.11.2026": "ВП"},
    ]


def test_without_the_timesheet_workbook_statuses_stay_as_they_are(tmp_path, monkeypatch):
    from utils import status_names

    monkeypatch.setattr(status_names, "TIMESHEET_SETTINGS_PATH", str(tmp_path / "немає.xlsx"))
    monkeypatch.setattr(status_names, "_CODES", None)

    assert status_names.to_code("Самовільно") == "Самовільно"


# --- завантаження для constants.py -------------------------------------------------


def test_load_settings_without_workbook_or_data_json_explains_what_to_do(tmp_path):
    with pytest.raises(SystemExit, match="generate_sample.py"):
        ps.load_settings(str(tmp_path / sb.SETTINGS_FILE_NAME), str(tmp_path / "data.json"), str(tmp_path))


def test_load_settings_turns_a_workbook_error_into_a_readable_exit(tmp_path, sample, capsys):
    """Пояснення з кроками - червоним (людина читає тільки червоне), а
    SystemExit - лише з кодом 1, щоб Python не друкував той самий текст білим."""
    path = _write(tmp_path, sample)
    workbook = load_workbook(path)
    del workbook[ps.GENERAL_SHEET]
    workbook.save(path)

    with pytest.raises(SystemExit, match='У книзі налаштувань помилка:\n.*немає таблиці "Загальні"') as stopped:
        ps.load_settings(str(path), None, str(tmp_path))

    assert stopped.value.code == 1
    out = capsys.readouterr().out
    assert "\x1b[31mПОМИЛКА В НАЛАШТУВАННЯХ - програму зупинено." in out
    assert f"  1. Відкрийте {path} в Excel." in out


def test_build_constants_or_exit_names_the_problem(tmp_path, sample):
    broken = dict(sample, categories=[dict(sample["categories"][0], grounds=["немає такої"], _row=2)])

    with pytest.raises(SystemExit, match='немає групи підстав "немає такої"'):
        ps.build_constants_or_exit(broken, _TODAY)


# --- перенесення колишніх data.json + constants.py ----------------------------------------


_LEGACY_DATA = {
    "unit": {"SHORT_UNIT_BATTALION": "0 бат", "SHORT_UNIT_BRIGADE": "00 обрбр", "FULL_UNIT_BUT": "0 батальйону",
             "FULL_MILITARY_UNIT": "військової частини А0000", "RTGR_JBD_UNIT_REFERENCE": "РТГР"},
    "br_task_order": {"BR_TASK_SECTION_TEMPLATE": "{unit} утримує район", "TASK_ORDER_VARIANTS": [[["А", "завдання"]], []]},
    "company_sections": {
        "SEQUENCE": ["1 рмп", "2 рмп", "3 рмп"],
        "PARAGRAPH_MAP": {"1рмп": "текст 1", "3рмп": "текст 3", "упр, штаб": "текст упр"},
        "PARAGRAPH_MAP_FOR_EXTRACT": {"1рмп": "витяг 1"},
    },
    "data_statik_report_for_get_additional_money": {
        "REPORT_TITLE": "РАПОРТ", "SECTIONS": {"30": "текст 30", "NOT_PAID": "не платити"}, "COMMANDER_SECTIONS": {"30": "мені 30"},
    },
}

# Стара версія constants.py: номери, підстави й категорії прямо в коді (як до книги).
_LEGACY_SOURCE = '''
import json, os
from datetime import datetime
with open(os.path.join("resources", "data.json"), encoding="utf-8") as f:
    _DATA = json.load(f)
SHORT_UNIT_BATTALION = _DATA["unit"]["SHORT_UNIT_BATTALION"]
_NOW = datetime.now().strftime("%d.%m.%Y")
NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY = {
    "31.10.2026": {"бат": "1", "брг": "2"},
    "01.11.2026": {"бат": "3", "брг": "4", "general_br_bat": "5 від 01.11.2026"},
    "31.11.2026": {"бат": "9", "брг": "9", "general_br_bat": ""},
}
NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK = {"01.11.2026": {"бат": "5", "посилання_брг": "6 від 01.10.2026", "посилання_бат": ""}}
NUMBER_OF_DOCUMENTS_BRS_SAVE = {"01.11.2026": {"бз_бат": "7", "бз_брг": "8"}}
MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR = [{"id": "JBD_1", "start": "01.10.2026", "end": _NOW, "lines": [f"ЖБД {SHORT_UNIT_BATTALION} №1"]}]
MONEY_REPORT_GENERAL_REFERENCES_BN = [{"id": "BN_1", "start": "01.10.2026", "end": "31.10.2026", "lines": ["БН №1"]}]
BR_HIGHT_UNIT = {"general": [{"start": "01.09.2026", "end": _NOW, "lines": ["ЖБД (Справа №2т)"]}],
                 "РТГр": [{"start": "01.09.2026", "end": _NOW, "lines": ["БН ТГр"]}]}
MONEY_REPORT_CHANGES_ORDER_REFERENCES = [{"id": "C_1", "start": "01.10.2026", "end": "31.10.2026", "lines": ["№1 від 07.11.2026"]}]
_FLAGS = {"use_brs": False, "use_brs_from_selected_folder": False, "exclude_general": [], "default": True,
          "generate_br": False, "include_to_report": True, "required_basis": True, "column_basis_from_tracking": None}
MONEY_REPORT_CATEGORIES = {
    30: {"general": BR_HIGHT_UNIT["general"], "РТГр": {**_FLAGS, "grounds": BR_HIGHT_UNIT["РТГр"]}},
    100: {"general": [*MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR, *MONEY_REPORT_GENERAL_REFERENCES_BN],
          "БД(СЗ)": {**_FLAGS, "use_brs": True, "grounds": []}},
    "РТГр_100": {"РТГр": {**_FLAGS, "exclude_general": "all", "grounds": [{"start": "01.10.2026", "end": _NOW, "lines": ["БР ком РТГр"]}]}},
    "NOT_PAID": {"column_basis_from_tracking": "ПІДСТАВИ ДЛЯ НЕ ВИПЛАТИ", "СЗЧ": {**_FLAGS, "default": False, "grounds": []}},
}
COMMANDER_MONEY_REPORT_CATEGORIES = {30: {"general": [], "column_basis_from_tracking": None,
                                          "ЖИТТЄДІЯЛЬНІСТЬ": {**_FLAGS, "grounds": [], "column_basis_from_tracking": ("ПІДСТАВИ 10К",)}}}
HIGHER_COMMANDER_TITLE = "Командир батальйону"
COMMANDER_TITLE = "Начальник штабу"
SEQUENCE_HIGHER_COMMANDER = ["ЗКБ", "НШ-ЗКБ"]
'''


@pytest.fixture
def legacy_project(tmp_path):
    (tmp_path / "resources").mkdir()
    (tmp_path / "resources" / "data.json").write_text(json.dumps(_LEGACY_DATA, ensure_ascii=False), encoding="utf-8")
    return tmp_path


def test_legacy_settings_are_migrated_into_an_equivalent_workbook(legacy_project, capsys):
    settings_path = legacy_project / "resources" / sb.SETTINGS_FILE_NAME
    legacy_json = legacy_project / "resources" / "data.json"
    today = datetime.now().strftime("%d.%m.%Y")
    legacy = sm.load_legacy_constants(_LEGACY_SOURCE, str(legacy_project))

    assert sm.migrate_legacy_settings(str(settings_path), str(legacy_json), str(legacy_project), legacy_source=_LEGACY_SOURCE)

    # Книга збіглася з колишніми значеннями - data.json більше не потрібен.
    assert not legacy_json.exists()
    settings = ps.read_settings(str(settings_path))
    built = ps.build_constants(settings, today)
    assert built.MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR == legacy["MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR"]
    assert list(built.BR_HIGHT_UNIT.values()) == list(legacy["BR_HIGHT_UNIT"].values())
    assert built.MONEY_REPORT_CATEGORIES[30]["general"] == legacy["MONEY_REPORT_CATEGORIES"][30]["general"]
    assert built.MONEY_REPORT_CATEGORIES[100]["general"] == legacy["MONEY_REPORT_CATEGORIES"][100]["general"]
    assert built.MONEY_REPORT_CATEGORIES["РТГр_100"]["РТГр"] == {**legacy["MONEY_REPORT_CATEGORIES"]["РТГр_100"]["РТГр"]}
    assert built.MONEY_REPORT_CATEGORIES["NOT_PAID"]["column_basis_from_tracking"] == "ПІДСТАВИ ДЛЯ НЕ ВИПЛАТИ"
    assert built.COMMANDER_MONEY_REPORT_CATEGORIES[30]["ЖИТТЄДІЯЛЬНІСТЬ"]["column_basis_from_tracking"] == "ПІДСТАВИ 10К"
    assert (built.HIGHER_COMMANDER_TITLE, built.SEQUENCE_HIGHER_COMMANDER) == ("Командир батальйону", ["ЗКБ", "НШ-ЗКБ"])
    assert built.MONEY_REPORT_STATIC["REPORT_TITLE"] == "РАПОРТ"
    assert built.MONEY_REPORT_STATIC["SECTIONS"] == {"30": "текст 30", "NOT_PAID": "не платити"}
    assert built.MONEY_REPORT_STATIC["COMMANDER_SECTIONS"] == {"30": "мені 30"}
    assert built.TASK_ORDER_VARIANTS == [[("А", "завдання")], []]

    # Групи підстав: спільні списки - відомими групами, решта - нова група з назвою пункту.
    groups = list(dict.fromkeys(row["group"] for row in settings["references"]))
    assert groups == ["ЖБД", "БН", "Бригада загальні", "РТГр", "РТГр_100"]
    # Колишнє datetime.now() ("діє досі") - порожній кінець, а не сьогоднішня дата назавжди.
    assert [row["end"] for row in settings["references"]] == [None, "31.10.2026", None, None, None]

    day, week, save = ps.build_number_dicts(settings["numbers"], 11, 2026)
    assert day == {
        "31.10.2026": {"бат": "1", "брг": "2", "general_br_bat": ""},
        "01.11.2026": {"бат": "3", "брг": "4", "general_br_bat": "5 від 01.11.2026"},
    }
    assert settings["numbers"][_DAY][-1]["day"] == 31  # номер "31.11.2026" не загубився - лежить у рядку 31
    assert week == {"01.11.2026": {"бат": "5", "посилання_брг": "6 від 01.10.2026", "посилання_бат": ""}}
    assert save == {"01.11.2026": {"бз_бат": "7", "бз_брг": "8"}}
    assert "не все збіглося" not in capsys.readouterr().out


def test_migrated_subdivisions_keep_both_orders():
    """Реальний випадок: порядок абзаців БР (дшр, рвп, мінбатр) відрізняється від
    порядку рядка "КОМАНДИРУ ..." (дшр, мінбатр, рвп), а абзаци є й для тих,
    кого в переліку немає (гбс, "упр, штаб") - обидва порядки лишаються як були."""
    sections = {
        "SEQUENCE": ["дшр", "мінбатр", "рвп"],
        "PARAGRAPH_MAP": {"дшр": "д", "рвп": "р", "мінбатр": "м", "гбс": "г", "упр, штаб": "у"},
        "PARAGRAPH_MAP_FOR_EXTRACT": {"дшр": "д2", "рвп": "р2"},
    }

    built = ps.build_constants({**_sample_module().settings_example(), "subdivisions": sm._subdivision_rows(sections),
                                "subdivision_texts": sm._subdivision_text_rows(sections)}, _TODAY)

    assert built.SEQUENCE == ["дшр", "мінбатр", "рвп"]
    assert list(built.PARAGRAPH_MAP.items()) == [("дшр", "д"), ("рвп", "р"), ("мінбатр", "м"), ("гбс", "г"), ("упр, штаб", "у")]
    assert built.PARAGRAPH_MAP_FOR_EXTRACT == {"дшр": "д2", "рвп": "р2"}


def test_section_text_keys_follow_the_point_names():
    """Реальний випадок: пункт перейменували "100_РТГр" -> "РТГр_100" лише в
    constants.py, а текст у data.json лишився під старою назвою, і рапорт падав
    з KeyError. Перенесення записує текст під назвою пункту."""
    points = ["30", "РТГр_100", "NOT_PAID"]

    assert sm._section_texts_by_point({"30": "а", "100_РТГр": "б", "NOT_PAID": "в", "50": "г"}, points) == {
        "30": "а", "РТГр_100": "б", "NOT_PAID": "в", "50": "г",
    }
    # Під правильною назвою текст уже є - нічого не перезаписується.
    assert sm._section_texts_by_point({"100_РТГр": "старий", "РТГр_100": "новий"}, points) == {
        "100_РТГр": "старий", "РТГр_100": "новий",
    }


def test_migration_renames_a_section_text_to_its_point_and_still_removes_data_json(tmp_path, capsys):
    data = json.loads(json.dumps(_LEGACY_DATA, ensure_ascii=False))
    data["data_statik_report_for_get_additional_money"]["SECTIONS"]["100_РТГр"] = "текст РТГр"
    (tmp_path / "resources").mkdir()
    legacy_json = tmp_path / "resources" / "data.json"
    legacy_json.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    settings_path = tmp_path / "resources" / sb.SETTINGS_FILE_NAME

    assert sm.migrate_legacy_settings(str(settings_path), str(legacy_json), str(tmp_path), legacy_source=_LEGACY_SOURCE)

    built = ps.build_constants(ps.read_settings(str(settings_path)), _TODAY)
    assert built.MONEY_REPORT_STATIC["SECTIONS"]["РТГр_100"] == "текст РТГр"
    assert "100_РТГр" not in built.MONEY_REPORT_STATIC["SECTIONS"]
    assert 'Текст пункту "100_РТГр" перенесено як "РТГр_100"' in capsys.readouterr().out
    assert not legacy_json.exists()


def test_migrated_paragraph_takes_the_spelling_of_the_subdivision_list():
    rows = sm._subdivision_text_rows({"SEQUENCE": ["1 рмп"], "PARAGRAPH_MAP": {"1рмп": "один", "упр, штаб": "у"}})

    assert [row["name"] for row in rows] == ["1 рмп", "упр, штаб"]


def test_legacy_data_json_stays_when_the_workbook_differs(legacy_project, monkeypatch, capsys):
    """Якщо книга, прочитана назад, у чомусь не збіглася з колишніми значеннями,
    data.json лишається, а користувач бачить, які розділи перевірити."""
    legacy_json = legacy_project / "resources" / "data.json"
    real_build = sm.ps.build_constants

    def build_with_a_lost_reference(settings, today):
        built = real_build(settings, today)
        built.MONEY_REPORT_GENERAL_REFERENCES_BN = []
        return built

    monkeypatch.setattr(sm.ps, "build_constants", build_with_a_lost_reference)

    migrated = sm.migrate_legacy_settings(
        str(legacy_project / "resources" / sb.SETTINGS_FILE_NAME), str(legacy_json), str(legacy_project), legacy_source=_LEGACY_SOURCE,
    )

    assert migrated is False
    assert legacy_json.exists()
    assert "не все збіглося з колишніми: MONEY_REPORT_GENERAL_REFERENCES_BN" in capsys.readouterr().out


def _current_settings_modules():
    """Поточні utils.project_settings/utils.settings_migration: у спільній сесії
    кількох проєктів кореневий conftest.py вивантажує utils.* перед тестом, а
    load_settings імпортує settings_migration ліниво - патчити треба саме їх, а
    не модулі, імпортовані на початку цього файлу."""
    import utils.project_settings as current_ps
    import utils.settings_migration as current_sm

    return current_ps, current_sm


def test_migration_without_the_old_constants_explains_what_to_do(legacy_project, monkeypatch):
    current_ps, current_sm = _current_settings_modules()
    monkeypatch.setattr(current_sm, "find_legacy_constants_source", lambda _project_dir: None)

    with pytest.raises(SystemExit, match="не знайдено попередньої версії constants.py"):
        current_ps.load_settings(
            str(legacy_project / "resources" / sb.SETTINGS_FILE_NAME), str(legacy_project / "resources" / "data.json"),
            str(legacy_project),
        )
    assert not (legacy_project / "resources" / sb.SETTINGS_FILE_NAME).exists()


def test_load_settings_migrates_once_and_then_reads_the_workbook(legacy_project, monkeypatch, capsys):
    current_ps, current_sm = _current_settings_modules()
    monkeypatch.setattr(current_sm, "find_legacy_constants_source", lambda _project_dir: _LEGACY_SOURCE)
    settings_path = legacy_project / "resources" / sb.SETTINGS_FILE_NAME
    legacy_json = legacy_project / "resources" / "data.json"

    first = current_ps.load_settings(str(settings_path), str(legacy_json), str(legacy_project))
    assert "без втрат, data.json видалено" in capsys.readouterr().out
    assert not legacy_json.exists()

    second = current_ps.load_settings(str(settings_path), str(legacy_json), str(legacy_project))
    assert _without_row_numbers(second) == _without_row_numbers(first)


@pytest.mark.skipif(shutil.which("git") is None, reason="потрібен git")
def test_old_constants_are_found_in_git_history(tmp_path):
    def git(*args):
        subprocess.run(["git", "-c", "user.name=t", "-c", "user.email=t@t", "-C", str(tmp_path), *args], check=True, capture_output=True)

    git("init", "-q")
    (tmp_path / "constants.py").write_text("NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY = {\n}\n", encoding="utf-8")
    git("add", "constants.py")
    git("commit", "-q", "-m", "old")
    (tmp_path / "constants.py").write_text("from utils import project_settings\n", encoding="utf-8")
    git("commit", "-q", "-am", "new")

    assert sm.find_legacy_constants_source(str(tmp_path)) == "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY = {\n}\n"


# --- папка етапу: книга номерів документів ---------------------------------------------


def _stage_dicts():
    return (
        {"31.10.2026": {"бат": "1", "брг": "", "general_br_bat": ""}, "01.11.2026": {"бат": "2", "брг": "3", "general_br_bat": ""}},
        {"01.11.2026": {"бат": "4", "посилання_брг": "", "посилання_бат": ""}},
        {"01.11.2026": {"бз_бат": "5", "бз_брг": "6"}},
    )


def _run_stage_sync(stage_dir, settings_path, dicts, monkeypatch, answer=False, month_folder=None):
    import sync.stage_folder as stage_folder

    monkeypatch.setattr(stage_folder, "_ask_yes_no", lambda *_args: answer)
    day, week, save = dicts
    return stage_folder.sync_stage_constants_backup(
        str(stage_dir), "11", month_folder or str(stage_dir.parent), True, str(settings_path),
        list(day), list(week), list(save), day, week, save,
    )


def test_new_stage_gets_its_own_numbers_workbook(tmp_path, monkeypatch):
    stage_dir = tmp_path / "11" / "1 етап"
    stage_dir.mkdir(parents=True)
    dicts = _stage_dicts()

    path = _run_stage_sync(stage_dir, tmp_path / "settings.xlsx", dicts, monkeypatch)

    assert path == str(stage_dir / "НОМЕРИ ДОКУМЕНТІВ.xlsx")
    assert ps.read_numbers_workbook(path, {_DAY: list(dicts[0])})[_DAY] == dicts[0]


def test_existing_stage_numbers_replace_the_project_ones(tmp_path, monkeypatch):
    stage_dir = tmp_path / "10" / "1 етап"  # не поточний місяць - питання не ставиться
    stage_dir.mkdir(parents=True)
    stored = _stage_dicts()
    ps.write_numbers_workbook(str(stage_dir / "НОМЕРИ ДОКУМЕНТІВ.xlsx"), {_DAY: stored[0], _WEEK: stored[1], _SAVE: stored[2]})
    in_memory = tuple({key: {field: "" for field in entry} for key, entry in numbers.items()} for numbers in _stage_dicts())

    _run_stage_sync(stage_dir, tmp_path / "settings.xlsx", in_memory, monkeypatch)

    assert in_memory == stored


def test_legacy_stage_backup_is_carried_into_the_numbers_workbook(tmp_path, monkeypatch):
    stage_dir = tmp_path / "11" / "1 етап"
    stage_dir.mkdir(parents=True)
    (stage_dir / "constants.py.bak").write_text(
        "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY = {\n"
        "    f'01.{datetime.strptime(DATA, \"%m.%Y\").strftime(\"%m.%Y\")}': {'бат': '350', 'брг': '3293'},\n"
        "}\nNUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK = {\n}\nNUMBER_OF_DOCUMENTS_BRS_SAVE = {\n}\n",
        encoding="utf-8",
    )
    dicts = _stage_dicts()

    path = _run_stage_sync(stage_dir, tmp_path / "settings.xlsx", dicts, monkeypatch)

    assert dicts[0] == {"01.11.2026": {"бат": "350", "брг": "3293"}}
    assert dicts[1] == {} and dicts[2] == {}
    assert os.path.isfile(path)
