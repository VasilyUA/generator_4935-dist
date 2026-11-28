"""Книга налаштувань resources/НАЛАШТУВАННЯ.xlsx (utils/project_settings.py,
utils/settings_book.py) - на синтетичних книгах у tmp_path, без справжніх
resources/."""
import gc
import json
from datetime import datetime

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell

from utils import project_settings as ps
from utils import settings_book as sb

_FIRST = "ПЕРШИЙ Перший Перший"
_SECOND = "ДРУГИЙ Другий Другий"


def _data(**overrides):
    data = {
        "unit": {
            "SHORT_UNIT_BATTALION": "0 бат",
            "SHORT_UNIT_BRIGADE": "00 обрбр",
            "FULL_UNIT_BUT": "0 батальйону",
            "FULL_MILITARY_UNIT": "військової частини А0000",
        },
        "groups": {"SPECIAL_TASK_GROUP": "РТГР «Назва»", "BCHS_SIGNAL_GROUP_NAME": "Група БЧС"},
        "INFORMATION_UNIT_KEEP_FILE_NAME": "Управління.xlsx",
        "company_sections": {"SEQUENCE": ["альфа", "бета"]},
        "CONST_PAYMENT_CATEGORY_100_MAX_DAYS": 27,
        "EXCLUDE_NAMES": [
            {"name": _FIRST, "days": 30, "color": False},
            {"name": _SECOND, "days": None, "color": "FF0000"},
        ],
        "EXCLUDE_ABSENT_FROM_INFORMATION_UNIT": [_SECOND],
        "EXCLUDE_FROM_PAYMENTS_REVIEW": [{"date": "05.11.2026", "name": _FIRST}],
        "EXCLUDE_FROM_REVIEW": [{"date": "05.11.2026", "name": _FIRST}, {"date": "06.11.2026", "name": _SECOND}],
        "IGNORE_LISTS": {
            "BOOK_61": [{"period": "31.10.2026", "name": _FIRST}],
            "CVS": [],
            "COMBAT_TRAINING_JOURNAL": [],
            "TREATMENT_JOURNAL": [{"period": "31.10.2026-30.11.2026", "name": _SECOND}],
        },
    }
    data.update(overrides)
    return data


def _write(tmp_path, data=None):
    path = tmp_path / sb.SETTINGS_FILE_NAME
    ps.write_settings(str(path), data or _data())
    return path


def _locate(workbook, table, header):
    """(аркуш, {заголовок: колонка}, рядки даних) таблиці table книги - від
    смуги "■ table" на її аркуші до наступної смуги; рядок заголовків - перший
    рядок таблиці, де є заголовок header."""
    sheet, _label, first_row, last_row = sb.locate_table(workbook, ps.PLACE[table], table)
    last_row = last_row or sheet.max_row
    for row in range(first_row, last_row + 1):
        headers = {cell.value: cell.column for cell in sheet[row] if isinstance(cell.value, str)}
        if header in headers:
            return sheet, headers, range(row + 1, last_row + 1)
    raise AssertionError(f"у таблиці {table} немає заголовка {header}")


def _table_rows(workbook, table, header):
    """Заповнені рядки таблиці: значення колонок аркуша від A до останнього заголовка таблиці."""
    sheet, headers, rows = _locate(workbook, table, header)
    width = max(headers.values())
    values = [tuple(sheet.cell(row=row, column=column).value for column in range(1, width + 1)) for row in rows]
    return [row for row in values if any(value is not None for value in row)]


def _empty_row(sheet, headers, rows):
    """Перший порожній рядок таблиці (під даними - кілька порожніх рядків з рамками)."""
    return next(row for row in rows if all(sheet.cell(row=row, column=column).value is None for column in headers.values()))


def _set_cells(path, table, find_header, values, row_index=0):
    """Записує values ({заголовок: значення}) у рядок row_index даних таблиці
    table (None - у перший порожній рядок). Повертає номер рядка книги."""
    workbook = sb.load_editable(str(path))
    sheet, headers, rows = _locate(workbook, table, find_header)
    row = _empty_row(sheet, headers, rows) if row_index is None else rows[row_index]
    for header, value in values.items():
        sheet.cell(row=row, column=headers[header]).value = value
    sb.save_workbook(workbook, str(path))
    return row


def _write_former_layout(path, data=None, skip=(), hidden=()):
    """Книга колишнього вигляду: кожна таблиця - на власному аркуші з назвою
    таблиці, без смуг "■"; таблиць skip у книзі немає, аркуші hidden приховані."""
    workbook = sb.new_workbook()
    for table, section in ps._sections(data or _data()).items():
        if table in skip:
            continue
        if section.settings is not None:
            sheet = sb.write_settings_sheet(workbook, table, section.settings, section.values or {})
        else:
            sheet = sb.write_table(workbook, table, section.columns, section.rows, band_by=section.band_by)
        if table in hidden:
            sheet.sheet_state = "hidden"
    sb.save_workbook(workbook, str(path))
    return path


def test_settings_survive_a_round_trip_through_the_workbook(tmp_path):
    path = _write(tmp_path)

    assert ps.read_settings(str(path)) == ps.with_default_statuses(_data())


def test_optional_settings_round_trip_when_present(tmp_path):
    data = _data(BRIGADE_MILITARY_UNITS=["А0001", "А0002"])
    data["groups"]["REPORTS_SIGNAL_GROUP_NAME"] = "Група рапортів"

    assert ps.read_settings(str(_write(tmp_path, data))) == ps.with_default_statuses(data)


def test_review_pairs_of_both_files_share_one_row(tmp_path):
    rows = _table_rows(load_workbook(_write(tmp_path)), ps.REVIEW_SHEET, "ПІБ")

    assert rows == [
        (_FIRST, datetime(2026, 11, 5), "Так", "Так"),
        (_SECOND, datetime(2026, 11, 6), "Так", "Ні"),
    ]


def test_empty_review_flags_mean_both_files(tmp_path):
    path = _write(tmp_path)
    _set_cells(path, ps.REVIEW_SHEET, "ПІБ", {"ОБЛІК для виплат": None}, row_index=1)

    data = ps.read_settings(str(path))

    assert {"date": "06.11.2026", "name": _SECOND} in data["EXCLUDE_FROM_PAYMENTS_REVIEW"]


def test_columns_may_be_reordered_and_extra_columns_are_ignored(tmp_path):
    path = _write(tmp_path)
    workbook = sb.load_editable(str(path))
    sheet, _headers, rows = _locate(workbook, ps.CATEGORY_100_SHEET, "ПІБ")
    header_row = rows[0] - 1
    for row in (header_row, *rows):
        for column in range(1, 5):
            sheet.cell(row=row, column=column).value = None
    table = [
        ["Примітка", "Колір", "ПІБ", "Днів"],
        ["моя примітка", "так", _FIRST, 31],
        [None, None, None, None],
        ["рядок лише з приміткою - пропускається", None, None, None],
    ]
    for row, values in zip((header_row, *rows), table):
        for column, value in enumerate(values, start=1):
            sheet.cell(row=row, column=column).value = value
    workbook.save(path)

    assert ps.read_settings(str(path))["EXCLUDE_NAMES"] == [{"name": _FIRST, "days": 31, "color": True}]

    sheet.cell(row=rows[3], column=4).value = 30
    workbook.save(path)
    with pytest.raises(sb.SettingsError, match=f'рядок {rows[3]}, колонка "ПІБ": клітинка не може бути порожньою'):
        ps.read_settings(str(path))


def test_book_may_be_named_by_its_technical_key(tmp_path):
    path = _write(tmp_path)
    _set_cells(path, ps.IGNORE_SHEET, "Книга", {"Книга": "treatment_journal"})

    lists = ps.read_settings(str(path))["IGNORE_LISTS"]

    assert lists["TREATMENT_JOURNAL"][0] == {"period": "31.10.2026", "name": _FIRST}


def test_unknown_book_names_the_allowed_ones(tmp_path):
    path = _write(tmp_path)
    _set_cells(path, ps.IGNORE_SHEET, "Книга", {"Книга": "книга 62"})

    with pytest.raises(sb.SettingsError, match="невідома книга 'книга 62' - можна: book61, ЦВС"):
        ps.read_settings(str(path))


def test_bad_date_names_the_sheet_row_and_column(tmp_path):
    path = _write(tmp_path)
    row = _set_cells(path, ps.REVIEW_SHEET, "Дата", {"Дата": "06.092026"})

    with pytest.raises(sb.SettingsError) as error:
        ps.read_settings(str(path))

    assert str(error.value) == (
        f'НАЛАШТУВАННЯ.xlsx, аркуш "Люди", таблиця "Без ручної перевірки", рядок {row}, колонка "Дата": '
        "очікується дата дд.мм.рррр, а не '06.092026'."
    )


def test_missing_required_setting_is_named(tmp_path):
    path = _write(tmp_path)
    workbook = sb.load_editable(str(path))
    sheet, headers, rows = _locate(workbook, ps.GENERAL_SHEET, sb.KEY_HEADER)
    row = next(row for row in rows if sheet.cell(row=row, column=headers[sb.KEY_HEADER]).value == "SHORT_UNIT_BATTALION")
    for column in headers.values():
        sheet.cell(row=row, column=column).value = None
    workbook.save(path)

    with pytest.raises(sb.SettingsError, match=r'немає рядка з параметром "Батальйон \(скорочено\)" \(SHORT_UNIT_BATTALION\)'):
        ps.read_settings(str(path))


def test_limit_must_be_a_whole_number(tmp_path):
    path = _write(tmp_path)
    workbook = sb.load_editable(str(path))
    sheet, headers, rows = _locate(workbook, ps.GENERAL_SHEET, sb.KEY_HEADER)
    row = next(row for row in rows if sheet.cell(row=row, column=headers[sb.KEY_HEADER]).value == "CONST_PAYMENT_CATEGORY_100_MAX_DAYS")
    sheet.cell(row=row, column=headers[sb.VALUE_HEADER]).value = "двадцять"
    workbook.save(path)

    with pytest.raises(sb.SettingsError, match="CONST_PAYMENT_CATEGORY_100_MAX_DAYS.*очікується ціле число"):
        ps.read_settings(str(path))


def test_list_sheets_may_be_missing_but_general_may_not(tmp_path):
    path = _write(tmp_path)
    workbook = sb.load_editable(str(path))
    sheet, _headers, rows = _locate(workbook, ps.ABSENT_SHEET, "ПІБ")  # остання таблиця аркуша "Люди"
    band_row = next(row for row in range(1, rows[0]) if sheet.cell(row=row, column=1).value == sb.section_band_text(ps.ABSENT_SHEET))
    for row in sheet.iter_rows(min_row=band_row, max_row=sheet.max_row):
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None
    workbook.save(path)

    assert ps.read_settings(str(path))["EXCLUDE_ABSENT_FROM_INFORMATION_UNIT"] == []

    del workbook[ps.GENERAL_BOOK_SHEET]
    workbook.save(path)
    with pytest.raises(sb.SettingsError, match='немає таблиці "Загальні"'):
        ps.read_settings(str(path))


def test_subdivisions_of_a_workbook_are_read_in_row_order(tmp_path):
    """Так табель читає й перелік з книги generator_br_and_report_for_money
    (той самий аркуш "Підрозділи"), а зайві колонки там не заважають."""
    path = tmp_path / "money.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = ps.SUBDIVISIONS_SHEET
    sheet.append(["Підрозділ", "Примітка"])
    sheet.append(["1 рмп", "перша"])
    sheet.append([None, None])
    sheet.append(["2 рмп", None])
    workbook.save(path)

    assert ps.read_subdivisions_file(str(path)) == ["1 рмп", "2 рмп"]


def test_subdivisions_table_is_found_on_any_sheet_of_the_money_workbook(tmp_path):
    """Теперішня книга generator_br_and_report_for_money: таблиця "■ Підрозділи"
    - друга на аркуші "Загальні" (під таблицею параметрів)."""
    path = tmp_path / "money.xlsx"
    workbook = sb.new_workbook()
    sb.write_sections(workbook, "Загальні", [
        sb.Section("Загальні", settings=(sb.Setting("NAME", label="Назва"),), values={"NAME": "Підрозділ"}),
        sb.Section(ps.SUBDIVISIONS_SHEET, columns=(sb.Column("name", "Підрозділ"),), rows=[{"name": "1 рмп"}, {"name": "2 рмп"}]),
    ], title="Загальні")
    sb.save_workbook(workbook, str(path))

    assert ps.read_subdivisions_file(str(path)) == ["1 рмп", "2 рмп"]


def test_load_settings_migrates_legacy_data_json_once_and_removes_it(tmp_path, capsys):
    legacy = tmp_path / "data.json"
    legacy.write_text(json.dumps(_data(), ensure_ascii=False), encoding="utf-8")
    settings_path = tmp_path / sb.SETTINGS_FILE_NAME

    assert ps.load_settings(str(settings_path), str(legacy)) == ps.with_default_statuses(_data())
    assert settings_path.exists()
    assert not legacy.exists()  # книга збіглася з data.json - він більше не потрібен
    assert "без втрат, data.json видалено" in capsys.readouterr().out

    assert ps.load_settings(str(settings_path), str(legacy)) == ps.with_default_statuses(_data())


def test_legacy_data_json_stays_when_the_workbook_differs(tmp_path, monkeypatch, capsys):
    """Якщо книга, прочитана назад, у чомусь не збіглася з data.json, він
    лишається, а користувач бачить, які розділи перевірити."""
    legacy = tmp_path / "data.json"
    legacy.write_text(json.dumps(_data(), ensure_ascii=False), encoding="utf-8")
    real_read = ps.read_settings
    monkeypatch.setattr(ps, "read_settings", lambda path: {**real_read(path), "EXCLUDE_NAMES": []})

    ps.load_settings(str(tmp_path / sb.SETTINGS_FILE_NAME), str(legacy))

    assert legacy.exists()
    assert "не все збіглося: EXCLUDE_NAMES" in capsys.readouterr().out


def test_unused_and_equivalent_legacy_values_are_not_differences():
    data = _data(company_sections={"SEQUENCE": ["альфа", "бета"], "MANAGEMENT": ["упр"]})
    data["groups"]["REPORTS_SIGNAL_GROUP_NAME"] = ""
    data["EXCLUDE_NAMES"][0]["color"] = None
    data["IGNORE_LISTS"]["TREATMENT_JOURNAL"][0]["period"] = "31.10.2026 - 30.11.2026"

    assert ps.migration_differences(data, ps.with_default_statuses(_data())) == []


def test_load_settings_without_workbook_or_legacy_file_explains_what_to_do(tmp_path):
    with pytest.raises(SystemExit, match="generate_sample.py"):
        ps.load_settings(str(tmp_path / sb.SETTINGS_FILE_NAME), str(tmp_path / "data.json"))


def test_load_settings_turns_a_workbook_error_into_a_readable_exit(tmp_path):
    path = _write(tmp_path)
    workbook = load_workbook(path)
    del workbook[ps.GENERAL_BOOK_SHEET]
    workbook.save(path)

    with pytest.raises(SystemExit, match='У книзі налаштувань помилка:\n.*немає таблиці "Загальні"'):
        ps.load_settings(str(path))


def _red_lines(output):
    """Рядки, надруковані червоним (sb.stop): усе між кодом червоного й
    скиданням кольору."""
    start, end = "\x1b[31m", "\x1b[0m"
    assert start in output, output
    return output.split(start, 1)[1].split(end, 1)[0].splitlines()


def test_settings_stop_ends_with_code_1_and_its_text_is_printed_once_in_red(capsys):
    """Реальний випадок: SystemExit з текстом Python друкував білим, а
    червоним лаунчер писав лише "Програма завершилась з помилкою" - людина,
    яка читає тільки червоне, не знала, що робити. Тепер пояснення з кроками -
    червоним, а сам SystemExit лише з кодом 1 (Python тексту вдруге не друкує)."""
    stop = sb.stop("Щось не так.", ("Перший крок.", "Другий крок."))

    assert isinstance(stop, SystemExit) and stop.code == 1
    assert _red_lines(capsys.readouterr().out) == [
        "ПОМИЛКА В НАЛАШТУВАННЯХ - програму зупинено.", "Щось не так.", "Що зробити:",
        "  1. Перший крок.", "  2. Другий крок.",
    ]
    assert str(stop).splitlines()[1] == "Щось не так."


def test_old_data_json_without_a_required_value_says_what_to_fill_in_and_where(tmp_path, capsys):
    """Реальний випадок: у старому data.json не було "Ліміт днів категорії
    100" - книгу записано, а прочитати не вдалося. Червоним: що це за
    параметр, що вписати, де (аркуш, рядок, колонка) і кроки. data.json
    лишається; наступний запуск каже те саме, доки значення не вписано, а
    після цього все працює."""
    legacy = tmp_path / "data.json"
    old = _data()
    del old["CONST_PAYMENT_CATEGORY_100_MAX_DAYS"]
    legacy.write_text(json.dumps(old, ensure_ascii=False), encoding="utf-8")
    settings_path = tmp_path / sb.SETTINGS_FILE_NAME

    with pytest.raises(SystemExit) as stopped:
        ps.load_settings(str(settings_path), str(legacy))

    assert stopped.value.code == 1
    red = "\n".join(_red_lines(capsys.readouterr().out))
    for expected in ("перенесено", 'аркуш "Загальні"', 'колонка "Значення"', '"Ліміт днів категорії 100"',
                     "Що вписати: Скільки днів поспіль", f"1. Відкрийте {settings_path} в Excel.",
                     "3. Збережіть файл, закрийте Excel і запустіть програму ще раз."):
        assert expected in red
    assert legacy.exists() and settings_path.exists()

    with pytest.raises(SystemExit, match='У книзі налаштувань помилка:\n.*"Ліміт днів категорії 100"'):
        ps.load_settings(str(settings_path), str(legacy))

    ps.write_settings(str(settings_path), _data())  # користувач вписав значення
    assert ps.load_settings(str(settings_path), str(legacy))["CONST_PAYMENT_CATEGORY_100_MAX_DAYS"] == 27


def test_unreadable_data_json_says_how_to_fix_it(tmp_path, capsys):
    legacy = tmp_path / "data.json"
    legacy.write_text("{не json", encoding="utf-8")

    with pytest.raises(SystemExit, match="Не вдалося перенести"):
        ps.load_settings(str(tmp_path / sb.SETTINGS_FILE_NAME), str(legacy))

    assert any("Виправте" in line and "generate_sample.py" in line for line in _red_lines(capsys.readouterr().out))


def test_invalid_value_found_by_constants_is_explained_in_red(tmp_path, capsys):
    """Помилку значення, яку знаходить constants.py під час імпорту (текст
    починається з назви книги), - червоним з кроками, а не трасуванням.
    Інші ValueError - як були."""
    settings_path = tmp_path / sb.SETTINGS_FILE_NAME
    with pytest.raises(SystemExit) as stopped:
        with ps.stop_on_invalid_settings(str(settings_path)):
            raise ValueError(f"{ps.source_of('IGNORE_LISTS')}: період '31.02.2026' - очікується \"дд.мм.рррр\".")

    assert stopped.value.code == 1
    red = _red_lines(capsys.readouterr().out)
    assert red[1] == "У книзі налаштувань помилка:" and red[2].startswith(sb.SETTINGS_FILE_NAME)
    assert f"  1. Відкрийте {settings_path} в Excel." in red

    with pytest.raises(ValueError, match="інша помилка"):
        with ps.stop_on_invalid_settings(str(settings_path)):
            raise ValueError("інша помилка")


def test_source_of_names_the_sheet_and_keeps_the_key():
    assert ps.source_of("EXCLUDE_NAMES") == (
        'НАЛАШТУВАННЯ.xlsx, аркуш "Люди", таблиця "Ліміт категорії 100" (EXCLUDE_NAMES)'
    )
    assert ps.source_of("SYNTHETIC") == "НАЛАШТУВАННЯ.xlsx (SYNTHETIC)"


def test_book_labels_match_the_kinds_known_to_constants():
    import constants

    assert [key for key, _label in ps.BOOKS] == list(constants.ADDON_IGNORE_LIST_KEYS)
    assert [label for _key, label in ps.BOOKS] == [
        constants.ADDON_BOOK_KINDS[kind] for kind in constants.ADDON_IGNORE_LIST_KEYS.values()
    ]


def test_workbook_has_titles_instructions_and_colored_tabs(tmp_path):
    workbook = load_workbook(_write(tmp_path))
    instructions = workbook[sb.INSTRUCTIONS_SHEET]
    links = {cell.value: cell.hyperlink.location for row in instructions.iter_rows() for cell in row if cell.hyperlink}

    assert workbook.sheetnames == [sb.INSTRUCTIONS_SHEET, *(sheet for sheet, *_rest in ps.BOOK_SHEETS)]
    assert len(workbook.sheetnames) <= 7
    assert all(sheet.sheet_state == "visible" for sheet in workbook.worksheets)
    assert all(workbook[name].sheet_properties.tabColor is not None for name in workbook.sheetnames)
    # Кожна таблиця - смуга "■ назва" на своєму аркуші, посилання з "Інструкції" веде саме на неї.
    for sheet_name, table in ps.BOOK_TABLES:
        sheet = workbook[sheet_name]
        band_row = next(row for row in range(1, sheet.max_row + 1) if sheet.cell(row=row, column=1).value == sb.section_band_text(table))
        assert links[f"{sheet_name} > {table}"] == f"'{sheet_name}'!A{band_row}"
    people = workbook[ps.PEOPLE_BOOK_SHEET]
    assert people.cell(row=sb.TITLE_ROW, column=1).value.startswith("Люди")
    review, headers, _rows = _locate(workbook, ps.REVIEW_SHEET, "ПІБ")
    assert sorted(headers, key=headers.get) == [column.header for column in ps.REVIEW_COLUMNS]
    general, headers, rows = _locate(workbook, ps.GENERAL_SHEET, sb.KEY_HEADER)
    labels = {general.cell(row=row, column=headers[sb.KEY_HEADER]).value: general.cell(row=row, column=headers[sb.LABEL_HEADER]).value
              for row in rows}
    assert labels["SHORT_UNIT_BATTALION"] == "Батальйон (скорочено)"


def test_tables_of_one_sheet_line_up_their_columns(tmp_path):
    """У таблиць аркуша спільні колонки аркуша: "ПІБ" - скрізь перша, тож
    вузькі колонки не розтягуються; широка колонка займає кілька колонок
    аркуша (об'єднані клітинки) і не розширює вузьку колонку сусідньої таблиці."""
    from utils import status_rules as sr

    workbook = load_workbook(_write(tmp_path))
    for table in (ps.REVIEW_SHEET, ps.IGNORE_SHEET, ps.CATEGORY_100_SHEET, ps.ABSENT_SHEET):
        assert _locate(workbook, table, "ПІБ")[1]["ПІБ"] == 1
    people = workbook[ps.PEOPLE_BOOK_SHEET]
    assert people.column_dimensions["B"].width < 25  # "Дата", "Книга", "Днів" - не ширина ПІБ
    reports = workbook[sr.REPORTS_BOOK_SHEET]
    _sheet, headers, rows = _locate(workbook, sr.SENTENCES_SHEET, "Шаблон речення")
    assert (headers["Шаблон речення"], headers["Статус"]) == (1, 5)  # шаблон - на чотири колонки аркуша
    assert f"A{rows[0]}:D{rows[0]}" in {str(merged) for merged in reports.merged_cells.ranges}
    assert reports.column_dimensions["A"].width < 30  # колонка "Статус" таблиці конфліктів - не 86


def test_color_cells_show_the_color_itself(tmp_path):
    sheet, headers, rows = _locate(load_workbook(_write(tmp_path)), ps.CATEGORY_100_SHEET, "ПІБ")
    cell = sheet.cell(row=rows[1], column=headers["Колір"])

    assert cell.value == "FF0000"
    assert cell.fill.start_color.rgb.endswith("FF0000")


def test_header_row_is_found_after_rows_were_inserted_above_it(tmp_path):
    """Користувач вставив рядок над заголовками таблиці й щось у ньому написав -
    заголовки шукаються самі."""
    path = _write(tmp_path)
    workbook = sb.load_editable(str(path))
    sheet, _headers, rows = _locate(workbook, ps.ABSENT_SHEET, "ПІБ")
    sheet.insert_rows(rows[0] - 1)
    sheet.cell(row=rows[0] - 1, column=1).value = "мої нотатки"
    workbook.save(path)

    assert ps.read_settings(str(path))["EXCLUDE_ABSENT_FROM_INFORMATION_UNIT"] == [_SECOND]


def test_old_layout_with_headers_in_the_first_row_is_still_read(tmp_path):
    """Книга, записана до оформлення (аркуш на таблицю, заголовки - у першому рядку), читається так само."""
    path = _write_former_layout(tmp_path / sb.SETTINGS_FILE_NAME)
    workbook = load_workbook(path)
    sheet = workbook[ps.SUBDIVISIONS_SHEET]
    for merged in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged))
    sheet.delete_rows(1, sb.HEADER_ROW - 1)
    workbook.save(path)

    assert ps.read_settings(str(path))["company_sections"] == {"SEQUENCE": ["альфа", "бета"]}


# --- статуси: значення, колір (заливка клітинки) і виплата ------------------------


def _statuses(path):
    return {row["status"]: row for row in ps.read_settings(str(path))["STATUSES"]}


def test_status_sheets_hold_the_former_constants_exactly(tmp_path):
    """Аркуші "Статуси", "Заміни для виплат", "Пріоритет статусів" нової книги дають
    рівно ті словники, що колись жили в constants.py (utils/default_statuses.py)."""
    from utils import default_statuses as ds

    tables = ps.status_tables(ps.read_settings(str(_write(tmp_path))))

    assert tables["STATUS_COLORS"] == ds.STATUS_COLORS
    assert tables["PAYMENT_STATUS_CONVERSIONS"] == ds.PAYMENT_STATUS_CONVERSIONS
    assert tables["PAYMENT_STATUS_OVERRIDES"] == ds.PAYMENT_STATUS_OVERRIDES
    assert tables["PAYMENT_STATUS_FALLBACKS"] == ds.PAYMENT_STATUS_FALLBACKS
    assert tables["PAYMENT_STATUS_PRIORITY"] == ds.PAYMENT_STATUS_PRIORITY
    assert all(tables["STATUS_MEANINGS"][status] == meaning for status, meaning in ds.STATUS_MEANINGS.items())
    assert 100 in tables["STATUS_COLORS"] and "100" not in tables["STATUS_COLORS"]  # категорія - числом, як у клітинці


def test_default_statuses_use_the_same_anchor_statuses_as_constants():
    import constants
    from utils import default_statuses as ds

    assert (ds.DEFAULT_STATUS, ds.TRANSFER_TO_BRIGADE_STATUS, ds.TRANSFER_TO_OTHER_UNIT_STATUS) == (
        constants.DEFAULT_STATUS, constants.TRANSFER_TO_BRIGADE_STATUS, constants.TRANSFER_TO_OTHER_UNIT_STATUS,
    )


def test_status_color_is_the_fill_the_user_picked(tmp_path):
    """Колір статусу - сама заливка клітинки "Колір": і код RGB, і колір теми з
    відтінком (верхні рядки палітри Excel), і вписаний текстом код."""
    from openpyxl.styles import PatternFill
    from openpyxl.styles.colors import Color

    path = _write(tmp_path)
    workbook = load_workbook(path)
    sheet, headers, data_rows = _locate(workbook, ps.STATUSES_SHEET, "Статус")
    status_column, color_column = headers["Статус"], headers["Колір"]
    rows = {sheet.cell(row=row, column=status_column).value: row for row in data_rows}
    sheet.cell(row=rows["ВП"], column=color_column).fill = PatternFill("solid", start_color="FF00AA11", end_color="FF00AA11")
    sheet.cell(row=rows["ШП"], column=color_column).fill = PatternFill("solid", fgColor=Color(theme=5, tint=0.5999938962981048))
    sheet.cell(row=rows["СЗЧ"], column=color_column).fill = PatternFill(fill_type=None)
    sheet.cell(row=rows["ВД"], column=color_column).fill = PatternFill(fill_type=None)
    sheet.cell(row=rows["ВД"], column=color_column).value = "#ffc000"
    workbook.save(path)

    statuses = _statuses(path)

    assert statuses["ВП"]["color"] == "FF00AA11"
    assert statuses["ШП"]["color"] == "FFE6B9B8"  # accent2 (C0504D) теми книги, світліше на 60%
    assert statuses["СЗЧ"]["color"] is None  # без заливки - клітинка в ОБЛІК лишається білою
    assert statuses["ВД"]["color"] == "FFFFC000"
    assert ps.status_tables(ps.read_settings(str(path)))["STATUS_COLORS"]["ВП"] == "FF00AA11"


def test_user_changes_payment_and_adds_a_status(tmp_path):
    path = _write(tmp_path)
    workbook = sb.load_editable(str(path))
    sheet, headers, rows = _locate(workbook, ps.STATUSES_SHEET, "Статус")
    row = next(row for row in rows if sheet.cell(row=row, column=headers["Статус"]).value == "ППД")
    sheet.cell(row=row, column=headers["Записати завжди"]).value = 30
    new_row = _empty_row(sheet, headers, rows)
    sheet.cell(row=new_row, column=headers["Статус"]).value = "НОВИЙ"
    sheet.cell(row=new_row, column=headers["Записати завжди"]).value = "10_МЕД"
    workbook.save(path)

    tables = ps.status_tables(ps.read_settings(str(path)))

    assert tables["PAYMENT_STATUS_CONVERSIONS"]["ППД"] == 30
    assert tables["PAYMENT_STATUS_CONVERSIONS"]["НОВИЙ"] == "10_МЕД"


def test_duplicate_status_is_an_error(tmp_path):
    path = _write(tmp_path)
    _set_cells(path, ps.STATUSES_SHEET, "Статус", {"Статус": "вп"}, row_index=None)

    with pytest.raises(sb.SettingsError, match='аркуш "Статуси", таблиця "Статуси", рядок \\d+, колонка "Статус": такий рядок уже є'):
        ps.read_settings(str(path))


def test_priority_winner_must_be_one_of_the_pair(tmp_path):
    path = _write(tmp_path)
    _set_cells(path, ps.PRIORITY_SHEET, "Перемагає", {"Перемагає": "ВП"})

    with pytest.raises(sb.SettingsError, match="має бути один з двох статусів рядка"):
        ps.read_settings(str(path))


def test_workbook_without_status_sheets_gets_them_with_the_defaults(tmp_path, capsys):
    """Книга, створена до появи таблиць статусів (і до кількох таблиць на
    аркуші): load_settings один раз переписує її в теперішній вигляд - ті самі
    дані й стандартні статуси; копія колишньої книги лишається поруч."""
    path = _write_former_layout(tmp_path / sb.SETTINGS_FILE_NAME,
                                skip=(ps.STATUSES_SHEET, ps.OVERRIDES_SHEET, ps.PRIORITY_SHEET))
    assert ps.read_settings(str(path))["STATUSES"] is None

    data = ps.load_settings(str(path))

    assert data == ps.with_default_statuses(_data())
    assert "переведено на новий вигляд" in capsys.readouterr().out
    assert load_workbook(path).sheetnames == [sb.INSTRUCTIONS_SHEET, *(sheet for sheet, *_rest in ps.BOOK_SHEETS)]
    assert ps.read_settings(str(path)) == data
    assert ps.STATUSES_SHEET not in load_workbook(tmp_path / "НАЛАШТУВАННЯ (до оновлення).xlsx").sheetnames

    ps.load_settings(str(path))  # уже в теперішньому вигляді - більше не переписується

    assert "новий вигляд" not in capsys.readouterr().out


def test_status_sheets_stay_missing_without_harm_when_the_workbook_is_locked(tmp_path, monkeypatch, capsys):
    path = _write_former_layout(tmp_path / sb.SETTINGS_FILE_NAME, skip=(ps.PRIORITY_SHEET,))
    before = path.read_bytes()

    def locked(*_args):
        raise sb.SettingsError("файл відкритий в Excel")

    monkeypatch.setattr(ps, "write_settings", locked)

    data = ps.load_settings(str(path))

    assert data["PAYMENT_STATUS_PRIORITY"] == ps.with_default_statuses({})["PAYMENT_STATUS_PRIORITY"]
    assert "спробую при наступному запуску" in capsys.readouterr().out
    assert path.read_bytes() == before


def _rename_status(path, code, name):
    workbook = sb.load_editable(str(path))
    sheet, headers, rows = _locate(workbook, ps.STATUSES_SHEET, "Код")
    row = next(row for row in rows if sb.cell_text(sheet.cell(row=row, column=headers["Код"]).value) == str(code))
    sheet.cell(row=row, column=headers["Статус"]).value = name
    sb.save_workbook(workbook, str(path))
    return row


def test_renamed_status_keeps_its_code_for_the_rules(tmp_path):
    """Користувач перейменував статус: правила програми (кольори, виплати,
    заміни) далі знають його за кодом, а назва - лише в STATUS_NAMES."""
    path = _write(tmp_path)
    _rename_status(path, "ЛМР", "Медрота")
    _rename_status(path, "10_МЕД", "10 мед")

    tables = ps.status_tables(ps.read_settings(str(path)))

    assert tables["STATUS_NAMES"] == {"ЛМР": "Медрота", "10_МЕД": "10 мед"}
    assert "ЛМР" in tables["STATUS_COLORS"] and "Медрота" not in tables["STATUS_COLORS"]
    assert tables["PAYMENT_STATUS_CONVERSIONS"]["ЛМР"] == "10_МЕД"
    assert tables["PAYMENT_STATUS_OVERRIDES"]["РВЗ"]["ЛМР"] == "10_МЕД"


def test_status_may_be_named_by_its_new_name_on_the_other_sheets(tmp_path):
    path = _write(tmp_path)
    _rename_status(path, "ЗВФН", "Без навантажень")
    _set_cells(path, ps.OVERRIDES_SHEET, "Записати",
               {"Статус в ОБЛІК": "ВД", "Подає підрозділ": "без навантажень", "Записати": "Без навантажень"}, row_index=None)

    overrides = ps.status_tables(ps.read_settings(str(path)))["PAYMENT_STATUS_OVERRIDES"]

    assert overrides["ВД"]["ЗВФН"] == "ЗВФН"


def test_status_name_may_not_be_another_status_code(tmp_path):
    path = _write(tmp_path)
    row = _rename_status(path, "ВПС", "ВП")

    with pytest.raises(sb.SettingsError, match=f'рядок {row}, колонка "Статус": назва "ВП" - це код статусу з рядка'):
        ps.read_settings(str(path))


def test_payment_category_numbers_are_not_renamed(tmp_path):
    path = _write(tmp_path)
    _rename_status(path, 100, "сто")

    with pytest.raises(sb.SettingsError, match="категорію виплати 100 не перейменовують"):
        ps.read_settings(str(path))


def test_status_rows_added_without_a_code_use_their_name(tmp_path):
    path = _write(tmp_path)
    _set_cells(path, ps.STATUSES_SHEET, "Статус", {"Статус": "НОВИЙ"}, row_index=None)

    statuses = _statuses(path)

    assert statuses["НОВИЙ"]["code"] == "НОВИЙ"
    assert ps.status_tables(ps.read_settings(str(path)))["STATUS_NAMES"] == {}


def _use_status_names(monkeypatch, names):
    """Перейменування для модулів табеля - як після constants.STATUS_NAMES з книги."""
    import content.status_names as status_names

    monkeypatch.setattr(status_names, "_DISPLAY", dict(names))
    monkeypatch.setattr(status_names, "_CODES", {status_names._norm(name): code for code, name in names.items()})
    return status_names


def test_status_names_translate_both_ways(monkeypatch):
    status_names = _use_status_names(monkeypatch, {"ЛМР": "Медрота", "ВП": "Відп"})

    assert [status_names.to_code(value) for value in ("медрота ", "ЛМР", "ВП", "Відп", 100, None)] == [
        "ЛМР", "ЛМР", "ВП", "ВП", 100, None,
    ]
    assert [status_names.to_display(value) for value in ("ЛМР", "ВП/ВПС", "ШП", 30)] == ["Медрота", "Відп/ВПС", "ШП", 30]


def test_timesheet_works_with_codes_and_writes_the_new_names(tmp_path, monkeypatch):
    """ОБЛІК.xlsx з новою назвою статусу: у пам'яті - код (на нього посилаються
    правила), у збереженому файлі - знову назва."""
    _use_status_names(monkeypatch, {"ВП": "Відп"})
    from content.oblik_timesheet import Timesheet

    source = tmp_path / "ОБЛІК.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Табель"
    sheet.append(["ПІБ", datetime(2026, 11, 1), datetime(2026, 11, 2), datetime(2026, 11, 3)])
    sheet.append([_FIRST, "Відп", "ВП", "ШП"])
    workbook.save(source)

    timesheet = Timesheet(str(source), "Табель")
    assert [timesheet.ws.cell(row=2, column=col).value for col, _day in timesheet.date_columns] == ["ВП", "ВП", "ШП"]

    output = tmp_path / "out.xlsx"
    timesheet.save(str(output))

    saved = load_workbook(output)["Табель"]
    assert [saved.cell(row=2, column=col).value for col in (2, 3, 4)] == ["Відп", "Відп", "ШП"]
    assert timesheet.ws.cell(row=2, column=2).value == "ВП"  # у пам'яті - знову код


def test_information_unit_values_with_the_new_name_become_codes(monkeypatch):
    _use_status_names(monkeypatch, {"ЛМР": "Медрота"})
    from content.information_unit_reader import _clean_cell_value

    assert [_clean_cell_value(value) for value in (" Медрота ", "ЛМР", "100", "ШП")] == ["ЛМР", "ЛМР", 100, "ШП"]


# --- правила статусів (колишні константи constants.py) ----------------------------


def _rules(path, special_group="група"):
    return ps.status_rules(ps.read_settings(str(path)), special_group, "<маркер>")


def test_rule_sheets_hold_the_former_constants_exactly(tmp_path):
    """Аркуші правил нової книги дають рівно ті значення, що колись були в
    constants.py (тепер - стандартні значення в utils/default_statuses.py)."""
    from utils import default_statuses as ds
    from utils import status_rules as sr

    rules = _rules(_write(tmp_path))

    for name in ("DEFAULT_STATUS", "AFTER_HOURS_ARRIVAL_STATUS", "VLK_STATUS", "ROZP_STATUS", "ADDON_LONG_TREATMENT_STATUS",
                 "TRANSFER_TO_BRIGADE_STATUS", "TRANSFER_TO_OTHER_UNIT_STATUS", "LEGACY_TRANSFER_STATUS", "LEAVE_STATUSES",
                 "STATUSES_OVER_ROZP", "PAYMENT_DUTY_CATEGORIES", "PAYMENT_MISMATCH_COMPARABLE_STATUSES",
                 "PAYMENT_ABSENT_FROM_FILES_EXEMPT_STATUSES", "ADDON_NOT_RETURNED_AFTER_SZCH_STATUSES",
                 "ADDON_ARRIVAL_DAY_STATUSES", "ADDON_ARRIVAL_DAY_SHEETS", "ADDON_NON_COMBAT_TREATMENT_STATUS",
                 "ADDON_KILLED_WITHOUT_CERTIFICATE_STATUS", "ADDON_TREATMENT_JOURNAL_STATUS", "ADDON_ONE_OF_STATUSES",
                 "STATUS_CONFLICT_PAID_BY_DEFAULT", "STATUS_CONFLICT_PRIORITY_EXEMPT", "IGNORED_TABLE_HEADING_KEYWORDS",
                 "ADDON_REQUIRED_BOOK61_SHEETS", "ADDON_TRACKED_STATUSES", "ADDON_TABLE_PRIORITY_STATUSES",
                 "ADDON_NESTED_STATUSES", "ADDON_TRAINING_JOURNAL_SHEETS"):
        assert rules[name] == getattr(ds, name), name
    assert rules["STATUS_HEADING_RULES"] == {
        keyword.replace("{спецгрупа}", "група"): "<маркер>" if status == ds.TRANSFER_HEADING_MARKER else status
        for keyword, status in ds.STATUS_HEADING_RULES.items()
    }
    assert list(rules["STATUS_HEADING_RULES"]) == [keyword.replace("{спецгрупа}", "група") for keyword in ds.STATUS_HEADING_RULES]
    assert rules["STATUS_CONFLICT_OVERRIDES"] == {frozenset(pair): winner for pair, winner in ds.STATUS_CONFLICT_OVERRIDES.items()}
    assert rules["ADDON_EQUIVALENT_STATUS_GROUPS"] == tuple(frozenset(group) for group in ds.ADDON_EQUIVALENT_STATUS_GROUPS)
    assert [(pattern.pattern, status) for pattern, status in rules["FREEFORM_STATUS_PATTERNS"]] == [
        (sr.compile_sentence(pattern, date_prefix).pattern, status) for (pattern, date_prefix), status in ds.FREEFORM_STATUS_PATTERNS
    ]


def test_rule_values_match_the_names_constants_keep_in_code():
    """Значення, які лишились у constants.py (не статуси), - ті самі, що й у
    стандартних правилах."""
    import constants
    from utils import default_statuses as ds
    from utils import status_rules as sr

    assert ds.ADDON_JOURNAL_OWN_UNIT == constants.ADDON_JOURNAL_OWN_UNIT
    assert sr.ADDON_BOOK_LABELS == constants.ADDON_BOOK_KINDS
    assert (sr.TRANSFER_HEADING, sr.IGNORED_HEADING, sr.SPECIAL_GROUP_PLACEHOLDER) == (
        ds.TRANSFER_HEADING_MARKER, ds.IGNORED_HEADING_MARKER, ds._SPECIAL_GROUP_NAME,
    )


def _set_group(path, key, value):
    from utils import status_rules as sr

    workbook = sb.load_editable(str(path))
    sheet, headers, rows = _locate(workbook, sr.GROUPS_SHEET, sb.KEY_HEADER)
    row = next(row for row in rows if sheet.cell(row=row, column=headers[sb.KEY_HEADER]).value == key)
    sheet.cell(row=row, column=headers[sb.VALUE_HEADER]).value = value
    sb.save_workbook(workbook, str(path))
    return row


def test_rules_follow_the_workbook_and_understand_new_status_names(tmp_path):
    from utils import status_rules as sr

    path = _write(tmp_path)
    _rename_status(path, "ВП", "Відпустка")
    _set_group(path, "LEAVE_STATUSES", "Відпустка, ВПС, ВПСЗ")
    _set_cells(path, sr.HEADINGS_SHEET, "Слова в заголовку",
               {"Слова в заголовку": "у декретну відпустку", "Статус": "відпустка"}, row_index=None)

    rules = _rules(path)

    assert rules["LEAVE_STATUSES"] == ("ВП", "ВПС", "ВПСЗ")  # нова назва - той самий код
    assert rules["STATUS_HEADING_RULES"]["у декретну відпустку"] == "ВП"


def test_single_status_group_must_have_exactly_one_status(tmp_path):
    path = _write(tmp_path)
    row = _set_group(path, "DEFAULT_STATUS", "РВЗ, ВП")

    with pytest.raises(sb.SettingsError, match=f'рядок {row}, колонка "Значення": "Статус за замовчуванням" - має бути рівно один'):
        ps.read_settings(str(path))


def test_broken_sentence_pattern_names_its_row(tmp_path):
    from utils import status_rules as sr

    path = _write(tmp_path)
    row = _set_cells(path, sr.SENTENCES_SHEET, "Шаблон речення", {"Шаблон речення": "вибув(у"})

    with pytest.raises(sb.SettingsError,
                       match=f'"Рапорти", таблиця "Речення рапорту", рядок {row}, колонка "Шаблон речення": помилка в шаблоні'):
        ps.read_settings(str(path))


def test_missing_rule_sheets_are_added_with_the_defaults(tmp_path):
    from utils import status_rules as sr

    path = _write_former_layout(tmp_path / sb.SETTINGS_FILE_NAME, skip=(sr.HEADINGS_SHEET, sr.ADDONS_SHEET))

    data = ps.load_settings(str(path))

    assert data == ps.with_default_statuses(_data())
    with sb.open_for_reading(str(path)) as workbook:
        assert sb.has_table(workbook, sr.REPORTS_BOOK_SHEET, sr.HEADINGS_SHEET)
        assert sb.has_table(workbook, sr.ADDONS_BOOK_SHEET, sr.ADDONS_SHEET)


def test_older_workbook_is_switched_to_the_current_layout_once(tmp_path, capsys):
    """Книга, де кожна таблиця мала власний аркуш (рідкісні - приховані):
    читається так само, а load_settings один раз переписує її в теперішній
    вигляд - кілька таблиць на аркуші, без прихованих аркушів, дані ті самі."""
    from utils import status_rules as sr

    path = _write_former_layout(tmp_path / sb.SETTINGS_FILE_NAME, hidden=(sr.GROUPS_SHEET, sr.SENTENCES_SHEET))
    before = ps.read_settings(str(path))

    data = ps.load_settings(str(path))

    assert data == before == ps.with_default_statuses(_data())
    assert "переведено на новий вигляд" in capsys.readouterr().out
    reopened = load_workbook(path)
    assert reopened.sheetnames == [sb.INSTRUCTIONS_SHEET, *(sheet for sheet, *_rest in ps.BOOK_SHEETS)]
    assert all(sheet.sheet_state == "visible" for sheet in reopened.worksheets)
    assert ps.read_settings(str(path)) == before
    assert (tmp_path / "НАЛАШТУВАННЯ (до оновлення).xlsx").exists()


def test_backup_of_an_earlier_upgrade_is_never_overwritten(tmp_path):
    path = _write_former_layout(tmp_path / sb.SETTINGS_FILE_NAME)
    earlier = tmp_path / "НАЛАШТУВАННЯ (до оновлення).xlsx"
    earlier.write_bytes(b"earlier backup")

    ps.load_settings(str(path))

    assert earlier.read_bytes() == b"earlier backup"
    assert len(load_workbook(tmp_path / "НАЛАШТУВАННЯ (до оновлення 2).xlsx").sheetnames) > 7


def test_unreadable_rewrite_brings_the_former_workbook_back(tmp_path, monkeypatch, capsys):
    path = _write_former_layout(tmp_path / sb.SETTINGS_FILE_NAME)
    before = path.read_bytes()
    real_read = ps.read_settings
    calls = []

    def read_settings(target):
        calls.append(target)
        if len(calls) == 2:  # книга, щойно переписана в новий вигляд
            raise sb.SettingsError("зіпсовано")
        return real_read(target)

    monkeypatch.setattr(ps, "read_settings", read_settings)

    data = ps.load_settings(str(path))

    assert "колишню повернуто" in capsys.readouterr().out
    assert path.read_bytes() == before
    assert data == ps.with_default_statuses(_data())


def test_equivalent_statuses_say_where_each_group_comes_from(tmp_path):
    from utils import status_rules as sr

    sheet, headers, rows = _locate(load_workbook(_write(tmp_path)), sr.EQUIVALENT_SHEET, "Статуси")
    notes = {
        sheet.cell(row=row, column=headers["Статуси"]).value: sheet.cell(row=row, column=headers[sr.EQUIVALENT_NOTE_HEADER]).value
        for row in rows if sheet.cell(row=row, column=headers["Статуси"]).value
    }

    assert notes["ЗВІЛЬНЕНИЙ, Звіль, Звіль."].startswith("book61, аркуш «Алф2»")
    assert all(notes.values())  # у кожної стандартної групи є довідка


def test_equivalent_sheet_without_the_reference_column_gets_it(tmp_path, capsys):
    """Аркуш "Рівнозначні статуси", записаний до появи довідки: при переписуванні
    книги в теперішній вигляд стандартні групи отримують довідку, решта - як були."""
    from utils import status_rules as sr

    path = _write_former_layout(tmp_path / sb.SETTINGS_FILE_NAME)
    workbook = sb.load_editable(str(path))
    del workbook[sr.EQUIVALENT_SHEET]
    sb.write_table(workbook, sr.EQUIVALENT_SHEET, sr.EQUIVALENT_COLUMNS[:1],
                   [{"statuses": ["БПШП", "ВПБП"]}, {"statuses": ["Свій", "Інший"]}])
    sb.save_workbook(workbook, str(path))

    data = ps.load_settings(str(path))

    assert data["EQUIVALENT_STATUSES"][0]["note"].startswith("book61, аркуш «ШП»")
    assert data["EQUIVALENT_STATUSES"][1] == {"statuses": ["Свій", "Інший"], "note": ""}
    assert "переведено на новий вигляд" in capsys.readouterr().out
    assert ps.read_settings(str(path))["EQUIVALENT_STATUSES"] == data["EQUIVALENT_STATUSES"]


def test_status_value_keeps_numbers_as_numbers():
    assert [ps.status_value(raw) for raw in (100, "100", " 10_МЕД ", "", None)] == [100, 100, "10_МЕД", None, None]


# --- settings_book: розбір значень і запис у наявну книгу ---------------------


@pytest.mark.parametrize("raw, expected", [
    (5, 5), (5.0, 5), ("05", 5), ("попередній", sb.PREVIOUS_MONTH_DAY), ("Попер.", sb.PREVIOUS_MONTH_DAY),
    (0, sb.PREVIOUS_MONTH_DAY), (None, None),
])
def test_day_values(raw, expected):
    assert sb.parse_value(raw, sb.DAY, "місце") == expected


@pytest.mark.parametrize("raw", [32, "завтра", 1.5])
def test_bad_day_values(raw):
    with pytest.raises(sb.SettingsError, match="місце: очікується день місяця"):
        sb.parse_value(raw, sb.DAY, "місце")


@pytest.mark.parametrize("raw, expected", [
    ("Так", True), ("так ", True), (True, True), (1, True), ("Ні", False), (False, False), ("-", False), (None, "типово"),
])
def test_bool_values(raw, expected):
    assert sb.parse_value(raw, sb.BOOL, "місце", default="типово") == expected


def test_bad_bool_value():
    with pytest.raises(sb.SettingsError, match='очікується "Так" або "Ні"'):
        sb.parse_value("може", sb.BOOL, "місце")


@pytest.mark.parametrize("raw, expected", [
    (350, "350"), (350.0, "350"), (" 350 ", "350"), (datetime(2026, 11, 5), "05.11.2026"), ("а\r\nб", "а\nб"), (None, ""),
])
def test_text_values(raw, expected):
    assert sb.parse_value(raw, sb.TEXT, "місце") == expected


def test_lines_and_lists():
    assert sb.parse_value("перший\n\n другий ", sb.LINES, "місце") == ["перший", "другий"]
    assert sb.parse_value("А0000, А0001; А0002", sb.LIST, "місце") == ["А0000", "А0001", "А0002"]


def test_text_starting_with_equals_is_not_a_formula(tmp_path):
    workbook = sb.new_workbook()
    columns = (sb.Column("text", "Текст"),)
    sb.write_table(workbook, "Аркуш", columns, [{"text": "=1+1"}])
    path = tmp_path / "book.xlsx"
    sb.save_workbook(workbook, str(path))

    with sb.open_for_reading(str(path)) as reopened:
        assert sb.read_rows(reopened, "Аркуш", columns) == [(sb.FIRST_DATA_ROW, {"text": "=1+1"})]


def test_update_table_cells_writes_only_the_given_fields_and_reports_missing_rows(tmp_path):
    columns = (sb.Column("day", "День", sb.DAY), sb.Column("bat", "БР бат", sb.NUMBER_TEXT), sb.Column("ref", "Посилання"))
    workbook = sb.new_workbook()
    sb.write_table(workbook, "Номери", columns, [
        {"day": sb.PREVIOUS_MONTH_DAY, "bat": "1", "ref": "старе"},
        {"day": 1, "bat": "", "ref": "лишається"},
    ])
    path = tmp_path / "book.xlsx"
    sb.save_workbook(workbook, str(path))

    missing = sb.update_table_cells(str(path), "Номери", columns[0], columns[1:], {
        1: {"bat": "350"}, sb.PREVIOUS_MONTH_DAY: {"bat": "", "ref": "нове"}, 31: {"bat": "9"},
    })

    assert missing == {31}
    with sb.open_for_reading(str(path)) as reopened:
        assert sb.read_rows(reopened, "Номери", columns) == [
            (sb.FIRST_DATA_ROW, {"day": sb.PREVIOUS_MONTH_DAY, "bat": "", "ref": "нове"}),
            (sb.FIRST_DATA_ROW + 1, {"day": 1, "bat": "350", "ref": "лишається"}),
        ]
    # Номер - числом, як його вписала б людина.
    assert load_workbook(path)["Номери"].cell(row=sb.FIRST_DATA_ROW + 1, column=2).value == 350


def test_reading_releases_the_file_at_once(tmp_path):
    """Прочитана книга одразу закрита: інакше на Windows запис у неї в тому
    самому запуску (номери документів) бачив би файл зайнятим і просив би
    закрити Excel. Збирач сміття вимкнено - він міг би закрити файл випадково."""
    path = _write(tmp_path)
    gc.disable()
    try:
        ps.read_settings(str(path))
        sb.save_workbook(_one_sheet_workbook(), str(path))  # заміна файлу на місці щойно прочитаного
    finally:
        gc.enable()


def test_save_workbook_asks_to_close_a_locked_file(tmp_path, monkeypatch):
    def locked(*_args):
        raise PermissionError("locked")

    monkeypatch.setattr(sb.os, "replace", locked)

    with pytest.raises(sb.SettingsError, match="Закрийте його"):
        sb.save_workbook(_one_sheet_workbook(), str(tmp_path / "book.xlsx"))
    assert list(tmp_path.iterdir()) == []  # тимчасовий файл прибрано


def _one_sheet_workbook():
    workbook = sb.new_workbook()
    workbook.create_sheet("Аркуш")
    return workbook
