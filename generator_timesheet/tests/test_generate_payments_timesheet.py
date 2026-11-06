from datetime import date, datetime

import pytest
from book61_fixtures import write_full_book61
from docx_fixtures import ARRIVAL_HEADER, write_daily_report
from openpyxl import Workbook, load_workbook
from treatment_fixtures import write_treatment_journal

from content.oblik_timesheet import normalize_name
from content.payment_mismatch_checker import _MATCH_FILL, _MISMATCH_FILL
from generators.check_addon_mismatches import check_addon_mismatches_file
import generators.generate_payments_timesheet as generate_payments_timesheet_module
from generators.generate_payments_timesheet import _drop_unresolved_already_in_mismatch_report, generate_payments_timesheet

_HEADER = ["ПОСАДА", "ЗВАННЯ", "ПІБ", datetime(2026, 7, 31), datetime(2026, 8, 1), datetime(2026, 8, 2)]


def _write_roster(path, rows, header=None):
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.title = "Табель"
    for col_idx, value in enumerate(header or _HEADER, start=1):
        ws.cell(row=1, column=col_idx, value=value)
    for row_idx, row_values in enumerate(rows, start=2):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    wb.save(str(path))
    return str(path)


def _write_info_unit_file(path, rows, dates):
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.title = "ТАБЕЛЬ"
    for col_idx, text in enumerate(["Посада", "Звання", "ПІБ", "ДАТА"], start=1):
        ws.cell(row=1, column=col_idx, value=text)
    for idx, value in enumerate(dates, start=4):
        ws.cell(row=2, column=idx, value=value)
    for row_idx, row_values in enumerate(rows, start=3):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    wb.save(str(path))
    return str(path)


def _write_info_unit_file_dates_in_row1(path, rows, dates):
    """Реальний структурний варіант: дати ОДРАЗУ в рядку 1 (без окремого
    групового заголовка "ДАТА" + рядка дат) - дані з рядка 2, а не з рядка 3."""
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.title = "ТАБЕЛЬ"
    for col_idx, text in enumerate(["Посада", "Звання", "ПІБ"], start=1):
        ws.cell(row=1, column=col_idx, value=text)
    for idx, value in enumerate(dates, start=4):
        ws.cell(row=1, column=idx, value=value)
    for row_idx, row_values in enumerate(rows, start=2):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    wb.save(str(path))
    return str(path)


def test_generate_payments_timesheet_replaces_rvz_with_payment_category(tmp_path):
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(
        report_dir / "01.08.2026 - щоденний рапорт.docx",
        groups=[(
            "ПРИБУЛИ до складу сил та засобів корпусу A0000:",
            [("З відпустки:", ARRIVAL_HEADER, [["1", "матрос", "ПЕРШИЙ Перший Перший", "Посада", "01.08.2026", "Прибув"]])],
        )],
    )
    roster_path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ВП", None, None],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "РВЗ", None, None],
    ])
    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_info_unit_file(
        info_unit_dir / "1 рота.xlsx",
        rows=[["Посада1", "сержант", "ПЕРШИЙ Перший Перший", 100, 100]],
        dates=[datetime(2026, 8, 1), datetime(2026, 8, 2)],
    )
    output_path = tmp_path / "output" / "ОБЛІК для виплат.xlsx"
    review_path = tmp_path / "output" / "review.txt"
    mismatch_report_path = tmp_path / "output" / "error_mis_statuses.xlsx"

    result_path, unresolved, _mismatch_open_path = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(output_path),
        review_file_name=str(review_path),
        mismatch_report_file_name=str(mismatch_report_path),
    )

    assert result_path == str(output_path)
    saved = load_workbook(str(output_path))["Табель"]
    # ПЕРШИЙ: З відпустки на 01.08 -> РВЗ -> замінено на 100 (подає 1 рота); 02.08 продовжує 100.
    assert [saved.cell(row=2, column=c).value for c in range(4, 6)] == ["ВП", 100]
    # ДРУГИЙ: РВЗ від самого початку, але 1 рота нічого про нього не подає - лишається на ручну перевірку.
    assert saved.cell(row=3, column=4).value == "РВЗ"
    assert any("ДРУГИЙ" in item.get("pib_raw", "") for item in unresolved)
    assert review_path.exists()
    # error_mis_statuses.xlsx: ПЕРШИЙ повністю збігається з "1 рота" (100==100),
    # а ДРУГОГО жоден файл не подає і він не "БЗ" - за рішенням користувача
    # це розбіжність ("Немає в жодному файлі information_unit").
    report = load_workbook(str(mismatch_report_path))["Розбіжності"]
    assert [report.cell(row=r, column=1).value for r in (2, 3)] == [
        "ОБЛІК для виплат (за рапортами)", "Немає в жодному файлі information_unit",
    ]
    assert report.cell(row=2, column=5).value == "ДРУГИЙ Другий Другий"


def test_generate_payments_timesheet_creates_neither_review_nor_mismatch_file_when_fully_clean(tmp_path):
    """Рішення користувача: коли прогін ПОВНІСТЮ чистий (жодного
    unresolved-запису, жодної розбіжності), НІ Потребує_ручної_перевірки_
    виплати.txt (так само, як і Потребує_ручної_перевірки.txt у
    generate_timesheet - _finish), НІ error_mis_statuses.xlsx НЕ
    створюються взагалі - немає сенсу писати файли, коли розповідати нема
    про що."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "01.08.2026 - щоденний рапорт.docx")

    roster_path = _write_roster(
        tmp_path / "ОБЛІК.xlsx",
        [["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ"]],
        header=["ПОСАДА", "ЗВАННЯ", "ПІБ", datetime(2026, 8, 1)],
    )
    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_info_unit_file(
        info_unit_dir / "1 рота.xlsx",
        rows=[["Посада1", "сержант", "ПЕРШИЙ Перший Перший", 100]],
        dates=[datetime(2026, 8, 1)],
    )

    review_path = tmp_path / "output" / "review.txt"
    mismatch_report_path = tmp_path / "output" / "error_mis_statuses.xlsx"

    _output_path, unresolved, mismatch_open_path = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(tmp_path / "output" / "ОБЛІК.xlsx"),
        review_file_name=str(review_path),
        mismatch_report_file_name=str(mismatch_report_path),
    )

    assert unresolved == []
    assert mismatch_open_path is None
    assert not review_path.exists()
    assert not mismatch_report_path.exists()


def _write_bchs_file(path, title_row1, rows):
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.title = "БЧС"
    ws.cell(row=1, column=1, value=title_row1)
    header = ["Посада", "Звання", "ПРІЗВИЩЕ ім'я по батькові", "БЧС", "Виплата додактової винагороди"]
    for col_idx, text in enumerate(header, start=1):
        ws.cell(row=2, column=col_idx, value=text)
    for row_idx, row_values in enumerate(rows, start=3):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    wb.save(str(path))
    return str(path)


def test_generate_payments_timesheet_replaces_rvz_using_bchs_sheet_only_source(tmp_path):
    """Регресійний тест на реальний випадок: людина отримує "РВЗ" у
    результаті, хоча ЄДИНЕ джерело даних про неї - файл information_unit,
    де ЛИШЕ аркуш БЧС (без жодного релевантного запису в ТАБЕЛЬ) прямо каже
    "100" для потрібної дати - число має підставитись, а не лишитись "РВЗ"."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "01.08.2026 - щоденний рапорт.docx")

    roster_path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", None, None],
    ])

    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_bchs_file(
        info_unit_dir / "БЧС Підрозділ1 01.08.2026.xlsx",
        title_row1="ВЗВОД ЗВ'ЯЗКУ 01.08.2026",
        rows=[["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", 100]],
    )

    output_path = tmp_path / "output" / "ОБЛІК для виплат.xlsx"

    _result_path, unresolved, _mismatch_open_path = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(output_path),
        review_file_name=str(tmp_path / "output" / "review.txt"),
        mismatch_report_file_name=str(tmp_path / "output" / "error_mis_statuses.xlsx"),
    )

    # 31.07 (базова колонка) - у файлі information_unit немає про неї даних,
    # лишається "РВЗ" на ручну перевірку - це ОКРЕМЕ, очікуване питання, не
    # те, що перевіряє цей тест.
    assert not any("01.08.2026" in item["reason"] for item in unresolved)
    saved = load_workbook(str(output_path))["Табель"]
    assert saved.cell(row=2, column=5).value == 100  # 01.08.2026 - замінено на 100, а не лишилось "РВЗ"


def test_generate_payments_timesheet_replaces_rvz_when_tabel_sheet_has_dates_in_row1(tmp_path):
    """Регресійний тест на реальний випадок: аркуш ТАБЕЛЬ файлу information_unit
    має дати ОДРАЗУ в рядку 1 (без окремого рядка-групи "ДАТА" в рядку 1 +
    самих дат у рядку 2, як у решти файлів) - раніше такий файл мовчки НЕ
    читався взагалі (жодної колонки-дати не знаходилось), тож людина
    лишалась "РВЗ" на кожен день, хоча файл прямо каже "100"."""
    header = ["ПОСАДА", "ЗВАННЯ", "ПІБ", datetime(2026, 7, 31), datetime(2026, 8, 1), datetime(2026, 8, 2), datetime(2026, 8, 3)]
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.title = "Табель"
    for col_idx, value in enumerate(header, start=1):
        ws.cell(row=1, column=col_idx, value=value)
    ws.cell(row=2, column=1, value="Командир роти")
    ws.cell(row=2, column=2, value="старший лейтенант")
    ws.cell(row=2, column=3, value="ПЕРШИЙ Перший Перший")
    ws.cell(row=2, column=4, value="РВЗ")
    roster_path = tmp_path / "ОБЛІК.xlsx"
    wb.save(str(roster_path))

    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "03.08.2026 - щоденний рапорт.docx")

    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_info_unit_file_dates_in_row1(
        info_unit_dir / "підрозділ 1.xlsx",
        rows=[["Командир роти", "старший лейтенант", "ПЕРШИЙ Перший Перший", 100, 100, 100]],
        dates=[datetime(2026, 8, 1), datetime(2026, 8, 2), datetime(2026, 8, 3)],
    )

    output_path = tmp_path / "output" / "ОБЛІК для виплат.xlsx"

    generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=str(roster_path),
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(output_path),
        review_file_name=str(tmp_path / "output" / "review.txt"),
        mismatch_report_file_name=str(tmp_path / "output" / "error_mis_statuses.xlsx"),
    )

    saved = load_workbook(str(output_path))["Табель"]
    # 01.08, 02.08, 03.08 - усі три дні замінено на 100, а не лишились "РВЗ".
    assert [saved.cell(row=2, column=c).value for c in range(5, 8)] == [100, 100, 100]


def test_generate_payments_timesheet_writes_mismatch_report_after_saving_payments_file(tmp_path):
    """Реальний випадок: файл information_unit каже "ВД" (не число) для
    людини - apply_payment_values НЕ підставляє нечислове значення (лишає
    "РВЗ"), тож звіт розбіжностей (сформований ОДРАЗУ ПІСЛЯ ОБЛІК для виплат,
    звіряючи саме ЦЕЙ згенерований результат) має явно показати "РВЗ" і "ВД"
    поряд, з назвою файлу-джерела."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "01.08.2026 - щоденний рапорт.docx")

    roster_path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", None, None],
    ])

    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_bchs_file(
        info_unit_dir / "БЧС Підрозділ1 01.08.2026.xlsx",
        title_row1="ВЗВОД ЗВ'ЯЗКУ 01.08.2026",
        rows=[["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ВД", 100]],
    )

    mismatch_report_path = tmp_path / "output" / "error_mis_statuses.xlsx"

    _output_path, _unresolved, mismatch_open_path = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(tmp_path / "output" / "ОБЛІК для виплат.xlsx"),
        review_file_name=str(tmp_path / "output" / "review.txt"),
        mismatch_report_file_name=str(mismatch_report_path),
    )

    # Є хоч один запис - третій елемент повертає шлях до звіту (index.py
    # відкриває його автоматично), а не None.
    assert mismatch_open_path == str(mismatch_report_path)
    saved = load_workbook(str(mismatch_report_path))["Розбіжності"]
    assert saved.cell(row=2, column=1).value == "ОБЛІК для виплат (за рапортами)"
    assert saved.cell(row=3, column=1).value == "БЧС Підрозділ1 01.08.2026.xlsx"
    assert saved.cell(row=2, column=5).value == "ПЕРШИЙ Перший Перший"
    # Рішення користувача: разом зі звітом, поруч із ним,
    # повідомлення командирам (значення - як у ОБЛІК).
    message = (tmp_path / "output" / "Повідомлення_командирам.txt").read_text(encoding="utf-8")
    assert message.startswith("Командири підрозділів нижчевказаних виправте табель")
    assert "ПЕРШИЙ Перший Перший - 01.08.2026 статус має бути (РВЗ)" in message


def _mismatch_record(**overrides):
    base = {
        "file_name": "файл1.xlsx", "roster_pib_raw": "ПЕРШИЙ Перший Перший", "file_pib_raw": "ПЕРШИЙ Перший Перший",
        "roster_pidrozdil": None, "roster_posada": "Посада1", "roster_zvannya": "сержант",
        "file_pidrozdil": None, "file_posada": "Посада1", "file_zvannya": "сержант",
        "dates": {date(2026, 8, 1): ("РВЗ", "ВЛК")},
    }
    base.update(overrides)
    return base


def test_drop_unresolved_already_in_mismatch_report_removes_the_duplicate():
    """Рішення користувача: unresolved-запис apply_payment_values
    для (людина, дата), яка ВЖЕ показана як розбіжність у
    error_mis_statuses.xlsx (records), - дублює той самий факт у ДВОХ файлах
    одразу, тож прибирається, лишаючи ЛИШЕ error_mis_statuses.xlsx."""
    unresolved = [{
        "reason": "\"ПЕРШИЙ Перший Перший\" 01.08.2026: підрозділ подає ['ВЛК'] - не єдине числове значення категорії виплати, статус \"РВЗ\" лишено як є.",
        "pib_raw": "ПЕРШИЙ Перший Перший", "report_date": date(2026, 8, 1),
    }]
    records = [_mismatch_record()]

    assert _drop_unresolved_already_in_mismatch_report(unresolved, records) == []


def test_drop_unresolved_already_in_mismatch_report_keeps_entries_without_a_matching_record():
    """Запис БЕЗ відповідника в records (напр. "Немає категорії виплати" -
    підрозділ узагалі нічого не подав, або ІНША дата/людина) НЕ прибирається -
    для НЬОГО error_mis_statuses.xlsx нічого не показує (немає даних для
    порівняння), тож лишається єдиним джерелом інформації."""
    unresolved = [
        {
            "reason": "Немає категорії виплати для \"ДРУГИЙ Другий Другий\" на 01.08.2026 - підрозділ ще не подав дані.",
            "pib_raw": "ДРУГИЙ Другий Другий", "report_date": date(2026, 8, 1),
        },
        {
            "reason": "...",
            "pib_raw": "ПЕРШИЙ Перший Перший", "report_date": date(2026, 8, 2),  # ІНША дата, не 01.08
        },
    ]
    records = [_mismatch_record()]

    assert _drop_unresolved_already_in_mismatch_report(unresolved, records) == unresolved


def test_generate_payments_timesheet_does_not_duplicate_non_numeric_mismatch_in_review_file(tmp_path):
    """Реальний випадок: файл information_unit каже "ВЛК" (не число) для
    людини - apply_payment_values лишає "РВЗ" і, БЕЗ цього виправлення, додав
    би унерозв'язаний запис "не єдине числове значення" в Потребує_ручної_
    перевірки_виплати.txt - хоча ТА САМА розбіжність ВЖЕ показана в
    error_mis_statuses.xlsx (ВЛК - "порівнюваний" статус). Рішення користувача: НЕ дублюється: unresolved НЕ містить цей запис,
    error_mis_statuses.xlsx і далі показує розбіжність."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "01.08.2026 - щоденний рапорт.docx")

    roster_path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ВОСЬМИЙ Восьмий Восьмий", "РВЗ", None, None],
    ])

    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_bchs_file(
        info_unit_dir / "БЧС Підрозділ1 01.08.2026.xlsx",
        title_row1="ВЗВОД ЗВ'ЯЗКУ 01.08.2026",
        rows=[["Посада1", "сержант", "ВОСЬМИЙ Восьмий Восьмий", "ВЛК", 100]],
    )

    mismatch_report_path = tmp_path / "output" / "error_mis_statuses.xlsx"

    _output_path, unresolved, mismatch_open_path = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(tmp_path / "output" / "ОБЛІК для виплат.xlsx"),
        review_file_name=str(tmp_path / "output" / "review.txt"),
        mismatch_report_file_name=str(mismatch_report_path),
    )

    # Дубльований запис (01.08 - "не єдине числове значення", ВЖЕ показано в
    # error_mis_statuses.xlsx) прибрано, АЛЕ 31.07 (базова колонка, "Немає
    # категорії виплати" - інша, НЕпов'язана причина, для НЕЇ немає розбіжності
    # у файлі, отже НЕ прибирається) - і далі лишається.
    assert not any("не єдине числове значення" in item["reason"] for item in unresolved)
    assert any("Немає категорії виплати" in item["reason"] for item in unresolved)
    assert mismatch_open_path == str(mismatch_report_path)
    saved = load_workbook(str(mismatch_report_path))["Розбіжності"]
    assert saved.cell(row=3, column=1).value == "БЧС Підрозділ1 01.08.2026.xlsx"


@pytest.mark.parametrize("pair_date, still_listed", [
    (date(2026, 7, 31), False),
    (date(2026, 8, 1), True),
])
def test_generate_payments_timesheet_pair_exclusion_drops_only_that_date_and_keeps_the_status(
    tmp_path, monkeypatch, pair_date, still_listed,
):
    """Пара (дата, ПІБ) з data.json прибирає запис для ручної перевірки ЛИШЕ на
    цю дату: той самий ПІБ на іншу дату лишається, а статус "РВЗ" у табелі від
    винятку НЕ змінюється - виняток діє лише на review-файл."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "01.08.2026 - щоденний рапорт.docx")

    roster_path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ВОСЬМИЙ Восьмий Восьмий", "РВЗ", None, None],
    ])

    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_bchs_file(
        info_unit_dir / "БЧС Підрозділ1 01.08.2026.xlsx",
        title_row1="ВЗВОД ЗВ'ЯЗКУ 01.08.2026",
        rows=[["Посада1", "сержант", "ВОСЬМИЙ Восьмий Восьмий", "ВЛК", 100]],
    )
    # Той самий модуль, що імпортований під час збору (рядковий шлях дав би свіжу копію).
    monkeypatch.setattr(
        generate_payments_timesheet_module,
        "PAYMENTS_REVIEW_EXCLUDE_PAIRS",
        ((pair_date, "ВОСЬМИЙ Восьмий Восьмий"),),
    )

    output_path, unresolved, _mismatch_open_path = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(tmp_path / "output" / "ОБЛІК для виплат.xlsx"),
        review_file_name=str(tmp_path / "output" / "review.txt"),
        mismatch_report_file_name=str(tmp_path / "output" / "error_mis_statuses.xlsx"),
    )

    assert any("Немає категорії виплати" in item["reason"] for item in unresolved) is still_listed
    assert load_workbook(output_path)["Табель"].cell(row=2, column=4).value == "РВЗ"


def test_generate_payments_timesheet_mismatch_report_includes_pidrozdil_column(tmp_path):
    """Роcтер із колонкою ПІДРОЗДІЛ - звіт розбіжностей має показати її поряд
    із Посадою/Званням, звірену з файлом information_unit (рішення користувача) - файл цього тесту не подає Підрозділ узагалі
    (None), тож роcтерове значення "підрозділ 1" - розбіжність, НЕ об'єднана
    клітинка."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "01.08.2026 - щоденний рапорт.docx")

    roster_path = _write_roster(
        tmp_path / "ОБЛІК.xlsx",
        [["підрозділ 1", "Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", None]],
        header=["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", datetime(2026, 7, 31), datetime(2026, 8, 1)],
    )

    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_bchs_file(
        info_unit_dir / "БЧС Підрозділ1 01.08.2026.xlsx",
        title_row1="ВЗВОД ЗВ'ЯЗКУ 01.08.2026",
        rows=[["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ВД", 100]],
    )

    mismatch_report_path = tmp_path / "output" / "error_mis_statuses.xlsx"

    generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(tmp_path / "output" / "ОБЛІК для виплат.xlsx"),
        review_file_name=str(tmp_path / "output" / "review.txt"),
        mismatch_report_file_name=str(mismatch_report_path),
    )

    saved = load_workbook(str(mismatch_report_path))["Розбіжності"]
    assert saved.cell(row=1, column=2).value == "Підрозділ"
    assert saved.cell(row=2, column=2).value == "підрозділ 1"
    merged = {str(rng) for rng in saved.merged_cells.ranges}
    assert "B2:B3" not in merged  # Підрозділ - файл None, роcтер "підрозділ 1" - розбіжність


def test_generate_payments_timesheet_mismatch_report_includes_people_missing_from_roster(tmp_path):
    """Реальний випадок: людина є у файлі information_unit, але ВЗАГАЛІ
    відсутня в роcтері (ОБЛІК.xlsx) - раніше мовчки не потрапляла в
    error_mis_statuses.xlsx узагалі (звіт порівнював лише людей з ростера)."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "01.08.2026 - щоденний рапорт.docx")

    roster_path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", None],
    ])

    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_bchs_file(
        info_unit_dir / "БЧС Підрозділ1 01.08.2026.xlsx",
        title_row1="ВЗВОД ЗВ'ЯЗКУ 01.08.2026",
        rows=[
            ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", 100],
            ["Посада3", "матрос", "ТРЕТІЙ Третій Третій", "РВЗ", 100],
        ],
    )

    mismatch_report_path = tmp_path / "output" / "error_mis_statuses.xlsx"

    generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(tmp_path / "output" / "ОБЛІК для виплат.xlsx"),
        review_file_name=str(tmp_path / "output" / "review.txt"),
        mismatch_report_file_name=str(mismatch_report_path),
    )

    saved = load_workbook(str(mismatch_report_path))["Розбіжності"]
    # Немає роcтерового рядка для ТРЕТІЙ - ПІБ НЕ об'єднана, ім'я показане
    # лише на рядку файлу (row_b); роcтерова сторона (row_a) показує
    # "Не знайдено в ОБЛІК.xlsx", а не порожню клітинку.
    pib_values = [saved.cell(row=r, column=5).value for r in range(2, saved.max_row + 1)]
    assert "ТРЕТІЙ Третій Третій" in pib_values
    row = next(r for r in range(2, saved.max_row + 1) if saved.cell(row=r, column=5).value == "ТРЕТІЙ Третій Третій")
    assert saved.cell(row=row, column=1).value == "БЧС Підрозділ1 01.08.2026.xlsx"
    assert saved.cell(row=row - 1, column=1).value == "ОБЛІК для виплат (за рапортами)"
    assert saved.cell(row=row - 1, column=5).value == "Не знайдено в ОБЛІК.xlsx"
    # Дата (01.08) - ОКРЕМА колонка (заголовок - сама дата), як в ОБЛІК.xlsx.
    date_col = next(c for c in range(6, saved.max_column + 1) if saved.cell(row=1, column=c).value == datetime(2026, 8, 1))
    assert saved.cell(row=row - 1, column=date_col).value == "Не знайдено в ОБЛІК.xlsx"
    assert saved.cell(row=row, column=date_col).value == "РВЗ"  # файл подав 100 - вид оплати людини на РВЗ


def test_generate_payments_timesheet_mismatch_report_flags_patronymic_spelling_difference(tmp_path):
    """Реальний випадок: роcтер (ОБЛІК.xlsx) і файл information_unit
    подають ТЕ САМЕ прізвище+ім'я, але РІЗНЕ по-батькові (друкарська
    помилка) - раніше це взагалі не потрапляло в error_mis_statuses.xlsx як
    порівнянна розбіжність ПІБ (normalize_name різний -> дані файлу не
    зіставлялись з роcтером), тепер має показати РЕАЛЬНЕ роcтерове ПІБ поряд
    із файловим написанням, а не мовчати чи показувати лише загальний
    напис "не знайдено"."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "01.08.2026 - щоденний рапорт.docx")

    roster_path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", None],
    ])

    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_bchs_file(
        info_unit_dir / "БЧС Підрозділ1 01.08.2026.xlsx",
        title_row1="ВЗВОД ЗВ'ЯЗКУ 01.08.2026",
        # "Перекший" замість роcтерового "Перший" - та сама людина, друкарська
        # помилка в по-батькові (реальний випадок з подібною одно-двобуквеною різницею).
        rows=[["Посада1", "сержант", "ПЕРШИЙ Перший Перекший", "РВЗ", 100]],
    )

    mismatch_report_path = tmp_path / "output" / "error_mis_statuses.xlsx"

    generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(tmp_path / "output" / "ОБЛІК для виплат.xlsx"),
        review_file_name=str(tmp_path / "output" / "review.txt"),
        mismatch_report_file_name=str(mismatch_report_path),
    )

    saved = load_workbook(str(mismatch_report_path))["Розбіжності"]
    assert saved.cell(row=2, column=1).value == "ОБЛІК для виплат (за рапортами)"
    assert saved.cell(row=3, column=1).value == "БЧС Підрозділ1 01.08.2026.xlsx"
    # РЕАЛЬНЕ роcтерове написання - НЕ загальний напис "не знайдено" -
    # показане поряд із файловим, щоб розбіжність у ПІБ було видно.
    assert saved.cell(row=2, column=5).value == "ПЕРШИЙ Перший Перший"
    assert saved.cell(row=3, column=5).value == "ПЕРШИЙ Перший Перекший"
    merged = {str(rng) for rng in saved.merged_cells.ranges}
    assert "E2:E3" not in merged  # ПІБ відрізняється - НЕ об'єднана
    assert saved.cell(row=2, column=5).fill.fgColor.rgb == "FFFFC7CE"
    assert saved.cell(row=3, column=5).fill.fgColor.rgb == "FFFFC7CE"


def test_generate_payments_timesheet_does_not_create_mismatch_report_when_everything_matches(tmp_path):
    """Рішення користувача: немає сенсу створювати
    error_mis_statuses.xlsx лише із заголовком, коли розповідати нема про що."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "01.08.2026 - щоденний рапорт.docx")

    roster_path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", None, None],
    ])
    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()

    mismatch_report_path = tmp_path / "output" / "error_mis_statuses.xlsx"

    _result_path, unresolved, mismatch_open_path = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(tmp_path / "output" / "ОБЛІК для виплат.xlsx"),
        review_file_name=str(tmp_path / "output" / "review.txt"),
        mismatch_report_file_name=str(mismatch_report_path),
    )

    assert unresolved != []  # ПЕРШИЙ лишається "РВЗ" без жодних даних про виплату
    assert mismatch_open_path is None  # немає жодного запису - нема сенсу відкривати звіт автоматично
    assert not mismatch_report_path.exists()


def test_generate_payments_timesheet_uses_the_same_output_file_name_as_generate_timesheet():
    """Рішення користувача: PAYMENTS_OUTPUT_FILE_NAME НАВМИСНО
    дорівнює OUTPUT_FILE_NAME (обидва - "ОБЛІК.xlsx"): обидва прогони завжди
    окремі (output/ очищається перед кожним, index.py), колізії імені
    немає."""
    from constants import OUTPUT_FILE_NAME, PAYMENT_MISMATCH_REPORT_FILE_NAME, PAYMENTS_OUTPUT_FILE_NAME
    assert PAYMENTS_OUTPUT_FILE_NAME == OUTPUT_FILE_NAME
    assert PAYMENT_MISMATCH_REPORT_FILE_NAME not in (OUTPUT_FILE_NAME, PAYMENTS_OUTPUT_FILE_NAME)


def test_vlk_between_leave_days_stays_the_leave_in_the_payments_oblik(tmp_path, monkeypatch):
    """Реальний випадок: рапорт - ВЛК, наступного дня знову ВПСЗ, у файлі БЧС
    увесь час ВПСЗ. ОБЛІК давав ВЛК, у звіті була розбіжність, а командиру йшло
    "статус має бути (ВЛК)". ВЛК під час відпустки не оплачується (рішення
    користувача), тож правий файл БЧС."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "02.08.2026 - щоденний рапорт.docx", groups=[])
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВЛК", "effective_date": date(2026, 8, 1)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВПСЗ", "effective_date": date(2026, 8, 2)},
    ]
    monkeypatch.setattr("generators.generate_timesheet.extract_status_events", lambda path: (events, []))
    days = [datetime(2026, 7, 31), datetime(2026, 8, 1), datetime(2026, 8, 2)]
    roster_path = _write_roster(
        tmp_path / "ОБЛІК.xlsx", [["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВПСЗ", None, None]],
        header=["ПОСАДА", "ЗВАННЯ", "ПІБ"] + days,
    )
    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_info_unit_file(info_unit_dir / "1 рота.xlsx", rows=[
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВПСЗ", "ВПСЗ", "ВПСЗ"],
    ], dates=days)
    output_dir = tmp_path / "output"

    output_path, _unresolved, report_path = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(output_dir / "ОБЛІК.xlsx"),
        review_file_name=str(output_dir / "review.txt"),
        mismatch_report_file_name=str(output_dir / "error_mis_statuses.xlsx"),
        mismatch_message_file_name=str(output_dir / "message.txt"),
    )

    assert report_path is None
    assert not (output_dir / "message.txt").exists()
    assert [load_workbook(output_path)["Табель"].cell(row=2, column=col).value for col in (4, 5, 6)] == ["ВПСЗ"] * 3


def test_text_statuses_become_categories_the_money_project_reads(tmp_path):
    """Реальні випадки: ЛМР, а підрозділ подає ШП; НОВ, а підрозділ подає БЗВП.
    Клітинки стають категоріями, і звірка з БЧС лишається без розбіжностей."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "01.08.2026 - щоденний рапорт.docx", groups=[])
    people = [
        ("ПЕРШИЙ Перший Перший", "ЛМР", "ШП", "10_МЕД"), ("ДРУГИЙ Другий Другий", "ЗВФН", "ЗВФН", "10_МЕД"),
        ("ТРЕТІЙ Третій Третій", "НОВ", "БЗВП", 10), ("ЧЕТВЕРТИЙ Четвертий Четвертий", "Адап", "Адап", 10),
        ("ВОСЬМИЙ Восьмий Восьмий", "НОВ", 100, 100),
    ]
    days = [datetime(2026, 7, 31), datetime(2026, 8, 1)]
    roster_path = _write_roster(
        tmp_path / "ОБЛІК.xlsx", [["Посада1", "матрос", name, status, status] for name, status, _unit, _result in people],
        header=["ПОСАДА", "ЗВАННЯ", "ПІБ"] + days,
    )
    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_info_unit_file(info_unit_dir / "1 рота.xlsx", rows=[
        ["Посада1", "матрос", name, submitted, submitted] for name, _status, submitted, _result in people
    ], dates=days)
    output_dir = tmp_path / "output"

    output_path, _unresolved, report_path = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(output_dir / "ОБЛІК.xlsx"),
        review_file_name=str(output_dir / "review.txt"),
        mismatch_report_file_name=str(output_dir / "error_mis_statuses.xlsx"),
        mismatch_message_file_name=str(output_dir / "message.txt"),
    )

    table = load_workbook(output_path)["Табель"]
    assert [table.cell(row=row, column=5).value for row in range(2, 7)] == [result for *_rest, result in people]
    assert report_path is None


def _generate_payments_for_category_100(tmp_path, paid_days, **options):
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(
        report_dir / "01.08.2026 - щоденний рапорт.docx",
        groups=[(
            "ПРИБУЛИ до складу сил та засобів корпусу A0000:",
            [("З відпустки:", ARRIVAL_HEADER, [["1", "матрос", "ПЕРШИЙ Перший Перший", "Посада", "01.08.2026", "Прибув"]])],
        )],
    )
    dates = [datetime(2026, 8, d) for d in range(1, 32)]
    roster_path = _write_roster(
        tmp_path / "ОБЛІК.xlsx",
        [["Посада1", "сержант", "ПЕРШИЙ Перший Перший"] + ["РВЗ"] * 31],
        header=["ПОСАДА", "ЗВАННЯ", "ПІБ"] + dates,
    )
    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_info_unit_file(
        info_unit_dir / "1 рота.xlsx",
        rows=[["Посада1", "сержант", "ПЕРШИЙ Перший Перший"] + [100] * paid_days],
        dates=dates[:paid_days],
    )
    output_path, _unresolved, _mismatch = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(tmp_path / "output" / "ОБЛІК.xlsx"),
        review_file_name=str(tmp_path / "output" / "review.txt"),
        mismatch_report_file_name=str(tmp_path / "output" / "error_mis_statuses.xlsx"),
        **options,
    )
    return output_path


@pytest.fixture
def august_limit_30(monkeypatch):
    """Ліміт категорії 100 задає користувач (data.json), тож тести його
    фіксують самі: 30 днів, і 100 увесь серпень - порушення."""
    monkeypatch.setattr("content.oblik_timesheet.CONST_PAYMENT_CATEGORY_100_MAX_DAYS", 30)


def _check_category_100(tmp_path, output_path, check_file, **options):
    from generators.generate_payments_timesheet import check_payment_category_100_file

    return check_payment_category_100_file(
        str(output_path), check_file_name=str(check_file), information_unit_dir=str(tmp_path / "information_unit"), **options,
    )


_YELLOW_PERSON = {"name": "ПЕРШИЙ Перший Перший", "days": None, "color": "FFFFFF00"}
_PLAIN_PERSON = {"name": "ПЕРШИЙ Перший Перший", "days": None, "color": None}


def test_generated_oblik_shows_the_color_of_the_person_in_the_name_cell(tmp_path):
    output_path = _generate_payments_for_category_100(tmp_path, paid_days=5, category_100_people=(_YELLOW_PERSON,))

    assert load_workbook(output_path)["Табель"]["C2"].fill.start_color.rgb == "FFFFFF00"


def test_generated_oblik_has_no_color_for_a_person_without_one(tmp_path, august_limit_30):
    output_path = _generate_payments_for_category_100(tmp_path, paid_days=5, category_100_people=(_PLAIN_PERSON,))

    assert load_workbook(output_path)["Табель"]["C2"].fill.fill_type is None


def test_generated_oblik_shows_a_person_from_the_category_100_table_in_yellow(tmp_path, august_limit_30):
    """Рішення користувача: хто є серед порушень у таблиці, той жовтий і в Табелі."""
    output_path = _generate_payments_for_category_100(tmp_path, paid_days=31, category_100_people=())
    check_file = tmp_path / "output" / "check.xlsx"

    violations, _check_path = _check_category_100(tmp_path, output_path, check_file, category_100_people=())

    assert [violation["pib_raw"] for violation in violations] == ["ПЕРШИЙ Перший Перший"]
    assert load_workbook(output_path)["Табель"]["C2"].fill.start_color.rgb == "FFFFFF00"


def test_category_100_table_shows_the_color_of_the_person(tmp_path, august_limit_30):
    output_path = _generate_payments_for_category_100(tmp_path, paid_days=31, category_100_people=())
    check_file = tmp_path / "output" / "check.xlsx"

    _check_category_100(tmp_path, output_path, check_file, category_100_people=(_YELLOW_PERSON,))

    assert load_workbook(check_file).worksheets[0].cell(row=2, column=2).fill.start_color.rgb == "FFFFFF00"


def test_marked_person_within_the_limit_is_in_the_table_but_is_not_a_violation(tmp_path):
    """Рішення користувача: людина з кольором є в таблиці й без порушення."""
    output_path = _generate_payments_for_category_100(tmp_path, paid_days=5, category_100_people=())
    check_file = tmp_path / "output" / "check.xlsx"

    violations, check_path = _check_category_100(tmp_path, output_path, check_file, category_100_people=(_YELLOW_PERSON,))

    assert violations == []
    assert check_path == str(check_file)
    row = [load_workbook(check_file).worksheets[0].cell(row=2, column=col) for col in (2, 5, 6)]
    assert [cell.value for cell in row] == ["ПЕРШИЙ Перший Перший", "01.08.2026 - 05.08.2026", 5]
    assert row[0].fill.start_color.rgb == "FFFFFF00"


def test_marked_person_with_a_violation_has_one_row_in_the_table(tmp_path, august_limit_30):
    output_path = _generate_payments_for_category_100(tmp_path, paid_days=31, category_100_people=())
    check_file = tmp_path / "output" / "check.xlsx"

    _check_category_100(tmp_path, output_path, check_file, category_100_people=(_YELLOW_PERSON,))

    report = load_workbook(check_file).worksheets[0]
    assert [report.cell(row=row, column=2).value for row in (2, 3)] == ["ПЕРШИЙ Перший Перший", None]


def test_category_100_table_lists_violations_first(tmp_path, august_limit_30):
    dates = [datetime(2026, 8, d) for d in range(1, 32)]
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший"] + [100] * 31,
        ["Посада2", "сержант", "ДРУГИЙ Другий Другий"] + [100] * 31,
    ], header=["ПОСАДА", "ЗВАННЯ", "ПІБ"] + dates)
    check_file = tmp_path / "check.xlsx"

    _check_category_100(tmp_path, oblik, check_file, category_100_people=({**_YELLOW_PERSON, "days": 31},))

    report = load_workbook(check_file).worksheets[0]
    assert [report.cell(row=row, column=2).value for row in (2, 3)] == ["ДРУГИЙ Другий Другий", "ПЕРШИЙ Перший Перший"]


def test_person_days_decide_whether_the_category_100_table_appears(tmp_path, august_limit_30):
    output_path = _generate_payments_for_category_100(tmp_path, paid_days=31, category_100_people=())
    check_file = tmp_path / "output" / "check.xlsx"

    violations, check_path = _check_category_100(
        tmp_path, output_path, check_file, category_100_people=({**_PLAIN_PERSON, "days": 31},),
    )

    assert violations == []
    assert check_path is None


def test_check_payment_category_100_file_writes_a_table_when_a_streak_exceeds_the_limit(tmp_path, august_limit_30):
    output_path = _generate_payments_for_category_100(tmp_path, paid_days=31)
    check_file = tmp_path / "output" / "check.xlsx"

    violations, check_path = _check_category_100(tmp_path, output_path, check_file)

    assert len(violations) == 1
    assert check_path == str(check_file)
    assert load_workbook(check_file).worksheets[0].cell(row=2, column=2).value == "ПЕРШИЙ Перший Перший"


def test_category_100_table_takes_the_tasks_from_the_bchs_sheet(tmp_path, august_limit_30):
    """Рішення користувача: завдання - з аркуша БЧС файлів information_unit."""
    output_path = _generate_payments_for_category_100(tmp_path, paid_days=31)
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.title = "БЧС"
    ws.cell(row=2, column=4, value="ПРІЗВИЩЕ Ім'я по Батькові")
    ws.cell(row=3, column=9, value="Завдання яке виконує військовослужбовець")
    ws.cell(row=3, column=11, value="Бойове завдання яке виконує військовослужбовець")
    ws.cell(row=4, column=4, value="ПЕРШИЙ Перший Перший")
    ws.cell(row=4, column=9, value="Чергування")
    ws.cell(row=4, column=11, value="Оборона позиції")
    wb.save(str(tmp_path / "information_unit" / "БЧС рота.xlsx"))
    check_file = tmp_path / "output" / "check.xlsx"

    _check_category_100(tmp_path, output_path, check_file)

    report = load_workbook(check_file).worksheets[0]
    assert [report.cell(row=2, column=col).value for col in (3, 4)] == ["Чергування", "Оборона позиції"]


def test_category_100_is_not_duplicated_in_the_payments_review_file(tmp_path, august_limit_30):
    """Рішення користувача: порушення категорії 100 - лише в окремому файлі
    перевірки (index.py пише його одразу після генерації), а не ще й у
    "Потребує_ручної_перевірки_виплати.txt"."""
    output_path = _generate_payments_for_category_100(tmp_path, paid_days=31)
    review_file = tmp_path / "output" / "review.txt"
    check_file = tmp_path / "output" / "check.xlsx"

    violations, _check_path = _check_category_100(tmp_path, output_path, check_file)

    assert len(violations) == 1
    assert check_file.exists()
    assert not review_file.exists() or "категорія виплати 100" not in review_file.read_text(encoding="utf-8")


def test_check_payment_category_100_file_writes_nothing_when_within_the_limit(tmp_path):
    output_path = _generate_payments_for_category_100(tmp_path, paid_days=5)
    check_file = tmp_path / "output" / "check.xlsx"

    violations, check_path = _check_category_100(tmp_path, output_path, check_file)

    assert violations == []
    assert check_path is None
    assert not check_file.exists()


def _generate_payments_with_rtgr(tmp_path, groups, submitted):
    """ПЕРШИЙ на 01.08 у РТГр, ДРУГИЙ - звичайний РВЗ на 30 (без нього не було б
    кого згадати в рапорті, коли ПЕРШОГО там немає). submitted - що подає
    підрозділ на ПЕРШОГО за 01.08 і 02.08. Повертає (ОБЛІК, шлях звіту чи None,
    шлях повідомлення)."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "02.08.2026 - щоденний рапорт.docx", groups=groups)
    dates = [datetime(2026, 8, 1), datetime(2026, 8, 2)]
    roster_path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РТГр", None],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "РВЗ", None],
    ], header=["ПОСАДА", "ЗВАННЯ", "ПІБ"] + dates)
    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_info_unit_file(info_unit_dir / "1 рота.xlsx", rows=[
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", *submitted],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", 30, 30],
    ], dates=dates)
    output_path, _unresolved, report_path = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(tmp_path / "output" / "ОБЛІК.xlsx"),
        review_file_name=str(tmp_path / "output" / "review.txt"),
        mismatch_report_file_name=str(tmp_path / "output" / "error_mis_statuses.xlsx"),
        mismatch_message_file_name=str(tmp_path / "output" / "message.txt"),
    )
    return load_workbook(output_path)["Табель"], report_path, tmp_path / "output" / "message.txt"


_SOMEONE_ELSE_RETURNED = [(
    "ПРИБУЛИ до складу сил та засобів корпусу A0000:",
    [("З відпустки:", ARRIVAL_HEADER, [
        ["1", "матрос", "ДРУГИЙ Другий Другий", "Посада", "02.08.2026", "Прибув в район виконання завдань з відпустки"],
    ])],
)]


def test_return_from_rtgr_in_the_report_gives_a_payment_and_no_mismatch(tmp_path):
    """Реальний випадок (рапорт 14.10): "З РТГр:" закінчує статус, далі в ОБЛІК
    виплата від підрозділу, тож звіт розбіжностей не потрібен."""
    groups = [(
        "ПРИБУЛИ до складу сил та засобів корпусу A0000:",
        [("З РТГр:", ARRIVAL_HEADER, [
            ["1", "сержант", "ПЕРШИЙ Перший Перший", "Посада", "02.08.2026", "Прибув в район виконання завдань з РТГр"],
        ])],
    )]

    sheet, report_path, message_path = _generate_payments_with_rtgr(tmp_path, groups, submitted=("РТГР", 30))

    assert (sheet.cell(row=2, column=4).value, sheet.cell(row=2, column=5).value) == ("РТГр", 30)
    assert report_path is None
    assert not message_path.exists()


def test_rtgr_without_a_return_in_the_report_goes_to_the_mismatch_report(tmp_path):
    """Рішення користувача: у рапортах людина й далі в РТГр, а підрозділ уже
    подає виплату. Статус у ОБЛІК лишається, а розбіжність видно у звіті (РТГр
    проти РВЗ) і в повідомленні командиру."""
    sheet, report_path, message_path = _generate_payments_with_rtgr(tmp_path, _SOMEONE_ELSE_RETURNED, submitted=("РТГР", 30))

    assert (sheet.cell(row=2, column=4).value, sheet.cell(row=2, column=5).value) == ("РТГр", "РТГр")
    assert report_path is not None
    report = load_workbook(report_path)["Розбіжності"]
    assert report.cell(row=2, column=5).value == "ПЕРШИЙ Перший Перший"
    august_2 = next(c for c in range(6, report.max_column + 1) if report.cell(row=1, column=c).value == datetime(2026, 8, 2))
    assert (report.cell(row=2, column=august_2).value, report.cell(row=3, column=august_2).value) == ("РТГр", "РВЗ")
    assert "02.08.2026 статус має бути (РТГр)" in message_path.read_text(encoding="utf-8")


def test_rtgr_agreed_by_the_report_and_the_unit_is_not_reported(tmp_path):
    sheet, report_path, message_path = _generate_payments_with_rtgr(tmp_path, _SOMEONE_ELSE_RETURNED, submitted=("РТГР", "РТГР"))

    assert (sheet.cell(row=2, column=4).value, sheet.cell(row=2, column=5).value) == ("РТГр", "РТГр")
    assert report_path is None
    assert not message_path.exists()


_AUGUST_DAYS = [datetime(2026, 8, day) for day in (1, 2, 3)]
_ROSTER_ROWS = [
    ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВД", None, None],
    ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "ППД", None, None],
    ["Посада3", "матрос", "ТРЕТІЙ Третій Третій", "РВЗ", None, None],
]


def _generate_payments_for_addon_check(tmp_path):
    """ОБЛІК для виплат: ПЕРШИЙ - ВД і ДРУГИЙ - ППД (обидва стають числом 10),
    ТРЕТІЙ - РВЗ на 100. Запис є лише про ПЕРШОГО: ВД з кінцем за планом 02.08.
    Повертає (ОБЛІК, book61, аргументи перерахунку статусів)."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "03.08.2026 - щоденний рапорт.docx", groups=[(
        "ПРИБУЛИ до складу сил та засобів корпусу A0000:",
        [("З відпустки:", ARRIVAL_HEADER, [
            ["1", "матрос", "ТРЕТІЙ Третій Третій", "Посада", "03.08.2026", "Прибув в район виконання завдань з відпустки"],
        ])],
    )])
    roster_path = _write_roster(tmp_path / "ОБЛІК.xlsx", _ROSTER_ROWS, header=["ПОСАДА", "ЗВАННЯ", "ПІБ"] + _AUGUST_DAYS)
    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_info_unit_file(info_unit_dir / "1 рота.xlsx", rows=[
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВД", "ВД", "ВД"],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", 10, 10, 10],
        ["Посада3", "матрос", "ТРЕТІЙ Третій Третій", 100, 100, 100],
    ], dates=_AUGUST_DAYS)
    output_path, _unresolved, _mismatch_report = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(tmp_path / "output" / "ОБЛІК.xlsx"),
        review_file_name=str(tmp_path / "output" / "review.txt"),
        mismatch_report_file_name=str(tmp_path / "output" / "error_mis_statuses.xlsx"),
        mismatch_message_file_name=str(tmp_path / "output" / "message.txt"),
    )
    book61 = write_full_book61(tmp_path / "book61.xlsm", {"ВД": [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "вдр", date(2026, 7, 14), None, None, 10, date(2026, 8, 2)],
    ]})
    return output_path, book61, {
        "report_dir": str(report_dir), "personel_list_file_name": roster_path, "information_unit_dir": str(info_unit_dir),
    }


def test_addon_check_of_the_payments_file_sees_statuses_behind_numbers(tmp_path):
    """Реальний випадок: людина у ВД до "Дата кінця" в book61, далі запису
    немає, а в ОБЛІК для виплат - 10 до кінця місяця. Так само ППД без жодного
    запису. РВЗ на 100 запису не потребує."""
    output_path, book61, status_sources = _generate_payments_for_addon_check(tmp_path)
    assert [cell.value for cell in load_workbook(output_path)["Табель"][2]][3:] == [10, 10, 10]
    report_path = str(tmp_path / "output" / "mis_match.xlsx")

    result = check_addon_mismatches_file(output_path, [book61], report_file_name=report_path, **status_sources)

    august = lambda day: date(2026, 8, day)
    assert {record["roster_pib_raw"]: record["dates"] for record in result.period_mismatches} == {
        "ПЕРШИЙ Перший Перший": {
            august(1): (10, "ВД"), august(2): (10, "ВД"), august(3): ("ВД", "немає запису на аркуші ВД"),
        },
        "ДРУГИЙ Другий Другий": {august(day): ("ППД", "немає запису на аркуші ППД") for day in (1, 2, 3)},
    }
    report = load_workbook(report_path)["book61"]
    first_row = next(row for row in range(2, report.max_row + 1) if report.cell(row=row, column=3).value == "ПЕРШИЙ Перший Перший")
    august_1, august_3 = (next(c for c in range(4, report.max_column + 1) if report.cell(row=1, column=c).value == day)
                          for day in (_AUGUST_DAYS[0], _AUGUST_DAYS[2]))
    assert report.cell(row=first_row, column=august_1).value == "ВД"  # не 10: звіт читають за статусами
    assert report.cell(row=first_row, column=august_1).fill.start_color.rgb == _MATCH_FILL.start_color.rgb
    assert (report.cell(row=first_row, column=august_3).value,
            report.cell(row=first_row + 1, column=august_3).value) == ("ВД", "немає запису на аркуші ВД")
    assert report.cell(row=first_row, column=august_3).fill.start_color.rgb == _MISMATCH_FILL.start_color.rgb


def test_addon_check_does_not_recount_statuses_unless_asked(tmp_path):
    """Без report_dir статусів за числами немає, тож index.py передає його завжди."""
    output_path, book61, _status_sources = _generate_payments_for_addon_check(tmp_path)

    result = check_addon_mismatches_file(output_path, [book61], report_file_name=str(tmp_path / "output" / "mis_match.xlsx"))

    assert result.period_mismatches == []


def test_statuses_behind_numbers_follow_the_person_not_the_row(tmp_path):
    """Файл для звірки користувач може обрати сам: рядки в ньому йдуть інакше,
    ніж у resources/ОБЛІК.xlsx, а людини з нього може там не бути."""
    _output_path, book61, status_sources = _generate_payments_for_addon_check(tmp_path)
    picked = _write_roster(tmp_path / "picked.xlsx", [
        ["Посада4", "матрос", "ЧЕТВЕРТИЙ Четвертий Четвертий", 10, 10, 10],
        ["Посада3", "матрос", "ТРЕТІЙ Третій Третій", 100, 100, 100],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", 10, 10, 10],
    ], header=["ПОСАДА", "ЗВАННЯ", "ПІБ"] + _AUGUST_DAYS)

    result = check_addon_mismatches_file(picked, [book61], report_file_name=str(tmp_path / "output" / "mis_match.xlsx"),
                                         **status_sources)

    assert [record["roster_pib_raw"] for record in result.period_mismatches] == ["ДРУГИЙ Другий Другий"]


def _generate_payments_with_evening_arrival(tmp_path, submitted):
    """ПЕРШИЙ прибув з ВЛК "01.08.2026 після 17.00" (рапорт 02.08), підрозділ
    подав за 31.07-03.08 значення submitted. Рапорт 03.08 - про ДРУГОГО, щоб
    наступний день мав рапорт. Повертає (ОБЛІК, звіт, повідомлення, book61,
    аргументи перерахунку)."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "02.08.2026 - щоденний рапорт.docx", groups=[(
        "ПРИБУЛИ до складу сил та засобів корпусу A0000:",
        [("З проходження військово-лікарської комісії:", ARRIVAL_HEADER, [
            ["1", "матрос", "ПЕРШИЙ Перший Перший", "Посада", "01.08.2026 після 17.00",
             "Прибув в район виконання завдань з військово-лікарської комісії військової частини A0000"],
        ])],
    )])
    write_daily_report(report_dir / "03.08.2026 - щоденний рапорт.docx", groups=[(
        "ПРИБУЛИ до складу сил та засобів корпусу A0000:",
        [("З відпустки:", ARRIVAL_HEADER, [
            ["1", "матрос", "ДРУГИЙ Другий Другий", "Посада", "03.08.2026", "Прибув в район виконання завдань з відпустки"],
        ])],
    )])
    days = [datetime(2026, 7, 31), *_AUGUST_DAYS]
    roster_path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВЛК", None, None, None],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "РВЗ", None, None, None],
    ], header=["ПОСАДА", "ЗВАННЯ", "ПІБ"] + days)
    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_info_unit_file(info_unit_dir / "1 рота.xlsx", rows=[
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВЛК", submitted, 30, 30],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", 30, 30, 30, 30],
    ], dates=days)
    output_dir = tmp_path / "output"
    output_path, _unresolved, report_path = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(output_dir / "ОБЛІК.xlsx"),
        review_file_name=str(output_dir / "review.txt"),
        mismatch_report_file_name=str(output_dir / "error_mis_statuses.xlsx"),
        mismatch_message_file_name=str(output_dir / "message.txt"),
    )
    book61 = write_full_book61(tmp_path / "book61.xlsm", {"ВЛК": [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 7, 20), date(2026, 8, 1)],
    ]})
    return output_path, report_path, output_dir / "message.txt", book61, {
        "report_dir": str(report_dir), "personel_list_file_name": roster_path, "information_unit_dir": str(info_unit_dir),
    }


def test_unit_paying_a_duty_day_for_an_evening_arrival_is_a_mismatch(tmp_path):
    """Реальний випадок: людина прибула з ВЛК "після 17.00", а підрозділ подав
    за цей день 30. У ОБЛІК день - ППД (10), далі 30, а командир бачить, що
    статус цього дня має бути ППД."""
    output_path, report_path, message_path, _book61, _sources = _generate_payments_with_evening_arrival(tmp_path, 30)

    assert [cell.value for cell in load_workbook(output_path)["Табель"][2]][4:] == [10, 30, 30]
    assert report_path is not None
    assert "01.08.2026 статус має бути (ППД)" in message_path.read_text(encoding="utf-8")


def test_unit_submitting_ppd_for_an_evening_arrival_is_not_reported(tmp_path):
    """Реальний випадок: підрозділ сам подав ППД на день прибуття "після 17.00"."""
    output_path, report_path, message_path, _book61, _sources = _generate_payments_with_evening_arrival(tmp_path, "ППД")

    assert [cell.value for cell in load_workbook(output_path)["Табель"][2]][4:] == [10, 30, 30]
    assert report_path is None
    assert not message_path.exists()


def test_addon_check_takes_the_evening_arrival_from_the_reports(tmp_path):
    """Книга веде лише ВЛК до дня прибуття, а ППД цього дня дає рапорт. Без
    рапортів перевірка цього не знає і показує день як розбіжність з ВЛК."""
    output_path, _report_path, _message_path, book61, status_sources = _generate_payments_with_evening_arrival(tmp_path, "ППД")
    report_path = str(tmp_path / "output" / "mis_match.xlsx")

    checked = check_addon_mismatches_file(output_path, [book61], report_file_name=report_path, **status_sources)
    unchecked = check_addon_mismatches_file(output_path, [book61], report_file_name=report_path)

    assert checked.period_mismatches == []
    assert [record["dates"][date(2026, 8, 1)] for record in unchecked.period_mismatches] == [(10, "ВЛК")]


def test_evening_arrival_follows_the_person_not_the_row(tmp_path):
    """Файл для звірки може мати інший порядок рядків (так само, як і в
    test_statuses_behind_numbers_follow_the_person_not_the_row)."""
    _output_path, _report_path, _message_path, book61, status_sources = _generate_payments_with_evening_arrival(tmp_path, "ППД")
    picked = _write_roster(tmp_path / "picked.xlsx", [
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", 30, 30, 30, 30],
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВЛК", 10, 30, 30],
    ], header=["ПОСАДА", "ЗВАННЯ", "ПІБ", datetime(2026, 7, 31), *_AUGUST_DAYS])

    result = check_addon_mismatches_file(picked, [book61], report_file_name=str(tmp_path / "output" / "mis_match.xlsx"),
                                         **status_sources)

    assert result.period_mismatches == []


def test_addon_check_without_reports_warns_and_still_compares(tmp_path, capsys):
    output_path, book61, status_sources = _generate_payments_for_addon_check(tmp_path)
    empty_dir = tmp_path / "empty"
    empty_dir.mkdir()

    result = check_addon_mismatches_file(output_path, [book61], report_file_name=str(tmp_path / "output" / "mis_match.xlsx"),
                                         **{**status_sources, "report_dir": str(empty_dir)})

    assert result.checked_files == [book61]
    assert result.period_mismatches == []
    assert "не знайдено жодного .docx рапорту" in capsys.readouterr().out


def _generate_payments_with_medical_category(tmp_path):
    """ОБЛІК для виплат: у ПЕРШОГО ВЛК, у ДРУГОГО ШП, в обох 10_МЕД. Записів у
    журналі лікування немає."""
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    write_daily_report(report_dir / "03.08.2026 - щоденний рапорт.docx", groups=[(
        "ПРИБУЛИ до складу сил та засобів корпусу A0000:",
        [("З відпустки:", ARRIVAL_HEADER, [
            ["1", "матрос", "ТРЕТІЙ Третій Третій", "Посада", "03.08.2026", "Прибув в район виконання завдань з відпустки"],
        ])],
    )])
    roster_path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВЛК", None, None],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "ШП", None, None],
        ["Посада3", "матрос", "ТРЕТІЙ Третій Третій", "РВЗ", None, None],
    ], header=["ПОСАДА", "ЗВАННЯ", "ПІБ"] + _AUGUST_DAYS)
    info_unit_dir = tmp_path / "information_unit"
    info_unit_dir.mkdir()
    _write_info_unit_file(info_unit_dir / "1 рота.xlsx", rows=[
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВЛК", "ВЛК", "ВЛК"],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "ШП", "ШП", "ШП"],
        ["Посада3", "матрос", "ТРЕТІЙ Третій Третій", 100, 100, 100],
    ], dates=_AUGUST_DAYS)
    output_dir = tmp_path / "output"
    output_path, _unresolved, _report = generate_payments_timesheet(
        report_dir=str(report_dir),
        personel_list_file_name=roster_path,
        personel_list_sheet_name="Табель",
        information_unit_dir=str(info_unit_dir),
        output_file_name=str(output_dir / "ОБЛІК.xlsx"),
        review_file_name=str(output_dir / "review.txt"),
        mismatch_report_file_name=str(output_dir / "error_mis_statuses.xlsx"),
        mismatch_message_file_name=str(output_dir / "message.txt"),
    )
    journal = write_treatment_journal(tmp_path / "журнал.xlsx", {
        "2026": [("ТРЕТІЙ Третій Третій", datetime(2026, 6, 1), datetime(2026, 6, 2))],
    })
    return output_path, journal, {
        "report_dir": str(report_dir), "personel_list_file_name": roster_path, "information_unit_dir": str(info_unit_dir),
    }


def test_journal_check_of_the_payments_file_tells_vlk_from_treatment(tmp_path):
    """У готовому файлі в обох 10_МЕД. Журнал лікування стосується лише ШП, а
    ВЛК він не веде; у розбіжності показується статус, а не категорія."""
    output_path, journal, status_sources = _generate_payments_with_medical_category(tmp_path)
    tabel = load_workbook(output_path)["Табель"]
    assert [cell.value for cell in tabel[2]][3:] == [cell.value for cell in tabel[3]][3:]

    result = check_addon_mismatches_file(output_path, [journal], report_file_name=str(tmp_path / "output" / "mis_match.xlsx"),
                                         **status_sources)

    (record,) = result.period_mismatches
    assert record["roster_pib_raw"] == "ДРУГИЙ Другий Другий"
    assert set(record["dates"].values()) == {("ШП", "немає запису в журналі лікування")}
    assert set(record["tabel_values"].values()) == {"ШП"}


def test_statuses_behind_categories_follow_the_person_not_the_row(tmp_path):
    _output_path, journal, status_sources = _generate_payments_with_medical_category(tmp_path)
    picked = _write_roster(tmp_path / "picked.xlsx", [
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "10_МЕД", "10_МЕД", "10_МЕД"],
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "10_МЕД", "10_МЕД", "10_МЕД"],
    ], header=["ПОСАДА", "ЗВАННЯ", "ПІБ"] + _AUGUST_DAYS)

    result = check_addon_mismatches_file(picked, [journal], report_file_name=str(tmp_path / "output" / "mis_match.xlsx"),
                                         **status_sources)

    assert [record["roster_pib_raw"] for record in result.period_mismatches] == ["ДРУГИЙ Другий Другий"]
