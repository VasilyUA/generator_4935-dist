from datetime import date, datetime

import pytest
import xlwt
from openpyxl import Workbook, load_workbook

import content.payment_mismatch_checker as pmc
from content.oblik_timesheet import Timesheet, normalize_name

_XLS_DATE_STYLE = xlwt.easyxf(num_format_str="DD.MM.YYYY")


def _write_roster(path, rows, header=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Табель"
    header = header or ["ПОСАДА", "ЗВАННЯ", "ПІБ", datetime(2026, 7, 31), datetime(2026, 8, 1), datetime(2026, 8, 2)]
    for col_idx, value in enumerate(header, start=1):
        ws.cell(row=1, column=col_idx, value=value)
    for row_idx, row_values in enumerate(rows, start=2):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    wb.save(str(path))
    return str(path)


@pytest.fixture
def roster(tmp_path):
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "ВП", "РВЗ", "РВЗ"],
    ])
    return Timesheet(path, "Табель")


@pytest.fixture
def roster_with_pidrozdil(tmp_path):
    path = _write_roster(
        tmp_path / "ОБЛІК.xlsx",
        [["підрозділ 1", "Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", "РВЗ", "РВЗ"]],
        header=["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", datetime(2026, 7, 31), datetime(2026, 8, 1), datetime(2026, 8, 2)],
    )
    return Timesheet(path, "Табель")


# -------------------------
# find_mismatches
# -------------------------

def test_find_mismatches_skips_when_file_reports_same_default_status(roster):
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", "РВЗ"),
        ],
    }

    assert pmc.find_mismatches(roster, snapshot) == []


def test_find_mismatches_detects_numeric_value_mismatch(roster):
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["file_name"] == "файл1.xlsx"
    assert records[0]["roster_pib_raw"] == "ПЕРШИЙ Перший Перший"
    assert records[0]["file_pib_raw"] == "ПЕРШИЙ Перший Перший"
    assert records[0]["dates"] == {date(2026, 8, 1): ("РВЗ", 100)}


def test_find_mismatches_detects_non_default_status_text_mismatch(roster):
    """Реальний випадок: аркуш БЧС каже "ВД" замість числа - теж розбіжність,
    так само як число, що не збігається."""
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", "ВД"),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert records[0]["dates"] == {date(2026, 8, 1): ("РВЗ", "ВД")}


def test_find_mismatches_detects_mismatch_when_final_status_is_a_resolved_number(roster):
    """Реальний випадок: роcтер уже застосував категорію виплати (число, не
    "РВЗ" - apply_payment_values спрацював), а ЦЕЙ САМИЙ чи ІНШИЙ файл
    information_unit подає ІНШЕ число для тієї ж дати - це РЕАЛЬНА
    розбіжність (підрозділи не погоджуються між собою), незалежно від того,
    що фінальний статус давно не буквально "РВЗ" - порівняння НЕ обмежується
    лише текстовим статусом "РВЗ", а й уже застосованими числами."""
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 200),
        ],
    }
    roster.ws.cell(row=2, column=5).value = 100  # 01.08 - вже застосована категорія 100 (не "РВЗ")

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["dates"] == {date(2026, 8, 1): (100, 200)}


def test_find_mismatches_ignores_dates_with_a_non_comparable_status(roster):
    """Реальний випадок: людина в СЗЧ - подальше звітування підрозділу про
    категорію виплати НЕ суперечить цьому статусу (людина або й далі
    отримує виплату на загальних підставах, або дані підрозділу вже
    застаріли) - не повинно породжувати запис у звіті розбіжностей,
    незалежно від того, що каже файл information_unit для цієї дати.
    "ВП"/"ВЛК"/"ВПЗС"/"ВПСЗ"/"ШП" тут НЕ перевіряються - вони, на відміну
    від СЗЧ, ТЕПЕР у PAYMENT_MISMATCH_COMPARABLE_STATUSES (див. тест
    нижче)."""
    roster.ws.cell(row=2, column=5).value = "СЗЧ"  # 01.08
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
        ],
    }

    assert pmc.find_mismatches(roster, snapshot) == []


@pytest.mark.parametrize("status", ["ВЛК", "ВП", "ВПЗС", "ВПСЗ", "ШП"])
def test_find_mismatches_detects_disagreement_on_vlk_or_vp(roster, status):
    """Реальні випадки: щоденний рапорт показує "ВЛК" і, наступного дня,
    "ВП" (типова послідовність "виписаний... ВЛК... потребує відпустки..."),
    чи "ВПЗС" (окремий статус, НЕ те саме, що "ВПСЗ"), чи "ВПСЗ" (вибуття у
    відпустку за станом здоров'я), чи "ШП" (день госпіталізації, коли
    PAYMENT_STATUS_OVERRIDES["ШП"] НЕ спрацював, бо information_unit подав
    НЕ "БПШП" - просто число), а файл information_unit ОДНОСТАЙНО й далі
    каже число (категорія виплати) для ТИХ САМИХ дат - підозріла
    суперечність, варта ручної перевірки, тож "ВЛК"/"ВП"/"ВПЗС"/"ВПСЗ"/"ШП"
    (на відміну від СЗЧ вище) - у PAYMENT_MISMATCH_COMPARABLE_STATUSES
    (constants.py)."""
    roster.ws.cell(row=2, column=5).value = status  # 01.08
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["dates"] == {date(2026, 8, 1): (status, 100)}


@pytest.mark.parametrize("roster_status", ["100_БЗ", "100_БПШП", "100_ВПБП"])
def test_find_mismatches_detects_disagreement_on_converted_payment_text(roster, roster_status):
    """Реальний випадок: "БЗ" вручну вписано в resources/ОБЛІК.xlsx (не з
    рапорту), apply_payment_values ОДРАЗУ конвертує в буквальний текст
    "100_БЗ" (PAYMENT_STATUS_CONVERSIONS) - а файл information_unit подає
    ЩОСЬ ІНШЕ (напр. число - людина насправді НЕ "БЗ") - підозріла
    суперечність, тож "100_БЗ"/"100_БПШП"/"100_ВПБП" (ТІ САМІ конвертовані
    значення, що й для "БПШП"/"ВПБП") - теж у
    PAYMENT_MISMATCH_COMPARABLE_STATUSES (constants.py)."""
    roster.ws.cell(row=2, column=5).value = roster_status  # 01.08
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 30),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["dates"] == {date(2026, 8, 1): (roster_status, 30)}


@pytest.mark.parametrize("roster_status, information_unit_status", [
    ("100_БЗ", "БЗ"), ("100_БЗ", "полон"), ("100_БЗ", "Інт"),
    ("100_БПШП", "БПШП"), ("100_ВПБП", "ВПБП"),
])
def test_find_mismatches_treats_converted_payment_text_as_equivalent_to_its_source_status(roster, roster_status, information_unit_status):
    """"100_БЗ"/"100_БПШП"/"100_ВПБП" ставши "порівнюваними" - НЕ повинні
    хибно позначати розбіжність, коли information_unit подає ТОЙ САМИЙ факт
    іншим написанням (вихідний статус, з якого й вийшла ця конвертація,
    PAYMENT_STATUS_CONVERSIONS) - _statuses_equivalent (generic перебір
    PAYMENT_STATUS_CONVERSIONS) і далі визнає їх рівними."""
    roster.ws.cell(row=2, column=5).value = roster_status  # 01.08
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", information_unit_status),
        ],
    }

    assert pmc.find_mismatches(roster, snapshot) == []


@pytest.mark.parametrize("oblik_status, information_unit_status, should_match", [
    ("ВП", "ВП", True),
    ("ВП", "ВПС", False),
    ("ВП", "ВПСЗ", False),
    ("ВП", "ВПБП", False),
    ("ВПСЗ", "ВПБП", False),
])
def test_find_mismatches_compares_vp_and_vpsz_text_statuses_against_each_other(roster, oblik_status, information_unit_status, should_match):
    """Реальні випадки, підтверджені користувачем - повна таблиця
    порівнянь для "ВП"/"ВПСЗ" проти ІНШИХ текстових статусів
    information_unit (не лише проти числа, як у тесті вище): збігається
    лише "ВП" з "ВП" (те саме) - решта пар (навіть спорідненого "родини
    відпусток" значення - "ВПС"/"ВПСЗ"/"ВПБП") - РІЗНІ реальні причини, тож
    РОЗБІЖНІСТЬ."""
    roster.ws.cell(row=2, column=5).value = oblik_status  # 01.08
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", information_unit_status),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    if should_match:
        assert records == []
    else:
        assert len(records) == 1
        assert records[0]["dates"] == {date(2026, 8, 1): (oblik_status, information_unit_status)}


def test_find_mismatches_includes_matching_dates_alongside_a_real_mismatch(roster):
    """Реальний випадок: файл information_unit має розбіжність лише на ОДНУ
    дату, але збігається з роcтером на ІНШУ - файл, у якого Є хоч ОДНА
    справжня розбіжність, показує ОБИДВІ дати (не лише розбіжну) - користувач
    не повинен бачити порожню клітинку там, де файл насправді дає збіжне
    значення."""
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
        ],
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 2)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", "РВЗ"),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["dates"] == {
        date(2026, 8, 1): ("РВЗ", 100),
        date(2026, 8, 2): ("РВЗ", "РВЗ"),
    }


def test_find_mismatches_excludes_file_with_only_matching_dates(roster):
    """Файл, що ПОВНІСТЮ погоджується з роcтером на всіх датах, де є дані -
    НЕМАЄ розбіжності, тож у звіт не потрапляє взагалі (не лише порожній
    рядок без жодної розбіжності)."""
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", "РВЗ"),
        ],
    }

    assert pmc.find_mismatches(roster, snapshot) == []


def test_find_mismatches_ignores_dates_where_final_value_is_none(roster_with_pidrozdil):
    """Дата-колонка роcтера ВЗАГАЛІ без значення (None) - нема з чим
    порівнювати, порівняння не відбувається (_is_comparable_status(None) є
    False, той самий шлях, що й для інших "непорівнюваних" статусів)."""
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 2)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", "підрозділ 1", "Посада1", "сержант", 100),
        ],
    }
    # roster_with_pidrozdil - 3 дати-колонки (31.07/01.08/02.08), усі "РВЗ" -
    # приберемо значення з 02.08, щоб перевірити None-гілку.
    col_by_date = {d: c for c, d in roster_with_pidrozdil.date_columns}
    roster_with_pidrozdil.ws.cell(row=2, column=col_by_date[date(2026, 8, 2)]).value = None

    assert pmc.find_mismatches(roster_with_pidrozdil, snapshot) == []


def test_find_mismatches_ignores_person_not_in_snapshot(roster):
    assert pmc.find_mismatches(roster, {}) == []


def test_find_mismatches_groups_multiple_mismatching_dates_into_one_record(roster):
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
        ],
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 2)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 30),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["dates"] == {
        date(2026, 8, 1): ("РВЗ", 100),
        date(2026, 8, 2): ("РВЗ", 30),
    }


def test_find_mismatches_creates_separate_records_for_different_files(roster):
    """Розбіжність з ДВОМА різними файлами - ДВА окремих записи (файл ЗАВЖДИ
    рівно один на запис), а не один запис на кількох файлів."""
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
            ("файл2.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 30),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 2
    assert sorted(r["file_name"] for r in records) == ["файл1.xlsx", "файл2.xlsx"]


def test_find_mismatches_tracks_both_roster_and_file_pib_raw(roster):
    """Реальний випадок: файл information_unit подає ПІБ, СХОЖИЙ, але не
    ІДЕНТИЧНИЙ (напр. зайвий пробіл) на роcтерове написання - обидва
    варіанти написання мають зберігатись окремо (roster_pib_raw/
    file_pib_raw), щоб write_mismatch_report міг порівняти й показати
    розбіжність, якщо вона є."""
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший ", None, "Посада1", "сержант", 100),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert records[0]["roster_pib_raw"] == "ПЕРШИЙ Перший Перший"
    assert records[0]["file_pib_raw"] == "ПЕРШИЙ Перший Перший "


def test_find_mismatches_includes_roster_and_file_pidrozdil(roster_with_pidrozdil):
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", "підрозділ 1", "Посада1", "сержант", 100),
        ],
    }

    records = pmc.find_mismatches(roster_with_pidrozdil, snapshot)

    assert records[0]["roster_pidrozdil"] == "підрозділ 1"
    assert records[0]["file_pidrozdil"] == "підрозділ 1"


def test_find_mismatches_reports_none_roster_pidrozdil_when_roster_has_no_such_column(roster):
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", "підрозділ 1", "Посада1", "сержант", 100),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert records[0]["roster_pidrozdil"] is None
    assert records[0]["file_pidrozdil"] == "підрозділ 1"


def test_find_mismatches_includes_people_present_only_in_information_unit(roster):
    """Реальний випадок: людина є в файлі information_unit, але ВЗАГАЛІ
    відсутня в ОБЛІК для виплат.xlsx (не в ростері) - раніше мовчки зникала
    зі звіту (основний прохід іде ПО РОСТЕРУ), тепер має зʼявитись."""
    snapshot = {
        (normalize_name("ТРЕТІЙ Третій Третій"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ТРЕТІЙ Третій Третій", "підрозділ 1", "Посада3", "матрос", 100),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["file_name"] == "файл1.xlsx"
    assert records[0]["roster_pib_raw"] == pmc._NOT_IN_ROSTER_LABEL
    assert records[0]["file_pib_raw"] == "ТРЕТІЙ Третій Третій"
    assert records[0]["roster_pidrozdil"] == pmc._NOT_IN_ROSTER_LABEL
    assert records[0]["roster_posada"] == pmc._NOT_IN_ROSTER_LABEL
    assert records[0]["roster_zvannya"] == pmc._NOT_IN_ROSTER_LABEL
    assert records[0]["file_pidrozdil"] == "підрозділ 1"
    assert records[0]["file_posada"] == "Посада3"
    assert records[0]["dates"] == {date(2026, 8, 1): (pmc._NOT_IN_ROSTER_LABEL, 100)}


def test_find_mismatches_groups_missing_person_dates_across_multiple_files(roster):
    snapshot = {
        (normalize_name("ТРЕТІЙ Третій Третій"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ТРЕТІЙ Третій Третій", None, "Посада3", "матрос", 100),
        ],
        (normalize_name("ТРЕТІЙ Третій Третій"), date(2026, 8, 2)): [
            ("файл1.xlsx", "ТРЕТІЙ Третій Третій", None, "Посада3", "матрос", 30),
            ("файл2.xlsx", "ТРЕТІЙ Третій Третій", None, "Посада3", "матрос", 100),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 2
    by_file = {r["file_name"]: r for r in records}
    assert by_file["файл1.xlsx"]["dates"] == {
        date(2026, 8, 1): (pmc._NOT_IN_ROSTER_LABEL, 100),
        date(2026, 8, 2): (pmc._NOT_IN_ROSTER_LABEL, 30),
    }
    assert by_file["файл2.xlsx"]["dates"] == {date(2026, 8, 2): (pmc._NOT_IN_ROSTER_LABEL, 100)}


def test_find_mismatches_dedupes_files_with_identical_content_for_missing_from_roster_person(roster):
    """Той самий дедуплікаційний принцип (test_find_mismatches_dedupes_files_
    with_identical_content), але для людини, ВІДСУТНЬОЇ в роcтері
    (_missing_from_roster_records) - ДВА файли з ідентичними даними теж
    мають зводитись до ОДНОГО запису."""
    snapshot = {
        (normalize_name("ТРЕТІЙ Третій Третій"), date(2026, 8, 1)): [
            ("файл_б.xlsx", "ТРЕТІЙ Третій Третій", "підрозділ 1", "Посада3", "матрос", 100),
            ("файл_а.xlsx", "ТРЕТІЙ Третій Третій", "підрозділ 1", "Посада3", "матрос", 100),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["file_name"] == "файл_а.xlsx"


def test_surname_firstname_key_returns_first_two_words():
    assert pmc._surname_firstname_key("ПЕРШИЙ ПЕРШИЙ ПЕРШИЙ") == "ПЕРШИЙ ПЕРШИЙ"


def test_surname_firstname_key_returns_whole_string_when_single_word():
    assert pmc._surname_firstname_key("ПЕРШИЙ") == "ПЕРШИЙ"


def test_find_mismatches_shows_real_roster_pib_when_only_patronymic_differs(roster):
    """Реальний випадок: файл information_unit подає ТЕ САМЕ
    прізвище+ім'я, але ІНШЕ по-батькові (друкарська помилка) -
    normalize_name дає інший ключ, тож людину раніше показували як
    "відсутню в ОБЛІК.xlsx" (загальний напис), хоча в ростері Є ЄДИНИЙ
    схожий кандидат - роcтерова сторона тепер має показати РЕАЛЬНЕ ПІБ і
    поточний статус цього кандидата, щоб розбіжність у написанні було видно
    поряд, а не ховалась за загальним написом."""
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перекший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перекший", None, "Посада1", "сержант", 100),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["roster_pib_raw"] == "ПЕРШИЙ Перший Перший"
    assert records[0]["file_pib_raw"] == "ПЕРШИЙ Перший Перекший"
    assert records[0]["roster_posada"] == "Посада1"
    assert records[0]["roster_zvannya"] == "сержант"
    assert records[0]["dates"] == {date(2026, 8, 1): ("РВЗ", 100)}


def test_find_mismatches_does_not_guess_when_multiple_roster_candidates_share_surname_and_firstname(tmp_path):
    """Двоє РІЗНИХ роcтерових людей з однаковим Прізвище+Ім'я (лише
    по-батькові різне) - неоднозначно, хто з них "той самий" - краще
    показати загальний напис "не знайдено", ніж вгадати неправильну
    людину."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада2", "матрос", "ПЕРШИЙ Перший Другший", "РВЗ", "РВЗ", "РВЗ"],
    ])
    roster_ambiguous = Timesheet(path, "Табель")
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Третший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Третший", None, "Посада3", "матрос", 100),
        ],
    }

    records = pmc.find_mismatches(roster_ambiguous, snapshot)

    assert len(records) == 1
    assert records[0]["roster_pib_raw"] == pmc._NOT_IN_ROSTER_LABEL


def test_find_mismatches_falls_back_to_placeholder_when_candidate_has_no_matching_date_column(roster):
    """Кандидат роcтера ЗНАЙДЕНИЙ (унікальний за Прізвище+Ім'я), але
    snapshot дата виходить за межі дат-колонок роcтера - для ЦІЄЇ дати
    роcтерове значення - загальний плейсхолдер, а не KeyError."""
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перекший"), date(2026, 9, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перекший", None, "Посада1", "сержант", 100),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert records[0]["dates"] == {date(2026, 9, 1): (pmc._NOT_IN_ROSTER_LABEL, 100)}


def test_find_mismatches_includes_matching_non_comparable_dates_alongside_a_real_mismatch(roster):
    """Реальний випадок (АНДРЕЇШИН/ТРОФІМОВ): дата з НЕ-порівнюваним статусом
    (напр. "ВПСЗ"), де файл ПОГОДЖУЄТЬСЯ з роcтером - раніше повністю
    пропускалась (навіть коли БУВ хоч ОДИН справжній mismatch на ІНШУ дату,
    що й так показує цей рядок звіту) - тепер має зʼявитись у "історії"
    файлу як звичайний збіг (зелена клітинка), а не порожня."""
    roster.ws.cell(row=2, column=5).value = "ВПСЗ"  # 01.08 - непорівнюваний, ЗБІГАЄТЬСЯ з файлом
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", "ВПСЗ"),
        ],
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 2)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),  # справжня розбіжність (РВЗ != 100)
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["dates"] == {
        date(2026, 8, 1): ("ВПСЗ", "ВПСЗ"),
        date(2026, 8, 2): ("РВЗ", 100),
    }


def test_find_mismatches_shows_non_comparable_disagreement_without_it_alone_triggering_inclusion(roster):
    """Не-порівнюваний статус, що НЕ збігається з файлом (напр. "ВЛК" vs
    100) - НЕ сам по собі створює запис (has_mismatch не спрацьовує через
    нього), але якщо рядок ВЖЕ показується через ІНШУ, справжню розбіжність
    - ця дата теж зʼявляється (як звичайна незбіжність - червона), для
    повної "історії" файлу."""
    roster.ws.cell(row=2, column=5).value = "ВЛК"  # 01.08 - непорівнюваний, НЕ збігається з файлом
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
        ],
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 2)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 999),  # справжня розбіжність
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["dates"] == {
        date(2026, 8, 1): ("ВЛК", 100),
        date(2026, 8, 2): ("РВЗ", 999),
    }


def test_find_mismatches_ignores_person_whose_only_dates_are_non_comparable_disagreements(roster):
    """Якщо ЄДИНА розбіжність людини - на непорівнюваній даті (напр. "СЗЧ"
    vs 100, без ЖОДНОЇ справжньої розбіжності на ІНШУ дату) - запис у звіт
    НЕ потрапляє взагалі (те саме, що test_find_mismatches_ignores_dates_
    with_a_non_comparable_status, але явно перевіряє, що dates-словник теж
    не "просочується" в порожній запис)."""
    roster.ws.cell(row=2, column=5).value = "СЗЧ"
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
        ],
    }

    assert pmc.find_mismatches(roster, snapshot) == []


def test_find_mismatches_treats_ppd_and_its_converted_number_as_equivalent(roster):
    """Реальний випадок: apply_payment_values (content/oblik_timesheet.py)
    ВЖЕ конвертував базовий статус "ППД" у фіксоване число 10
    (PAYMENT_STATUS_CONVERSIONS) з боку ОБЛІК для виплат, а файл
    information_unit - свій ВЛАСНИЙ, незалежний словник статусів - і далі
    каже текстом "ППД" для тієї самої дати. Це ТА САМА реальна ситуація,
    записана по-різному, а НЕ розбіжність."""
    roster.ws.cell(row=2, column=5).value = 10  # 01.08 - вже конвертоване ППД -> 10
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", "ППД"),
        ],
    }

    assert pmc.find_mismatches(roster, snapshot) == []


def test_find_mismatches_treats_bz_and_rozp_as_equivalent_via_priority(roster):
    """PAYMENT_STATUS_PRIORITY (constants.py, керовано користувачем
    самостійно) - реальний випадок, підтверджений користувачем: роcтерове
    ФІНАЛЬНЕ значення "100_БЗ" (вже конвертоване "БЗ" - зниклий безвісті) і
    "Розп" (розпорядження) з information_unit - "БЗ" МАЄ визначеного
    переможця над "Розп" у цій парі, тож ЦЕ НЕ розбіжність, а лише
    підтвердження того самого факту з нижчим пріоритетом - error_mis_
    statuses.xlsx НЕ повинен показувати цей запис."""
    roster.ws.cell(row=2, column=5).value = "100_БЗ"  # 01.08 - вже конвертоване БЗ
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", "Розп"),
        ],
    }

    assert pmc.find_mismatches(roster, snapshot) == []


@pytest.mark.parametrize("information_unit_status", ["БЗВП", "Адап"])
def test_find_mismatches_treats_bzvp_and_adap_and_their_converted_number_as_equivalent(roster, information_unit_status):
    """Той самий принцип, що й ППД/10 вище, але через PAYMENT_STATUS_
    OVERRIDES[DEFAULT_STATUS] (constants.py) замість PAYMENT_STATUS_
    CONVERSIONS: apply_payment_values ВЖЕ замінив "РВЗ" на 10, бо файл
    information_unit ОДНОЗНАЧНО казав "БЗВП"/"Адап" - цей звіт звіряється з
    ТИМ САМИМ файлом, тож "10" і "БЗВП"/"Адап" - ТА САМА реальна ситуація, а
    не розбіжність. "НОВ" тут НЕ бере участі - на відміну від "БЗВП"/"Адап",
    "НОВ" замінюється на буквальний текст "НОВ" (не на 10), тож звичайного
    текстового збігу вже досить, без цієї спеціальної еквівалентності."""
    roster.ws.cell(row=2, column=5).value = 10  # 01.08 - вже замінене БЗВП/Адап -> 10
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", information_unit_status),
        ],
    }

    assert pmc.find_mismatches(roster, snapshot) == []


def test_find_mismatches_treats_vd_bzvp_and_its_converted_number_as_equivalent(roster):
    """Реальний випадок: apply_payment_values
    ВЖЕ замінив "ВД" на 10, бо файл information_unit ОДНОЗНАЧНО казав
    "БЗВП" (PAYMENT_STATUS_OVERRIDES["ВД"], constants.py) - цей звіт
    звіряється з ТИМ САМИМ файлом, тож "10" і "БЗВП" - ТА САМА реальна
    ситуація, а не розбіжність (на відміну від "ВД" vs "БЗВП" САМИХ ПО
    СОБІ, коли override НЕ спрацював - це й далі справжня розбіжність)."""
    roster.ws.cell(row=2, column=5).value = 10  # 01.08 - вже замінене ВД+БЗВП -> 10
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", "БЗВП"),
        ],
    }

    assert pmc.find_mismatches(roster, snapshot) == []


def test_find_mismatches_dedupes_files_with_identical_content(roster):
    """Реальний випадок: ДВА файли з ледь різними назвами (напр. одна зайва
    пробіл) подають АБСОЛЮТНО ІДЕНТИЧНІ дані для ОДНІЄЇ людини - типова
    причина: файл випадково задубльовано на диску. Лишає лише ОДИН запис
    (файл з алфавітно першою назвою), а не два однакових рядки в звіті."""
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл_б.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
            ("файл_а.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["file_name"] == "файл_а.xlsx"


def test_find_mismatches_keeps_files_with_differing_content_separate(roster):
    """Два файли з РІЗНИМИ значеннями (навіть якщо для ІНШИХ людей вони
    можуть частково збігатись) - НЕ дублікати, лишаються окремими записами -
    дедуплікація не повинна ховати справжні розбіжності між підрозділами."""
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
            ("файл2.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 30),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 2


def test_find_mismatches_does_not_duplicate_people_that_are_in_the_roster(roster):
    """Людина, яка Є в ростері - НЕ повинна ще й потрапити в
    _missing_from_roster_records, навіть якщо вона теж є в snapshot."""
    snapshot = {
        (normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [
            ("файл1.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
        ],
    }

    records = pmc.find_mismatches(roster, snapshot)

    assert len(records) == 1
    assert records[0]["dates"] == {date(2026, 8, 1): ("РВЗ", 100)}  # не _NOT_IN_ROSTER_LABEL


# -------------------------
# read_information_unit_snapshot
# -------------------------

def test_read_information_unit_snapshot_includes_file_attribution_from_tabel(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "ТАБЕЛЬ"
    for col_idx, text in enumerate(["Посада", "Звання", "ПІБ", "ДАТА"], start=1):
        ws.cell(row=1, column=col_idx, value=text)
    ws.cell(row=2, column=4, value=datetime(2026, 8, 1))
    ws.cell(row=3, column=1, value="Посада1")
    ws.cell(row=3, column=2, value="сержант")
    ws.cell(row=3, column=3, value="ПЕРШИЙ Перший Перший")
    ws.cell(row=3, column=4, value=100)
    wb.save(str(tmp_path / "рота.xlsx"))

    snapshot, skipped = pmc.read_information_unit_snapshot(str(tmp_path))

    assert skipped == []
    assert snapshot[(normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1))] == [
        ("рота.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
    ]


def test_read_information_unit_snapshot_reads_tabel_pidrozdil_column(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "ТАБЕЛЬ"
    for col_idx, text in enumerate(["ПІДРОЗДІЛ", "Посада", "Звання", "ПІБ", "ДАТА"], start=1):
        ws.cell(row=1, column=col_idx, value=text)
    ws.cell(row=2, column=5, value=datetime(2026, 8, 1))
    ws.cell(row=3, column=1, value="підрозділ 1")
    ws.cell(row=3, column=2, value="Посада1")
    ws.cell(row=3, column=3, value="сержант")
    ws.cell(row=3, column=4, value="ПЕРШИЙ Перший Перший")
    ws.cell(row=3, column=5, value=100)
    wb.save(str(tmp_path / "рота.xlsx"))

    snapshot, skipped = pmc.read_information_unit_snapshot(str(tmp_path))

    assert skipped == []
    assert snapshot[(normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1))] == [
        ("рота.xlsx", "ПЕРШИЙ Перший Перший", "підрозділ 1", "Посада1", "сержант", 100),
    ]


def test_read_information_unit_snapshot_reads_tabel_with_dates_in_row1(tmp_path):
    """Реальний випадок: аркуш ТАБЕЛЬ з датами ОДРАЗУ в рядку 1 (дані - з
    рядка 2, а не з рядка 3) - той самий структурний варіант, що й
    information_unit_reader (той самий _tabel_date_row_and_columns)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "ТАБЕЛЬ"
    for col_idx, text in enumerate(["Посада", "Звання", "ПІБ"], start=1):
        ws.cell(row=1, column=col_idx, value=text)
    ws.cell(row=1, column=4, value=datetime(2026, 8, 1))
    ws.cell(row=2, column=1, value="Командир роти")
    ws.cell(row=2, column=2, value="старший лейтенант")
    ws.cell(row=2, column=3, value="ПЕРШИЙ Перший Перший")
    ws.cell(row=2, column=4, value=100)
    wb.save(str(tmp_path / "підрозділ 1.xlsx"))

    snapshot, skipped = pmc.read_information_unit_snapshot(str(tmp_path))

    assert skipped == []
    assert snapshot[(normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1))] == [
        ("підрозділ 1.xlsx", "ПЕРШИЙ Перший Перший", None, "Командир роти", "старший лейтенант", 100),
    ]


def test_read_information_unit_snapshot_reads_xls_tabel_sheet(tmp_path):
    """Реальний випадок: файл цієї теки у старому бінарному форматі (.xls) -
    раніше мовчки потрапляв у skipped, тепер читається так само, як .xlsx
    (content/information_unit_reader._load_workbook/_XlsWorkbookAdapter,
    той самий шлях, що й read_payment_values)."""
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("ТАБЕЛЬ")
    for col_idx, text in enumerate(["Посада", "Звання", "ПІБ"]):
        ws.write(0, col_idx, text)
    ws.write(0, 3, datetime(2026, 8, 1), _XLS_DATE_STYLE)
    ws.write(1, 0, "Посада1")
    ws.write(1, 1, "сержант")
    ws.write(1, 2, "ПЕРШИЙ Перший Перший")
    ws.write(1, 3, 100)
    wb.save(str(tmp_path / "підрозділ 1.xls"))

    snapshot, skipped = pmc.read_information_unit_snapshot(str(tmp_path))

    assert skipped == []
    assert snapshot[(normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1))] == [
        ("підрозділ 1.xls", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
    ]


def test_read_information_unit_snapshot_includes_file_attribution_from_bchs(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "БЧС"
    ws.cell(row=1, column=1, value="ВЗВОД ЗВ'ЯЗКУ 08.08.2026")
    header = ["Посада", "Звання", "ПРІЗВИЩЕ ім'я по батькові", "БЧС", "Виплата додактової винагороди"]
    for col_idx, text in enumerate(header, start=1):
        ws.cell(row=2, column=col_idx, value=text)
    ws.cell(row=3, column=1, value="Посада1")
    ws.cell(row=3, column=2, value="матрос")
    ws.cell(row=3, column=3, value="ПЕРШИЙ Перший Перший")
    ws.cell(row=3, column=4, value="РВЗ")
    ws.cell(row=3, column=5, value=100)
    wb.save(str(tmp_path / "БЧС ВЗ 08.08.2026.xlsx"))

    snapshot, skipped = pmc.read_information_unit_snapshot(str(tmp_path))

    assert skipped == []
    assert snapshot[(normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 8))] == [
        ("БЧС ВЗ 08.08.2026.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "матрос", 100),
    ]


def test_read_information_unit_snapshot_reads_bchs_pidrozdil_column(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "БЧС"
    ws.cell(row=1, column=1, value="ВЗВОД ЗВ'ЯЗКУ 08.08.2026")
    header = ["ПІДРОЗДІЛ", "Посада", "Звання", "ПРІЗВИЩЕ ім'я по батькові", "БЧС", "Виплата додактової винагороди"]
    for col_idx, text in enumerate(header, start=1):
        ws.cell(row=2, column=col_idx, value=text)
    ws.cell(row=3, column=1, value="ВЗ")
    ws.cell(row=3, column=2, value="Посада1")
    ws.cell(row=3, column=3, value="матрос")
    ws.cell(row=3, column=4, value="ПЕРШИЙ Перший Перший")
    ws.cell(row=3, column=5, value="РВЗ")
    ws.cell(row=3, column=6, value=100)
    wb.save(str(tmp_path / "БЧС ВЗ 08.08.2026.xlsx"))

    snapshot, skipped = pmc.read_information_unit_snapshot(str(tmp_path))

    assert skipped == []
    assert snapshot[(normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 8))] == [
        ("БЧС ВЗ 08.08.2026.xlsx", "ПЕРШИЙ Перший Перший", "ВЗ", "Посада1", "матрос", 100),
    ]


def test_read_information_unit_snapshot_skips_unsupported_extension(tmp_path):
    (tmp_path / "стара.doc").write_text("не таблиця взагалі")

    snapshot, skipped = pmc.read_information_unit_snapshot(str(tmp_path))

    assert snapshot == {}
    assert len(skipped) == 1
    assert "не підтримується" in skipped[0]["reason"]


def test_read_information_unit_snapshot_skips_office_lock_file(tmp_path):
    """Той самий випадок, що й read_payment_values - lock-файл Excel
    ("~$Назва.xlsx") тихо пропускається, а не спричиняє крах читання."""
    (tmp_path / "~$рота.xlsx").write_text("lock-файл, не справжня книга")

    snapshot, skipped = pmc.read_information_unit_snapshot(str(tmp_path))

    assert snapshot == {}
    assert skipped == []


def test_read_information_unit_snapshot_skips_a_corrupted_file_and_still_reads_the_rest(tmp_path):
    """Той самий випадок, що й read_payment_values - ОДИН пошкоджений файл
    (не дійсний zip-архів попри розширення .xlsx) не повинен зупиняти
    прогін для решти, справних файлів; записується в skipped із причиною."""
    (tmp_path / "пошкоджена.xlsx").write_bytes(b"garbage, not a real workbook")
    wb = Workbook()
    ws = wb.active
    ws.title = "Табель"
    for col_idx, text in enumerate(["Посада", "Звання", "ПІБ", "ДАТА"], start=1):
        ws.cell(row=1, column=col_idx, value=text)
    ws.cell(row=2, column=4, value=date(2026, 8, 1))
    ws.cell(row=3, column=1, value="Посада1")
    ws.cell(row=3, column=2, value="сержант")
    ws.cell(row=3, column=3, value="ПЕРШИЙ Перший Перший")
    ws.cell(row=3, column=4, value=100)
    wb.save(str(tmp_path / "рота.xlsx"))

    snapshot, skipped = pmc.read_information_unit_snapshot(str(tmp_path))

    assert snapshot[(normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1))] == [
        ("рота.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
    ]
    assert len(skipped) == 1
    assert "пошкоджена.xlsx" in skipped[0]["reason"]


def test_read_information_unit_snapshot_skips_file_with_neither_source_readable(tmp_path):
    wb = Workbook()
    wb.active.title = "Щось інше"
    wb.save(str(tmp_path / "порожній.xlsx"))

    snapshot, skipped = pmc.read_information_unit_snapshot(str(tmp_path))

    assert snapshot == {}
    assert len(skipped) == 1
    assert "не знайдено аркуш ТАБЕЛЬ" in skipped[0]["reason"]


def test_read_information_unit_snapshot_records_non_default_bchs_status(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "БЧС"
    ws.cell(row=1, column=1, value="ВЗВОД ЗВ'ЯЗКУ 08.08.2026")
    header = ["Посада", "Звання", "ПРІЗВИЩЕ ім'я по батькові", "БЧС", "Виплата додактової винагороди"]
    for col_idx, text in enumerate(header, start=1):
        ws.cell(row=2, column=col_idx, value=text)
    ws.cell(row=3, column=1, value="Посада1")
    ws.cell(row=3, column=2, value="матрос")
    ws.cell(row=3, column=3, value="ДРУГИЙ Другий Другий")
    ws.cell(row=3, column=4, value="ВД")
    ws.cell(row=3, column=5, value=100)
    wb.save(str(tmp_path / "БЧС ВЗ 08.08.2026.xlsx"))

    snapshot, skipped = pmc.read_information_unit_snapshot(str(tmp_path))

    assert skipped == []
    assert snapshot[(normalize_name("ДРУГИЙ Другий Другий"), date(2026, 8, 8))] == [
        ("БЧС ВЗ 08.08.2026.xlsx", "ДРУГИЙ Другий Другий", None, "Посада1", "матрос", "ВД"),
    ]


def test_normalize_for_compare_treats_none_as_empty_string():
    assert pmc._normalize_for_compare(None) == ""


def test_tabel_label_columns_ignores_non_string_header_cells():
    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value=123)  # нетекстова клітинка - не має падати
    ws.cell(row=1, column=2, value="ПІДРОЗДІЛ")
    ws.cell(row=1, column=3, value="ПОСАДА")
    ws.cell(row=1, column=4, value="ЗВАННЯ")

    pidrozdil_col, posada_col, zvannya_col = pmc._tabel_label_columns(ws)

    assert pidrozdil_col == 2
    assert posada_col == 3
    assert zvannya_col == 4


def test_read_tabel_snapshot_reports_when_pib_column_missing():
    wb = Workbook()
    ws = wb.active
    ws.title = "ТАБЕЛЬ"
    ws.cell(row=1, column=1, value="Щось")
    ws.cell(row=2, column=2, value=datetime(2026, 8, 1))

    reason = pmc._read_tabel_snapshot(wb, "рота.xlsx", "рота.xlsx", {})

    assert reason is not None
    assert "колонку ПІБ" in reason


def test_read_tabel_snapshot_skips_rows_with_empty_pib():
    wb = Workbook()
    ws = wb.active
    ws.title = "ТАБЕЛЬ"
    for col_idx, text in enumerate(["Посада", "Звання", "ПІБ", "ДАТА"], start=1):
        ws.cell(row=1, column=col_idx, value=text)
    ws.cell(row=2, column=4, value=datetime(2026, 8, 1))
    # рядок 3 - порожній ПІБ, має бути пропущений без падіння.
    ws.cell(row=4, column=1, value="Посада1")
    ws.cell(row=4, column=2, value="сержант")
    ws.cell(row=4, column=3, value="ПЕРШИЙ Перший Перший")
    ws.cell(row=4, column=4, value=100)

    snapshot = {}
    reason = pmc._read_tabel_snapshot(wb, "рота.xlsx", "рота.xlsx", snapshot)

    assert reason is None
    assert snapshot[(normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1))] == [
        ("рота.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "сержант", 100),
    ]


def test_read_bchs_snapshot_reports_when_no_date_found():
    wb = Workbook()
    ws = wb.active
    ws.title = "БЧС"
    ws.cell(row=1, column=1, value="ЯКИЙСЬ ВЗВОД")

    reason = pmc._read_bchs_snapshot(wb, "без_дати.xlsx", "без_дати.xlsx", {})

    assert reason is not None
    assert "не вдалось визначити дату" in reason


def test_read_bchs_snapshot_reports_when_pib_column_missing():
    wb = Workbook()
    ws = wb.active
    ws.title = "БЧС"
    ws.cell(row=1, column=1, value="ВЗВОД 08.08.2026")
    ws.cell(row=2, column=1, value="Посада")

    reason = pmc._read_bchs_snapshot(wb, "рота.xlsx", "рота.xlsx", {})

    assert reason is not None
    assert "колонку ПІБ" in reason


def test_read_bchs_snapshot_skips_rows_with_empty_pib():
    wb = Workbook()
    ws = wb.active
    ws.title = "БЧС"
    ws.cell(row=1, column=1, value="ВЗВОД ЗВ'ЯЗКУ 08.08.2026")
    header = ["Посада", "Звання", "ПРІЗВИЩЕ ім'я по батькові", "БЧС", "Виплата додактової винагороди"]
    for col_idx, text in enumerate(header, start=1):
        ws.cell(row=2, column=col_idx, value=text)
    # рядок 3 - порожній ПІБ (напр. рядок-підзаголовок), має бути пропущений.
    ws.cell(row=4, column=1, value="Посада1")
    ws.cell(row=4, column=2, value="матрос")
    ws.cell(row=4, column=3, value="ПЕРШИЙ Перший Перший")
    ws.cell(row=4, column=4, value="РВЗ")
    ws.cell(row=4, column=5, value=100)

    snapshot = {}
    reason = pmc._read_bchs_snapshot(wb, "рота.xlsx", "рота.xlsx", snapshot)

    assert reason is None
    assert snapshot[(normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 8))] == [
        ("рота.xlsx", "ПЕРШИЙ Перший Перший", None, "Посада1", "матрос", 100),
    ]


# -------------------------
# write_mismatch_report
# -------------------------

def _record(**overrides):
    base = {
        "file_name": "файл1.xlsx",
        "roster_pib_raw": "ПЕРШИЙ Перший Перший", "file_pib_raw": "ПЕРШИЙ Перший Перший",
        "roster_pidrozdil": "підрозділ 1", "roster_posada": "Посада1", "roster_zvannya": "сержант",
        "file_pidrozdil": "підрозділ 1", "file_posada": "Посада1", "file_zvannya": "сержант",
        "dates": {date(2026, 8, 1): ("РВЗ", 100)},
    }
    base.update(overrides)
    return base


def test_write_mismatch_report_merges_pib_and_matching_labels(tmp_path):
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record()], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    assert ws.cell(row=1, column=1).value == "Назва файлу"
    assert ws.cell(row=1, column=2).value == "Підрозділ"
    assert ws.cell(row=1, column=5).value == "ПІБ"
    assert ws.cell(row=1, column=6).value == datetime(2026, 8, 1)  # дата - ОКРЕМА колонка, як в ОБЛІК.xlsx
    assert ws.cell(row=2, column=1).value == "ОБЛІК для виплат (за рапортами)"
    assert ws.cell(row=3, column=1).value == "файл1.xlsx"
    assert ws.cell(row=2, column=5).value == "ПЕРШИЙ Перший Перший"
    merged = {str(rng) for rng in ws.merged_cells.ranges}
    assert "B2:B3" in merged  # Підрозділ збігається - об'єднана
    assert "C2:C3" in merged  # Посада збігається - об'єднана
    assert "D2:D3" in merged  # Звання збігається - об'єднана
    assert "E2:E3" in merged  # ПІБ збігається - об'єднана
    assert ws.cell(row=2, column=6).value == "РВЗ"
    assert ws.cell(row=3, column=6).value == 100


def test_write_mismatch_report_uses_a_separate_column_per_date(tmp_path):
    """За прямою вказівкою користувача - структура колонок як в ОБЛІК.xlsx:
    кожна дата - ОКРЕМА колонка (а не спільна колонка "Дата"). Людина+файл з
    КІЛЬКОМА розбіжними датами - ОДНА пара рядків із КІЛЬКОМА заповненими
    колонками-датами, а не окрема пара рядків на кожну дату."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report(
        [_record(dates={date(2026, 8, 1): ("РВЗ", 100), date(2026, 8, 2): ("РВЗ", 30)})],
        str(output_path),
    )

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    assert ws.max_row == 3  # заголовок + ОДНА пара рядків
    assert ws.cell(row=1, column=6).value == datetime(2026, 8, 1)
    assert ws.cell(row=1, column=7).value == datetime(2026, 8, 2)
    assert ws.cell(row=2, column=6).value == "РВЗ"
    assert ws.cell(row=3, column=6).value == 100
    assert ws.cell(row=2, column=7).value == "РВЗ"
    assert ws.cell(row=3, column=7).value == 30


def test_write_mismatch_report_appends_manual_entry_columns_after_the_dates(tmp_path):
    """За прямою вказівкою користувача - "ПІДСТАВИ" і "ДАТА В СТАТУС
    СПЕЦКОНТИНГЕНТУ" - дві останні колонки, ПІСЛЯ всіх дат (немає джерела
    даних для них у проєкті - завжди лишаються порожніми, для РУЧНОГО
    заповнення)."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report(
        [_record(dates={date(2026, 8, 1): ("РВЗ", 100), date(2026, 8, 2): ("РВЗ", 30)})],
        str(output_path),
    )

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    assert ws.max_column == 9  # 5 колонок-міток + 2 дати + 2 колонки ручного заповнення
    assert ws.cell(row=1, column=8).value == "ПІДСТАВИ"
    assert ws.cell(row=1, column=9).value == "ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ"
    assert ws.cell(row=2, column=8).value is None
    assert ws.cell(row=2, column=9).value is None
    assert ws.cell(row=2, column=8).fill.fill_type is None
    merged = {str(rng) for rng in ws.merged_cells.ranges}
    assert "H2:H3" in merged
    assert "I2:I3" in merged


def test_write_mismatch_report_leaves_other_records_date_columns_blank(tmp_path):
    """Дата, розбіжна лише для ОДНОГО з кількох записів, - заповнена лише в
    ЙОГО рядку; в рядку іншого запису, де ЦІЄЇ розбіжності немає, колонка
    лишається порожньою (але з межею - див. test_write_mismatch_report_adds_
    borders_to_header_and_data_cells)."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report(
        [
            _record(dates={date(2026, 8, 1): ("РВЗ", 100)}),
            _record(
                file_name="файл2.xlsx", roster_pib_raw="ДРУГИЙ Другий Другий", file_pib_raw="ДРУГИЙ Другий Другий",
                dates={date(2026, 8, 2): ("РВЗ", 30)},
            ),
        ],
        str(output_path),
    )

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    assert ws.cell(row=1, column=6).value == datetime(2026, 8, 1)
    assert ws.cell(row=1, column=7).value == datetime(2026, 8, 2)
    assert ws.cell(row=2, column=6).value == "РВЗ"  # запис 1 - 01.08 заповнена
    assert ws.cell(row=2, column=7).value is None  # запис 1 - 02.08 його не стосується
    assert ws.cell(row=4, column=6).value is None  # запис 2 - 01.08 його не стосується
    assert ws.cell(row=4, column=7).value == "РВЗ"  # запис 2 - 02.08 заповнена


def test_write_mismatch_report_uses_placeholder_for_missing_label_values(tmp_path):
    """Порожнє (None) значення в ПОРІВНЮВАНІЙ колонці (напр. роcтер без
    колонки ЗВАННЯ) - показане як _EMPTY_PLACEHOLDER, а не як буквально
    порожня клітинка."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record(roster_zvannya=None, file_zvannya=None)], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    assert ws.cell(row=2, column=4).value == pmc._EMPTY_PLACEHOLDER


def test_write_mismatch_report_uses_placeholder_for_missing_pidrozdil(tmp_path):
    """Порожнє (None) значення з ОБОХ сторін (роcтер і файл ОБИДВА не подають
    Підрозділ) - однакове (збігається), тож об'єднана клітинка, показана як
    _EMPTY_PLACEHOLDER, а не буквально порожня."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record(roster_pidrozdil=None, file_pidrozdil=None)], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    assert ws.cell(row=2, column=2).value == pmc._EMPTY_PLACEHOLDER
    merged = {str(rng) for rng in ws.merged_cells.ranges}
    assert "B2:B3" in merged


def test_write_mismatch_report_splits_differing_pib(tmp_path):
    """Реальний випадок: ПІБ у файлі information_unit НЕ збігається з
    роcтеровим написанням (напр. друкарська помилка) - клітинка має
    розділитись на 2 значення, а не мовчки показати лише одне з них."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report(
        [_record(roster_pib_raw="ПЕРШИЙ Перший Перший", file_pib_raw="ПЕРШИЙ Перший Перека")],
        str(output_path),
    )

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    assert ws.cell(row=2, column=5).value == "ПЕРШИЙ Перший Перший"
    assert ws.cell(row=3, column=5).value == "ПЕРШИЙ Перший Перека"
    merged = {str(rng) for rng in ws.merged_cells.ranges}
    assert "E2:E3" not in merged


def test_write_mismatch_report_merges_pib_despite_cosmetic_whitespace_difference(tmp_path):
    """Зайвий пробіл (типова "бруднота" реальних даних) - НЕ справжня
    розбіжність, ПІБ все одно має об'єднатись."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report(
        [_record(roster_pib_raw="ПЕРШИЙ  Перший Перший", file_pib_raw="ПЕРШИЙ Перший Перший ")],
        str(output_path),
    )

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    merged = {str(rng) for rng in ws.merged_cells.ranges}
    assert "E2:E3" in merged


def test_write_mismatch_report_merges_equivalent_ppd_and_converted_number(tmp_path):
    """10 (вже застосована конвертація ППД -> 10, apply_payment_values) і
    текстовий статус "ППД" з файлу information_unit - ЕКВІВАЛЕНТНІ
    (PAYMENT_STATUS_CONVERSIONS з constants.py) - об'єднана клітинка, зелена
    заливка, а не розбіжність, попри те, що самі значення текстово різні."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record(dates={date(2026, 8, 1): (10, "ППД")})], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    merged = {str(rng) for rng in ws.merged_cells.ranges}
    assert "F2:F3" in merged
    assert ws.cell(row=2, column=6).fill.fgColor.rgb == "FFC6EFCE"


def test_write_mismatch_report_merges_prvd_and_perevd_spelling_variant(tmp_path):
    """Реальний випадок: "ПРВД" (переведений в інший підрозділ) з
    ОБЛІК і "Перевд." з файлу information_unit - ТОЙ САМИЙ факт, лише інше
    написання (PAYMENT_STATUS_OVERRIDES["ПРВД"] з constants.py - запис-
    ідентичність заради _statuses_equivalent) - об'єднана клітинка, зелена
    заливка, а не розбіжність, навіть коли рядок Є у звіті через ІНШУ,
    реальну розбіжність на іншу дату."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record(dates={date(2026, 8, 1): ("ПРВД", "Перевд.")})], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    merged = {str(rng) for rng in ws.merged_cells.ranges}
    assert "F2:F3" in merged
    assert ws.cell(row=2, column=6).value == "ПРВД"
    assert ws.cell(row=2, column=6).fill.fgColor.rgb == "FFC6EFCE"


def test_write_mismatch_report_merges_matching_non_comparable_status(tmp_path):
    """Реальний випадок (АНДРЕЇШИН/ТРОФІМОВ): дата, де ОБЛІК і файл
    погоджуються на непорівнюваному статусі (напр. "ВПСЗ") - має показуватись
    ЗЕЛЕНОЮ об'єднаною клітинкою, а не порожньою."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record(dates={date(2026, 8, 1): ("ВПСЗ", "ВПСЗ")})], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    merged = {str(rng) for rng in ws.merged_cells.ranges}
    assert "F2:F3" in merged
    assert ws.cell(row=2, column=6).value == "ВПСЗ"
    assert ws.cell(row=2, column=6).fill.fgColor.rgb == "FFC6EFCE"


def test_write_mismatch_report_splits_differing_pidrozdil(tmp_path):
    """За прямою вказівкою користувача - Підрозділ ТЕЖ звіряється, як і
    Посада/Звання/ПІБ: різні значення роcтера й файлу - 2 окремих значення
    (по одному на рядок), а не мовчки лише роcтерове."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record(roster_pidrozdil="підрозділ 1", file_pidrozdil="підрозділ 2")], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    assert ws.cell(row=2, column=2).value == "підрозділ 1"
    assert ws.cell(row=3, column=2).value == "підрозділ 2"
    merged = {str(rng) for rng in ws.merged_cells.ranges}
    assert "B2:B3" not in merged  # Підрозділ відрізняється - НЕ об'єднана
    assert "C2:C3" in merged  # Посада все ще збігається - об'єднана


def test_write_mismatch_report_splits_differing_posada_but_merges_matching_zvannya(tmp_path):
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record(roster_posada="Посада1", file_posada="Посада2")], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    assert ws.cell(row=2, column=3).value == "Посада1"
    assert ws.cell(row=3, column=3).value == "Посада2"
    merged = {str(rng) for rng in ws.merged_cells.ranges}
    assert "C2:C3" not in merged  # Посада відрізняється - НЕ об'єднана
    assert "D2:D3" in merged  # Звання все ще збігається - об'єднана


def test_write_mismatch_report_multiple_records_use_separate_row_pairs(tmp_path):
    records = [
        _record(),
        _record(
            file_name="файл2.xlsx", roster_pib_raw="ДРУГИЙ Другий Другий", file_pib_raw="ДРУГИЙ Другий Другий",
            roster_pidrozdil="підрозділ 2", file_pidrozdil="підрозділ 2",
            roster_posada="Посада2", file_posada="Посада2",
            roster_zvannya="матрос", file_zvannya="матрос",
            dates={date(2026, 8, 2): ("РВЗ", 30)},
        ),
    ]
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report(records, str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    assert ws.cell(row=4, column=1).value == "ОБЛІК для виплат (за рапортами)"
    assert ws.cell(row=5, column=1).value == "файл2.xlsx"
    assert ws.cell(row=4, column=5).value == "ДРУГИЙ Другий Другий"


def test_write_mismatch_report_colors_matching_cell_green_and_mismatching_red(tmp_path):
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record(roster_posada="Посада1", file_posada="Посада2")], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    # Посада (col 3) відрізняється - червоний.
    assert ws.cell(row=2, column=3).fill.fgColor.rgb == "FFFFC7CE"
    assert ws.cell(row=3, column=3).fill.fgColor.rgb == "FFFFC7CE"
    # Дата (col 6) - розбіжна за побудовою (інакше не потрапила б у звіт) - червона.
    assert ws.cell(row=2, column=6).fill.fgColor.rgb == "FFFFC7CE"
    assert ws.cell(row=3, column=6).fill.fgColor.rgb == "FFFFC7CE"
    # "Назва файлу" - не порівнюване значення, без заливки.
    assert ws.cell(row=2, column=1).fill.fill_type is None
    # Підрозділ (col 2) збігається за замовчуванням у _record() - зелений.
    assert ws.cell(row=2, column=2).fill.fgColor.rgb == "FFC6EFCE"


def test_write_mismatch_report_body_uses_times_new_roman_14pt(tmp_path):
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record()], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    body_font = ws.cell(row=2, column=1).font
    assert body_font.name == "Times New Roman"
    assert body_font.size == 14


def test_write_mismatch_report_header_looks_like_oblik_file(tmp_path):
    """За прямою вказівкою користувача - заголовок error_mis_statuses.xlsx
    має мати той самий вигляд, що й заголовок ОБЛІК.xlsx (resources/
    ОБЛІК.xlsx: Bahnschrift Light SemiCondensed 14 жирний білий текст на
    темно-бірюзовій заливці FF035C6B), а не довільний стиль."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record()], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    header_font = ws.cell(row=1, column=1).font
    assert header_font.name == "Bahnschrift Light SemiCondensed"
    assert header_font.size == 14
    assert header_font.bold is True
    assert header_font.color.rgb == "FFFFFFFF"
    assert ws.cell(row=1, column=1).fill.fgColor.rgb == "FF035C6B"


def test_write_mismatch_report_wraps_permission_error_with_a_friendly_message(tmp_path, monkeypatch):
    """Реальний випадок: error_mis_statuses.xlsx уже відкритий у Excel -
    wb.save() падає з PermissionError - обгортаємо зрозумілим повідомленням
    українською, той самий підхід, що й Timesheet.save (content/oblik_timesheet.py)."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    from openpyxl import Workbook as _Workbook

    def _raise(self, *args, **kwargs):
        raise PermissionError("[WinError 32] заблоковано іншою програмою")

    monkeypatch.setattr(_Workbook, "save", _raise)

    with pytest.raises(PermissionError, match="Закрийте його"):
        pmc.write_mismatch_report([_record()], str(output_path))


def test_write_mismatch_report_with_no_records_does_not_create_a_file(tmp_path):
    """За прямою вказівкою користувача - немає сенсу створювати файл лише із
    заголовком, коли розповідати нема про що."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    result_path = pmc.write_mismatch_report([], str(output_path))

    assert result_path is None
    assert not output_path.exists()


def test_write_mismatch_report_adds_borders_to_header_and_data_cells(tmp_path):
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record()], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    for row in (1, 2, 3):
        for col_idx in range(1, 9):  # 5 колонок-міток + 1 дата + 2 колонки для ручного заповнення
            cell = ws.cell(row=row, column=col_idx)
            assert cell.border.left.style == "thin"
            assert cell.border.right.style == "thin"
            assert cell.border.bottom.style == "thin"
            # top - лише для рядка 1 (заголовок) і рядка 2 (row_a, "якір"
            # об'єднаного діапазону) - openpyxl очищує top другого рядка
            # об'єднання (внутрішній шов) при збереженні, це очікувано і не
            # впливає на візуальний вигляд (Excel не малює внутрішні лінії
            # всередині об'єднаної клітинки).
            if row != 3:
                assert cell.border.top.style == "thin"


def test_write_mismatch_report_does_not_wrap_body_text(tmp_path):
    """За прямою вказівкою користувача - текст ДАНИХ (не заголовка) НЕ
    переноситься в кілька рядків (раніше через wrap_text=True + мала висота
    рядка текст "накладався" на сусідній рядок і ставав нечитабельним).
    Заголовок - wrap_text=True (той самий вигляд, що й в ОБЛІК.xlsx, короткі
    підписи колонок - переносу там не станеться)."""
    output_path = tmp_path / "error_mis_statuses.xlsx"

    pmc.write_mismatch_report([_record()], str(output_path))

    wb = load_workbook(str(output_path))
    ws = wb["Розбіжності"]
    # wrap_text=False - той самий default, що й Alignment() - openpyxl не
    # серіалізує його явно, тож при читанні назад це None, а не False.
    assert not ws.cell(row=2, column=1).alignment.wrap_text
