import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import pytest
from docx import Document
from openpyxl import Workbook

import constants
from constants import SHORT_UNIT_BATTALION, SHORT_UNIT_BRIGADE, HIGHER_COMMANDER_TITLE
from content import report_changes
import content.money_report_helpers as mrh

_BR = f"БР {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE}"


def _write_xlsx(path, headers, rows, sheet_name="Аркуш1"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)
    return path


def _write_docx(path, paragraphs=()):
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(path)
    return str(path)


def _row(pib, posada, pidrozdil, days, zvannya="сержант"):
    row = {"ПІБ": pib, "ПОСАДА": posada, "ПІДРОЗДІЛ": pidrozdil, "ЗВАННЯ": zvannya}
    row.update(days)
    return row


D1 = datetime(2026, 6, 1)
D2 = datetime(2026, 6, 2)

# Мінімальний MONEY_REPORT_CATEGORIES-подібний словник для тестів diff-логіки -
# без реальних grounds/use_brs (required_basis: False скрізь, щоб порожня
# ПІДСТАВА була ОЧІКУВАНОЮ і не заважала перевіряти саме ПЕРІОД/ДНІ).
_CATEGORIES = {
    30: {
        "general": [],
        "ЖИТТЄДІЯЛЬНІСТЬ": {
            "grounds": [], "use_brs": False, "use_brs_from_selected_folder": False, "exclude_general": [],
            "default": True, "include_to_report": True, "required_basis": False,
        },
    },
    100: {
        "general": [],
        "БД(СЗ)": {
            "grounds": [], "use_brs": False, "use_brs_from_selected_folder": False, "exclude_general": [],
            "default": True, "include_to_report": True, "required_basis": False,
        },
    },
}


# -------------------------
# _rows_from_check_file
# -------------------------
def test_rows_from_check_file_converts_date_keys_to_datetime(tmp_path):
    path = _write_xlsx(tmp_path / "ОБЛІК.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1, D2], [
        ["Підрозділ 1", "Стрілець", "сержант", "Перший Перший", 30, 100],
    ])

    rows_with_data, date_columns = report_changes._rows_from_check_file(str(path))

    assert date_columns == [D1, D2]
    assert all(isinstance(d, datetime) for d in date_columns)
    assert rows_with_data == [
        {"ПІДРОЗДІЛ": "Підрозділ 1", "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Перший Перший", D1: 30, D2: 100},
    ]


def test_rows_from_check_file_empty_file_returns_empty(tmp_path):
    path = _write_xlsx(tmp_path / "ОБЛІК.xlsx", ["ПІБ", D1], [])

    rows_with_data, date_columns = report_changes._rows_from_check_file(str(path))

    assert rows_with_data == []
    assert date_columns == [D1]


# -------------------------
# _backfill_missing_labels
# -------------------------
def test_backfill_missing_labels_fills_from_other_file_by_pib():
    """prev-файл без ПОСАДА (лише ЗВАННЯ+ПІБ) - підставляється з actual-файлу за
    тим самим (нормалізованим) ПІБ."""
    rows = [{"ЗВАННЯ": "сержант", "ПІБ": "Шістнадцятий Шістнадцятий"}]
    other_rows = [{"ПІДРОЗДІЛ": "Підрозділ 1", "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "ШІСТНАДЦЯТИЙ ШІСТНАДЦЯТИЙ"}]

    result = report_changes._backfill_missing_labels(rows, other_rows)

    assert result == [{"ЗВАННЯ": "сержант", "ПІБ": "Шістнадцятий Шістнадцятий", "ПІДРОЗДІЛ": "Підрозділ 1", "ПОСАДА": "Стрілець"}]


def test_backfill_missing_labels_leaves_row_unchanged_when_person_missing_from_other_file():
    rows = [{"ЗВАННЯ": "сержант", "ПІБ": "ПЕРШИЙ Перший"}]

    result = report_changes._backfill_missing_labels(rows, [])

    assert result == rows


def test_backfill_missing_labels_does_not_override_already_present_value():
    rows = [{"ПОСАДА": "Стрілець", "ПІБ": "Перший Перший"}]
    other_rows = [{"ПОСАДА": "Зовсім інша посада", "ПІБ": "Перший Перший"}]

    result = report_changes._backfill_missing_labels(rows, other_rows)

    assert result[0]["ПОСАДА"] == "Стрілець"


# -------------------------
# _resolve_file_year_month
# -------------------------
def test_resolve_file_year_month_returns_none_for_empty_columns():
    assert report_changes._resolve_file_year_month([]) == (None, None)


def test_resolve_file_year_month_picks_most_common_year_month():
    dates = [datetime(2026, 6, 1), datetime(2026, 6, 2), datetime(2026, 7, 1)]
    assert report_changes._resolve_file_year_month(dates) == (2026, 6)


# -------------------------
# _months_before_selected
# -------------------------
def test_months_before_selected_positive_for_earlier_months(monkeypatch):
    monkeypatch.setattr(report_changes, "MONTH", "07")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    assert report_changes._months_before_selected(2026, 6) == 1
    assert report_changes._months_before_selected(2026, 5) == 2


def test_months_before_selected_handles_year_boundary(monkeypatch):
    """Грудень ПОПЕРЕДНЬОГО року - "1 місяць тому" відносно січня наступного,
    а не "пізніше" через голе порівняння номерів місяців (12 > 01)."""
    monkeypatch.setattr(report_changes, "MONTH", "01")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    assert report_changes._months_before_selected(2025, 12) == 1


def test_months_before_selected_zero_or_negative_for_current_or_future(monkeypatch):
    monkeypatch.setattr(report_changes, "MONTH", "07")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    assert report_changes._months_before_selected(2026, 7) == 0
    assert report_changes._months_before_selected(2026, 8) == -1


# -------------------------
# _changes_category_config - use_brs/use_brs_from_selected_folder завжди
# вимкнені в цьому розділі (навіть для пункту 100) - підтверджено користувачем:
# "чомусь підтягує дані з NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY... хоча не мало би,
# я вказував що тільки MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR/_BN" - обрана
# папка місяця синхронізує ці константи (sync.changes_folder), і БЕЗ цього
# force-вимкнення "БД(СЗ)" пункту 100 (use_brs_from_selected_folder=True в
# реальному constants.py) мовчки підхопила б їх сюди, хоча підстава пункту 100
# в цьому розділі має йти ЛИШЕ з extra_grounds_by_person (find_person_document_references).
# -------------------------
def test_changes_category_config_disables_use_brs_for_non_100_point():
    categories = {30: {"general": [], "X": {
        "grounds": [], "use_brs": True, "use_brs_from_selected_folder": True, "exclude_general": [],
        "default": True, "include_to_report": True, "required_basis": False,
    }}}
    config = report_changes._changes_category_config(30, "X", categories)
    assert config["use_brs"] is False
    assert config["use_brs_from_selected_folder"] is False


def test_changes_category_config_disables_use_brs_for_point_100_too():
    categories = {100: {"general": [], "X": {
        "grounds": [], "use_brs": True, "use_brs_from_selected_folder": True, "exclude_general": [],
        "default": True, "include_to_report": True, "required_basis": False,
    }}}
    config = report_changes._changes_category_config(100, "X", categories)
    assert config["use_brs"] is False
    assert config["use_brs_from_selected_folder"] is False


def test_delta_rows_for_point_ignores_synced_brs_for_every_point(monkeypatch):
    """Навіть якщо категорія налаштована з use_brs=True (як у звичайному рапорті),
    у розділі змін вона НЕ повинна тягнути дані з НОВОСИНХРОНІЗОВАНОГО
    NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK (папка з підставами за попередній місяць) -
    підтверджено користувачем - для ЖОДНОГО пункту, включно зі 100.

    LOG_WAR/BN монкіпатчено в [] - точки 100/30 тут перевизначаються
    _changes_category_config'ом на РЕАЛЬНІ (з constants.py) LOG_WAR/BN
    (_CHANGES_POINT_PRIORITY_ORDER), а не на "general": [] з тестового
    categories - без цього тест перевіряв би не use_brs-ізоляцію (свою мету), а
    випадково збігся б чи ні з реальним виробничим вмістом LOG_WAR/BN на дату D1."""
    monkeypatch.setattr(report_changes, "MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR", [])
    monkeypatch.setattr(report_changes, "MONEY_REPORT_GENERAL_REFERENCES_BN", [])
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {"01.06.2026": {"бат": "117", "посилання_брг": "", "посилання_бат": ""}})
    categories = {
        30: {"general": [], "X": {
            "grounds": [], "use_brs": True, "use_brs_from_selected_folder": False, "exclude_general": [],
            "default": True, "include_to_report": True, "required_basis": False,
        }},
        100: {"general": [], "Y": {
            "grounds": [], "use_brs": True, "use_brs_from_selected_folder": False, "exclude_general": [],
            "default": True, "include_to_report": True, "required_basis": False,
        }},
    }
    prev_rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {D1: "ВД"})]
    actual_rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {D1: 30})]
    add_30 = report_changes._delta_rows_for_point(prev_rows, [D1], actual_rows, [D1], 30, categories)

    prev_rows_100 = [_row("Петров Петро", "Стрілець", "Підрозділ 1", {D1: "ВД"})]
    actual_rows_100 = [_row("Петров Петро", "Стрілець", "Підрозділ 1", {D1: 100})]
    add_100 = report_changes._delta_rows_for_point(prev_rows_100, [D1], actual_rows_100, [D1], 100, categories)

    assert add_30[0]["ПІДСТАВА"] == ""
    assert add_100[0]["ПІДСТАВА"] == ""


# -------------------------
# build_appendix_pairs_for_month - декілька пунктів разом (реклассифікація МІЖ пунктами)
# -------------------------
def test_build_appendix_pairs_for_month_reclassified_between_points(tmp_path):
    """Сценарій ГАРБУЗ (з новою логікою, підтверджено користувачем): у prev - усі
    30 днів під пунктом 30; у actual - 15 (ПІДМНОЖИНА prev) під 30 і 15 (нові) під
    100. Пункт 100 - ЛИШЕ "додати" (у prev не було жодного дня цієї людини під
    100, days=15). Пункт 30 - "Викласти в новій редакції" З РЕШТОЮ 15 днів (01-15) -
    підтверджено користувачем на прикладі АРТЬОМОВА: сам факт, що ЧАСТИНА старого
    30-денного запису "переїхала" у 100, означає, що ВЕСЬ запис Додатка 3 змінився й
    має бути переоформлений з АКТУАЛЬНИМ (меншим) периодом, а не лишатись
    незгаданим лише тому, що жодного НОВОГО 30-дня не з'явилось."""
    days_range = list(range(1, 31))
    all_june = {datetime(2026, 6, d): 30 for d in days_range}
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx",
        ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ"] + list(all_june.keys()),
        [["Підрозділ 1", "Командир взводу", "капітан", "ДРУГИЙ Другий"] + list(all_june.values())],
    )
    actual_values = {datetime(2026, 6, d): (30 if d <= 15 else 100) for d in days_range}
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx",
        ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ"] + list(actual_values.keys()),
        [["Підрозділ 1", "Командир взводу", "капітан", "ДРУГИЙ Другий"] + list(actual_values.values())],
    )

    pairs = report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [], _CATEGORIES)
    by_point = {point: (exclude_rows, add_rows) for point, _num, exclude_rows, add_rows in pairs}

    assert set(by_point) == {100, 30}
    point_100_exclude, point_100_add = by_point[100]
    assert point_100_exclude == []
    assert [r["ПІБ"] for r in point_100_add] == ["ДРУГИЙ Другий"]
    assert point_100_add[0]["ДНІ"] == 15

    point_30_exclude, point_30_add = by_point[30]
    assert point_30_exclude == []
    assert [r["ПІБ"] for r in point_30_add] == ["ДРУГИЙ Другий"]
    assert point_30_add[0]["ДНІ"] == 15
    assert point_30_add[0]["ПЕРІОД"] == "01.06.2026-15.06.2026"
    # В prev УЖЕ БУВ запис пункту 30 (весь місяць) - переоформлюється, тож
    # "Викласти в новій редакції", а НЕ звичайне "доповнити".
    assert point_30_add[0]["_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ"] is True


def test_build_appendix_pairs_for_month_restates_point_30_for_genuinely_new_days(tmp_path):
    """"Викласти в новій редакції" (Додаток 30) З'ЯВЛЯЄТЬСЯ, коли в actual дійсно
    є ДЕНЬ, що НЕ рахувався 30-м у prev (тут - був порожнім/відпусткою) - і, на
    відміну від звичайного "доповнити", показує ПОВНИЙ актуальний 30-денний
    період людини (усі дні actual, а не лише новий) - підтверджено користувачем:
    "Викласти в новій редакції" замінює ВЕСЬ запис Додатка, тож має описувати
    його ПОВНІСТЮ."""
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1, D2], [["Сімнадцятий Сімнадцятий", 30, "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1, D2], [["Сімнадцятий Сімнадцятий", 30, 30]])

    pairs = report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [], _CATEGORIES)
    by_point = {point: (exclude_rows, add_rows) for point, _num, exclude_rows, add_rows in pairs}

    assert set(by_point) == {30}
    exclude_rows, add_rows = by_point[30]
    assert exclude_rows == []
    assert [r["ПІБ"] for r in add_rows] == ["Сімнадцятий Сімнадцятий"]
    assert add_rows[0]["ДНІ"] == 2
    assert add_rows[0]["ПЕРІОД"] == "01.06.2026-02.06.2026"
    # D1 (30) вже був у prev - є що "перевидавати".
    assert add_rows[0]["_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ"] is True


def test_build_appendix_pairs_for_month_restate_point_new_participant_not_flagged_as_restate(tmp_path):
    """РЕГРЕСІЯ (підтверджено користувачем): людина, яка в prev НЕ мала жодного
    дня ні пункту 30, ні пункту 100 (щойно долучилась/повернулась) - тепер
    з'являється з 30-днями в actual - "_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ" МАЄ БУТИ
    False (немає що "перевидавати" - запису раніше не існувало взагалі), тож
    рендериться звичайним "Додаток 3 доповнити", а не "Викласти в новій
    редакції"."""
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1, D2], [["Новенький Іван", "ВД", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1, D2], [["Новенький Іван", 30, 30]])

    pairs = report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [], _CATEGORIES)
    by_point = {point: (exclude_rows, add_rows) for point, _num, exclude_rows, add_rows in pairs}

    assert set(by_point) == {30}
    exclude_rows, add_rows = by_point[30]
    assert [r["ПІБ"] for r in add_rows] == ["Новенький Іван"]
    assert add_rows[0]["ДНІ"] == 2
    assert add_rows[0]["_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ"] is False


def test_build_appendix_pairs_for_month_new_participant_with_both_100_and_30_uses_add_for_both(tmp_path):
    """Реальний сценарій, підтверджений користувачем: людина без жодного дня
    пункту 30 ЧИ 100 у prev, у актуального місяця має ОБИДВІ категорії -
    ОБИДВА Додатки (1 і 3) мають рендеритись звичайним "доповнити", жоден -
    "Викласти в новій редакції"."""
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1, D2], [["Новенький Іван", "ВД", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1, D2], [["Новенький Іван", 30, 100]])

    pairs = report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [], _CATEGORIES)
    by_point = {point: (exclude_rows, add_rows) for point, _num, exclude_rows, add_rows in pairs}

    assert set(by_point) == {100, 30}
    assert by_point[100][1][0]["_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ"] is False
    assert by_point[30][1][0]["_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ"] is False


def test_build_appendix_pairs_for_month_appendix_numbers_are_compressed(tmp_path):
    """Якщо змінився лише пункт 100 (пункт 30 - НЕЗМІННИЙ) - "Додаток 1" все одно
    відповідає пункту 100 (стиснуто - позиція серед пунктів, що ЗМІНИЛИСЬ, а НЕ
    фіксована прив'язка "30 - завжди Додаток 1")."""
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1],
        [["Перший Перший", "ВД"]],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1],
        [["Перший Перший", 100]],
    )

    pairs = report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [], _CATEGORIES)

    assert len(pairs) == 1
    point, appendix_number, exclude_rows, add_rows = pairs[0]
    assert point == 100
    assert appendix_number == 1
    assert exclude_rows == []
    assert [r["ПІБ"] for r in add_rows] == ["Перший Перший"]


def test_build_appendix_pairs_for_month_orders_appendices_100_then_50_then_30(tmp_path, monkeypatch):
    """Пріоритет нумерації додатків у ЦЬОМУ розділі - 100, 50, 30 (підтверджено
    користувачем), а НЕ пріоритет звичайних пунктів рапорту (30, 100, ...) і не
    порядок появи в categories. Пункт 50 - ЩЕ НЕ ІСНУЄ в реальному MONEY_REPORT_CATEGORIES
    (дані на нього поки не надані), але вже підтримується "наперед" - коли користувач
    додасть 50 в MONEY_REPORT_CATEGORIES/COMMANDER_MONEY_REPORT_CATEGORIES, він одразу
    стане "Додаток 2" без жодних змін коду, як тут.

    resolve_day_value_and_category (money_report_helpers.py) звіряє числове значення
    комірки з ГЛОБАЛЬНИМ mrh.MONEY_REPORT_CATEGORIES (а не з переданим у
    build_appendix_pairs_for_month аргументом categories) - тож для симуляції "точка
    50 вже існує" тут монкіпатчиться САМЕ ГЛОБАЛЬНИЙ словник, інакше комірка зі
    значенням 50 не розпізнається як день участі взагалі (те саме обмеження діє й
    для звичайного рапорту - REAL constants.MONEY_REPORT_CATEGORIES теж має отримати
    запис 50, не лише переданий сюди categories)."""
    categories = {
        30: {"general": [], "A": {
            "grounds": [], "use_brs": False, "use_brs_from_selected_folder": False, "exclude_general": [],
            "default": True, "include_to_report": True, "required_basis": False,
        }},
        50: {"general": [], "B": {
            "grounds": [], "use_brs": False, "use_brs_from_selected_folder": False, "exclude_general": [],
            "default": True, "include_to_report": True, "required_basis": False,
        }},
        100: {"general": [], "C": {
            "grounds": [], "use_brs": False, "use_brs_from_selected_folder": False, "exclude_general": [],
            "default": True, "include_to_report": True, "required_basis": False,
        }},
    }
    monkeypatch.setitem(mrh.MONEY_REPORT_CATEGORIES, 50, categories[50])
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [
        ["Тридцятник Іван", "ВД"], ["П'ятдесятник Петро", "ВД"], ["Сотник Сидір", "ВД"],
    ])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [
        ["Тридцятник Іван", 30], ["П'ятдесятник Петро", 50], ["Сотник Сидір", 100],
    ])

    pairs = report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [], categories)

    assert [(point, appendix_number) for point, appendix_number, _exclude, _add in pairs] == [(100, 1), (50, 2), (30, 3)]


def test_build_appendix_pairs_for_month_appendix_numbers_stay_fixed_when_point_50_absent(tmp_path):
    """РЕГРЕСІЯ (двічі підтверджено користувачем): коли точка 50 БЕЗ ЗМІН цього
    місяця (звичайний, майже завжди реальний випадок - 50 ще навіть не введена
    в MONEY_REPORT_CATEGORIES), а точки 100 і 30 - ОБИДВІ мають зміни - точка 30
    однаково лишається "Додатком 3", а НЕ "зсувається" на "Додаток 2" через те,
    що 50 просто відсутня в результаті. "Додаток 2" тут не з'являється ВЗАГАЛІ."""
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", "ВД"], ["Другий Другий", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 100], ["Другий Другий", 30]])

    pairs = report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [], _CATEGORIES)

    assert [(point, appendix_number) for point, appendix_number, _exclude, _add in pairs] == [(100, 1), (30, 3)]


def test_build_appendix_pairs_for_month_content_search_applies_only_to_point_100(tmp_path, monkeypatch):
    """content_search - реальний .docx у папці, що згадує ТРИНАДЦЯТОГО у тексті: пункт
    100 (де ТРИНАДЦЯТИЙ реально змінився) отримує цю знахідку в ПІДСТАВІ, пункт 30 (де
    змінився ЧОТИРНАДЦЯТИЙ, чиє ім'я НЕ згадане в жодному файлі) - ні, навіть якщо в тій
    самій папці лежить документ, що згадує ЧОТИРНАДЦЯТИМ під іншим номером - лише реальна
    згадка в тексті вирішує, а не сам факт присутності людини серед учасників.

    LOG_WAR/BN монкіпатчено в [] - див. test_delta_rows_for_point_ignores_synced_brs_for_every_point:
    без цього ПІДСТАВА точки 100/30 містила б ще й реальні виробничі рядки LOG_WAR/BN,
    заважаючи перевірити точний ефект саме content_search."""
    monkeypatch.setattr(report_changes, "MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR", [])
    monkeypatch.setattr(report_changes, "MONEY_REPORT_GENERAL_REFERENCES_BN", [])
    _write_docx(tmp_path / "№79 ЗАВДАННЯ 05.06.2026.docx", paragraphs=["ТРИНАДЦЯТИЙ Тринадцятий Тринадцятий"])

    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [
        ["ТРИНАДЦЯТИЙ Тринадцятий Тринадцятий", 30], ["Чотирнадцятий Чотирнадцятий", "ВД"],
    ])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [
        ["ТРИНАДЦЯТИЙ Тринадцятий Тринадцятий", 100], ["Чотирнадцятий Чотирнадцятий", 30],
    ])

    content_search = {"folder": str(tmp_path), "scope": "all"}
    pairs = report_changes.build_appendix_pairs_for_month(
        str(prev_path), str(actual_path), [], _CATEGORIES, content_search,
    )
    by_point = {point: (exclude_rows, add_rows) for point, _num, exclude_rows, add_rows in pairs}

    point_100_add = by_point[100][1]
    assert [row["ПІДСТАВА"] for row in point_100_add if row["ПІБ"] == "ТРИНАДЦЯТИЙ Тринадцятий Тринадцятий"] == [
        f"{_BR} №79 від 05.06.2026",
    ]

    point_30_add = by_point[30][1]
    assert [row["ПІДСТАВА"] for row in point_30_add if row["ПІБ"] == "Чотирнадцятий Чотирнадцятий"] == [""]


def test_build_appendix_pairs_for_month_170_folds_into_point_100_not_its_own_appendix(tmp_path, monkeypatch):
    """170 (і 70) НЕ мають своєї окремої точки в цьому розділі - як і в звичайному
    щомісячному рапорті (_COMBINED_TARGET_CELL_VALUES, content/money_report_helpers.py),
    їхні дні рахуються ДОДАТКОВИМ рядком усередині пункту 100. Раніше 170 помилково
    отримував ВЛАСНИЙ "Додаток 2" - той самий рядок дублювався і в 1, і в 2; тепер
    ЛИШЕ один - у Додатку 1 (точка 100), з тим самим content_search у ПІДСТАВІ.

    LOG_WAR/BN монкіпатчено в [] - див. test_delta_rows_for_point_ignores_synced_brs_for_every_point."""
    monkeypatch.setattr(report_changes, "MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR", [])
    monkeypatch.setattr(report_changes, "MONEY_REPORT_GENERAL_REFERENCES_BN", [])
    _write_docx(tmp_path / "№79 ЗАВДАННЯ 05.06.2026.docx", paragraphs=["СІМНАДЦЯТИЙ Сімнадцятий Сімнадцятий"])

    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["СІМНАДЦЯТИЙ Сімнадцятий Сімнадцятий", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["СІМНАДЦЯТИЙ Сімнадцятий Сімнадцятий", 170]])

    content_search = {"folder": str(tmp_path), "scope": "all"}
    pairs = report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [], _CATEGORIES, content_search)
    by_point = {point: (exclude_rows, add_rows) for point, _num, exclude_rows, add_rows in pairs}

    assert set(by_point) == {100}
    point_100_add = by_point[100][1]
    assert [row["ПІДСТАВА"] for row in point_100_add] == [f"{_BR} №79 від 05.06.2026"]


def test_build_appendix_pairs_for_month_collects_missing_document_coverage_warnings(tmp_path):
    """Для пункту 100 (у _FOLDER_BASED_GROUNDS_POINTS) - якщо в actual-файлі є
    ДВА дні участі, а документ (ЗАВДАННЯ) з ПІБ людини - лише на ОДИН із них -
    попередження саме про день БЕЗ документа збирається в переданий список
    (а НЕ друкується в термінал - підтверджено користувачем); без переданого
    списку (missing_coverage_warnings=None, за замовчуванням) - нічого не збирається."""
    D3 = datetime(2026, 6, 3)
    _write_docx(tmp_path / "№79 ЗАВДАННЯ 01.06.2026.docx", paragraphs=["ШІСТНАДЦЯТИЙ Шістнадцятий Шістнадцятий"])

    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1, D3], [["ШІСТНАДЦЯТИЙ Шістнадцятий Шістнадцятий", "ВД", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1, D3], [["ШІСТНАДЦЯТИЙ Шістнадцятий Шістнадцятий", 100, 100]])

    content_search = {"folder": str(tmp_path), "scope": "all"}
    warnings = []
    report_changes.build_appendix_pairs_for_month(
        str(prev_path), str(actual_path), [], _CATEGORIES, content_search, missing_coverage_warnings=warnings,
    )

    assert len(warnings) == 1
    assert "03.06.2026" in warnings[0]
    assert "ШІСТНАДЦЯТИЙ" in warnings[0]
    assert "01.06.2026" not in warnings[0]

    # Без переданого списку (типове використання, коли попередження не потрібні) -
    # не падає і нічого нікуди не пише.
    report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [], _CATEGORIES, content_search)


def test_build_appendix_pairs_for_month_missing_coverage_warning_includes_rank(tmp_path):
    """rank_by_pib/subdivision_by_pib (ЗВАННЯ/ПІДРОЗДІЛ з actual-рядка)
    підставляються ПЕРЕД ПІБ у самому попередженні, у порядку "підрозділ
    звання ПІБ" - підтверджено користувачем."""
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1],
        [["Підрозділ 1", "Стрілець", "сержант", "ТРЕТІЙ Третій Третій", "ВД"]],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1],
        [["Підрозділ 1", "Стрілець", "сержант", "ТРЕТІЙ Третій Третій", 100]],
    )

    content_search = {"folder": str(tmp_path), "scope": "all"}
    warnings = []
    report_changes.build_appendix_pairs_for_month(
        str(prev_path), str(actual_path), [], _CATEGORIES, content_search, missing_coverage_warnings=warnings,
    )

    assert warnings == ["Немає документа (ЩОДЕННА/ЗАВДАННЯ) за 01.06.2026 для Підрозділ 1 сержант ТРЕТІЙ Третій Третій - підстава на цей день відсутня."]


def test_build_appendix_pairs_for_month_point_without_changes_is_absent(tmp_path):
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])

    pairs = report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [], _CATEGORIES)

    assert pairs == []


def test_build_appendix_pairs_for_month_ignores_pseudo_points_outside_priority_order(tmp_path, monkeypatch):
    """Регресія: псевдо-пункти на кшталт "100_ШП"/"NOT_PAID" - звичайні top-level
    ключі MONEY_REPORT_CATEGORIES (потрібні лише щомісячному рапорту,
    generate_report_for_get_money._build_categories, де НОВУ точку підхоплює
    _ordered_points АВТОМАТИЧНО) - НЕ повинні самі собою ставати новим додатком
    тут, навіть якщо людина реально змінилась між prev/actual саме за цією
    категорією: додатки цього розділу обмежені РІВНО _CHANGES_POINT_PRIORITY_ORDER
    (100, 50, 30). Без цієї гарантії людина з "100_ШП" помилково опинялась у
    зайвому "Додатку 3", хоча реальних додатків мало бути лише два (100 і 30)."""
    categories = {
        **_CATEGORIES,
        "100_ШП": {
            "general": [],
            "ЛІКУВАННЯ_ПО_ПОРАНЕННЮ": {
                "grounds": [], "use_brs": False, "use_brs_from_selected_folder": False, "exclude_general": [],
                "default": True, "include_to_report": True, "required_basis": False,
            },
        },
    }
    monkeypatch.setitem(mrh.MONEY_REPORT_CATEGORIES, "100_ШП", categories["100_ШП"])

    # D1 - переходить у псевдо-пункт "100_ШП" (не повинен створити свій додаток);
    # D2 - ОКРЕМИЙ, СПРАВЖНІЙ новий 30-день (щоб точка 30 взагалі мала що показати -
    # "Викласти в новій редакції" рахує лише НОВІ дні, а сам факт "втратив D1"
    # більше не рахується як зміна цього пункту).
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1, D2], [["ЧЕТВЕРТИЙ Четвертий", 30, "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1, D2], [["ЧЕТВЕРТИЙ Четвертий", "100_ШП", 30]])

    pairs = report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [], categories)

    assert {point for point, *_ in pairs} == {30}


def test_build_appendix_pairs_for_month_uses_global_categories_by_default(tmp_path, monkeypatch):
    monkeypatch.setattr(report_changes, "MONEY_REPORT_CATEGORIES", _CATEGORIES)
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])

    pairs = report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [])

    assert [point for point, *_ in pairs] == [30]


# -------------------------
# commander_and_tvo_only - КБ/ТВО виключені з головного рапорту, включені у власний
# -------------------------
def test_build_appendix_pairs_for_month_excludes_commander_by_default(tmp_path):
    """За замовчуванням (commander_and_tvo_only=False, головний рапорт) - штатний
    командир батальйону (ПОСАДА == HIGHER_COMMANDER_TITLE) НЕ фігурує в жодній
    зміні - для нього є окремий рапорт від першої особи
    (generate_report_for_commander_money.py), підтверджено користувачем."""
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1],
        [["Штаб", HIGHER_COMMANDER_TITLE, "підполковник", "ЧЕТВЕРТИЙ Четвертий", 30]],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1],
        [["Штаб", HIGHER_COMMANDER_TITLE, "підполковник", "ЧЕТВЕРТИЙ Четвертий", "ВД"]],
    )

    pairs = report_changes.build_appendix_pairs_for_month(str(prev_path), str(actual_path), [], _CATEGORIES)

    assert pairs == []


def test_build_appendix_pairs_for_month_commander_and_tvo_only_keeps_only_commander(tmp_path):
    """commander_and_tvo_only=True (рапорт КБ/ТВО) - навпаки, лишає ЛИШЕ зміни
    штатного командира/ТВО - решта підлеглих не фігурує - підтверджено
    користувачем."""
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1],
        [
            ["Штаб", HIGHER_COMMANDER_TITLE, "підполковник", "ЧЕТВЕРТИЙ Четвертий", "ВД"],
            ["Підрозділ 1", "Стрілець", "сержант", "ШОСТИЙ Шостий", "ВД"],
        ],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", D1],
        [
            ["Штаб", HIGHER_COMMANDER_TITLE, "підполковник", "ЧЕТВЕРТИЙ Четвертий", 30],
            ["Підрозділ 1", "Стрілець", "сержант", "ШОСТИЙ Шостий", 30],
        ],
    )

    pairs = report_changes.build_appendix_pairs_for_month(
        str(prev_path), str(actual_path), [], _CATEGORIES, commander_and_tvo_only=True,
    )

    assert len(pairs) == 1
    point, _num, _exclude_rows, add_rows = pairs[0]
    assert point == 30
    assert [r["ПІБ"] for r in add_rows] == ["ЧЕТВЕРТИЙ Четвертий"]


# -------------------------
# _basis_appeared_add_rows - ретроактивне 100_СПЕЦКОНТИНГЕНТ/100_БПШП, коли
# ПІДСТАВИ з'явилась лише в actual-файлі
# -------------------------
def test_basis_appeared_add_rows_adds_retroactive_change_when_basis_newly_documented(tmp_path):
    """BASIS_REQUIRED_POINTS ('100_СПЕЦКОНТИНГЕНТ') - людина, чий статус ("полон")
    лишався НЕЗМІННИМ між prev і actual, але в prev-місяці ще не було "ПІДСТАВИ
    СПЕЦКОНТИНГЕНТУ" (тож звичайний щомісячний рапорт за ТОЙ місяць виключив би
    її цілком - generate_report_for_get_money._build_categories), а в actual
    вона вже з'явилась - має РЕТРОАКТИВНО з'явитись "як у звичайному рапорті"
    (SECTIONS у resources/data.json, ПОЗА межами "Прошу внести зміни...") - з
    датами PREV, але з ПІДСТАВОЮ, яка щойно з'явилась (actual) - підтверджено
    користувачем. Колонка - ВЛАСНА для цього статусу (BASIS_REQUIRED_POINT_COLUMN_NAMES,
    constants.py), а не спільна "ПІДСТАВИ"."""
    headers = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", "ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ", "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ", D1]
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ВОСЬМИЙ Восьмий", "01.06.2026", "", "полон"]],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ВОСЬМИЙ Восьмий", "01.06.2026", "Наказ №5 від 15.06.2026", "полон"]],
    )

    sections = report_changes._basis_appeared_extra_points_for_month(
        str(prev_path), str(actual_path), [], constants.MONEY_REPORT_CATEGORIES, False, None,
    )
    by_point = {section["point"]: section["rows"] for section in sections}

    assert "100_СПЕЦКОНТИНГЕНТ" in by_point
    add_rows = by_point["100_СПЕЦКОНТИНГЕНТ"]
    assert [r["ПІБ"] for r in add_rows] == ["ВОСЬМИЙ Восьмий"]
    assert add_rows[0]["ПІДСТАВА"] == "Наказ №5 від 15.06.2026"


def test_basis_appeared_add_rows_absent_when_basis_already_present_in_prev(tmp_path):
    """Якщо ПІДСТАВИ вже БУЛА заповнена в prev - це вже "не зміна" (людина й так
    потрапила б у тодішній рапорт) - жодного ретроактивного пункту не
    з'являється."""
    headers = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ", D1]
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ВОСЬМИЙ Восьмий", "Наказ №5 від 15.06.2026", "полон"]],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ВОСЬМИЙ Восьмий", "Наказ №5 від 15.06.2026", "полон"]],
    )

    sections = report_changes._basis_appeared_extra_points_for_month(
        str(prev_path), str(actual_path), [], constants.MONEY_REPORT_CATEGORIES, False, None,
    )

    assert "100_СПЕЦКОНТИНГЕНТ" not in {section["point"] for section in sections}


def test_basis_appeared_add_rows_absent_when_basis_still_missing_in_actual(tmp_path):
    """Якщо ПІДСТАВИ й далі порожня в actual - людина не мала б потрапити нікуди
    (звичайне виключення - результат ЦЬОГО місяця обробляється окремо, не тут) -
    жодного ретроактивного пункту."""
    headers = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ", D1]
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ВОСЬМИЙ Восьмий", "", "полон"]],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ВОСЬМИЙ Восьмий", "", "полон"]],
    )

    sections = report_changes._basis_appeared_extra_points_for_month(
        str(prev_path), str(actual_path), [], constants.MONEY_REPORT_CATEGORIES, False, None,
    )

    assert "100_СПЕЦКОНТИНГЕНТ" not in {section["point"] for section in sections}


def test_basis_appeared_add_rows_covers_vpbp_too(tmp_path):
    """BASIS_REQUIRED_POINTS (constants.py, керується користувачем без змін коду) -
    та сама ретроактивна логіка стосується і 100_ВПБП, не лише 100_СПЕЦКОНТИНГЕНТ/
    100_БПШП - підтверджено користувачем. Колонка - ВЛАСНА для цього статусу
    ("ПІДСТАВИ ВПБП", BASIS_REQUIRED_POINT_COLUMN_NAMES, constants.py)."""
    headers = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", "ПІДСТАВИ ВПБП", D1]
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ДЕСЯТИЙ Десятий", "", "ВПБП"]],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ДЕСЯТИЙ Десятий", "Наказ №9 від 20.06.2026", "ВПБП"]],
    )

    sections = report_changes._basis_appeared_extra_points_for_month(
        str(prev_path), str(actual_path), [], constants.MONEY_REPORT_CATEGORIES, False, None,
    )
    by_point = {section["point"]: section["rows"] for section in sections}

    assert "100_ВПБП" in by_point
    add_rows = by_point["100_ВПБП"]
    assert [r["ПІБ"] for r in add_rows] == ["ДЕСЯТИЙ Десятий"]
    assert add_rows[0]["ПІДСТАВА"] == "Наказ №9 від 20.06.2026"


def test_basis_appeared_add_rows_columns_do_not_cross_contaminate(tmp_path):
    """РЕГРЕСІЯ (підтверджено користувачем): "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ" заповнена
    в actual для людини зі статусом "ВПБП" - НЕ повинна хибно "розблокувати" її
    для 100_ВПБП (та навпаки) - кожен статус BASIS_REQUIRED_POINTS дивиться
    ЛИШЕ на СВОЮ колонку (BASIS_REQUIRED_POINT_COLUMN_NAMES), а не на будь-яку
    непорожню підставу взагалі. Раніше всі три статуси ділили ОДНУ спільну
    "ПІДСТАВИ" - підстава СПЕЦКОНТИНГЕНТУ помилково "розблоковувала" б і ВПБП."""
    headers = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ", "ПІДСТАВИ ВПБП", D1]
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ДЕСЯТИЙ Десятий", "", "", "ВПБП"]],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ДЕСЯТИЙ Десятий", "Наказ №9 від 20.06.2026", "", "ВПБП"]],
    )

    sections = report_changes._basis_appeared_extra_points_for_month(
        str(prev_path), str(actual_path), [], constants.MONEY_REPORT_CATEGORIES, False, None,
    )

    assert "100_ВПБП" not in {section["point"] for section in sections}


def test_new_participation_rows_for_basis_required_point_person_absent_from_prev(tmp_path):
    """РЕГРЕСІЯ (реальний випадок, підтверджений користувачем) - людина БЕЗ
    жодного дня "100_БПШП" у prev (там -
    зовсім інша категорія), у якої з'явились НОВІ дні в actual - МАЄ з'явитись
    (на відміну від _basis_appeared_add_rows, яка тут нічого не знайшла б, бо
    будує рядки ЗІ САМОГО prev_rows) - за умови непорожньої ВЛАСНОЇ підстави."""
    headers = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", "ПІДСТАВИ БПШП", D1, D2]
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ШОСТИЙ Шостий", "", 170, 170]],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ШОСТИЙ Шостий", "Довідка про травму", "100_БПШП", "100_БПШП"]],
    )

    sections = report_changes._basis_appeared_extra_points_for_month(
        str(prev_path), str(actual_path), [], constants.MONEY_REPORT_CATEGORIES, False, None,
    )
    by_point = {section["point"]: section["rows"] for section in sections}

    assert "100_БПШП" in by_point
    rows = by_point["100_БПШП"]
    assert [r["ПІБ"] for r in rows] == ["ШОСТИЙ Шостий"]
    assert rows[0]["ДНІ"] == 2
    assert rows[0]["ПІДСТАВА"] == "Довідка про травму"


def test_new_participation_rows_for_basis_required_point_excluded_without_own_basis(tmp_path):
    """Та сама нова участь, АЛЕ без непорожньої ВЛАСНОЇ підстави ("ПІДСТАВИ ВПБП"
    порожня, навіть якщо "ПІДСТАВИ БПШП" заповнена для ІНШОГО статусу) -
    ВИКЛЮЧАЄТЬСЯ ЦІЛКОМ (required_basis) - підтверджено користувачем: НЕ
    підставляти чужу колонку як запасний варіант."""
    headers = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", "ПІДСТАВИ БПШП", "ПІДСТАВИ ВПБП", D1]
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ОДИНАДЦЯТИЙ Одинадцятий", "", "", 170]],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ОДИНАДЦЯТИЙ Одинадцятий", "Довідка про травму", "", "100_ВПБП"]],
    )

    sections = report_changes._basis_appeared_extra_points_for_month(
        str(prev_path), str(actual_path), [], constants.MONEY_REPORT_CATEGORIES, False, None,
    )

    assert "100_ВПБП" not in {section["point"] for section in sections}


def test_basis_appeared_extra_points_returns_empty_when_a_file_has_no_date_columns(tmp_path):
    """Один із файлів (тут - actual) не має ЖОДНОЇ колонки-дати (напр.
    зіпсований/не той файл) - _rows_from_check_file повертає порожній
    date_columns, і функція безпечно повертає [] замість падіння на
    подальшій обробці, якій потрібна хоча б одна дата з обох боків."""
    headers_with_date = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ", D1]
    headers_without_date = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ"]
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", headers_with_date,
        [["Підрозділ 1", "Стрілець", "сержант", "ВОСЬМИЙ Восьмий", "", "полон"]],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", headers_without_date,
        [["Підрозділ 1", "Стрілець", "сержант", "ВОСЬМИЙ Восьмий", "Наказ №5 від 15.06.2026"]],
    )

    sections = report_changes._basis_appeared_extra_points_for_month(
        str(prev_path), str(actual_path), [], constants.MONEY_REPORT_CATEGORIES, False, None,
    )

    assert sections == []


def test_resolve_day_value_and_category_accepts_shpbp_typo_spelling():
    """РЕГРЕСІЯ (реальний випадок): комірка дня в реальному ОБЛІК.xlsx містить
    буквальний текст "100_ШПБП" (літери переставлені) замість "100_БПШП" - той
    самий типовий одрук, що й "ПІДСТАВИ ШПБП" (BASIS_REQUIRED_POINT_COLUMN_NAMES) -
    має розпізнаватись як точка "100_БПШП", а не губитись як нерозпізнаний текст."""
    point, category = mrh.resolve_day_value_and_category("100_ШПБП", None)
    assert point == "100_БПШП"
    assert category == "БПШП|100_БПШП|100_ШПБП"


# -------------------------
# build_changes_entries - фільтрація/сортування по місяцях
# -------------------------
def test_build_changes_entries_skips_pair_when_actual_file_has_no_dates(tmp_path, monkeypatch):
    monkeypatch.setattr(report_changes, "MONTH", "07")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ"], [["Перший Перший"]])
    pairs = [{"month_token": "06", "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused"}]

    entries = report_changes.build_changes_entries([], _CATEGORIES, changes_file_pairs=pairs)

    assert entries == []

def test_build_changes_entries_skips_pair_with_no_actual_changes(tmp_path, monkeypatch):
    monkeypatch.setattr(report_changes, "MONTH", "07")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])
    pairs = [{"month_token": "06", "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused"}]

    entries = report_changes.build_changes_entries([], _CATEGORIES, changes_file_pairs=pairs)

    assert entries == []


def test_build_changes_entries_skips_pair_for_current_or_future_month(tmp_path, monkeypatch):
    """actual-файл описує ПОТОЧНИЙ (чи майбутній) відносно обраного місяць -
    НЕ потрапляє в "зміни за попередні місяці" взагалі (навіть якщо містить
    реальні зміни)."""
    monkeypatch.setattr(report_changes, "MONTH", "06")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", "ВД"]])
    pairs = [{"month_token": "06", "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused"}]

    entries = report_changes.build_changes_entries([], _CATEGORIES, changes_file_pairs=pairs)

    assert entries == []


# -------------------------
# _month_date_range / _resolve_admin_order_reference
# -------------------------
def test_month_date_range_returns_first_and_last_calendar_day():
    start, end = report_changes._month_date_range(2026, 6)
    assert start == datetime(2026, 6, 1)
    assert end == datetime(2026, 6, 30)


def test_resolve_admin_order_reference_returns_none_when_no_match():
    references = [{"start": "01.05.2026", "end": "31.05.2026", "lines": ["№10 від 01.05.2026"]}]
    assert report_changes._resolve_admin_order_reference(datetime(2026, 6, 1), datetime(2026, 6, 30), references) is None


def test_resolve_admin_order_reference_returns_first_overlapping_match():
    references = [
        {"start": "01.05.2026", "end": "31.05.2026", "lines": ["№10 від 01.05.2026"]},
        {"start": "01.06.2026", "end": "30.06.2026", "lines": ["№45 від 05.06.2026"]},
    ]
    result = report_changes._resolve_admin_order_reference(datetime(2026, 6, 1), datetime(2026, 6, 30), references)
    assert result == "№45 від 05.06.2026"


def test_build_changes_entries_attaches_matching_order_reference(tmp_path, monkeypatch):
    monkeypatch.setattr(report_changes, "MONTH", "07")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    monkeypatch.setattr(report_changes, "MONEY_REPORT_CHANGES_ORDER_REFERENCES", [
        {"start": "01.06.2026", "end": "30.06.2026", "lines": ["№45 від 05.06.2026"]},
    ])
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])
    pairs = [{"month_token": "06", "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused"}]

    entries = report_changes.build_changes_entries([], _CATEGORIES, changes_file_pairs=pairs)

    assert entries[0]["order_reference"] == "№45 від 05.06.2026"


def test_build_changes_entries_reads_content_search_folder_off_pair(tmp_path, monkeypatch):
    """sync.changes_folder.prompt_grounds_folders_for_changes записує
    content_search_folder/content_search_scope прямо в елемент changes_file_pairs -
    build_changes_entries має підхопити їх і передати в build_appendix_pairs_for_month
    (пункт 100 отримує знахідку з реального .docx у цій папці).

    LOG_WAR/BN монкіпатчено в [] - див. test_delta_rows_for_point_ignores_synced_brs_for_every_point."""
    monkeypatch.setattr(report_changes, "MONTH", "07")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    monkeypatch.setattr(report_changes, "MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR", [])
    monkeypatch.setattr(report_changes, "MONEY_REPORT_GENERAL_REFERENCES_BN", [])
    _write_docx(tmp_path / "№79 ЗАВДАННЯ 05.06.2026.docx", paragraphs=["Перший Перший"])
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 100]])
    pairs = [{
        "month_token": "06", "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused",
        "content_search_folder": str(tmp_path), "content_search_scope": "all",
    }]

    entries = report_changes.build_changes_entries([], _CATEGORIES, changes_file_pairs=pairs)

    add_rows = entries[0]["appendix_pairs"][0][3]
    assert add_rows[0]["ПІДСТАВА"] == f"{_BR} №79 від 05.06.2026"


def test_build_changes_entries_order_reference_none_when_no_match(tmp_path, monkeypatch):
    monkeypatch.setattr(report_changes, "MONTH", "07")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    monkeypatch.setattr(report_changes, "MONEY_REPORT_CHANGES_ORDER_REFERENCES", [])
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])
    pairs = [{"month_token": "06", "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused"}]

    entries = report_changes.build_changes_entries([], _CATEGORIES, changes_file_pairs=pairs)

    assert entries[0]["order_reference"] is None


def test_build_changes_entries_orders_nearest_month_first(tmp_path, monkeypatch):
    """Обрано 08 (серпень) - записи за 07 (липень, 1 місяць тому) мають йти
    ПЕРЕД записами за 06 (червень, 2 місяці тому), незалежно від порядку у
    вхідному changes_file_pairs."""
    monkeypatch.setattr(report_changes, "MONTH", "08")
    monkeypatch.setattr(report_changes, "YEAR", "2026")

    def _pair(month_token, month_day_date, pib):
        prev_path = _write_xlsx(tmp_path / f"prev_{month_token}_ОБЛІК.xlsx", ["ПІБ", month_day_date], [[pib, "ВД"]])
        actual_path = _write_xlsx(tmp_path / f"actual_{month_token}_ОБЛІК.xlsx", ["ПІБ", month_day_date], [[pib, 30]])
        return {"month_token": month_token, "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused"}

    june_pair = _pair("06", datetime(2026, 6, 1), "Червневий Іван")
    july_pair = _pair("07", datetime(2026, 7, 1), "Липневий Петро")

    entries = report_changes.build_changes_entries([], _CATEGORIES, changes_file_pairs=[june_pair, july_pair])

    assert len(entries) == 2
    first_pib = entries[0]["appendix_pairs"][0][3][0]["ПІБ"]
    second_pib = entries[1]["appendix_pairs"][0][3][0]["ПІБ"]
    assert first_pib == "Липневий Петро"
    assert second_pib == "Червневий Іван"


def test_build_changes_entries_defaults_to_no_pairs():
    """Без явного changes_file_pairs - порожній кортеж за замовчуванням, не падає."""
    assert report_changes.build_changes_entries([]) == []


# -------------------------
# _extra_point_rows_for_month / build_changes_entries "extra_points" (70/170
# "як у звичайному рапорті", поза межами "Прошу внести зміни...")
# -------------------------
_CATEGORIES_WITH_70 = {
    **_CATEGORIES,
    70: {"general": [], "ОБОРОНА": {
        "grounds": [], "use_brs": False, "use_brs_from_selected_folder": False, "exclude_general": [],
        "default": True, "include_to_report": True, "required_basis": False,
    }},
}


def test_extra_point_rows_for_month_includes_only_genuinely_own_valued_days(tmp_path):
    """Точка 70 рахує ЛИШЕ дні, буквально позначені 70 - день, позначений 100
    (уже показаний у Додатку 100), сюди НЕ входить, навіть для тієї самої людини."""
    prev_path = _write_xlsx(tmp_path / "prev_07_ОБЛІК.xlsx", ["ПІБ", D1, D2], [["Перший Перший", "ВД", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_07_ОБЛІК.xlsx", ["ПІБ", D1, D2], [["Перший Перший", 70, 100]])

    sections = report_changes._extra_point_rows_for_month(
        str(prev_path), str(actual_path), [], _CATEGORIES_WITH_70, False, None, None,
    )

    assert len(sections) == 1
    assert sections[0]["point"] == 70
    assert sections[0]["rows"][0]["ДНІ"] == 1
    assert sections[0]["rows"][0]["ПЕРІОД"] == "01.06.2026-01.06.2026"


def test_extra_point_rows_for_month_ignores_synced_brs(tmp_path, monkeypatch):
    """Навіть якщо точка 70 налаштована з use_brs=True (як у звичайному рапорті),
    "extra_points" НЕ повинні тягнути дані з NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK -
    підтверджено користувачем: "чомусь підтягує дані з NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY...
    хоча не мало би". Раніше цей шлях (на відміну від diff_point_categories) помилково
    минав _changes_category_config, викликаючи resolve_category_config напряму."""
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {"01.06.2026": {"бат": "999", "посилання_брг": "", "посилання_бат": ""}})
    categories = {70: {"general": [], "ОБОРОНА": {
        "grounds": [], "use_brs": True, "use_brs_from_selected_folder": True, "exclude_general": [],
        "default": True, "include_to_report": True, "required_basis": False,
    }}}
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 70]])

    sections = report_changes._extra_point_rows_for_month(str(prev_path), str(actual_path), [], categories, False, None, None)

    assert sections[0]["rows"][0]["ПІДСТАВА"] == ""


def test_extra_point_rows_for_month_absent_point_produces_no_section(tmp_path):
    """Точки 70/170 немає серед categories взагалі (як у _CATEGORIES без 70) -
    жодного extra-пункту, навіть якщо в файлі є день зі значенням 70."""
    prev_path = _write_xlsx(tmp_path / "prev_07_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_07_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 70]])

    sections = report_changes._extra_point_rows_for_month(str(prev_path), str(actual_path), [], _CATEGORIES, False, None, None)

    assert sections == []


def test_extra_point_rows_for_month_defaults_to_global_categories(tmp_path):
    """categories=None - підставляється constants.MONEY_REPORT_CATEGORIES (де 70
    реально існує), а не падає з KeyError/TypeError."""
    prev_path = _write_xlsx(tmp_path / "prev_07_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_07_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 70]])

    sections = report_changes._extra_point_rows_for_month(str(prev_path), str(actual_path), [], None, False, None, None)

    assert [s["point"] for s in sections] == [70]


def test_extra_point_rows_for_month_no_dates_returns_empty(tmp_path):
    prev_path = _write_xlsx(tmp_path / "prev_07_ОБЛІК.xlsx", ["ПІБ"], [["Перший Перший"]])
    actual_path = _write_xlsx(tmp_path / "actual_07_ОБЛІК.xlsx", ["ПІБ"], [["Перший Перший"]])

    assert report_changes._extra_point_rows_for_month(str(prev_path), str(actual_path), [], _CATEGORIES_WITH_70, False, None, None) == []


def test_extra_point_rows_for_month_absent_when_unchanged_from_prev(tmp_path):
    """РЕГРЕСІЯ (підтверджено користувачем): "якщо в файлі 170 не змінювалось,
    то не потрібно добавляти цей пункт" - людина з ІДЕНТИЧНИМ 70-днем і в prev,
    і в actual (жодного НОВОГО дня) - НЕ рендериться взагалі, навіть якщо вона
    Й ДАЛІ має цей день в actual."""
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 70]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 70]])

    sections = report_changes._extra_point_rows_for_month(
        str(prev_path), str(actual_path), [], _CATEGORIES_WITH_70, False, None, None,
    )

    assert sections == []


def test_extra_point_rows_for_month_applies_content_search_grounds(tmp_path):
    _write_docx(tmp_path / "№79 ЗАВДАННЯ 01.07.2026.docx", paragraphs=["ВІСІМНАДЦЯТИЙ Вісімнадцятий Вісімнадцятий"])
    prev_path = _write_xlsx(tmp_path / "prev_07_ОБЛІК.xlsx", ["ПІБ", D1], [["ВІСІМНАДЦЯТИЙ Вісімнадцятий Вісімнадцятий", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_07_ОБЛІК.xlsx", ["ПІБ", D1], [["ВІСІМНАДЦЯТИЙ Вісімнадцятий Вісімнадцятий", 70]])
    content_search = {"folder": str(tmp_path), "scope": "all"}

    sections = report_changes._extra_point_rows_for_month(
        str(prev_path), str(actual_path), [], _CATEGORIES_WITH_70, False, None, content_search,
    )

    assert sections[0]["rows"][0]["ПІДСТАВА"] == f"{_BR} №79 від 01.07.2026"


def test_build_changes_entries_covers_extra_point_dates_via_point_100_check(tmp_path, monkeypatch):
    """_extra_point_rows_for_month (70/170) саме НЕ перевіряє покриття документами
    само - воно вже охоплене build_appendix_pairs_for_month'ів перевіркою для
    точки 100 (_COMBINED_TARGET_CELL_VALUES включає 70/170) - build_changes_entries
    (реальний виклик з обома функціями разом) все одно ловить пропущений день
    людини, чиї дні фігурують ЛИШЕ в extra_points (немає жодної зміни для Додатку 100)."""
    monkeypatch.setattr(report_changes, "MONTH", "08")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    D3 = datetime(2026, 6, 3)
    _write_docx(tmp_path / "№79 ЗАВДАННЯ 01.06.2026.docx", paragraphs=["ДВАДЦЯТИЙ Двадцятий Двадцятий"])
    # prev - "100" (НЕ "ВД") на ТІ САМІ дати - 70/170 УЖЕ входять у комбіноване
    # target_cell_values точки 100 ({70,100,170}), тож "70" замість "100" на ТУ
    # САМУ дату НЕ дає точці 100 жодного НОВОГО дня (сама дата вже рахувалась у
    # комбінованій участі), а ось для САМОСТІЙНОЇ точки 70 (target_cell_values=
    # {70} лише) це ГЕНУЇНО новий день - за НОВИМ правилом ("якщо 170/70 не
    # змінювалось - не додавати") лише ТАКА зміна лишає extra_points непорожнім
    # при порожньому appendix_pairs.
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1, D3], [["ДВАДЦЯТИЙ Двадцятий Двадцятий", 100, 100]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1, D3], [["ДВАДЦЯТИЙ Двадцятий Двадцятий", 70, 70]])
    pairs = [{
        "month_token": "06", "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused",
        "content_search_folder": str(tmp_path), "content_search_scope": "all",
    }]

    entries = report_changes.build_changes_entries([], _CATEGORIES_WITH_70, changes_file_pairs=pairs)

    assert len(entries) == 1
    assert entries[0]["appendix_pairs"] == []
    assert [s["point"] for s in entries[0]["extra_points"]] == [70]
    warnings = entries[0]["missing_coverage_warnings"]
    assert len(warnings) == 1
    assert "03.06.2026" in warnings[0]
    assert "01.06.2026" not in warnings[0]


def test_build_changes_entries_includes_extra_points_even_without_appendix_changes(tmp_path, monkeypatch):
    """Місяць, де НЕМАЄ жодної зміни в Додатках, але Є НОВИЙ 70-день у actual
    (якого НЕ БУЛО в prev - там та сама дата рахувалась як "100", яке 70/170
    вже охоплюють комбіновано, тож для точки 100 це НЕ новий день) - однаково
    потрапляє в результат (appendix_pairs порожній, extra_points - ні), з
    номером місяця для рендеру заголовка."""
    monkeypatch.setattr(report_changes, "MONTH", "08")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 100]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 70]])
    pairs = [{"month_token": "06", "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused"}]

    entries = report_changes.build_changes_entries([], _CATEGORIES_WITH_70, changes_file_pairs=pairs)

    assert len(entries) == 1
    assert entries[0]["appendix_pairs"] == []
    assert entries[0]["month"] == 6
    assert [s["point"] for s in entries[0]["extra_points"]] == [70]


def test_build_changes_entries_pair_with_no_changes_at_all_produces_no_entry(tmp_path, monkeypatch):
    """РЕГРЕСІЯ (підтверджено користувачем): "якщо в файлі 170 не змінювалось,
    то не потрібно добавляти цей пункт" - місяць, де НІЧОГО не змінилось (ні
    Додатки, ні 70/170), НЕ потрапляє в результат УЗАГАЛІ - раніше 70/170
    рахувались НАПРЯМУ з actual (без diff проти prev), тож ця сама пара
    помилково давала б запис лише через незмінну 70-участь."""
    monkeypatch.setattr(report_changes, "MONTH", "08")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 70]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 70]])
    pairs = [{"month_token": "06", "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused"}]

    entries = report_changes.build_changes_entries([], _CATEGORIES_WITH_70, changes_file_pairs=pairs)

    assert entries == []


def test_build_changes_entries_merges_basis_required_points_into_extra_points(tmp_path, monkeypatch):
    """BASIS_REQUIRED_POINTS ("100_СПЕЦКОНТИНГЕНТ" тощо) - НІКОЛИ не потрапляють
    у "appendix_pairs" ("Додаток N") - лише в "extra_points" (об'єднано з
    70/170, якщо є), рендеряться "як у звичайному рапорті" - підтверджено
    користувачем."""
    monkeypatch.setattr(report_changes, "MONTH", "07")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    headers = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", "ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ", "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ", D1]
    prev_path = _write_xlsx(
        tmp_path / "prev_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ВОСЬМИЙ Восьмий", "01.06.2026", "", "полон"]],
    )
    actual_path = _write_xlsx(
        tmp_path / "actual_06_ОБЛІК.xlsx", headers,
        [["Підрозділ 1", "Стрілець", "сержант", "ВОСЬМИЙ Восьмий", "01.06.2026", "Наказ №5 від 15.06.2026", "полон"]],
    )
    pairs = [{"month_token": "06", "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused"}]

    entries = report_changes.build_changes_entries([], constants.MONEY_REPORT_CATEGORIES, changes_file_pairs=pairs)

    assert len(entries) == 1
    assert "100_СПЕЦКОНТИНГЕНТ" not in {point for point, _num, _e, _a in entries[0]["appendix_pairs"]}
    assert [s["point"] for s in entries[0]["extra_points"]] == ["100_СПЕЦКОНТИНГЕНТ"]
    assert [r["ПІБ"] for r in entries[0]["extra_points"][0]["rows"]] == ["ВОСЬМИЙ Восьмий"]


# -------------------------
# describe_eligible_changes_months
# -------------------------
def test_describe_eligible_changes_months_orders_nearest_first_with_ukrainian_labels(tmp_path, monkeypatch):
    """Той самий критерій елігібельності й порядок (найближчий місяць першим), що
    й build_changes_entries - лише замінює диференціацію на людський підпис
    "MM (назва місяця)" для запиту папки з підставами."""
    monkeypatch.setattr(report_changes, "MONTH", "08")
    monkeypatch.setattr(report_changes, "YEAR", "2026")

    def _pair(month_token, month_day_date, pib):
        prev_path = _write_xlsx(tmp_path / f"prev_{month_token}_ОБЛІК.xlsx", ["ПІБ", month_day_date], [[pib, 30]])
        actual_path = _write_xlsx(tmp_path / f"actual_{month_token}_ОБЛІК.xlsx", ["ПІБ", month_day_date], [[pib, "ВД"]])
        return {"month_token": month_token, "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused"}

    june_pair = _pair("06", datetime(2026, 6, 1), "Червневий Іван")
    july_pair = _pair("07", datetime(2026, 7, 1), "Липневий Петро")

    described = report_changes.describe_eligible_changes_months([june_pair, july_pair])

    assert [label for _pair, label in described] == ["07 (липень)", "06 (червень)"]
    assert described[0][0] is july_pair
    assert described[1][0] is june_pair


def test_describe_eligible_changes_months_excludes_current_or_future_month(tmp_path, monkeypatch):
    monkeypatch.setattr(report_changes, "MONTH", "06")
    monkeypatch.setattr(report_changes, "YEAR", "2026")
    prev_path = _write_xlsx(tmp_path / "prev_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", 30]])
    actual_path = _write_xlsx(tmp_path / "actual_06_ОБЛІК.xlsx", ["ПІБ", D1], [["Перший Перший", "ВД"]])
    pairs = [{"month_token": "06", "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused"}]

    assert report_changes.describe_eligible_changes_months(pairs) == []
