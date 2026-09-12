import os
import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import pytest

import constants
import content.money_report_helpers as mrh
from helpers_for_TEST_WORK import read_docx_text


# -------------------------
# Назви місяців
# -------------------------
def test_month_nominative_upper():
    assert mrh.month_nominative_upper(6) == "ЧЕРВЕНЬ"
    assert mrh.month_nominative_upper(1) == "СІЧЕНЬ"


def test_month_genitive_lower():
    assert mrh.month_genitive_lower(6) == "червня"
    assert mrh.month_genitive_lower(1) == "січня"


# -------------------------
# normalize_name
# -------------------------
def test_normalize_name_non_string_returns_empty():
    assert mrh.normalize_name(None) == ""
    assert mrh.normalize_name(123) == ""


def test_normalize_name_collapses_whitespace_and_uppercases():
    assert mrh.normalize_name("  Другий   Другий  Другий ") == "ДРУГИЙ ДРУГИЙ ДРУГИЙ"


def test_normalize_name_fixes_apostrophe_space():
    assert mrh.normalize_name("Тринадцятий В' осьмий Чотирнадцятий") == "ТРИНАДЦЯТИЙ В'ОСЬМИЙ ЧОТИРНАДЦЯТИЙ"


# -------------------------
# resolve_status_category
# -------------------------
def test_resolve_status_category_mp_subdivision_always_medic():
    assert mrh.resolve_status_category(None, "мп") == "МЕДИК"
    assert mrh.resolve_status_category("ОБОРОНА", "МП") == "МЕДИК"
    assert mrh.resolve_status_category("ОБОРОНА", "  мп  ") == "МЕДИК"


def test_resolve_status_category_by_status_value():
    # Назви категорій ніде не дублюються окремими константами - самі рядки нижче
    # мають збігатись із ключами реального constants.MONEY_REPORT_CATEGORIES.
    assert mrh.resolve_status_category("МЕДИК") == "МЕДИК"
    assert mrh.resolve_status_category("РТГр") == "РТГр"
    assert mrh.resolve_status_category("БД(СЗ)") == "БД(СЗ)"


def test_resolve_status_category_unknown_returns_none():
    assert mrh.resolve_status_category("щось інше", "Підрозділ 1") is None
    assert mrh.resolve_status_category(None, None) is None


def test_resolve_status_category_matches_regardless_of_case():
    """Статус завжди звіряється у ВЕРХНЬОМУ регістрі, а деякі ключі
    MONEY_REPORT_CATEGORIES (напр. "РТГр") мають рядкові літери - тож
    порівняння має ігнорувати регістр (інакше "РТГР" з файлу ніколи не збігся б
    із ключем "РТГр", і людина мовчки потрапляла б у catch-all замість своєї
    категорії - саме такий баг був знайдений і виправлений)."""
    assert mrh.resolve_status_category("РТГР") == "РТГр"
    assert mrh.resolve_status_category("ртгр") == "РТГр"
    assert mrh.resolve_status_category("70_РТГР") == "70_РТГр"
    assert mrh.resolve_status_category("170_ртгр") == "170_РТГр"


def test_resolve_status_category_matches_any_pipe_separated_alternative():
    """Ключ категорії ("NOT_PAID"/"Задув.|задув|задут" у реальному
    constants.py) може містити кілька варіантів написання, розділених "|" -
    БУДЬ-ЯКИЙ з них рахується як збіг, повертається ЦІЛИЙ ключ (з усіма
    варіантами), а не лише той, що фактично збігся."""
    assert mrh.resolve_status_category("задув") == "Задув.|задув|задут"
    assert mrh.resolve_status_category("Задув.") == "Задув.|задув|задут"
    assert mrh.resolve_status_category("ЗАДУТ") == "Задув.|задув|задут"
    assert mrh.resolve_status_category("задувши") is None


# -------------------------
# _category_name_matches
# -------------------------
def test_category_name_matches_single_name_without_pipe():
    assert mrh._category_name_matches("РТГр", "РТГР") is True
    assert mrh._category_name_matches("РТГр", "ІНШЕ") is False


def test_category_name_matches_any_pipe_separated_alternative():
    assert mrh._category_name_matches("Задув.|задув|задут", "ЗАДУВ.") is True
    assert mrh._category_name_matches("Задув.|задув|задут", "ЗАДУВ") is True
    assert mrh._category_name_matches("Задув.|задув|задут", "ЗАДУТ") is True
    assert mrh._category_name_matches("Задув.|задув|задут", "ЗАДУВАННЯ") is False


# -------------------------
# get_date_columns
# -------------------------
def test_get_date_columns_filters_and_sorts_datetimes():
    cols = ["ПІБ", datetime(2026, 7, 3), "ПОСАДА", datetime(2026, 7, 1)]
    assert mrh.get_date_columns(cols) == [datetime(2026, 7, 1), datetime(2026, 7, 3)]


# -------------------------
# resolve_day_value_and_category
# -------------------------
def test_resolve_day_value_and_category_plain_number_uses_base_status(monkeypatch):
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {30: {"general": []}, 100: {"general": []}})
    assert mrh.resolve_day_value_and_category(30, "ЗВРез") == (30, "ЗВРез")
    assert mrh.resolve_day_value_and_category(100, "ОБОРОНА") == (100, "ОБОРОНА")


def test_resolve_day_value_and_category_text_override_ignores_base_status(monkeypatch):
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {
        30: {"general": [], "ЗВРез": {}},
        100: {"general": [], "РТГр": {}},
    })
    assert mrh.resolve_day_value_and_category("РТГр", "ОБОРОНА") == (100, "РТГр")
    assert mrh.resolve_day_value_and_category("ЗВРез", "ОБОРОНА") == (30, "ЗВРез")


def test_resolve_day_value_and_category_leave_code_or_unknown_text_returns_none_none(monkeypatch):
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {30: {"general": []}, 100: {"general": []}})
    assert mrh.resolve_day_value_and_category("ВД", "ОБОРОНА") == (None, None)
    assert mrh.resolve_day_value_and_category(None, "ОБОРОНА") == (None, None)


def test_resolve_day_value_and_category_text_override_matches_regardless_of_case(monkeypatch):
    """Той самий регістр-незалежний збіг, що й у resolve_status_category, але для
    тексту-оверрайду в комірці дня - результат має повертати КАНОНІЧНЕ написання
    з MONEY_REPORT_CATEGORIES ("РТГр"), а не те, як його ввели в комірці ("ртгр")."""
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {
        30: {"general": []},
        100: {"general": [], "РТГр": {}},
    })
    assert mrh.resolve_day_value_and_category("ртгр", "ОБОРОНА") == (100, "РТГр")
    assert mrh.resolve_day_value_and_category("РТГР", "ОБОРОНА") == (100, "РТГр")


def test_resolve_day_value_and_category_matches_any_pipe_separated_alternative(monkeypatch):
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {
        "NOT_PAID": {"СЗЧ": {}, "Задув.|задув|задут": {}},
    })
    assert mrh.resolve_day_value_and_category("Задув.", None) == ("NOT_PAID", "Задув.|задув|задут")
    assert mrh.resolve_day_value_and_category("задув", None) == ("NOT_PAID", "Задув.|задув|задут")
    assert mrh.resolve_day_value_and_category("ЗАДУТ", None) == ("NOT_PAID", "Задув.|задув|задут")


# -------------------------
# get_tvo_commander_pibs
# -------------------------
def test_get_tvo_commander_pibs_includes_overlapping_tvo_record():
    rows_tvo = [{
        "ПОСАДА": "Командир батальйону", "ТВО": True,
        "Start": "01.07.2026", "End": "31.07.2026", "ПІБ": "Перший Перший",
    }]
    result = mrh.get_tvo_commander_pibs(rows_tvo, "Командир батальйону", datetime(2026, 7, 10), datetime(2026, 7, 20))
    assert result == {"ПЕРШИЙ ПЕРШИЙ"}


def test_get_tvo_commander_pibs_ignores_wrong_position_or_not_tvo():
    rows_tvo = [
        {"ПОСАДА": "Інша посада", "ТВО": True, "Start": "01.07.2026", "End": "31.07.2026", "ПІБ": "Перший Перший"},
        {"ПОСАДА": "Командир батальйону", "ТВО": False, "Start": "01.07.2026", "End": "31.07.2026", "ПІБ": "Другий Другий"},
    ]
    result = mrh.get_tvo_commander_pibs(rows_tvo, "Командир батальйону", datetime(2026, 7, 10), datetime(2026, 7, 20))
    assert result == set()


def test_get_tvo_commander_pibs_ignores_non_overlapping_period():
    rows_tvo = [{
        "ПОСАДА": "Командир батальйону", "ТВО": True,
        "Start": "01.01.2020", "End": "31.01.2020", "ПІБ": "Перший Перший",
    }]
    result = mrh.get_tvo_commander_pibs(rows_tvo, "Командир батальйону", datetime(2026, 7, 10), datetime(2026, 7, 20))
    assert result == set()


def test_get_tvo_commander_pibs_skips_rows_with_missing_dates():
    rows_tvo = [{"ПОСАДА": "Командир батальйону", "ТВО": True, "ПІБ": "Перший Перший"}]
    result = mrh.get_tvo_commander_pibs(rows_tvo, "Командир батальйону", datetime(2026, 7, 10), datetime(2026, 7, 20))
    assert result == set()


# -------------------------
# build_period_text_and_days
# -------------------------
def test_build_period_text_and_days_empty():
    assert mrh.build_period_text_and_days([]) == ("", 0)


def test_build_period_text_and_days_single_day():
    text, count = mrh.build_period_text_and_days([datetime(2026, 7, 1)])
    assert text == "01.07.2026-01.07.2026"
    assert count == 1


def test_build_period_text_and_days_contiguous_range():
    dates = [datetime(2026, 7, d) for d in (5, 6, 7)]
    text, count = mrh.build_period_text_and_days(dates)
    assert text == "05.07.2026-07.07.2026"
    assert count == 3


def test_build_period_text_and_days_multiple_ranges():
    dates = [datetime(2026, 7, d) for d in (1, 2, 5, 6, 6 + 1)]
    text, count = mrh.build_period_text_and_days(dates)
    assert text == "01.07.2026-02.07.2026; 05.07.2026-07.07.2026"
    assert count == 5


# -------------------------
# build_day_value_periods
# -------------------------
def test_build_day_value_periods_excludes_recognized_points_and_categories(monkeypatch):
    """Значення, що вже розпізнаються MONEY_REPORT_CATEGORIES (числа-точки 30/100
    чи назви категорій ЖИТТЄДІЯЛЬНІСТЬ/РТГр) НЕ повинні потрапляти в лог - вони й
    так є в самому рапорті; лог - лише для того, чого там немає."""
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {
        30: {"general": [], "ЖИТТЄДІЯЛЬНІСТЬ": {}},
        100: {"general": [], "РТГр": {}},
    })
    d1, d2, d3, d4 = (datetime(2026, 7, d) for d in (1, 2, 3, 4))
    row = {d1: 30, d2: 100, d3: "РТГр", d4: "ВД"}
    result = mrh.build_day_value_periods(row, [d1, d2, d3, d4])
    assert result == {"ВД": ("04.07.2026-04.07.2026", 1)}


def test_build_day_value_periods_groups_unrecognized_values_by_period(monkeypatch):
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {30: {"general": []}, 100: {"general": []}})
    d1, d2, d3 = (datetime(2026, 7, d) for d in (1, 2, 3))
    row = {d1: "ВД", d2: "ВД", d3: "ПРВД"}
    result = mrh.build_day_value_periods(row, [d1, d2, d3])
    assert result == {
        "ВД": ("01.07.2026-02.07.2026", 2),
        "ПРВД": ("03.07.2026-03.07.2026", 1),
    }


def test_build_day_value_periods_skips_empty_cells():
    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 2)
    row = {d1: None, d2: ""}
    assert mrh.build_day_value_periods(row, [d1, d2]) == {}


# -------------------------
# _lines_for_period
# -------------------------
def test_lines_for_period_includes_overlapping_and_excludes_others():
    references = [
        {"start": "01.05.2026", "end": "31.05.2026", "lines": ["травень"]},
        {"start": "01.06.2026", "end": None, "lines": ["червень і далі"]},
        {"start": "01.01.2020", "end": "31.01.2020", "lines": ["не має потрапити"]},
    ]
    lines = mrh._lines_for_period(references, datetime(2026, 6, 10), datetime(2026, 6, 15))
    assert lines == ["червень і далі"]


def test_lines_for_period_open_ended_end_uses_period_end():
    references = [{"start": "01.06.2026", "end": None, "lines": ["x"]}]
    # Період повністю ДО старту референсу - не має перетнутись.
    lines = mrh._lines_for_period(references, datetime(2026, 5, 1), datetime(2026, 5, 20))
    assert lines == []


# -------------------------
# build_brs_chain_lines / build_legal_basis_text
# -------------------------
@pytest.fixture(autouse=True)
def _no_general_money_references(monkeypatch):
    """Ізолює тести build_legal_basis_text/build_category_rows від реального
    constants.NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK, який продовжує наповнюватись реальними
    даними користувача (category_config-и в цих тестах самі визначають свій
    "general"/"grounds" напряму, тож MONEY_REPORT_CATEGORIES їх не стосується -
    лише 'бат'/'посилання_бат'/'посилання_брг' шукаються по реальних датах
    навіть із власним category_config, тому саме цю константу й ізолюємо)."""
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {})


EMPTY_CATEGORY = {"grounds": [], "use_brs": False, "general": []}


# -------------------------
# resolve_category_config
# -------------------------
def test_resolve_category_config_includes_all_general_by_default(monkeypatch):
    bn = {"id": "1", "start": "01.01.2026", "end": "31.01.2026", "lines": ["BN"]}
    jbd = {"id": "2", "start": "01.01.2026", "end": "31.01.2026", "lines": ["JBD"]}
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {
        30: {
            "general": [bn, jbd],
            "ТЕСТ": {"grounds": ["own"], "use_brs": True, "exclude_general": []},
        },
    })
    result = mrh.resolve_category_config(30, "ТЕСТ")
    assert result == {"grounds": ["own"], "use_brs": True, "use_brs_from_selected_folder": False, "general": [bn, jbd], "required_basis": True}


def test_resolve_category_config_excludes_by_id(monkeypatch):
    bn = {"id": "1", "start": "01.01.2026", "end": "31.01.2026", "lines": ["BN"]}
    jbd = {"id": "2", "start": "01.01.2026", "end": "31.01.2026", "lines": ["JBD"]}
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {
        100: {
            "general": [bn, jbd],
            "МЕДИК": {"grounds": [], "use_brs": True, "exclude_general": ["2"]},
        },
    })
    result = mrh.resolve_category_config(100, "МЕДИК")
    assert result["general"] == [bn]


def test_resolve_category_config_all_excludes_everything(monkeypatch):
    bn = {"id": "1", "start": "01.01.2026", "end": "31.01.2026", "lines": ["BN"]}
    jbd = {"id": "2", "start": "01.01.2026", "end": "31.01.2026", "lines": ["JBD"]}
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {
        100: {
            "general": [bn, jbd],
            "РТГр": {"grounds": [], "use_brs": False, "exclude_general": "all"},
        },
    })
    result = mrh.resolve_category_config(100, "РТГр")
    assert result["general"] == []


def test_resolve_category_config_defaults_when_fields_missing(monkeypatch):
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {30: {"general": [], "МІНІМУМ": {}}})
    result = mrh.resolve_category_config(30, "МІНІМУМ")
    assert result == {"grounds": [], "use_brs": False, "use_brs_from_selected_folder": False, "general": [], "required_basis": True}


def test_resolve_category_config_passes_through_required_basis_false(monkeypatch):
    """"required_basis": False (напр. 10к) переноситься в готовий category_config -
    саме за ним _log_missing_legal_basis вирішує, чи попереджати про порожню підставу."""
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {
        10: {"general": [], "ЖИТТЄДІЯЛЬНІСТЬ": {"grounds": [], "use_brs": False, "exclude_general": [], "required_basis": False}},
    })
    result = mrh.resolve_category_config(10, "ЖИТТЄДІЯЛЬНІСТЬ")
    assert result["required_basis"] is False


def test_resolve_category_config_passes_through_use_brs_from_selected_folder_true(monkeypatch):
    """"use_brs_from_selected_folder": True (напр. МЕДИК пункту 100) переноситься в готовий
    category_config - за ним build_legal_basis_text вирішує, з якого джерела (щоденні
    чи тижневі номери БР) будувати підставу."""
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {
        100: {"general": [], "МЕДИК": {"grounds": [], "use_brs_from_selected_folder": True, "exclude_general": []}},
    })
    result = mrh.resolve_category_config(100, "МЕДИК")
    assert result["use_brs_from_selected_folder"] is True


def test_resolve_category_config_uses_explicit_categories_instead_of_global(monkeypatch):
    """Явно переданий categories (напр. COMMANDER_MONEY_REPORT_CATEGORIES) повністю
    підміняє джерело - зміна глобального MONEY_REPORT_CATEGORIES його НЕ зачіпає."""
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {
        30: {"general": [], "ТЕСТ": {"grounds": ["з головного"], "use_brs": False, "exclude_general": []}},
    })
    own_categories = {
        30: {"general": [], "ТЕСТ": {"grounds": ["з окремого"], "use_brs": True, "exclude_general": []}},
    }
    result = mrh.resolve_category_config(30, "ТЕСТ", categories=own_categories)
    assert result == {"grounds": ["з окремого"], "use_brs": True, "use_brs_from_selected_folder": False, "general": [], "required_basis": True}


def test_build_legal_basis_text_empty_dates_returns_empty():
    assert mrh.build_legal_basis_text([], EMPTY_CATEGORY) == ""


def test_build_legal_basis_text_use_brs_false_ignores_brs_chain(monkeypatch):
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {"01.07.2026": {"бат": "999"}})
    result = mrh.build_legal_basis_text([datetime(2026, 7, 1)], EMPTY_CATEGORY)
    assert result == ""


def test_build_legal_basis_text_use_brs_true_builds_chain_with_priority_and_dedup(monkeypatch):
    fake_brs = {
        "01.07.2026": {"бат": "10", "посилання_бат": "67 від 20.06.2026", "посилання_брг": "100 від 01.06.2026"},
        "02.07.2026": {"бат": "10", "посилання_бат": "67 від 20.06.2026", "посилання_брг": "100 від 01.06.2026"},
        "03.07.2026": {"бат": "11", "посилання_бат": "", "посилання_брг": ""},
    }
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", fake_brs)

    dates = [datetime(2026, 7, 1), datetime(2026, 7, 2), datetime(2026, 7, 3)]
    category_config = {"grounds": [], "use_brs": True, "general": []}
    result = mrh.build_legal_basis_text(dates, category_config)
    lines = result.split("\n")

    # 'бат' дедуплікується за значенням, але дата в тексті - першого разу, коли зустрівся
    assert lines[0] == f"БР {constants.SHORT_UNIT_BATTALION} {constants.SHORT_UNIT_BRIGADE} №10 від 01.07.2026"
    assert lines[1] == f"БР {constants.SHORT_UNIT_BATTALION} {constants.SHORT_UNIT_BRIGADE} №11 від 03.07.2026"
    assert lines[2] == f"БР {constants.SHORT_UNIT_BATTALION} {constants.SHORT_UNIT_BRIGADE} №67 від 20.06.2026"
    assert lines[3] == f"БР {constants.SHORT_UNIT_BRIGADE} №100 від 01.06.2026"
    assert "бз_бат" not in result and "БР КБ" not in result


def test_build_legal_basis_text_dedups_posylannia_bat_against_earlier_bat_number(monkeypatch):
    """'посилання_бат' зберігається як текст "NNN від ДАТА" - якщо цей САМИЙ номер
    NNN вже друкувався раніше через власне поле 'бат' іншої дати (типова ситуація:
    наказ від дня X пізніше згадується як "посилання" в записі дня Y), він НЕ
    повинен друкуватись іще раз - інакше один і той самий наказ дублюється двічі."""
    fake_brs = {
        "01.07.2026": {"бат": "87", "посилання_бат": "67 від 20.06.2026", "посилання_брг": ""},
        "07.07.2026": {"бат": "102", "посилання_бат": "87 від 01.07.2026", "посилання_брг": ""},
    }
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", fake_brs)

    dates = [datetime(2026, 7, 1), datetime(2026, 7, 7)]
    category_config = {"grounds": [], "use_brs": True, "general": []}
    result = mrh.build_legal_basis_text(dates, category_config)
    lines = result.split("\n")

    assert lines == [
        f"БР {constants.SHORT_UNIT_BATTALION} {constants.SHORT_UNIT_BRIGADE} №87 від 01.07.2026",
        f"БР {constants.SHORT_UNIT_BATTALION} {constants.SHORT_UNIT_BRIGADE} №102 від 07.07.2026",
        f"БР {constants.SHORT_UNIT_BATTALION} {constants.SHORT_UNIT_BRIGADE} №67 від 20.06.2026",
    ]


def test_build_legal_basis_text_general_references_applied_in_listed_order():
    category_config = {
        "grounds": [],
        "use_brs": False,
        "general": [
            {"start": "01.07.2026", "end": "31.07.2026", "lines": ["БН ТЕСТ"]},
            {"start": "01.07.2026", "end": "31.07.2026", "lines": ["ЖБД ТЕСТ"]},
        ],
    }
    result = mrh.build_legal_basis_text([datetime(2026, 7, 1)], category_config)
    assert result.split("\n") == ["БН ТЕСТ", "ЖБД ТЕСТ"]


def test_build_legal_basis_text_general_only_lists_what_category_includes():
    bn = {"start": "01.07.2026", "end": "31.07.2026", "lines": ["БН ТЕСТ"]}
    # Категорія просто НЕ включає JBD до свого "general" (resolve_category_config
    # вже виключив його через exclude_general) - тут це вже готовий плаский список.
    category_config = {"grounds": [], "use_brs": False, "general": [bn]}
    result = mrh.build_legal_basis_text([datetime(2026, 7, 1)], category_config)
    assert result == "БН ТЕСТ"


def test_build_legal_basis_text_grounds_appended_after_general():
    category_config = {
        "grounds": [{"start": "01.07.2026", "end": "31.07.2026", "lines": ["ВЛАСНА ПІДСТАВА КАТЕГОРІЇ"]}],
        "use_brs": False,
        "general": [],
    }
    result = mrh.build_legal_basis_text([datetime(2026, 7, 1)], category_config)
    assert result == "ВЛАСНА ПІДСТАВА КАТЕГОРІЇ"


def test_build_legal_basis_text_appends_extra_lines_after_grounds():
    category_config = {
        "grounds": [{"start": "01.07.2026", "end": "31.07.2026", "lines": ["ЗАГАЛЬНА ПІДСТАВА"]}],
        "use_brs": False,
        "general": [],
    }
    result = mrh.build_legal_basis_text([datetime(2026, 7, 1)], category_config, extra_lines=["ОСОБИСТА ПІДСТАВА"])
    assert result == "ЗАГАЛЬНА ПІДСТАВА\nОСОБИСТА ПІДСТАВА"


def test_build_legal_basis_text_extra_lines_none_or_empty_has_no_effect():
    category_config = {"grounds": [], "use_brs": False, "general": []}
    assert mrh.build_legal_basis_text([datetime(2026, 7, 1)], category_config, extra_lines=None) == ""
    assert mrh.build_legal_basis_text([datetime(2026, 7, 1)], category_config, extra_lines=[]) == ""


def test_build_legal_basis_text_dedups_duplicate_lines_within_grounds():
    """Той самий рядок, випадково продубльований у "lines" ОДНОГО запису grounds
    (напр. рукописна помилка при заповненні constants.py) - має з'явитись у
    підставі лише ОДИН раз, зі збереженням позиції першої появи."""
    category_config = {
        "grounds": [{"start": "01.07.2026", "end": "31.07.2026", "lines": ["А", "Б", "Б", "В"]}],
        "use_brs": False,
        "general": [],
    }
    result = mrh.build_legal_basis_text([datetime(2026, 7, 1)], category_config)
    assert result == "А\nБ\nВ"


def test_build_legal_basis_text_dedups_duplicate_lines_across_general_and_grounds():
    """Той самий текст, що трапляється і в "general", і у власних "grounds"
    (напр. два записи з різних довідників, що перетинаються за періодом) -
    так само не повинен дублюватись у підсумковій підставі."""
    category_config = {
        "grounds": [{"start": "01.07.2026", "end": "31.07.2026", "lines": ["СПІЛЬНИЙ РЯДОК", "ВЛАСНЕ"]}],
        "use_brs": False,
        "general": [{"start": "01.07.2026", "end": "31.07.2026", "lines": ["СПІЛЬНИЙ РЯДОК"]}],
    }
    result = mrh.build_legal_basis_text([datetime(2026, 7, 1)], category_config)
    assert result == "СПІЛЬНИЙ РЯДОК\nВЛАСНЕ"


def test_build_legal_basis_text_full_priority_order(monkeypatch):
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {"01.07.2026": {"бат": "10"}})
    category_config = {
        "grounds": [{"start": "01.07.2026", "end": "31.07.2026", "lines": ["ВЛАСНА"]}],
        "use_brs": True,
        "general": [
            {"start": "01.07.2026", "end": "31.07.2026", "lines": ["БН"]},
            {"start": "01.07.2026", "end": "31.07.2026", "lines": ["ЖБД"]},
        ],
    }
    result = mrh.build_legal_basis_text([datetime(2026, 7, 1)], category_config)
    assert result.split("\n") == [
        f"БР {constants.SHORT_UNIT_BATTALION} {constants.SHORT_UNIT_BRIGADE} №10 від 01.07.2026",
        "БН", "ЖБД", "ВЛАСНА",
    ]


def test_build_legal_basis_text_use_brs_from_selected_folder_false_ignores_daily_chain(monkeypatch):
    """"use_brs_from_selected_folder" не встановлено (чи False) - навіть якщо
    NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY заповнений, підстава лишається порожньою."""
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", {"01.07.2026": {"бат": "999"}})
    result = mrh.build_legal_basis_text([datetime(2026, 7, 1)], EMPTY_CATEGORY)
    assert result == ""


def test_build_legal_basis_text_use_brs_from_selected_folder_true_builds_chain_with_dedup(monkeypatch):
    fake_brs_every_day = {
        "01.07.2026": {"бат": "10", "брг": "5"},
        "02.07.2026": {"бат": "10", "брг": "5"},
        "03.07.2026": {"бат": "11", "брг": ""},
    }
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", fake_brs_every_day)

    dates = [datetime(2026, 7, 1), datetime(2026, 7, 2), datetime(2026, 7, 3)]
    category_config = {"grounds": [], "use_brs_from_selected_folder": True, "general": []}
    result = mrh.build_legal_basis_text(dates, category_config)
    lines = result.split("\n")

    assert lines == [
        f"БР {constants.SHORT_UNIT_BATTALION} {constants.SHORT_UNIT_BRIGADE} №10 від 01.07.2026",
        f"БР {constants.SHORT_UNIT_BATTALION} {constants.SHORT_UNIT_BRIGADE} №11 від 03.07.2026",
        f"БР {constants.SHORT_UNIT_BRIGADE} №5 від 01.07.2026",
    ]


def test_build_legal_basis_text_use_brs_from_selected_folder_returns_empty_when_folder_not_read(monkeypatch):
    """Якщо папку з документами не зчитували цього запуску (відповідь "Ні" на
    "Зчитати обрану папку з документами (для номерів БАТ / БЗ)?") -
    NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY лишається порожнім для відповідних дат, і
    підстава просто виходить порожньою (без падіння) - саме цей сценарій пізніше
    ловить _log_missing_legal_basis у generate_report_for_get_money.py."""
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", {})
    category_config = {"grounds": [], "use_brs_from_selected_folder": True, "general": []}
    result = mrh.build_legal_basis_text([datetime(2026, 7, 1)], category_config)
    assert result == ""


def test_build_legal_basis_text_combines_use_brs_and_use_brs_from_selected_folder(monkeypatch):
    """Обидва прапорці можуть бути True одночасно - тижневий і щоденний ланцюжки
    БР просто йдуть один за одним (спершу тижневий, потім щоденний)."""
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {"01.07.2026": {"бат": "10"}})
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", {"01.07.2026": {"бат": "20"}})
    category_config = {"grounds": [], "use_brs": True, "use_brs_from_selected_folder": True, "general": []}
    result = mrh.build_legal_basis_text([datetime(2026, 7, 1)], category_config)
    assert result.split("\n") == [
        f"БР {constants.SHORT_UNIT_BATTALION} {constants.SHORT_UNIT_BRIGADE} №10 від 01.07.2026",
        f"БР {constants.SHORT_UNIT_BATTALION} {constants.SHORT_UNIT_BRIGADE} №20 від 01.07.2026",
    ]


# -------------------------
# build_category_rows
# -------------------------
def _row(pib, posada, pidrozdil, days, zvannya="сержант"):
    row = {"ПІБ": pib, "ПОСАДА": posada, "ПІДРОЗДІЛ": pidrozdil, "ЗВАННЯ": zvannya}
    row.update(days)
    return row


def test_build_category_rows_no_longer_hardcodes_battalion_commander_exclusion():
    """Штатний командир батальйону більше НЕ виключається жорстко всередині
    build_category_rows - хто саме потрапляє в rows_with_data (з ним чи без),
    вирішує викликач (generate_report_for_get_money.py прибирає його зі свого
    списку сам, а generate_report_for_commander_money.py - навпаки, будує список
    САМЕ з нього, для окремого рапорту від першої особи)."""
    d1 = datetime(2026, 7, 1)
    rows = [_row("Третій Третій", "Командир батальйону", "упр", {d1: 30})]
    result = mrh.build_category_rows(rows, [d1], {30}, {}, status_filter=None, category_config=EMPTY_CATEGORY)
    assert [r["ПІБ"] for r in result] == ["Третій Третій"]


def test_build_category_rows_skips_person_with_no_matching_dates():
    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: "ВД"})]
    result = mrh.build_category_rows(rows, [d1], {30, 100}, {}, status_filter=None, category_config=EMPTY_CATEGORY)
    assert result == []


def test_build_category_rows_basic_thirty_section():
    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 2)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30, d2: 30})]
    result = mrh.build_category_rows(rows, [d1, d2], {30}, {}, status_filter=None, category_config=EMPTY_CATEGORY)
    assert len(result) == 1
    row = result[0]
    assert row["ПОСАДА"] == "Стрілець"
    assert row["ЗВАННЯ"] == "сержант"
    assert row["ПІБ"] == "Другий Другий"
    assert row["ПЕРІОД"] == "01.07.2026-02.07.2026"
    assert row["ДНІ"] == 2


def test_build_category_rows_status_filter_medic():
    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший", "Бойовий медик", "Підрозділ 1", {d1: 100}),
        _row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 100}),
    ]
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "МЕДИК"}
    result = mrh.build_category_rows(rows, [d1], {100}, status_lookup, status_filter="МЕДИК", category_config=EMPTY_CATEGORY)
    assert [r["ПІБ"] for r in result] == ["Перший Перший"]


def test_build_category_rows_status_filter_accepts_collection():
    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший", "Бойовий медик", "Підрозділ 1", {d1: 30}),
        _row("Третій Третій", "Стрілець", "Підрозділ 1", {d1: 30}),
        _row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30}),
    ]
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "МЕДИК", "ТРЕТІЙ ТРЕТІЙ": "БД(СЗ)"}
    result = mrh.build_category_rows(
        rows, [d1], {30}, status_lookup, status_filter={"МЕДИК", "БД(СЗ)"}, category_config=EMPTY_CATEGORY,
    )
    assert {r["ПІБ"] for r in result} == {"Перший Перший", "Третій Третій"}


def test_build_category_rows_mp_subdivision_forced_medic_even_with_other_status():
    d1 = datetime(2026, 7, 1)
    rows = [_row("Перший Перший", "Начальник медичного пункту", "МП", {d1: 30})]
    status_lookup = {"ЛІКАР ОЛЬГА": "ОБОРОНА"}
    result = mrh.build_category_rows(rows, [d1], {30, 100}, status_lookup, status_filter="МЕДИК", category_config=EMPTY_CATEGORY)
    assert len(result) == 1
    assert result[0]["ПІБ"] == "Перший Перший"


def test_build_category_rows_exclude_status_skips_medics():
    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший", "Бойовий медик", "Підрозділ 1", {d1: 30}),
        _row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30}),
    ]
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "МЕДИК"}
    result = mrh.build_category_rows(
        rows, [d1], {30}, status_lookup, status_filter=None, category_config=EMPTY_CATEGORY, exclude_status="МЕДИК",
    )
    assert [r["ПІБ"] for r in result] == ["Другий Другий"]


def test_build_category_rows_exclude_status_accepts_collection():
    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший", "Бойовий медик", "Підрозділ 1", {d1: 30}),
        _row("Третій Третій", "Стрілець", "Підрозділ 1", {d1: 30}),
        _row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30}),
    ]
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "МЕДИК", "ТРЕТІЙ ТРЕТІЙ": "БД(СЗ)"}
    result = mrh.build_category_rows(
        rows, [d1], {30}, status_lookup, status_filter=None, category_config=EMPTY_CATEGORY,
        exclude_status={"МЕДИК", "БД(СЗ)"},
    )
    assert [r["ПІБ"] for r in result] == ["Другий Другий"]


def test_build_category_rows_preserves_oblik_file_order_not_subdivision_order():
    """Порядок рядків у рапорті має точно збігатись з порядком в ОБЛІК.xlsx - без
    жодного перевпорядкування за підрозділом чи будь-чим іншим."""
    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший", "Стрілець", "Підрозділ 2", {d1: 30}),
        _row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30}),
    ]
    result = mrh.build_category_rows(rows, [d1], {30}, {}, status_filter=None, category_config=EMPTY_CATEGORY)
    assert [r["ПІБ"] for r in result] == ["Перший Перший", "Другий Другий"]


def test_build_category_rows_uses_category_config_for_legal_basis():
    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30})]
    category_config = {
        "grounds": [{"start": "01.07.2026", "end": "31.07.2026", "lines": ["ПІДСТАВА КАТЕГОРІЇ"]}],
        "use_brs": False,
        "general": [],
    }
    result = mrh.build_category_rows(rows, [d1], {30}, {}, status_filter=None, category_config=category_config)
    assert result[0]["ПІДСТАВА"] == "ПІДСТАВА КАТЕГОРІЇ"


def test_build_category_rows_applies_extra_grounds_only_to_matching_person():
    d1 = datetime(2026, 7, 1)
    rows = [
        _row("ОДИНАДЦЯТИЙ Одинадцятий Одинадцятий", "Стрілець", "Підрозділ 1", {d1: 30}),
        _row("ДВАНАДЦЯТИЙ Дванадцятий Дванадцятий", "Стрілець", "Підрозділ 1", {d1: 30}),
    ]
    category_config = {"grounds": [], "use_brs": False, "general": []}
    extra_line = f"БР {constants.SHORT_UNIT_BATTALION} {constants.SHORT_UNIT_BRIGADE} №79 від 05.06.2026"
    extra_grounds_by_person = {"ОДИНАДЦЯТИЙ ОДИНАДЦЯТИЙ ОДИНАДЦЯТИЙ": [extra_line]}

    result = mrh.build_category_rows(
        rows, [d1], {30}, {}, status_filter=None, category_config=category_config,
        extra_grounds_by_person=extra_grounds_by_person,
    )

    by_pib = {row["ПІБ"]: row["ПІДСТАВА"] for row in result}
    assert by_pib["ОДИНАДЦЯТИЙ Одинадцятий Одинадцятий"] == extra_line
    assert by_pib["ДВАНАДЦЯТИЙ Дванадцятий Дванадцятий"] == ""


def test_build_category_rows_threads_required_basis_flag_into_row():
    """category_config["required_basis"] потрапляє в кожен побудований рядок
    (внутрішній прапорець "_ПІДСТАВА_ОБОВ'ЯЗКОВА") - саме за ним
    _log_missing_legal_basis вирішує, чи попереджати про порожню підставу."""
    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 10})]
    category_config = {"grounds": [], "use_brs": False, "general": [], "required_basis": False}

    result = mrh.build_category_rows(rows, [d1], {10}, {}, status_filter=None, category_config=category_config)

    assert result[0]["ПІДСТАВА"] == ""
    assert result[0]["_ПІДСТАВА_ОБОВ'ЯЗКОВА"] is False


def test_build_category_rows_defaults_required_basis_to_true():
    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30})]

    result = mrh.build_category_rows(rows, [d1], {30}, {}, status_filter=None, category_config=EMPTY_CATEGORY)

    assert result[0]["_ПІДСТАВА_ОБОВ'ЯЗКОВА"] is True


def test_build_category_rows_threads_use_brs_from_selected_folder_flag_and_category_name_into_row(monkeypatch):
    """category_config["use_brs_from_selected_folder"] і category_name потрапляють у кожен
    побудований рядок (внутрішні прапорці "_ВИКОРИСТОВУЄ_ОБРАНУ_ПАПКУ"/"_КАТЕГОРІЯ") -
    саме за ними _log_missing_legal_basis вирішує, чи показувати спеціальне
    повідомлення "завдання не вдалось зчитати" замість загального попередження."""
    monkeypatch.setattr(mrh, "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", {})
    d1 = datetime(2026, 7, 1)
    rows = [_row("Перший Перший", "Бойовий медик", "Підрозділ 1", {d1: 100})]
    category_config = {"grounds": [], "use_brs_from_selected_folder": True, "general": []}

    result = mrh.build_category_rows(
        rows, [d1], {100}, {}, status_filter=None, category_config=category_config, category_name="МЕДИК",
    )

    assert result[0]["ПІДСТАВА"] == ""
    assert result[0]["_ВИКОРИСТОВУЄ_ОБРАНУ_ПАПКУ"] is True
    assert result[0]["_КАТЕГОРІЯ"] == "МЕДИК"


def test_build_category_rows_defaults_use_brs_from_selected_folder_and_category_name():
    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30})]

    result = mrh.build_category_rows(rows, [d1], {30}, {}, status_filter=None, category_config=EMPTY_CATEGORY)

    assert result[0]["_ВИКОРИСТОВУЄ_ОБРАНУ_ПАПКУ"] is False
    assert result[0]["_КАТЕГОРІЯ"] is None


def test_build_category_rows_day_text_override_ignores_base_status_for_that_day(monkeypatch):
    """Людина базово (по колонці СТАТУС) - "ОБОРОНА", але один конкретний день у
    ОБЛІК.xlsx позначено текстом "РТГр" замість числа - цей день має потрапити в
    РТГр, попри те, що на решту місяця людина рахується як ОБОРОНА."""
    monkeypatch.setattr(mrh, "MONEY_REPORT_CATEGORIES", {
        100: {"general": [], "РТГр": {}, "ОБОРОНА": {}},
    })
    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 2)
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 100, d2: "РТГр"})]
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "ОБОРОНА"}

    rtgr_rows = mrh.build_category_rows(
        rows, [d1, d2], {100}, status_lookup, status_filter="РТГр", category_config=EMPTY_CATEGORY,
    )
    oborona_rows = mrh.build_category_rows(
        rows, [d1, d2], {100}, status_lookup, status_filter="ОБОРОНА", category_config=EMPTY_CATEGORY,
    )

    assert [r["ДНІ"] for r in rtgr_rows] == [1]
    assert rtgr_rows[0]["ПЕРІОД"] == "02.07.2026-02.07.2026"
    assert [r["ДНІ"] for r in oborona_rows] == [1]
    assert oborona_rows[0]["ПЕРІОД"] == "01.07.2026-01.07.2026"


# -------------------------
# Наскрізна генерація рапорту (реальні resources/*, ізольована output-директорія)
# -------------------------
def test_process_generate_report_for_get_money_returns_none_without_dates():
    from process_generate_report_for_get_additional_money import process_generate_report_for_get_money

    result = process_generate_report_for_get_money([], [], ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ"])
    assert result is None


# -------------------------
# _center_basis_column
# -------------------------
def _make_table_for(report_module, rows):
    from docx import Document
    from formatting.docx_utils import create_table

    doc = Document()
    data_rows = [
        [str(i + 1), r["ПОСАДА"], r["ЗВАННЯ"], r["ПІБ"], r["ПЕРІОД"], str(r["ДНІ"]), r["ПІДСТАВА"]]
        for i, r in enumerate(rows)
    ]
    create_table(doc, headers=report_module.TABLE_HEADERS, data_rows=data_rows, widths_cm=report_module.TABLE_WIDTHS_CM, font_size=report_module._TABLE_FONT_SIZE_PT)
    return doc, doc.tables[-1]


def test_center_basis_column_sets_center_alignment_only_for_that_column():
    """"Підстава для виплати" - по центру комірки (горизонтально; вертикально
    центрується окремо через create_table(cell_alignment=CENTER))."""
    import generators.generate_report_for_get_money as report_module
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", "ПЕРІОД": "01.07.2026-01.07.2026", "ДНІ": 1, "ПІДСТАВА": "Рядок 1\nРядок 2"}]
    doc, table = _make_table_for(report_module, rows)

    report_module._center_basis_column(table)

    for paragraph in table.cell(0, report_module.BASIS_COLUMN_INDEX).paragraphs:
        assert paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
    for paragraph in table.cell(1, report_module.BASIS_COLUMN_INDEX).paragraphs:
        assert paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
    # Інші колонки не зачіпаються.
    assert table.cell(1, 3).paragraphs[0].alignment != WD_ALIGN_PARAGRAPH.CENTER


# -------------------------
# _place_rows_for_page_fill
# -------------------------
def test_place_rows_for_page_fill_pulls_smaller_row_forward_to_fill_gap():
    """Якщо наступний за порядком рядок не влазить у залишок сторінки, а десь
    далі є менший, що влазить - саме він підставляється замість нього, щоб
    сторінка не лишалась заповненою лише частково."""
    import generators.generate_report_for_get_money as report_module

    small_row = {"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "МАЛИЙ", "ПЕРІОД": "01.07.2026-01.07.2026", "ДНІ": 1, "ПІДСТАВА": "X"}
    big_row = {"ПОСАДА": "П" * 200, "ЗВАННЯ": "с", "ПІБ": "ВЕЛИКИЙ", "ПЕРІОД": "01.07.2026-01.07.2026", "ДНІ": 1, "ПІДСТАВА": ""}
    rows = [big_row, small_row]

    font_pt = 10
    usable_height_cm = report_module._row_height_cm(small_row, font_pt) + 0.01
    placements, final_page_used_cm = report_module._place_rows_for_page_fill(rows, usable_height_cm, initial_used_cm=0.0, font_size_pt=font_pt)

    # Великий рядок (не влазить у залишок цієї крихітної "сторінки") відсунуто,
    # малий - підтягнуто вперед; фінальна позиція рахується за базовою висотою -
    # кегль ніколи не змінюється.
    assert [row["ПІБ"] for row in placements] == ["МАЛИЙ", "ВЕЛИКИЙ"]
    assert final_page_used_cm == report_module._row_height_cm(big_row, font_pt)


def test_place_rows_for_page_fill_never_shrinks_font_even_when_nothing_fits():
    """Кегль ЗАВЖДИ лишається базовим - підтверджено користувачем: у табличках
    рапорту й "змін в наказі" шрифт має бути рівно 10, без винятків. Раніше тут
    стискався шрифт, коли всі рядки, що лишились, приблизно однаково великі
    (типово - лише командний склад із великою підставою в кожного); тепер
    рядок, що не влазить у залишок поточної сторінки, просто переходить на
    нову - кожен рядок тут влазить на ОКРЕМУ чисту сторінку за базовим кеглем."""
    import generators.generate_report_for_get_money as report_module

    def r(pib, tag):
        return {"ПОСАДА": "Командир", "ЗВАННЯ": "капітан", "ПІБ": pib, "ПЕРІОД": "01.07.2026-31.07.2026", "ДНІ": 31, "ПІДСТАВА": "\n".join(f"Наказ {tag}№{i}" for i in range(30))}

    rows = [r("ОДИН", "А"), r("ДВА", "Б")]
    font_pt = 10
    first_height = report_module._row_height_cm(rows[0], font_pt)
    initial_used_cm = 2.0
    usable_height_cm = first_height + 0.5

    placements, final_page_used_cm = report_module._place_rows_for_page_fill(rows, usable_height_cm, initial_used_cm, font_size_pt=font_pt)

    # Обидва рядки - за базовим кеглем, кожен на своїй (щойно почато чистій)
    # сторінці; природний порядок збережено.
    assert [row["ПІБ"] for row in placements] == ["ОДИН", "ДВА"]
    assert final_page_used_cm == first_height


def test_place_rows_for_page_fill_places_row_too_big_for_any_page_as_is():
    """Рядок, що не влазить навіть на щойно почату чисту сторінку (напр. дуже
    об'ємна підстава в однієї людини) - розміщується як є (Word сам розіб'є
    таку таблицю на кілька сторінок природним чином), кегль не чіпається."""
    import generators.generate_report_for_get_money as report_module

    huge_row = {
        "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "ІВАН", "ПЕРІОД": "01.07.2026-01.07.2026", "ДНІ": 1,
        "ПІДСТАВА": "\n".join(f"Наказ №{i}" for i in range(200)),
    }
    font_pt = 10
    usable_height_cm = 23.0  # реалістичний розмір сторінки - рядок все одно більший.
    assert report_module._row_height_cm(huge_row, font_pt) > usable_height_cm

    placements, final_page_used_cm = report_module._place_rows_for_page_fill([huge_row], usable_height_cm, initial_used_cm=5.0, font_size_pt=font_pt)

    assert placements == [huge_row]
    assert final_page_used_cm == report_module._row_height_cm(huge_row, font_pt)


def test_place_rows_for_page_fill_preserves_natural_order_when_everything_fits():
    import generators.generate_report_for_get_money as report_module

    def r(pib):
        return {"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": pib, "ПЕРІОД": "01.07.2026-01.07.2026", "ДНІ": 1, "ПІДСТАВА": ""}

    rows = [r("ОДИН"), r("ДВА"), r("ТРИ")]
    placements, final_page_used_cm = report_module._place_rows_for_page_fill(rows, usable_height_cm=1000.0, initial_used_cm=0.0, font_size_pt=10)

    assert [row["ПІБ"] for row in placements] == ["ОДИН", "ДВА", "ТРИ"]
    assert final_page_used_cm > 0


# -------------------------
# _usable_text_width_cm / _estimate_paragraph_height_cm / carried_over_cm threading
# -------------------------
def test_estimate_paragraph_height_cm_grows_with_text_length():
    import generators.generate_report_for_get_money as report_module

    short_h = report_module._estimate_paragraph_height_cm("Короткий текст.", 14, usable_width_cm=17.0)
    long_h = report_module._estimate_paragraph_height_cm("Дуже " * 200 + "довгий текст.", 14, usable_width_cm=17.0)
    assert long_h > short_h


def test_add_category_section_carries_page_position_into_return_value(monkeypatch):
    """carried_over_cm, передане на вході, впливає на те, скільки саме
    сторінки лишиться зайнятим на виході - підтверджує, що стан коректно
    протікає через функцію (а не ігнорується)."""
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    monkeypatch.setitem(report_module.STATIK["SECTIONS"], "TEST_TEMPLATE", "Текст пункту {start_date_day} {manth} {year}:")

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", "ПЕРІОД": "01.07.2026-01.07.2026", "ДНІ": 1, "ПІДСТАВА": "X"}]

    result_from_zero = report_module._add_category_section(doc, 1, "TEST_TEMPLATE", rows, 7, 1, 31, carried_over_cm=0.0)

    doc2 = Document()
    report_module.set_margins(doc2, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    result_from_ten = report_module._add_category_section(doc2, 1, "TEST_TEMPLATE", rows, 7, 1, 31, carried_over_cm=10.0)

    assert result_from_ten > result_from_zero


def test_add_category_section_table_spans_full_usable_page_width(monkeypatch):
    """TABLE_WIDTHS_CM задає лише СПІВВІДНОШЕННЯ колонок - фактична таблиця має
    рівно заповнювати ширину сторінки між полями, а не власну (наближену) суму."""
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches, Cm

    monkeypatch.setitem(report_module.STATIK["SECTIONS"], "TEST_TEMPLATE", "Текст пункту {start_date_day} {manth} {year}:")

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", "ПЕРІОД": "01.07.2026-01.07.2026", "ДНІ": 1, "ПІДСТАВА": "X"}]

    report_module._add_category_section(doc, 1, "TEST_TEMPLATE", rows, 7, 1, 31, carried_over_cm=0.0)

    table = doc.tables[-1]
    total_width_cm = sum(col.width.cm for col in table.columns)
    usable_width_cm = report_module._usable_text_width_cm(doc)
    assert total_width_cm == pytest.approx(usable_width_cm, abs=0.01)


def test_add_category_section_heading_uses_configured_font_size(monkeypatch):
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches, Pt

    monkeypatch.setitem(report_module.STATIK["SECTIONS"], "TEST_TEMPLATE", "Текст пункту {start_date_day} {manth} {year}:")

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", "ПЕРІОД": "01.07.2026-01.07.2026", "ДНІ": 1, "ПІДСТАВА": "X"}]

    report_module._add_category_section(doc, 1, "TEST_TEMPLATE", rows, 7, 1, 31, carried_over_cm=0.0)

    heading_paragraph = doc.paragraphs[0]
    assert heading_paragraph.runs[0].font.size == Pt(report_module._HEADING_FONT_SIZE_PT)
    assert report_module._HEADING_FONT_SIZE_PT == 13.5


# -------------------------
# _bold_money_amounts
# -------------------------
def test_bold_money_amounts_finds_amount_regardless_of_point_number():
    """Регулярний вираз, а не хардкод по номеру точки - нова точка (напр. 50)
    отримує виділення суми жирним без жодних змін коду."""
    import generators.generate_report_for_get_money as report_module

    text = "Виплатити додаткову винагороду у розмірі 50 000 грн. 00 коп. військовослужбовцям..."
    assert report_module._bold_money_amounts(text) == ["50 000 грн. 00 коп."]


def test_bold_money_amounts_no_match_returns_empty_list():
    import generators.generate_report_for_get_money as report_module

    assert report_module._bold_money_amounts("Виключити пункти в Додатку 2:") == []


def test_add_category_section_heading_bolds_money_amount(monkeypatch):
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    monkeypatch.setitem(
        report_module.STATIK["SECTIONS"], "TEST_TEMPLATE",
        "Виплатити у розмірі 50 000 грн. 00 коп. за період з {start_date_day} {manth} {year}:",
    )
    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", "ПЕРІОД": "01.07.2026-01.07.2026", "ДНІ": 1, "ПІДСТАВА": "X"}]

    report_module._add_category_section(doc, 1, "TEST_TEMPLATE", rows, 7, 1, 31, carried_over_cm=0.0)

    heading_paragraph = doc.paragraphs[0]
    bold_runs = [r for r in heading_paragraph.runs if r.font.bold]
    assert any(r.text == "50 000 грн. 00 коп." for r in bold_runs)
    plain_runs = [r for r in heading_paragraph.runs if not r.font.bold]
    assert any("Виплатити у розмірі" in r.text for r in plain_runs)


# -------------------------
# _add_changes_intro / _add_changes_subsection / _add_changes_sections
# -------------------------
def test_add_changes_intro_leaves_blank_when_no_order_reference():
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))

    result = report_module._add_changes_intro(doc, 11, None, carried_over_cm=0.0)

    expected_text = report_module.STATIK["CHANGES_INTRO"].format(order_reference=report_module._BLANK_ORDER_REFERENCE)
    assert doc.paragraphs[0].text == f"11. {expected_text}"
    assert result > 0.0
    assert len(doc.tables) == 0


def test_add_changes_intro_substitutes_order_reference_when_given():
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))

    report_module._add_changes_intro(doc, 11, "№45 від 05.06.2026", carried_over_cm=0.0)

    expected_text = report_module.STATIK["CHANGES_INTRO"].format(order_reference="№45 від 05.06.2026")
    assert doc.paragraphs[0].text == f"11. {expected_text}"


def test_add_changes_subsection_formats_two_level_heading_with_appendix_number():
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", "ПЕРІОД": "01.06.2026-30.06.2026", "ДНІ": 30, "ПІДСТАВА": "X"}]

    report_module._add_changes_subsection(doc, 11, 1, "CHANGES_EXCLUDE_LABEL", 2, rows, carried_over_cm=0.0)

    assert doc.paragraphs[0].text == "11.1 Виключити пункти в Додатку 2:"
    assert len(doc.tables) == 1


def test_add_changes_sections_numbers_points_sequentially_and_advances_point_number():
    """Кожен запис changes_entries отримує СВІЙ ЦІЛИЙ номер пункту (11, 12, ...);
    підномери (N.1, N.2, ...) скидаються на 1 для КОЖНОГО запису окремо.
    "Виключити" НІКОЛИ не рендериться, навіть якщо exclude_rows непорожній
    (перший запис) - підтверджено користувачем."""
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    exclude_rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", "ПЕРІОД": "01.06.2026-30.06.2026", "ДНІ": 30, "ПІДСТАВА": "X"}]
    add_rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Перший Перший", "ПЕРІОД": "01.06.2026-15.06.2026", "ДНІ": 15, "ПІДСТАВА": "Y"}]
    changes_entries = [
        {"appendix_pairs": [(10, 1, exclude_rows, add_rows)], "order_reference": "№45 від 05.06.2026"},
        {"appendix_pairs": [(100, 1, [], add_rows)]},
    ]

    next_point_number, carried_over_cm = report_module._add_changes_sections(doc, 11, changes_entries, carried_over_cm=0.0)

    assert next_point_number == 13
    assert carried_over_cm > 0.0
    heading_texts = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
    first_intro = report_module.STATIK["CHANGES_INTRO"].format(order_reference="№45 від 05.06.2026")
    second_intro = report_module.STATIK["CHANGES_INTRO"].format(order_reference=report_module._BLANK_ORDER_REFERENCE)
    assert heading_texts[0] == f"11. {first_intro}"
    assert heading_texts[1] == "11.1 Додаток 1 доповнити наступними пунктами:"
    assert heading_texts[2] == f"12. {second_intro}"
    assert heading_texts[3] == "12.1 Додаток 1 доповнити наступними пунктами:"


def test_add_changes_sections_restate_point_uses_restate_label_not_add():
    """RESTATE_POINTS (report_changes.py, зараз - лише 30) - рядок із
    "_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ": True рендериться під "CHANGES_RESTATE_LABEL"
    ("Викласти в новій редакції..."), а НЕ "CHANGES_ADD_LABEL" ("...доповнити
    наступними пунктами") - на відміну від звичайних пунктів (напр. 100
    нижче, де цей прапорець узагалі не перевіряється)."""
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    restate_rows = [{
        "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Перший Перший",
        "ПЕРІОД": "02.06.2026-02.06.2026", "ДНІ": 1, "ПІДСТАВА": "", "_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ": True,
    }]
    add_rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Перший Перший", "ПЕРІОД": "02.06.2026-02.06.2026", "ДНІ": 1, "ПІДСТАВА": ""}]
    changes_entries = [{"appendix_pairs": [(30, 1, [], restate_rows), (100, 2, [], add_rows)]}]

    report_module._add_changes_sections(doc, 1, changes_entries, carried_over_cm=0.0)

    heading_texts = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
    assert heading_texts[1] == "1.1 Викласти в новій редакції в Додатку 1:"
    assert heading_texts[2] == "1.2 Додаток 2 доповнити наступними пунктами:"


def test_add_changes_sections_restate_point_new_participant_uses_add_label():
    """РЕГРЕСІЯ (підтверджено користувачем): людина без жодного дня цієї
    категорії в prev (ні 30, ні 100 - щойно долучилась) - "_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ"
    відсутній/False - рядок рендериться ЗВИЧАЙНИМ "Додаток N доповнити", а НЕ
    "Викласти в новій редакції" (немає що "перевидавати" - запису раніше не
    існувало)."""
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    new_rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", "ПЕРІОД": "02.06.2026-02.06.2026", "ДНІ": 1, "ПІДСТАВА": ""}]
    changes_entries = [{"appendix_pairs": [(30, 3, [], new_rows)]}]

    report_module._add_changes_sections(doc, 1, changes_entries, carried_over_cm=0.0)

    heading_texts = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
    assert heading_texts[1] == "1.1 Додаток 3 доповнити наступними пунктами:"


def test_add_changes_sections_restate_point_mixed_renders_both_subsections():
    """Той самий пункт (30)/місяць МОЖЕ дати ОБИДВА варіанти одночасно (для
    РІЗНИХ людей) - тоді рендеряться ДВОМА окремими підпунктами під ТИМ САМИМ
    номером Додатка, "Викласти в новій редакції" ПЕРШИМ, "доповнити" другим -
    підтверджено користувачем (реальний підсумок: пункт 100 - "доповнити
    Додаток 1"; пункт 30 - 2 підпункти, 1-й "Викласти в новій редакції", 2-й
    "доповнити")."""
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    mixed_rows = [
        {"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", "ПЕРІОД": "02.06.2026-02.06.2026", "ДНІ": 1, "ПІДСТАВА": ""},
        {
            "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Перший Перший",
            "ПЕРІОД": "01.06.2026-05.06.2026", "ДНІ": 5, "ПІДСТАВА": "", "_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ": True,
        },
    ]
    changes_entries = [{"appendix_pairs": [(30, 3, [], mixed_rows)]}]

    report_module._add_changes_sections(doc, 1, changes_entries, carried_over_cm=0.0)

    heading_texts = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
    assert heading_texts[1] == "1.1 Викласти в новій редакції в Додатку 3:"
    assert heading_texts[2] == "1.2 Додаток 3 доповнити наступними пунктами:"


def test_add_changes_sections_empty_entries_leaves_document_unchanged():
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))

    next_point_number, carried_over_cm = report_module._add_changes_sections(doc, 5, [], carried_over_cm=2.0)

    assert next_point_number == 5
    assert carried_over_cm == 2.0
    assert doc.paragraphs == []


def test_add_changes_sections_skips_entry_with_no_appendix_pairs():
    """Запис лише з extra_points (без жодного appendix_pairs) - тут пропускається
    цілком (без порожнього "N. Прошу внести зміни..."); рендериться окремо,
    через _add_extra_point_sections."""
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    changes_entries = [{"appendix_pairs": [], "extra_points": [{"point": 70, "rows": []}]}]

    next_point_number, carried_over_cm = report_module._add_changes_sections(doc, 5, changes_entries, carried_over_cm=0.0)

    assert next_point_number == 5
    assert carried_over_cm == 0.0
    assert doc.paragraphs == []


# -------------------------
# _add_extra_point_sections (70/170 "як у звичайному рапорті", поза "Прошу внести зміни...")
# -------------------------
def test_add_extra_point_sections_renders_own_numbered_point_per_section():
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    rows_70 = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", "ПЕРІОД": "01.07.2026-31.07.2026", "ДНІ": 31, "ПІДСТАВА": "X"}]
    changes_entries = [{
        "appendix_pairs": [], "month": 7,
        "extra_points": [{"point": 70, "rows": rows_70, "period_start_day": 1, "period_end_day": 31}],
    }]

    next_point_number, carried_over_cm = report_module._add_extra_point_sections(doc, 5, changes_entries, carried_over_cm=0.0)

    assert next_point_number == 6
    assert carried_over_cm > 0.0
    heading_texts = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
    expected_text = report_module.STATIK["SECTIONS"]["70"].format(
        start_date_day="01", end_date_day="31", manth=report_module.month_genitive_lower(7), year=report_module.YEAR,
    )
    assert heading_texts[0] == f"5. {expected_text}"


def test_add_extra_point_sections_renders_basis_required_points_too():
    """BASIS_REQUIRED_POINTS ("100_СПЕЦКОНТИНГЕНТ"/"100_БПШП"/"100_ВПБП") -
    ТОЙ САМИЙ рендер, що й 70/170 (str(section["point"]) уже готовий рядок-
    ключ) - "як у звичайному рапорті" (SECTIONS у resources/data.json), а НЕ
    "Додаток N доповнити" - підтверджено користувачем."""
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    rows = [{
        "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Перший Перший", "ПЕРІОД": "01.06.2026-15.06.2026", "ДНІ": 15,
        "ДАТА_ЗНИКНЕННЯ": "01.06.2026", "ПІДСТАВА": "Наказ №5 від 15.06.2026",
    }]
    changes_entries = [{
        "appendix_pairs": [], "month": 6,
        "extra_points": [{"point": "100_СПЕЦКОНТИНГЕНТ", "rows": rows, "period_start_day": 1, "period_end_day": 15}],
    }]

    next_point_number, carried_over_cm = report_module._add_extra_point_sections(doc, 5, changes_entries, carried_over_cm=0.0)

    assert next_point_number == 6
    heading_texts = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
    expected_text = report_module.STATIK["SECTIONS"]["100_СПЕЦКОНТИНГЕНТ"].format(
        start_date_day="01", end_date_day="15", manth=report_module.month_genitive_lower(6), year=report_module.YEAR,
    )
    assert heading_texts[0] == f"5. {expected_text}"
    assert "Додаток" not in heading_texts[0]


def test_add_extra_point_sections_no_entries_leaves_document_unchanged():
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))

    next_point_number, carried_over_cm = report_module._add_extra_point_sections(doc, 5, [], carried_over_cm=0.0)

    assert next_point_number == 5
    assert carried_over_cm == 0.0
    assert doc.paragraphs == []


# -------------------------
# _add_missing_coverage_section (список "відсутні документи" - у сам WORD-документ,
# НЕ в термінал - підтверджено користувачем)
# -------------------------
def test_add_missing_coverage_section_renders_heading_and_each_warning():
    import generators.generate_report_for_get_money as report_module
    from docx import Document

    doc = Document()
    changes_entries = [
        {"missing_coverage_warnings": ["Немає документа за 03.06.2026 для Другий Другий - підстава на цей день відсутня."]},
        {"missing_coverage_warnings": ["Немає документа за 01.07.2026 для Перший Перший - підстава на цей день відсутня."]},
    ]

    report_module._add_missing_coverage_section(doc, changes_entries)

    texts = [p.text for p in doc.paragraphs]
    assert texts[0] == report_module.STATIK["MISSING_COVERAGE_HEADING"]
    assert "Другий Другий" in texts[1]
    assert "Перший Перший" in texts[2]


def test_add_missing_coverage_section_no_warnings_leaves_document_unchanged():
    import generators.generate_report_for_get_money as report_module
    from docx import Document

    doc = Document()
    changes_entries = [{"missing_coverage_warnings": []}, {}]

    report_module._add_missing_coverage_section(doc, changes_entries)

    assert doc.paragraphs == []


# -------------------------
# _add_category_section (_NOT_PAID_POINT)
# -------------------------
def test_add_category_section_renders_not_paid_heading_and_last_column_header():
    """_NOT_PAID_POINT рендериться через ЗВИЧАЙНИЙ _add_category_section
    (section_map_key="SECTIONS", headers=NOT_PAID_TABLE_HEADERS) - окремої
    функції для цього більше немає (псевдо-пункт у самому словнику категорій)."""
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))
    rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Перший Перший", "ПЕРІОД": "05.07.2026-05.07.2026", "ДНІ": 1, "ПІДСТАВА": "СЗЧ"}]

    report_module._add_category_section(
        doc, 4, report_module._NOT_PAID_POINT, rows, 7, 1, 31, carried_over_cm=0.0,
        headers=report_module.NOT_PAID_TABLE_HEADERS,
    )

    assert doc.paragraphs[0].text == f"4. {report_module.STATIK['SECTIONS']['NOT_PAID']}"
    header_row = doc.tables[-1].rows[0]
    assert header_row.cells[-1].text == "Підстава для не виплати"
    assert doc.tables[-1].cell(1, 1).text == "Стрілець"
    assert doc.tables[-1].cell(1, len(report_module.NOT_PAID_TABLE_HEADERS) - 1).text == "СЗЧ"


# -------------------------
# _add_changes_only_intro
# -------------------------
def test_add_changes_only_intro_writes_addressee_and_title_without_accounting_section():
    """Той самий адресат+заголовок, що й _add_intro, але БЕЗ START_SECTION/
    SECTION_INFORMATION (текст про звичайний облік поточного місяця тут не
    застосовний - документ одразу переходить до пунктів "Прошу внести зміни...")."""
    import generators.generate_report_for_get_money as report_module
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    report_module.set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.4003))

    report_module._add_changes_only_intro(doc)

    texts = [p.text for p in doc.paragraphs]
    assert report_module.STATIK["HIGH_COMMANDER"] in texts
    assert report_module.STATIK["REPORT_TITLE"] in texts
    assert not any(report_module.STATIK["START_SECTION"] in t for t in texts)


# -------------------------
# _flatten_changes_rows_for_logging
# -------------------------
def test_flatten_changes_rows_for_logging_builds_section_keys_per_point():
    import generators.generate_report_for_get_money as report_module

    exclude_rows = [{"ПІБ": "Другий Другий"}]
    add_rows = [{"ПІБ": "Перший Перший"}]
    changes_entries = [{"appendix_pairs": [(30, 1, exclude_rows, add_rows), (100, 2, [], add_rows)]}]

    flattened = report_module._flatten_changes_rows_for_logging(changes_entries)

    assert flattened == [
        ("30_CHANGES_EXCLUDE", exclude_rows),
        ("30_CHANGES_ADD", add_rows),
        ("100_CHANGES_ADD", add_rows),
    ]


def test_flatten_changes_rows_for_logging_includes_extra_points():
    """extra_points (70/170 "як у звичайному рапорті", поза "Прошу внести
    зміни...") теж потрапляють у результат - під ключем "{point}_CHANGES_EXTRA"."""
    import generators.generate_report_for_get_money as report_module

    rows = [{"ПІБ": "Перший Перший"}]
    changes_entries = [{"appendix_pairs": [], "extra_points": [{"point": 70, "rows": rows}]}]

    flattened = report_module._flatten_changes_rows_for_logging(changes_entries)

    assert flattened == [("70_CHANGES_EXTRA", rows)]


def test_flatten_changes_rows_for_logging_empty_entries_returns_empty():
    import generators.generate_report_for_get_money as report_module

    assert report_module._flatten_changes_rows_for_logging([]) == []


# -------------------------
# _exclude_by_excluded_day_value
# -------------------------
def test_exclude_by_excluded_day_value_removes_row_with_prvd_day_cell():
    """Виключається людина, у якої ХОЧ ОДНА комірка дати за місяць - "ПРВД"
    (незалежно від регістру), навіть якщо решта днів - звичайні 30/100."""
    import generators.generate_report_for_get_money as report_module

    d1, d2, d3 = (datetime(2026, 7, d) for d in (1, 2, 3))
    rows = [
        {"ПІБ": "Другий Другий", d1: 100, d2: "првд", d3: 100},  # регістр не має значення
        {"ПІБ": "Перший Перший", d1: 100, d2: 100, d3: 30},
    ]

    result = report_module._exclude_by_excluded_day_value(rows, [d1, d2, d3])

    assert [row["ПІБ"] for row in result] == ["Перший Перший"]


def test_exclude_by_excluded_day_value_keeps_szch_rows():
    """"СЗЧ" БІЛЬШЕ не виключає людину з рапорту цілком (на відміну від "ПРВД") -
    це звичайна категорія пункту "NOT_PAID" (MONEY_REPORT_CATEGORIES), рахується
    day-by-day через _build_categories, а не бланкетним виключенням тут."""
    import generators.generate_report_for_get_money as report_module

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 2)
    rows = [{"ПІБ": "Перший Перший", d1: 100, d2: "СЗЧ"}]

    result = report_module._exclude_by_excluded_day_value(rows, [d1, d2])

    assert result == rows


def test_exclude_by_excluded_day_value_keeps_rows_without_excluded_values():
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [{"ПІБ": "БЕЗ ВИКЛЮЧЕНЬ", d1: 100}, {"ПІБ": "ПОРОЖНІЙ ДЕНЬ", d1: None}]
    result = report_module._exclude_by_excluded_day_value(rows, [d1])
    assert result == rows


# -------------------------
# _exclude_by_szch_started_this_month - СЗЧ, що почався ЦЬОГО місяця (на
# відміну від СЗЧ, перенесеного з попереднього - можливе повернення)
# -------------------------
def test_szch_started_this_month_excludes_when_no_return():
    """Кейс 1: 01-10 - звичайний статус, 11 - СЗЧ (без повернення до кінця
    місяця) - людина виключається ЦІЛКОМ (не потрапляє в рапорт/БР)."""
    import content.money_report_helpers as mrh

    d1, d2, d3 = datetime(2026, 7, 1), datetime(2026, 7, 11), datetime(2026, 7, 20)
    rows = [{"ПІБ": "Перший Перший", d1: 100, d2: "СЗЧ", d3: "СЗЧ"}]

    result = mrh._exclude_by_szch_started_this_month(rows, [d1, d2, d3])

    assert result == []


def test_szch_started_this_month_excludes_even_with_return_same_month():
    """Кейс 2: 01-10 - звичайний статус, 11 - СЗЧ, 17 - повернення (звичайний
    статус) - УСЕ ОДНО виключається: раз СЗЧ ПОЧАВСЯ цього місяця, подальше
    повернення в тому ж місяці нічого не змінює - місяць уже "зіпсований"."""
    import content.money_report_helpers as mrh

    d1, d2, d3 = datetime(2026, 7, 1), datetime(2026, 7, 11), datetime(2026, 7, 17)
    rows = [{"ПІБ": "Перший Перший", d1: 100, d2: "СЗЧ", d3: 30}]

    result = mrh._exclude_by_szch_started_this_month(rows, [d1, d2, d3])

    assert result == []


def test_szch_from_previous_month_with_return_is_not_excluded():
    """Кейс 3: 01-10 - СЗЧ (уже на самому початку місяця - перенесено з
    попереднього), 11 - звичайний статус (10/30/100/70/170) - НЕ виключається -
    це "повернення", а не "втеча цього місяця"."""
    import content.money_report_helpers as mrh

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 11)
    rows = [{"ПІБ": "Перший Перший", d1: "СЗЧ", d2: 30}]

    result = mrh._exclude_by_szch_started_this_month(rows, [d1, d2])

    assert result == rows


def test_szch_from_previous_month_return_to_zvrez_is_still_excluded():
    """Нюанс: СЗЧ з попереднього місяця, повернення - але ПЕРШИЙ статус ПІСЛЯ
    СЗЧ-періоду - САМЕ "ЗВРез" - усе одно виключається, попри "повернення"."""
    import content.money_report_helpers as mrh

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 11)
    rows = [{"ПІБ": "Перший Перший", d1: "СЗЧ", d2: "ЗВРез"}]

    result = mrh._exclude_by_szch_started_this_month(rows, [d1, d2])

    assert result == []


def test_szch_from_previous_month_return_to_zvrez_is_case_insensitive():
    import content.money_report_helpers as mrh

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 11)
    rows = [{"ПІБ": "Перший Перший", d1: "сзч", d2: "звРЕЗ"}]

    result = mrh._exclude_by_szch_started_this_month(rows, [d1, d2])

    assert result == []


def test_szch_entire_month_with_no_return_is_not_excluded_by_this_rule():
    """Увесь місяць - СЗЧ (без жодного дня повернення) - НЕ виключається ЦІЄЮ
    функцією (нюанс ЗВРез тут не застосовний - немає дня "після" СЗЧ) -
    звичайна per-day категоризація ("NOT_PAID"/"СЗЧ") і далі рахує ці дні як
    неоплачувані, без окремого повного виключення."""
    import content.money_report_helpers as mrh

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 2)
    rows = [{"ПІБ": "УВЕСЬ МІСЯЦЬ СЗЧ", d1: "СЗЧ", d2: "СЗЧ"}]

    result = mrh._exclude_by_szch_started_this_month(rows, [d1, d2])

    assert result == rows


def test_szch_started_this_month_case_insensitive_and_empty_cells_skipped():
    """Порожня комірка ПЕРШОГО дня (людина ще не мала запису) НЕ рахується як
    "перший день був СЗЧ" (першим КАЛЕНДАРНИМ днем тут є САМЕ порожня комірка
    d1, а не наступний непорожній день d2=100) - тож це відхід ЦЬОГО місяця
    (d3="сзч" з'явився ПІЗНІШЕ); регістр не має значення ("сзч")."""
    import content.money_report_helpers as mrh

    d1, d2, d3 = datetime(2026, 7, 1), datetime(2026, 7, 5), datetime(2026, 7, 15)
    rows = [{"ПІБ": "Перший Перший", d1: None, d2: 100, d3: "сзч"}]

    result = mrh._exclude_by_szch_started_this_month(rows, [d1, d2, d3])

    assert result == []


def test_szch_no_szch_at_all_is_not_excluded():
    import content.money_report_helpers as mrh

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 2)
    rows = [{"ПІБ": "Перший Перший", d1: 100, d2: 30}]

    result = mrh._exclude_by_szch_started_this_month(rows, [d1, d2])

    assert result == rows


def test_szch_started_this_month_not_excluded_when_no_date_columns():
    """date_columns=[] (напр. обраний місяць узагалі не має жодної колонки) -
    _row_started_szch_this_month не має що сканувати, тож нікого не виключає,
    а не падає з IndexError на date_columns[0]."""
    import content.money_report_helpers as mrh

    rows = [{"ПІБ": "Перший Перший"}]

    result = mrh._exclude_by_szch_started_this_month(rows, [])

    assert result == rows


def test_szch_leading_blank_days_then_only_szch_is_started_this_month():
    """РЕГРЕСІЯ (виявлено користувачем на реальному прикладі): комірки ПЕРШИХ
    днів місяця ПОРОЖНІ (немає запису), а далі - СУЦІЛЬНИЙ "СЗЧ" аж до кінця
    місяця, без жодного ІНШОГО значення. Раніше це помилково розпізнавалось як
    "перший НЕПОРОЖНІЙ день - уже СЗЧ" (кейс 3 - перенесено з попереднього
    місяця, НЕ виключати), хоча порожня комірка НЕ доводить жодного статусу
    "до" СЗЧ - однаково ЗАВЖДИ виключається (СЗЧ почався ЦЬОГО місяця)."""
    import content.money_report_helpers as mrh

    d1, d2, d3 = datetime(2026, 8, 1), datetime(2026, 8, 2), datetime(2026, 8, 3)
    rows = [{"ПІБ": "ПЕРШИЙ Перший Перший", d1: None, d2: None, d3: "СЗЧ"}]

    result = mrh._exclude_by_szch_started_this_month(rows, [d1, d2, d3])

    assert result == []


# -------------------------
# _blank_non_szch_days_for_szch_started_this_month - рапорт на додаткову
# винагороду (на відміну від БР): людина, чий СЗЧ почався цього місяця, НЕ
# зникає з рапорту цілком - лише не отримує оплату за жоден звичайний пункт,
# АЛЕ й далі з'являється в "NOT_PAID" за свої дні СЗЧ.
# -------------------------
def test_blank_non_szch_days_blanks_regular_days_but_keeps_szch_days():
    import content.money_report_helpers as mrh

    d1, d2, d3 = datetime(2026, 8, 1), datetime(2026, 8, 2), datetime(2026, 8, 3)
    rows = [{"ПІБ": "ДРУГИЙ Другий Другий", d1: 30, d2: "СЗЧ", d3: "СЗЧ"}]

    result = mrh._blank_non_szch_days_for_szch_started_this_month(rows, [d1, d2, d3])

    assert result == [{"ПІБ": "ДРУГИЙ Другий Другий", d1: None, d2: "СЗЧ", d3: "СЗЧ"}]


def test_blank_non_szch_days_leaves_carried_over_szch_untouched():
    """Кейс 3 (СЗЧ з попереднього місяця, НЕ почався цього) - рядок лишається
    ПОВНІСТЮ незмінним (жодна комірка не порожниться)."""
    import content.money_report_helpers as mrh

    d1, d2 = datetime(2026, 8, 1), datetime(2026, 8, 11)
    rows = [{"ПІБ": "Перший Перший", d1: "СЗЧ", d2: 30}]

    result = mrh._blank_non_szch_days_for_szch_started_this_month(rows, [d1, d2])

    assert result == rows


def test_blank_non_szch_days_leaves_rows_without_szch_untouched():
    import content.money_report_helpers as mrh

    d1 = datetime(2026, 8, 1)
    rows = [{"ПІБ": "Другий Другий", d1: 100}]

    result = mrh._blank_non_szch_days_for_szch_started_this_month(rows, [d1])

    assert result == rows


# -------------------------
# _row_ignorable_not_paid_statuses / _blank_not_paid_days_confirmed_elsewhere -
# NOT_PAID_IGNORED_WHEN_CONFIRMED_BY (constants.py, за замовчуванням
# {"СЗЧ": {"РОЗП"}}) - "СЗЧ", ПІЗНІШЕ підтверджений "РОЗП" (людину офіційно
# оголошено в розшуку) - ігнорується (реальний випадок, підтверджений
# користувачем: ТРЕТІЙ Третій Третій, 01.08-04.08 - "СЗЧ", з 05.08 -
# "РОЗП"). Керовано з constants.py - тести нижче звіряються з РЕАЛЬНИМ
# словником (не монкіпатчать його), крім окремого теста на розширюваність.
# -------------------------
def test_row_ignorable_not_paid_statuses_when_confirmed():
    import content.money_report_helpers as mrh

    d1, d2, d3 = datetime(2026, 8, 1), datetime(2026, 8, 4), datetime(2026, 8, 5)
    row = {"ПІБ": "ТРЕТІЙ Третій Третій", d1: "СЗЧ", d2: "СЗЧ", d3: "РОЗП"}

    assert mrh._row_ignorable_not_paid_statuses(row, [d1, d2, d3]) == {"СЗЧ"}


def test_row_ignorable_not_paid_statuses_empty_when_not_confirmed():
    import content.money_report_helpers as mrh

    d1, d2 = datetime(2026, 8, 1), datetime(2026, 8, 4)
    row = {"ПІБ": "ЗВИЧАЙНИЙ СЗЧ", d1: "СЗЧ", d2: "СЗЧ"}

    assert mrh._row_ignorable_not_paid_statuses(row, [d1, d2]) == set()


def test_row_ignorable_not_paid_statuses_regardless_of_order():
    """"РОЗП" ДО "СЗЧ" (а не після) - усе одно підтверджує (порядок НЕ важливий -
    реальний випадок, ЧЕТВЕРТИЙ Четвертий Четвертий: СЗЧ 01.08-17.08, РОЗП
    18.08-30.08, знову СЗЧ 31.08 - "РОЗП" тут не останній хронологічно, але
    людину все одно вже офіційно визнано в розшуку)."""
    import content.money_report_helpers as mrh

    d1, d2 = datetime(2026, 8, 1), datetime(2026, 8, 4)
    row = {"ПІБ": "НЕЗВИЧНИЙ ПОРЯДОК", d1: "РОЗП", d2: "СЗЧ"}

    assert mrh._row_ignorable_not_paid_statuses(row, [d1, d2]) == {"СЗЧ"}


def test_row_ignorable_not_paid_statuses_when_szch_reappears_after_rozp():
    """Реальний випадок (ЧЕТВЕРТИЙ Четвертий Четвертий): СЗЧ 01-17.08, РОЗП
    18-30.08, знову СЗЧ 31.08 (позначка на межі місяця) - усі "СЗЧ"-дні (і ДО,
    і ПІСЛЯ "РОЗП") мають ігноруватись однаково."""
    import content.money_report_helpers as mrh

    d1, d2, d3, d4 = datetime(2026, 8, 1), datetime(2026, 8, 17), datetime(2026, 8, 18), datetime(2026, 8, 31)
    row = {"ПІБ": "ЧЕТВЕРТИЙ Четвертий Четвертий", d1: "СЗЧ", d2: "СЗЧ", d3: "РОЗП", d4: "СЗЧ"}

    assert mrh._row_ignorable_not_paid_statuses(row, [d1, d2, d3, d4]) == {"СЗЧ"}


def test_row_ignorable_not_paid_statuses_case_insensitive_and_skips_empty():
    import content.money_report_helpers as mrh

    d1, d2, d3 = datetime(2026, 8, 1), datetime(2026, 8, 2), datetime(2026, 8, 3)
    row = {"ПІБ": "РЕГІСТР", d1: None, d2: "сзч", d3: "розп"}

    assert mrh._row_ignorable_not_paid_statuses(row, [d1, d2, d3]) == {"СЗЧ"}


def test_row_ignorable_not_paid_statuses_driven_by_constants_not_hardcoded(monkeypatch):
    """РОЗШИРЮВАНІСТЬ (мета цього рефакторингу - винесено в constants.py, щоб
    користувач міг керувати без змін коду): монкіпатчимо
    NOT_PAID_IGNORED_WHEN_CONFIRMED_BY зовсім ІНШОЮ парою - код не повинен
    десь усередині лишати хардкод "СЗЧ"/"РОЗП"."""
    import content.money_report_helpers as mrh

    monkeypatch.setattr(mrh, "NOT_PAID_IGNORED_WHEN_CONFIRMED_BY", {"Задув.": {"Повер"}})

    d1, d2 = datetime(2026, 8, 1), datetime(2026, 8, 4)
    row = {"ПІБ": "ПЕРШИЙ ПЕРШИЙ", d1: "Задув.", d2: "Повер"}

    assert mrh._row_ignorable_not_paid_statuses(row, [d1, d2]) == {"ЗАДУВ."}
    # Стара пара "СЗЧ"/"РОЗП" вже НЕ підтверджує нічого, поки монкіпатч активний.
    szch_row = {"ПІБ": "СЗЧ РОЗП", d1: "СЗЧ", d2: "РОЗП"}
    assert mrh._row_ignorable_not_paid_statuses(szch_row, [d1, d2]) == set()


def test_blank_not_paid_days_confirmed_elsewhere_clears_only_ignorable_cells():
    """Замінюються ЛИШЕ комірки зі значенням буквально "СЗЧ" - "РОЗП"-дні й
    решта категорій людини лишаються без змін."""
    import content.money_report_helpers as mrh

    d1, d2, d3, d4 = (datetime(2026, 8, d) for d in (1, 2, 3, 4))
    rows = [{"ПІБ": "ТРЕТІЙ Третій Третій", d1: "СЗЧ", d2: "СЗЧ", d3: "РОЗП", d4: 100}]

    result = mrh._blank_not_paid_days_confirmed_elsewhere(rows, [d1, d2, d3, d4])

    assert result == [{"ПІБ": "ТРЕТІЙ Третій Третій", d1: None, d2: None, d3: "РОЗП", d4: 100}]
    # Оригінальний rows не мутується на місці.
    assert rows[0][d1] == "СЗЧ"


def test_blank_not_paid_days_confirmed_elsewhere_leaves_other_people_unchanged():
    import content.money_report_helpers as mrh

    d1, d2 = datetime(2026, 8, 1), datetime(2026, 8, 2)
    rows = [
        {"ПІБ": "ТРЕТІЙ Третій Третій", d1: "СЗЧ", d2: "РОЗП"},
        {"ПІБ": "ЗВИЧАЙНИЙ СЗЧ", d1: "СЗЧ", d2: "СЗЧ"},
    ]

    result = mrh._blank_not_paid_days_confirmed_elsewhere(rows, [d1, d2])

    assert result[0] == {"ПІБ": "ТРЕТІЙ Третій Третій", d1: None, d2: "РОЗП"}
    assert result[1] == rows[1]


def test_build_categories_ignores_szch_days_confirmed_by_rozp():
    """Наскрізь: людина, чиї "СЗЧ"-дні пізніше підтверджені "РОЗП" - НЕ
    з'являється в пункті NOT_PAID узагалі (ні рядком, ні попередженням "Немає
    підстави для не виплати") - підтверджено користувачем на реальному
    випадку."""
    import generators.generate_report_for_get_money as report_module

    d1, d2, d3, d4 = (datetime(2026, 8, d) for d in (1, 2, 4, 5))
    rows = [_row("ТРЕТІЙ Третій Третій", "Стрілець", "Підрозділ 1", {d1: "СЗЧ", d2: "СЗЧ", d3: "СЗЧ", d4: "РОЗП"})]

    date_columns = [d1, d2, d3, d4]
    filtered_rows = mrh._blank_not_paid_days_confirmed_elsewhere(rows, date_columns)
    categories = report_module._build_categories(filtered_rows, date_columns, {})

    not_paid_rows = dict(categories).get("NOT_PAID", [])
    assert "ТРЕТІЙ Третій Третій" not in {r["ПІБ"] for r in not_paid_rows}


def test_build_categories_ignores_szch_days_confirmed_by_rozp_even_when_szch_reappears_after():
    """Наскрізь, реальний випадок (ЧЕТВЕРТИЙ Четвертий Четвертий): СЗЧ 01-17.08,
    РОЗП 18-30.08, знову СЗЧ 31.08 - РАНІШЕ (коли перевірка вимагала "РОЗП
    строго ПІСЛЯ останнього СЗЧ") попередження й далі показувалось за ОБИДВА
    періоди "СЗЧ" (01.08-17.08 і 31.08-31.08), бо останній хронологічно
    статус - знову "СЗЧ", а не "РОЗП". Тепер - жодного рядка/попередження
    взагалі, незалежно від порядку."""
    import generators.generate_report_for_get_money as report_module

    d1, d2, d3, d4 = datetime(2026, 8, 1), datetime(2026, 8, 17), datetime(2026, 8, 18), datetime(2026, 8, 31)
    rows = [_row("ЧЕТВЕРТИЙ Четвертий Четвертий", "Стрілець", "Підрозділ 1", {d1: "СЗЧ", d2: "СЗЧ", d3: "РОЗП", d4: "СЗЧ"})]

    date_columns = [d1, d2, d3, d4]
    filtered_rows = mrh._blank_not_paid_days_confirmed_elsewhere(rows, date_columns)
    categories = report_module._build_categories(filtered_rows, date_columns, {})

    not_paid_rows = dict(categories).get("NOT_PAID", [])
    assert "ЧЕТВЕРТИЙ Четвертий Четвертий" not in {r["ПІБ"] for r in not_paid_rows}


# -------------------------
# build_pidstavy_extra_grounds_by_person
# -------------------------
def test_build_pidstavy_extra_grounds_by_person_collects_nonblank_cells():
    rows = [
        {"ПІБ": "Перший Перший", "ПІДСТАВИ": "  Наказ №12 від 03.07.2026  "},
        {"ПІБ": "Другий Другий Другий", "ПІДСТАВИ": ""},
        {"ПІБ": "Третій Третій Третій"},
    ]

    result = mrh.build_pidstavy_extra_grounds_by_person(rows)

    assert result == {"ПЕРШИЙ ПЕРШИЙ": ["Наказ №12 від 03.07.2026"]}


def test_build_pidstavy_extra_grounds_by_person_treats_nan_as_blank():
    import numpy as np

    rows = [{"ПІБ": "Перший Перший", "ПІДСТАВИ": np.nan}]
    assert mrh.build_pidstavy_extra_grounds_by_person(rows) == {}


def test_build_pidstavy_extra_grounds_by_person_accepts_singular_column_name():
    """Реальний ОБЛІК.xlsx називає цю колонку в однині - "ПІДСТАВА" (не
    "ПІДСТАВИ") - обидва написання мають рівноправно розпізнаватись
    (utils.excel_reader.OPTIONAL_PERSONEL_COLUMN_NAMES)."""
    rows = [{"ПІБ": "ШОСТИЙ Шостий", "ПІДСТАВА": "№100\\132/n від 30.06.2026"}]

    result = mrh.build_pidstavy_extra_grounds_by_person(rows)

    assert result == {"ШОСТИЙ ШОСТИЙ": ["№100\\132/n від 30.06.2026"]}


def test_build_pidstavy_extra_grounds_by_person_prefers_plural_when_both_present():
    rows = [{"ПІБ": "Перший Перший", "ПІДСТАВИ": "Множина", "ПІДСТАВА": "Однина"}]
    assert mrh.build_pidstavy_extra_grounds_by_person(rows) == {"ПЕРШИЙ ПЕРШИЙ": ["Множина"]}


def test_build_pidstavy_extra_grounds_by_person_prefers_new_general_column_name():
    """"ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ" (GENERAL_PIDSTAVY_COLUMN_NAMES, constants.py) -
    новий, пріоритетний запис загальної підстави - перевіряється ПЕРШИМ, перед
    старими "ПІДСТАВИ"/"ПІДСТАВА" (лишеними як запасний варіант)."""
    rows = [{"ПІБ": "Перший Перший", "ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ": "Нова", "ПІДСТАВИ": "Множина", "ПІДСТАВА": "Однина"}]
    assert mrh.build_pidstavy_extra_grounds_by_person(rows) == {"ПЕРШИЙ ПЕРШИЙ": ["Нова"]}


def test_build_pidstavy_extra_grounds_by_person_accepts_custom_column_names():
    """column_names - явно передані назви колонок (BASIS_REQUIRED_POINT_COLUMN_NAMES,
    constants.py) - ІГНОРУЄ загальні "ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ"/"ПІДСТАВИ"/"ПІДСТАВА",
    навіть якщо вони непорожні - підтверджено користувачем: кожен статус
    BASIS_REQUIRED_POINTS має власну, незалежну колонку."""
    rows = [{
        "ПІБ": "Перший Перший",
        "ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ": "Загальна - не повинна враховуватись",
        "ПІДСТАВИ ВПБП": "Наказ №1 від 01.07.2026",
    }]
    result = mrh.build_pidstavy_extra_grounds_by_person(rows, ("ПІДСТАВИ ВПБП",))
    assert result == {"ПЕРШИЙ ПЕРШИЙ": ["Наказ №1 від 01.07.2026"]}


# -------------------------
# "NOT_PAID" - СЗЧ/Задув. як звичайні категорії (_build_categories)
# -------------------------
def test_build_categories_puts_szch_and_zaduv_days_under_not_paid_point():
    """"СЗЧ"/"Задув." - звичайні названі категорії пункту "NOT_PAID" (та сама
    механіка резолюції дня за текстом комірки, що й "100_ШП"/"РТГр" тощо) -
    "Підстава" береться з колонки "ПІДСТАВИ" (build_pidstavy_extra_grounds_by_person),
    бо "grounds"/"general" категорій "СЗЧ"/"Задув." порожні."""
    import generators.generate_report_for_get_money as report_module

    d1, d2, d3 = (datetime(2026, 7, d) for d in (1, 2, 3))
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 100, d2: "СЗЧ", d3: "Задув."})]
    rows[0]["ПІДСТАВИ"] = "Довідка №5 від 03.07.2026"

    categories = report_module._build_categories(rows, [d1, d2, d3], {})

    not_paid_rows = dict(categories)["NOT_PAID"]
    assert {r["ПІБ"] for r in not_paid_rows} == {"Перший Перший"}
    assert all(r["ПІДСТАВА"] == "Довідка №5 від 03.07.2026" for r in not_paid_rows)
    # 100-день і далі рахується у звичайному пункті 100 - СЗЧ не виключає решту місяця.
    assert "Перший Перший" in {r["ПІБ"] for r in dict(categories)["100"]}


def test_build_categories_pidstavy_column_applies_to_every_point():
    """Колонка "ПІДСТАВИ" - НЕ лише для "NOT_PAID": build_pidstavy_extra_grounds_by_person
    передається в КОЖЕН виклик _build_combined_section, тож текст додається до
    "Підстава для виплати" й у звичайних пунктах (тут - 30), поверх наявних grounds."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30})]
    rows[0]["ПІДСТАВИ"] = "Особиста примітка"

    categories = report_module._build_categories(rows, [d1], {})

    thirty_row = dict(categories)["30"][0]
    assert "Особиста примітка" in thirty_row["ПІДСТАВА"]


# -------------------------
# _exclude_commander_and_tvo
# -------------------------
def test_exclude_commander_and_tvo_removes_regular_commander_and_period_tvo():
    """Штатний командир батальйону завжди виключається; ТВО - якщо його період
    (Start/End) перетинається з ЗВІТНИМ періодом, а не лише з "сьогодні"."""
    import generators.generate_report_for_get_money as report_module

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 31)
    rows = [
        _row("Третій Третій", "Командир батальйону", "упр", {d1: 100}),
        _row("Другий Другий", "Заступник командира батальйону", "упр", {d1: 100}),
        _row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 100}),
    ]
    rows_with_tvo_data = [
        {"ПІБ": "Другий Другий", "ПОСАДА": "Командир батальйону", "Start": "01.07.2026", "End": "31.07.2026", "ТВО": True},
    ]

    result = report_module._exclude_commander_and_tvo(rows, rows_with_tvo_data, [d1, d2], {"ТВО": True})

    assert [r["ПІБ"] for r in result] == ["Перший Перший"]


def test_exclude_commander_and_tvo_ignores_tvo_outside_reporting_period():
    import generators.generate_report_for_get_money as report_module

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 31)
    rows = [_row("Другий Другий", "Заступник командира батальйону", "упр", {d1: 100})]
    rows_with_tvo_data = [
        {"ПІБ": "Другий Другий", "ПОСАДА": "Командир батальйону", "Start": "01.09.2026", "End": "30.09.2026", "ТВО": True},
    ]

    result = report_module._exclude_commander_and_tvo(rows, rows_with_tvo_data, [d1, d2], {"ТВО": True})

    assert [r["ПІБ"] for r in result] == ["Другий Другий"]


def test_log_personnel_day_value_periods_writes_file_and_prints(tmp_path, monkeypatch, capsys):
    import generators.generate_report_for_get_money as report_module

    d1, d2, d3 = (datetime(2026, 7, d) for d in (1, 2, 3))
    rows = [
        {"ПІБ": "Другий Другий", d1: "ВД", d2: "ВД", d3: "ПРВД"},
        {"ПІБ": "Перший Перший", d1: None, d2: "", d3: None},
    ]
    log_path = tmp_path / "log.txt"

    report_module._log_personnel_day_value_periods(rows, [d1, d2, d3], str(log_path))

    content = log_path.read_text(encoding="utf-8")
    assert content == "Другий Другий - ВД: 01.07.2026-02.07.2026 (2 дн.); ПРВД: 03.07.2026-03.07.2026 (1 дн.)"
    assert "Перший Перший" not in content

    printed = capsys.readouterr().out
    assert "Другий Другий" in printed
    assert "Перший Перший" not in printed


def test_log_missing_legal_basis_prints_only_rows_without_basis(capsys):
    import generators.generate_report_for_get_money as report_module

    categories = [
        ("70", [
            {"ПІБ": "Перший Перший", "ПЕРІОД": "01.07.2026-31.07.2026", "ПІДСТАВА": ""},
            {"ПІБ": "Другий Другий", "ПЕРІОД": "01.07.2026-31.07.2026", "ПІДСТАВА": "БР ... від 01.07.2026"},
        ]),
    ]

    report_module._log_missing_legal_basis(categories)

    printed = capsys.readouterr().out
    assert "Перший Перший" in printed
    assert "70" in printed
    assert "Другий Другий" not in printed


def test_log_missing_legal_basis_skips_rows_with_optional_basis(capsys):
    """Категорії з "required_basis": False (напр. 10к) не повинні попереджати про
    порожню підставу - вона там ОЧІКУВАНА, а не помилка."""
    import generators.generate_report_for_get_money as report_module

    categories = [
        ("10", [
            {"ПІБ": "Другий Другий", "ПЕРІОД": "01.07.2026-31.07.2026", "ПІДСТАВА": "", "_ПІДСТАВА_ОБОВ'ЯЗКОВА": False},
        ]),
        ("70", [
            {"ПІБ": "Перший Перший", "ПЕРІОД": "01.07.2026-31.07.2026", "ПІДСТАВА": "", "_ПІДСТАВА_ОБОВ'ЯЗКОВА": True},
        ]),
    ]

    report_module._log_missing_legal_basis(categories)

    printed = capsys.readouterr().out
    assert "Другий Другий" not in printed


def test_log_missing_legal_basis_reports_specific_message_when_daily_brs_folder_not_read(capsys):
    """Категорії з "use_brs_from_selected_folder": True (напр. МЕДИК пункту 100), у яких
    підстава вийшла порожньою (папку з документами не зчитували цього запуску),
    отримують ОКРЕМЕ, конкретне повідомлення "завдання не вдалось зчитати" замість
    загального "Немає підстави для виплати: ПІБ (...)"."""
    import generators.generate_report_for_get_money as report_module

    categories = [
        ("100", [
            {
                "ПІБ": "Перший Перший", "ПЕРІОД": "01.07.2026-31.07.2026", "ПІДСТАВА": "",
                "_ПІДСТАВА_ОБОВ'ЯЗКОВА": True, "_ВИКОРИСТОВУЄ_ОБРАНУ_ПАПКУ": True, "_КАТЕГОРІЯ": "МЕДИК",
            },
        ]),
    ]

    report_module._log_missing_legal_basis(categories)

    printed = capsys.readouterr().out
    assert "Для 100 статус МЕДИК завдання не вдалось зчитати." in printed
    assert "Перший Перший" not in printed
    assert "Немає підстави для виплати" not in printed


def test_log_missing_legal_basis_dedups_specific_message_per_point_and_category(capsys):
    """Кілька людей з тієї самої (пункт, категорія) і порожньою підставою -
    повідомлення друкується ОДИН раз, а не по одному на кожну людину."""
    import generators.generate_report_for_get_money as report_module

    categories = [
        ("100", [
            {
                "ПІБ": "Перший Перший", "ПЕРІОД": "01.07.2026-31.07.2026", "ПІДСТАВА": "",
                "_ПІДСТАВА_ОБОВ'ЯЗКОВА": True, "_ВИКОРИСТОВУЄ_ОБРАНУ_ПАПКУ": True, "_КАТЕГОРІЯ": "МЕДИК",
            },
            {
                "ПІБ": "Другий Другий", "ПЕРІОД": "01.07.2026-31.07.2026", "ПІДСТАВА": "",
                "_ПІДСТАВА_ОБОВ'ЯЗКОВА": True, "_ВИКОРИСТОВУЄ_ОБРАНУ_ПАПКУ": True, "_КАТЕГОРІЯ": "МЕДИК",
            },
        ]),
    ]

    report_module._log_missing_legal_basis(categories)

    printed = capsys.readouterr().out
    assert printed.count("Для 100 статус МЕДИК завдання не вдалось зчитати.") == 1


def test_log_missing_legal_basis_ignores_use_brs_from_selected_folder_flag_when_basis_present(capsys):
    """Якщо підстава ВЖЕ заповнена (папку зчитали успішно), прапорець
    "_ВИКОРИСТОВУЄ_ОБРАНУ_ПАПКУ" не має жодного ефекту - жодне повідомлення не друкується."""
    import generators.generate_report_for_get_money as report_module

    categories = [
        ("100", [
            {
                "ПІБ": "Перший Перший", "ПЕРІОД": "01.07.2026-31.07.2026", "ПІДСТАВА": "БР ... від 01.07.2026",
                "_ПІДСТАВА_ОБОВ'ЯЗКОВА": True, "_ВИКОРИСТОВУЄ_ОБРАНУ_ПАПКУ": True, "_КАТЕГОРІЯ": "МЕДИК",
            },
        ]),
    ]

    report_module._log_missing_legal_basis(categories)

    printed = capsys.readouterr().out
    assert printed == ""


def test_log_missing_legal_basis_uses_not_paid_specific_message(capsys):
    """NOT_PAID (і секції "Прошу внести зміни..." NOT_PAID_CHANGES_EXCLUDE/_ADD) -
    ОКРЕМЕ формулювання "Немає підстави ДЛЯ НЕ ВИПЛАТИ" (не "для виплати", як для
    решти пунктів) - підтверджено користувачем: цей пункт описує причину невиплати
    (СЗЧ/задув.), а не підставу для виплати."""
    import generators.generate_report_for_get_money as report_module

    categories = [
        ("NOT_PAID", [
            {"ПІБ": "Другий Другий", "ПЕРІОД": "01.07.2026-31.07.2026", "ПІДСТАВА": ""},
        ]),
        ("NOT_PAID_CHANGES_ADD", [
            {"ПІБ": "Третій Третій", "ПЕРІОД": "01.06.2026-30.06.2026", "ПІДСТАВА": ""},
        ]),
        ("70", [
            {"ПІБ": "Перший Перший", "ПЕРІОД": "01.07.2026-31.07.2026", "ПІДСТАВА": ""},
        ]),
    ]

    report_module._log_missing_legal_basis(categories)

    printed = capsys.readouterr().out
    assert "Немає підстави для не виплати: Другий Другий" in printed
    assert "Немає підстави для не виплати: Третій Третій" in printed
    assert "Немає підстави для виплати: Перший Перший" in printed
    assert "Немає підстави для виплати: ЗАДУВ" not in printed


def test_build_categories_bpshp_requires_nonblank_pidstavy_column():
    """100_БПШП - гейт за ВЛАСНОЮ колонкою "ПІДСТАВИ БПШП" (BASIS_REQUIRED_POINT_COLUMN_NAMES,
    constants.py) - НЕ спільна "ПІДСТАВИ" - підтверджено користувачем: кожен
    статус BASIS_REQUIRED_POINTS має власну колонку, незалежну від інших."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: "БПШП"}),
        _row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: "БПШП"}),
    ]
    rows[0]["ПІДСТАВИ БПШП"] = "Наказ №1 від 01.07.2026"
    rows[1]["ПІДСТАВИ БПШП"] = ""

    categories = report_module._build_categories(rows, [d1], {})

    section = dict(categories).get("100_БПШП", [])
    assert {r["ПІБ"] for r in section} == {"Перший Перший Перший"}


def test_build_categories_bpshp_accepts_real_file_column_spelling_shpbp():
    """РЕГРЕСІЯ (підтверджено користувачем на реальному випадку - Перший Перший
    Перший помилково виключений з рапорту): реальний ОБЛІК.xlsx називає цю
    колонку "ПІДСТАВИ ШПБП" (літери переставлені відносно самого ключа пункту
    "100_БПШП") - BASIS_REQUIRED_POINT_COLUMN_NAMES (constants.py) визнає ОБИДВА
    написання, "ШПБП" і "БПШП"."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: "БПШП"})]
    rows[0]["ПІДСТАВИ ШПБП"] = "Довідка про обставини травми №2371/2026-д5/136 від 23.06.2026"

    categories = report_module._build_categories(rows, [d1], {})

    section = dict(categories).get("100_БПШП", [])
    assert {r["ПІБ"] for r in section} == {"Перший Перший Перший"}


def test_build_categories_vpbp_requires_nonblank_pidstavy_column():
    """100_ВПБП - гейт за ВЛАСНОЮ колонкою "ПІДСТАВИ ВПБП" (BASIS_REQUIRED_POINT_COLUMN_NAMES,
    constants.py) - НЕ спільна "ПІДСТАВИ" - підтверджено користувачем: кожен
    статус BASIS_REQUIRED_POINTS (100_СПЕЦКОНТИНГЕНТ/100_БПШП/100_ВПБП, там же
    в constants.py) має власну колонку, незалежну від інших."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: "ВПБП"}),
        _row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: "ВПБП"}),
    ]
    rows[0]["ПІДСТАВИ ВПБП"] = "Наказ №1 від 01.07.2026"
    rows[1]["ПІДСТАВИ ВПБП"] = ""

    categories = report_module._build_categories(rows, [d1], {})

    section = dict(categories).get("100_ВПБП", [])
    assert {r["ПІБ"] for r in section} == {"Перший Перший Перший"}


def test_build_categories_basis_required_points_do_not_share_columns():
    """РЕГРЕСІЯ (підтверджено користувачем): "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ" заповнена
    для людини зі статусом "ВПБП" - НЕ повинна хибно "розблокувати" її для
    100_ВПБП - кожен статус BASIS_REQUIRED_POINTS дивиться ЛИШЕ на СВОЮ колонку
    (BASIS_REQUIRED_POINT_COLUMN_NAMES, constants.py). Раніше всі три статуси
    ділили ОДНУ спільну "ПІДСТАВИ" - будь-яка з них "розблоковувала" всі три."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: "ВПБП"})]
    rows[0]["ПІДСТАВИ СПЕЦКОНТИНГЕНТУ"] = "Наказ №1 від 01.07.2026"

    categories = report_module._build_categories(rows, [d1], {})

    assert dict(categories).get("100_ВПБП", []) == []


def test_build_categories_warns_when_excluding_vpbp_person_without_basis(capsys):
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: "ВПБП"})]

    report_module._build_categories(rows, [d1], {})

    printed = capsys.readouterr().out
    assert "Немає підстави для 100_ВПБП: Другий Другий Другий" in printed
    assert "виключено з рапорту" in printed


def test_build_categories_warns_when_excluding_person_without_basis(capsys):
    """Виключення з BASIS_REQUIRED_POINTS (100_СПЕЦКОНТИНГЕНТ/100_БПШП) через
    порожню ПІДСТАВИ друкує попередження в термінал з іменем людини й пунктом -
    раніше виключення було мовчазним - підтверджено користувачем."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: "полон"})]

    report_module._build_categories(rows, [d1], {})

    printed = capsys.readouterr().out
    assert "Немає підстави для 100_СПЕЦКОНТИНГЕНТ: Другий Другий Другий" in printed
    assert "виключено з рапорту" in printed


def test_build_categories_does_not_warn_when_basis_present(capsys):
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: "полон"})]
    rows[0]["ПІДСТАВИ СПЕЦКОНТИНГЕНТУ"] = "Наказ №1 від 01.07.2026"

    report_module._build_categories(rows, [d1], {})

    printed = capsys.readouterr().out
    assert "Немає підстави для 100_СПЕЦКОНТИНГЕНТ" not in printed


def test_log_missing_posada_prints_only_rows_without_posada(capsys):
    import generators.generate_report_for_get_money as report_module

    categories = [
        ("30", [
            {"ПІБ": "ОДИНАДЦЯТИЙ Одинадцятий Одинадцятий", "ПОСАДА": ""},
            {"ПІБ": "Другий Другий", "ПОСАДА": "Стрілець"},
        ]),
    ]

    report_module._log_missing_posada(categories)

    printed = capsys.readouterr().out
    assert "для ОДИНАДЦЯТИЙ Одинадцятий Одинадцятий немає посади" in printed
    assert "Другий Другий" not in printed


def test_log_missing_posada_adds_changes_section_suffix_for_changes_keys(capsys):
    """section_key з "_CHANGES_EXCLUDE"/"_CHANGES_ADD" (розділ "Прошу внести
    зміни...") отримує уточнення в тексті повідомлення - звичайні пункти рапорту
    (без цього суфікса в ключі) лишаються без нього."""
    import generators.generate_report_for_get_money as report_module

    categories = [
        ("100_CHANGES_EXCLUDE", [{"ПІБ": "ОДИНАДЦЯТИЙ Одинадцятий Одинадцятий", "ПОСАДА": ""}]),
        ("100_CHANGES_ADD", [{"ПІБ": "ОДИНАДЦЯТИЙ Одинадцятий Одинадцятий", "ПОСАДА": ""}]),
        ("30", [{"ПІБ": "Другий Другий", "ПОСАДА": ""}]),
    ]

    report_module._log_missing_posada(categories)

    printed = capsys.readouterr().out
    assert printed.count("для ОДИНАДЦЯТИЙ Одинадцятий Одинадцятий немає посади в секції зміни до наказу") == 2
    assert "для Другий Другий немає посади" in printed
    assert "для Другий Другий немає посади в секції зміни до наказу" not in printed


def test_process_generate_report_for_get_money_end_to_end(tmp_path, monkeypatch):
    import generators.generate_report_for_get_money as report_module
    from process_generate_report_for_get_additional_money import process_generate_report_for_get_money
    from utils.excel_reader import read_datafile, read_optional_datafile, PERSONEL_LIST_FALLBACK_SHEET_NAMES

    monkeypatch.setattr(report_module, "OUTPUT_DIR", str(tmp_path))

    rows_data = read_datafile(
        constants.PERSONEL_LIST_FILE_NAME, constants.PERSONEL_LIST_SHEET_NAME, constants.PERSONEL_LIST_COLUMNS_LETTERS,
        extra_fallback_sheet_names=PERSONEL_LIST_FALLBACK_SHEET_NAMES,
    )
    tvo_rows = read_optional_datafile(constants.PERSONEL_LIST_FILE_NAME, constants.TVO_LIST_SHEET_NAME, constants.TVO_LIST_COLUMNS_LETTERS, sheet_can_be_missing=True)

    result_path = process_generate_report_for_get_money(
        rows_data.get("rows", []), tvo_rows, rows_data.get("columns", []),
    )

    assert result_path is not None
    assert os.path.isfile(result_path)
    assert os.path.dirname(os.path.abspath(result_path)) == str(tmp_path)

    text = read_docx_text(result_path)
    assert "РАПОРТ" in text
    assert "1. " in text


def test_process_generate_report_for_get_money_adds_not_paid_section_for_szch_and_zaduv(tmp_path, monkeypatch):
    """Пункт "Не виплачувати додаткову винагороду..." з'являється в кінці рапорту
    для тих, у кого хоч один день місяця - "СЗЧ" чи "Задув." (звичайні категорії
    пункту "NOT_PAID"). ЖОДНЕ з двох значень більше НЕ виключає людину з рапорту
    цілком (на відміну від "ПРВД") - її ІНШІ дні місяця й далі рахуються за
    своїми звичайними категоріями як завжди, тож людина цілком очікувано
    з'являється і в звичайному пункті (за свій робочий день), і в "NOT_PAID"
    (за свій день СЗЧ/Задув.).

    Перший Перший: d1 (ПЕРШИЙ день) - "СЗЧ" (перенесено з ПОПЕРЕДНЬОГО місяця -
    НЕ "почався цього місяця", інакше спрацював би ОКРЕМИЙ виняток
    _blank_non_szch_days_for_szch_started_this_month і d2 теж занулився б - див.
    той тест окремо), d2 - звичайне повернення (100)."""
    import generators.generate_report_for_get_money as report_module
    from process_generate_report_for_get_additional_money import process_generate_report_for_get_money

    monkeypatch.setattr(report_module, "OUTPUT_DIR", str(tmp_path))

    # _add_footer підписує рапорт командиром - потрібен ТВО-запис на HIGHER_COMMANDER_TITLE,
    # чинний на СЬОГОДНІ (find_higher_commander перевіряє саме datetime.now(), а не
    # період рапорту), інакше підпис не збудується взагалі (не пов'язано з цим пунктом).
    tvo_rows = [{"ПОСАДА": constants.HIGHER_COMMANDER_TITLE, "ТВО": True, "Start": "01.01.2020", "End": "31.12.2030", "ЗВАННЯ": "підполковник", "ПІБ": "Третій Третій Третій"}]

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 2)
    rows_with_data = [
        {"ПІДРОЗДІЛ": "Підрозділ 1", "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", d1: 100, d2: 100},
        {"ПІДРОЗДІЛ": "Підрозділ 1", "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Перший Перший", d1: "СЗЧ", d2: 100},
        {"ПІДРОЗДІЛ": "Підрозділ 2", "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Четвертий Четвертий", d1: "Задув.", d2: 100},
    ]
    columns = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", d1, d2]

    result_path = process_generate_report_for_get_money(rows_with_data, tvo_rows, columns)

    assert result_path is not None
    text = read_docx_text(result_path)

    assert report_module.STATIK["SECTIONS"]["NOT_PAID"] in text
    assert "Четвертий Четвертий" in text
    # Перший Перший - у ЗВИЧАЙНОМУ пункті (100, за свій робочий d2) І в "NOT_PAID"
    # (за свій СЗЧ-день d1) - двічі, а не виключений з рапорту цілком.
    assert text.count("Перший Перший") == 2


def test_process_generate_report_for_get_money_orders_not_paid_before_changes_section(tmp_path, monkeypatch):
    """РЕГРЕСІЯ (підтверджено користувачем): обраний (серпневий) місяць - у
    ПРІОРИТЕТІ - усі його пункти, включно з "NOT_PAID" ("Не виплачувати..."),
    мають йти ПЕРШИМИ, а "Прошу внести зміни..." за ПОПЕРЕДНІ місяці - лише
    ПІСЛЯ них. Раніше "NOT_PAID" рендерився В САМОМУ КІНЦІ документа (ПІСЛЯ
    розділу змін), через що виглядав як частина попередніх місяців, хоча
    стосується САМЕ обраного місяця."""
    import generators.generate_report_for_get_money as report_module
    import content.report_changes as report_changes_module
    from openpyxl import Workbook
    from process_generate_report_for_get_additional_money import process_generate_report_for_get_money

    monkeypatch.setattr(report_module, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(report_module, "MONTH", "08")
    monkeypatch.setattr(report_module, "YEAR", "2026")
    monkeypatch.setattr(report_changes_module, "MONTH", "08")
    monkeypatch.setattr(report_changes_module, "YEAR", "2026")

    tvo_rows = [{"ПОСАДА": constants.HIGHER_COMMANDER_TITLE, "ТВО": True, "Start": "01.01.2020", "End": "31.12.2030", "ЗВАННЯ": "підполковник", "ПІБ": "Третій Третій Третій"}]

    d1 = datetime(2026, 8, 1)
    rows_with_data = [
        {"ПІДРОЗДІЛ": "Підрозділ 1", "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", d1: 100},
        {"ПІДРОЗДІЛ": "Підрозділ 1", "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Перший Перший", d1: "СЗЧ"},
    ]
    columns = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", d1]

    def _write_xlsx(path, headers, rows):
        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        for row in rows:
            ws.append(row)
        wb.save(path)
        return path

    d_july = datetime(2026, 7, 1)
    prev_path = _write_xlsx(tmp_path / "prev_07_ОБЛІК.xlsx", ["ПІБ", d_july], [["Четвертий Четвертий", "ВД"]])
    actual_path = _write_xlsx(tmp_path / "actual_07_ОБЛІК.xlsx", ["ПІБ", d_july], [["Четвертий Четвертий", 30]])
    changes_file_pairs = [{
        "month_token": "07", "prev_path": str(prev_path), "actual_path": str(actual_path), "changes_path": "unused",
    }]

    result_path = process_generate_report_for_get_money(rows_with_data, tvo_rows, columns, changes_file_pairs=changes_file_pairs)

    assert result_path is not None
    from docx import Document
    doc = Document(result_path)
    heading_texts = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]

    not_paid_index = next(i for i, text in enumerate(heading_texts) if "Не виплачувати" in text)
    changes_index = next(i for i, text in enumerate(heading_texts) if "Прошу внести зміни" in text)
    assert not_paid_index < changes_index


def test_process_generate_report_for_get_money_omits_not_paid_section_when_no_matches(tmp_path, monkeypatch):
    import generators.generate_report_for_get_money as report_module
    from process_generate_report_for_get_additional_money import process_generate_report_for_get_money

    monkeypatch.setattr(report_module, "OUTPUT_DIR", str(tmp_path))

    tvo_rows = [{"ПОСАДА": constants.HIGHER_COMMANDER_TITLE, "ТВО": True, "Start": "01.01.2020", "End": "31.12.2030", "ЗВАННЯ": "підполковник", "ПІБ": "Третій Третій Третій"}]

    d1 = datetime(2026, 7, 1)
    rows_with_data = [{"ПІДРОЗДІЛ": "Підрозділ 1", "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", d1: 100}]
    columns = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", d1]

    result_path = process_generate_report_for_get_money(rows_with_data, tvo_rows, columns)

    text = read_docx_text(result_path)
    assert report_module.STATIK["SECTIONS"]["NOT_PAID"] not in text


def test_process_generate_report_for_get_money_szch_started_this_month_appears_only_in_not_paid(tmp_path, monkeypatch):
    """РЕГРЕСІЯ (підтверджено користувачем на реальному прикладі): в/сл, чий
    СЗЧ ПОЧАВСЯ ЦЬОГО місяця (без повернення) - НЕ отримує оплату за жоден
    звичайний пункт (30/70/100/170/10, навіть за дні ДО СЗЧ того самого
    місяця), АЛЕ й далі з'являється в "NOT_PAID" за свої дні СЗЧ - раніше
    зникав з рапорту ЦІЛКОМ (жодного сліду ні в звичайних пунктах, ні в
    "NOT_PAID")."""
    import generators.generate_report_for_get_money as report_module
    from process_generate_report_for_get_additional_money import process_generate_report_for_get_money

    monkeypatch.setattr(report_module, "OUTPUT_DIR", str(tmp_path))

    tvo_rows = [{"ПОСАДА": constants.HIGHER_COMMANDER_TITLE, "ТВО": True, "Start": "01.01.2020", "End": "31.12.2030", "ЗВАННЯ": "підполковник", "ПІБ": "Третій Третій Третій"}]

    d1, d2, d3 = datetime(2026, 8, 1), datetime(2026, 8, 2), datetime(2026, 8, 3)
    rows_with_data = [
        {"ПІДРОЗДІЛ": "Підрозділ 1", "ПОСАДА": "Стрілець", "ЗВАННЯ": "матрос", "ПІБ": "СЬОМИЙ Сьомий", d1: 30, d2: "СЗЧ", d3: "СЗЧ"},
    ]
    columns = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", d1, d2, d3]

    result_path = process_generate_report_for_get_money(rows_with_data, tvo_rows, columns)

    assert result_path is not None
    text = read_docx_text(result_path)

    assert report_module.STATIK["SECTIONS"]["NOT_PAID"] in text
    # З'являється РІВНО один раз - лише в "NOT_PAID" (2 дні СЗЧ, одним рядком
    # періоду), а НЕ ще й у пункті 30 (d1 заблоковано, а не переведено туди).
    assert text.count("СЬОМИЙ Сьомий") == 1


def test_process_generate_changes_report_no_entries_returns_none(monkeypatch, capsys):
    """RUN_MODE_CHANGES_ONLY, але немає жодної підтвердженої зміни (напр. немає
    елігібельних пар у resources/changes) - жодного файлу не створюється."""
    import generators.generate_report_for_get_money as report_module

    monkeypatch.setattr(report_module, "build_changes_entries", lambda *args, **kwargs: [])

    result = report_module.process_generate_changes_report([], [], changes_file_pairs=[])

    assert result is None
    assert "не сформовано" in capsys.readouterr().out


def test_process_generate_changes_report_end_to_end(tmp_path, monkeypatch):
    """Документ містить ЛИШЕ розділи "Прошу внести зміни..." (адресат/заголовок +
    _add_changes_sections + підпис) - без жодного зі звичайних пунктів рапорту."""
    import generators.generate_report_for_get_money as report_module
    from process_generate_report_for_get_additional_money import process_generate_changes_report
    from utils.excel_reader import read_datafile, read_optional_datafile, PERSONEL_LIST_FALLBACK_SHEET_NAMES

    monkeypatch.setattr(report_module, "OUTPUT_DIR", str(tmp_path))

    exclude_rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Другий Другий", "ПЕРІОД": "01.06.2026-30.06.2026", "ДНІ": 30, "ПІДСТАВА": "X"}]
    add_rows = [{"ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": "Перший Перший", "ПЕРІОД": "01.06.2026-15.06.2026", "ДНІ": 15, "ПІДСТАВА": "Y"}]
    changes_entries = [{"appendix_pairs": [(30, 1, exclude_rows, add_rows)]}]
    monkeypatch.setattr(report_module, "build_changes_entries", lambda *args, **kwargs: changes_entries)

    rows_data = read_datafile(
        constants.PERSONEL_LIST_FILE_NAME, constants.PERSONEL_LIST_SHEET_NAME, constants.PERSONEL_LIST_COLUMNS_LETTERS,
        extra_fallback_sheet_names=PERSONEL_LIST_FALLBACK_SHEET_NAMES,
    )
    tvo_rows = read_optional_datafile(constants.PERSONEL_LIST_FILE_NAME, constants.TVO_LIST_SHEET_NAME, constants.TVO_LIST_COLUMNS_LETTERS, sheet_can_be_missing=True)

    result_path = process_generate_changes_report(rows_data.get("rows", []), tvo_rows, changes_file_pairs=[])

    assert result_path is not None
    assert os.path.isfile(result_path)
    assert os.path.dirname(os.path.abspath(result_path)) == str(tmp_path)

    text = read_docx_text(result_path)
    expected_intro = report_module.STATIK["CHANGES_INTRO"].format(order_reference=report_module._BLANK_ORDER_REFERENCE)
    assert report_module.STATIK["REPORT_TITLE"] in text
    assert "1. " + expected_intro in text
    assert report_module.STATIK["START_SECTION"] not in text


def test_build_categories_merges_medic_enemy_defense_into_single_point_two():
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший", "Бойовий медик", "Підрозділ 1", {d1: 100}),
        _row("Другий Другий", "Стрілець", "Підрозділ 2", {d1: 100}),
        _row("Третій Третій", "Стрілець", "Підрозділ 3", {d1: 100}),
    ]
    status_lookup = {
        "ПЕРШИЙ ПЕРШИЙ": "МЕДИК",
        "ДРУГИЙ ДРУГИЙ": "РТГр",
        "ТРЕТІЙ ТРЕТІЙ": "ОБОРОНА",
    }
    categories = report_module._build_categories(rows, [d1], status_lookup)

    template_keys = [key for key, _ in categories]
    assert template_keys == ["30", "100"]

    hundred_rows = dict(categories)["100"]
    assert {r["ПІБ"] for r in hundred_rows} == {"Перший Перший", "Другий Другий", "Третій Третій"}


def test_build_categories_skips_hundred_point_when_empty():
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30})]
    categories = report_module._build_categories(rows, [d1], {})

    assert [key for key, _ in categories] == ["30"]


def test_build_categories_splits_thirty_point_by_zvrez_and_life_support():
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 30}),
        _row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30}),
    ]
    status_lookup = {"Перший Перший": "ЗВРез"}
    categories = report_module._build_categories(rows, [d1], status_lookup)

    thirty_rows = dict(categories)["30"]
    assert {r["ПІБ"] for r in thirty_rows} == {"Перший Перший", "Другий Другий"}


def test_build_categories_status_change_mid_month_splits_between_points_by_day_value():
    """СТАТУС в ОБЛІК.xlsx - одне значення на весь місяць, тож людину, що фактично
    перейшла з ЗВРЕЗ на ОБОРОНУ посеред місяця (а колонка й далі показує "ЗВРез"),
    не можна розпізнати як "ОБОРОНА" за статусом. Розв'язання - "ОБОРОНА" є
    catch-all для пункту 2 (усі 100-дні, крім МЕДИК/РТГр), тож 30-дні (до зміни)
    все одно йдуть у ЗВРЕЗ (пункт 1), а 100-дні (після зміни) - у ОБОРОНУ (пункт 2),
    виключно за рахунок різних значень 30/100 по днях, без статусу "ОБОРОНА" явно."""
    import generators.generate_report_for_get_money as report_module

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 14)
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 30, d2: 100})]
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "ЗВРез"}
    categories = report_module._build_categories(rows, [d1, d2], status_lookup)

    thirty_rows = dict(categories)["30"]
    hundred_rows = dict(categories)["100"]
    assert [r["ДНІ"] for r in thirty_rows if r["ПІБ"] == "Перший Перший"] == [1]
    assert [r["ДНІ"] for r in hundred_rows if r["ПІБ"] == "Перший Перший"] == [1]


def test_build_categories_defense_catchall_excludes_medic_and_enemy_territory():
    """ОБОРОНА (catch-all пункту 2) не повинна дублювати людей, що вже враховані
    в МЕДИК чи РТГр (у них теж 100-дні, але своя окрема підстава)."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший", "Бойовий медик", "Підрозділ 1", {d1: 100}),
        _row("Другий Другий", "Стрілець", "Підрозділ 2", {d1: 100}),
    ]
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "МЕДИК", "ДРУГИЙ ДРУГИЙ": "РТГр"}
    categories = report_module._build_categories(rows, [d1], status_lookup)

    hundred_rows = dict(categories)["100"]
    assert {r["ПІБ"] for r in hundred_rows} == {"Перший Перший", "Другий Другий"}


def test_build_categories_hundred_section_preserves_oblik_file_order_across_categories():
    """Рядки 100_SECTION складаються з окремих виликів build_category_rows
    (по одному на кожну категорію - МЕДИК, РТГр, ОБОРОНА), тож просте конкатенування
    групувало б їх по категоріях, а не за файлом. Тут порядок у файлі навмисно
    "перемішаний" (ОБОРОНА, потім МЕДИК, потім РТГр) - результат має бути точно
    таким самим, а не згрупованим по категоріях."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Третій Третій", "Стрілець", "Підрозділ 1", {d1: 100}),
        _row("Другий Другий", "Бойовий медик", "Підрозділ 1", {d1: 100}),
        _row("Перший Перший", "Стрілець", "Підрозділ 2", {d1: 100}),
        _row("Четвертий Четвертий", "Стрілець", "Підрозділ 1", {d1: 100}),
    ]
    status_lookup = {
        "ТРЕТІЙ ТРЕТІЙ": "ОБОРОНА",
        "ДРУГИЙ ДРУГИЙ": "МЕДИК",
        "ПЕРШИЙ ПЕРШИЙ": "РТГр",
        "ЧЕТВЕРТИЙ ЧЕТВЕРТИЙ": "ОБОРОНА",
    }
    categories = report_module._build_categories(rows, [d1], status_lookup)

    hundred_rows = dict(categories)["100"]
    assert [r["ПІБ"] for r in hundred_rows] == [
        "Третій Третій", "Другий Другий", "Перший Перший", "Четвертий Четвертий",
    ]


def test_build_combined_section_defaults_to_global_categories_when_none_given():
    """_build_combined_section, викликана НАПРЯМУ (без _build_categories, який
    завжди резолвить categories заздалегідь) з categories=None, сама підставляє
    глобальний MONEY_REPORT_CATEGORIES - той самий запасний варіант, що й
    resolve_category_config за замовчуванням."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 100})]
    status_lookup = {"ДРУГИЙ ДРУГИЙ": "РТГр"}

    result = report_module._build_combined_section(rows, [d1], status_lookup, 100, {100}, categories=None)

    assert [r["ПІБ"] for r in result] == ["Другий Другий"]


def test_build_combined_section_returns_empty_when_point_missing_from_categories():
    """Викликана НАПРЯМУ (не через _build_categories, яка тепер сама динамічно йде
    лише по точках, що ДІЙСНО є в categories, і ніколи не викличе цю функцію з
    відсутньою точкою) - _build_combined_section все одно захищена від цього
    напряму: відсутня точка -> порожній список, без падіння."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 100})]

    result = report_module._build_combined_section(rows, [d1], {}, 100, {100}, categories={30: {}})

    assert result == []


def test_build_combined_section_without_catchall_flag_drops_unmatched_people():
    """Якщо в categories[point] жодна категорія не позначена "default" - у пункту
    просто немає catch-all: хто не потрапив у названу категорію, ніде не з'являється
    (а не падає і не потрапляє кудись за замовчуванням, як це було б із захардкодженою
    назвою catch-all)."""
    import generators.generate_report_for_get_money as report_module

    categories = {100: {"general": [], "РТГр": {"grounds": [], "use_brs": False, "exclude_general": []}}}
    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 100}),
        _row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 100}),
    ]
    status_lookup = {"ДРУГИЙ ДРУГИЙ": "РТГр"}

    result = report_module._build_combined_section(rows, [d1], status_lookup, 100, {100}, categories=categories)

    assert [r["ПІБ"] for r in result] == ["Другий Другий"]


def test_build_combined_section_named_category_excluded_from_report_is_hidden_but_not_in_catchall():
    """"include_to_report": False на НАЗВАНІЙ категорії - її дні не друкуються в
    рапорті, але й НЕ переходять у catch-all (вони все одно розпізнані саме як ця
    категорія - просто не показуються)."""
    import generators.generate_report_for_get_money as report_module

    categories = {
        100: {
            "general": [],
            "РТГр": {"grounds": [], "use_brs": False, "exclude_general": [], "include_to_report": False},
            "БД(СЗ)": {"grounds": [], "use_brs": False, "exclude_general": [], "default": True},
        },
    }
    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 100})]
    status_lookup = {"ДРУГИЙ ДРУГИЙ": "РТГр"}

    result = report_module._build_combined_section(rows, [d1], status_lookup, 100, {100}, categories=categories)

    assert result == []


def test_build_combined_section_catchall_excluded_from_report_still_shows_named_rows():
    """"include_to_report": False на catch-all - лише сам catch-all не друкується,
    названі категорії того самого пункту виводяться як завжди."""
    import generators.generate_report_for_get_money as report_module

    categories = {
        100: {
            "general": [],
            "РТГр": {"grounds": [], "use_brs": False, "exclude_general": []},
            "БД(СЗ)": {"grounds": [], "use_brs": False, "exclude_general": [], "default": True, "include_to_report": False},
        },
    }
    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 100}),
        _row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 100}),
    ]
    status_lookup = {"ДРУГИЙ ДРУГИЙ": "РТГр"}

    result = report_module._build_combined_section(rows, [d1], status_lookup, 100, {100}, categories=categories)

    assert [r["ПІБ"] for r in result] == ["Другий Другий"]


def test_build_categories_seventy_and_one_seventy_sections_use_oborona_catchall():
    """70к/170к - той самий "catch-all ОБОРОНА + названі категорії" механізм, що й
    100к: людина без розпізнаного статусу все одно потрапляє в catch-all."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 70}),
        _row("Третій Третій", "Стрілець", "Підрозділ 1", {d1: 170}),
    ]
    categories = report_module._build_categories(rows, [d1], {})

    assert dict(categories)["70"][0]["ПІБ"] == "Другий Другий"
    assert dict(categories)["170"][0]["ПІБ"] == "Третій Третій"


def test_build_categories_seventy_and_one_seventy_specific_category_excluded_from_catchall():
    """"70_РТГр"/"170_РТГр" - названі категорії своїх точок, як "РТГр" у 100к:
    не повинні дублюватись у catch-all ОБОРОНА тієї самої точки."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 70}),
        _row("Четвертий Четвертий", "Стрілець", "Підрозділ 1", {d1: 170}),
    ]
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "70_РТГр", "ЧЕТВЕРТИЙ ЧЕТВЕРТИЙ": "170_РТГр"}
    categories = report_module._build_categories(rows, [d1], status_lookup)

    assert [r["ПІБ"] for r in dict(categories)["70"]] == ["Перший Перший"]
    assert [r["ПІБ"] for r in dict(categories)["170"]] == ["Четвертий Четвертий"]


def test_build_categories_seventy_and_one_seventy_days_also_add_row_in_hundred():
    """70-дні й 170-дні (catch-all ОБОРОНА чи "*_РТГр") ДОДАТКОВО потрапляють ще й
    окремим рядком у 100_SECTION - з власною (100-ю) підставою, поверх їхніх
    рядків у 70_SECTION/170_SECTION, а не замість них."""
    import generators.generate_report_for_get_money as report_module

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 2)
    rows = [
        _row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 70}),
        _row("Третій Третій", "Стрілець", "Підрозділ 1", {d1: 70, d2: 170}),
    ]
    categories = report_module._build_categories(rows, [d1, d2], {})
    by_key = dict(categories)

    assert [r["ПІБ"] for r in by_key["70"]] == ["Другий Другий", "Третій Третій"]
    assert [r["ПІБ"] for r in by_key["170"]] == ["Третій Третій"]
    # обидва (і 70-only, і 70+170) отримують рядок і в 100-й теж
    assert {r["ПІБ"] for r in by_key["100"]} == {"Другий Другий", "Третій Третій"}
    obydva_hundred_row = next(r for r in by_key["100"] if r["ПІБ"] == "Третій Третій")
    assert obydva_hundred_row["ДНІ"] == 2


def test_build_categories_thirty_point_recognizes_real_file_spelling_30_rtgr():
    """РЕГРЕСІЯ (виявлено користувачем на реальному прикладі):
    реальний ОБЛІК.xlsx позначає цю категорію пункту 30 САМЕ "30_РТГр" (з
    префіксом номера пункту, як "70_РТГр"/"170_РТГр" у своїх пунктах) - раніше
    constants.MONEY_REPORT_CATEGORIES[30] реєстрував категорію як голе "РТГр",
    яке НІКОЛИ не збігалося з реальним текстом комірки - людина мовчки зникала
    з рапорту ЦІЛКОМ (без жодного попередження в термінал), а не потрапляла в
    catch-all "ЖИТТЄДІЯЛЬНІСТЬ"."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 8, 1)
    rows = [_row("Перший Перший", "Стрілець-санітар", "рв", {d1: "30_РТГр"})]
    categories = report_module._build_categories(rows, [d1], {})

    assert [r["ПІБ"] for r in dict(categories)["30"]] == ["Перший Перший"]


def test_build_categories_rtgr_cross_tier_extra_row_lands_in_hundred_rtgr_not_generic_hundred():
    """РЕГРЕСІЯ (підтверджено користувачем на реальному прикладі): РТГр-в/сл,
    позначені "70_РТГр"/"170_РТГр" (як і решта - test_build_categories_seventy_
    and_one_seventy_days_also_add_row_in_hundred), теж отримують "зайвий" рядок
    пункту 100, АЛЕ САМЕ в секції "100_РТГр" (разом з рештою РТГр, напр. тими,
    чия комірка буквально "100_РТГр"), а НЕ в загальній "100"/"БД(СЗ)" - раніше
    цей рядок помилково потрапляв у загальну "100"."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 70}),
        _row("Четвертий Четвертий", "Стрілець", "Підрозділ 1", {d1: 170}),
        _row("Третій Третій", "Стрілець", "Підрозділ 1", {d1: "100_РТГр"}),
    ]
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "70_РТГр", "ЧЕТВЕРТИЙ ЧЕТВЕРТИЙ": "170_РТГр"}
    categories = report_module._build_categories(rows, [d1], status_lookup)
    by_key = dict(categories)

    assert {r["ПІБ"] for r in by_key["100_РТГр"]} == {
        "Перший Перший", "Четвертий Четвертий", "Третій Третій",
    }
    generic_hundred_pibs = {r["ПІБ"] for r in by_key.get("100", [])}
    assert "Перший Перший" not in generic_hundred_pibs
    assert "Четвертий Четвертий" not in generic_hundred_pibs


def test_rtgr_cross_tier_rows_returns_empty_without_hundred_rtgr_point():
    """rtgr_cross_tier_rows - безпечний no-op, коли categories взагалі не має
    пункту "100_РТГр" (напр. COMMANDER_MONEY_REPORT_CATEGORIES)."""
    d1 = datetime(2026, 7, 1)
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 70})]
    categories = {70: {"general": [], "70_РТГр": {"grounds": [], "exclude_general": []}}}

    assert mrh.rtgr_cross_tier_rows(rows, [d1], {}, categories) == []


def test_rtgr_cross_tier_rows_returns_empty_when_hundred_rtgr_point_has_no_categories():
    """"100_РТГр" присутній у categories, але має ЛИШЕ "general" (жодної
    справжньої категорії) - resolve_reportable_categories повертає порожній
    specs, тож rtgr_cross_tier_rows так само безпечний no-op."""
    d1 = datetime(2026, 7, 1)
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 70})]
    categories = {"100_РТГр": {"general": []}}

    assert mrh.rtgr_cross_tier_rows(rows, [d1], {}, categories) == []


def test_resolve_reportable_categories_excludes_rtgr_cross_tier_categories_when_present_elsewhere():
    """catch-all пункту 100 не приймає "70_РТГр"/"170_РТГр" - але ЛИШЕ якщо вони
    дійсно існують як названа категорія ІНШОГО пункту цього ж categories (інакше
    виняток "протікав" би й у categories, що цієї бізнес-логіки не моделюють -
    див. test_resolve_reportable_categories_excludes_named_category_with_include_
    to_report_false, де "РТГр" - звичайна названа категорія пункту 100 самого,
    а не 70_РТГр/170_РТГр)."""
    categories = {
        70: {"general": [], "70_РТГр": {"grounds": [], "exclude_general": []}},
        170: {"general": [], "170_РТГр": {"grounds": [], "exclude_general": []}},
        100: {
            "general": [],
            "МЕДИК": {"grounds": [], "exclude_general": []},
            "БД(СЗ)": {"grounds": [], "exclude_general": [], "default": True},
        },
    }
    specs = mrh.resolve_reportable_categories(100, categories)

    catchall_spec = next(s for s in specs if s["category_name"] == "БД(СЗ)")
    assert catchall_spec["exclude_status"] == {"МЕДИК", "70_РТГр", "170_РТГр"}


def test_ordered_points_uses_fixed_priority_not_ascending_number():
    """30, 100, 70, 170, 10 - САМЕ в цьому порядку (бізнес-правило пріоритету
    додатків), а НЕ за зростанням номера точки (10, 30, 70, 100, 170), навіть
    якщо всі п'ять присутні в categories."""
    categories = {10: {}, 30: {}, 70: {}, 100: {}, 170: {}}
    assert mrh._ordered_points(categories) == [30, 100, 70, 170, 10]


def test_ordered_points_skips_missing_known_points_without_breaking_order():
    """Немає 10 і 70 - решта (30, 100, 170) лишаються у ТОМУ САМОМУ відносному
    порядку, просто без пропущених точок (перелік коротшає, а не ламається)."""
    categories = {30: {}, 100: {}, 170: {}}
    assert mrh._ordered_points(categories) == [30, 100, 170]


def test_ordered_points_appends_unknown_points_sorted_after_known_ones():
    """Нова точка (напр. "50"), додана лише в MONEY_REPORT_CATEGORIES без жодних
    змін коду - з'являється ПІСЛЯ п'яти пріоритетних точок, за зростанням номера
    серед самих невідомих точок (тут - лише одна, "50")."""
    categories = {30: {}, 50: {}, 100: {}}
    assert mrh._ordered_points(categories) == [30, 100, 50]


def test_ordered_points_empty_categories_returns_empty():
    assert mrh._ordered_points({}) == []


def test_ordered_points_handles_mixed_int_and_string_unknown_points():
    """Пункт може бути й довільним рядком (напр. "100_ШП" - окрема категорія, не
    прив'язана до числового пункту виплати), не лише числом - раніше це падало
    з TypeError ('<' not supported between instances of 'str' and 'int') у
    sorted(). Числові "інші" пункти йдуть за зростанням, рядкові - окремим
    блоком після них, за алфавітом."""
    categories = {30: {}, 50: {}, "100_ШП": {}, "10_ІНШЕ": {}}
    assert mrh._ordered_points(categories) == [30, 50, "100_ШП", "10_ІНШЕ"]


def test_build_categories_orders_populated_points_by_priority_not_by_number():
    """Наскрізна перевірка (а не лише _ordered_points ізольовано): коли ОДНОЧАСНО
    заповнені кілька пунктів, підсумковий СПИСОК (а не dict - порядок реально
    важливий для нумерації пунктів у самому рапорті) іде в порядку 30, 100, 70,
    170, 10 - НЕ за зростанням номера точки."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 10}),
        _row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 70}),
        _row("Третій Третій", "Стрілець", "Підрозділ 1", {d1: 100}),
        _row("Четвертий Четвертий", "Стрілець", "Підрозділ 1", {d1: 170}),
        _row("Шостий Шостий", "Стрілець", "Підрозділ 1", {d1: 30}),
    ]

    categories = report_module._build_categories(rows, [d1], {})

    section_keys = [key for key, rows in categories if rows]
    assert section_keys == ["30", "100", "70", "170", "10"]


# -------------------------
# resolve_reportable_categories
# -------------------------
def test_resolve_reportable_categories_returns_named_then_catchall_specs():
    categories = {
        30: {
            "general": [],
            "ЗВРез": {"grounds": [], "exclude_general": [], "include_to_report": True},
            "ЖИТТЄДІЯЛЬНІСТЬ": {"grounds": [], "exclude_general": [], "default": True, "include_to_report": True},
        },
    }
    specs = mrh.resolve_reportable_categories(30, categories)

    assert [s["category_name"] for s in specs] == ["ЗВРез", "ЖИТТЄДІЯЛЬНІСТЬ"]
    assert specs[0] == {"category_name": "ЗВРез", "status_filter": "ЗВРез", "exclude_status": None}
    assert specs[1] == {"category_name": "ЖИТТЄДІЯЛЬНІСТЬ", "status_filter": None, "exclude_status": {"ЗВРез"}}


def test_resolve_reportable_categories_excludes_named_category_with_include_to_report_false():
    """Категорія з "include_to_report": False не потрапляє в результат (не
    друкується), АЛЕ її назва все одно є серед exclude_status catch-all - її дні
    НЕ стають "усім іншим" лише тому, що сама категорія прихована з рапорту."""
    categories = {
        100: {
            "general": [],
            "РТГр": {"grounds": [], "exclude_general": [], "include_to_report": False},
            "БД(СЗ)": {"grounds": [], "exclude_general": [], "default": True, "include_to_report": True},
        },
    }
    specs = mrh.resolve_reportable_categories(100, categories)

    assert [s["category_name"] for s in specs] == ["БД(СЗ)"]
    assert specs[0]["exclude_status"] == {"РТГр"}


def test_resolve_reportable_categories_hidden_catchall_omitted_entirely():
    categories = {
        100: {
            "general": [],
            "БД(СЗ)": {"grounds": [], "exclude_general": [], "default": True, "include_to_report": False},
        },
    }
    assert mrh.resolve_reportable_categories(100, categories) == []


def test_resolve_reportable_categories_no_catchall_returns_only_named():
    categories = {30: {"general": [], "ЗВРез": {"grounds": [], "exclude_general": [], "include_to_report": True}}}
    specs = mrh.resolve_reportable_categories(30, categories)
    assert [s["category_name"] for s in specs] == ["ЗВРез"]


def test_resolve_reportable_categories_missing_point_returns_empty():
    assert mrh.resolve_reportable_categories(999, {30: {}}) == []


def test_build_categories_uses_explicit_categories_independent_of_global(monkeypatch):
    """categories, переданий у _build_categories, повністю визначає підстави -
    підміна не залежить від того, що записано в глобальному MONEY_REPORT_CATEGORIES
    (так рапорт на командира/ТВО отримує НЕЗАЛЕЖНІ підстави від головного рапорту)."""
    import generators.generate_report_for_get_money as report_module

    monkeypatch.setattr(report_module, "MONEY_REPORT_CATEGORIES", {
        30: {
            "general": [],
            "ЖИТТЄДІЯЛЬНІСТЬ": {"grounds": [{"start": "01.07.2026", "end": "31.07.2026", "lines": ["З ГОЛОВНОГО"]}], "use_brs": False, "exclude_general": [], "default": True},
            "ЗВРез": {"grounds": [], "use_brs": False, "exclude_general": []},
        },
    })
    # _build_categories завжди звертається до всіх 4 пунктів (30/70/100/170) - тож
    # навіть коли тест цікавить лише 30, довелось додати мінімальну структуру й для
    # решти (лише "general" + catch-all "ОБОРОНА", без власних підстав).
    _minimal_point = {"general": [], "БД(СЗ)": {"grounds": [], "use_brs": False, "exclude_general": [], "default": True}}
    own_categories = {
        30: {
            "general": [],
            "ЖИТТЄДІЯЛЬНІСТЬ": {"grounds": [{"start": "01.07.2026", "end": "31.07.2026", "lines": ["З ОКРЕМОГО"]}], "use_brs": False, "exclude_general": [], "default": True},
            "ЗВРез": {"grounds": [], "use_brs": False, "exclude_general": []},
        },
        70: _minimal_point,
        100: _minimal_point,
        170: _minimal_point,
    }
    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30})]

    categories = report_module._build_categories(rows, [d1], {}, categories=own_categories)

    assert dict(categories)["30"][0]["ПІДСТАВА"] == "З ОКРЕМОГО"


def test_build_categories_thirty_point_tolerates_missing_named_category():
    """Пункт 30, як і 70к/100к/170к, не повинен ламатись, якщо переданий categories
    не містить ОДНІЄЇ з "типових" назв (напр. "ЗВРез" видалено з
    COMMANDER_MONEY_REPORT_CATEGORIES) - набір названих категорій для 30-го пункту
    вираховується динамічно з самого словника, а не захардкоджений. Жодних крос-
    пунктових винятків (напр. для "МЕДИК") тут більше немає - хто не потрапив у жодну
    названу категорію ЦЬОГО пункту, за замовчуванням потрапляє в catch-all, як і всі
    інші."""
    import generators.generate_report_for_get_money as report_module

    _minimal_point = {"general": [], "БД(СЗ)": {"grounds": [], "use_brs": False, "exclude_general": [], "default": True}}
    categories_without_zvrez = {
        30: {
            "general": [],
            "ЖИТТЄДІЯЛЬНІСТЬ": {"grounds": [], "use_brs": True, "exclude_general": [], "default": True},
        },
        70: _minimal_point,
        100: _minimal_point,
        170: _minimal_point,
    }
    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30}),
        _row("Перший Перший", "Бойовий медик", "Підрозділ 1", {d1: 30}),
    ]
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "МЕДИК"}

    categories = report_module._build_categories(rows, [d1], status_lookup, categories=categories_without_zvrez)

    thirty_rows = dict(categories)["30"]
    assert {r["ПІБ"] for r in thirty_rows} == {"Другий Другий", "Перший Перший"}


def test_build_categories_tolerates_missing_point_entirely():
    """Якщо ЦІЛИЙ пункт (напр. 70) видалено з переданого categories - жодна категорія
    цього пункту нікуди не потрапляє (ні в цей рапорт, ні деінде), без падіння; решта
    пунктів будуються як зазвичай."""
    import generators.generate_report_for_get_money as report_module

    _minimal_point = {"general": [], "БД(СЗ)": {"grounds": [], "use_brs": False, "exclude_general": [], "default": True}}
    categories_without_seventy = {
        30: {"general": [], "ЖИТТЄДІЯЛЬНІСТЬ": {"grounds": [], "use_brs": False, "exclude_general": [], "default": True}},
        100: _minimal_point,
        170: _minimal_point,
    }
    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 70})]

    categories = report_module._build_categories(rows, [d1], {}, categories=categories_without_seventy)

    assert "70" not in dict(categories)


def test_build_categories_thirty_point_recognizes_zvrez_when_present_in_custom_categories():
    """Симетрично попередньому тесту - якщо переданий categories[30] МІСТИТЬ "ЗВРез"
    (як і в MONEY_REPORT_CATEGORIES), вона й далі розпізнається як окрема названа
    категорія (власний status_filter, власна підстава), а не автоматично зливається
    з catch-all."""
    import generators.generate_report_for_get_money as report_module

    _minimal_point = {"general": [], "БД(СЗ)": {"grounds": [], "use_brs": False, "exclude_general": [], "default": True}}
    categories_with_zvrez = {
        30: {
            "general": [],
            "ЖИТТЄДІЯЛЬНІСТЬ": {"grounds": [], "use_brs": True, "exclude_general": [], "default": True},
            "ЗВРез": {"grounds": [{"start": "01.07.2026", "end": "31.07.2026", "lines": ["ОКРЕМА ПІДСТАВА ЗВРЕЗ"]}], "use_brs": False, "exclude_general": []},
        },
        70: _minimal_point,
        100: _minimal_point,
        170: _minimal_point,
    }
    d1 = datetime(2026, 7, 1)
    rows = [_row("Третій Третій", "Стрілець", "Підрозділ 1", {d1: 30})]
    status_lookup = {"ТРЕТІЙ ТРЕТІЙ": "ЗВРез"}

    categories = report_module._build_categories(rows, [d1], status_lookup, categories=categories_with_zvrez)

    thirty_rows = dict(categories)["30"]
    assert thirty_rows[0]["ПІДСТАВА"] == "ОКРЕМА ПІДСТАВА ЗВРЕЗ"


# -------------------------
# пункт 10к - тепер такий самий, як будь-який інший пункт: кожен день рахується
# незалежно, без жодної перевірки "чи є за місяць 30/70/100/170" - людина цілком
# нормально потрапляє одночасно і в 10_SECTION, і в інший пункт, якщо в різні дні
# місяця в неї різні значення.
# -------------------------
def test_build_categories_ten_section_includes_person_with_only_ten_days():
    import generators.generate_report_for_get_money as report_module

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 2)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 10, d2: 10})]

    categories = report_module._build_categories(rows, [d1, d2], {})

    ten_rows = dict(categories)["10"]
    assert [r["ПІБ"] for r in ten_rows] == ["Другий Другий"]
    assert ten_rows[0]["ДНІ"] == 2


def test_build_categories_ten_section_and_other_point_both_populated_for_same_person():
    """Людина з "10" одні дні й "100" інші - потрапляє ОБОМА рядками одразу: і в
    10_SECTION (за свої 10-дні), і в 100_SECTION (за свої 100-дні) - без жодного
    взаємовиключення між пунктами."""
    import generators.generate_report_for_get_money as report_module

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 2)
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {d1: 10, d2: 100})]

    categories = report_module._build_categories(rows, [d1, d2], {})
    by_key = dict(categories)

    ten_rows = by_key["10"]
    hundred_rows = by_key["100"]
    assert [r["ПІБ"] for r in ten_rows] == ["Перший Перший"]
    assert ten_rows[0]["ДНІ"] == 1
    assert [r["ПІБ"] for r in hundred_rows] == ["Перший Перший"]
    assert hundred_rows[0]["ДНІ"] == 1


def test_build_categories_omits_ten_section_when_nobody_qualifies():
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 100})]

    categories = report_module._build_categories(rows, [d1], {})

    assert "10" not in dict(categories)


# -------------------------
# oblik_disappearance_date (спільна з checker_accounting.report_checker - обидва
# написання колонки: "ДАТА ЗНИКНЕННЯ"/"ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ")
# -------------------------
def test_oblik_disappearance_date_accepts_either_column_name():
    assert mrh.oblik_disappearance_date({"ДАТА ЗНИКНЕННЯ": "11.08.2024"}) == "11.08.2024"
    assert mrh.oblik_disappearance_date({"ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ": "11.08.2024"}) == "11.08.2024"
    assert mrh.oblik_disappearance_date({"ПІБ": "х"}) is None


# -------------------------
# _disappearance_accrual_dates - статус "Очк.БЗ": період НЕ зі сканування
# комірок дат, а напряму від дати зникнення до останньої дати date_columns.
# -------------------------
def test_disappearance_accrual_dates_spans_from_disappearance_to_last_date_column():
    date_columns = [datetime(2026, 7, d) for d in range(1, 32)]
    row = {"ДАТА ЗНИКНЕННЯ": "29.07.2026"}

    dates = mrh._disappearance_accrual_dates(row, date_columns)

    assert dates == [datetime(2026, 7, 29), datetime(2026, 7, 30), datetime(2026, 7, 31)]


def test_disappearance_accrual_dates_spans_back_across_previous_months():
    """Головний сценарій, підтверджений користувачем: людина зникла ЗАДОВГО до
    поточного місяця генерації - період має тягнутись через УСІ місяці, що
    минули від моменту зникнення, а не лише обраний місяць."""
    date_columns = [datetime(2026, 7, d) for d in range(1, 32)]
    row = {"ДАТА ЗНИКНЕННЯ": "11.08.2024"}

    dates = mrh._disappearance_accrual_dates(row, date_columns)

    assert dates[0] == datetime(2024, 8, 11)
    assert dates[-1] == datetime(2026, 7, 31)
    assert len(dates) == (datetime(2026, 7, 31) - datetime(2024, 8, 11)).days + 1


def test_disappearance_accrual_dates_accepts_singular_column_name_variant():
    date_columns = [datetime(2026, 7, d) for d in range(1, 32)]
    row = {"ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ": "31.07.2026"}
    assert mrh._disappearance_accrual_dates(row, date_columns) == [datetime(2026, 7, 31)]


def test_disappearance_accrual_dates_empty_without_disappearance_column():
    assert mrh._disappearance_accrual_dates({}, [datetime(2026, 7, 1)]) == []


def test_disappearance_accrual_dates_empty_for_blank_disappearance_value():
    import numpy as np
    assert mrh._disappearance_accrual_dates({"ДАТА ЗНИКНЕННЯ": np.nan}, [datetime(2026, 7, 1)]) == []


def test_disappearance_accrual_dates_empty_for_unparseable_date():
    row = {"ДАТА ЗНИКНЕННЯ": "не дата"}
    assert mrh._disappearance_accrual_dates(row, [datetime(2026, 7, 1)]) == []


def test_disappearance_accrual_dates_empty_when_disappearance_date_after_last_date_column():
    date_columns = [datetime(2026, 7, d) for d in range(1, 32)]
    row = {"ДАТА ЗНИКНЕННЯ": "01.08.2026"}
    assert mrh._disappearance_accrual_dates(row, date_columns) == []


def test_disappearance_accrual_dates_empty_without_date_columns():
    assert mrh._disappearance_accrual_dates({"ДАТА ЗНИКНЕННЯ": "01.07.2026"}, []) == []


# -------------------------
# build_category_rows - категорія "Очк.БЗ" (_DISAPPEARANCE_ACCRUAL_CATEGORY)
# -------------------------
def test_build_category_rows_ochk_bz_uses_disappearance_date_span_not_day_cells():
    """Категорія "Очк.БЗ" НЕ сканує комірки дат узагалі (для таких людей їх
    просто не заповнюють, на відміну від решти категорій СПЕЦКОНТИНГЕНТУ) -
    період рахується напряму від дати зникнення до останньої дати
    date_columns (кінець обраного місяця генерації), підтверджено користувачем."""
    date_columns = [datetime(2026, 7, d) for d in range(1, 32)]
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {})]
    rows[0]["ДАТА ЗНИКНЕННЯ"] = "20.07.2026"
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "Очк.БЗ"}

    result = mrh.build_category_rows(
        rows, date_columns, {"100_СПЕЦКОНТИНГЕНТ"}, status_lookup,
        status_filter="Очк.БЗ", category_config=EMPTY_CATEGORY, category_name="Очк.БЗ",
    )

    assert len(result) == 1
    assert result[0]["ПІБ"] == "Перший Перший"
    assert result[0]["ПЕРІОД"] == "20.07.2026-31.07.2026"
    assert result[0]["ДНІ"] == 12


def test_build_category_rows_ochk_bz_spans_back_to_previous_months():
    date_columns = [datetime(2026, 7, d) for d in range(1, 32)]
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {})]
    rows[0]["ДАТА ЗНИКНЕННЯ"] = "11.08.2024"
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "Очк.БЗ"}

    result = mrh.build_category_rows(
        rows, date_columns, {"100_СПЕЦКОНТИНГЕНТ"}, status_lookup,
        status_filter="Очк.БЗ", category_config=EMPTY_CATEGORY, category_name="Очк.БЗ",
    )

    assert result[0]["ПЕРІОД"] == "11.08.2024-31.07.2026"


def test_build_category_rows_ochk_bz_skips_person_without_disappearance_date():
    date_columns = [datetime(2026, 7, d) for d in range(1, 32)]
    rows = [_row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {})]
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ ПЕРШИЙ": "Очк.БЗ"}

    result = mrh.build_category_rows(
        rows, date_columns, {"100_СПЕЦКОНТИНГЕНТ"}, status_lookup,
        status_filter="Очк.БЗ", category_config=EMPTY_CATEGORY, category_name="Очк.БЗ",
    )

    assert result == []


def test_build_category_rows_ochk_bz_skips_person_with_different_base_status():
    """Категорія "Очк.БЗ" рахується лише для людей, чий БАЗОВИЙ статус (колонка
    СТАТУС) - саме "Очк.БЗ", навіть якщо в них Є дата зникнення (напр. вони вже
    "полон", а дата зникнення лишилась в історії їхнього запису)."""
    date_columns = [datetime(2026, 7, d) for d in range(1, 32)]
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {})]
    rows[0]["ДАТА ЗНИКНЕННЯ"] = "01.07.2026"
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "полон"}

    result = mrh.build_category_rows(
        rows, date_columns, {"100_СПЕЦКОНТИНГЕНТ"}, status_lookup,
        status_filter="Очк.БЗ", category_config=EMPTY_CATEGORY, category_name="Очк.БЗ",
    )

    assert result == []


def test_build_category_rows_ochk_bz_applies_extra_grounds_by_person():
    """"ПІДСТАВА" для "Очк.БЗ" будується тим самим шляхом (build_legal_basis_text +
    extra_grounds_by_person), що й для решти категорій - жодної спеціальної
    логіки для тексту підстави, лише для самих ДАТ участі."""
    date_columns = [datetime(2026, 7, d) for d in range(1, 32)]
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {})]
    rows[0]["ДАТА ЗНИКНЕННЯ"] = "20.07.2026"
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "Очк.БЗ"}
    extra_grounds_by_person = {"ПЕРШИЙ ПЕРШИЙ": ["Наказ №1164-ОД від 30.09.2024"]}

    result = mrh.build_category_rows(
        rows, date_columns, {"100_СПЕЦКОНТИНГЕНТ"}, status_lookup,
        status_filter="Очк.БЗ", category_config=EMPTY_CATEGORY, category_name="Очк.БЗ",
        extra_grounds_by_person=extra_grounds_by_person,
    )

    assert result[0]["ПІДСТАВА"] == "Наказ №1164-ОД від 30.09.2024"


# -------------------------
# resolve_status_category / resolve_reportable_categories - "Очк.БЗ" в
# реальному constants.MONEY_REPORT_CATEGORIES["100_СПЕЦКОНТИНГЕНТ"]
# -------------------------
def test_resolve_status_category_recognizes_ochk_bz():
    assert mrh.resolve_status_category("Очк.БЗ") == "Очк.БЗ"
    assert mrh.resolve_status_category("очк.бз") == "Очк.БЗ"


def test_resolve_reportable_categories_includes_ochk_bz_for_spetskontyngent_point():
    specs = mrh.resolve_reportable_categories("100_СПЕЦКОНТИНГЕНТ", constants.MONEY_REPORT_CATEGORIES)
    assert {spec["category_name"] for spec in specs} >= {"полон", "інд", "Очк.БЗ"}


# -------------------------
# _build_categories - "100_СПЕЦКОНТИНГЕНТ" вимагає непорожню колонку "ПІДСТАВИ"
# -------------------------
def test_build_categories_spetskontyngent_requires_nonblank_pidstavy_column():
    """Користувач підтвердив: секція "100_СПЕЦКОНТИНГЕНТ" МАЄ включати людину в
    рапорт ЛИШЕ якщо колонка "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ" ОБЛІК.xlsx (ВЛАСНА для
    цього статусу - BASIS_REQUIRED_POINT_COLUMN_NAMES, constants.py, а НЕ
    спільна "ПІДСТАВИ") для неї непорожня - на відміну від решти пунктів
    (test_build_categories_pidstavy_column_applies_to_every_point), де ця
    колонка лише ДОПОВНЮЄ вже побудовану підставу, не впливаючи на саме
    включення людини в рапорт."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [
        _row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: "100_СПЕЦКОНТИНГЕНТ"}),
        _row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: "100_СПЕЦКОНТИНГЕНТ"}),
    ]
    rows[0]["ПІДСТАВИ СПЕЦКОНТИНГЕНТУ"] = "Наказ №1 від 01.07.2026"
    rows[0]["ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ"] = "01.07.2026"
    rows[1]["ПІДСТАВИ СПЕЦКОНТИНГЕНТУ"] = ""
    rows[1]["ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ"] = "01.07.2026"
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ ПЕРШИЙ": "полон", "ДРУГИЙ ДРУГИЙ ДРУГИЙ": "полон"}

    categories = report_module._build_categories(rows, [d1], status_lookup)

    section = dict(categories).get("100_СПЕЦКОНТИНГЕНТ", [])
    assert {r["ПІБ"] for r in section} == {"Перший Перший Перший"}


def test_build_categories_spetskontyngent_excludes_person_without_pidstavy_column_at_all():
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Третій Третій Третій", "Стрілець", "Підрозділ 1", {d1: "100_СПЕЦКОНТИНГЕНТ"})]
    status_lookup = {"ТРЕТІЙ ТРЕТІЙ ТРЕТІЙ": "полон"}

    categories = report_module._build_categories(rows, [d1], status_lookup)

    assert "100_СПЕЦКОНТИНГЕНТ" not in dict(categories)


def test_build_categories_spetskontyngent_section_absent_when_everyone_lacks_basis():
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: "100_СПЕЦКОНТИНГЕНТ"})]
    status_lookup = {"ДРУГИЙ ДРУГИЙ ДРУГИЙ": "полон"}

    categories = report_module._build_categories(rows, [d1], status_lookup)

    assert "100_СПЕЦКОНТИНГЕНТ" not in dict(categories)


def test_build_categories_other_points_unaffected_by_pidstavy_gate():
    """Гейт "непорожня ПІДСТАВИ" стосується ЛИШЕ "100_СПЕЦКОНТИНГЕНТ" - звичайні
    пункти (тут - 30) і далі включають людину без жодної колонки "ПІДСТАВИ"."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 7, 1)
    rows = [_row("Другий Другий", "Стрілець", "Підрозділ 1", {d1: 30})]

    categories = report_module._build_categories(rows, [d1], {})

    assert [r["ПІБ"] for r in dict(categories)["30"]] == ["Другий Другий"]


def test_build_categories_ochk_bz_appears_in_spetskontyngent_section_with_full_period():
    """Наскрізь: людина зі статусом "Очк.БЗ" і непорожньою "ПІДСТАВИ" - потрапляє
    в "100_СПЕЦКОНТИНГЕНТ" з періодом, що тягнеться від дати зникнення (а не
    лише обраний місяць генерації)."""
    import generators.generate_report_for_get_money as report_module

    date_columns = [datetime(2026, 7, d) for d in range(1, 32)]
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {})]
    rows[0]["ДАТА ЗНИКНЕННЯ"] = "11.08.2024"
    rows[0]["ПІДСТАВИ СПЕЦКОНТИНГЕНТУ"] = "Наказ №1164-ОД від 30.09.2024"
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "Очк.БЗ"}

    categories = report_module._build_categories(rows, date_columns, status_lookup)

    section = dict(categories)["100_СПЕЦКОНТИНГЕНТ"]
    assert len(section) == 1
    assert section[0]["ПЕРІОД"] == "11.08.2024-31.07.2026"
    assert section[0]["ПІДСТАВА"] == "Наказ №1164-ОД від 30.09.2024"


def test_build_categories_ochk_bz_excluded_without_pidstavy():
    import generators.generate_report_for_get_money as report_module

    date_columns = [datetime(2026, 7, d) for d in range(1, 32)]
    rows = [_row("Перший Перший", "Стрілець", "Підрозділ 1", {})]
    rows[0]["ДАТА ЗНИКНЕННЯ"] = "11.08.2024"
    status_lookup = {"ПЕРШИЙ ПЕРШИЙ": "Очк.БЗ"}

    categories = report_module._build_categories(rows, date_columns, status_lookup)

    assert "100_СПЕЦКОНТИНГЕНТ" not in dict(categories)
