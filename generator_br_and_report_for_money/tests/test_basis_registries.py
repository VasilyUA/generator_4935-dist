import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import numpy as np

import constants
import content.money_report_helpers as mrh


def _row(pib, posada, pidrozdil, days, zvannya="сержант"):
    row = {"ПІБ": pib, "ПОСАДА": posada, "ПІДРОЗДІЛ": pidrozdil, "ЗВАННЯ": zvannya}
    row.update(days)
    return row


def _basis_column_names_declared_in_registries():
    names = set(constants.GENERAL_PIDSTAVY_COLUMN_NAMES)
    for column_names in constants.BASIS_REQUIRED_POINT_COLUMN_NAMES.values():
        names.update(column_names)
    for column_names in constants.POINT_OWN_PIDSTAVY_COLUMN_NAMES.values():
        names.update((column_names,) if isinstance(column_names, str) else column_names)
    return names


# -------------------------
# Реєстри підстав узгоджені між собою
# -------------------------
def test_basis_required_points_match_their_column_registry():
    """Кожен пункт, що вимагає підставу, має свою колонку (і навпаки) - інакше
    пункт або ігнорує свою колонку, або має колонку, яку ніхто не читає."""
    assert set(constants.BASIS_REQUIRED_POINTS) == set(constants.BASIS_REQUIRED_POINT_COLUMN_NAMES)


def test_basis_required_points_exist_as_money_report_categories():
    assert set(constants.BASIS_REQUIRED_POINTS) <= set(constants.MONEY_REPORT_CATEGORIES)


def test_every_basis_column_name_is_read_from_the_file():
    """РЕГРЕСІЯ (двічі підтверджено користувачем: "ПІДСТАВИ 10К", "ПІДСТАВИ
    ШПБПДЛ"/"БПШПДЛ"): назва колонки, оголошена в реєстрі підстав, але ВІДСУТНЯ
    в OPTIONAL_PERSONEL_COLUMN_NAMES, мовчки не читається з ОБЛІК.xlsx - клітинка
    завжди порожня, хоч логіка її використання правильна."""
    assert _basis_column_names_declared_in_registries() <= set(constants.OPTIONAL_PERSONEL_COLUMN_NAMES)


# -------------------------
# 100_БПШПДЛ - та сама логіка, що й 100_БПШП (власна колонка, гейт за підставою)
# -------------------------
def test_bpshpdl_point_is_basis_required():
    assert "100_БПШПДЛ" in constants.BASIS_REQUIRED_POINTS


def test_bpshpdl_takes_basis_text_from_its_own_column():
    """Саме цей випадок користувач бачив: людина з підставою в колонці
    "ПІДСТАВИ БПШПДЛ" отримувала порожню "Підстава для виплати" і помилку
    "Немає підстави для виплати"."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 9, 1)
    rows = [_row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: "БПШПДЛ"})]
    rows[0]["ПІДСТАВИ БПШПДЛ"] = "Висновок ВЛК №7 від 01.09.2026"

    categories = report_module._build_categories(rows, [d1], {})

    section = dict(categories).get("100_БПШПДЛ", [])
    assert [r["ПІБ"] for r in section] == ["Перший Перший Перший"]
    assert "Висновок ВЛК №7 від 01.09.2026" in section[0]["ПІДСТАВА"]


def test_bpshpdl_accepts_shpbpdl_spelling():
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 9, 1)
    rows = [_row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: "БПШПДЛ"})]
    rows[0]["ПІДСТАВИ ШПБПДЛ"] = "Висновок ВЛК №8 від 01.09.2026"

    categories = report_module._build_categories(rows, [d1], {})

    section = dict(categories).get("100_БПШПДЛ", [])
    assert "Висновок ВЛК №8 від 01.09.2026" in section[0]["ПІДСТАВА"]


def test_bpshpdl_requires_nonblank_basis_column():
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 9, 1)
    rows = [
        _row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: "БПШПДЛ"}),
        _row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: "БПШПДЛ"}),
    ]
    rows[0]["ПІДСТАВИ БПШПДЛ"] = "Висновок ВЛК №7 від 01.09.2026"
    rows[1]["ПІДСТАВИ БПШПДЛ"] = ""

    categories = report_module._build_categories(rows, [d1], {})

    section = dict(categories).get("100_БПШПДЛ", [])
    assert {r["ПІБ"] for r in section} == {"Перший Перший Перший"}


def test_bpshpdl_person_without_basis_is_excluded_with_warning(capsys):
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 9, 1)
    rows = [_row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: "БПШПДЛ"})]

    report_module._build_categories(rows, [d1], {})

    printed = capsys.readouterr().out
    assert "Немає підстави для 100_БПШПДЛ: Другий Другий Другий" in printed
    assert "виключено з рапорту" in printed


# -------------------------
# build_pidstavy_extra_grounds_by_person - одиночний рядок як одна назва колонки
# -------------------------
def test_build_pidstavy_extra_grounds_by_person_treats_bare_string_as_one_column_name():
    rows = [{"ПІБ": "Перший Перший", "ПІДСТАВИ 10К": "Наказ №7 від 01.07.2026"}]

    assert mrh.build_pidstavy_extra_grounds_by_person(rows, "ПІДСТАВИ 10К") == {
        "ПЕРШИЙ ПЕРШИЙ": ["Наказ №7 від 01.07.2026"],
    }


# -------------------------
# _row_szch_whole_month - порожні/NaN комірки (гілки перевірки)
# -------------------------
def test_row_szch_whole_month_false_for_none_and_nan_cells():
    d1, d2 = datetime(2026, 9, 1), datetime(2026, 9, 2)

    assert mrh._row_szch_whole_month({d1: "СЗЧ", d2: None}, [d1, d2]) is False
    assert mrh._row_szch_whole_month({d1: "СЗЧ", d2: np.nan}, [d1, d2]) is False
