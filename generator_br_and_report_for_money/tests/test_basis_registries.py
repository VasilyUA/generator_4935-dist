import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import numpy as np
import pytest

import constants
import content.money_report_helpers as mrh
import content.report_changes as rc


def _row(pib, posada, pidrozdil, days, zvannya="сержант"):
    row = {"ПІБ": pib, "ПОСАДА": posada, "ПІДРОЗДІЛ": pidrozdil, "ЗВАННЯ": zvannya}
    row.update(days)
    return row


def _as_column_names(value):
    """Значення реєстру - кортеж назв АБО одиночний рядок (пропущена кома:
    ("ПІДСТАВИ 10К") - це рядок, а не кортеж). Ітерація по рядку дала б літери,
    тож розгортаємо його в одну назву."""
    return (value,) if isinstance(value, str) else tuple(value)


def _basis_column_names_declared_in_registries():
    names = set(constants.GENERAL_PIDSTAVY_COLUMN_NAMES)
    for column_names in constants.BASIS_REQUIRED_POINT_COLUMN_NAMES.values():
        names.update(_as_column_names(column_names))
    for column_names in constants.POINT_OWN_PIDSTAVY_COLUMN_NAMES.values():
        names.update(_as_column_names(column_names))
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


def test_read_allowlist_holds_whole_column_names_not_characters():
    """РЕГРЕСІЯ (2026-10-04: КУРГУЗ/МІРОШНІЧЕНКО випадали з рапорту й зі змін):
    значення реєстру, записане рядком без коми, розгорталось у ПОСИМВОЛЬНИЙ список
    ("П", "І", "Д", ...) - справжня назва колонки не потрапляла в
    OPTIONAL_PERSONEL_COLUMN_NAMES, і колонка мовчки не читалась з ОБЛІК.xlsx.
    Тест вище цього не ловив: він теж ітерував по рядку й порівнював літери самі з собою."""
    assert all(len(name) > 1 for name in constants.OPTIONAL_PERSONEL_COLUMN_NAMES)


def test_changes_label_columns_hold_whole_basis_column_names():
    """Той самий клас помилки для _PIDSTAVY_LABEL_COLUMNS (content.report_changes):
    без повної назви _read_check_file не збирає підставу з prev/actual файлів, і
    зміни для людини з новою підставою (напр. 100_СПЕЦКОНТИНГЕНТ) не з'являються."""
    basis_names = set(constants.GENERAL_PIDSTAVY_COLUMN_NAMES)
    for column_names in constants.BASIS_REQUIRED_POINT_COLUMN_NAMES.values():
        basis_names.update(_as_column_names(column_names))
    assert basis_names <= set(rc._PIDSTAVY_LABEL_COLUMNS)
    assert all(len(name) > 1 for name in rc._PIDSTAVY_LABEL_COLUMNS)


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


# -------------------------
# Реєстри приймають БУДЬ-ЯКУ форму запису (utils/column_names.py) - редагуйте словники довільно
# -------------------------
@pytest.mark.parametrize("value", [
    "ПІДСТАВИ 10К",
    ("ПІДСТАВИ 10К"),
    ("ПІДСТАВИ 10К",),
    ["ПІДСТАВИ 10К"],
    {"ПІДСТАВИ 10К"},
    "підстави  10к",
])
def test_column_names_of_accepts_every_registry_form(value):
    from utils.column_names import column_names_of

    assert column_names_of(value) == ("ПІДСТАВИ 10К",)


def test_column_names_of_treats_a_non_iterable_value_as_one_name():
    from utils.column_names import column_names_of

    assert column_names_of(10) == ("10",)


def test_column_names_of_joins_several_names_and_drops_duplicates_keeping_order():
    from utils.column_names import column_names_of

    assert column_names_of(("ПІДСТАВИ ШПБП", "підстави  шпбп", "ПІДСТАВИ БПШП")) == ("ПІДСТАВИ ШПБП", "ПІДСТАВИ БПШП")


def test_column_names_of_treats_empty_values_as_no_names():
    from utils.column_names import column_names_of

    assert column_names_of(None) == ()
    assert column_names_of("") == ()
    assert column_names_of(("", "   ", None)) == ()


def test_basis_source_for_point_routes_by_registry_keys():
    """Пункт обирається за КЛЮЧЕМ реєстру: BASIS (власна колонка з виключенням), OWN (власна
    без виключення) або загальні колонки - у будь-якій формі значення."""
    basis = {"10_МЕД": "ПІДСТАВИ МЕДИЧНІ НА 10К"}
    own = {10: ("ПІДСТАВИ 10К",)}

    assert mrh.basis_source_for_point("10_МЕД", basis, own) == ("basis", ("ПІДСТАВИ МЕДИЧНІ НА 10К",))
    assert mrh.basis_source_for_point(10, basis, own) == ("own", ("ПІДСТАВИ 10К",))
    assert mrh.basis_source_for_point(30, basis, own) == ("general", ())


def test_basis_source_for_point_prefers_basis_when_point_is_in_both_registries():
    basis = {"10_МЕД": ("ПІДСТАВИ МЕДИЧНІ НА 10К",)}
    own = {"10_МЕД": ("ПІДСТАВИ 10К",)}

    assert mrh.basis_source_for_point("10_МЕД", basis, own)[0] == "basis"


def test_basis_source_for_point_rejects_registry_value_without_any_column_name():
    """Порожнє значення - помилка з назвою пункту, а не мовчазне виключення всіх людей з пункту."""
    with pytest.raises(ValueError, match="10_МЕД"):
        mrh.basis_source_for_point("10_МЕД", {"10_МЕД": ""}, {})


def test_basis_source_for_point_reads_live_registry_when_no_registries_given(monkeypatch):
    monkeypatch.setitem(mrh.BASIS_REQUIRED_POINT_COLUMN_NAMES, "ТЕСТОВИЙ_ПУНКТ", "ПІДСТАВИ ТЕСТ")

    assert mrh.basis_source_for_point("ТЕСТОВИЙ_ПУНКТ") == ("basis", ("ПІДСТАВИ ТЕСТ",))


def test_build_pidstavy_extra_grounds_finds_column_in_any_case_and_spacing():
    rows = [{"ПІБ": "Перший Перший", "ПІДСТАВИ МЕДИЧНІ НА 10К": "Висновок №7"}]

    assert mrh.build_pidstavy_extra_grounds_by_person(rows, ("підстави  медичні на 10к",)) == {
        "ПЕРШИЙ ПЕРШИЙ": ["Висновок №7"],
    }


def test_detect_personel_list_columns_letters_matches_header_in_any_case_and_spacing(tmp_path):
    """Заголовок ОБЛІК.xlsx з іншим регістром і пробілами (з переносом рядка) знаходить колонку
    з реєстру, і сам запис у реєстрі теж може бути в будь-якому регістрі - обидва варіанти."""
    import utils.excel_reader as excel_reader_module
    from datetime import datetime as dt
    from openpyxl import Workbook

    path = tmp_path / "oblik.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Аркуш1"
    ws.append(["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", dt(2026, 7, 1), "Підстави  медичні\nна 10к"])
    wb.save(str(path))

    for registry_value in ("ПІДСТАВИ МЕДИЧНІ НА 10К", "підстави медичні на 10к"):
        letters = excel_reader_module.detect_personel_list_columns_letters(
            str(path), "Аркуш1", "07", "2026", optional_column_names=(registry_value,),
        )
        assert letters == ["A", "B", "C", "D", "E", "F"]


def test_build_categories_reads_basis_from_a_registry_value_written_as_a_bare_string(capsys):
    """Пункт, чию власну колонку записано РЯДКОМ (без коми), бере підставу саме з неї; людина
    без неї виключається з пункту з попередженням - так само, як для будь-якого пункту, чий
    ключ є в BASIS_REQUIRED_POINT_COLUMN_NAMES. Реєстр передається явно (без підміни модуля)."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 9, 1)
    with_basis = _row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: "БПШП"})
    with_basis["ПІДСТАВИ ТЕСТ"] = "Висновок ВЛК №9"
    without_basis = _row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: "БПШП"})

    categories = report_module._build_categories(
        [with_basis, without_basis], [d1], {}, basis_registry={"100_БПШП": "ПІДСТАВИ ТЕСТ"},
    )

    section = dict(categories).get("100_БПШП", [])
    assert [r["ПІБ"] for r in section] == ["Перший Перший Перший"]
    assert "Висновок ВЛК №9" in section[0]["ПІДСТАВА"]
    assert "Немає підстави для 100_БПШП: Другий Другий Другий" in capsys.readouterr().out
