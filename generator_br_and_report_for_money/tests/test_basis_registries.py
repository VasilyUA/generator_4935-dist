import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import numpy as np
import pytest

import constants
import content.money_report_helpers as mrh
import content.report_changes as rc
from utils.column_names import collect_column_basis_names, column_names_of


def _row(pib, posada, pidrozdil, days, zvannya="сержант"):
    row = {"ПІБ": pib, "ПОСАДА": posada, "ПІДРОЗДІЛ": pidrozdil, "ЗВАННЯ": zvannya}
    row.update(days)
    return row


def _basis_column_names_declared_in_categories():
    """Усі назви колонок підстави, фактично дописані в "column_basis_from_tracking"
    якоїсь названої категорії реального constants.py (обидва словники -
    MONEY_REPORT_CATEGORIES, COMMANDER_MONEY_REPORT_CATEGORIES)."""
    names = set(collect_column_basis_names(constants.MONEY_REPORT_CATEGORIES))
    names.update(collect_column_basis_names(constants.COMMANDER_MONEY_REPORT_CATEGORIES))
    return names


# -------------------------
# basis_required_points/column_basis_from_tracking - узгоджені з реальним constants.py
# -------------------------
def test_basis_required_points_match_live_constants_categories():
    """basis_required_points(categories) - похідний набір, а не окремий реєстр: пункт
    потрапляє туди РІВНО тоді, коли ХОЧ ОДНА його названа категорія задала (своє, або
    успадкувала None з рівня пункту - як "NOT_PAID") непорожнє "column_basis_from_tracking"
    з required_basis=True."""
    expected = set()
    for point, point_categories in constants.MONEY_REPORT_CATEGORIES.items():
        point_level_tracking = point_categories.get("column_basis_from_tracking")
        for category_name, config in point_categories.items():
            if category_name == "general" or not isinstance(config, dict):
                continue
            raw_value = config.get("column_basis_from_tracking")
            if raw_value is None:
                raw_value = point_level_tracking
            if raw_value is not None and config.get("required_basis", True):
                expected.add(point)
                break
    assert mrh.basis_required_points(constants.MONEY_REPORT_CATEGORIES) == expected
    assert expected <= set(constants.MONEY_REPORT_CATEGORIES)


@pytest.mark.parametrize("categories, label", [
    (constants.MONEY_REPORT_CATEGORIES, "MONEY_REPORT_CATEGORIES"),
    (constants.COMMANDER_MONEY_REPORT_CATEGORIES, "COMMANDER_MONEY_REPORT_CATEGORIES"),
])
def test_every_named_category_has_the_column_basis_from_tracking_key(categories, label):
    """Рішення користувача: поле "column_basis_from_tracking" МАЄ бути присутнє в
    кожній названій категорії явно (None, якщо своєї колонки немає - у пунктів зі
    спільним значенням на рівні пункту, напр. "NOT_PAID", None тут означає
    "успадкувати його", а не "точно без колонки") - видимість поля для кожного,
    хто читає constants.py, окрема, явно поставлена вимога, незалежно від того,
    чи є в пункту своє "column_basis_from_tracking"."""
    missing = [
        (point, category_name)
        for point, point_categories in categories.items()
        for category_name, config in point_categories.items()
        if category_name != "general" and isinstance(config, dict) and "column_basis_from_tracking" not in config
    ]
    assert missing == [], f"{label}: бракує column_basis_from_tracking у {missing}"


def test_every_basis_column_name_is_read_from_the_file():
    """РЕГРЕСІЯ (двічі підтверджено користувачем: "ПІДСТАВИ 10К", "ПІДСТАВИ
    ШПБПДЛ"/"БПШПДЛ"): назва колонки, дописана в "column_basis_from_tracking" якоїсь
    категорії, але ВІДСУТНЯ в OPTIONAL_PERSONEL_COLUMN_NAMES, мовчки не читається з
    ОБЛІК.xlsx - клітинка завжди порожня, хоч логіка її використання правильна."""
    assert _basis_column_names_declared_in_categories() <= set(constants.OPTIONAL_PERSONEL_COLUMN_NAMES)


def test_read_allowlist_holds_whole_column_names_not_characters():
    """РЕГРЕСІЯ (2026-10-04: КУРГУЗ/МІРОШНІЧЕНКО випадали з рапорту й зі змін):
    значення поля, записане рядком без коми, розгорталось у ПОСИМВОЛЬНИЙ список
    ("П", "І", "Д", ...) - справжня назва колонки не потрапляла в
    OPTIONAL_PERSONEL_COLUMN_NAMES, і колонка мовчки не читалась з ОБЛІК.xlsx.
    Тест вище цього не ловив: він теж ітерував по рядку й порівнював літери самі з собою."""
    assert all(len(name) > 1 for name in constants.OPTIONAL_PERSONEL_COLUMN_NAMES)


def test_changes_label_columns_hold_whole_basis_column_names():
    """Той самий клас помилки для _PIDSTAVY_LABEL_COLUMNS (content.report_changes):
    без повної назви _read_check_file не збирає підставу з prev/actual файлів, і
    зміни для людини з новою підставою (напр. 100_СПЕЦКОНТИНГЕНТ) не з'являються."""
    assert _basis_column_names_declared_in_categories() <= set(rc._PIDSTAVY_LABEL_COLUMNS)
    assert all(len(name) > 1 for name in rc._PIDSTAVY_LABEL_COLUMNS)


def test_collect_column_basis_names_skips_general_and_none_and_dedupes():
    categories = {
        10: {
            "general": [],
            "А": {"column_basis_from_tracking": "ПІДСТАВИ 10К"},
            "Б": {"column_basis_from_tracking": None},
        },
        30: {
            "general": [],
            "В": {"column_basis_from_tracking": ("ПІДСТАВИ 10К", "ІНША")},
        },
    }

    assert collect_column_basis_names(categories) == ("ПІДСТАВИ 10К", "ІНША")


# -------------------------
# Пункт з власною обов'язковою підставою (та сама логіка, що й у реальних
# "100_БПШП"/"100_ДТЛ"/"100_СПЕЦКОНТИНГЕНТ" - required_basis=True зі своєю
# column_basis_from_tracking) - на ВИГАДАНІЙ категорії, переданій явно через
# categories=, а не на реальному записі з live constants.py. РЕГРЕСІЯ
# (виявлено користувачем на реальному прикладі): раніше ці тести перевіряли
# РІВНО це саме на пункті "100_БПШПДЛ" із constants.MONEY_REPORT_CATEGORIES за
# замовчуванням - коли користувач перейменував його на "100_ДТЛ" (реальний
# статус і справді міняється з часом), тести миттєво зламались, хоча сама
# логіка не постраждала. Тест на ЛОГІКУ не повинен залежати від того, як САМЕ
# зараз називається реальний статус.
#
# Пункт тут - число 10 (а не вигаданий рядок, як категорія) НЕ довільне:
# resolve_day_value_and_category/_resolve_category_point_and_name
# (content/money_report_helpers.py) звіряють ЧИСЛОВЕ значення комірки дня з
# ГЛОБАЛЬНИМ constants.MONEY_REPORT_CATEGORIES напряму (а не з categories,
# переданим у _build_categories) - вигаданий ЧИСЛОВИЙ пункт (напр. 9999) чи
# вигадане ТЕКСТОВЕ значення дня (напр. "ТЕСТ") там ніколи не знайдеться, і
# рядок людини мовчки не потрапив би НІКУДИ. 10/30/70/100/170 - реальні,
# структурні (не статусні, тому стабільні) ключі глобального реєстру, тож
# лишаються єдиним безпечним "носієм" дня для вигаданої категорії; категорія
# під ним ("ТЕСТ") позначена "default": True, щоб чисте число 10 (без жодного
# статусу) потрапило САМЕ в неї (catch-all), а не загубилось.
# -------------------------
def _own_categories_with_required_basis():
    return {
        10: {
            "general": [],
            "ТЕСТ": {
                "grounds": [],
                "use_brs": False,
                "use_brs_from_selected_folder": False,
                "exclude_general": [],
                "default": True,
                "generate_br": False,
                "include_to_report": True,
                "required_basis": True,
                "column_basis_from_tracking": ("ПІДСТАВИ А", "ПІДСТАВИ Б"),
            },
        },
    }


def test_point_with_required_basis_column_is_in_basis_required_points():
    assert 10 in mrh.basis_required_points(_own_categories_with_required_basis())


def test_category_takes_basis_text_from_its_own_column():
    """Той самий випадок, що й користувач колись бачив на реальному статусі: людина
    з підставою у своїй колонці отримувала порожню "Підстава для виплати" і
    помилку "Немає підстави для виплати"."""
    import generators.generate_report_for_get_money as report_module

    own_categories = _own_categories_with_required_basis()
    d1 = datetime(2026, 9, 1)
    rows = [_row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: 10})]
    rows[0]["ПІДСТАВИ А"] = "Висновок ВЛК №7 від 01.09.2026"

    categories = report_module._build_categories(rows, [d1], {}, categories=own_categories)

    section = dict(categories).get("10", [])
    assert [r["ПІБ"] for r in section] == ["Перший Перший Перший"]
    assert "Висновок ВЛК №7 від 01.09.2026" in section[0]["ПІДСТАВА"]


def test_category_accepts_alternate_basis_column_spelling():
    import generators.generate_report_for_get_money as report_module

    own_categories = _own_categories_with_required_basis()
    d1 = datetime(2026, 9, 1)
    rows = [_row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: 10})]
    rows[0]["ПІДСТАВИ Б"] = "Висновок ВЛК №8 від 01.09.2026"

    categories = report_module._build_categories(rows, [d1], {}, categories=own_categories)

    section = dict(categories).get("10", [])
    assert "Висновок ВЛК №8 від 01.09.2026" in section[0]["ПІДСТАВА"]


def test_category_requires_nonblank_basis_column():
    import generators.generate_report_for_get_money as report_module

    own_categories = _own_categories_with_required_basis()
    d1 = datetime(2026, 9, 1)
    rows = [
        _row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: 10}),
        _row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: 10}),
    ]
    rows[0]["ПІДСТАВИ А"] = "Висновок ВЛК №7 від 01.09.2026"
    rows[1]["ПІДСТАВИ А"] = ""

    categories = report_module._build_categories(rows, [d1], {}, categories=own_categories)

    section = dict(categories).get("10", [])
    assert {r["ПІБ"] for r in section} == {"Перший Перший Перший"}


def test_category_person_without_basis_is_excluded_with_warning(capsys):
    import generators.generate_report_for_get_money as report_module

    own_categories = _own_categories_with_required_basis()
    d1 = datetime(2026, 9, 1)
    rows = [_row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: 10})]

    report_module._build_categories(rows, [d1], {}, categories=own_categories)

    printed = capsys.readouterr().out
    assert "Немає підстави для 10: Другий Другий Другий" in printed
    assert "виключено з рапорту" in printed


# -------------------------
# NOT_PAID (СЗЧ/Задув) - власна колонка "ПІДСТАВИ ДЛЯ НЕ ВИПЛАТИ", та сама для ОБОХ
# -------------------------
def test_not_paid_statuses_share_the_same_basis_column():
    """СЗЧ і "Задув.|задув|задут" - рішення користувача: ОДНА й та сама колонка
    "ПІДСТАВИ ДЛЯ НЕ ВИПЛАТИ" для обох - повторення назви між різними категоріями
    того самого пункту дозволене й очікуване, не конфлікт."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 9, 1)
    rows = [
        _row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: "СЗЧ"}),
        _row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: "задув"}),
    ]
    rows[0]["ПІДСТАВИ ДЛЯ НЕ ВИПЛАТИ"] = "Рапорт від 01.09.2026"
    rows[1]["ПІДСТАВИ ДЛЯ НЕ ВИПЛАТИ"] = "Акт від 01.09.2026"

    categories = report_module._build_categories(rows, [d1], {})

    not_paid = {r["ПІБ"]: r["ПІДСТАВА"] for r in dict(categories).get("NOT_PAID", [])}
    assert "Рапорт від 01.09.2026" in not_paid["Перший Перший Перший"]
    assert "Акт від 01.09.2026" in not_paid["Другий Другий Другий"]


def test_not_paid_person_without_basis_is_excluded_with_warning(capsys):
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 9, 1)
    rows = [_row("Третій Третій Третій", "Стрілець", "Підрозділ 1", {d1: "СЗЧ"})]

    categories = report_module._build_categories(rows, [d1], {})

    assert "Третій Третій Третій" not in {r["ПІБ"] for r in dict(categories).get("NOT_PAID", [])}
    printed = capsys.readouterr().out
    assert "Немає підстави для NOT_PAID: Третій Третій Третій" in printed


def test_build_categories_raises_when_file_has_both_alternate_basis_spellings_at_once():
    """Рішення користувача: якщо в ОБЛІК.xlsx одночасно присутні ОБИДВІ альтернативні
    назви колонки однієї підстави - помилка, а не мовчазний вибір першої за
    порядком (вигаданий пункт/категорія - та сама логіка, що й у реальних
    "100_БПШП"/"100_ДТЛ", див. коментар над _own_categories_with_required_basis)."""
    import generators.generate_report_for_get_money as report_module

    own_categories = _own_categories_with_required_basis()
    d1 = datetime(2026, 9, 1)
    rows = [_row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: 10})]
    rows[0]["ПІДСТАВИ А"] = "Висновок А"
    rows[0]["ПІДСТАВИ Б"] = "Висновок Б"

    with pytest.raises(ValueError, match="ПІДСТАВИ"):
        report_module._build_categories(rows, [d1], {}, categories=own_categories)


# -------------------------
# build_pidstavy_extra_grounds_by_person - одиночний рядок як одна назва колонки;
# помилка, якщо в файлі одночасно є кілька альтернативних написань
# -------------------------
def test_build_pidstavy_extra_grounds_by_person_treats_bare_string_as_one_column_name():
    rows = [{"ПІБ": "Перший Перший", "ПІДСТАВИ 10К": "Наказ №7 від 01.07.2026"}]

    assert mrh.build_pidstavy_extra_grounds_by_person(rows, "ПІДСТАВИ 10К") == {
        "ПЕРШИЙ ПЕРШИЙ": ["Наказ №7 від 01.07.2026"],
    }


def test_build_pidstavy_extra_grounds_by_person_rejects_both_alternate_columns_present():
    """Будь-який виклик із кількома альтернативними назвами (не лише через
    column_basis_from_tracking) - обидві присутні одночасно в файлі - помилка,
    безумовно (рішення користувача: жодного "м'якого" режиму більше немає)."""
    rows = [{
        "ПІБ": "Перший Перший",
        "ПІДСТАВИ ШПБПДЛ": "Висновок А",
        "ПІДСТАВИ БПШПДЛ": "Висновок Б",
    }]

    with pytest.raises(ValueError, match="ПІДСТАВИ ШПБПДЛ"):
        mrh.build_pidstavy_extra_grounds_by_person(rows, ("ПІДСТАВИ ШПБПДЛ", "ПІДСТАВИ БПШПДЛ"))


def test_build_pidstavy_extra_grounds_by_person_allows_only_one_alternate_column_present():
    rows = [{"ПІБ": "Перший Перший", "ПІДСТАВИ БПШПДЛ": "Висновок Б"}]

    assert mrh.build_pidstavy_extra_grounds_by_person(rows, ("ПІДСТАВИ ШПБПДЛ", "ПІДСТАВИ БПШПДЛ")) == {
        "ПЕРШИЙ ПЕРШИЙ": ["Висновок Б"],
    }


def test_build_pidstavy_extra_grounds_by_person_defaults_to_no_columns():
    """column_names=() за замовчуванням - категорія без власної column_basis_from_tracking
    (None) не читає жодної колонки з файлу (рішення користувача: жодної спільної
    "загальної" колонки для категорій без власної більше немає)."""
    rows = [{"ПІБ": "Перший Перший", "ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ": "Текст", "ПІДСТАВИ": "Текст"}]

    assert mrh.build_pidstavy_extra_grounds_by_person(rows) == {}


# -------------------------
# _row_szch_whole_month - порожні/NaN комірки (гілки перевірки)
# -------------------------
def test_row_szch_whole_month_false_for_none_and_nan_cells():
    d1, d2 = datetime(2026, 9, 1), datetime(2026, 9, 2)

    assert mrh._row_szch_whole_month({d1: "СЗЧ", d2: None}, [d1, d2]) is False
    assert mrh._row_szch_whole_month({d1: "СЗЧ", d2: np.nan}, [d1, d2]) is False


# -------------------------
# column_basis_from_tracking приймає БУДЬ-ЯКУ форму запису (utils/column_names.py)
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
    assert column_names_of(value) == ("ПІДСТАВИ 10К",)


def test_column_names_of_treats_a_non_iterable_value_as_one_name():
    assert column_names_of(10) == ("10",)


def test_column_names_of_joins_several_names_and_drops_duplicates_keeping_order():
    assert column_names_of(("ПІДСТАВИ ШПБП", "підстави  шпбп", "ПІДСТАВИ БПШП")) == ("ПІДСТАВИ ШПБП", "ПІДСТАВИ БПШП")


def test_column_names_of_treats_empty_values_as_no_names():
    assert column_names_of(None) == ()
    assert column_names_of("") == ()
    assert column_names_of(("", "   ", None)) == ()


# -------------------------
# basis_source_for_point - читає "column_basis_from_tracking" напряму з categories
# -------------------------
def test_basis_source_for_point_routes_by_category_field():
    """Пункт класифікується за полями категорій САМОГО categories: required_basis=True
    - "basis" (власна колонка з виключенням), required_basis=False - "own" (власна без
    виключення), жодної "column_basis_from_tracking" - "general"."""
    categories = {
        "10_МЕД": {
            "general": [],
            "МЕДИЧНІ": {"required_basis": True, "column_basis_from_tracking": "ПІДСТАВИ МЕДИЧНІ НА 10К"},
        },
        10: {
            "general": [],
            "ЖИТТЄДІЯЛЬНІСТЬ": {"required_basis": False, "column_basis_from_tracking": ("ПІДСТАВИ 10К",)},
        },
        30: {"general": []},
    }

    assert mrh.basis_source_for_point("10_МЕД", categories) == ("basis", ("ПІДСТАВИ МЕДИЧНІ НА 10К",))
    assert mrh.basis_source_for_point(10, categories) == ("own", ("ПІДСТАВИ 10К",))
    assert mrh.basis_source_for_point(30, categories) == ("general", ())


def test_basis_source_for_point_ignores_categories_without_the_field():
    """Категорія БЕЗ "column_basis_from_tracking" (чи None) не заважає ІНШІЙ названій
    категорії того самого пункту визначити "basis"/"own" - лише та, що сама нічого не
    задала, просто пропускається."""
    categories = {
        "10_МЕД": {
            "general": [],
            "МЕДИЧНІ": {"required_basis": True, "column_basis_from_tracking": "ПІДСТАВИ МЕДИЧНІ НА 10К"},
            "БЕЗ_ВЛАСНОЇ": {"required_basis": True, "column_basis_from_tracking": None},
        },
    }

    assert mrh.basis_source_for_point("10_МЕД", categories) == ("basis", ("ПІДСТАВИ МЕДИЧНІ НА 10К",))


def test_basis_source_for_point_category_with_explicit_none_inherits_point_level_value():
    """Питання користувача: якщо в пункту є власне "column_basis_from_tracking", а
    категорія явно поставила СВОЄ None (а не просто не дописала ключ) - має
    спрацювати так само, як і без ключа взагалі - успадкувати значення пункту,
    а НЕ "ніякої колонки" (точний приклад користувача: "100_СПЕЦКОНТИНГЕНТ" з
    власним "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ" на рівні пункту, підстатус - None)."""
    categories = {
        "100_СПЕЦКОНТИНГЕНТ": {
            "general": [],
            "column_basis_from_tracking": ("ПІДСТАВИ СПЕЦКОНТИНГЕНТУ"),
            "полон": {"required_basis": True, "column_basis_from_tracking": None},
        },
    }

    assert mrh.basis_source_for_point("100_СПЕЦКОНТИНГЕНТ", categories) == ("basis", ("ПІДСТАВИ СПЕЦКОНТИНГЕНТУ",))


def test_basis_source_for_point_category_with_real_value_overrides_point_level_value():
    """Дзеркально до попереднього - категорія з НЕПОРОЖНІМ власним значенням
    перевизначає значення пункту лише для себе (напр. якщо пізніше знадобиться
    ОКРЕМА колонка лише для одного підстатусу з кількох, що ділять спільну)."""
    categories = {
        "NOT_PAID": {
            "column_basis_from_tracking": "ПІДСТАВИ ДЛЯ НЕ ВИПЛАТИ",
            "СЗЧ": {"required_basis": True, "column_basis_from_tracking": None},
            "Арешт": {"required_basis": True, "column_basis_from_tracking": "ПІДСТАВИ ЗА АРЕШТ"},
        },
    }

    with pytest.raises(ValueError, match="NOT_PAID"):
        mrh.basis_source_for_point("NOT_PAID", categories)


def test_basis_source_for_point_rejects_categories_disagreeing_on_column_name():
    """Дві названі категорії ОДНОГО пункту задають РІЗНІ column_basis_from_tracking -
    незрозуміло, яку з колонок брати для пункту в цілому, тож помилка, а не мовчазний
    вибір однієї з них."""
    categories = {
        "10_МЕД": {
            "general": [],
            "А": {"required_basis": True, "column_basis_from_tracking": "ПІДСТАВИ А"},
            "Б": {"required_basis": True, "column_basis_from_tracking": "ПІДСТАВИ Б"},
        },
    }

    with pytest.raises(ValueError, match="10_МЕД"):
        mrh.basis_source_for_point("10_МЕД", categories)


def test_basis_source_for_point_rejects_categories_disagreeing_on_required_basis():
    categories = {
        "10_МЕД": {
            "general": [],
            "А": {"required_basis": True, "column_basis_from_tracking": "ПІДСТАВИ МЕДИЧНІ НА 10К"},
            "Б": {"required_basis": False, "column_basis_from_tracking": "ПІДСТАВИ МЕДИЧНІ НА 10К"},
        },
    }

    with pytest.raises(ValueError, match="10_МЕД"):
        mrh.basis_source_for_point("10_МЕД", categories)


def test_basis_source_for_point_rejects_column_basis_from_tracking_without_any_column_name():
    """Порожнє значення - помилка з назвою пункту, а не мовчазне виключення всіх людей з пункту."""
    categories = {"10_МЕД": {"general": [], "А": {"required_basis": True, "column_basis_from_tracking": ""}}}

    with pytest.raises(ValueError, match="10_МЕД"):
        mrh.basis_source_for_point("10_МЕД", categories)


def test_basis_source_for_point_reads_live_constants_when_categories_not_given():
    assert mrh.basis_source_for_point("100_БПШП") == mrh.basis_source_for_point("100_БПШП", constants.MONEY_REPORT_CATEGORIES)
    assert mrh.basis_source_for_point("100_БПШП")[0] == "basis"


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


def test_build_categories_reads_basis_from_a_bare_string_column_basis_from_tracking(capsys):
    """Категорія, чию "column_basis_from_tracking" записано РЯДКОМ (без коми), бере
    підставу саме з неї; людина без неї виключається з пункту з попередженням - так
    само, як для будь-якого пункту з required_basis=True. categories передається явно
    (без підміни модуля чи реального constants.py)."""
    import generators.generate_report_for_get_money as report_module

    d1 = datetime(2026, 9, 1)
    with_basis = _row("Перший Перший Перший", "Стрілець", "Підрозділ 1", {d1: "БПШП"})
    with_basis["ПІДСТАВИ ТЕСТ"] = "Висновок ВЛК №9"
    without_basis = _row("Другий Другий Другий", "Стрілець", "Підрозділ 1", {d1: "БПШП"})

    test_categories = {
        "100_БПШП": {
            "general": [],
            "БПШП": {
                "grounds": [],
                "use_brs": False,
                "use_brs_from_selected_folder": False,
                "exclude_general": [],
                "default": True,
                "generate_br": False,
                "include_to_report": True,
                "required_basis": True,
                "column_basis_from_tracking": ("ПІДСТАВИ ТЕСТ"),
            },
        },
    }

    categories = report_module._build_categories(
        [with_basis, without_basis], [d1], {}, categories=test_categories,
    )

    section = dict(categories).get("100_БПШП", [])
    assert [r["ПІБ"] for r in section] == ["Перший Перший Перший"]
    assert "Висновок ВЛК №9" in section[0]["ПІДСТАВА"]
    assert "Немає підстави для 100_БПШП: Другий Другий Другий" in capsys.readouterr().out
