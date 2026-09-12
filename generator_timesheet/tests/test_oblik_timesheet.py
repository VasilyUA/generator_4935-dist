from datetime import date, datetime

import pytest
from openpyxl import Workbook, load_workbook

import content.oblik_timesheet as ot
from constants import (
    DEFAULT_STATUS,
    FREEFORM_STATUS_PATTERNS,
    STATUS_COLORS,
    STATUS_CONFLICT_OVERRIDES,
    STATUS_HEADING_RULES,
)


def _write_workbook(path, header, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Табель"
    for col_idx, value in enumerate(header, start=1):
        ws.cell(row=1, column=col_idx, value=value)
    for row_idx, row_values in enumerate(rows, start=2):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    wb.save(str(path))
    return str(path)


_HEADER = [
    "ПОСАДА", "ЗВАННЯ", "ПІБ",
    datetime(2026, 7, 31), datetime(2026, 8, 1), datetime(2026, 8, 2), datetime(2026, 8, 3),
    "ПІДСТАВИ",
]


@pytest.fixture
def base_workbook(tmp_path):
    return _write_workbook(tmp_path / "ОБЛІК.xlsx", _HEADER, [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", None, None, None, ""],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "ВП", None, None, None, "якась підстава"],
    ])


def test_timesheet_detects_label_and_date_columns(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")

    assert sheet.label_columns["ПІБ"] == 3
    assert sheet.label_columns["ПОСАДА"] == 1
    assert [d for _, d in sheet.date_columns] == [date(2026, 7, 31), date(2026, 8, 1), date(2026, 8, 2), date(2026, 8, 3)]


def test_timesheet_requires_pib_column(tmp_path):
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", ["ПОСАДА", datetime(2026, 8, 1)], [["Посада1", None]])
    with pytest.raises(ValueError, match="ПІБ"):
        ot.Timesheet(path, "Табель")


def test_timesheet_requires_at_least_one_date_column(tmp_path):
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", ["ПОСАДА", "ПІБ"], [["Посада1", "ПЕРШИЙ Перший Перший"]])
    with pytest.raises(ValueError, match="дати"):
        ot.Timesheet(path, "Табель")


def test_normalize_name_returns_empty_string_for_non_string_values():
    assert ot.normalize_name(None) == ""
    assert ot.normalize_name(12345) == ""


def test_apply_events_returns_early_without_any_report_dates(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [{"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВП", "effective_date": date(2026, 8, 1)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value is None  # нічого не заповнено - немає жодного рапорту


def test_person_rows_normalizes_pib(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    rows = sheet.person_rows()

    assert rows[ot.normalize_name("Перший Перший Перший")] == rows["ПЕРШИЙ ПЕРШИЙ ПЕРШИЙ"] == 2
    assert rows["ДРУГИЙ ДРУГИЙ ДРУГИЙ"] == 3


def test_apply_events_carries_status_forward_and_applies_change(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [{"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВП", "effective_date": date(2026, 8, 2)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1), date(2026, 8, 2)])

    assert unresolved == []
    row = 2  # ПЕРШИЙ
    assert sheet.ws.cell(row=row, column=5).value == "РВЗ"  # 01.08 - без подій, продовжує з базового 31.07
    assert sheet.ws.cell(row=row, column=6).value == "ВП"   # 02.08 - подія (останній рапорт)
    assert sheet.ws.cell(row=row, column=7).value is None  # 03.08 - після останнього рапорту, не займається
    # ПІДСТАВИ (позаду останньої колонки-дати) не займається взагалі - лишається тим, чим було для ДРУГОГО.
    assert sheet.ws.cell(row=3, column=8).value == "якась підстава"


def test_apply_events_adopts_a_manually_entered_status_within_the_report_range_as_a_new_anchor(base_workbook):
    """Реальний випадок: 31.07 - базовий "РВЗ",
    змінюється за рапортами як завжди, АЛЕ на 02.08 (дата В МЕЖАХ діапазону
    наявних рапортів - 01.08-03.08, а НЕ після нього) уже вручну вписано
    "БЗ", і жоден рапорт НЕ подає події САМЕ на цю дату - за прямою вказівкою
    користувача, цей вручну вписаний статус ВСЕ ОДНО стає новим "якорем"
    (як і подія з рапорту), а НЕ мовчки перезаписується carried-forward
    статусом ("РВЗ") - і ПРОДОВЖУЄТЬСЯ (forward-fill) далі (03.08), доки не
    трапиться подія з рапорту чи ще один вручну вписаний статус."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=6, value="БЗ")  # ПЕРШИЙ, 02.08 - вручну вписано, В МЕЖАХ діапазону рапортів
    events = []

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1), date(2026, 8, 2), date(2026, 8, 3)])

    assert unresolved == []
    row = 2  # ПЕРШИЙ
    assert sheet.ws.cell(row=row, column=5).value == "РВЗ"  # 01.08 - продовжує базовий (подій немає)
    assert sheet.ws.cell(row=row, column=6).value == "БЗ"   # 02.08 - вручну вписаний якір (подій немає, НЕ перезаписано)
    assert sheet.ws.cell(row=row, column=7).value == "БЗ"   # 03.08 - продовжено від якоря (подій немає)


def test_apply_events_report_event_still_overrides_a_manually_entered_status(base_workbook):
    """На відміну від тесту вище - якщо на ту САМУ дату є подія З РАПОРТУ,
    вона Й ДАЛІ перемагає (події завжди мають пріоритет, вручну вписаний
    статус - лише запасний "якір", коли події немає САМЕ на цю дату)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=6, value="БЗ")  # ПЕРШИЙ, 02.08 - вручну вписано
    events = [{"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВП", "effective_date": date(2026, 8, 2)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1), date(2026, 8, 2)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=6).value == "ВП"  # подія перемагає вручну вписаний статус


def test_apply_events_continues_a_manually_entered_status_after_the_last_report(base_workbook):
    """Реальний випадок: 31.07 - базовий "РВЗ",
    змінюється за рапортами як завжди, АЛЕ поза діапазоном наявних рапортів
    (тут - лише 01.08) хтось УЖЕ вручну вписав статус ("БЗ") у ПІЗНІШУ
    колонку (03.08, а не одразу наступну 02.08) - за прямою вказівкою
    користувача, цей статус ПРОДОВЖУЄТЬСЯ (forward-fill) у ВСІ НАСТУПНІ
    порожні колонки ПІСЛЯ нього (тут таких немає - 03.08 останній), а
    колонки МІЖ останнім рапортом і цим записом (02.08 - порожня, БЕЗ
    підтвердженого "якоря") лишаються порожніми, як і раніше."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=7, value="БЗ")  # ПЕРШИЙ, 03.08 - вручну вписано наперед
    events = []

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert unresolved == []
    row = 2  # ПЕРШИЙ
    assert sheet.ws.cell(row=row, column=5).value == "РВЗ"  # 01.08 - продовжує базовий (останній рапорт)
    assert sheet.ws.cell(row=row, column=6).value is None   # 02.08 - поза рапортами, БЕЗ якоря - порожня
    assert sheet.ws.cell(row=row, column=7).value == "БЗ"   # 03.08 - вручну вписаний якір, лишається як є


def test_apply_events_continues_a_manually_entered_status_into_later_blank_columns(base_workbook):
    """Той самий принцип, що й тест вище, але з колонкою ПІСЛЯ якоря, яку є
    куди продовжувати - "БЗ", вписаний у 02.08 (одразу після останнього
    рапорту - 01.08), ПРОДОВЖУЄТЬСЯ (forward-fill) у 03.08 (порожню
    колонку), а не лишається ізольованим лише в одній клітинці."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=6, value="БЗ")  # ПЕРШИЙ, 02.08 - вручну вписано
    events = []

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert unresolved == []
    row = 2  # ПЕРШИЙ
    assert sheet.ws.cell(row=row, column=5).value == "РВЗ"  # 01.08 - продовжує базовий (останній рапорт)
    assert sheet.ws.cell(row=row, column=6).value == "БЗ"   # 02.08 - вручну вписаний якір
    assert sheet.ws.cell(row=row, column=7).value == "БЗ"   # 03.08 - продовжено від якоря


def test_apply_events_colors_changed_cell_according_to_status_colors(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [{"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ШП", "effective_date": date(2026, 8, 1)}]

    ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    cell = sheet.ws.cell(row=2, column=5)
    assert cell.value == "ШП"
    assert cell.fill.fgColor.rgb == STATUS_COLORS["ШП"]


def test_apply_events_colors_carried_forward_cell_too_not_only_changed_ones(base_workbook):
    """Колір застосовується для КОЖНОЇ заповненої клітинки, а не лише для тих,
    де БУЛА подія - клітина, що просто "продовжує" попередній статус (немає
    події на цю дату), теж має відповідний колір, інакше сітка виглядала б
    непослідовно (частина клітинок кольорова, частина - ні, хоча статус той
    самий текст)."""
    sheet = ot.Timesheet(base_workbook, "Табель")

    ot.apply_events(sheet, events=[], report_dates=[date(2026, 8, 1)])

    cell = sheet.ws.cell(row=2, column=5)
    assert cell.value == "РВЗ"
    assert cell.fill.fgColor.rgb == STATUS_COLORS[DEFAULT_STATUS]


def test_apply_events_colors_baseline_column_without_changing_its_value(base_workbook):
    """Базова (найперша) колонка-дата - ЄДИНА, значення якої apply_events
    НІКОЛИ не змінює (див. докстрінг apply_events), але користувач очікує, що
    її заливка все одно відповідатиме кольоровій гамі, як і решта рядка -
    інакше саме перша колонка виглядала б "забутою" серед пофарбованих
    сусідів."""
    sheet = ot.Timesheet(base_workbook, "Табель")

    ot.apply_events(sheet, events=[], report_dates=[date(2026, 8, 1)])

    baseline_cell = sheet.ws.cell(row=2, column=4)  # 31.07 - "РВЗ" у base_workbook
    assert baseline_cell.value == "РВЗ"  # значення НЕ змінилось
    assert baseline_cell.fill.fgColor.rgb == STATUS_COLORS[DEFAULT_STATUS]

    baseline_cell_vp = sheet.ws.cell(row=3, column=4)  # 31.07 - "ВП" у base_workbook (ДРУГИЙ)
    assert baseline_cell_vp.value == "ВП"
    assert baseline_cell_vp.fill.fgColor.rgb == STATUS_COLORS["ВП"]


def test_apply_events_leaves_unknown_baseline_status_uncolored(base_workbook):
    """Базова колонка з текстом, якого немає в STATUS_COLORS (напр. вручну
    введений статус, невідомий цій програмі) - лишається БЕЗ заливки, а не
    падає з помилкою й не отримує якийсь довільний колір за замовчуванням."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="НЕВІДОМИЙ СТАТУС")

    ot.apply_events(sheet, events=[], report_dates=[date(2026, 8, 1)])

    baseline_cell = sheet.ws.cell(row=2, column=4)
    assert baseline_cell.value == "НЕВІДОМИЙ СТАТУС"
    assert baseline_cell.fill.fill_type is None


def test_apply_payment_values_colors_substituted_cell_with_the_categorys_own_color(base_workbook):
    """За прямою вказівкою користувача - КОЖНА категорія виплати (число) має
    ВЛАСНИЙ, ВІДМІННИЙ від інших колір (STATUS_COLORS[100], НЕ зелений
    DEFAULT_STATUS для всіх чисел підряд) - щоб категорії було видно на
    очах, не лише читаючи текст."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")
    sheet.ws.cell(row=2, column=5, value="РВЗ")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [100]}

    ot.apply_payment_values(sheet, payment_values)

    cell = sheet.ws.cell(row=2, column=5)
    assert cell.value == 100
    assert cell.fill.fgColor.rgb == STATUS_COLORS[100]


def test_apply_payment_values_colors_baseline_column_too(base_workbook):
    """apply_payment_values (на відміну від apply_events) підставляє
    категорію виплати і в БАЗОВУ (найпершу) колонку-дату - вона повинна
    отримати колір так само, як і решта, хоча apply_events її взагалі не
    торкається й ніколи не фарбує."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): [100]}

    ot.apply_payment_values(sheet, payment_values)

    cell = sheet.ws.cell(row=2, column=4)
    assert cell.value == 100
    assert cell.fill.fgColor.rgb == STATUS_COLORS[100]


def test_status_colors_cover_every_status_the_app_can_actually_produce():
    """Регресійний тест: якщо хтось додасть новий статус (новий рядок у
    STATUS_HEADING_RULES/FREEFORM_STATUS_PATTERNS/STATUS_CONFLICT_OVERRIDES),
    але забуде додати для нього колір - ОБЛІК.xlsx мовчки лишить цю клітинку
    без заливки замість помітної кольорової гами. Не помилка виконання (див.
    _fill_for_status - None просто не фарбує), але саме цей тест мав би про
    це попередити."""
    all_statuses = (
        {DEFAULT_STATUS}
        | set(STATUS_HEADING_RULES.values())
        | {status for _, status in FREEFORM_STATUS_PATTERNS}
        | set(STATUS_CONFLICT_OVERRIDES.values())
    )

    assert all_statuses <= set(STATUS_COLORS)


@pytest.mark.parametrize("category", [10, 30, 70, 100, 170])
def test_status_colors_covers_every_payment_category(category):
    """Регресійний тест, дзеркальний до test_status_colors_cover_every_
    status_the_app_can_actually_produce - але для КАТЕГОРІЙ ВИПЛАТИ (числа,
    не текстові статуси): якщо хтось прибере колір для однієї з них,
    apply_payment_values мовчки лишить клітинку без заливки."""
    assert category in STATUS_COLORS


def test_status_colors_gives_every_payment_category_a_distinct_color():
    """За прямою вказівкою користувача - КОЖНА категорія виплати має
    ВІДМІННИЙ від інших колір (не лише "якийсь" колір - вони мають РІЗНИТИСЬ
    між собою, інакше 30 і 100, напр., виглядали б однаково на очах)."""
    colors = [STATUS_COLORS[category] for category in (10, 30, 70, 100, 170)]

    assert len(set(colors)) == len(colors)


def test_apply_events_does_not_touch_dates_after_last_report(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    ot.apply_events(sheet, events=[], report_dates=[date(2026, 8, 1)])

    row = 2
    assert sheet.ws.cell(row=row, column=5).value == "РВЗ"  # 01.08 - в межах звіту, продовжено
    assert sheet.ws.cell(row=row, column=6).value is None  # 02.08 - після останнього рапорту, не займається
    assert sheet.ws.cell(row=row, column=7).value is None  # 03.08


def test_apply_events_reports_unknown_pib(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [{"pib_raw": "НЕВІДОМИЙ Хтось Хтозна", "status": "ВП", "effective_date": date(2026, 8, 1)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert len(unresolved) == 1
    assert "не знайдено" in unresolved[0]["reason"]
    assert "друкарська помилка" not in unresolved[0]["reason"]  # немає схожого написання в роcтері - без підказки


def test_apply_events_auto_applies_status_when_a_unique_similar_pib_exists(tmp_path):
    """Реальний випадок: рапорт написав ПІБ з переставленими літерами в
    прізвищі порівняно з роcтеровим написанням - за прямою вказівкою
    користувача, подія ВСЕ ОДНО ЗАСТОСОВУЄТЬСЯ до цього єдиного схожого
    кандидата (а не лишається лише підказкою), і ОДНОЧАСНО потрапляє в
    unresolved як ІНФОРМАЦІЙНИЙ запис (не потребує дії, лише для прозорості
    - щоб було видно, що й кого торкнулась автоматична корекція)."""
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", _HEADER, [
        ["Посада1", "сержант", "ЧОТИРНАДЦЯТИЙ Чотирнадцятий Чотирнадцятий", "РВЗ", None, None, None, ""],
    ])
    sheet = ot.Timesheet(path, "Табель")
    events = [{"pib_raw": "ЧОТИРНАЦДЯТИЙ Чотирнадцятий Чотирнадцятий", "status": "ВП", "effective_date": date(2026, 8, 1)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert len(unresolved) == 1
    assert "не збігається буквально" in unresolved[0]["reason"]
    assert "друкарську помилку" in unresolved[0]["reason"]
    assert "ЧОТИРНАДЦЯТИЙ Чотирнадцятий Чотирнадцятий" in unresolved[0]["reason"]
    # Подія ЗАСТОСОВАНА до єдиного схожого роcтерового кандидата - "ВП", а
    # не просто перенесений базовий статус "РВЗ".
    assert sheet.ws.cell(row=2, column=5).value == "ВП"


def test_apply_events_does_not_auto_apply_when_multiple_similar_pib_exist(tmp_path):
    """Два роcтерові написання ОБИДВА достатньо схожі на ПІБ з рапорту -
    неоднозначно, яке з них мається на увазі, тож подія НЕ застосовується
    автоматично (краще нічого не застосувати, ніж застосувати до НЕ ТІЄЇ
    людини) - обидва роcтерові рядки лишаються на своєму базовому статусі."""
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", _HEADER, [
        ["Посада1", "сержант", "ЧОТИРНАДЦЯТИЙ Чотирнадцятий Чотирнадцятий", "РВЗ", None, None, None, ""],
        ["Посада2", "матрос", "ЧОТИРНАЦЯДТИЙ Чотирнадцятий Чотирнадцятий", "РВЗ", None, None, None, ""],
    ])
    sheet = ot.Timesheet(path, "Табель")
    events = [{"pib_raw": "ЧОТИРНАЦДЯТИЙ Чотирнадцятий Чотирнадцятий", "status": "ВП", "effective_date": date(2026, 8, 1)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert len(unresolved) == 1
    assert "не знайдено" in unresolved[0]["reason"]
    assert "друкарська помилка" not in unresolved[0]["reason"]
    assert sheet.ws.cell(row=2, column=5).value == "РВЗ"
    assert sheet.ws.cell(row=3, column=5).value == "РВЗ"


def test_apply_events_does_not_auto_apply_when_similar_pib_has_a_different_surname(tmp_path):
    """Реальний випадок (баг): ДВІ РІЗНІ людини з ТИМ САМИМ ім'ям+по-батькові
    (поширене поєднання), але РІЗНИМИ прізвищами - ratio ЦІЛОГО ПІБ хибно
    перевищував поріг (спільна велика частина імені+по-батькові "переважує"
    різницю в прізвищі), хоча самі ПРІЗВИЩА - геть різні слова. Роcтер має
    лише "ДРУГИЙ" - рапорт помилково (чи через плутанину) згадує "ПЕРШИЙ" з
    тим самим ім'ям і по-батькові. Прізвища НЕ мають пройти поріг схожості
    окремо, тож подія НЕ застосовується автоматично - роcтерова людина
    лишається на своєму базовому статусі, а unresolved НЕ згадує "друкарську
    помилку" (як для справжньої одруку - див.
    test_apply_events_auto_applies_status_when_a_unique_similar_pib_exists)."""
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", _HEADER, [
        ["Посада1", "сержант", "ДРУГИЙ Чотирнадцятий Чотирнадцятий", "РВЗ", None, None, None, ""],
    ])
    sheet = ot.Timesheet(path, "Табель")
    events = [{"pib_raw": "ПЕРШИЙ Чотирнадцятий Чотирнадцятий", "status": "ВП", "effective_date": date(2026, 8, 1)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert len(unresolved) == 1
    assert "не знайдено" in unresolved[0]["reason"]
    assert "друкарську помилку" not in unresolved[0]["reason"]
    assert sheet.ws.cell(row=2, column=5).value == "РВЗ"


def test_apply_events_silently_skips_date_before_the_earliest_column_when_it_confirms_the_baseline(base_workbook):
    """Реальний випадок: подія зі щоденного рапорту вказує дату, що передує
    ВЖЕ наявному в ОБЛІК.xlsx базовому стану (напр. "з 21.07 у відпустці", а
    базова колонка табеля починається щойно з 31.07) - НЕ друкарська помилка
    (дата цілком правильна, просто раніша за відстежуваний діапазон), і сам
    СТАТУС події ЗБІГАЄТЬСЯ з тим, що ВЖЕ стоїть у базовій колонці ("РВЗ" у
    base_workbook для ПЕРШОГО) - подія лише ПІДТВЕРДЖУЄ вже наявний запис,
    базова колонка НІКОЛИ не змінюється, тож подія просто НЕМА КУДИ
    застосовувати - мовчки пропускається, БЕЗ запису на ручну перевірку."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [{"pib_raw": "ПЕРШИЙ Перший Перший", "status": "РВЗ", "effective_date": date(2026, 7, 2)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert unresolved == []


def test_apply_events_reports_date_before_the_earliest_column_when_it_contradicts_the_baseline(base_workbook):
    """На відміну від тесту вище - дата РАНІШЕ найпершої колонки САМА ПО СОБІ
    ще НЕ доказ, що це не друкарська помилка: якщо статус події НЕ збігається
    з тим, що вже стоїть у базовій колонці ("РВЗ" у base_workbook, подія тут
    - "ВП"), сумнів не знято - лишається на ручну перевірку, як і раніше."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [{"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВП", "effective_date": date(2026, 7, 2)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert len(unresolved) == 1
    assert "відсутня серед колонок" in unresolved[0]["reason"]


def test_apply_events_reports_out_of_range_date_within_the_tracked_period(base_workbook):
    """На відміну від дати РАНІШЕ найпершої колонки (мовчки пропускається,
    тест вище) - дата, якої немає серед колонок з ІНШОЇ причини (тут - вона
    ПІЗНІША за останню колонку табеля, типова причина - друкарська помилка
    в самому рапорті) - і далі лишається на ручну перевірку, як раніше."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [{"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВП", "effective_date": date(2026, 8, 10)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert len(unresolved) == 1
    assert "відсутня серед колонок" in unresolved[0]["reason"]


def test_apply_events_silently_skips_date_exactly_one_day_after_the_latest_column(base_workbook):
    """Реальний випадок, підтверджений користувачем: рапорт ОСТАННЬОГО дня
    місяця описує складену подію ("виписаний ... та проходить ВЛК ...,
    потребує відпустки ..." - content/daily_report_reader.
    _discharge_vlk_leave_events), де відпустка (друга подія) обчислюється
    як "наступний день" - РІВНО день ПІСЛЯ останньої наявної колонки-дати
    (тут - 03.08.2026, тож "наступний день" - 04.08.2026, якого табель ще
    не охоплює). НЕ друкарська помилка (сам НАСТУПНИЙ місяць просто ще не
    існує як колонки цього табеля) - мовчки пропускається, БЕЗ запису на
    ручну перевірку, на відміну від дати, що ПРОСТО пізніша за останню
    колонку з ІНШОЇ причини (тест вище, "2026-08-10" - на 7 днів пізніше,
    а не рівно на 1)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [{"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВПСЗ", "effective_date": date(2026, 8, 4)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 3)])

    assert unresolved == []


def test_apply_events_warns_on_conflicting_same_day_statuses(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВП", "effective_date": date(2026, 8, 1)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ШП", "effective_date": date(2026, 8, 1)},
    ]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert len(unresolved) == 1
    assert "суперечливі статуси" in unresolved[0]["reason"]
    assert sheet.ws.cell(row=2, column=5).value == "ШП"  # немає override для цієї пари - останній перемагає


def test_apply_events_resolves_vp_vpsz_conflict_without_manual_review(base_workbook):
    """Реальний випадок: "ВП" і "ВПСЗ" для однієї людини/дати не є справжньою
    суперечністю (ВПСЗ - той самий факт відпустки, лише уточнена причина) -
    ВПСЗ має застосуватись напряму, без запису на ручну перевірку, незалежно
    від порядку появи в рапорті."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВП", "effective_date": date(2026, 8, 1)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВПСЗ", "effective_date": date(2026, 8, 1)},
    ]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == "ВПСЗ"


def test_apply_events_resolves_vpsz_vp_conflict_regardless_of_order(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВПСЗ", "effective_date": date(2026, 8, 1)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВП", "effective_date": date(2026, 8, 1)},
    ]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == "ВПСЗ"


def test_apply_events_resolves_vpbp_vpsz_conflict_without_manual_review(base_workbook):
    """Реальний випадок, підтверджений користувачем: один рапорт дає "ВПБП"
    (сам факт "потребує відпустки для лікування після поранення"), а
    ПІЗНІШИЙ рапорт ОКРЕМИМ реченням підтверджує ТОЙ САМИЙ день як "вибув у
    відпустку ЗА СТАНОМ ЗДОРОВ'Я" (ВПСЗ) - той самий факт відпустки, лише
    підтверджений іншим формулюванням причини - НЕ справжня суперечність,
    "ВПСЗ" застосовується напряму, без запису на ручну перевірку."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВПБП", "effective_date": date(2026, 8, 1)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВПСЗ", "effective_date": date(2026, 8, 1)},
    ]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == "ВПСЗ"


def test_apply_payment_values_replaces_default_status_with_the_numeric_value(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value="РВЗ")  # ПЕРШИЙ, 01.08
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [100]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == 100


def test_apply_payment_values_leaves_non_default_status_untouched(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано
    # ДРУГИЙ 31.07 = "ВП" (з base_workbook), не "РВЗ" - саме це тут і перевіряється.
    payment_values = {(ot.normalize_name("ДРУГИЙ Другий Другий"), date(2026, 7, 31)): [100]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=3, column=4).value == "ВП"


def test_apply_payment_values_overrides_vp_with_vpbp_to_the_literal_compound_text(base_workbook):
    """PAYMENT_STATUS_OVERRIDES["ВП"] (constants.py) - той самий принцип, що
    й "ШП" нижче: файл information_unit ОДНОЗНАЧНО подає "ВПБП" для тієї
    самої дати - підрозділ підтверджує категорію "100_ВПБП"
    (PAYMENT_STATUS_CONVERSIONS["ВПБП"]), тож "ВП" ЗАМІНЮЄТЬСЯ на буквальний
    текст "100_ВПБП"."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ, 31.07
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): ["ВПБП"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "100_ВПБП"
    assert sheet.ws.cell(row=2, column=4).fill.fgColor.rgb == STATUS_COLORS["100_ВПБП"]


def test_apply_payment_values_overrides_vp_with_vps_to_itself(base_workbook):
    """PAYMENT_STATUS_OVERRIDES["ВП"]["ВПС"] - ідентичність (та сама логіка,
    що й "ПРВД"/"Перевд." вище): файл information_unit подає "ВПС"
    (відпустка за сімейними обставинами) для тієї самої дати - "ВП"
    замінюється на "ВПС" (без числової категорії - лише підтверджує
    конкретнішу причину)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ, 31.07
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): ["ВПС"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "ВПС"
    assert sheet.ws.cell(row=2, column=4).fill.fgColor.rgb == STATUS_COLORS["ВПС"]


def test_apply_payment_values_overrides_shp_with_bpshp_to_the_literal_compound_text(base_workbook):
    """Реальний випадок: "ШП" - день виписки з лікарні, конкретна датована
    подія щоденного рапорту, а файл information_unit ОДНОЗНАЧНО подає "БПШП"
    для тієї самої дати - підрозділ підтверджує, що ця дата вже враховується
    в категорії "100_БПШП" (PAYMENT_STATUS_CONVERSIONS["БПШП"]), тож "ШП"
    ЗАМІНЮЄТЬСЯ на буквальний текст "100_БПШП"."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ШП")  # ПЕРШИЙ, 31.07
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): ["БПШП"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "100_БПШП"
    assert sheet.ws.cell(row=2, column=4).fill.fgColor.rgb == STATUS_COLORS["100_БПШП"]


def test_apply_payment_values_overrides_vpsz_with_vpbp_to_the_literal_compound_text(base_workbook):
    """Той самий принцип, що й "ШП" -> "100_БПШП" вище, але для "ВПСЗ"
    (реальний випадок: вибуття у відпустку за станом здоров'я, конкретна
    датована подія щоденного рапорту) і "ВПБП" - файл information_unit
    ОДНОЗНАЧНО подає "ВПБП" для тієї самої дати - підрозділ підтверджує, що
    ця дата вже враховується в категорії "100_ВПБП"
    (PAYMENT_STATUS_CONVERSIONS["ВПБП"]), тож "ВПСЗ" ЗАМІНЮЄТЬСЯ на
    буквальний текст "100_ВПБП"."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВПСЗ")  # ПЕРШИЙ, 31.07
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): ["ВПБП"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "100_ВПБП"
    assert sheet.ws.cell(row=2, column=4).fill.fgColor.rgb == STATUS_COLORS["100_ВПБП"]


def test_apply_payment_values_overrides_prvd_with_perevd_to_itself(base_workbook):
    """Реальний випадок: "ПРВД" (переведений в інший підрозділ) в ОБЛІК, а
    файл information_unit ОДНОЗНАЧНО каже "Перевд." - ТОЙ САМИЙ факт, лише
    написання information_unit - тож "ПРВД" замінюється на "ПРВД" (без
    видимої зміни значення). ЦЕЙ запис-ідентичність існує заради
    _statuses_equivalent (content/payment_mismatch_checker.py) - щоб пара
    ("ПРВД","Перевд.") визнавалась ЕКВІВАЛЕНТНОЮ (не розбіжністю) у "повній
    історії" файлу з РЕАЛЬНОЮ розбіжністю на ІНШУ дату."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ПРВД")  # ПЕРШИЙ, 31.07
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): ["Перевд."]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "ПРВД"
    assert sheet.ws.cell(row=2, column=4).fill.fgColor.rgb == STATUS_COLORS["ПРВД"]


def test_apply_payment_values_colors_100_bpshp_the_same_regardless_of_path(base_workbook):
    """Реальний випадок: "100_БПШП" може з'явитись у клітинці ДВОМА РІЗНИМИ
    шляхами - напряму, коли базовий статус ВЖЕ "БПШП" (PAYMENT_STATUS_
    CONVERSIONS, фарбує ЗА СТАРИМ статусом "БПШП"), або через PAYMENT_
    STATUS_OVERRIDES, коли базовий статус "ШП", а information_unit
    підтверджує "БПШП" (фарбує ЗА НОВИМ значенням "100_БПШП"). Обидва шляхи
    МАЮТЬ давати ОДНАКОВИЙ колір для ОДНАКОВОГО тексту - інакше та сама
    клітинка "100_БПШП" виглядала б по-різному залежно від того, звідки вона
    взялась."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="БПШП")  # ПЕРШИЙ, 31.07 - шлях 1: PAYMENT_STATUS_CONVERSIONS
    sheet.ws.cell(row=2, column=5, value="ШП")  # ПЕРШИЙ, 01.08 - шлях 2: PAYMENT_STATUS_OVERRIDES
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): ["БПШП"]}

    ot.apply_payment_values(sheet, payment_values)

    cell_via_conversion = sheet.ws.cell(row=2, column=4)
    cell_via_override = sheet.ws.cell(row=2, column=5)
    assert cell_via_conversion.value == cell_via_override.value == "100_БПШП"
    assert cell_via_conversion.fill.fgColor.rgb == cell_via_override.fill.fgColor.rgb == STATUS_COLORS["100_БПШП"]


def test_apply_payment_values_converts_ppd_to_fixed_number_unconditionally(base_workbook):
    """PAYMENT_STATUS_CONVERSIONS (constants.py) - за прямою вказівкою
    користувача: "ППД" завжди замінюється на ФІКСОВАНЕ число 10, незалежно
    від payment_values (не з файлів information_unit - той самий статус
    завжди дає той самий запис)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ППД")

    unresolved = ot.apply_payment_values(sheet, payment_values={})

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == 10


def test_apply_payment_values_converts_bpshp_to_literal_compound_text(base_workbook):
    """"БПШП" - за прямою вказівкою користувача, замінюється на БУКВАЛЬНИЙ
    текст "100_БПШП" (не число окремо + причина десь-інде)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="БПШП")

    unresolved = ot.apply_payment_values(sheet, payment_values={})

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "100_БПШП"


@pytest.mark.parametrize("status, expected", [
    ("ППД", 10),
    ("БПШП", "100_БПШП"),
    ("ВПБП", "100_ВПБП"),
    ("БЗ", "100_БЗ"),
    ("полон", "100_БЗ"),
    ("Інт", "100_БЗ"),
])
def test_apply_payment_values_covers_every_fixed_conversion(base_workbook, status, expected):
    """Регресійний тест на КОЖЕН рядок PAYMENT_STATUS_CONVERSIONS - "БЗ"/
    "полон"/"Інт" усі троє дають ОДНАКОВЕ значення "100_БЗ" (підтверджено
    користувачем), а не кожен своє."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value=status)

    unresolved = ot.apply_payment_values(sheet, payment_values={})

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == expected


def test_apply_payment_values_colors_converted_cell_with_original_status_color(base_workbook):
    """Заливка конвертованої клітинки - за КОЛИШНІМ статусом ("ППД", світло-
    сірий), а не якимось новим кольором за замовчуванням - той самий підхід,
    що й для DEFAULT_STATUS ("РВЗ") вище (заливка лишається "зеленою" й
    ПІСЛЯ підстановки числа)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ППД")

    ot.apply_payment_values(sheet, payment_values={})

    assert sheet.ws.cell(row=2, column=4).fill.fgColor.rgb == STATUS_COLORS["ППД"]


@pytest.mark.parametrize("status", ["БЗ", "полон", "Інт"])
def test_apply_payment_values_colors_shared_conversion_target_consistently(base_workbook, status):
    """На відміну від "ППД" (ЄДИНЕ джерело для 10) - "БЗ"/"полон"/"Інт" усі
    троє ведуть до ОДНОГО й ТОГО САМОГО "100_БЗ" (PAYMENT_STATUS_CONVERSIONS)
    - заливка ТУТ фарбується ЗА НОВИМ значенням ("100_БЗ"), а НЕ за колишнім
    статусом, інакше та сама клітинка "100_БЗ" виглядала б ТРЬОМА РІЗНИМИ
    кольорами залежно від походження (реальний випадок, той самий клас
    багу, що й "100_БПШП" раніше)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value=status)

    ot.apply_payment_values(sheet, payment_values={})

    assert sheet.ws.cell(row=2, column=4).value == "100_БЗ"
    assert sheet.ws.cell(row=2, column=4).fill.fgColor.rgb == STATUS_COLORS["100_БЗ"]


def test_apply_payment_values_matches_conversion_status_despite_whitespace(base_workbook):
    """Реальний ризик: клітинка ОБЛІК зі статусом "БПШП " (зайвий пробіл,
    вручну введений текст) - раніше точний (не нормалізований) пошук у
    PAYMENT_STATUS_CONVERSIONS мовчки НЕ спрацьовував би, і клітинка
    лишалась би незміненою, без жодного повідомлення про помилку."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="БПШП ")

    unresolved = ot.apply_payment_values(sheet, payment_values={})

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "100_БПШП"


def test_apply_payment_values_matches_override_value_despite_case_difference(base_workbook):
    """Реальний ризик: файл information_unit подає "бпшп" (інший регістр,
    ніж "БПШП" - ключ PAYMENT_STATUS_OVERRIDES["ШП"]) - без нормалізації
    override мовчки НЕ спрацював би."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ШП")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): ["бпшп"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "100_БПШП"


def test_apply_payment_values_treats_case_varying_duplicate_values_as_agreement(base_workbook):
    """Реальний випадок: ДВА файли information_unit подають ТОЙ САМИЙ статус
    РІЗНИМ регістром ("БПШП" і "бпшп") - це ТА САМА відповідь, а не
    суперечність (на відміну від СПРАВДІ різних значень, де override не
    застосовується, а лишається на числову гілку/ручну перевірку)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ШП")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): ["БПШП", "бпшп"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "100_БПШП"


def test_apply_payment_values_flags_missing_payment_data(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")
    sheet.ws.cell(row=2, column=5, value="РВЗ")

    unresolved = ot.apply_payment_values(sheet, payment_values={})

    assert len(unresolved) == 1
    assert "Немає категорії виплати" in unresolved[0]["reason"]
    assert sheet.ws.cell(row=2, column=5).value == "РВЗ"  # не займаний


def test_apply_payment_values_fills_a_blank_cell_for_a_newly_arrived_person(base_workbook):
    """Реальний випадок, підтверджений користувачем: новоприбулу людину
    додано до роcтера, але вона В РАПОРТАХ ВЗАГАЛІ НЕ ФІГУРУЄ - усі
    колонки-дати лишились ПОРОЖНІМИ (не "РВЗ" - буквально None), лише
    information_unit подає для неї реальні дані. Порожня клітинка ТЕПЕР
    отримує категорію виплати ТИМ САМИМ шляхом, що й "РВЗ" - без жодного
    "усі колонки порожні" запису, коли ХОЧА Б ОДНА дата щось отримала."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ - не "РВЗ", щоб не заважати перевірці нижче
    sheet.ws.append(["Посада3", "матрос", "ТРЕТІЙ Третій Третій", None, None, None, None, ""])
    payment_values = {(ot.normalize_name("ТРЕТІЙ Третій Третій"), date(2026, 8, 1)): [100]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=4, column=5).value == 100  # ТРЕТІЙ, 01.08 - підставлено з information_unit
    assert sheet.ws.cell(row=4, column=4).value is None  # 31.07 - і далі порожньо (немає даних саме на цю дату)


def test_apply_payment_values_applies_default_status_override_to_a_blank_cell(base_workbook):
    """Реальний випадок (баг, підтверджений користувачем на реальних даних):
    новоприбула людина, чия клітинка ПОРОЖНЯ (не "РВЗ"), а information_unit
    подає "НОВ" - замінник, який ВЖЕ працює для "РВЗ"
    (PAYMENT_STATUS_OVERRIDES[DEFAULT_STATUS]["НОВ"] = 10, той самий шлях, що
    й "БЗВП"/"Адап"). Порожня клітинка МАЄ шукати той самий override_map, що
    й "РВЗ" (не "свій" відсутній запис для None) - інакше "НОВ" мовчки
    "падав" би в "не єдине числове значення" на КОЖНУ дату, попри те, що це
    ОДНОЗНАЧНИЙ, відомий замінник."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ - не "РВЗ", щоб не заважати перевірці нижче
    sheet.ws.append(["Посада3", "матрос", "ТРЕТІЙ Третій Третій", None, None, None, None, ""])
    payment_values = {(ot.normalize_name("ТРЕТІЙ Третій Третій"), date(2026, 8, 1)): ["НОВ"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=4, column=5).value == 10


def test_apply_payment_values_applies_conversion_to_a_blank_cell(base_workbook):
    """Реальний випадок (баг, підтверджений користувачем на реальних
    даних): новоприбула людина, чия клітинка ПОРОЖНЯ - information_unit
    подає "ВПБП" (ключ PAYMENT_STATUS_CONVERSIONS, не PAYMENT_STATUS_
    OVERRIDES) - конвертується на "100_ВПБП", той самий факт, що й коли
    роcтерова клітинка сама писала б "ВПБП" напряму."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ - не "РВЗ", щоб не заважати перевірці нижче
    sheet.ws.append(["Посада3", "матрос", "ТРЕТІЙ Третій Третій", None, None, None, None, ""])
    payment_values = {(ot.normalize_name("ТРЕТІЙ Третій Третій"), date(2026, 8, 1)): ["ВПБП"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=4, column=5).value == "100_ВПБП"


def test_apply_payment_values_applies_an_unmapped_status_text_to_a_blank_cell_silently(base_workbook):
    """Реальний випадок (баг, підтверджений користувачем на реальних
    даних): новоприбула людина, чия клітинка ПОРОЖНЯ - information_unit
    ОДНОСТАЙНО подає "СЗЧ" - НЕ число, НЕ ключ PAYMENT_STATUS_CONVERSIONS/
    PAYMENT_STATUS_OVERRIDES[DEFAULT_STATUS], АЛЕ сам по собі - змістовний,
    розпізнаваний статус (STATUS_COLORS). Для ПОРОЖНЬОЇ клітинки - за
    прямою вказівкою користувача - застосовується НАПРЯМУ, буквальним
    текстом, БЕЗ жодного запису на ручну перевірку (на відміну від "РВЗ",
    де той самий випадок і далі лишається на ручну перевірку - див. тест
    нижче)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ - не "РВЗ", щоб не заважати перевірці нижче
    sheet.ws.append(["Посада3", "матрос", "ТРЕТІЙ Третій Третій", None, None, None, None, ""])
    payment_values = {(ot.normalize_name("ТРЕТІЙ Третій Третій"), date(2026, 8, 1)): ["СЗЧ"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=4, column=5).value == "СЗЧ"


def test_apply_payment_values_leaves_a_blank_cell_untouched_on_a_genuine_conflict(base_workbook):
    """Той самий "порожня клітинка" сценарій, але information_unit подає
    ДВА РІЗНІ, суперечливі значення для ОДНІЄЇ дати ("ВП" і "СЗЧ" - ні одне
    з них не "перемагає", жодного запису PAYMENT_STATUS_PRIORITY для цієї
    пари) - немає підстави вибрати ЯКЕСЬ одне, клітинка лишається
    порожньою, і, за прямою вказівкою користувача, БЕЗ запису на ручну
    перевірку на ЦЮ дату (як і "не вистачає даних" вище)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ - не "РВЗ", щоб не заважати перевірці нижче
    sheet.ws.append(["Посада3", "матрос", "ТРЕТІЙ Третій Третій", None, None, None, None, ""])
    payment_values = {(ot.normalize_name("ТРЕТІЙ Третій Третій"), date(2026, 8, 1)): ["ВП", "СЗЧ"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert len(unresolved) == 1  # лише підсумковий "усі колонки порожні"
    assert "порожні" in unresolved[0]["reason"]
    assert sheet.ws.cell(row=4, column=5).value is None


def test_apply_payment_values_reports_a_person_whose_entire_row_stays_empty(base_workbook):
    """Той самий новоприбулий сценарій, але information_unit ТЕЖ не подає
    ЖОДНОГО значення для ЖОДНОЇ дати - ось тоді дійсно нема на чому
    будувати подальшу обробку, тож ОДИН явний запис на ручну перевірку
    (а НЕ по одному запису "Немає категорії виплати" на КОЖНУ з 4 порожніх
    колонок-дат - надто галасливо)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ - не "РВЗ", щоб не заважати перевірці нижче
    sheet.ws.append(["Посада3", "матрос", "ТРЕТІЙ Третій Третій", None, None, None, None, ""])

    unresolved = ot.apply_payment_values(sheet, payment_values={})

    assert len(unresolved) == 1
    assert "ТРЕТІЙ Третій Третій" in unresolved[0]["reason"]
    assert "порожні" in unresolved[0]["reason"]
    assert sheet.ws.cell(row=2, column=4).value == "ВП"  # ПЕРШИЙ не постраждав


def test_apply_payment_values_flags_conflicting_numbers_from_different_files(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")
    sheet.ws.cell(row=2, column=5, value="РВЗ")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [100, 30]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert len(unresolved) == 1
    assert "не єдине числове значення" in unresolved[0]["reason"]
    assert sheet.ws.cell(row=2, column=5).value == "РВЗ"


def test_apply_payment_values_resolves_bz_vs_rozp_conflict_using_priority(base_workbook):
    """PAYMENT_STATUS_PRIORITY (constants.py, керовано користувачем
    самостійно) - реальний випадок, підтверджений користувачем: РІЗНІ файли
    information_unit подають суперечливі текстові статуси "БЗ" і "Розп" для
    ОДНІЄЇ людини й дати - на відміну від звичайної суперечності
    (test_apply_payment_values_flags_conflicting_numbers_from_different_files
    вище), ЦЯ пара МАЄ визначеного переможця ("БЗ") -
    _resolve_value_priority_conflict РОЗВ'ЯЗУЄ конфлікт автоматично ДО
    подальшої обробки, тож unresolved-запис (якщо він таки трапиться) згадує
    ЛИШЕ переможця "БЗ", а не обидва сирі суперечливі значення. "БЗ" сам по
    собі НЕ зареєстрований ключ PAYMENT_STATUS_OVERRIDES[DEFAULT_STATUS]
    (лише "БЗВП" там є) і не число - тож ОСТАТОЧНО все одно лишається на
    ручну перевірку, а не підставляється мовчки в "РВЗ"."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")
    sheet.ws.cell(row=2, column=5, value="РВЗ")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): ["БЗ", "Розп"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert len(unresolved) == 1
    assert "['БЗ']" in unresolved[0]["reason"]
    assert "Розп" not in unresolved[0]["reason"]
    assert sheet.ws.cell(row=2, column=5).value == "РВЗ"


def test_apply_payment_values_flags_non_numeric_value(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")
    sheet.ws.cell(row=2, column=5, value="РВЗ")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): ["ВП"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert len(unresolved) == 1
    assert "не єдине числове значення" in unresolved[0]["reason"]
    assert sheet.ws.cell(row=2, column=5).value == "РВЗ"


def test_apply_payment_values_applies_when_duplicate_files_agree(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")
    sheet.ws.cell(row=2, column=5, value="РВЗ")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [100, 100]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == 100


def test_apply_payment_values_overrides_status_when_information_unit_gives_a_more_specific_reason(base_workbook):
    """PAYMENT_STATUS_OVERRIDES (constants.py) - реальний випадок:
    "ВД" (відрядження) в ОБЛІК, а файл information_unit
    ОДНОЗНАЧНО каже "Навч" (навчання) для тієї самої дати - підрозділ
    уточнює РЕАЛЬНУ причину, тож "ВД" замінюється на "Навч"."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВД")  # ПЕРШИЙ, 31.07
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): ["Навч"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "Навч"


def test_apply_payment_values_leaves_vd_unchanged_for_unmapped_information_unit_value(base_workbook):
    """PAYMENT_STATUS_OVERRIDES["ВД"] (constants.py) має лише "Навч"/"ВД" як
    визнані замінники - "БЗВП" серед них НЕМАЄ. Для "ВД"-подібних статусів
    (на відміну від DEFAULT_STATUS) apply_payment_values навмисно М'ЯКШИЙ:
    значення, якого немає серед ключів мапи, лишає статус як є, БЕЗ жодного
    запису на ручну перевірку ("ВД" сам собою не потребує уваги, це не "ще
    не визначений" стан) - на відміну від "РВЗ", де таке саме значення
    потрапило б в unresolved."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВД")  # ПЕРШИЙ, 31.07
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): ["БЗВП"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "ВД"


def test_apply_payment_values_overrides_status_colors_cell_with_the_new_statuss_color(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВД")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): ["Навч"]}

    ot.apply_payment_values(sheet, payment_values)

    assert sheet.ws.cell(row=2, column=4).fill.fgColor.rgb == STATUS_COLORS["Навч"]


def test_apply_payment_values_leaves_override_eligible_status_untouched_without_matching_data(base_workbook):
    """Немає даних information_unit узагалі для цієї людини й дати - "ВД"
    лишається як є, БЕЗ жодного запису на ручну перевірку (на відміну від
    DEFAULT_STATUS - "ВД" сам собою не є "ще не визначеним" станом, що
    потребує уваги)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВД")

    unresolved = ot.apply_payment_values(sheet, payment_values={})

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "ВД"


def test_apply_payment_values_leaves_override_eligible_status_untouched_when_value_is_not_accepted(base_workbook):
    """Файл information_unit подає щось ІНШЕ, ніж прийнятний замінник для
    "ВД" (PAYMENT_STATUS_OVERRIDES["ВД"] - лише "Навч") - "ВД" лишається як
    є, без заміни й без запису на ручну перевірку."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВД")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): ["ШП"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "ВД"


def test_apply_payment_values_leaves_override_eligible_status_untouched_on_conflicting_data(base_workbook):
    """Кілька РІЗНИХ значень з різних файлів (навіть якщо ОДНЕ з них -
    прийнятний замінник) - неоднозначно, "ВД" лишається як є."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВД")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): ["Навч", "ШП"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == "ВД"


def test_apply_payment_values_overrides_default_status_with_bzvp_to_the_number_10(base_workbook):
    """PAYMENT_STATUS_OVERRIDES[DEFAULT_STATUS] (constants.py) - "РВЗ" (ще
    не визначено), а файл information_unit ОДНОЗНАЧНО каже "БЗВП" для тієї
    самої дати - це, по суті, ТА САМА причина, що й "ППД" (категорія виплати
    10), тож замінюється на ЧИСЛО 10 (а не на буквальний текст "БЗВП"),
    ЗАМІСТЬ запису на ручну перевірку. Заливка - за DEFAULT_STATUS (як і в
    числовій гілці), бо саме число кольору не має."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value="РВЗ")  # ПЕРШИЙ, 01.08
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): ["БЗВП"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == 10
    assert sheet.ws.cell(row=2, column=5).fill.fgColor.rgb == STATUS_COLORS[10]


def test_apply_payment_values_overrides_default_status_with_adap_to_the_number_10(base_workbook):
    """Реальні файли resources/information_unit пишуть цей статус САМЕ як
    "Адап" (без крапки, на відміну від "Адап." зі STATUS_MEANINGS/
    STATUS_COLORS) - приймається так, як реально надходить, і, так само як
    "БЗВП", замінюється на ЧИСЛО 10 (та сама причина - категорія виплати),
    а не на буквальний текст "Адап"."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value="РВЗ")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): ["Адап"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == 10
    assert sheet.ws.cell(row=2, column=5).fill.fgColor.rgb == STATUS_COLORS[10]


def test_apply_payment_values_overrides_default_status_with_nov_to_the_number_10(base_workbook):
    """PAYMENT_STATUS_OVERRIDES[DEFAULT_STATUS]["НОВ"] = 10 - той самий шлях,
    що й "БЗВП"/"Адап" (категорія виплати 10), а НЕ буквальний текст "НОВ"."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value="РВЗ")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): ["НОВ"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == 10
    assert sheet.ws.cell(row=2, column=5).fill.fgColor.rgb == STATUS_COLORS[10]


@pytest.mark.parametrize("status", ["ЗВФН", "ЛМР"])
def test_apply_payment_values_overrides_default_status_with_literal_text(base_workbook, status):
    """На відміну від "БЗВП"/"Адап"/"НОВ" (замінюються на ЧИСЛО 10) - "ЗВФН"
    і "ЛМР" замінюють "РВЗ" на буквальний ТЕКСТ (та сама логіка, що й
    "ВД" -> "Навч") - PAYMENT_STATUS_OVERRIDES[DEFAULT_STATUS] мапить КОЖЕН
    ключ на ВЛАСНЕ значення для запису, не всі на одне й те саме."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value="РВЗ")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [status]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == status
    assert sheet.ws.cell(row=2, column=5).fill.fgColor.rgb == STATUS_COLORS[status]


@pytest.mark.parametrize("information_unit_status", ["ПРВД", "Перевд.", "Првд", "Првд."])
def test_apply_payment_values_overrides_default_status_with_prvd_regardless_of_spelling(base_workbook, information_unit_status):
    """Реальний випадок (баг, підтверджений користувачем): новоприбула
    людина (порожня клітинка - той самий шлях, що й "РВЗ") - information_unit
    ОДНОЗНАЧНО каже "ПРВД" (переведений в інший підрозділ) чи один із ТИХ
    САМИХ варіантів написання, що й PAYMENT_STATUS_OVERRIDES["ПРВД"]
    ("Перевд."/"Првд"/"Првд.") - УСІ ЧОТИРИ ведуть до ОДНОГО канонічного
    тексту "ПРВД" (на відміну від "ПРВД"-keyed мапи, де "Перевд." стає
    "ПРВД", а "Првд"/"Првд." лишаються собою - тут, для НОВОПРИБУЛОЇ людини,
    немає "вже наявного" написання, яке варто зберегти, тож усі варіанти
    зводяться до одного). Раніше - "не єдине числове значення" на КОЖНУ
    дату, попри однозначний, відомий текстовий замінник."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ - не "РВЗ", щоб не заважати перевірці нижче
    sheet.ws.append(["Посада3", "матрос", "ТРЕТІЙ Третій Третій", None, None, None, None, ""])
    payment_values = {(ot.normalize_name("ТРЕТІЙ Третій Третій"), date(2026, 8, 1)): [information_unit_status]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=4, column=5).value == "ПРВД"


def test_apply_payment_values_still_applies_a_number_when_default_status_has_no_text_override(base_workbook):
    """Регресія: "РВЗ" + ЧИСЛО (без жодного текстового замінника серед
    значень) - і далі застосовується як РАНІШЕ (числова гілка), новий
    override-шлях не мав цього зламати."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value="РВЗ")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [100]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == 100


def test_apply_payment_values_flags_default_status_when_override_text_conflicts_with_a_number(base_workbook):
    """"РВЗ" з ДВОМА різними значеннями від різних файлів - одне число, одне
    - прийнятний текстовий замінник - неоднозначно (справжня суперечність
    між підрозділами), тож лишається на ручну перевірку, а не мовчки
    застосовується одне з двох."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value="РВЗ")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [100, "БЗВП"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert len(unresolved) == 1
    assert "не єдине числове значення" in unresolved[0]["reason"]
    assert sheet.ws.cell(row=2, column=5).value == "РВЗ"


def test_strip_external_links_clears_the_external_links_list():
    """Реальний випадок: resources/ОБЛІК.xlsx має осиротіле зовнішнє
    посилання (Excel External Reference - залишок від колись видаленого
    Data Validation, що вказувала на конкретний, застарілий файл
    information_unit) - openpyxl НЕ вміє коректно перезберегти книгу з таким
    посиланням (Excel потім показує "Знайдено проблему з вмістом" при
    відкритті результату), тож _strip_external_links прибирає його
    БЕЗУМОВНО - жодна формула/правило цим посиланням фактично не
    користується."""
    class _FakeWorkbook:
        _external_links = ["осиротіле посилання"]

    wb = _FakeWorkbook()
    ot._strip_external_links(wb)

    assert wb._external_links == []


def test_save_removes_external_links_before_saving(base_workbook, tmp_path):
    """Timesheet.save() прибирає зовнішні посилання ПЕРЕД self.wb.save() -
    інакше (осиротіле посилання пережило б збереження, як діагностовано на
    реальному resources/ОБЛІК.xlsx) openpyxl впав би з AttributeError,
    намагаючись серіалізувати непридатний запис."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.wb._external_links = ["осиротіле посилання (заглушка тесту)"]
    output_path = tmp_path / "out" / "ОБЛІК.xlsx"

    sheet.save(str(output_path))

    assert load_workbook(str(output_path))._external_links == []


def test_save_preserves_workbook_and_writes_to_new_path(base_workbook, tmp_path):
    sheet = ot.Timesheet(base_workbook, "Табель")
    output_path = tmp_path / "out" / "ОБЛІК.xlsx"

    sheet.save(str(output_path))

    assert output_path.exists()
    saved = load_workbook(str(output_path))
    assert saved["Табель"].cell(row=2, column=3).value == "ПЕРШИЙ Перший Перший"


def test_save_wraps_permission_error_with_a_friendly_message(base_workbook, tmp_path, monkeypatch):
    """Реальний випадок: цільовий файл зараз відкритий у Excel -
    self.wb.save() падає з PermissionError і всипає "сирий" traceback
    (WinError 32 і т.п.). Обгортаємо зрозумілим повідомленням українською -
    той самий текст, що бачить користувач (index.py ловить його й друкує без
    traceback), а НЕ сам закриваємо файл (ризик втратити незбережені зміни)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    output_path = tmp_path / "out" / "ОБЛІК.xlsx"

    def _raise(*args, **kwargs):
        raise PermissionError("[WinError 32] заблоковано іншою програмою")

    monkeypatch.setattr(sheet.wb, "save", _raise)

    with pytest.raises(PermissionError, match="Закрийте його"):
        sheet.save(str(output_path))


def test_export_for_money_project_renames_the_sole_sheet_and_saves_a_separate_copy(base_workbook, tmp_path):
    """За прямою вказівкою користувача - готує вже згенерований "ОБЛІК для
    виплат" для сусіднього проєкту: перейменовує аркуш на точну назву, яку
    той проєкт шукає ("ТРАВЕНЬ"), і зберігає це як ОКРЕМУ копію - джерело
    (source_path, "Табель") лишається незмінним."""
    output_path = tmp_path / "out" / "ОБЛІК.xlsx"

    ot.export_for_money_project(base_workbook, str(output_path), "ТРАВЕНЬ")

    assert output_path.exists()
    saved = load_workbook(str(output_path))
    assert saved.sheetnames == ["ТРАВЕНЬ"]
    assert saved["ТРАВЕНЬ"].cell(row=2, column=3).value == "ПЕРШИЙ Перший Перший"
    # Джерело - НЕ перейменоване, лишається "Табель".
    assert load_workbook(base_workbook).sheetnames == ["Табель"]


def test_export_for_money_project_removes_external_links_before_saving(base_workbook, tmp_path, monkeypatch):
    """Той самий принцип, що й Timesheet.save вище - зовнішні посилання
    (реальний випадок - resources/ОБЛІК.xlsx) прибираються ПЕРЕД збереженням
    копії для money-проєкту, а не переносяться в неї."""
    real_load_workbook = ot.load_workbook

    def _load_with_fake_external_link(*args, **kwargs):
        wb = real_load_workbook(*args, **kwargs)
        wb._external_links = ["осиротіле посилання (заглушка тесту)"]
        return wb

    monkeypatch.setattr(ot, "load_workbook", _load_with_fake_external_link)
    output_path = tmp_path / "out" / "ОБЛІК.xlsx"

    ot.export_for_money_project(base_workbook, str(output_path), "ТРАВЕНЬ")

    assert load_workbook(str(output_path))._external_links == []


def test_export_for_money_project_wraps_permission_error_with_a_friendly_message(base_workbook, tmp_path, monkeypatch):
    """Той самий принцип, що й Timesheet.save вище - цільовий файл зайнятий
    іншою програмою дає зрозуміле повідомлення, а не сирий traceback."""
    from openpyxl import Workbook as _Workbook

    output_path = tmp_path / "out" / "ОБЛІК.xlsx"

    def _raise(*args, **kwargs):
        raise PermissionError("[WinError 32] заблоковано іншою програмою")

    monkeypatch.setattr(_Workbook, "save", _raise)

    with pytest.raises(PermissionError, match="спробуйте ще раз"):
        ot.export_for_money_project(base_workbook, str(output_path), "ТРАВЕНЬ")
