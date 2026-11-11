import unicodedata
import zipfile
from datetime import date, datetime, timedelta

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

import content.oblik_timesheet as ot
from constants import (
    AFTER_HOURS_ARRIVAL_STATUS,
    DEFAULT_STATUS,
    FREEFORM_STATUS_PATTERNS,
    STATUS_COLORS,
    STATUS_CONFLICT_OVERRIDES,
    STATUS_HEADING_RULES,
    TRANSFER_HEADING_MARKER,
    TRANSFER_STATUSES,
)


def _write_workbook(path, header, rows):
    wb = Workbook()
    ws = wb.worksheets[0]
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


def _add_text_custom_filters(path):
    """Автофільтр так, як його пише Excel: "(Непусті)" - notEqual " ", і
    "дорівнює <текст>" - openpyxl сам такий файл не записав би."""
    import zipfile

    with zipfile.ZipFile(path) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    sheet_xml = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    parts["xl/worksheets/sheet1.xml"] = sheet_xml.replace("</sheetData>", (
        '</sheetData><autoFilter ref="A1:H3"><filterColumn colId="7"><customFilters>'
        '<customFilter operator="notEqual" val=" "/><customFilter operator="equal" val="ВП"/>'
        '</customFilters></filterColumn><filterColumn colId="3"><customFilters>'
        '<customFilter operator="greaterThan" val="30"/></customFilters></filterColumn></autoFilter>'
    ), 1).encode("utf-8")
    with zipfile.ZipFile(path, "w") as target:
        for name, data in parts.items():
            target.writestr(name, data)


def test_timesheet_opens_a_book_with_a_text_custom_filter_and_keeps_it(base_workbook, tmp_path):
    """Реальний випадок: фільтр на колонці Табеля в resources/ОБЛІК.xlsx -
    openpyxl 3.1.2 сам по собі відмовлявся відкривати книгу ("Value must be
    either numerical or a string containing a wildcard"), content/
    excel_compat.py - відкривається, фільтр зберігається як був."""
    _add_text_custom_filters(base_workbook)
    output_path = str(tmp_path / "out.xlsx")
    ot.Timesheet(base_workbook, "Табель").save(output_path)

    saved = ot.Timesheet(output_path, "Табель")
    text_filters, number_filters = (column.customFilters.customFilter for column in saved.ws.auto_filter.filterColumn)
    assert [(f.operator, f.val) for f in text_filters] == [("notEqual", " "), ("equal", "ВП")]
    assert [(f.operator, f.val) for f in number_filters] == [("greaterThan", 30.0)]


def test_timesheet_requires_pib_column(tmp_path):
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", ["ПОСАДА", datetime(2026, 8, 1)], [["Посада1", None]])
    with pytest.raises(ValueError, match="ПІБ"):
        ot.Timesheet(path, "Табель")


def test_timesheet_requires_at_least_one_date_column(tmp_path):
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", ["ПОСАДА", "ПІБ"], [["Посада1", "ПЕРШИЙ Перший Перший"]])
    with pytest.raises(ValueError, match="дати"):
        ot.Timesheet(path, "Табель")


def test_timesheet_reports_a_workbook_without_the_tabel_sheet(tmp_path):
    """Реальний випадок: для перевірки обрано книгу без аркуша "Табель"
    (напр. відомість з аркушем місяця) - зрозуміла ValueError, а не
    KeyError, що валив увесь скрипт (index.py ловить ValueError)."""
    wb = Workbook()
    wb.worksheets[0].title = "ЖОВТЕНЬ"
    path = str(tmp_path / "Відомість.xlsx")
    wb.save(path)

    with pytest.raises(ValueError, match="немає аркуша 'Табель'"):
        ot.Timesheet(path, "Табель")


def test_timesheet_reports_a_file_that_is_not_an_excel_workbook(tmp_path):
    path = tmp_path / "нотатки.txt"
    path.write_text("не книга", encoding="utf-8")

    with pytest.raises(ValueError, match="не книга Excel"):
        ot.Timesheet(str(path), "Табель")


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


def test_person_rows_skips_a_row_without_a_name(base_workbook):
    """Рядок без ПІБ (напр. хвіст аркуша з окремим числом) не вважається людиною."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    blank_name_row = sheet.ws.max_row + 1
    sheet.ws.cell(row=blank_name_row, column=5, value=10)

    rows = sheet.person_rows()

    assert ot.normalize_name("ПЕРШИЙ Перший Перший") in rows
    assert blank_name_row not in rows.values()


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
    "БЗ", і жоден рапорт НЕ подає події САМЕ на цю дату - за рішенням користувача, цей вручну вписаний статус ВСЕ ОДНО стає новим "якорем"
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


def test_apply_events_manually_entered_status_overrides_a_conflicting_report_event(base_workbook):
    """Рішення користувача (ЗМІНА попередньої поведінки) - якщо
    на ту САМУ дату resources/ОБЛІК.xlsx ВЖЕ має вручну вписаний статус,
    ВІДМІННИЙ від того, що дає подія з рапорту, - вручну вписаний статус
    МАЄ ВИЩИЙ ПРІОРИТЕТ і перемагає (реальний випадок: людина фактично
    прибула на постійний пункт дислокації, хоча рапорт каже лише загальне
    "прибув в район виконання завдань"). Рішення користувача: ПОВНІСТЮ МОВЧКИ, БЕЗ жодного запису на ручну перевірку (людина вже
    свідомо контролює це значення)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=6, value="БЗ")  # ПЕРШИЙ, 02.08 - вручну вписано
    events = [{"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВП", "effective_date": date(2026, 8, 2)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1), date(2026, 8, 2)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=6).value == "БЗ"  # вручну вписаний статус перемагає подію


def test_apply_events_manually_entered_ppd_overrides_default_status_report_event(base_workbook):
    """Реальний випадок, підтверджений користувачем: рапорт каже, що людина
    "прибула в район виконання завдань" (DEFAULT_STATUS) - типове,
    неточне формулювання, що трапляється в БАГАТЬОХ реальних рапортах - але
    користувач ВРУЧНУ вписав у resources/ОБЛІК.xlsx конкретніший "ППД" (людина
    фактично прибула на постійний пункт дислокації, чого сам текст рапорту
    не уточнює). Вручну вписаний "ППД" МАЄ перемогти, ПОВНІСТЮ МОВЧКИ."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=6, value="ППД")  # ПЕРШИЙ, 02.08 - вручну вписано
    events = [{"pib_raw": "ПЕРШИЙ Перший Перший", "status": DEFAULT_STATUS, "effective_date": date(2026, 8, 2)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1), date(2026, 8, 2)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=6).value == "ППД"


def test_apply_events_manually_entered_status_carries_forward_after_overriding_an_event(base_workbook):
    """Вручну вписаний статус, що переміг подію з рапорту (тест вище), і
    ДАЛІ стає "якорем" - ПРОДОВЖУЄТЬСЯ у наступну (порожню) колонку-дату,
    той самий принцип "продовжує статус", що й для звичайних подій."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=6, value="ППД")  # ПЕРШИЙ, 02.08 - вручну вписано
    events = [{"pib_raw": "ПЕРШИЙ Перший Перший", "status": DEFAULT_STATUS, "effective_date": date(2026, 8, 2)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1), date(2026, 8, 2), date(2026, 8, 3)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=6).value == "ППД"
    assert sheet.ws.cell(row=2, column=7).value == "ППД"  # 03.08 - продовжено


def test_apply_events_continues_a_manually_entered_status_after_the_last_report(base_workbook):
    """Реальний випадок: 31.07 - базовий "РВЗ",
    змінюється за рапортами як завжди, АЛЕ поза діапазоном наявних рапортів
    (тут - лише 01.08) хтось УЖЕ вручну вписав статус ("БЗ") у ПІЗНІШУ
    колонку (03.08, а не одразу наступну 02.08) - за рішенням користувача, цей статус ПРОДОВЖУЄТЬСЯ (forward-fill) у ВСІ НАСТУПНІ
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


def test_apply_events_forward_fills_tail_status_without_a_color_as_plain_value(base_workbook, monkeypatch):
    """Ручний статус ПІСЛЯ останнього рапорту, для якого немає кольору, продовжується
    як звичайне значення - без заливки і без падіння."""
    monkeypatch.setattr(ot, "_fill_for_status", lambda status: None)
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=7, value="БЗ")  # ПЕРШИЙ, 03.08 - вручну вписано наперед

    unresolved = ot.apply_events(sheet, [], report_dates=[date(2026, 8, 1)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=6).value is None  # 02.08 - поза рапортами, без якоря
    assert sheet.ws.cell(row=2, column=7).value == "БЗ"


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
    """Рішення користувача: КОЖНА категорія виплати (число) має
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
        | (set(STATUS_HEADING_RULES.values()) - {TRANSFER_HEADING_MARKER})
        | set(TRANSFER_STATUSES)  # заголовок "Переведені:" дає один із них (за колонкою "Куди вибув")
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
    """Рішення користувача: КОЖНА категорія виплати має
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
    прізвищі порівняно з роcтеровим написанням - за рішенням користувача, подія ВСЕ ОДНО ЗАСТОСОВУЄТЬСЯ до цього єдиного схожого
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
    """"ВП" і "Розп" - НАВМИСНО пара, де жоден автоматичний механізм (ні
    STATUS_CONFLICT_OVERRIDES, ні _automatic_conflict_winner - ЖОДЕН із
    двох не оплачується за замовчуванням і ЖОДЕН не в STATUS_CONFLICT_
    PRIORITY_EXEMPT) не дає відповіді - демонструє запасний шлях "дійсно
    невідомо, лишається на ручну перевірку"."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВП", "effective_date": date(2026, 8, 1)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "Розп", "effective_date": date(2026, 8, 1)},
    ]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert len(unresolved) == 1
    assert "суперечливі статуси" in unresolved[0]["reason"]
    assert sheet.ws.cell(row=2, column=5).value == "Розп"  # немає переможця для цієї пари - останній перемагає


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


def test_apply_events_treats_vpsz_ppd_as_a_genuine_conflict(base_workbook):
    """Рішення користувача: "ВПСЗ" (прогноз відпустки з
    "потребує відпустки ...") і "ППД" (фактичне прибуття на постійний пункт
    дислокації) для однієї людини/дати - НЕ зареєстрована в
    STATUS_CONFLICT_OVERRIDES пара (на відміну від "ВП"/"ВПСЗ") - справжня
    суперечність, лишається на ручну перевірку, останній перемагає."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВПСЗ", "effective_date": date(2026, 8, 1)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ППД", "effective_date": date(2026, 8, 1)},
    ]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert len(unresolved) == 1
    assert "суперечливі статуси" in unresolved[0]["reason"]
    assert sheet.ws.cell(row=2, column=5).value == "ППД"


_LEAVES = ["ВП", "ВПС", "ВПСЗ", "ВПБП"]


def _statuses_after_events(tmp_path, *statuses, manual=None):
    """Статуси ПЕРШОГО за 01.08... після apply_events: перший - базовий стан, далі
    подія кожного дня (None - подій немає). manual - {номер дня: вписаний вручну}."""
    manual = manual or {}
    days = len(statuses)
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [statuses[0]] + [manual.get(day) for day in range(2, days + 1)], days=days)
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": status, "effective_date": date(2026, 8, day)}
        for day, status in enumerate(statuses, start=1) if status and day > 1
    ]
    ot.apply_events(sheet, events, report_dates=[date(2026, 8, day) for day in range(1, days + 1)])
    return [sheet.ws.cell(row=2, column=col).value for col, _day in sheet.date_columns]


@pytest.mark.parametrize("leave", _LEAVES)
def test_vlk_between_days_of_a_leave_stays_the_leave(tmp_path, leave):
    """Рішення користувача: ВЛК під час відпустки не оплачується, періоди йдуть
    паралельно, і відпустка будь-якого типу важливіша: ВП - ВЛК - ВП це ВП."""
    assert _statuses_after_events(tmp_path, leave, "ВЛК", leave) == [leave, leave, leave]


def test_several_vlk_days_between_days_of_a_leave_stay_the_leave(tmp_path):
    assert _statuses_after_events(tmp_path, "ВПСЗ", "ВЛК", None, None, "ВПСЗ") == ["ВПСЗ"] * 5


def test_vlk_between_two_different_leaves_continues_the_first_one(tmp_path):
    assert _statuses_after_events(tmp_path, "ВПСЗ", "ВЛК", "ВП") == ["ВПСЗ", "ВПСЗ", "ВП"]


def test_vlk_day_inside_a_leave_gets_the_color_of_the_leave(tmp_path):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", ["ВПСЗ", None, None], days=3)
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВЛК", "effective_date": date(2026, 8, 2)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВПСЗ", "effective_date": date(2026, 8, 3), "is_forecast": True},
    ]

    ot.apply_events(sheet, events, report_dates=[date(2026, 8, day) for day in (1, 2, 3)])

    cell = sheet.ws.cell(row=2, column=5)
    assert cell.value == "ВПСЗ"
    assert cell.fill.start_color.rgb == STATUS_COLORS["ВПСЗ"]


@pytest.mark.parametrize("statuses", [
    ("РВЗ", "ВЛК", "ВПСЗ"),  # відпустка лише після ВЛК
    ("ВПСЗ", "ВЛК", "РВЗ"),  # відпустка лише перед ВЛК
    ("ВПСЗ", "ВЛК"),  # що далі, ще невідомо
    ("ШП", "ВЛК", "ШП"),
    ("ВПСЗ", "ВЛК", "РВЗ", "ВПСЗ", "ВПСЗ"),
    ("ВЛК", "ВПСЗ"),  # базова колонка не змінюється
])
def test_vlk_stays_when_it_is_not_between_days_of_a_leave(tmp_path, statuses):
    assert _statuses_after_events(tmp_path, *statuses) == list(statuses)


def test_other_status_between_days_of_a_leave_stays(tmp_path):
    """Правило лише про ВЛК: лікування посеред відпустки лишається лікуванням."""
    assert _statuses_after_events(tmp_path, "ВПСЗ", "ШП", "ВПСЗ") == ["ВПСЗ", "ШП", "ВПСЗ"]


def test_vlk_written_by_hand_between_days_of_a_leave_stays(tmp_path):
    """Вписане вручну в resources/ОБЛІК.xlsx важливіше за правила (рішення користувача)."""
    assert _statuses_after_events(tmp_path, "ВПСЗ", None, "ВПСЗ", manual={2: "ВЛК"}) == ["ВПСЗ", "ВЛК", "ВПСЗ"]


@pytest.mark.parametrize("leave", _LEAVES)
@pytest.mark.parametrize("first_is_vlk", [True, False])
def test_leave_wins_a_same_day_conflict_with_vlk(tmp_path, leave, first_is_vlk):
    """Рішення користувача: ВЛК під час відпустки не оплачується, тож правило "хто
    оплачується" тут більше не діє."""
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", ["РВЗ", None], days=2)
    statuses = ["ВЛК", leave] if first_is_vlk else [leave, "ВЛК"]
    events = [{"pib_raw": "ПЕРШИЙ Перший Перший", "status": status, "effective_date": date(2026, 8, 2)} for status in statuses]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1), date(2026, 8, 2)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == leave


def test_automatic_conflict_winner_prefers_the_paid_status_over_unpaid():
    """За ШП належить виплата, а ВПСЗ сама по собі не оплачується, тож ШП
    автоматично перемагає, незалежно від порядку аргументів. ВЛК проти відпустки
    сюди не доходить: цю пару вирішує STATUS_CONFLICT_OVERRIDES."""
    assert ot._automatic_conflict_winner("ВПСЗ", "ШП") == "ШП"
    assert ot._automatic_conflict_winner("ШП", "ВПСЗ") == "ШП"


def test_automatic_conflict_winner_returns_none_when_both_statuses_are_paid():
    """"ШП" і "ВЛК" - ОБИДВА оплачуються за замовчуванням - автоматичне
    правило НЕ має сенсу тут (яка саме причина відсутності правильна -
    питання до людини, не до платіжної категорії)."""
    assert ot._automatic_conflict_winner("ШП", "ВЛК") is None


def test_automatic_conflict_winner_returns_none_when_neither_status_is_paid():
    """"ВП" і "ВПСЗ" НЕ в STATUS_CONFLICT_PAID_BY_DEFAULT (ця пара й так уже
    має власний запис у STATUS_CONFLICT_OVERRIDES, перевірений РАНІШЕ - тут
    перевіряється САМА функція, напряму, без цього раннього запису)."""
    assert ot._automatic_conflict_winner("ВП", "ВПСЗ") is None


def test_automatic_conflict_winner_returns_none_when_either_status_is_exempt():
    """"СЗЧ" - у STATUS_CONFLICT_PRIORITY_EXEMPT - автоматичне правило
    НІКОЛИ не застосовується, коли він бере участь у суперечності, навіть
    проти статусу, що оплачується ("ВЛК") - адміністративний/дисциплінарний
    статус, а не "хто більше заробляє"."""
    assert ot._automatic_conflict_winner("СЗЧ", "ВЛК") is None
    assert ot._automatic_conflict_winner("ВЛК", "СЗЧ") is None


def test_apply_events_resolves_vpsz_shp_conflict_via_automatic_payment_priority(base_workbook):
    """Наскрізний (apply_events) випадок: "ШП" автоматично перемагає "ВПСЗ"
    (STATUS_CONFLICT_PAID_BY_DEFAULT) без запису на ручну перевірку: цієї пари
    немає в статичному STATUS_CONFLICT_OVERRIDES."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВПСЗ", "effective_date": date(2026, 8, 1)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ШП", "effective_date": date(2026, 8, 1)},
    ]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == "ШП"


def test_apply_events_automatic_priority_never_overrides_an_exempt_status(base_workbook):
    """Реальний випадок, підтверджений користувачем: "СЗЧ" (у STATUS_
    CONFLICT_PRIORITY_EXEMPT) НЕ повинен поступатись статусу, що оплачується
    ("ВЛК"), лише тому, що той оплачується - лишається СПРАВЖНЬОЮ
    суперечністю на ручну перевірку, так само як до цієї можливості."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "СЗЧ", "effective_date": date(2026, 8, 1)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВЛК", "effective_date": date(2026, 8, 1)},
    ]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert len(unresolved) == 1
    assert "суперечливі статуси" in unresolved[0]["reason"]
    assert sheet.ws.cell(row=2, column=5).value == "ВЛК"  # немає переможця - останній перемагає, як і раніше


def test_apply_events_manual_conflict_table_still_wins_over_automatic_priority(base_workbook):
    """КРИТИЧНИЙ регресійний тест - "не поламавши логіку", за прямою
    вимогою користувача: {РВЗ,ВЛК}->РВЗ (STATUS_CONFLICT_OVERRIDES) - давно
    підтверджений виняток (комісія не забрала весь день, людина того самого
    дня повернулась) - "ВЛК" оплачується, а РВЗ (DEFAULT_STATUS) - ні, тож
    САМА ПО СОБІ автоматична формула дала б ПРОТИЛЕЖНУ, неправильну
    відповідь ("ВЛК"). Статична таблиця МАЄ перевірятись і перемагати
    ПЕРШОЮ - цей реальний випадок не повинен зламатись."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВЛК", "effective_date": date(2026, 8, 1)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": DEFAULT_STATUS, "effective_date": date(2026, 8, 1)},
    ]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == DEFAULT_STATUS


def test_apply_events_a_real_fact_always_beats_a_forecast_regardless_of_order(base_workbook):
    """Реальний випадок, підтверджений користувачем: рапорт передбачає
    відпустку "наступного дня" після ВЛК (is_forecast=True - сама дата
    відпустки НЕ прочитана з тексту, лише ОБЧИСЛЕНА), а ПІЗНІШИЙ рапорт
    фактично повідомляє, що людина того самого дня повернулась на ППД, а
    відпустка почалась лише кількома днями пізніше - справжній ФАКТ (без
    is_forecast) МАЄ перемогти прогноз, БЕЗ запису на ручну перевірку,
    незалежно від того, який із двох прийшов першим."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    forecast = {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВПСЗ", "effective_date": date(2026, 8, 1), "is_forecast": True}
    fact = {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ППД", "effective_date": date(2026, 8, 1)}

    unresolved = ot.apply_events(sheet, [forecast, fact], report_dates=[date(2026, 8, 1)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == "ППД"

    sheet = ot.Timesheet(base_workbook, "Табель")
    unresolved = ot.apply_events(sheet, [fact, forecast], report_dates=[date(2026, 8, 1)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == "ППД"


def test_apply_events_resolves_same_day_vlk_round_trip_to_default_status(base_workbook):
    """Реальний випадок, підтверджений користувачем: ОДИН рапорт містить і
    "Вибув ... до [ВЛК]" (заголовок "На проходженн...", статус "ВЛК"), і
    "Прибув ... з [ВЛК]" (заголовок "З проходженн...", DEFAULT_STATUS) - З
    ТІЄЮ САМОЮ датою - людина вибула на комісію й ОДРАЗУ Ж того самого дня
    повернулась - увесь день МАЄ рахуватись звичайним "у районі виконання
    завдань" (DEFAULT_STATUS), а не "ВЛК", незалежно від порядку появи в
    рапорті, без запису на ручну перевірку."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВЛК", "effective_date": date(2026, 8, 1)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": DEFAULT_STATUS, "effective_date": date(2026, 8, 1)},
    ]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == DEFAULT_STATUS

    sheet = ot.Timesheet(base_workbook, "Табель")
    unresolved = ot.apply_events(sheet, list(reversed(events)), report_dates=[date(2026, 8, 1)])

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == DEFAULT_STATUS


def test_apply_events_resolves_same_day_vlk_completion_and_ppd_arrival_to_ppd(base_workbook):
    """Реальний випадок, підтверджений користувачем: рапорт одного дня -
    "ДАТА закінчив проходження ВЛК" (ВЛК), рапорт наступного дня - таблиця
    "З проходження ВЛК:" з "Прибув до пункту постійної дислокації" (ППД) з
    ТІЄЮ САМОЮ датою - день і подальші дні МАЮТЬ бути "ППД" (той самий
    принцип, що й {РВЗ,ВЛК}->РВЗ), незалежно від порядку, без ручної
    перевірки - а не "ВЛК" за правилом "хто оплачується"."""
    events = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВЛК", "effective_date": date(2026, 8, 1)},
        {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ППД", "effective_date": date(2026, 8, 1)},
    ]
    for ordered in (events, list(reversed(events))):
        sheet = ot.Timesheet(base_workbook, "Табель")
        unresolved = ot.apply_events(sheet, ordered, report_dates=[date(2026, 8, 1), date(2026, 8, 2)])

        assert unresolved == []
        assert sheet.ws.cell(row=2, column=5).value == "ППД"
        assert sheet.ws.cell(row=2, column=6).value == "ППД"


def test_apply_events_applies_an_overridden_after_hours_event_from_the_next_day(base_workbook):
    """Реальний випадок, підтверджений користувачем: на 01.08 вручну "ППД",
    рапорт - "прибув в район виконання завдань 01.08 після 17:00" (РВЗ) -
    01.08 лишається вручну вписаним "ППД", з 02.08 - РВЗ. Без "після 17:00"
    вручну вписаний статус, як і раніше, продовжується далі; вручну вписаний
    статус і на наступний день теж перемагає."""
    def run(after_hours, manual_next_day=None):
        sheet = ot.Timesheet(base_workbook, "Табель")
        sheet.ws.cell(row=2, column=4, value="ВЛК")
        sheet.ws.cell(row=2, column=5, value="ППД")
        sheet.ws.cell(row=2, column=6, value=manual_next_day)
        event = {"pib_raw": "ПЕРШИЙ Перший Перший", "status": DEFAULT_STATUS, "effective_date": date(2026, 8, 1)}
        if after_hours:
            event["after_hours"] = True
        unresolved = ot.apply_events(sheet, [event], report_dates=[date(2026, 8, 1), date(2026, 8, 2), date(2026, 8, 3)])
        assert unresolved == []
        return [sheet.ws.cell(row=2, column=c).value for c in (5, 6, 7)]

    assert run(after_hours=True) == ["ППД", DEFAULT_STATUS, DEFAULT_STATUS]
    assert run(after_hours=False) == ["ППД", "ППД", "ППД"]
    assert run(after_hours=True, manual_next_day="ППД") == ["ППД", "ППД", "ППД"]


_EVENING_ARRIVAL = {
    "pib_raw": "ПЕРШИЙ Перший Перший", "status": DEFAULT_STATUS, "effective_date": date(2026, 8, 1), "after_hours": True,
}
_ARRIVAL_WEEK = [date(2026, 8, 1), date(2026, 8, 2), date(2026, 8, 3)]


def _days_after_arrival(sheet):
    return [sheet.ws.cell(row=2, column=col).value for col in (5, 6, 7)]


@pytest.mark.parametrize("previous_status", ["ВЛК", "ВД", "Навч", "ППД", "ШП"])
def test_apply_events_evening_arrival_is_ppd_that_day_and_duty_from_the_next(base_workbook, previous_status):
    """Рапорт складається станом на 17 годину, тож хто прибув у район
    виконання завдань "після 17.00", того дня ще не там: день - ППД, РВЗ - з
    наступного дня (рішення користувача). Звідки людина прибула, не важить."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value=previous_status)

    assert ot.apply_events(sheet, [_EVENING_ARRIVAL], _ARRIVAL_WEEK) == []

    assert _days_after_arrival(sheet) == [AFTER_HOURS_ARRIVAL_STATUS, DEFAULT_STATUS, DEFAULT_STATUS]
    assert sheet.after_hours_arrivals == {(2, date(2026, 8, 1))}


def test_apply_events_same_evening_arrival_gives_the_same_days_with_or_without_a_manual_cell(base_workbook):
    """Реальний випадок: двоє людей з однаковим записом рапорту ("дата після
    17.00, прибув з ВЛК"). Одному ППД вписали вручну, і він мав ППД, РВЗ, а
    другий без вписаного ППД одразу отримав РВЗ."""
    def days(manual):
        sheet = ot.Timesheet(base_workbook, "Табель")
        sheet.ws.cell(row=2, column=4, value="ВЛК")
        sheet.ws.cell(row=2, column=5, value=manual)
        ot.apply_events(sheet, [_EVENING_ARRIVAL], _ARRIVAL_WEEK)
        return _days_after_arrival(sheet), sheet.after_hours_arrivals

    assert days(manual=AFTER_HOURS_ARRIVAL_STATUS) == days(manual=None)


def test_apply_events_other_manual_status_on_an_evening_arrival_day_still_wins(base_workbook):
    """Вписане вручну важливіше за рапорт, і таким днем людина керує сама: ППД
    тут немає, тож і позначки для звірки з книгами немає."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=5, value="ВЛК")

    ot.apply_events(sheet, [_EVENING_ARRIVAL], _ARRIVAL_WEEK)

    assert _days_after_arrival(sheet) == ["ВЛК", DEFAULT_STATUS, DEFAULT_STATUS]
    assert sheet.after_hours_arrivals == set()


def test_apply_events_event_of_the_next_day_beats_the_evening_arrival(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    next_day = {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ШП", "effective_date": date(2026, 8, 2)}

    ot.apply_events(sheet, [_EVENING_ARRIVAL, next_day], _ARRIVAL_WEEK)

    assert _days_after_arrival(sheet) == [AFTER_HOURS_ARRIVAL_STATUS, "ШП", "ШП"]


def test_apply_events_arrival_without_a_time_is_duty_at_once(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    arrival = {key: value for key, value in _EVENING_ARRIVAL.items() if key != "after_hours"}

    ot.apply_events(sheet, [arrival], _ARRIVAL_WEEK)

    assert _days_after_arrival(sheet) == [DEFAULT_STATUS, DEFAULT_STATUS, DEFAULT_STATUS]
    assert sheet.after_hours_arrivals == set()


def test_apply_events_departure_after_hours_stays_on_its_own_date(base_workbook):
    """Реальні випадки: "вибув до лікарні 20.10 після 17.00", і підрозділ теж
    подає БПШП із 20.10. Правило лише про прибуття в район виконання завдань."""
    sheet = ot.Timesheet(base_workbook, "Табель")

    ot.apply_events(sheet, [{**_EVENING_ARRIVAL, "status": "ШП"}], _ARRIVAL_WEEK)

    assert _days_after_arrival(sheet) == ["ШП", "ШП", "ШП"]
    assert sheet.after_hours_arrivals == set()


def test_apply_events_evening_arrival_on_the_last_report_day_leaves_the_next_day_empty(base_workbook):
    """Дата після останнього рапорту не заповнюється: РВЗ на неї покаже наступний рапорт."""
    sheet = ot.Timesheet(base_workbook, "Табель")

    ot.apply_events(sheet, [_EVENING_ARRIVAL], [date(2026, 8, 1)])

    assert _days_after_arrival(sheet) == [AFTER_HOURS_ARRIVAL_STATUS, None, None]


def test_apply_payment_values_replaces_default_status_with_the_numeric_value(base_workbook):
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value="РВЗ")  # ПЕРШИЙ, 01.08
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [100]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == 100


def test_apply_payment_values_remembers_the_status_behind_a_number(tmp_path):
    """Число - лише вид оплати, тож статус, який воно замінило, лишається в
    Timesheet.statuses_behind_numbers: ППД (фіксоване перетворення) і БЗВП
    (подав підрозділ). Категорія поверх "РВЗ" і текстове "100_БПШП" - ні."""
    header = ["ПОСАДА", "ЗВАННЯ", "ПІБ"] + [datetime(2026, 8, day) for day in range(1, 5)]
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", header, [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ППД", "РВЗ", "РВЗ", "БПШП"],
    ])
    sheet = ot.Timesheet(path, "Табель")
    key = ot.normalize_name("ПЕРШИЙ Перший Перший")

    ot.apply_payment_values(sheet, {(key, date(2026, 8, 2)): ["БЗВП"], (key, date(2026, 8, 3)): [100]})

    assert [sheet.ws.cell(row=2, column=col).value for col in range(4, 8)] == [10, 10, 100, "100_БПШП"]
    assert sheet.statuses_behind_numbers == {(2, date(2026, 8, 1)): "ППД", (2, date(2026, 8, 2)): "БЗВП"}


def test_apply_payment_values_remembers_the_status_behind_a_category(tmp_path):
    """Текст-категорія теж лише вид оплати: 10_МЕД - це і ШП, і ЛМР, і ВЛК, а
    звірці з журналом лікування треба знати, що саме. Звіти показують
    категорію, тож це окремий словник."""
    header = ["ПОСАДА", "ЗВАННЯ", "ПІБ"] + [datetime(2026, 8, day) for day in range(1, 6)]
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", header, [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ШП", "ШП", "ВЛК", "РВЗ", "БПШП"],
    ])
    sheet = ot.Timesheet(path, "Табель")
    key = ot.normalize_name("ПЕРШИЙ Перший Перший")

    ot.apply_payment_values(sheet, {
        (key, date(2026, 8, 1)): ["ШП"], (key, date(2026, 8, 2)): ["БПШП"], (key, date(2026, 8, 3)): ["ВЛК"],
        (key, date(2026, 8, 4)): ["ЛМР"],
    })

    assert sheet.statuses_behind_categories == {
        (2, date(2026, 8, 1)): "ШП", (2, date(2026, 8, 2)): "БПШП", (2, date(2026, 8, 3)): "ВЛК",
        (2, date(2026, 8, 4)): "ЛМР", (2, date(2026, 8, 5)): "БПШП",
    }
    assert sheet.statuses_behind_numbers == {}


def test_apply_payment_values_keeps_no_status_behind_the_same_text(tmp_path):
    """"Повер" у клітинці - це вже сам статус."""
    header = ["ПОСАДА", "ЗВАННЯ", "ПІБ", datetime(2026, 8, 1)]
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", header, [["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ"]])
    sheet = ot.Timesheet(path, "Табель")

    ot.apply_payment_values(sheet, {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): ["Повер"]})

    assert sheet.ws.cell(row=2, column=4).value == "Повер"
    assert sheet.statuses_behind_categories == {}


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
    """PAYMENT_STATUS_OVERRIDES["ВП"]["ВПС"] - ідентичність: файл information_unit подає "ВПС"
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


@pytest.mark.parametrize("status", ["ПРВДБ", "ПРВДВЧ"])
def test_apply_payment_values_keeps_a_transfer_status_and_its_colour(base_workbook, status):
    """Рішення користувача: замість "ПРВД" - "ПРВДБ" (у підрозділи
    бригади) і "ПРВДВЧ" (в іншу в/ч): такий статус в ОБЛІК лишається як є
    (має власний колір), information_unit його не перекриває."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value=status)  # ПЕРШИЙ, 31.07
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): [status]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == status
    assert status in STATUS_COLORS


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
    """PAYMENT_STATUS_CONVERSIONS (constants.py) - за рішенням користувача: "ППД" завжди замінюється на ФІКСОВАНЕ число 10, незалежно
    від payment_values (не з файлів information_unit - той самий статус
    завжди дає той самий запис)."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ППД")

    unresolved = ot.apply_payment_values(sheet, payment_values={})

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=4).value == 10


def test_apply_payment_values_converts_bpshp_to_literal_compound_text(base_workbook):
    """"БПШП" - за рішенням користувача, замінюється на БУКВАЛЬНИЙ
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
    розпізнаваний статус (STATUS_COLORS). Для ПОРОЖНЬОЇ клітинки - за рішенням користувача - застосовується НАПРЯМУ, буквальним
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
    порожньою, і, БЕЗ запису на ручну
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


@pytest.mark.parametrize("status,expected", [("НОВ", 10), ("ЛМР", "10_МЕД"), ("ЗВФН", "10_МЕД")])
def test_apply_payment_values_overrides_default_status_with_its_payment_category(base_workbook, status, expected):
    """PAYMENT_STATUS_OVERRIDES[DEFAULT_STATUS]: "НОВ" -> число 10, "ЛМР" -> код "10_МЕД"
    (категорія виплати медиків - окремий код, а не число 10, як і "ШП"/"ВЛК" в тому ж
    словнику). Обидва - той самий шлях категорії виплати, а НЕ буквальний текст статусу."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value="РВЗ")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): [status]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == expected
    assert sheet.ws.cell(row=2, column=5).fill.fgColor.rgb == STATUS_COLORS[expected]


@pytest.mark.parametrize("cell_status,payment,expected", [
    ("Адап", [100], 100),       # реальний випадок: підрозділ подає 100 - адаптація вже скінчилась
    ("НОВ", [170], 170),
    ("БЗВП", [30], 30),
    ("Ррез", [100], 100),       # реальний випадок, написання "Ррез" - той самий статус
    ("Адап", [10], 10),         # 10 - той самий статус, не кінець
    ("РТГр", [100], "РТГр"),    # користувач сам прибрав заміну для РТГр
])
def test_apply_payment_values_ends_a_newcomer_or_reserve_status_by_a_regular_category(
        base_workbook, cell_status, payment, expected):
    """Рішення користувача ("командир подавав все правильно"):
    новоприбулі (Адап/НОВ/БЗВП) і резерв (РРез) закінчуються без жодної події в
    щоденному рапорті - звичайна категорія виплати від підрозділу (100/30/70/
    170) означає, що статус уже завершився, і записується вона."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value=cell_status)
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): payment}

    ot.apply_payment_values(sheet, payment_values)

    assert sheet.ws.cell(row=2, column=5).value == expected


@pytest.mark.parametrize("cell_status,payment,expected", [
    ("БПШП", None, "100_БПШП"),  # наявний статус конвертується: гілка конвертації клітинки
    (None, ["ППД"], 10),         # порожня клітинка + конвертоване значення
    ("РВЗ", ["ЛМР"], "10_МЕД"),  # перевизначення за замовчуванням
    (None, ["ВП"], "ВП"),        # порожня клітинка + текстовий статус без конвертації
    (None, [30], 30),            # порожня клітинка + число
])
def test_apply_payment_values_writes_the_value_even_without_a_fill_color(base_workbook, monkeypatch, cell_status, payment, expected):
    """Для значення, якому немає кольору (_fill_for_status -> None), клітинка все одно
    отримує значення - заливка просто не ставиться."""
    monkeypatch.setattr(ot, "_fill_for_status", lambda status: None)
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value=cell_status)
    payment_values = {} if payment is None else {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): payment}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == expected


def test_apply_payment_values_leaves_a_status_without_an_override_as_is(base_workbook):
    """Статус клітинки, для якого немає перевизначення (і він не РВЗ), лишається як є -
    навіть коли information_unit дає значення на цю дату."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value="Навч")  # немає власного перевизначення для непорожньої клітинки
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): ["ППД"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == "Навч"


def test_apply_payment_values_overrides_default_status_with_literal_text(base_workbook):
    """На відміну від "БЗВП"/"Адап"/"НОВ" (замінюються на ЧИСЛО 10) -
    "Повер" замінює "РВЗ" на буквальний ТЕКСТ (та сама логіка, що й "ВД" ->
    "Навч") - PAYMENT_STATUS_OVERRIDES[DEFAULT_STATUS] мапить КОЖЕН ключ на
    ВЛАСНЕ значення для запису, не всі на одне й те саме."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ 31.07 - нейтралізовано, поза увагою цього тесту
    sheet.ws.cell(row=2, column=5, value="РВЗ")
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1)): ["Повер"]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=2, column=5).value == "Повер"
    assert sheet.ws.cell(row=2, column=5).fill.fgColor.rgb == STATUS_COLORS["Повер"]


def _one_day_payment(tmp_path, status, submitted):
    """Табель одного дня зі статусом status після apply_payment_values; submitted -
    що подає підрозділ (None - нічого)."""
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", ["ПОСАДА", "ЗВАННЯ", "ПІБ", datetime(2026, 8, 1)], [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", status],
    ])
    sheet = ot.Timesheet(path, "Табель")
    key = (ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 8, 1))
    ot.apply_payment_values(sheet, {} if submitted is None else {key: submitted})
    return sheet


@pytest.mark.parametrize("status", ["ЗВФН", "ЛМР", " лмр "])
@pytest.mark.parametrize("submitted", [None, ["ШП"], ["ЛМР"], ["ЗВФН"], ["ВП"], [100], ["ШП", "ВЛК"]])
def test_medical_status_always_becomes_10_med(tmp_path, status, submitted):
    """Рішення користувача, реальний випадок: у Табелі ЛМР, підрозділ подає ШП, і
    клітинка лишалась текстом. Такої категорії money-проєкт не знає, тож ці дні
    випадали б із виплат."""
    sheet = _one_day_payment(tmp_path, status, submitted)

    cell = sheet.ws.cell(row=2, column=4)
    assert cell.value == "10_МЕД"
    assert cell.fill.fgColor.rgb == STATUS_COLORS["10_МЕД"]


@pytest.mark.parametrize("status", ["ЗВФН", "ЛМР"])
def test_status_behind_10_med_is_remembered(tmp_path, status):
    """Звіти показують статус, а не код виплати."""
    assert _one_day_payment(tmp_path, status, None).statuses_behind_categories == {(2, date(2026, 8, 1)): status}


@pytest.mark.parametrize("status", ["НОВ", "Адап", "БЗВП", " нов "])
@pytest.mark.parametrize("submitted", [None, ["НОВ"], ["Адап"], ["БЗВП"], [10], ["ВП"], [100, 30]])
def test_newcomer_status_becomes_10(tmp_path, status, submitted):
    """Рішення користувача, реальний випадок: у Табелі НОВ, підрозділ подає БЗВП чи
    НОВ, і клітинка лишалась текстом (151 день)."""
    sheet = _one_day_payment(tmp_path, status, submitted)

    cell = sheet.ws.cell(row=2, column=4)
    assert cell.value == 10
    assert cell.fill.fgColor.rgb == STATUS_COLORS[10]


@pytest.mark.parametrize("status", ["НОВ", "Адап", "БЗВП"])
def test_status_behind_a_newcomer_10_is_remembered(tmp_path, status):
    """10 - лише вид оплати, звіти показують статус (рішення користувача)."""
    sheet = _one_day_payment(tmp_path, status, ["БЗВП"])

    assert sheet.statuses_behind_numbers == {(2, date(2026, 8, 1)): status}


@pytest.mark.parametrize("status", ["НОВ", "Адап", "БЗВП"])
@pytest.mark.parametrize("category", [100, 30, 70, 170])
def test_regular_category_from_the_unit_still_ends_a_newcomer_status(tmp_path, status, category):
    """Рішення користувача: людина вже на службі, тож це не 10."""
    assert _one_day_payment(tmp_path, status, [category]).ws.cell(row=2, column=4).value == category


@pytest.mark.parametrize("status,submitted", [("РРез", ["Ррез"]), ("Ррез", None), ("РТГр", [100]), ("ВП", None)])
def test_other_statuses_stay_text(tmp_path, status, submitted):
    """Запасне 10 лише для новоприбулих: резерв, наприклад, не оплачується."""
    assert _one_day_payment(tmp_path, status, submitted).ws.cell(row=2, column=4).value == status


@pytest.mark.parametrize("information_unit_status", ["ПРВДБ", "ПРВДВЧ"])
def test_apply_payment_values_overrides_default_status_with_a_transfer_status(base_workbook, information_unit_status):
    """Новоприбула людина (порожня клітинка - той самий шлях, що й "РВЗ") -
    information_unit ОДНОЗНАЧНО каже "ПРВДБ"/"ПРВДВЧ" - записується як є."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.ws.cell(row=2, column=4, value="ВП")  # ПЕРШИЙ - не "РВЗ", щоб не заважати перевірці нижче
    sheet.ws.append(["Посада3", "матрос", "ТРЕТІЙ Третій Третій", None, None, None, None, ""])
    payment_values = {(ot.normalize_name("ТРЕТІЙ Третій Третій"), date(2026, 8, 1)): [information_unit_status]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert unresolved == []
    assert sheet.ws.cell(row=4, column=5).value == information_unit_status


@pytest.mark.parametrize("old_spelling", ["ПРВД", "Перевд.", "Првд"])
def test_apply_payment_values_does_not_guess_the_kind_of_transfer_from_old_spellings(base_workbook, old_spelling):
    """Колишній "ПРВД" і його написання не кажуть, куди переведено (бригада
    чи інша в/ч) - за рішенням користувача "ПРВД" видалено, тож для
    "РВЗ" такий текст від information_unit не підставляється, а лишається на
    перевірку."""
    sheet = ot.Timesheet(base_workbook, "Табель")  # ПЕРШИЙ, 31.07 - "РВЗ"
    payment_values = {(ot.normalize_name("ПЕРШИЙ Перший Перший"), date(2026, 7, 31)): [old_spelling]}

    unresolved = ot.apply_payment_values(sheet, payment_values)

    assert sheet.ws.cell(row=2, column=4).value == DEFAULT_STATUS
    assert any("ПЕРШИЙ" in str(item.get("reason", "")) for item in unresolved)


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
    wb = Workbook()
    wb._external_links = ["осиротіле посилання"]  # pyright: ignore[reportAttributeAccessIssue]

    ot._strip_external_links(wb)

    assert wb._external_links == []  # pyright: ignore[reportAttributeAccessIssue]


def _workbook_xml(path):
    with zipfile.ZipFile(path) as book:
        return book.read("xl/workbook.xml").decode("utf-8")


def test_save_removes_names_that_point_to_the_removed_external_links(base_workbook, tmp_path):
    """Реальний випадок: імена на [1] і [2] лишались після видалення зв'язків,
    і Excel показував "виявлено проблему із вмістом". Без таких імен той самий
    файл відкривався."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.wb.defined_names["Словник"] = DefinedName("Словник", attr_text="[1]СЛОВНИК!$G$2:$G$998")
    sheet.ws.defined_names["Оклад"] = DefinedName("Оклад", attr_text="_xlfn.XLOOKUP(1,[2]Тарифний!$A$1:$A$9,[2]Тарифний!$D$1:$D$9)")
    output_path = str(tmp_path / "ОБЛІК.xlsx")

    sheet.save(output_path)

    xml = _workbook_xml(output_path)
    assert "[1]" not in xml and "[2]" not in xml
    assert "Словник" not in xml and "Оклад" not in xml


def test_save_keeps_names_without_external_links(base_workbook, tmp_path):
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.wb.defined_names["Загальне"] = DefinedName("Загальне", attr_text="Табель!$A$1:$A$9")
    sheet.ws.defined_names["Місцеве"] = DefinedName("Місцеве", attr_text="Табель!$B$1:$B$9")
    output_path = str(tmp_path / "ОБЛІК.xlsx")

    sheet.save(output_path)

    saved = load_workbook(output_path)
    assert "Загальне" in saved.defined_names
    assert "Місцеве" in saved["Табель"].defined_names


def test_export_for_money_project_removes_names_of_external_links(base_workbook, tmp_path):
    wb = load_workbook(base_workbook)
    wb.defined_names["Словник"] = DefinedName("Словник", attr_text="[1]СЛОВНИК!$G$2:$G$998")
    source = str(tmp_path / "джерело.xlsx")
    wb.save(source)
    output_path = str(tmp_path / "ОБЛІК.xlsx")

    ot.export_for_money_project(source, output_path, "ТРАВЕНЬ")

    assert "[1]" not in _workbook_xml(output_path)


def test_save_removes_external_links_before_saving(base_workbook, tmp_path):
    """Timesheet.save() прибирає зовнішні посилання ПЕРЕД self.wb.save() -
    інакше (осиротіле посилання пережило б збереження, як діагностовано на
    реальному resources/ОБЛІК.xlsx) openpyxl впав би з AttributeError,
    намагаючись серіалізувати непридатний запис."""
    sheet = ot.Timesheet(base_workbook, "Табель")
    sheet.wb._external_links = ["осиротіле посилання (заглушка тесту)"]  # pyright: ignore[reportAttributeAccessIssue]
    output_path = tmp_path / "out" / "ОБЛІК.xlsx"

    sheet.save(str(output_path))

    assert load_workbook(str(output_path))._external_links == []  # pyright: ignore[reportAttributeAccessIssue]


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
    """Рішення користувача: готує вже згенерований "ОБЛІК для
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
        wb._external_links = ["осиротіле посилання (заглушка тесту)"]  # pyright: ignore[reportAttributeAccessIssue]
        return wb

    monkeypatch.setattr(ot, "load_workbook", _load_with_fake_external_link)
    output_path = tmp_path / "out" / "ОБЛІК.xlsx"

    ot.export_for_money_project(base_workbook, str(output_path), "ТРАВЕНЬ")

    assert load_workbook(str(output_path))._external_links == []  # pyright: ignore[reportAttributeAccessIssue]


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


def _wide_timesheet(tmp_path, pib, values, days=40, first_day=datetime(2026, 8, 1)):
    days = max(days, len(values))
    header = ["ПОСАДА", "ЗВАННЯ", "ПІБ"] + [first_day + timedelta(days=i) for i in range(days)]
    row = ["Посада1", "сержант", pib] + values + [None] * (days - len(values))
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", header, [row])
    return ot.Timesheet(path, "Табель")


def _person(name, days=None, color=None):
    """Запис EXCLUDE_NAMES так, як його віддає constants.py."""
    return {"name": name, "days": days, "color": color}


@pytest.fixture
def limit_30(monkeypatch):
    """Спільний ліміт задає користувач у data.json, тож тести правила не
    залежать від того, яке число там зараз."""
    monkeypatch.setattr(ot, "CONST_PAYMENT_CATEGORY_100_MAX_DAYS", 30)


def test_payment_category_100_limit_is_whatever_the_user_sets_in_data_json(tmp_path, monkeypatch):
    """Рішення користувача: один спільний ліміт без поділу на місяці."""
    monkeypatch.setattr(ot, "CONST_PAYMENT_CATEGORY_100_MAX_DAYS", 20)
    over = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 21)
    at_limit = _wide_timesheet(tmp_path, "ДРУГИЙ Другий Другий", [100] * 20)

    unresolved = ot.check_payment_category_100_days(over, people=())

    assert len(unresolved) == 1
    assert "21 день поспіль (01.08.2026-21.08.2026), дозволено не більше 20." in unresolved[0]["reason"]
    assert ot.check_payment_category_100_days(at_limit, people=()) == []


def test_payment_category_100_person_days_can_allow_more_than_the_common_limit(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 31)

    assert ot.check_payment_category_100_days(sheet, people=[_person("ПЕРШИЙ Перший Перший", days=31)]) == []


def test_payment_category_100_person_days_can_allow_less_than_the_common_limit(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 15)

    (violation,) = ot.check_payment_category_100_days(sheet, people=[_person("ПЕРШИЙ Перший Перший", days=14)])

    assert violation["limit"] == 14
    assert "15 днів поспіль (01.08.2026-15.08.2026), дозволено не більше 14." in violation["reason"]


def test_payment_category_100_empty_person_days_mean_the_common_limit(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 31)

    (violation,) = ot.check_payment_category_100_days(sheet, people=[_person("ПЕРШИЙ Перший Перший", days=None)])

    assert violation["limit"] == 30


def test_payment_category_100_zero_person_days_is_a_limit_not_an_empty_field(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100])

    (violation,) = ot.check_payment_category_100_days(sheet, people=[_person("ПЕРШИЙ Перший Перший", days=0)])

    assert violation["limit"] == 0


def test_payment_category_100_person_days_apply_only_to_that_person(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 31)

    (violation,) = ot.check_payment_category_100_days(sheet, people=[_person("ДРУГИЙ Другий Другий", days=31)])

    assert violation["limit"] == 30


def test_category_100_violation_carries_the_color_of_the_person(tmp_path, limit_30):
    """Колір потрібен таблиці Перевірка_категорії_100.xlsx."""
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 31)

    (violation,) = ot.check_payment_category_100_days(sheet, people=[_person("ПЕРШИЙ Перший Перший", color="FFFF0000")])

    assert violation["color"] == "FFFF0000"


def test_category_100_violation_of_a_person_without_color_has_none(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 31)

    (violation,) = ot.check_payment_category_100_days(sheet, people=())

    assert violation["color"] is None


def test_marked_person_within_the_limit_gets_a_row_with_the_longest_series(tmp_path, limit_30):
    """Рішення користувача: людина з кольором є в таблиці й без порушення."""
    values = [100] * 5 + ["РВЗ"] + [100] * 20 + ["РВЗ"]
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", values)
    people = [_person("ПЕРШИЙ Перший Перший", days=31, color="FFFFFF00")]

    (row,) = ot.category_100_marked_rows(sheet, ot.check_payment_category_100_days(sheet, people), people)

    assert {key: row[key] for key in ("pib_raw", "pidrozdil", "start", "end", "days", "limit", "color")} == {
        "pib_raw": "ПЕРШИЙ Перший Перший", "pidrozdil": None, "start": date(2026, 8, 7), "end": date(2026, 8, 26),
        "days": 20, "limit": 31, "color": "FFFFFF00",
    }


def test_marked_person_takes_the_first_of_two_equal_series(tmp_path, limit_30):
    values = [100] * 3 + ["РВЗ"] + [100] * 3
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", values)

    (row,) = ot.category_100_marked_rows(sheet, [], [_person("ПЕРШИЙ Перший Перший", color="FFFFFF00")])

    assert row["start"] == date(2026, 8, 1)


def test_marked_person_with_empty_days_shows_the_common_limit(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 3)

    (row,) = ot.category_100_marked_rows(sheet, [], [_person("ПЕРШИЙ Перший Перший", days=None, color="FFFFFF00")])

    assert row["limit"] == 30


def test_marked_person_without_100_days_still_gets_a_row(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", ["РВЗ"] * 5)

    (row,) = ot.category_100_marked_rows(sheet, [], [_person("ПЕРШИЙ Перший Перший", color="FFFFFF00")])

    assert (row["start"], row["end"], row["days"]) == (None, None, 0)


def test_marked_person_with_a_violation_is_not_listed_twice(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 31)
    people = [_person("ПЕРШИЙ Перший Перший", color="FFFFFF00")]

    violations = ot.check_payment_category_100_days(sheet, people)

    assert len(violations) == 1
    assert ot.category_100_marked_rows(sheet, violations, people) == []


def test_person_without_color_is_not_listed_without_a_violation(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 3)

    assert ot.category_100_marked_rows(sheet, [], [_person("ПЕРШИЙ Перший Перший", days=31, color=None)]) == []


def test_marked_rows_find_the_person_regardless_of_case_and_spacing(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 3)

    rows = ot.category_100_marked_rows(sheet, [], [_person("  перший   ПЕРШИЙ перший ", color="FFFFFF00")])

    assert [row["pib_raw"] for row in rows] == ["ПЕРШИЙ Перший Перший"]


def test_marked_person_who_is_not_in_the_timesheet_has_no_row(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 3)

    assert ot.category_100_marked_rows(sheet, [], [_person("ДРУГИЙ Другий Другий", color="FFFFFF00")]) == []


def _pib_fill(sheet, row=2):
    return sheet.ws.cell(row=row, column=sheet.label_columns["ПІБ"]).fill


def test_highlight_fills_the_name_cell_with_the_color_of_the_person(tmp_path):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100])

    ot.highlight_category_100_people(sheet, [_person("ПЕРШИЙ Перший Перший", color="FFFFFF00")])

    assert _pib_fill(sheet).fill_type == "solid"
    assert _pib_fill(sheet).start_color.rgb == "FFFFFF00"


def test_highlight_finds_the_person_regardless_of_case_and_spacing(tmp_path):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100])

    ot.highlight_category_100_people(sheet, [_person("  перший   ПЕРШИЙ перший ", color="FFFFFF00")])

    assert _pib_fill(sheet).fill_type == "solid"


def test_highlight_fills_a_person_with_a_violation_yellow(tmp_path, limit_30):
    """Рішення користувача: порушники з таблиці Перевірка_категорії_100.xlsx
    жовті й у Табелі, навіть якщо їх немає в EXCLUDE_NAMES."""
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 31)

    ot.highlight_category_100_people(sheet, people=())

    assert _pib_fill(sheet).start_color.rgb == "FFFFFF00"


def test_highlight_fills_a_person_with_a_violation_even_if_their_color_is_off(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 31)

    ot.highlight_category_100_people(sheet, [_person("ПЕРШИЙ Перший Перший", color=None)])

    assert _pib_fill(sheet).start_color.rgb == "FFFFFF00"


def test_highlight_keeps_the_own_color_of_a_person_with_a_violation(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 31)

    ot.highlight_category_100_people(sheet, [_person("ПЕРШИЙ Перший Перший", color="FFFF0000")])

    assert _pib_fill(sheet).start_color.rgb == "FFFF0000"


def test_highlight_leaves_a_person_within_the_limit_as_is(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 30)

    ot.highlight_category_100_people(sheet, people=())

    assert _pib_fill(sheet).fill_type is None


def test_highlight_respects_the_days_of_the_person(tmp_path, limit_30):
    """Людина, якій дозволено 31 день, не порушник, тож і не жовта."""
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 31)

    ot.highlight_category_100_people(sheet, [_person("ПЕРШИЙ Перший Перший", days=31, color=None)])

    assert _pib_fill(sheet).fill_type is None


def test_highlight_fills_only_the_row_of_the_person_with_a_violation(tmp_path, limit_30):
    header = ["ПОСАДА", "ЗВАННЯ", "ПІБ"] + [datetime(2026, 8, day) for day in range(1, 32)]
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", header, [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший"] + [100] * 30 + ["РВЗ"],
        ["Посада2", "сержант", "ДРУГИЙ Другий Другий"] + [100] * 31,
    ])
    sheet = ot.Timesheet(path, "Табель")

    ot.highlight_category_100_people(sheet, people=())

    assert _pib_fill(sheet, row=2).fill_type is None
    assert _pib_fill(sheet, row=3).start_color.rgb == "FFFFFF00"


def test_highlight_leaves_a_person_without_color_as_is(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100])

    ot.highlight_category_100_people(sheet, [_person("ПЕРШИЙ Перший Перший", color=None)])

    assert _pib_fill(sheet).fill_type is None


def test_highlight_touches_only_the_name_cell(tmp_path):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100])

    ot.highlight_category_100_people(sheet, [_person("ПЕРШИЙ Перший Перший", color="FFFFFF00")])

    pib_col = sheet.label_columns["ПІБ"]
    others = [cell for cell in sheet.ws[2] if cell.column != pib_col]
    assert others and all(cell.fill.fill_type is None for cell in others)


def test_highlight_skips_a_person_who_is_not_in_the_timesheet(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100])

    ot.highlight_category_100_people(sheet, [_person("ДРУГИЙ Другий Другий", color="FFFFFF00")])

    assert _pib_fill(sheet).fill_type is None


def test_category_100_violation_keeps_what_the_table_needs(tmp_path, monkeypatch):
    monkeypatch.setattr(ot, "CONST_PAYMENT_CATEGORY_100_MAX_DAYS", 20)
    header =["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ"] + [datetime(2026, 8, day) for day in range(1, 23)]
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", header, [["1 абв", "Посада1", "сержант", "ПЕРШИЙ Перший Перший"] + [100] * 22])

    (violation,) = ot.check_payment_category_100_days(ot.Timesheet(path, "Табель"), people=())

    assert {key: violation[key] for key in ("pib_raw", "pidrozdil", "start", "end", "days", "limit")} == {
        "pib_raw": "ПЕРШИЙ Перший Перший", "pidrozdil": "1 абв", "start": date(2026, 8, 1), "end": date(2026, 8, 22),
        "days": 22, "limit": 20,
    }


def test_category_100_violation_without_a_subdivision_column(tmp_path, monkeypatch):
    monkeypatch.setattr(ot, "CONST_PAYMENT_CATEGORY_100_MAX_DAYS", 20)

    (violation,) = ot.check_payment_category_100_days(_wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 21), people=())

    assert violation["pidrozdil"] is None


@pytest.mark.parametrize("value", [None, "", "28", True, 28.5, 32, -1])
def test_common_limit_in_data_json_must_be_a_number_of_days(value, monkeypatch):
    import constants

    monkeypatch.setitem(constants._DATA, "CONST_PAYMENT_CATEGORY_100_MAX_DAYS", value)

    with pytest.raises(ValueError, match="CONST_PAYMENT_CATEGORY_100_MAX_DAYS"):
        constants._category_100_max_days()


def test_common_limit_must_be_present_in_data_json(monkeypatch):
    import constants

    monkeypatch.delitem(constants._DATA, "CONST_PAYMENT_CATEGORY_100_MAX_DAYS")

    with pytest.raises(ValueError, match="CONST_PAYMENT_CATEGORY_100_MAX_DAYS"):
        constants._category_100_max_days()


@pytest.mark.parametrize("value", [0, 28, 31])
def test_common_limit_takes_the_number_from_data_json(value, monkeypatch):
    import constants

    monkeypatch.setitem(constants._DATA, "CONST_PAYMENT_CATEGORY_100_MAX_DAYS", value)

    assert constants._category_100_max_days() == value


_NAME = "ПЕРШИЙ Перший Перший"


def _people_from_data_json(monkeypatch, entries):
    import constants

    monkeypatch.setitem(constants._DATA, "EXCLUDE_NAMES", entries)
    return constants._category_100_people("EXCLUDE_NAMES")


@pytest.mark.parametrize("entry", [
    _NAME,  # старий запис рядком
    {"days": 31, "color": False},
    {"name": "", "days": 31, "color": False},
    {"name": "   ", "days": 31, "color": False},
    {"name": 5, "days": 31, "color": False},
    {"name": _NAME, "color": False},
    {"name": _NAME, "days": 31},
    {"name": _NAME, "days": "31", "color": False},
    {"name": _NAME, "days": True, "color": False},
    {"name": _NAME, "days": 31.5, "color": False},
    {"name": _NAME, "days": 32, "color": False},
    {"name": _NAME, "days": -1, "color": False},
    {"name": _NAME, "days": 31, "color": "жовтий"},
    {"name": _NAME, "days": 31, "color": "FF00"},
    {"name": _NAME, "days": 31, "color": "#FFFF0"},
    {"name": _NAME, "days": 31, "color": "FFFF00 жовтий"},
    {"name": _NAME, "days": 31, "color": 1},
    {"name": _NAME, "days": 31, "color": 0},
])
def test_exclude_names_entry_that_cannot_be_read_stops_the_run(entry, monkeypatch):
    """Як і в _date_name_pairs: помилка одразу, а не тихо проігнорований запис."""
    with pytest.raises(ValueError, match="EXCLUDE_NAMES"):
        _people_from_data_json(monkeypatch, [entry])


@pytest.mark.parametrize("entries", [_NAME, "", {"name": _NAME, "days": 31, "color": False}, {}])
def test_exclude_names_must_be_a_list(entries, monkeypatch):
    with pytest.raises(ValueError, match="EXCLUDE_NAMES.*має бути списком"):
        _people_from_data_json(monkeypatch, entries)


@pytest.mark.parametrize("value, expected", [(31, 31), (5, 5), (0, 0), (None, None), ("", None)])
def test_exclude_names_days(value, expected, monkeypatch):
    (person,) = _people_from_data_json(monkeypatch, [{"name": _NAME, "days": value, "color": False}])

    assert person["days"] == expected


@pytest.mark.parametrize("value, expected", [
    (False, None), (None, None), ("", None),
    (True, "FFFFFF00"),  # жовтий (рішення користувача)
    ("FF0000", "FFFF0000"), ("#ff0000", "FFFF0000"), ("80ff0000", "80FF0000"),
])
def test_exclude_names_color(value, expected, monkeypatch):
    (person,) = _people_from_data_json(monkeypatch, [{"name": _NAME, "days": 31, "color": value}])

    assert person["color"] == expected


def test_exclude_names_keeps_the_name_and_the_order(monkeypatch):
    people = _people_from_data_json(monkeypatch, [
        {"name": _NAME, "days": 31, "color": False},
        {"name": "ДРУГИЙ Другий Другий", "days": None, "color": True},
    ])

    assert [person["name"] for person in people] == [_NAME, "ДРУГИЙ Другий Другий"]


def test_exclude_names_may_be_empty(monkeypatch):
    assert _people_from_data_json(monkeypatch, []) == ()


def test_payment_category_100_flags_a_whole_month_without_a_break(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 31)

    unresolved = ot.check_payment_category_100_days(sheet, people=())

    assert [item["reason"] for item in unresolved] == [
        "\"ПЕРШИЙ Перший Перший\": категорія виплати 100 - 31 день поспіль (01.08.2026-31.08.2026), "
        "дозволено не більше 30. Якщо так має бути - вкажіть людині кількість днів в EXCLUDE_NAMES (resources/data.json).",
    ]
    assert unresolved[0]["pib_raw"] == "ПЕРШИЙ Перший Перший"


def test_payment_category_100_allows_one_other_day_in_the_month(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 30 + ["РВЗ"])

    assert ot.check_payment_category_100_days(sheet, people=()) == []


def test_payment_category_100_counts_from_the_first_day_of_the_month(tmp_path, limit_30):
    """Реальний випадок: жовтневий Табель починається з 30.09 (останній день
    попереднього місяця), і 100 з 30.09 по 29.10 рахувалось як "30 днів (з
    30.09)". Серія рахується з 1-го числа: 31.07 + 01.08-30.08 - не
    порушення, а 31.07 + увесь серпень - порушення з 01.08."""
    first_day = datetime(2026, 7, 31)
    with_a_break = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 31 + ["РВЗ"], first_day=first_day)
    whole_month = _wide_timesheet(tmp_path, "ДРУГИЙ Другий Другий", [100] * 32, first_day=first_day)

    assert ot.check_payment_category_100_days(with_a_break, people=()) == []
    unresolved = ot.check_payment_category_100_days(whole_month, people=())
    assert len(unresolved) == 1
    assert "31 день поспіль (01.08.2026-31.08.2026)" in unresolved[0]["reason"]


def test_payment_category_100_counts_every_month_on_its_own(tmp_path, monkeypatch):
    """Увесь серпень (31) і увесь вересень (30) - два окремі порушення, а не
    одна серія; лютий (28) уміщається в той самий ліміт."""
    monkeypatch.setattr(ot, "CONST_PAYMENT_CATEGORY_100_MAX_DAYS", 29)
    summer = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * 61)
    february = _wide_timesheet(tmp_path, "ДРУГИЙ Другий Другий", [100] * 28, first_day=datetime(2026, 2, 1))

    reasons = [item["reason"] for item in ot.check_payment_category_100_days(summer, people=())]
    assert len(reasons) == 2
    assert "31 день поспіль (01.08.2026-31.08.2026), дозволено не більше 29." in reasons[0]
    assert "30 днів поспіль (01.09.2026-30.09.2026), дозволено не більше 29." in reasons[1]
    assert ot.check_payment_category_100_days(february, people=()) == []


def test_payment_category_100_breaks_the_streak_on_any_other_value(tmp_path, limit_30):
    values = [100] * 15 + ["РВЗ"] + [100] * 15
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", values)

    assert ot.check_payment_category_100_days(sheet, people=()) == []


def test_payment_category_100_breaks_the_streak_on_an_empty_day(tmp_path, limit_30):
    values = [100] * 15 + [None] + [100] * 15
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", values)

    assert ot.check_payment_category_100_days(sheet, people=()) == []


def test_payment_category_100_matches_people_regardless_of_case_and_spacing(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ТРЕТІЙ Третій Третій", [100] * 31)

    assert ot.check_payment_category_100_days(sheet, people=[_person("  третій   ТРЕТІЙ третій ", days=31)]) == []


def test_payment_category_100_ignores_compound_text_categories(tmp_path, limit_30):
    sheet = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", ["100_БПШП"] * 31)

    assert ot.check_payment_category_100_days(sheet, people=()) == []


def test_payment_category_100_uses_the_limit_from_data_json(tmp_path):
    from constants import CONST_PAYMENT_CATEGORY_100_MAX_DAYS as limit

    at_limit =_wide_timesheet(tmp_path, "ДРУГИЙ Другий Другий", [100] * min(limit, 31))
    assert ot.check_payment_category_100_days(at_limit) == []
    if limit < 31:
        over_by_one = _wide_timesheet(tmp_path, "ПЕРШИЙ Перший Перший", [100] * (limit + 1))
        assert len(ot.check_payment_category_100_days(over_by_one)) == 1


def test_days_word_agrees_with_the_number():
    assert [f"{n} {ot._days_word(n)}" for n in (1, 2, 5, 11, 12, 21, 22, 25, 30, 31)] == [
        "1 день", "2 дні", "5 днів", "11 днів", "12 днів", "21 день", "22 дні", "25 днів", "30 днів", "31 день",
    ]


def test_payment_category_100_config_exposes_whole_people_not_characters():
    from constants import PAYMENT_CATEGORY_100_PEOPLE

    assert isinstance(PAYMENT_CATEGORY_100_PEOPLE, tuple)
    assert all(set(person) == {"name", "days", "color"} and len(person["name"]) > 1 for person in PAYMENT_CATEGORY_100_PEOPLE)


def _relative_luminance(argb):
    rgb = [int(argb[i:i + 2], 16) / 255 for i in (2, 4, 6)]
    linear = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in rgb]
    return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]


def test_status_colors_keep_black_text_readable():
    """Колір заливки має давати контраст з ЧОРНИМ текстом не менше 4.5:1 (WCAG AA) -
    інакше значення клітинки (напр. "Загиб") стає практично невидимим."""
    for key, argb in STATUS_COLORS.items():
        contrast = (_relative_luminance(argb) + 0.05) / 0.05
        assert contrast >= 4.5, f"{key!r} {argb} контраст {contrast:.1f}"


def test_core_status_colors_are_visibly_distinct():
    """Статуси з РІЗНИМ значенням (не аліаси) мають відрізнятись на око - мінімальна
    відстань у RGB між будь-якою парою не менше 20."""
    core = [
        DEFAULT_STATUS, "ВП", "ВПСЗ", "ШП", "СЗЧ", "ВД", "РТГр", "ВЛК", "ПРВДБ", "ПРВДВЧ", "ЗВФН",
        "ЛМР", "ВПБП", "БПШП", "Загиб", "Смерть", 10, 30, 70, 100, 170,
    ]
    rgb = {key: STATUS_COLORS[key][2:] for key in core}
    for i, a in enumerate(core):
        for b in core[i + 1:]:
            distance = sum((int(rgb[a][j:j + 2], 16) - int(rgb[b][j:j + 2], 16)) ** 2 for j in (0, 2, 4)) ** 0.5
            assert distance >= 20, f"{a!r} {rgb[a]} та {b!r} {rgb[b]} надто схожі (відстань {distance:.0f})"


def test_normalize_name_ignores_the_apostrophe_but_spelling_differences_are_detectable():
    """Рішення користувача (реальний випадок: в ОБЛІК - з
    апострофом, в інших документах - без, і людину не було зіставлено зовсім,
    тож і розбіжність ніде не показувалась): для зіставлення апостроф (будь-
    яке написання, з пробілом після нього, є він чи немає) - не різниця;
    spelled_differently - є апостроф / немає - різниця, а вид апострофа,
    пробіли й регістр - ні."""
    variants = ["П'ЯТИЙ П'ятий П'ятий", "П’ЯТИЙ П’ятий П’ятий", "ПʼЯТИЙ П' ятий Пʼятий", "ПЯТИЙ Пятий Пятий"]
    assert {ot.normalize_name(name) for name in variants} == {"ПЯТИЙ ПЯТИЙ ПЯТИЙ"}
    assert not ot.spelled_differently("П'ЯТИЙ П'ятий П'ятий", "П’ЯТИЙ  П’ятий п’ятий")
    assert not ot.spelled_differently("П'ЯТИЙ П'ятий П'ятий", "П' ЯТИЙ П'ятий П'ятий")
    assert ot.spelled_differently("П'ЯТИЙ П'ятий П'ятий", "ПЯТИЙ Пятий Пятий")
    assert ot.spelled_differently("Не знайдено в ОБЛІК.xlsx", "ПЯТИЙ Пятий Пятий")
    assert not ot.spelled_differently(None, None)


def test_a_letter_written_with_two_symbols_is_the_same_name():
    """Реальний випадок: у книзі "й" записана двома символами ("и" і знак над
    нею). На вигляд ПІБ те саме, а порівняння вважало його іншим, і звіт
    показував два однакові на вигляд ПІБ як невідповідність."""
    name = "ПЕРШИЙ Перший Їжак"
    two_symbols = unicodedata.normalize("NFD", name)
    assert two_symbols != name

    assert ot.normalize_name(two_symbols) == ot.normalize_name(name)
    assert not ot.spelled_differently(two_symbols, name)


def test_apply_events_matches_a_name_spelled_without_the_apostrophe_and_notes_it(tmp_path):
    """Рапорт пише ПІБ без апострофа, ростер - з ним: статус застосовано тій
    самій людині, а в перевірці - що написання різне."""
    path = _write_workbook(tmp_path / "ОБЛІК.xlsx", _HEADER, [
        ["Посада1", "сержант", "П'ЯТИЙ П'ятий П'ятий", "РВЗ", None, None, None, ""],
    ])
    sheet = ot.Timesheet(path, "Табель")
    events = [{"pib_raw": "ПЯТИЙ Пятий Пятий", "status": "ВП", "effective_date": date(2026, 8, 1)}]

    unresolved = ot.apply_events(sheet, events, report_dates=[date(2026, 8, 1)])

    assert sheet.ws.cell(row=2, column=5).value == "ВП"
    assert [item["reason"] for item in unresolved] == [
        "ПІБ \"ПЯТИЙ Пятий Пятий\" з рапорту написано інакше, ніж у роcтері \"П'ЯТИЙ П'ятий П'ятий\" "
        "(апостроф) - статус застосовано; перевірте, яке написання правильне."
    ]
