"""Звірка ЩОЙНО згенерованого ОБЛІК.xlsx (аркуш "Табель") з книгами
resources/addons - book61 (content/addon_reader.py, інструкція -
resources/ОБЛІК.xlsx, аркуш "map") і ЦВС (content/cvs_reader.py, аркуш "map"
самої книги) - чи відповідає статус, що реально стоїть у Табелі на кожну
дату, тому, що case-за-case виводиться з цих книг (періоди відпусток/
відряджень/лікування/..., адміністративні накази, втрати).

Будує mis_match.xlsx - той самий стиль (заливки/шрифти/межі), що й
error_mis_statuses.xlsx (write_mismatch_report, content/
payment_mismatch_checker.py) - стилі ЗВІДТИ й перевикористовуються
напряму, а не копіюються."""

import os
import re
from bisect import bisect_left, bisect_right

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta

from constants import (
    ADDON_ARRIVAL_DAY_SHEETS,
    ADDON_ARRIVAL_DAY_STATUSES,
    ADDON_EQUIVALENT_STATUS_GROUPS,
    ADDON_TABEL_PRIORITY_STATUSES,
    DATE_NUMBER_FORMAT,
    PAYMENT_STATUS_CONVERSIONS,
    PAYMENT_STATUS_OVERRIDES,
    PIB_COLUMN_NAME,
    TRANSFER_STATUSES,
)
from content.addon_reader import ANY_STATUS
from content.oblik_timesheet import normalize_name
from content.payment_mismatch_checker import (
    _BORDER,
    _CENTER,
    _display_value,
    _FONT,
    _HEADER_ALIGNMENT,
    _HEADER_FILL,
    _HEADER_FONT,
    _is_number,
    _MATCH_FILL,
    _MISMATCH_FILL,
    _normalize_for_compare,
)


def _is_payment_category_value(value):
    """Чи "value" - запис КАТЕГОРІЇ ВИПЛАТИ (число, чи текст вигляду
    "100_..."/"10_..." - PAYMENT_STATUS_CONVERSIONS/PAYMENT_STATUS_OVERRIDES,
    constants.py), а НЕ звичайний текстовий статус. Обидва словники містять
    і такі "категорійні" значення, і прості ІДЕНТИЧНОСТІ (напр. "Навч" ->
    "Навч", "ВПС" -> "ВПС", "Повер" -> "Повер" - information_unit лише
    ПІДТВЕРДЖУЄ конкретнішу причину, без зміни категорії виплати) -
    розрізнення потрібне, щоб _payment_converted_forms нижче не плутала ці
    два випадки: лише перше - "та сама ситуація, інший запис", друге - й
    далі змістовно ІНША, розрізнювана причина."""
    if _is_number(value):
        return True
    if isinstance(value, str) and "_" in value:
        return value.split("_", 1)[0].isdigit()
    return False


def _payment_converted_forms(status):
    """Усі значення, які apply_payment_values (content/oblik_timesheet.py)
    МІГ БИ реально записати в клітинку Табеля для дати, чий ПОДІЄВИЙ статус
    (щоденний рапорт) - "status" (PAYMENT_STATUS_CONVERSIONS[status] -
    фіксована конвертація; PAYMENT_STATUS_OVERRIDES[status].values() - за
    підтвердженням information_unit) - АЛЕ лише "категорійні" значення
    (_is_payment_category_value). book61 (джерело "очікуваного" статусу
    тут) звіряється зі своїм ВЛАСНИМ, незалежним визначенням факту, тож
    "100_БПШП" замість "БПШП" чи число 10 замість "ВД" - ТА САМА реальна
    ситуація, лише записана по-іншому вже ПІСЛЯ book61, а не розбіжність.

    Реальні випадки, підтверджені на output/ОБЛІК.xlsx: book61 (аркуш "ШП")
    очікує "БПШП", Табель каже "100_БПШП"; book61 (аркуш "ВД") очікує "ВД",
    Табель каже число 10; аналогічно "ВПСЗ" (аркуш "ВП") і Табель
    "100_ВПБП".

    НАВМИСНО НЕ зводить "ВД" із текстовим "Навч" (PAYMENT_STATUS_OVERRIDES
    ["ВД"]["Навч"] = "Навч" - ІДЕНТИЧНІСТЬ, не категорія виплати, відхиляється
    _is_payment_category_value) - ВД/Навч - справжня, змістовна різниця
    (content/daily_report_reader.py::_is_training_destination), а не той
    самий факт по-іншому записаний. Так само книга НЕ матиме записів для
    "Навч" як базового статусу (DEFAULT_STATUS/"ВД" - єдині там), тож
    forms для очікуваного "Навч" - порожній набір: фактичне число 10 проти
    очікуваного book61 "Навч" і далі коректно лишається розбіжністю, а не
    мовчки гаситься через "10 ~ ВД"."""
    forms = set()
    converted = PAYMENT_STATUS_CONVERSIONS.get(status)
    if converted is not None and _is_payment_category_value(converted):
        forms.add(converted)
    for replacement in PAYMENT_STATUS_OVERRIDES.get(status, {}).values():
        if _is_payment_category_value(replacement):
            forms.add(replacement)
    # Статус, яким information_unit уточнює ІНШИЙ базовий статус, - теж:
    # "БЗВП"/"Адап"/"НОВ" у клітинці "РВЗ" записуються числом 10
    # (PAYMENT_STATUS_OVERRIDES[DEFAULT_STATUS]) - реальний випадок: журнал
    # обліку БП очікує "БЗВП", Табель для виплат каже 10.
    for override_map in PAYMENT_STATUS_OVERRIDES.values():
        replacement = override_map.get(status)
        if replacement is not None and _is_payment_category_value(replacement):
            forms.add(replacement)
    return forms


def _paid_as(status):
    """Категорії виплати, якими статус оплачується за будь-яких обставин:
    фіксована конвертація (ППД -> 10) чи заміна, коли саме цей статус подає
    information_unit (БЗВП/Адап/НОВ/ВД -> 10) - без "можливих уточнень"
    базового статусу (РВЗ лише МОЖЕ стати 10)."""
    forms = {PAYMENT_STATUS_CONVERSIONS.get(status)}
    forms.update(override_map.get(status) for override_map in PAYMENT_STATUS_OVERRIDES.values())
    return {_normalize_for_compare(form) for form in forms if form is not None and _is_payment_category_value(form)}


def _statuses_match(expected, actual):
    if expected == ANY_STATUS:
        return True
    expected_normalized = _normalize_for_compare(expected)
    actual_normalized = _normalize_for_compare(actual)
    if expected_normalized == actual_normalized:
        return True
    # Очікується сама категорія виплати (журнал обліку БП: "закордон" - 10) -
    # збіг і для статусу в Табелі, що нею ЗАВЖДИ оплачується (ППД/БЗВП/НОВ/
    # Адап/ВД -> 10), але не для "РВЗ", який лише МОЖЕ стати 10.
    if _is_number(expected) and expected_normalized in _paid_as(actual):
        return True
    for group in ADDON_EQUIVALENT_STATUS_GROUPS:
        normalized_group = {_normalize_for_compare(status) for status in group}
        if expected_normalized in normalized_group and actual_normalized in normalized_group:
            return True
    for expected_status, tabel_statuses in ADDON_TABEL_PRIORITY_STATUSES.items():
        if expected_normalized == _normalize_for_compare(expected_status) and any(
            actual_normalized == _normalize_for_compare(status) for status in tabel_statuses
        ):
            return True
    return any(actual_normalized == _normalize_for_compare(form) for form in _payment_converted_forms(expected))


# За прямою вказівкою користувача - процес прибуття з аркуша "Прибулі у ВД"
# (РРез -> ППД -> будь-який статус у день прийому посади, content/
# addon_reader.py::read_rrez_tvo_periods) ПЕРЕВАЖАЄ будь-який паралельний
# book61-період на ВЕСЬ свій проміжок (реальний випадок: ТВО і "ППД" з тієї
# самої дати; "ППД" з аркуша ППД продовжується лише після прийому посади) -
# тож накладається ОСТАННІМ, поверх решти, незалежно від дати початку.
_PRIORITY_SOURCE_SHEETS = ("Прибулі у ВД",)


def _accepted_on_arrival_day(record, date_value, actual):
    """День прибуття з ADDON_ARRIVAL_DAY_SHEETS (constants.py) - за прямою
    вказівкою користувача "РВЗ" (30, 100) пріоритетніший: значення з
    ADDON_ARRIVAL_DAY_STATUSES у Табелі в цей день - НЕ розбіжність (сам
    статус аркуша теж лишається правильним - _statuses_match)."""
    if record["source_sheet"] not in ADDON_ARRIVAL_DAY_SHEETS or record.get("arrival_date") != date_value:
        return False
    accepted = {str(_normalize_for_compare(status)) for status in ADDON_ARRIVAL_DAY_STATUSES}
    return str(_normalize_for_compare(actual)) in accepted


def _paint(timeline, record, start, end, dates):
    """dates - відсортовані дати шкали; проходяться лише дні самого запису
    (звірка книг між собою - шкала за кілька років, а не за місяць)."""
    for date_value in dates[bisect_left(dates, start):bisect_right(dates, end)]:
        timeline[date_value] = record


# Остаточні/адміністративні статуси - не можуть "повернутись" після
# пізніших подій: будь-який новіший період означає, що вони вже не діють
# (людина знову на службі чи запис book61 застарів). Реальний випадок,
# підтверджений користувачем: "Звіль" з Алф2 за 2025 рік (відкритий період)
# "повертався" після закінчення ВЛК у 2026, хоча пізніші ШП/ВП/ВЛК прямо
# свідчать, що людина знову на службі. "СЗЧ" - так само: пізніша
# відпустка/відрядження означає, що людина вже повернулась (реальний
# випадок: СЗЧ з лютого 2026 "повертався" після ВП у жовтні).
_SUPERSEDED_BY_LATER_STATUSES = ("Звіль", *TRANSFER_STATUSES, "полон", "Загиб", "Смерть", "БЗ", "Розп", "СЗЧ")

# Відкритий період (кінця в книзі немає) решти статусів так само закінчується
# перед пізнішою подією - реальний випадок (книга ЦВС): лікування без "Дата
# виписки" поверталось після пізніших ВЛК/ВП і тягнулось до кінця місяця.
# Виняток - ППД: людина повертається туди після лікування/ТВО.
_RESUMING_OPEN_STATUSES = ("ППД",)


def _ends_at_next_event(record):
    return record["status"] in _SUPERSEDED_BY_LATER_STATUSES or (
        record["end_date"] is None and record["status"] not in _RESUMING_OPEN_STATUSES
    )


def _known_end_of_open_period(record, records):
    """Кінець відкритого періоду, відомий з ІНШОГО запису того самого
    статусу (інша книга знає завершення того самого епізоду) - найраніший
    кінець такого запису, що не закінчився до початку цього; None - немає.
    Реальний випадок, підтверджений користувачем: book61 "ВЛК" з 14.10 без
    "Кінець (ВЛК)", ЦВС - ВЛК 14.10-20.10 (документ "від 20.10.2026") - ВЛК
    тягнулась до кінця місяця, бо запис ЦВС починався того самого дня, а
    не пізніше (_ends_at_next_event його не враховує)."""
    if record["end_date"] is not None or record["status"] in _RESUMING_OPEN_STATUSES:
        return None
    known_ends = [
        other["end_date"] for other in records
        if other is not record and other["status"] == record["status"]
        and other["end_date"] is not None and other["end_date"] >= record["start_date"]
    ]
    return min(known_ends) if known_ends else None


def _paint_timeline(records, date_columns, last_date):
    """{дата Табеля: period-запис, що її покриває} - ОДНА послідовна шкала
    з УСІХ періодів людини (усі книги resources/addons). Записи накладаються
    в порядку дати початку, пізніший - поверх попереднього: кожен період
    фактично триває до дня перед початком наступного ("ППД -> ШП -> СЗЧ");
    якщо ж пізніший закінчився раніше, ніж (за самою книгою) мав
    закінчитись попередній, - попередній продовжується після нього (напр.
    ППД після ТВО).

    Винятки (_ends_at_next_event): _SUPERSEDED_BY_LATER_STATUSES і
    відкриті періоди (end_date - None, крім _RESUMING_OPEN_STATUSES)
    закінчуються в день перед БУДЬ-ЯКОЮ пізнішою подією й ніколи не
    продовжуються після неї; відкритий період, чий кінець знає інша книга
    (_known_end_of_open_period), закінчується там само; записи з
    _PRIORITY_SOURCE_SHEETS ("Прибулі у ВД") накладаються останніми, поверх
    усього. Відкритий період без пізніших подій - до останньої колонки-дати."""
    ordered = sorted(records, key=lambda r: (r["source_sheet"] in _PRIORITY_SOURCE_SHEETS, r["start_date"]))
    dates = sorted(date_value for _col_idx, date_value in date_columns)
    timeline = {}
    for record in ordered:
        start = record["start_date"]
        end = record["end_date"] or _known_end_of_open_period(record, records) or last_date
        if _ends_at_next_event(record):
            later_starts = [other["start_date"] for other in records if other["start_date"] > start]
            if later_starts:
                end = min(end, min(later_starts) - timedelta(days=1))
        _paint(timeline, record, start, end, dates)
    return timeline


def find_period_mismatches(timesheet, period_records):
    """Для КОЖНОЇ людини з Табеля - ОДНА послідовна шкала очікуваних
    статусів з УСІХ її book61-аркушів (_paint_timeline), звірена з
    фактичним статусом у Табелі на кожну дату-колонку - ОДИН рядок звіту на
    людину. За прямою вказівкою користувача (реальний випадок): book61
    "ППД" 13.08-04.10, "ШП" 03.10-06.10, "СЗЧ" з 06.10 - раніше ДВА окремі
    рядки звіту з датами, що перекривались; тепер ОДИН: ППД до 02.10 -> ШП
    03.10-05.10 -> СЗЧ з 06.10 (до наступного статусу, напр. "Розп").

    Людина, якої НЕМАЄ в Табелі (normalize_name не серед person_rows), -
    НЕ звіряється ВЗАГАЛІ, з жодного аркуша book61 - за прямою вказівкою
    користувача, БЕЗ винятків: Табель - ЄДИНЕ джерело того, КОГО звіряти,
    book61 лише каже, що ОЧІКУЄТЬСЯ для вже наявних у ньому людей.
    Реальні випадки: люди з давніми записами "СЗЧ" та "ВД", яких давно
    немає в Табелі, хибно потрапляли в mis_match.xlsx - раніше відсутність
    у Табелі для ЧАСТИНИ аркушів (ВП/ВД/ШП/ВЛК/ППД) рахувалась
    розбіжністю - це правило ПОВНІСТЮ скасоване.

    Шкала ОДНА на людину - навіть коли її записи йдуть із КІЛЬКОХ аркушів
    чи КІЛЬКОХ addon-файлів (правила _paint_timeline діють через усі книги);
    за прямою вказівкою користувача звіт розділено за addon-файлами: кожна
    дата належить файлу, чий запис визначив очікуваний статус, і на кожен
    файл - ОКРЕМИЙ запис людини (один рядок на аркуші цього файлу, без
    дублювання ПІБ).

    Повертає список {"source_file", "sources" ("файл: аркуш" цього файлу у
    тому порядку, в якому вони змінюють одне одного на шкалі - для колонки
    "Джерело"), "roster_pidrozdil" (ПІДРОЗДІЛ з Табеля, для колонки
    "Підрозділ"), "roster_pib_raw", "addon_pib_raw", "dates": {date:
    (actual, expected)}} - лише пари людина+файл, де є ХОЧА Б ОДНА справжня
    розбіжність (_statuses_match - False)."""
    person_rows = timesheet.person_rows()
    pib_col = timesheet.label_columns[PIB_COLUMN_NAME]
    date_columns = timesheet.date_columns
    if not date_columns:
        return []
    last_date = date_columns[-1][1]

    grouped = {}
    for record in period_records:
        grouped.setdefault(normalize_name(record["pib_raw"]), []).append(record)

    results = []
    for normalized, records in sorted(grouped.items()):
        row_idx = person_rows.get(normalized)
        if row_idx is None:
            continue
        timeline = _paint_timeline(records, date_columns, last_date)

        by_file = {}
        for col_idx, date_value in date_columns:
            record = timeline.get(date_value)
            if record is None:
                continue
            actual = timesheet.ws.cell(row=row_idx, column=col_idx).value
            # Дата ще ПОРОЖНЯ (None) - реальний випадок: дата ПІСЛЯ
            # останнього наявного рапорту, apply_events (content/
            # oblik_timesheet.py) НАВМИСНО лишає такі колонки порожніми.
            # Це НЕ розбіжність (ще просто НЕМАЄ даних, щоб судити).
            if actual is None:
                continue
            expected = actual if _accepted_on_arrival_day(record, date_value, actual) else record["status"]
            entry = by_file.setdefault(record["source_file"], {"dates": {}, "sources": [], "pib_raw": record["pib_raw"]})
            entry["dates"][date_value] = (actual, expected)
            source = f"{record['source_file']}: {record['source_sheet']}"
            if not entry["sources"] or entry["sources"][-1] != source:
                entry["sources"].append(source)

        for source_file, entry in by_file.items():
            if all(_statuses_match(expected, actual) for actual, expected in entry["dates"].values()):
                continue
            results.append({
                "source_file": source_file,
                "sources": entry["sources"],
                "roster_pidrozdil": _roster_pidrozdil(timesheet, row_idx),
                "roster_pib_raw": timesheet.ws.cell(row=row_idx, column=pib_col).value,
                "addon_pib_raw": entry["pib_raw"],
                "dates": entry["dates"],
            })
    return results


def _roster_pidrozdil(timesheet, row_idx):
    """ПІДРОЗДІЛ людини з Табеля - None, якщо в Табелі немає такої колонки."""
    pidrozdil_col = timesheet.label_columns.get("ПІДРОЗДІЛ")
    return timesheet.ws.cell(row=row_idx, column=pidrozdil_col).value if pidrozdil_col else None


def unmapped_people_in_roster(timesheet, unmapped_people):
    """Лише ті з content/addon_reader.py::read_unmapped_alf2_people, хто Є в
    Табелі - те саме правило, що й find_period_mismatches: кого немає в
    Табелі, той не звіряється з resources/addons узагалі, навіть як застереження.
    Кожна людина - ОДИН раз (Алф2 може мати кілька рядків з тим самим ПІБ;
    за прямою вказівкою користувача ПІБ у звіті не дублюються); до запису
    додається її ПІДРОЗДІЛ з Табеля ("roster_pidrozdil")."""
    person_rows = timesheet.person_rows()
    kept = {}
    for person in unmapped_people:
        normalized = normalize_name(person["pib_raw"])
        if normalized in person_rows and normalized not in kept:
            kept[normalized] = {**person, "roster_pidrozdil": _roster_pidrozdil(timesheet, person_rows[normalized])}
    return list(kept.values())


_UNMAPPED_WARNING_HEADERS = ["Підрозділ", "ПІБ (book61)", "Статус в Алф2 (book61)", "Примітка"]
_UNMAPPED_WARNING_COLUMN_WIDTHS = (26, 30, 24, 70)


def _source_label(record):
    """Колонка "Джерело": "book61: ППД -> book61: ШП -> book61: СЗЧ"."""
    return " -> ".join(record["sources"])


def _write_period_mismatch_rows(ws, records, start_row, date_col_by_date):
    current_row = start_row
    for record in records:
        row_a, row_b = current_row, current_row + 1
        ws.cell(row=row_a, column=1, value="Табель (ОБЛІК.xlsx)")
        ws.cell(row=row_b, column=1, value=_source_label(record))

        ws.cell(row=row_a, column=2, value=_display_value(record["roster_pidrozdil"]))
        ws.merge_cells(start_row=row_a, start_column=2, end_row=row_b, end_column=2)

        roster_display = _display_value(record["roster_pib_raw"])
        addon_display = _display_value(record["addon_pib_raw"])
        cell_a = ws.cell(row=row_a, column=3, value=roster_display)
        if _normalize_for_compare(record["roster_pib_raw"]) == _normalize_for_compare(record["addon_pib_raw"]):
            ws.merge_cells(start_row=row_a, start_column=3, end_row=row_b, end_column=3)
            cell_a.fill = _MATCH_FILL
        else:
            cell_a.fill = _MISMATCH_FILL
            ws.cell(row=row_b, column=3, value=addon_display).fill = _MISMATCH_FILL

        for date_value, (actual, expected) in record["dates"].items():
            col_idx = date_col_by_date[date_value]
            cell_actual = ws.cell(row=row_a, column=col_idx, value=_display_value(actual))
            if _statuses_match(expected, actual):
                ws.merge_cells(start_row=row_a, start_column=col_idx, end_row=row_b, end_column=col_idx)
                cell_actual.fill = _MATCH_FILL
            else:
                cell_actual.fill = _MISMATCH_FILL
                ws.cell(row=row_b, column=col_idx, value=_display_value(expected)).fill = _MISMATCH_FILL

        for row in (row_a, row_b):
            ws.row_dimensions[row].height = 20
        current_row += 2
    return current_row


_UNMAPPED_SHEET_TITLE = "Застереження"
_NO_MISMATCHES_TEXT = "Розбіжностей з цим файлом немає."
_SHEET_TITLE_FORBIDDEN_RE = re.compile(r"[\[\]:*?/\\]")
_SHEET_TITLE_MAX_LENGTH = 31


def _sheet_title(name, used):
    """Назва аркуша з імені addon-файлу за правилами Excel: без []:*?/\\, не
    довше 31 символу, унікальна в книзі (used - уже зайняті, у нижньому
    регістрі)."""
    base = _SHEET_TITLE_FORBIDDEN_RE.sub("_", str(name)).strip() or "addon"
    title = base[:_SHEET_TITLE_MAX_LENGTH]
    number = 1
    while title.lower() in used:
        number += 1
        suffix = f" ({number})"
        title = base[:_SHEET_TITLE_MAX_LENGTH - len(suffix)] + suffix
    used.add(title.lower())
    return title


def write_addon_mismatch_report(period_mismatches, unmapped_people, output_path, source_files=()):
    """Будує output_path (mis_match.xlsx) - за прямою вказівкою користувача
    ОКРЕМИЙ аркуш на кожен addon-файл (назва аркуша - ім'я файлу, записи -
    за "source_file"): розбіжності періодів у тому самому "два рядки,
    зелена/червона заливка на весь рядочок" стилі, що й error_mis_statuses.xlsx
    (_write_period_sheet). source_files - усі звірені файли: файл без
    розбіжностей теж отримує аркуш з написом _NO_MISMATCHES_TEXT (видно, що
    його звірено). ЯКЩО є хоч один нерозпізнаний статус Алф2 ("самостріл") -
    ОКРЕМИЙ аркуш "Застереження" (текстові рядки, не по датах -
    content/addon_reader.py::read_unmapped_alf2_people).

    Повертає output_path, АБО None, якщо розповідати нема про що взагалі
    (жодної розбіжності й жодного попередження) - той самий принцип, що й
    write_mismatch_report."""
    if not period_mismatches and not unmapped_people:
        return None

    by_file = {source_file: [] for source_file in source_files}
    for record in period_mismatches:
        by_file.setdefault(record["source_file"], []).append(record)

    wb = Workbook()
    wb.remove(wb.active)
    used_titles = set()
    for source_file, records in by_file.items():
        _write_period_sheet(wb.create_sheet(_sheet_title(source_file, used_titles)), records)

    if unmapped_people:
        _write_unmapped_sheet(wb.create_sheet(_sheet_title(_UNMAPPED_SHEET_TITLE, used_titles)), unmapped_people)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    try:
        wb.save(output_path)
    except PermissionError as e:
        raise PermissionError(
            f"Не вдалось зберегти '{output_path}': файл зараз відкритий в іншій програмі "
            "(напр. Excel). Закрийте його та запустіть генерацію знову."
        ) from e
    return output_path


def _write_period_sheet(ws, records):
    """Аркуш одного addon-файлу: Джерело/Підрозділ/ПІБ + колонка на кожну
    дату його записів (_write_period_mismatch_rows); без записів - лише
    шапка й _NO_MISMATCHES_TEXT."""
    all_dates = sorted({date_value for record in records for date_value in record["dates"]})

    date_headers = [datetime.combine(d, datetime.min.time()) for d in all_dates]
    headers = ["Джерело", "Підрозділ", "ПІБ"] + date_headers
    for col_idx, value in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=value)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGNMENT
        cell.border = _BORDER
        if isinstance(value, datetime):
            cell.number_format = DATE_NUMBER_FORMAT
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 45

    date_col_by_date = {d: idx for idx, d in enumerate(all_dates, start=4)}
    if records:
        next_row = _write_period_mismatch_rows(ws, records, 2, date_col_by_date)
    else:
        ws.cell(row=2, column=1, value=_NO_MISMATCHES_TEXT)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
        next_row = 3

    for row in range(2, next_row):
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row, column=col_idx)
            cell.font = _FONT
            cell.alignment = _CENTER
            cell.border = _BORDER

    ws.column_dimensions["A"].width = max([30] + [len(_source_label(record)) + 2 for record in records])
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 30
    for col_idx in range(4, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 24


def _write_unmapped_sheet(warn_ws, unmapped_people):
    for col_idx, value in enumerate(_UNMAPPED_WARNING_HEADERS, start=1):
        cell = warn_ws.cell(row=1, column=col_idx, value=value)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGNMENT
        cell.border = _BORDER
    warn_ws.freeze_panes = "A2"
    warn_row = 2
    for person in unmapped_people:
        warn_ws.cell(row=warn_row, column=1, value=_display_value(person["roster_pidrozdil"]))
        warn_ws.cell(row=warn_row, column=2, value=person["pib_raw"])
        warn_ws.cell(row=warn_row, column=3, value=person["alf2_status"])
        warn_ws.cell(row=warn_row, column=4, value=(
            "Статус book61 (Алф2) не зіставлено з жодним статусом ОБЛІК у довіднику - перевірити вручну."
        ))
        warn_row += 1
    for row in range(2, warn_row):
        for col_idx in range(1, len(_UNMAPPED_WARNING_HEADERS) + 1):
            cell = warn_ws.cell(row=row, column=col_idx)
            cell.font = _FONT
            cell.alignment = _CENTER
            cell.border = _BORDER
    for col_idx, width in enumerate(_UNMAPPED_WARNING_COLUMN_WIDTHS, start=1):
        warn_ws.column_dimensions[get_column_letter(col_idx)].width = width
