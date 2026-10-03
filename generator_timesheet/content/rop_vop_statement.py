import calendar
import os
from collections import defaultdict
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from constants import (
    FULL_MILITARY_UNIT,
    FULL_UNIT_BUT,
    MONTH_NAMES_NOMINATIVE_UPPER,
    PIB_COLUMN_NAME,
    RECOGNIZED_SUBDIVISIONS,
    SCHEDULE_VOP_ALIASES,
)

# Категорії виплати (числа, ВЖЕ підставлені apply_payment_values,
# content/oblik_timesheet.py) - за прямою вказівкою користувача, 70 = РОП
# (ротний опорний пункт), 170 = ВОП (взводний опорний пункт) - STATUS_MEANINGS
# (constants.py) підтверджує той самий сенс словами. Будь-яке ІНШЕ значення
# клітинки (30/100/текстовий статус) - порожня клітинка у Відомості.
_ROP_CATEGORY = 70
_VOP_CATEGORY = 170


def _month_columns(timesheet, year, month):
    """[(col_idx, номер_дня), ...] - колонки-дати timesheet, що належать
    year/month, відсортовані за днем - спільна основа для
    _compute_day_marks/build_statement_rows нижче."""
    return sorted(
        (col_idx, date_value.day)
        for col_idx, date_value in timesheet.date_columns
        if date_value.year == year and date_value.month == month
    )


def _compute_day_marks(timesheet, year, month):
    """{row_idx: {день: "роп"/"воп"}} для КОЖНОЇ людини роcтера (БЕЗ фільтру
    "хоча б один день" - на відміну від build_statement_rows нижче, яка
    саме ЦЕЙ фільтр застосовує). Порожній словник для людини без жодного
    дня РОП/ВОП - навмисно не пропускається взагалі (а не відсутній ключ),
    щоб find_schedule_mismatches (content/schedule_reader.py звіряння)
    однаково легко перебирав УСІХ людей роcтера, а не лише тих, хто вже
    потрапив би у Відомість."""
    month_cols = _month_columns(timesheet, year, month)
    result = {}
    for row_idx in timesheet.person_rows().values():
        day_marks = {}
        for col_idx, day in month_cols:
            value = timesheet.ws.cell(row=row_idx, column=col_idx).value
            if value == _ROP_CATEGORY:
                day_marks[day] = "роп"
            elif value == _VOP_CATEGORY:
                day_marks[day] = "воп"
        result[row_idx] = day_marks
    return result


def build_statement_rows(timesheet, year, month):
    """Для КОЖНОЇ людини роcтера - позначки "роп"/"воп" на КОЖЕН день
    year/month, де клітинка-дата timesheet ВЖЕ має відповідну категорію
    виплати (70/170 - apply_payment_values мала відпрацювати ДО виклику цієї
    функції), і лічильники днів РОП/ВОП. Людина потрапляє в результат, ЛИШЕ
    якщо має ХОЧА Б ОДИН такий день за місяць (реальний зразок - resources/
    {SHORT_UNIT_BATTALION без пробілу}ВОП-РОП_ЛИПЕНЬ_.xlsx - той самий фільтр: НЕ весь роcтер, а підмножина
    з хоча б одним днем РОП/ВОП). Порядок результату - порядок роcтера
    (person_rows зберігає порядок рядків аркуша)."""
    days_in_month = calendar.monthrange(year, month)[1]
    all_day_marks = _compute_day_marks(timesheet, year, month)

    pib_col = timesheet.label_columns[PIB_COLUMN_NAME]
    posada_col = timesheet.label_columns.get("ПОСАДА")
    zvannya_col = timesheet.label_columns.get("ЗВАННЯ")

    records = []
    for row_idx, day_marks in all_day_marks.items():
        if not day_marks:
            continue

        rop_count = sum(1 for status in day_marks.values() if status == "роп")
        vop_count = sum(1 for status in day_marks.values() if status == "воп")
        records.append({
            "pib_raw": timesheet.ws.cell(row=row_idx, column=pib_col).value,
            "posada": timesheet.ws.cell(row=row_idx, column=posada_col).value if posada_col else None,
            "zvannya": timesheet.ws.cell(row=row_idx, column=zvannya_col).value if zvannya_col else None,
            "day_marks": day_marks,
            "rop_count": rop_count,
            "vop_count": vop_count,
        })

    return records, days_in_month


_NOT_IN_ROSTER_LABEL = "Не знайдено в ОБЛІК.xlsx"


def _schedule_statuses_equivalent(computed_status, schedule_status):
    """Дві сторони одного дня вважаються ЕКВІВАЛЕНТНИМИ (не розбіжністю),
    якщо вони буквально збігаються, АБО обидві - псевдоніми ВОП
    (SCHEDULE_VOP_ALIASES, constants.py, керовано користувачем самостійно) -
    за прямою вказівкою користувача: "воп"/"тмп"/"тот" - РІЗНІ слова для
    ТІЄЇ САМОЇ категорії виплати 170, тоді як обчислене (_compute_day_marks)
    рахує ЛИШЕ ОДНУ текстову форму - "воп". Обидва аргументи мають бути
    справжніми (не None) псевдонімами - "не звірялось" (schedule_status is
    None) і "жодного дня на ВОП взагалі" (computed_status is None)
    обробляються ОКРЕМО, до виклику цієї функції."""
    if computed_status == schedule_status:
        return True
    return computed_status in SCHEDULE_VOP_ALIASES and schedule_status in SCHEDULE_VOP_ALIASES


def find_schedule_mismatches(timesheet, year, month, schedule_marks, schedule_people, recognized_subdivisions=RECOGNIZED_SUBDIVISIONS):
    """За прямою вказівкою користувача - звіряє ЩОЙНО ОБЧИСЛЕНІ (з
    information_unit, apply_payment_values - _compute_day_marks вище)
    позначки "роп"/"воп" з "офіційною" відомістю resources/schedule
    (content/schedule_reader.read_schedule_marks - marks: {(normalize_name(
    ПІБ), день): "роп"/"воп"}, people: {normalize_name(ПІБ): {"pib_raw",
    "posada", "zvannya", "pidrozdil"}}).

    За прямою вказівкою користувача - порівнюються ЛИШЕ дні, де resources/
    schedule Є НЕПОРОЖНІМ (файл подається щодня й поступово заповнюється,
    тож дні, ще не заповнені там, - НЕ "стверджено порожньо", а просто "ще
    не дійшли" - порівнювати їх було б передчасно). День, де schedule
    непорожній, - розбіжність, якщо ОБЧИСЛЕНЕ значення НЕ ЕКВІВАЛЕНТНЕ
    (_schedule_statuses_equivalent - враховує SCHEDULE_VOP_ALIASES,
    constants.py: "воп"/"тмп"/"тот" - РІЗНІ слова тієї самої категорії 170,
    НЕ розбіжність між собою) schedule-значенню (в т.ч. коли обчислене -
    взагалі порожньо - завжди розбіжність, "порожньо" не входить у жодний
    псевдонім). Дні, де schedule ПОРОЖНІЙ, - НЕ порівнюються взагалі,
    НАВІТЬ якщо обчислено щось непорожнє (це НЕ "підрозділ ще не подав
    дані" - обчислене значення там МОЖЕ бути цілком правильним, просто
    "офіційна" відомість ще не дійшла до цього дня).

    ОКРЕМО - записи resources/schedule, чий normalize_name НЕ збігається з
    ЖОДНИМ роcтеровим рядком (typo в ПІБ чи людина відсутня в роcтері) - той
    самий принцип, що й _missing_from_roster_records
    (content/payment_mismatch_checker.py): роcтерова сторона (Посада/
    Звання/ПІБ) отримує _NOT_IN_ROSTER_LABEL, а schedule-сторона - РЕАЛЬНІ
    дані з schedule_people (а НЕ ще один _NOT_IN_ROSTER_LABEL) - за прямою
    вказівкою користувача, щоб було видно, ХТО саме ця людина, а не лише
    факт "не знайдено" на обидві клітинки одразу.

    Повертає список {
        "roster_pib_raw", "roster_posada", "roster_zvannya",
        "schedule_pib_raw", "schedule_posada", "schedule_zvannya",
        "pidrozdil", "dates": {date: (обчислено, schedule)},
        "full_days": {date: (обчислено, schedule)},
    } ЛИШЕ для людей/записів із ХОЧА Б ОДНІЄЮ розбіжною датою.

    "dates" - лише РОЗБІЖНІ дати (визначає, чи людина взагалі потрапляє в
    результат, і які колонки-дати матиме звіт - write_schedule_mismatch_
    report). "full_days" - УСІ ПОРІВНЮВАНІ дати цієї людини (schedule
    непорожній), включно зі ЗБІЖНИМИ - звіт має показувати ЦІ збіжні дні
    теж (зеленим), а не лишати їх порожньою клітинкою, якщо вони потрапили
    в колонки звіту через розбіжність ІНШОЇ людини.

    "pidrozdil" - ІНФОРМАЦІЙНЕ поле (НЕ звіряється роcтер/schedule, на
    відміну від Посада/Звання/ПІБ) - роcтерове значення колонки ПІДРОЗДІЛ
    для знайдених людей, або дужкова дописка прикріплення з schedule для НЕ
    знайдених. НЕ звіряється навмисно: roster.ПІДРОЗДІЛ - "рідна" колонка,
    заповнена для КОЖНОЇ людини роcтера, тоді як schedule фіксує дужкою
    ЛИШЕ ВИНЯТКОВІ прикріплення з ІНШОГО підрозділу - порівняння цих двох
    як "розбіжність" лише тому, що schedule про це мовчить, дало б хибну
    розбіжність практично в КОЖНОГО знайденого запису.

    За прямою вказівкою користувача - людина взагалі НЕ потрапляє в
    результат, якщо ЇЇ pidrozdil (роcтерове ПІДРОЗДІЛ для ЗНАЙДЕНИХ людей,
    дужкова дописка прикріплення зі schedule для НЕ ЗНАЙДЕНИХ) -
    НЕПОРОЖНІЙ текст, якого НЕМАЄ серед recognized_subdivisions (за
    замовчуванням - RECOGNIZED_SUBDIVISIONS, constants.py: "бойові" роти/
    батареї/взводи зі СПІЛЬНОГО з generator_br_and_report_for_money переліку
    - resources/data.json::company_sections.SEQUENCE того проєкту, а не
    штаб/управління тощо). Реальний випадок для НЕ знайдених - дописка
    прикріплення часто вказує на людину з ІНШОГО, стороннього підрозділу
    (НЕ рота/взвод/батарея ЦЬОГО батальйону) - її відсутність у роcтері
    ОЧІКУВАНА (вона й не має тут бути), а НЕ помилка, тож звірка НЕ
    повинна її показувати взагалі.
    pidrozdil ВІДСУТНІЙ (роcтер узагалі не має колонки ПІДРОЗДІЛ/ця
    клітинка порожня для знайденої людини, АБО schedule не має дужкової
    дописки для не знайденої) - фільтр НЕ застосовується (немає даних -
    немає підстави виключати, той самий принцип, що й Посада/Звання:
    відсутнє значення НЕ трактується як помилкове). ЗНАЙДЕНА людина, ЯКУ
    виключено фільтром, ВИКЛЮЧАЄТЬСЯ ЦІЛКОМ (а не лише позначається якось
    інакше) - її normalize_name усе одно рахується "знайденим"
    (matched_normalized), тож вона НЕ з'явиться і в другому проході (не
    породжує фантомний "не знайдено в ОБЛІК.xlsx" запис)."""
    days_in_month = calendar.monthrange(year, month)[1]
    all_day_marks = _compute_day_marks(timesheet, year, month)
    person_rows = timesheet.person_rows()

    schedule_by_person = defaultdict(dict)
    for (normalized, day), status in schedule_marks.items():
        schedule_by_person[normalized][day] = status

    pib_col = timesheet.label_columns[PIB_COLUMN_NAME]
    posada_col = timesheet.label_columns.get("ПОСАДА")
    zvannya_col = timesheet.label_columns.get("ЗВАННЯ")
    pidrozdil_col = timesheet.label_columns.get("ПІДРОЗДІЛ")

    records = []
    matched_normalized = set()
    for normalized, row_idx in person_rows.items():
        computed_days = all_day_marks.get(row_idx, {})
        schedule_days = schedule_by_person.get(normalized, {})
        if normalized in schedule_by_person:
            matched_normalized.add(normalized)

        roster_pidrozdil = timesheet.ws.cell(row=row_idx, column=pidrozdil_col).value if pidrozdil_col else None
        if isinstance(roster_pidrozdil, str) and roster_pidrozdil.strip() and roster_pidrozdil.strip().lower() not in recognized_subdivisions:
            continue

        full_days = {
            date(year, month, day): (computed_days.get(day), schedule_status)
            for day, schedule_status in schedule_days.items()
            if day <= days_in_month
        }
        dates = {d: pair for d, pair in full_days.items() if not _schedule_statuses_equivalent(*pair)}

        if not dates:
            continue
        schedule_info = schedule_people.get(normalized, {})
        records.append({
            "roster_pib_raw": timesheet.ws.cell(row=row_idx, column=pib_col).value,
            "roster_posada": timesheet.ws.cell(row=row_idx, column=posada_col).value if posada_col else None,
            "roster_zvannya": timesheet.ws.cell(row=row_idx, column=zvannya_col).value if zvannya_col else None,
            "schedule_pib_raw": schedule_info.get("pib_raw"),
            "schedule_posada": schedule_info.get("posada"),
            "schedule_zvannya": schedule_info.get("zvannya"),
            "pidrozdil": roster_pidrozdil,
            "dates": dates,
            "full_days": full_days,
        })

    for normalized, schedule_days in schedule_by_person.items():
        if normalized in matched_normalized:
            continue

        schedule_info = schedule_people.get(normalized, {})
        schedule_pidrozdil = schedule_info.get("pidrozdil")
        if isinstance(schedule_pidrozdil, str) and schedule_pidrozdil.strip() and schedule_pidrozdil.strip().lower() not in recognized_subdivisions:
            continue

        full_days = {
            date(year, month, day): (None, status)
            for day, status in schedule_days.items()
            if day <= days_in_month
        }
        if not full_days:
            continue
        records.append({
            "roster_pib_raw": _NOT_IN_ROSTER_LABEL,
            "roster_posada": _NOT_IN_ROSTER_LABEL,
            "roster_zvannya": _NOT_IN_ROSTER_LABEL,
            "schedule_pib_raw": schedule_info.get("pib_raw"),
            "schedule_posada": schedule_info.get("posada"),
            "schedule_zvannya": schedule_info.get("zvannya"),
            "pidrozdil": schedule_pidrozdil,
            "dates": full_days,
            "full_days": full_days,
        })

    return records


_HEADER_FONT = Font(name="Times New Roman", size=11, bold=True)
_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
_BODY_FONT = Font(name="Times New Roman", size=11)
_CENTER = Alignment(horizontal="center", vertical="center")
_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_LABEL_HEADERS = ["№\nз/п", "Військове звання", "Посада", "Прізвище,\nвласне ім'я"]
_LABEL_COLUMN_WIDTHS = (6, 20, 30, 30)
_DAY_COLUMN_WIDTH = 4
# За прямою вказівкою користувача - ТО/ТМП/ТП немає джерела даних у
# generator_timesheet (лише РОП/ВОП обчислюються), тож лічильники цих трьох
# категорій завжди 0 - колонки лишено (структура зразка), значення - навмисно
# статичне.
_COUNT_HEADERS = ["кількість\n днів на ВОП", "кількість\n днів на ТО", "кількість\n днів на ТМП", "кількість\n днів на ТП", "кількість\n днів на РОП"]
_COUNT_COLUMN_WIDTH = 10


def write_statement(records, days_in_month, output_path, year, month):
    """Будує output_path (.xlsx) - за зразком реального файлу resources/
    {SHORT_UNIT_BATTALION без пробілу}ВОП-РОП_ЛИПЕНЬ_.xlsx: заголовок (FULL_UNIT_BUT/FULL_MILITARY_UNIT з
    resources/data.json + назва місяця/рік), колонки № з/п / Військове
    звання / Посада / ПІБ / дні 1..days_in_month / 5 лічильників (ВОП, ТО,
    ТМП, ТП, РОП - лише ВОП/РОП обчислюються, решта завжди 0), підписний
    блок наприкінці (фіксований текст, без реальних даних). НОВА книга (не
    завантаження зразка) - той самий підхід, що й write_mismatch_report
    (content/payment_mismatch_checker.py): не залежить від того, скільки
    людей увійде цього місяця, немає потреби зсувати підписний блок.

    Повертає output_path."""
    month_name = MONTH_NAMES_NOMINATIVE_UPPER[month]

    wb = Workbook()
    ws = wb.active
    ws.title = month_name

    day_headers = [f"{day:02d}" for day in range(1, days_in_month + 1)]
    headers = _LABEL_HEADERS + day_headers + _COUNT_HEADERS
    last_col = len(headers)

    ws.cell(row=2, column=1, value="Відомість").font = _HEADER_FONT
    ws.cell(row=3, column=1, value="обліку днів участі військовослужбовців у виконанні бойових (спеціальних завдань)").font = _HEADER_FONT
    ws.cell(row=4, column=1, value=f"{FULL_UNIT_BUT} {FULL_MILITARY_UNIT} за {month_name} {year} року").font = _HEADER_FONT
    for row in (2, 3, 4):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
        ws.cell(row=row, column=1).alignment = _CENTER

    header_row = 6
    for col_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=text)
        cell.font = _HEADER_FONT
        cell.alignment = _HEADER_ALIGNMENT
        cell.border = _BORDER
        if col_idx <= len(_LABEL_HEADERS):
            width = _LABEL_COLUMN_WIDTHS[col_idx - 1]
        elif col_idx <= len(_LABEL_HEADERS) + days_in_month:
            width = _DAY_COLUMN_WIDTH
        else:
            width = _COUNT_COLUMN_WIDTH
        ws.column_dimensions[cell.column_letter].width = width
    ws.freeze_panes = ws.cell(row=header_row + 1, column=len(_LABEL_HEADERS) + 1).coordinate

    current_row = header_row + 1
    for number, record in enumerate(records, start=1):
        row_values = [number, record["zvannya"], record["posada"], record["pib_raw"]]
        row_values += [record["day_marks"].get(day, None) for day in range(1, days_in_month + 1)]
        row_values += [record["vop_count"], 0, 0, 0, record["rop_count"]]
        for col_idx, value in enumerate(row_values, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=value)
            cell.font = _BODY_FONT
            cell.alignment = _CENTER
            cell.border = _BORDER
        current_row += 1

    _write_signature_block(ws, current_row + 1, last_col)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    try:
        wb.save(output_path)
    except PermissionError as e:
        raise PermissionError(
            f"Не вдалось зберегти '{output_path}': файл зараз відкритий в іншій програмі "
            "(напр. Excel). Закрийте його та спробуйте ще раз."
        ) from e
    return output_path


def _write_signature_block(ws, start_row, last_col):
    """Фіксований підписний блок (текст ролей, без реальних даних - "___" -
    буквальний текст зразка, заповнюється вручну на папері) - той самий
    вигляд, що й у resources/{SHORT_UNIT_BATTALION без пробілу}ВОП-РОП_ЛИПЕНЬ_.xlsx."""
    rows = [
        ("Начальник штабу  ___ батальйону (дивізіону)", None),
        ("_________________", "____________________"),
        ("(військове звання)", "(ім'я та прізвище)"),
        (None, None),
        ("Особа відповідальна за облік особового складу", None),
        ("_________________", "____________________"),
        ("(військове звання)", "(ім'я та прізвище)"),
    ]
    for offset, (left, right) in enumerate(rows):
        row = start_row + offset
        if left is not None:
            cell = ws.cell(row=row, column=1, value=left)
            cell.font = _BODY_FONT
        if right is not None:
            cell = ws.cell(row=row, column=4, value=right)
            cell.font = _BODY_FONT


# Стиль - той самий "вигляд error_mis_statuses.xlsx" (Times New Roman 14,
# зелена/червона заливка), що й content/payment_mismatch_checker.py - за
# прямою вказівкою користувача, "щось на зразок error_mis_status, тільки з
# відомістю". ДУБЛЬОВАНО тут, а не імпортовано звідти - той самий принцип
# незалежності модулів, що й скрізь у проєкті (content/rop_vop_statement.py
# і content/payment_mismatch_checker.py не залежать одне від одного).
_MISMATCH_FONT = Font(name="Times New Roman", size=14)
_MISMATCH_HEADER_FONT = Font(name="Bahnschrift Light SemiCondensed", size=14, bold=True, color="FFFFFFFF")
_MISMATCH_HEADER_FILL = PatternFill(start_color="FF035C6B", end_color="FF035C6B", fill_type="solid")
_MISMATCH_MATCH_FILL = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")
_MISMATCH_MISMATCH_FILL = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
# Дата, ЯКА є колонкою звіту (бо розбіжна для ХОЧА Б ОДНОЇ людини), але для
# ЦІЄЇ конкретної людини schedule на цей день ПОРОЖНІЙ (не звірялась
# взагалі) - НЕ "збіг" (зелений) і НЕ "розбіжність" (червоний), окремий
# нейтральний колір, щоб не виглядало порожньою/пропущеною клітинкою.
_MISMATCH_NOT_COMPARED_FILL = PatternFill(start_color="FFF2F2F2", end_color="FFF2F2F2", fill_type="solid")
_MISMATCH_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)
_MISMATCH_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
# За прямою вказівкою користувача - заголовки дат розвернуті на 90° (текст
# знизу вгору) - при 31 колонці-даті це вдвічі компактніше по ширині, ніж
# горизонтальний заголовок "ДД.ММ.РРРР".
_MISMATCH_DATE_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True, text_rotation=90)
_MISMATCH_THIN_SIDE = Side(style="thin")
_MISMATCH_BORDER = Border(left=_MISMATCH_THIN_SIDE, right=_MISMATCH_THIN_SIDE, top=_MISMATCH_THIN_SIDE, bottom=_MISMATCH_THIN_SIDE)

_MISMATCH_COMPUTED_LABEL = "Обчислено (information_unit)"
_MISMATCH_SCHEDULE_LABEL = "resources/schedule (файл)"
_MISMATCH_EMPTY_PLACEHOLDER = "—"
_MISMATCH_LABEL_HEADERS = ["Джерело", "Підрозділ", "Посада", "Звання", "ПІБ"]
_MISMATCH_LABEL_WIDTHS = (28, 20, 30, 24, 30)
# Вужче, ніж раніше (10) - заголовок дати тепер вертикальний (текст_
# rotation=90), під нього достатньо ширини лише для короткого значення
# ("роп"/"воп"/"—") у самих рядках даних.
_MISMATCH_DATE_WIDTH = 6
_MISMATCH_HEADER_ROW_HEIGHT = 85


def _mismatch_display(value):
    return _MISMATCH_EMPTY_PLACEHOLDER if value is None else value


def _write_compared_label_cell(ws, row_a, row_b, col_idx, roster_value, schedule_value):
    """Одна "порівнювана" колонка (Посада/Звання/ПІБ) - за прямою вказівкою
    користувача, той самий принцип, що й дата-колонки (і той самий, що й
    _write_compared_cell, content/payment_mismatch_checker.py - ДУБЛЬОВАНО
    тут, а не імпортовано, той самий принцип незалежності модулів): ОДНА
    об'єднана зелена клітинка, якщо роcтер і schedule дають ОДНАКОВЕ
    значення, інакше - 2 окремих значення (по одному на рядок), червона
    заливка на ОБОХ - щоб було видно, ДЕ саме є дані, а де немає, замість
    одного напису "Не знайдено в ОБЛІК.xlsx" на обидві клітинки одразу."""
    is_match = roster_value == schedule_value
    cell_a = ws.cell(row=row_a, column=col_idx, value=_mismatch_display(roster_value))
    if is_match:
        ws.merge_cells(start_row=row_a, start_column=col_idx, end_row=row_b, end_column=col_idx)
        cell_a.fill = _MISMATCH_MATCH_FILL
    else:
        cell_a.fill = _MISMATCH_MISMATCH_FILL
        ws.cell(row=row_b, column=col_idx, value=_mismatch_display(schedule_value)).fill = _MISMATCH_MISMATCH_FILL


def write_schedule_mismatch_report(records, output_path):
    """Будує output_path (.xlsx) - "щось на зразок error_mis_statuses.xlsx,
    тільки з відомістю" (за прямою вказівкою користувача): колонки Джерело/
    Підрозділ/Посада/Звання/ПІБ, потім по ОКРЕМІЙ колонці на КОЖНУ розбіжну
    дату (лише дати, що трапляються в records["dates"] - той самий підхід,
    що й write_mismatch_report, content/payment_mismatch_checker.py;
    заголовок дати - розвернутий на 90°). Для КОЖНОГО запису - ДВА рядки
    (_MISMATCH_COMPUTED_LABEL/_MISMATCH_SCHEDULE_LABEL у колонці "Джерело").

    За прямою вказівкою користувача - Посада/Звання/ПІБ ЗВІРЯЮТЬСЯ, як і
    дата-колонки (_write_compared_label_cell): ОДНА зелена клітинка, якщо
    роcтер і schedule збігаються, інакше - 2 окремих значення (роcтер/
    schedule), червона заливка - щоб для запису "Не знайдено в ОБЛІК.xlsx"
    було видно РЕАЛЬНІ Посада/Звання/ПІБ зі schedule, а не той самий напис
    трижды поспіль. "Підрозділ" - НЕ звіряється (лише ІНФОРМАЦІЙНА об'єднана
    клітинка, roster.ПІДРОЗДІЛ для знайдених/дужкова дописка прикріплення
    для НЕ знайдених - schedule не веде "рідну" підрозділову приналежність
    для КОЖНОЇ людини, тож звіряти її, як решту полів, дало б хибну
    розбіжність практично завжди).

    За прямою вказівкою користувача - у таблиці НЕМАЄ порожніх клітинок:
    КОЖНА дата-колонка заповнюється для КОЖНОГО запису, за його
    "full_days" (а не лише "dates"), трьома можливими станами:
    - schedule на цей день для ЦІЄЇ людини ПОРОЖНІЙ (дата - колонка звіту
      через розбіжність ІНШОЇ людини, але тут просто "ще не звірялось") -
      нейтральна сіра об'єднана клітинка з "—";
    - "обчислено" й "schedule" ЕКВІВАЛЕНТНІ (_schedule_statuses_equivalent -
      буквальний збіг, АБО обидва - псевдоніми ВОП, SCHEDULE_VOP_ALIASES,
      constants.py: "воп"/"тмп"/"тот" - НЕ розбіжність між собою) - зелена
      об'єднана клітинка зі значенням schedule (як подано, без заміни на
      "воп");
    - інакше (справжня розбіжність) - 2 окремих значення, червона заливка.

    Повертає output_path, або None, якщо records порожній - за прямою
    вказівкою користувача, як і write_mismatch_report - немає сенсу
    створювати файл лише із заголовком."""
    if not records:
        return None

    all_dates = sorted({date_value for record in records for date_value in record["dates"]})

    wb = Workbook()
    ws = wb.active
    ws.title = "Розбіжності"

    date_headers = [datetime.combine(d, datetime.min.time()) for d in all_dates]
    headers = _MISMATCH_LABEL_HEADERS + date_headers
    for col_idx, value in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=value)
        cell.font = _MISMATCH_HEADER_FONT
        cell.fill = _MISMATCH_HEADER_FILL
        cell.border = _MISMATCH_BORDER
        if isinstance(value, datetime):
            cell.alignment = _MISMATCH_DATE_HEADER_ALIGNMENT
            cell.number_format = "DD.MM.YYYY"
        else:
            cell.alignment = _MISMATCH_HEADER_ALIGNMENT
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = _MISMATCH_HEADER_ROW_HEIGHT

    date_col_by_date = {d: idx for idx, d in enumerate(all_dates, start=len(_MISMATCH_LABEL_HEADERS) + 1)}

    current_row = 2
    for record in records:
        row_a, row_b = current_row, current_row + 1
        full_days = record["full_days"]

        ws.cell(row=row_a, column=1, value=_MISMATCH_COMPUTED_LABEL)
        ws.cell(row=row_b, column=1, value=_MISMATCH_SCHEDULE_LABEL)

        ws.cell(row=row_a, column=2, value=_mismatch_display(record["pidrozdil"]))
        ws.merge_cells(start_row=row_a, start_column=2, end_row=row_b, end_column=2)

        for col_idx, roster_value, schedule_value in (
            (3, record["roster_posada"], record["schedule_posada"]),
            (4, record["roster_zvannya"], record["schedule_zvannya"]),
            (5, record["roster_pib_raw"], record["schedule_pib_raw"]),
        ):
            _write_compared_label_cell(ws, row_a, row_b, col_idx, roster_value, schedule_value)

        for date_value in all_dates:
            col_idx = date_col_by_date[date_value]
            computed_status, schedule_status = full_days.get(date_value, (None, None))
            if schedule_status is None:
                cell_a = ws.cell(row=row_a, column=col_idx, value=_MISMATCH_EMPTY_PLACEHOLDER)
                ws.merge_cells(start_row=row_a, start_column=col_idx, end_row=row_b, end_column=col_idx)
                cell_a.fill = _MISMATCH_NOT_COMPARED_FILL
            elif _schedule_statuses_equivalent(computed_status, schedule_status):
                # schedule_status (не computed_status) - за прямою вказівкою
                # користувача, "воп"/"тмп"/"тот" еквівалентні, але РІЗНІ
                # слова - клітинка показує САМЕ ТЕ, що написано в schedule
                # (джерело, яке звіряється), а не завжди канонічне "воп".
                cell_a = ws.cell(row=row_a, column=col_idx, value=_mismatch_display(schedule_status))
                ws.merge_cells(start_row=row_a, start_column=col_idx, end_row=row_b, end_column=col_idx)
                cell_a.fill = _MISMATCH_MATCH_FILL
            else:
                cell_a = ws.cell(row=row_a, column=col_idx, value=_mismatch_display(computed_status))
                cell_a.fill = _MISMATCH_MISMATCH_FILL
                ws.cell(row=row_b, column=col_idx, value=_mismatch_display(schedule_status)).fill = _MISMATCH_MISMATCH_FILL

        for row in (row_a, row_b):
            ws.row_dimensions[row].height = 20
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row, column=col_idx)
                cell.font = _MISMATCH_FONT
                cell.alignment = _MISMATCH_CENTER
                cell.border = _MISMATCH_BORDER

        current_row += 2

    for col_idx in range(1, len(headers) + 1):
        letter = ws.cell(row=1, column=col_idx).column_letter
        width = _MISMATCH_LABEL_WIDTHS[col_idx - 1] if col_idx <= len(_MISMATCH_LABEL_HEADERS) else _MISMATCH_DATE_WIDTH
        ws.column_dimensions[letter].width = width

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    try:
        wb.save(output_path)
    except PermissionError as e:
        raise PermissionError(
            f"Не вдалось зберегти '{output_path}': файл зараз відкритий в іншій програмі "
            "(напр. Excel). Закрийте його та спробуйте ще раз."
        ) from e
    return output_path
