import calendar
from collections import defaultdict
from datetime import date, datetime

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from constants import (
    DATE_NUMBER_FORMAT,
    FULL_MILITARY_UNIT,
    FULL_UNIT_BUT,
    MONTH_NAMES_NOMINATIVE_UPPER,
    PIB_COLUMN_NAME,
    RECOGNIZED_SUBDIVISIONS,
    SCHEDULE_VOP_ALIASES,
)
from content.commander_message import (
    NOT_IN_ROSTER_ISSUE,
    SCHEDULE_SOURCE,
    compose_message,
    date_issues,
    update_message_file,
)
from content.excel_compat import save_workbook, workbook_with_sheet
from content.oblik_timesheet import normalize_name
from content.report_pib_column import PIB_ALIGNMENT, pib_column_width

# Категорії виплати, які вже підставив apply_payment_values
# (content/oblik_timesheet.py): 70 - РОП (ротний опорний пункт), 170 - ВОП
# (взводний опорний пункт), як і в STATUS_MEANINGS. Будь-яке інше значення
# (30, 100, текстовий статус) - порожня клітинка у Відомості.
_ROP_CATEGORY = 70
_VOP_CATEGORY = 170


def _month_columns(timesheet, year, month):
    """[(col_idx, номер дня), ...] - колонки-дати timesheet за year/month,
    за зростанням дня."""
    return sorted(
        (col_idx, date_value.day)
        for col_idx, date_value in timesheet.date_columns
        if date_value.year == year and date_value.month == month
    )


def _compute_day_marks(timesheet, year, month):
    """{row_idx: {день: "роп"/"воп"}} для кожної людини Табеля, навіть без
    жодного дня РОП/ВОП (порожній словник, а не відсутній ключ) - щоб
    find_schedule_mismatches однаково перебирав усіх людей. Фільтр "хоча б
    один день" застосовує лише build_statement_rows."""
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
    """Позначки "роп"/"воп" на кожен день year/month, де клітинка-дата вже має
    категорію 70/170 (apply_payment_values має відпрацювати до цього), і
    лічильники днів РОП/ВОП. Як і в реальному зразку
    resources/{SHORT_UNIT_BATTALION без пробілу}ВОП-РОП_ЛИПЕНЬ_.xlsx, у
    результат потрапляють лише люди з хоча б одним таким днем за місяць - у
    порядку рядків Табеля."""
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
    """Чи однакові обидві сторони дня: буквальний збіг або обидва значення -
    псевдоніми ВОП (SCHEDULE_VOP_ALIASES, перелік веде користувач).
    "воп"/"тмп"/"тот" - різні слова для тієї самої категорії 170, а
    обчислене (_compute_day_marks) завжди пише "воп". None з будь-якого боку
    ("не звірялось" чи "не на ВОП") обробляється ще до виклику."""
    if computed_status == schedule_status:
        return True
    return computed_status in SCHEDULE_VOP_ALIASES and schedule_status in SCHEDULE_VOP_ALIASES


def find_schedule_mismatches(timesheet, year, month, schedule_marks, schedule_people, recognized_subdivisions=RECOGNIZED_SUBDIVISIONS):
    """Звіряє обчислені позначки "роп"/"воп" (_compute_day_marks: ОБЛІК за
    рапортами, у клітинки "РВЗ" якого apply_payment_values уже поставив
    категорії з information_unit) з підтвердженою відомістю
    resources/schedule (read_schedule_marks: marks - {(normalize_name(ПІБ),
    день): "роп"/"воп"}, people - {normalize_name(ПІБ): {"pib_raw", "posada",
    "zvannya", "pidrozdil"}}).

    Порівнюються лише дні, заповнені у відомості: її подають щодня й
    заповнюють поступово, тож порожній день означає "ще не дійшли", а не
    "порожньо". Заповнений день - розбіжність, якщо обчислене значення не
    еквівалентне (_schedule_statuses_equivalent) значенню відомості, зокрема
    якщо обчислено порожньо. Незаповнені у відомості дні не порівнюються
    взагалі, навіть якщо обчислено щось непорожнє.

    Окремо - записи відомості, чий ПІБ не знайдено в Табелі (описка чи людини
    там немає), як і _missing_from_roster_records у payment_mismatch_checker:
    сторона Табеля (Посада/Звання/ПІБ) отримує _NOT_IN_ROSTER_LABEL, а сторона
    відомості - справжні дані з schedule_people, щоб було видно, хто це.

    Повертає [{"roster_pib_raw", "roster_posada", "roster_zvannya",
    "schedule_pib_raw", "schedule_posada", "schedule_zvannya", "pidrozdil",
    "dates": {date: (обчислено, відомість)}, "full_days": {date: (обчислено,
    відомість)}, "roster_values": {date: значення клітинки ОБЛІК}}] лише для
    людей і записів з хоча б однією розбіжною датою ("roster_values" - лише
    в знайдених у Табелі).
    "dates" - лише розбіжні дати: від них залежить, чи людина потрапляє у звіт
    і які колонки-дати він має. "full_days" - усі порівняні дати людини,
    зокрема збіжні: звіт показує їх зеленим, а не порожньою клітинкою, якщо
    колонка з'явилась через розбіжність іншої людини.

    "pidrozdil" - довідкове поле, яке не звіряється: ПІДРОЗДІЛ з Табеля для
    знайдених людей або дописка прикріплення в дужках з відомості для
    незнайдених. У Табелі ця колонка заповнена для всіх, а відомість дужками
    позначає лише винятки, тож порівняння давало б хибну розбіжність майже в
    кожному записі.

    Людина зовсім не потрапляє в результат, якщо її pidrozdil - непорожній
    текст поза recognized_subdivisions (за замовчуванням RECOGNIZED_SUBDIVISIONS:
    бойові роти, батареї, взводи зі спільного з generator_br_and_report_for_money
    переліку company_sections.SEQUENCE, без штабу чи управління). Дописка в
    незнайдених часто вказує на сторонній підрозділ, і відсутність такої
    людини в Табелі очікувана, а не помилка. Якщо pidrozdil немає (колонки
    ПІДРОЗДІЛ немає, клітинка порожня чи дописки немає), фільтр не діє: немає
    даних - немає підстав виключати. Знайдена людина, яку виключив фільтр,
    виключається повністю - її ПІБ однаково вважається знайденим
    (matched_normalized), тож у другому проході вона не з'явиться як
    "не знайдено в ОБЛІК.xlsx"."""
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

    col_by_day = {day: col_idx for col_idx, day in _month_columns(timesheet, year, month)}

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
            # Що саме стоїть в ОБЛІК у ці дні (напр. "Аршт" чи 30; за числом,
            # що замінило статус, - сам статус: ППД, а не 10) - звіт показує це
            # замість порожнього "—", щоб було видно причину.
            "roster_values": {
                d: timesheet.statuses_behind_numbers.get(
                    (row_idx, d), timesheet.ws.cell(row=row_idx, column=col_by_day[d.day]).value,
                )
                for d in full_days if d.day in col_by_day
            },
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
# ТО/ТМП/ТП в generator_timesheet не мають джерела даних (обчислюються лише
# РОП/ВОП), тож їхні лічильники завжди 0. Колонки лишено, бо вони є в зразку.
_COUNT_HEADERS = ["кількість\n днів на ВОП", "кількість\n днів на ТО", "кількість\n днів на ТМП", "кількість\n днів на ТП", "кількість\n днів на РОП"]
_COUNT_COLUMN_WIDTH = 10


def write_statement(records, days_in_month, output_path, year, month):
    """Будує output_path (.xlsx) за зразком реального файлу
    resources/{SHORT_UNIT_BATTALION без пробілу}ВОП-РОП_ЛИПЕНЬ_.xlsx: заголовок
    (FULL_UNIT_BUT/FULL_MILITARY_UNIT з resources/data.json, назва місяця й
    рік), колонки № з/п / Військове звання / Посада / ПІБ / дні
    1..days_in_month / 5 лічильників (ВОП, ТО, ТМП, ТП, РОП - обчислюються
    лише ВОП і РОП, решта завжди 0) і підписний блок з фіксованим текстом.
    Книга створюється заново, а не з зразка, як і write_mismatch_report: тоді
    не треба зсувати підписний блок під кількість людей.

    Повертає output_path."""
    month_name = MONTH_NAMES_NOMINATIVE_UPPER[month]

    wb, ws = workbook_with_sheet(month_name)

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

    save_workbook(wb, output_path)
    return output_path


def _write_signature_block(ws, start_row, last_col):
    """Фіксований підписний блок (ролі без реальних даних; "___" - як у
    зразку, заповнюється вручну на папері) - такий самий, як у
    resources/{SHORT_UNIT_BATTALION без пробілу}ВОП-РОП_ЛИПЕНЬ_.xlsx."""
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


# Вигляд як в error_mis_statuses.xlsx (Times New Roman 14, зелена й червона
# заливка) - користувач просив "щось на зразок error_mis_status, тільки з
# відомістю". Стилі продубльовано, а не імпортовано з
# content/payment_mismatch_checker.py, щоб ці модулі не залежали один від
# одного.
_MISMATCH_FONT = Font(name="Times New Roman", size=14)
_MISMATCH_HEADER_FONT = Font(name="Bahnschrift Light SemiCondensed", size=14, bold=True, color="FFFFFFFF")
_MISMATCH_HEADER_FILL = PatternFill(start_color="FF035C6B", end_color="FF035C6B", fill_type="solid")
_MISMATCH_MATCH_FILL = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")
_MISMATCH_MISMATCH_FILL = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
# Дата - колонка звіту (бо розбіжна для когось іншого), але для цієї людини
# відомість цього дня порожня й не звірялась: не зелений і не червоний, а
# окремий нейтральний колір, щоб клітинка не виглядала пропущеною.
_MISMATCH_NOT_COMPARED_FILL = PatternFill(start_color="FFF2F2F2", end_color="FFF2F2F2", fill_type="solid")
_MISMATCH_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)
_MISMATCH_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
# Заголовки дат повернуті на 90° (текст знизу вгору): з 31 колонкою-датою це
# вдвічі вужче за горизонтальне "ДД.ММ.РРРР".
_MISMATCH_DATE_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True, text_rotation=90)
_MISMATCH_THIN_SIDE = Side(style="thin")
_MISMATCH_BORDER = Border(left=_MISMATCH_THIN_SIDE, right=_MISMATCH_THIN_SIDE, top=_MISMATCH_THIN_SIDE, bottom=_MISMATCH_THIN_SIDE)

# Обчислене береться не з самого information_unit, а з ОБЛІК: категорія з
# файлу БЧС потрапляє лише в клітинку "РВЗ". Реальний випадок: підрозділ і
# відомість кажуть ВОП, а в ОБЛІК "Аршт" (повернення з арешту немає в
# рапортах) - колись звіт показував тут "Обчислено (information_unit)" і "—",
# і причину не було видно.
_MISMATCH_COMPUTED_LABEL = "ОБЛІК (рапорти + БЧС)"
_MISMATCH_SCHEDULE_LABEL = "resources/schedule (файл)"
_MISMATCH_EMPTY_PLACEHOLDER = "—"
_MISMATCH_LABEL_HEADERS = ["Джерело", "Підрозділ", "Посада", "Звання", "ПІБ"]
_MISMATCH_LABEL_WIDTHS = (28, 20, 30, 24, 30)
_MISMATCH_PIB_COLUMN = 5  # ширина - за найдовшим ПІБ (content/report_pib_column.py)
# Заголовок дати вертикальний, тож ширини досить для короткого значення
# ("роп"/"воп"/"—").
_MISMATCH_DATE_WIDTH = 6
_MISMATCH_HEADER_ROW_HEIGHT = 85


def _mismatch_display(value):
    return _MISMATCH_EMPTY_PLACEHOLDER if value is None else value


def _computed_display(record, date_value, computed_status):
    """Обчислена сторона розбіжного дня: "роп"/"воп", а якщо їх немає - те,
    що стоїть в ОБЛІК ("Аршт", 30...), і лише за порожньої клітинки "—"."""
    if computed_status is not None:
        return computed_status
    return _mismatch_display(record.get("roster_values", {}).get(date_value))


def _write_compared_label_cell(ws, row_a, row_b, col_idx, roster_value, schedule_value):
    """Одна порівнювана колонка (Посада/Звання/ПІБ) - так само, як колонки-дати
    (і як _write_compared_cell у payment_mismatch_checker, тут продубльовано).
    Однакові значення в Табелі й відомості - одна об'єднана зелена клітинка,
    різні - два окремі значення (по одному на рядок) з червоною заливкою, щоб
    було видно, де дані є, а де немає."""
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
    тільки з відомістю": колонки Джерело/Підрозділ/Посада/Звання/ПІБ і окрема
    колонка на кожну розбіжну дату (лише дати з records["dates"], як у
    write_mismatch_report; заголовок дати повернуто на 90°). На кожен запис -
    два рядки (_MISMATCH_COMPUTED_LABEL і _MISMATCH_SCHEDULE_LABEL у колонці
    "Джерело").

    Посада/Звання/ПІБ звіряються так само, як дати (_write_compared_label_cell):
    збіг - одна зелена клітинка, різниця - два значення червоним. Так для
    запису "Не знайдено в ОБЛІК.xlsx" видно справжні Посаду/Звання/ПІБ з
    відомості. "Підрозділ" не звіряється (лише довідкова об'єднана клітинка):
    відомість не веде підрозділ кожної людини, і порівняння майже завжди
    давало б хибну розбіжність.

    Порожніх клітинок у таблиці немає: кожна колонка-дата заповнюється для
    кожного запису за його "full_days" одним із трьох станів:
    - відомість цього дня для цієї людини порожня (колонка є через
      розбіжність іншої людини) - сіра об'єднана клітинка з "—";
    - обчислене й відомість еквівалентні (_schedule_statuses_equivalent) -
      зелена об'єднана клітинка зі значенням з відомості, як подано (без
      заміни на "воп");
    - інакше (справжня розбіжність) - два значення з червоною заливкою.

    Повертає output_path або None, якщо records порожній - як і
    write_mismatch_report, файл лише із заголовком не потрібен."""
    if not records:
        return None

    all_dates = sorted({date_value for record in records for date_value in record["dates"]})

    wb, ws = workbook_with_sheet("Розбіжності")

    date_headers = [datetime.combine(d, datetime.min.time()) for d in all_dates]
    headers = _MISMATCH_LABEL_HEADERS + date_headers
    for col_idx, value in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=value)
        cell.font = _MISMATCH_HEADER_FONT
        cell.fill = _MISMATCH_HEADER_FILL
        cell.border = _MISMATCH_BORDER
        if isinstance(value, datetime):
            cell.alignment = _MISMATCH_DATE_HEADER_ALIGNMENT
            cell.number_format = DATE_NUMBER_FORMAT
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
                # Показуємо саме те, що написано у відомості: "воп"/"тмп"/"тот"
                # еквівалентні, але це різні слова, а не завжди "воп".
                cell_a = ws.cell(row=row_a, column=col_idx, value=_mismatch_display(schedule_status))
                ws.merge_cells(start_row=row_a, start_column=col_idx, end_row=row_b, end_column=col_idx)
                cell_a.fill = _MISMATCH_MATCH_FILL
            else:
                cell_a = ws.cell(row=row_a, column=col_idx, value=_computed_display(record, date_value, computed_status))
                cell_a.fill = _MISMATCH_MISMATCH_FILL
                ws.cell(row=row_b, column=col_idx, value=_mismatch_display(schedule_status)).fill = _MISMATCH_MISMATCH_FILL

        for row in (row_a, row_b):
            ws.row_dimensions[row].height = 20
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row, column=col_idx)
                cell.font = _MISMATCH_FONT
                cell.alignment = PIB_ALIGNMENT if col_idx == _MISMATCH_PIB_COLUMN else _MISMATCH_CENTER
                cell.border = _MISMATCH_BORDER

        current_row += 2

    # Статус з ОБЛІК буває довшим за "роп"/"воп" (напр. "100_БПШП").
    date_width = max([_MISMATCH_DATE_WIDTH] + [
        len(str(ws.cell(row=row, column=col_idx).value or "")) + 2
        for row in range(2, current_row) for col_idx in date_col_by_date.values()
    ])
    for col_idx in range(1, len(headers) + 1):
        letter = get_column_letter(col_idx)
        if col_idx == _MISMATCH_PIB_COLUMN:
            width = pib_column_width(
                [record["roster_pib_raw"] for record in records] + [record["schedule_pib_raw"] for record in records]
            )
        else:
            width = _MISMATCH_LABEL_WIDTHS[col_idx - 1] if col_idx <= len(_MISMATCH_LABEL_HEADERS) else date_width
        ws.column_dimensions[letter].width = width

    save_workbook(wb, output_path)
    return output_path


def _bchs_value(schedule_status):
    """Що має стояти у файлі БЧС за відомістю: замість РОП - 70, замість ВОП
    і його псевдонімів SCHEDULE_VOP_ALIASES - 170; інше - як у відомості."""
    if schedule_status == "роп":
        return _ROP_CATEGORY
    if schedule_status in SCHEDULE_VOP_ALIASES:
        return _VOP_CATEGORY
    return schedule_status


def _same_value(a, b):
    """170 і "170" (чи 170.0) - те саме значення файлу БЧС."""
    return str(a).strip().removesuffix(".0").casefold() == str(b).strip().removesuffix(".0").casefold()


def _schedule_message_entries(records, submitted_values=None):
    """Записи повідомлення командирам (content/commander_message.py) за тими
    самими записами find_schedule_mismatches, з яких пишеться
    error_mis_statuses_Відомість.xlsx. "Як має бути" - значення з відомості
    (вона підтверджене джерело) у вигляді категорії файлу БЧС (_bchs_value).

    submitted_values - те, що подали самі підрозділи ({(normalize_name(ПІБ),
    дата): [значення, ...]}, read_payment_values). День, де файл БЧС уже
    подає саме це, командиру не пишеться: виправляти йому нічого, причина в
    ОБЛІК. Реальний випадок: підрозділ подав 170, відомість - ВОП, а в ОБЛІК
    "Аршт" без повернення в рапортах - командиру йшло "має бути (170)" поруч
    із "має бути (Аршт)"."""
    submitted_values = submitted_values or {}
    entries = []
    for record in records:
        if record["roster_pib_raw"] == _NOT_IN_ROSTER_LABEL:
            entries.append((record["pidrozdil"], record["schedule_pib_raw"], [NOT_IN_ROSTER_ISSUE]))
            continue
        normalized = normalize_name(record["roster_pib_raw"])
        expected = {}
        for date_value, (_computed, schedule_status) in record["dates"].items():
            value = _bchs_value(schedule_status)
            if not any(_same_value(value, sent) for sent in submitted_values.get((normalized, date_value), ())):
                expected[date_value] = value
        if expected:
            entries.append((record["pidrozdil"], record["roster_pib_raw"], date_issues(expected)))
    return entries


def schedule_mismatch_message(records, submitted_values=None):
    """Текст повідомлення лише за цими записами; "" - нема про що писати."""
    return compose_message(_schedule_message_entries(records, submitted_values))


def write_schedule_mismatch_message(records, output_path, submitted_values=None):
    """Додає ці записи в спільний файл повідомлення output_path
    (update_message_file); None - нема про що писати."""
    return update_message_file(output_path, SCHEDULE_SOURCE, _schedule_message_entries(records, submitted_values))
