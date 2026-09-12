import calendar
import os
import re
from datetime import datetime, timedelta
from functools import lru_cache

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment

from constants import (
    OUTPUT_DIR, PERSONEL_LIST_FILE_NAME, PERSONEL_LIST_SHEET_NAME, PERSONEL_LIST_COLUMNS_LETTERS, MONTH, YEAR,
    BR_HIGHT_UNIT, SHORT_UNIT_BRIGADE, NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY, NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK,
)
from content.document_content_search import extract_document_text
from content.money_report_helpers import normalize_name, oblik_disappearance_date, build_brs_chain_lines
from content.report_document_reader import extract_report_entries, _dates_from_period_text, SPETSKONTYNGENT_VALUE
from utils.date_utils import to_date
from utils.excel_reader import read_datafile, PERSONEL_LIST_FALLBACK_SHEET_NAMES
from utils.logging_utils import print_green, print_red
from utils.excel_writer import save_workbook_safely

# Звірка ГОТОВОГО рапорту (report_document_reader.extract_report_entries - ПЕРІОД/
# ДНІ/ПІБ, ЯК НАПИСАНО в самому рапорті) з РЕАЛЬНИМИ даними ОБЛІК.xlsx поточного
# місяця (constants.PERSONEL_LIST_FILE_NAME/SHEET_NAME/COLUMNS_LETTERS - вже
# автоматично звужені до обраного місяця, як і для звичайної генерації) - на
# відміну від табеля (generators/generate_timetable.py), тут МЕТА - саме
# ЗІСТАВЛЕННЯ рапорту з джерелом, а не побудова нового документа.
_RESULT_FILE_NAME = "Звірка рапорту з ОБЛІК.xlsx"
# Окремі колонки (а не один суцільний текст) - підтверджено користувачем: з
# результатом зручніше працювати (сортувати/фільтрувати в Excel за ПІБ чи за
# ресурсом), коли ПІБ/сам опис/джерело звірки/шлях до нього - окремі значення,
# а не одна змішана фраза. "Ресурс" - ЩО саме звірялось (назва файлу ОБЛІК.xlsx
# для _check_pib_spelling/_check_periods_and_days/_check_spetskontyngent, чи
# текст самої підстави "БР"/"ЖБД" для _check_basis_documents); "Шлях" - ДЕ
# саме (порожньо для перевірок проти ОБЛІК.xlsx - сам "Ресурс" уже шлях; шлях
# до обраної папки OUTPUT для _check_basis_documents). "Дата періоду"/"Документ" -
# ЛИШЕ для _check_basis_documents (підтверджено користувачем): конкретна дата
# ПЕРІОДУ людини, за яку не підтвердилась участь, і назва конкретного файлу,
# що перевірявся за цю дату (порожньо для решти перевірок - у них немає такого
# поняття, як "конкретний документ за конкретний день").
_RESULT_HEADERS = ["ПІБ", "Повідомлення", "Ресурс", "Шлях", "Дата періоду", "Документ"]
_RESULT_COLUMN_WIDTHS = [30, 90, 40, 55, 16, 45]


def _finding(pib, message, resource, path="", period_date="", document=""):
    return {"ПІБ": pib, "Повідомлення": message, "Ресурс": resource, "Шлях": path, "Дата періоду": period_date, "Документ": document}


# Маркер загибелі в комірці дня ОБЛІК.xlsx: реальні дані показують, що це
# буквальний текст "загибель" (не число) - "200"
# лишається підтримуваним про всяк випадок (це слово користувач вживав як усне
# позначення поняття "загибель" в обговоренні, а не як точний текст комірки).
# З дня, наступного за ОСТАННІМ днем зі статусом СПЕЦКОНТИНГЕНТ, з'являється цей
# маркер - людина вибуває зі статусу, а не просто "пропущений день".
_DEATH_VALUE = 200
_DEATH_TEXT_VALUE = "загибель"

def _blank(value):
    return value is None or value == "" or (not isinstance(value, str) and pd.isna(value))


def _is_death_marker(value):
    if value == _DEATH_VALUE or (isinstance(value, str) and value.strip() == str(_DEATH_VALUE)):
        return True
    return isinstance(value, str) and value.strip().lower() == _DEATH_TEXT_VALUE


def _oblik_rows_by_normalized_pib():
    """{normalize_name(ПІБ): рядок ОБЛІК.xlsx} - джерело істини для звірки."""
    rows = read_datafile(
        PERSONEL_LIST_FILE_NAME, PERSONEL_LIST_SHEET_NAME, PERSONEL_LIST_COLUMNS_LETTERS,
        extra_fallback_sheet_names=PERSONEL_LIST_FALLBACK_SHEET_NAMES,
    ).get("rows", [])
    return {normalize_name(row.get("ПІБ", "")): row for row in rows}


def _check_pib_spelling(entries, oblik_by_pib):
    """Для КОЖНОЇ (унікальної) людини з рапорту: чи є вона в ОБЛІК.xlsx узагалі
    (за нормалізованим ПІБ - без урахування регістру/пробілів/апострофа), і чи
    написано її ПІБ ТОЧНО так само (побуквено), як в ОБЛІК.xlsx - різне написання
    (зайвий пробіл, інша літера, розрив апострофа) не завадило б звичайній
    генерації (вона звіряє за normalize_name), але видає одруківку/помилку, яку
    варто виправити вручну в одному з двох файлів."""
    findings = []
    seen = set()
    for entry in entries:
        pib_raw = entry["ПІБ"]
        normalized = normalize_name(pib_raw)
        if normalized in seen:
            continue
        seen.add(normalized)

        oblik_row = oblik_by_pib.get(normalized)
        if oblik_row is None:
            findings.append(_finding(pib_raw, "З рапорту не знайдено в ОБЛІК.xlsx (помилка написання чи людину прибрано з ОБЛІК).", PERSONEL_LIST_FILE_NAME))
            continue

        oblik_pib = oblik_row.get("ПІБ", "")
        if oblik_pib != pib_raw:
            findings.append(_finding(pib_raw, f"ПІБ написано по-різному: рапорт \"{pib_raw}\" - ОБЛІК.xlsx \"{oblik_pib}\".", PERSONEL_LIST_FILE_NAME))
    return findings


def _check_periods_self_consistent(entries, resource):
    """Для КОЖНОГО рядка рапорту, КРІМ SPETSKONTYNGENT_VALUE (у нього ЗОВСІМ
    інша таблиця - "Дата зникнення безвісти"/"Період виплати" замість "Період
    участі"/"Кількість днів", де "ДНІ" узагалі не про кількість днів - див.
    _check_spetskontyngent, окрема логіка, підтверджено користувачем): чи
    справді "ПЕРІОД" дає стільки днів, скільки написано в "ДНІ"
    (самоузгодженість САМОГО рапорту - без звірки з жодним зовнішнім джерелом,
    підтверджено користувачем: окремий легкий інструмент, де ОБЛІК.xlsx
    узагалі не потрібен/не читається).

    resource - що показати в колонці "Ресурс" результату (сам рапорт, що
    перевіряється, - на відміну від check_report_against_oblik, тут немає
    зовнішнього джерела звірки)."""
    findings = []
    for entry in entries:
        if entry["raw_value"] == SPETSKONTYNGENT_VALUE:
            continue

        pib_raw, period_text, days_text, point = entry["ПІБ"], entry["ПЕРІОД"], entry["ДНІ"], entry["raw_value"]
        dates = _dates_from_period_text(period_text)

        try:
            written_days = int(days_text)
        except ValueError:
            findings.append(_finding(pib_raw, f"(пункт {point}): кількість днів \"{days_text}\" - не число (період \"{period_text}\").", resource))
            continue

        if written_days != len(dates):
            findings.append(_finding(
                pib_raw, f"(пункт {point}): період \"{period_text}\" дає {len(dates)} дн., а в рапорті написано {written_days}.", resource,
            ))
    return findings


def _check_periods_and_days(entries, oblik_by_pib):
    """Те саме, що _check_periods_self_consistent, ПЛЮС: чи МАЄ ОБЛІК.xlsx
    НЕПОРОЖНЄ значення на КОЖНУ дату періоду для цієї людини (рапорт
    стверджує участь того дня - джерело має це підтверджувати)."""
    findings = _check_periods_self_consistent(entries, PERSONEL_LIST_FILE_NAME)

    for entry in entries:
        if entry["raw_value"] == SPETSKONTYNGENT_VALUE:
            continue

        pib_raw, period_text, point = entry["ПІБ"], entry["ПЕРІОД"], entry["raw_value"]
        oblik_row = oblik_by_pib.get(normalize_name(pib_raw))
        if oblik_row is None:
            continue  # вже є окреме повідомлення в _check_pib_spelling

        for date in _dates_from_period_text(period_text):
            if _blank(oblik_row.get(date)):
                findings.append(_finding(
                    pib_raw, f"(пункт {point}): рапорт вказує участь {date.strftime('%d.%m.%Y')}, але в ОБЛІК.xlsx комірка за цю дату порожня.",
                    PERSONEL_LIST_FILE_NAME,
                ))
    return findings


def _month_end_or_death_date(oblik_row, month_int, year_int):
    """Останній день, до якого людина МАЄ рахуватись СПЕЦКОНТИНГЕНТОМ в обраному
    місяці (constants.MONTH/YEAR) - останній календарний день місяця, АБО день
    ПЕРЕД першим маркером загибелі (_is_death_marker) в ОБЛІК.xlsx, якщо він
    трапляється РАНІШЕ. Маркер з'являється, ПОЧИНАЮЧИ з дня, наступного за
    останнім активним статусом (підтверджено реальними даними: людина мала
    "100_СПЕЦКОНТИНГЕНТ" по 22.07 включно, а вже з
    23.07 - "загибель") - тому очікуваний кінець періоду це день ПЕРЕД тим, що
    з маркером, а не сам цей день."""
    last_day = calendar.monthrange(year_int, month_int)[1]
    for day in range(1, last_day + 1):
        date = datetime(year_int, month_int, day)
        if _is_death_marker(oblik_row.get(date)):
            return date - timedelta(days=1)
    return datetime(year_int, month_int, last_day)


def _check_spetskontyngent(entries, oblik_by_pib):
    """Окрема логіка ЛИШЕ для SPETSKONTYNGENT_VALUE ("СПЕЦКОНТИНГЕНТ" - полон/
    зниклі безвісти/заручники), підтверджено користувачем: (1) БЕЗ перевірки
    "ДНІ" (_check_periods_and_days її й не робить для цього пункту - таблиця
    інша); (2) звіряє "Дата зникнення безвісти" рапорту з відповідною датою в
    ОБЛІК.xlsx (oblik_disappearance_date); (3) день-за-днем - якщо в ОБЛІК.xlsx
    на якийсь день трапляється маркер загибелі (_is_death_marker), участь ПІСЛЯ
    цієї дати вже НЕ перевіряється; (4) сам ПЕРІОД у рапорті МАЄ закінчуватись
    РІВНО в очікувану дату (_month_end_or_death_date - кінець місяця, АБО день
    перед загибеллю, якщо вона трапляється раніше) - як КОРОТШИЙ період без
    причини (людина мала б і далі рахуватись СПЕЦКОНТИНГЕНТОМ), так і ДОВШИЙ
    (рапорт продовжує рахувати участь ПІСЛЯ загибелі) - обидва розбіжність.

    Перевірки (2) і (3) вимагають рядка ОБЛІК.xlsx і тому пропускаються, якщо
    людину в ньому не знайдено (вже є окреме повідомлення в _check_pib_spelling).
    Перевірка (4), навпаки, НЕ гейтиться знайденою людиною - вона самодостатня
    (звіряє ПЕРІОД рапорту з календарем обраного місяця), а відсутність людини
    в ОБЛІК.xlsx означає лише "немає даних про дату '200'", тобто очікуваний
    кінець - просто кінець місяця (_month_end_or_death_date з {} нічого не
    знайде і поверне саме це). Раніше (4) також ховалась за тим самим "не
    знайдено" - і саме тому користувач не побачив розбіжність періоду для
    людей, яких прибрано з поточного ОБЛІК.xlsx (напр. після визнання зниклими
    безвісти)."""
    findings = []
    entries_by_pib = {}
    for entry in entries:
        if entry["raw_value"] == SPETSKONTYNGENT_VALUE:
            entries_by_pib.setdefault(normalize_name(entry["ПІБ"]), []).append(entry)

    for normalized, person_entries in entries_by_pib.items():
        pib_raw = person_entries[0]["ПІБ"]
        oblik_row = oblik_by_pib.get(normalized)

        if oblik_row is not None:
            for entry in person_entries:
                report_date_text = entry["ДАТА_ЗНИКНЕННЯ"]
                oblik_date_raw = oblik_disappearance_date(oblik_row)
                if report_date_text:
                    try:
                        dates_match = not _blank(oblik_date_raw) and to_date(oblik_date_raw) == to_date(report_date_text)
                    except ValueError:
                        dates_match = False
                    if not dates_match:
                        findings.append(_finding(
                            pib_raw, f"(СПЕЦКОНТИНГЕНТ): дата зникнення в рапорті - \"{report_date_text}\", а в ОБЛІК.xlsx - {oblik_date_raw!r}.",
                            PERSONEL_LIST_FILE_NAME,
                        ))

        all_dates = sorted({date for entry in person_entries for date in _dates_from_period_text(entry["ПЕРІОД"])})

        if oblik_row is not None:
            for date in all_dates:
                if _is_death_marker(oblik_row.get(date)):
                    break
                if _blank(oblik_row.get(date)):
                    findings.append(_finding(
                        pib_raw, f"(СПЕЦКОНТИНГЕНТ): рапорт вказує участь {date.strftime('%d.%m.%Y')}, але в ОБЛІК.xlsx комірка за цю дату порожня.",
                        PERSONEL_LIST_FILE_NAME,
                    ))

        expected_end = _month_end_or_death_date(oblik_row or {}, int(MONTH), int(YEAR))
        month_end = datetime(int(YEAR), int(MONTH), calendar.monthrange(int(YEAR), int(MONTH))[1])
        boundary = "дати загибелі" if expected_end != month_end else "кінця місяця"
        actual_end = all_dates[-1] if all_dates else None

        if actual_end is not None and actual_end > expected_end:
            findings.append(_finding(
                pib_raw,
                f"(СПЕЦКОНТИНГЕНТ): період у рапорті триває до {actual_end.strftime('%d.%m.%Y')}, хоча має закінчуватись до {boundary} "
                f"({expected_end.strftime('%d.%m.%Y')}) - участь після цієї дати не рахується.",
                PERSONEL_LIST_FILE_NAME,
            ))
        elif actual_end is None or actual_end < expected_end:
            findings.append(_finding(
                pib_raw,
                f"(СПЕЦКОНТИНГЕНТ): період у рапорті закінчується {actual_end.strftime('%d.%m.%Y') if actual_end else '(немає жодної дати)'}, "
                f"а має тривати до {boundary} ({expected_end.strftime('%d.%m.%Y')}).",
                PERSONEL_LIST_FILE_NAME,
            ))
    return findings


# Рядок "Підстава для виплати" - НАДЗВИЧАЙНО різноманітний за форматом (вручну
# написані constants.MONEY_REPORT_GENERAL_REFERENCES*/"grounds": роздільники
# "від"/пробіл/крапка/дефіс/"·", номер написаний ДО чи ПІСЛЯ дати, номери ВИЩИХ
# штабів на кшталт "119/1/2500т/9" замість простого числа) - надійно зіставити
# САМЕ номер із назвою локального файлу неможливо, а от ДАТА завжди присутня в
# однозначному форматі "ДД.ММ.РРРР" - тож звірка з файлами тримається лише на
# ній (і типі - "БР" чи "ЖБД"), не на номері. Номер (коли вдається розпізнати)
# усе одно повертається й показується користувачу в підсумкових повідомленнях -
# ЛИШЕ для зручності читання, у самій звірці з файлами участі не бере.
_BASIS_MARKER_RE = re.compile(r"\b(БР|БН|ЖБД|ПозБД)\b", re.IGNORECASE)
_BASIS_DATE_RE = re.compile(r"\d{2}\.\d{2}\.\d{4}")
# "№2217" -> "2217"; зупиняється ПЕРЕД "від" (з можливими пробілом/комою/дефісом
# перед ним) чи на першому пробілі - інакше "від"/сама дата потрапили б у номер
# (напр. "№119/1/2500т/9-від-12.05.2026" -> "119/1/2500т/9", а не все до кінця).
_BASIS_NUMBER_RE = re.compile(r"№\s*([^\s]+?)(?=[\s,;)-]*від\b|\s|$)", re.IGNORECASE)
# Ведучі не-буквено-цифрові символи в номері - лише косметика написання (той
# самий номер може бути "№-1622" (constants.py, вручну) чи "№ 1622"/"№1622"
# (як його фактично передрукували в готовий рапорт) - обрізаються, щоб таке
# саме число не вважалось РІЗНИМ при звірці з _IGNORED_BASIS_CITATIONS нижче.
_BASIS_NUMBER_LEADING_PUNCTUATION = " -·."
# Підтверджено користувачем: перевіряються ЛИШЕ "БР"/"ЖБД" - "БН"/"ПозБД" часто
# посилаються на накази ВИЩОГО штабу (не локальні файли цього проєкту), тож не
# звіряються з файлами, але й далі рахуються МЕЖЕЮ сегмента нижче (щоб не
# "проковтнути" дату НАСТУПНОЇ підстави в тому самому багаторядковому тексті).
_BASIS_TRACKED_MARKERS = {"БР", "ЖБД"}


def _number_and_date(text):
    """(нормалізований номер чи None, дата "ДД.ММ.РРРР" чи None) - ПЕРША дата
    (_BASIS_DATE_RE) і ПЕРШИЙ номер (_BASIS_NUMBER_RE, без ведучої пунктуації -
    _BASIS_NUMBER_LEADING_PUNCTUATION) де-небудь у text. (None, None), якщо
    дати немає взагалі (номер без дати - недостатньо, щоб вважати це підставою,
    яку взагалі можна перевірити чи виключити)."""
    date_match = _BASIS_DATE_RE.search(text)
    if not date_match:
        return None, None
    number_match = _BASIS_NUMBER_RE.search(text)
    number = number_match.group(1).lstrip(_BASIS_NUMBER_LEADING_PUNCTUATION) if number_match else None
    return number, date_match.group()


def _reference_line_citation(line):
    """(тип "БР"/"ЖБД", номер, дата) чи None - якщо line (ОДИН самодостатній
    рядок підстави, напр. з constants.BR_HIGHT_UNIT чи
    content.money_report_helpers.build_brs_chain_lines) починається зі
    "БР"/"ЖБД" (_BASIS_TRACKED_MARKERS) і має дату десь у своєму тексті. None -
    інакше (БН/ПозБД, чи взагалі без дати - напр. "ЖБД (Справа №2т)")."""
    normalized = " ".join(line.split())
    marker_match = _BASIS_MARKER_RE.search(normalized)
    if marker_match is None or marker_match.group(1).upper() not in _BASIS_TRACKED_MARKERS:
        return None
    number, date = _number_and_date(normalized)
    if date is None:
        return None
    return marker_match.group(1).upper(), number, date


def _reference_citations(lines):
    """{(тип, номер, дата), ...} - _reference_line_citation КОЖНОГО рядка
    lines, пропускаючи ті, що повернули None. Окрема функція (а не сам вираз
    одразу в _IGNORED_BASIS_CITATIONS нижче) - щоб цю логіку можна було
    перевірити на СИНТЕТИЧНОМУ переліку рядків, незалежно від ПОТОЧНОГО
    (щомісяця змінюваного) вмісту реальних джерел."""
    return {citation for citation in (_reference_line_citation(line) for line in lines) if citation is not None}


def _all_reference_lines(references_by_group):
    """[рядок, ...] - з УСІХ "lines" УСІХ записів УСІХ груп словника форми
    constants.BR_HIGHT_UNIT ({назва групи: [{"start", "end", "lines": [...]}, ...]})."""
    return [
        line
        for group in references_by_group.values()
        for reference in group
        for line in reference.get("lines", [])
    ]


def _dates_from_date_keyed_dict(date_keyed_dict):
    """[datetime, ...] - з РЯДКОВИХ ключів "ДД.ММ.РРРР" словника форми
    constants.NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK/_EVERY_DAY - для передачі в
    content.money_report_helpers.build_brs_chain_lines, якій для роботи
    потрібен лише список дат, а не сирий словник напряму. Ключ, який не
    розбирається як дата - пропускається (не мало б траплятись у реальних
    даних, але без падіння, як і решта опціональних перевірок проєкту)."""
    dates = []
    for date_str in date_keyed_dict:
        try:
            dates.append(datetime.strptime(date_str, "%d.%m.%Y"))
        except ValueError:
            continue
    return dates


def _every_day_brg_lines():
    """{"БР {SHORT_UNIT_BRIGADE} №{N} від {дата}", ...} - РІВНО той самий
    формат, що й content.money_report_helpers._build_brs_chain_lines_from_selected_folder
    будує з поля 'брг' constants.NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY. Підтверджено
    користувачем: 'брг' цього словника - номер ВИЩОГО (бригадного) штабу за
    той самий день, не локальний файл цього проєкту, тож НЕ перевіряється - на
    відміну від поля 'бат' ЦЬОГО Ж словника (справжній щоденний БР батальйону -
    локальний файл ЩОДЕННОЇ, і далі перевіряється як завжди)."""
    return {
        f"БР {SHORT_UNIT_BRIGADE} №{entry['брг']} від {date_str}"
        for date_str, entry in NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY.items()
        if entry.get("брг")
    }


# Підтверджено користувачем: джерела підстав, які НІКОЛИ не перевіряються з
# файлами (навіть попри "БР"/"ЖБД" і дату в самому тексті) - жодне з них НЕ є
# локальним файлом ЩОДЕННОЇ/Витяга з ЖБД цього проєкту:
# 1. constants.BR_HIGHT_UNIT - підстави ВИЩОГО штабу ("9 АК", командир бригади тощо);
# 2. constants.NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK - ЦІЛКОМ (усі поля 'бат'/
#    'посилання_бат'/'посилання_брг' - тижневий ЗАВДАННЯ, не щоденна ЩОДЕННА) -
#    ЗА НОМЕРОМ, окремо нижче (_IGNORED_WEEKLY_BR_NUMBERS), не тут;
# 3. constants.NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY - ЛИШЕ поле 'брг' (_every_day_brg_lines,
#    'бат' цього самого словника - навпаки, і далі перевіряється).
# "general_br_bat" (constants.NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY) тут НЕ згадано:
# це поле взагалі не потрапляє в "Підстава для виплати" рапорту (лише в текст
# заголовка самого документа ЩОДЕННОЇ - generators/generate_documents_br_every_day.py),
# тож _basis_citations його й так ніколи не побачить.
#
# Порівняння - за (тип, номер, дата) (_reference_citations), а НЕ за дослівним
# текстом рядка: РЕАЛЬНІ дані показали, що той самий запис BR_HIGHT_UNIT
# з'являється в ГОТОВОМУ рапорті записаним ЗОВСІМ інакше, ніж він написаний у
# ПОТОЧНОМУ constants.py (напр. рапорт: "БР 9 АК №119/1/2500т/9 від
# 12.05.2026;", а constants.py сьогодні: "БР-9·АК №119/1/2500т/9-від-12.05.2026") -
# рапорт зберігає текст ТАКИМ, яким BR_HIGHT_UNIT був НА МОМЕНТ його генерації,
# а сам BR_HIGHT_UNIT редагується (переформатовується) далі - дослівне
# порівняння тоді ніколи б не збіглося, хоча йдеться про ТУ САМУ підставу.
_IGNORED_BASIS_CITATIONS = (
    _reference_citations(_all_reference_lines(BR_HIGHT_UNIT))
    | _reference_citations(_every_day_brg_lines())
)


def _ignored_weekly_br_numbers():
    """{номер, ...} - усі 'бат'-номери, які build_brs_chain_lines будує з
    constants.NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK - ЛИШЕ номер, БЕЗ дати
    (на відміну від _IGNORED_BASIS_CITATIONS вище). Підтверджено користувачем
    на реальному випадку: тижневий БР лишається "чинним" кілька днів ПІСЛЯ
    дня, яким він виданий (доти, доки не з'явиться наступний тижневий) - тому
    в ГОТОВОМУ рапорті той самий номер часто цитується з ІНШОЮ датою (датою
    участі людини того тижня, а не датою видання, що є єдиним ключем цього
    словника) - зіставлення за (тип, номер, ДАТА) тоді ніколи не збіглося б,
    хоча йдеться про ТОЙ САМИЙ тижневий номер."""
    return {
        citation[1]
        for citation in _reference_citations(build_brs_chain_lines(_dates_from_date_keyed_dict(NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK)))
        if citation[1] is not None
    }


_IGNORED_WEEKLY_BR_NUMBERS = _ignored_weekly_br_numbers()

# "№117 ЩОДЕННА 13.07.2026.docx"/"№87 ЗАВДАННЯ 01.07.2026.pdf" (можливо з
# префіксом "Витяг " - content.document_content_search._DAILY_OR_TASK_FILENAME_RE
# має ту саму форму, але тут не використовується напряму: там ще й номер/тип
# документа окремими групами, тут потрібна лише дата).
_BR_DOCUMENT_FILENAME_RE = re.compile(r"№\d+\s*(?:ЩОДЕННА|ЗАВДАННЯ)\s+(\d{2}\.\d{2}\.\d{4})\.(?:docx|doc|pdf)$", re.IGNORECASE)

# "Витяг з ЖБД для БР - №117 13.07.2026 ЩОДЕННА.docx"/"...ЩОТИЖНЕВА.docx"
# (formatting.docx_utils.save_combat_log_extract_war) - дата тут ("13.07.2026") -
# ВЛАСНА календарна дата ЦЬОГО КОНКРЕТНОГО витяга (день, що в ньому описаний), а
# НЕ дата "від DATE" із самої підстави рапорту: та дата - дата РЕЄСТРАЦІЇ запису
# constants.MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR (ОДНА й та сама для ВСЬОГО
# періоду дії цього запису, тобто спільна для БАГАТЬОХ різних витягів з РІЗНИМИ
# власними датами) - підтверджено користувачем, тому для ЖБД дата НЕ бере
# участі в пошуку файлу взагалі (лише номер справи, _log_war_files_for_number
# нижче, і власна дата файлу - для звірки з ПЕРІОДОМ людини).
_LOG_WAR_EXTRACT_DATE_RE = re.compile(r"№\d+\s+(\d{2}\.\d{2}\.\d{4})(?:\s*(?:ЩОДЕННА|ЩОТИЖНЕВА))?\.(?:docx|doc)$", re.IGNORECASE)


@lru_cache(maxsize=None)
def _cached_document_text(path):
    """extract_document_text(path), кешовано за шляхом - той самий файл може
    читатись і НА НОМЕР СПРАВИ (_log_war_files_for_number), і НА ПІБ
    (_check_basis_documents) кілька разів за один запуск звірки."""
    return extract_document_text(path)


def _br_files_for_date(folder, date_str):
    """Файли folder (рекурсивно, os.walk) типу ЩОДЕННА/ЗАВДАННЯ
    (_BR_DOCUMENT_FILENAME_RE), чия ВЛАСНА дата (з назви файлу) - РІВНО date_str."""
    matches = []
    for dirpath, _dirnames, filenames in os.walk(folder):
        for filename in sorted(filenames):
            match = _BR_DOCUMENT_FILENAME_RE.search(filename)
            if match and match.group(1) == date_str:
                matches.append(os.path.join(dirpath, filename))
    return matches


def _log_war_files_for_number(folder, number):
    """[(шлях, datetime власної дати файлу), ...] - файли folder (рекурсивно,
    os.walk), чия НАЗВА відповідає формату "Витяг з ЖБД..." (_LOG_WAR_EXTRACT_DATE_RE)
    і чий ЗМІСТ (заголовок - formatting.docx_utils.generate_combat_log_extract_title)
    містить number (напр. "359дск/7") - цей номер справи дослівно повторюється
    в заголовку КОЖНОГО витяга за ввесь період дії відповідного запису
    constants.MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR, тому саме він (а не дата
    з тексту підстави рапорту) - надійний спосіб знайти ВСІ відповідні файли.

    number порожній (не вдалось розпізнати в самій підставі) - порожній
    результат: немає за чим шукати."""
    if not number:
        return []
    matches = []
    for dirpath, _dirnames, filenames in os.walk(folder):
        for filename in sorted(filenames):
            date_match = _LOG_WAR_EXTRACT_DATE_RE.search(filename)
            if not date_match:
                continue
            path = os.path.join(dirpath, filename)
            if number in _cached_document_text(path):
                matches.append((path, datetime.strptime(date_match.group(1), "%d.%m.%Y")))
    return matches


def _basis_citations(basis_text):
    """[(тип "БР"/"ЖБД", номер чи None, дата "ДД.ММ.РРРР"), ...] - для КОЖНОГО
    згаданого в basis_text посилання на "БР"/"ЖБД" (_BASIS_TRACKED_MARKERS), що
    має дату десь у своєму тексті. basis_text спершу зводиться до одного рядка
    (пробіли/переноси рядків - до одного пробілу): один запис підстави інколи
    сам переносить дату на НАСТУПНИЙ рядок формулювання (напр.
    constants.MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR: "ЖБД ... №359дск/N\\n
    від {дата}") - без цього зведення дата й маркер опинились би на РІЗНИХ
    "рядках" і жодного зіставлення не сталось би.

    Далі текст ділиться на сегменти за КОЖНОЮ згадкою будь-якого маркера
    (БР/БН/ЖБД/ПозБД - усі чотири, а не лише "БР"/"ЖБД") - сегмент триває від
    самого маркера до ПОЧАТКУ наступного (чи кінця тексту): це не дає одному
    сегменту "поглинути" дату вже НАСТУПНОЇ підстави, коли кілька написані
    поспіль в одному багаторядковому тексті. У готовому сегменті шукається
    ПЕРША дата і ПЕРШИЙ номер (_number_and_date) - номер повертається головно
    для показу користувачу, але береться участь і в звірці з
    _IGNORED_BASIS_CITATIONS нижче.

    (тип, номер, дата) сегмента, що збігається з одним із _IGNORED_BASIS_CITATIONS
    (constants.BR_HIGHT_UNIT, поле 'брг' NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY) -
    жодне з них не відповідає локальному файлу цього проєкту - пропускається
    взагалі, підтверджено користувачем. Порівняння - за (тип, номер, дата), НЕ
    за дослівним текстом сегмента: те саме реальне посилання ЧАСТО записане в
    готовому рапорті ЗОВСІМ іншими словами/розділовими знаками, ніж воно ж
    виглядає в ПОТОЧНОМУ (уже відредагованому користувачем відтоді) constants.py.

    "БР", чий номер - один із _IGNORED_WEEKLY_BR_NUMBERS
    (constants.NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK) - пропускається так само,
    але ЛИШЕ за номером (дата НЕ звіряється взагалі): тижневий БР лишається
    чинним кілька днів ПІСЛЯ дня видання, тож у рапорті цитується з ІНШОЮ
    датою (датою участі, а не видання), підтверджено користувачем на
    реальному випадку."""
    if not basis_text:
        return []
    normalized = " ".join(basis_text.split())
    markers = list(_BASIS_MARKER_RE.finditer(normalized))

    citations = []
    for index, marker in enumerate(markers):
        marker_type = marker.group(1).upper()
        if marker_type not in _BASIS_TRACKED_MARKERS:
            continue
        segment_end = markers[index + 1].start() if index + 1 < len(markers) else len(normalized)
        segment = normalized[marker.start():segment_end].strip()
        number, date = _number_and_date(segment)
        if date is None:
            continue
        if marker_type == "БР" and number in _IGNORED_WEEKLY_BR_NUMBERS:
            continue
        citation = (marker_type, number, date)
        if citation not in _IGNORED_BASIS_CITATIONS:
            citations.append(citation)
    return citations


def _files_by_date_for_citation(folder, doc_type, number, date_str, person_dates):
    """{datetime дня: [шлях, ...]} - файли folder, релевантні для ОДНІЄЇ
    підстави "БР"/"ЖБД", згруповані за ВЛАСНОЮ датою КОЖНОГО файлу.

    "БР" - РІВНО одна дата (сама дата з підстави рапорту, розпарсена з
    date_str) - _br_files_for_date, як і раніше (файл(и), якщо є, під цією
    ЄДИНОЮ датою).

    "ЖБД" - дата "від DATE" у підставі рапорту - дата РЕЄСТРАЦІЇ запису
    constants.MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR (спільна для БАГАТЬОХ
    витягів різних днів), а не дата конкретного файлу чи дня участі людини -
    тому файли шукаються за НОМЕРОМ справи в заголовку
    (_log_war_files_for_number), групуючись за ВЛАСНОЮ датою КОЖНОГО з них,
    лише серед днів, які людина РЕАЛЬНО заявляє у своєму "ПЕРІОД" (person_dates,
    _dates_from_period_text) - перевіряється саме "чи є людина за той період,
    який вказаний в рапорті", а не за буквальною (спільною для багатьох днів)
    датою з тексту підстави."""
    if doc_type == "ЖБД":
        files_by_date = {}
        for path, file_date in _log_war_files_for_number(folder, number):
            if file_date in person_dates:
                files_by_date.setdefault(file_date, []).append(path)
        return files_by_date

    citation_date = datetime.strptime(date_str, "%d.%m.%Y")
    br_files = _br_files_for_date(folder, date_str)
    return {citation_date: br_files} if br_files else {}


def _check_basis_documents(entries, folder):
    """Для КОЖНОЇ унікальної (людина, тип, номер, дата) підстави "БР"/"ЖБД"
    (_basis_citations) - шукає відповідні файли в folder (_files_by_date_for_citation,
    згруповані за ВЛАСНОЮ датою КОЖНОГО файлу - для "БР" це рівно одна дата, для
    "ЖБД" - усі дні періоду людини, підтверджені окремими витягами) і, ДЛЯ
    КОЖНОЇ такої дати, перевіряє, чи згадане повне ПІБ людини в тексті ХОЧ
    ОДНОГО з файлів САМЕ цієї дати (той самий підхід порівняння, що й
    content.document_content_search.find_person_document_references).

    Підтверджено користувачем: розбіжність тепер ЗА КОЖНУ ОКРЕМУ дату періоду
    окремо (а не одна узагальнена на всю підставу) - "Дата періоду"/"Документ"
    у результаті показують САМЕ ту дату й САМЕ той файл (чи "не знайдено
    документ", якщо на цю дату взагалі немає відповідного файлу), де ПІБ не
    підтвердився - людині, яка згадана в ОДНИХ витягах періоду, але не в ІНШИХ,
    тепер видно РІВНО ЯКІ дні саме бракує підтвердження.

    Підстава без жодної згадки "БР"/"ЖБД" з датою - не перевіряється взагалі
    (файл не шукається). Та сама (людина, тип, номер, дата) підстава, згадана в
    кількох рядках рапорту (напр. кілька пунктів тієї самої людини) -
    перевіряється ОДИН раз."""
    findings = []
    seen = set()
    for entry in entries:
        pib_raw = entry["ПІБ"]
        normalized_pib = normalize_name(pib_raw)
        person_dates = set(_dates_from_period_text(entry.get("ПЕРІОД")))

        for doc_type, number, date_str in _basis_citations(entry.get("ПІДСТАВА")):
            key = (normalized_pib, doc_type, number, date_str)
            if key in seen:
                continue
            seen.add(key)

            label = f"{doc_type} №{number} від {date_str}" if number else f"{doc_type} від {date_str}"
            files_by_date = _files_by_date_for_citation(folder, doc_type, number, date_str, person_dates)

            if doc_type == "ЖБД":
                if not files_by_date:
                    # Жодного файлу з цим номером справи серед ПЕРІОДУ людини
                    # взагалі - ОДНЕ загальне "не знайдено" (а не помилкове
                    # "не знайдено" на КОЖЕН день періоду - решта днів періоду
                    # цілком законно можуть належати ІНШІЙ підставі "ЖБД" з
                    # іншим номером справи, підтверджено користувачем).
                    findings.append(_finding(pib_raw, f"Не знайдено документ \"{label}\" у {folder}.", label, folder))
                    continue
                dates_to_check = sorted(files_by_date)
            else:
                dates_to_check = [datetime.strptime(date_str, "%d.%m.%Y")]

            for date in dates_to_check:
                period_date = date.strftime("%d.%m.%Y")
                files_for_date = files_by_date.get(date, [])

                if not files_for_date:
                    findings.append(_finding(
                        pib_raw, f"Не знайдено документ \"{label}\" за {period_date} у {folder}.",
                        label, folder, period_date=period_date,
                    ))
                    continue

                mentioned = any(normalized_pib in " ".join(_cached_document_text(path).split()).upper() for path in files_for_date)
                if not mentioned:
                    document_names = "; ".join(os.path.basename(path) for path in files_for_date)
                    findings.append(_finding(
                        pib_raw, f"Документ \"{label}\" знайдено, але ПІБ за {period_date} там не згадано.",
                        label, folder, period_date=period_date, document=document_names,
                    ))
    return findings


def _write_findings_workbook(findings, output_path, message_when_clean):
    """Спільна для check_report_against_oblik/check_report_periods логіка
    запису результату (_RESULT_HEADERS - ПІБ/Повідомлення/Ресурс/Шлях/Дата
    періоду/Документ, як окремі колонки, підтверджено користувачем:
    зручніше сортувати/фільтрувати, ніж один суцільний текст) - розбіжностей
    немає - файл усе одно створюється (з текстом про це), щоб користувач
    бачив, що перевірка відбулась. Друкує кожну розбіжність у консоль
    (червоним) чи message_when_clean (зеленим), якщо розбіжностей немає."""
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Звірка"
    for col_idx, header in enumerate(_RESULT_HEADERS, start=1):
        ws.cell(row=1, column=col_idx, value=header)

    if not findings:
        ws.cell(row=2, column=1, value=message_when_clean)
    else:
        for row_idx, finding in enumerate(findings, start=2):
            for col_idx, header in enumerate(_RESULT_HEADERS, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=finding[header])
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    for col_idx, width in enumerate(_RESULT_COLUMN_WIDTHS, start=1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width
    save_workbook_safely(wb, output_path)

    if findings:
        for finding in findings:
            print_red(f"⚠ {finding['ПІБ']}: {finding['Повідомлення']}")
    else:
        print_green(message_when_clean)

    return output_path


def check_report_against_oblik(report_path, output_path=None, basis_folder=None):
    """Звіряє report_path (готовий рапорт .docx) з ОБЛІК.xlsx поточного обраного
    місяця (constants.MONTH/YEAR) за трьома ознаками: правильність написання ПІБ,
    самоузгодженість "ПЕРІОД"/"ДНІ" в самому рапорті, і чи справді ОБЛІК.xlsx
    підтверджує КОЖЕН день, який рапорт заявляє як день участі.

    basis_folder - опційно, шлях до папки OUTPUT (обраної користувачем) - якщо
    задано, додатково перевіряє КОЖНЕ посилання "БР"/"ЖБД" у колонці "Підстава
    для виплати" (_check_basis_documents): чи існує в цій папці відповідний
    документ (за датою) і чи згадане в ньому повне ПІБ людини. Не задано (за
    замовчуванням) - ця перевірка пропускається взагалі (як і раніше).

    Зберігає список розбіжностей у output_path (за замовчуванням -
    OUTPUT_DIR/_RESULT_FILE_NAME) і повертає цей шлях."""
    entries = extract_report_entries(report_path)
    oblik_by_pib = _oblik_rows_by_normalized_pib()

    findings = (
        _check_pib_spelling(entries, oblik_by_pib)
        + _check_periods_and_days(entries, oblik_by_pib)
        + _check_spetskontyngent(entries, oblik_by_pib)
    )
    if basis_folder:
        findings += _check_basis_documents(entries, basis_folder)

    return _write_findings_workbook(
        findings, output_path or os.path.join(OUTPUT_DIR, _RESULT_FILE_NAME),
        message_when_clean="Розбіжностей між рапортом і ОБЛІК.xlsx не знайдено.",
    )


# Окремий, ЛЕГКИЙ інструмент (RUN_MODE_CHECK_REPORT_PERIODS) - підтверджено
# користувачем: не всі хочуть звіряти з ОБЛІК.xlsx (де вже є своя, ширша
# перевірка "Звірка рапорту з ОБЛІК.xlsx (періоди/дні/ПІБ)") - лише швидко
# обрати рапорт і перевірити, чи сам він внутрішньо узгоджений (період і
# написана кількість днів збігаються), без будь-якого зовнішнього джерела.
_PERIODS_RESULT_FILE_NAME = "Перевірка рапорту (періоди-дні).xlsx"


def check_report_periods(report_path, output_path=None):
    """Перевіряє report_path (готовий рапорт .docx) САМ НА СЕБЕ (без
    ОБЛІК.xlsx чи будь-якого іншого зовнішнього джерела): чи справді
    "ПЕРІОД" кожного рядка дає стільки днів, скільки написано в "ДНІ"
    (_check_periods_self_consistent). СПЕЦКОНТИНГЕНТ (SPETSKONTYNGENT_VALUE)
    пропускається - у нього зовсім інша таблиця (див. docstring
    _check_periods_self_consistent).

    Зберігає список розбіжностей у output_path (за замовчуванням -
    OUTPUT_DIR/_PERIODS_RESULT_FILE_NAME) і повертає цей шлях."""
    entries = extract_report_entries(report_path)
    findings = _check_periods_self_consistent(entries, os.path.basename(report_path))

    return _write_findings_workbook(
        findings, output_path or os.path.join(OUTPUT_DIR, _PERIODS_RESULT_FILE_NAME),
        message_when_clean="Невідповідностей між періодом і кількістю днів не знайдено.",
    )
