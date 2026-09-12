import re
from datetime import datetime, timedelta

import pandas as pd

from constants import (
    NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK,
    NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY,
    SHORT_UNIT_BATTALION,
    SHORT_UNIT_BRIGADE,
    MONEY_REPORT_CATEGORIES,
    HIGHER_COMMANDER_TITLE,
    BASIS_REQUIRED_POINTS,
    GENERAL_PIDSTAVY_COLUMN_NAMES,
    NOT_PAID_IGNORED_WHEN_CONFIRMED_BY,
)

# Підрозділ "МП" (медичний пункт) завжди означає медика, незалежно від колонки СТАТУС.
# Це - єдиний десь захардкоджений виняток; сама назва "МЕДИК" далі ніде окремо не
# дублюється - і resolve_status_category, і сам ОБЛІК.xlsx звіряються з ключами
# constants.MONEY_REPORT_CATEGORIES напряму.
_MP_SUBDIVISION_CATEGORY = "МЕДИК"
from utils.date_utils import to_date, date_to_str
from utils.logging_utils import print_red

MONTH_NAMES_NOMINATIVE_UPPER = {
    1: "СІЧЕНЬ", 2: "ЛЮТИЙ", 3: "БЕРЕЗЕНЬ", 4: "КВІТЕНЬ", 5: "ТРАВЕНЬ", 6: "ЧЕРВЕНЬ",
    7: "ЛИПЕНЬ", 8: "СЕРПЕНЬ", 9: "ВЕРЕСЕНЬ", 10: "ЖОВТЕНЬ", 11: "ЛИСТОПАД", 12: "ГРУДЕНЬ",
}

MONTH_NAMES_GENITIVE_LOWER = {
    1: "січня", 2: "лютого", 3: "березня", 4: "квітня", 5: "травня", 6: "червня",
    7: "липня", 8: "серпня", 9: "вересня", 10: "жовтня", 11: "листопада", 12: "грудня",
}


def month_nominative_upper(month_int):
    return MONTH_NAMES_NOMINATIVE_UPPER[month_int]


def month_genitive_lower(month_int):
    return MONTH_NAMES_GENITIVE_LOWER[month_int]


def _clean_pib_spacing(name):
    """Виправляє розрив апострофа (В' Ярош -> В'Ярош) і зайві пробіли в ПІБ - спільний
    перший крок normalize_name (тут) і content.br_helpers.convert_to_short_name (де
    далі йде ще й розбиття на прізвище/ім'я/по батькові, чого тут не потрібно)."""
    cleaned = re.sub(r"(В|в)'[\s]+([а-яА-ЯІіЇїЄєҐґ])", r"\1'\2", name.strip())
    return " ".join(cleaned.split())


def normalize_name(name):
    """Нормалізує ПІБ для зіставлення в межах ОБЛІК.xlsx (пробіли, апостроф, регістр)."""
    if not isinstance(name, str):
        return ""
    return _clean_pib_spacing(name).upper()


def _category_name_matches(category_name, normalized_value):
    """Чи збігається normalized_value (ВЖЕ у ВЕРХНЬОМУ регістрі, обрізаний) з
    назвою категорії category_name. Назва категорії може містити КІЛЬКА варіантів
    написання, розділених "|" (напр. "Задув.|задув|задут" - одна категорія, кілька
    написань того самого слова, які реально трапляються в ОБЛІК.xlsx, замість
    окремої категорії на кожен варіант) - збіг з БУДЬ-ЯКИМ із варіантів рахується
    як збіг із самою категорією. Назва без "|" - звичайне порівняння, як і раніше."""
    return normalized_value in {alt.strip().upper() for alt in category_name.split("|")}


def resolve_status_category(raw_status, pidrozdil=None, categories=None):
    """Повертає назву категорії (ключ у constants.MONEY_REPORT_CATEGORIES[30 чи 100],
    напр. "МЕДИК"/"РТГр"/"БД(СЗ)"/"ЗВРез") для сирого значення колонки СТАТУС, або
    None, якщо жодна категорія не відповідає. Категорії тут НЕ дублюються окремими
    константами - назва є лише один раз, у самому MONEY_REPORT_CATEGORIES.

    categories - те саме джерело категорій, що й в resolve_category_config/
    resolve_reportable_categories (за замовчуванням - constants.MONEY_REPORT_CATEGORIES;
    звіт на командира/ТВО передає constants.COMMANDER_MONEY_REPORT_CATEGORIES). Раніше
    завжди звірялось лише з глобальним MONEY_REPORT_CATEGORIES, незалежно від того, який
    саме набір категорій передав викликач build_category_rows - категорія, що існує лише
    в COMMANDER_MONEY_REPORT_CATEGORIES (чи в переданому тестом власному наборі), просто
    ніколи не розпізнавалась.

    Порівняння - БЕЗ урахування регістру (значення статусу завжди звіряється у
    ВЕРХНЬОМУ регістрі, щоб "ртгр"/"Ртгр"/"РТГр" з файлу вважались однаковим
    статусом), тож порівнювати напряму з такими ключами, як "РТГр"/"ЗВРез" (де є
    рядкові літери), звичайним "==" не можна - вони ніколи не збіглися б (див.
    _category_name_matches - також підтримує кілька варіантів написання через "|").
    Повертається КАНОНІЧНЕ написання ключа з MONEY_REPORT_CATEGORIES (цілий ключ,
    разом з усіма "|"-варіантами, якщо їх кілька), а не те, як його ввели у файлі."""
    if str(pidrozdil or "").strip().upper() == "МП":
        return _MP_SUBDIVISION_CATEGORY

    normalized_status = str(raw_status or "").strip().upper()
    if not normalized_status:
        return None

    if categories is None:
        categories = MONEY_REPORT_CATEGORIES

    for point_categories in categories.values():
        for category_name in point_categories:
            if category_name != "general" and _category_name_matches(category_name, normalized_status):
                return category_name

    return None


def get_tvo_commander_periods(rows_with_tvo_data, position_title, period_start, period_end):
    """Повертає список {"pib", "start", "end"} для КОЖНОГО запису ОБЛІК.xlsx (аркуш "ТВО"),
    де хтось значиться ТВО на посаду position_title (типово - HIGHER_COMMANDER_TITLE) з
    періодом дії (Start/End), що перетинається з [period_start; period_end] - "start"/"end"
    кожного запису вже обрізані до меж цього перетину (date, не datetime).

    generate_report_for_commander_money.py використовує САМЕ періоди (а не лише ПІБ) -
    кожен ТВО має рахуватись лише за дні своєї фактичної заміни, а не за весь місяць."""
    period_start = to_date(period_start)
    period_end = to_date(period_end)

    periods = []
    for row in rows_with_tvo_data:
        if row.get("ПОСАДА") != position_title or not row.get("ТВО"):
            continue
        try:
            row_start = to_date(row["Start"])
            row_end = to_date(row["End"])
        except (KeyError, ValueError):
            continue
        if row_start <= period_end and row_end >= period_start:
            periods.append({
                "pib": normalize_name(row.get("ПІБ", "")),
                "start": max(row_start, period_start),
                "end": min(row_end, period_end),
            })
    return periods


def get_tvo_commander_pibs(rows_with_tvo_data, position_title, period_start, period_end):
    """Повертає нормалізовані ПІБ усіх, хто в ОБЛІК.xlsx (аркуш "ТВО") значиться ТВО на
    посаду position_title (типово - HIGHER_COMMANDER_TITLE) з періодом дії (Start/End), що
    перетинається з [period_start; period_end]. Такі люди НЕ потрапляють у жоден пункт
    ГОЛОВНОГО рапорту на додаткову винагороду - для тимчасово виконуючих обов'язки
    командира є окремий рапорт (generators/generate_report_for_commander_money.py)."""
    return {
        period["pib"]
        for period in get_tvo_commander_periods(rows_with_tvo_data, position_title, period_start, period_end)
    }


# Якщо ХОЧ В ОДНІЙ комірці дати (не в колонці СТАТУС!) за місяць стоїть одне з
# цих значень - людина взагалі не потрапляє в рапорт, незалежно від значень у
# решти її комірок дат. "СЗЧ" тут свідомо НЕМАЄ - це звичайна названа категорія
# пункту "NOT_PAID" (MONEY_REPORT_CATEGORIES), рахується за конкретний день, а
# не виключає людину з рапорту цілком (див. _exclude_by_excluded_day_value нижче).
EXCLUDED_DAY_VALUES = {"ПРВД"}


def _row_has_excluded_day_value(row, date_columns):
    """Чи є хоч ОДИН день у місяці, де сама комірка ОБЛІК.xlsx (не СТАТУС людини,
    а значення в колонці конкретної дати) дорівнює одному зі значень
    EXCLUDED_DAY_VALUES ("ПРВД") - незалежно від регістру."""
    for col in date_columns:
        raw_value = row.get(col)
        if raw_value is None or raw_value == "" or pd.isna(raw_value):
            continue
        if str(raw_value).strip().upper() in EXCLUDED_DAY_VALUES:
            return True
    return False


def _exclude_by_excluded_day_value(rows_with_data, date_columns):
    """Прибирає з рапорту в/сл, у кого ХОЧ ЗА ОДИН день місяця в комірці
    ОБЛІК.xlsx стоїть "ПРВД" - людина взагалі не потрапляє в рапорт (не лише за
    ЦЕЙ конкретний день), незалежно від того, які значення стоять у решти її
    комірок дат (навіть якщо там, наприклад, 100 чи 30 за інші дні). "СЗЧ" сюди
    НЕ входить - це звичайна названа категорія пункту "NOT_PAID"
    (constants.MONEY_REPORT_CATEGORIES) і рахується day-by-day, як і решта
    категорій: сам день СЗЧ потрапляє в пункт "Не виплачувати...", а ІНШІ дні
    місяця людини й далі рахуються за своїми категоріями як завжди - на відміну
    від "ПРВД" (переведення в інший підрозділ/частину), що й далі виключає
    людину з рапорту ЦІЛКОМ. "СЗЧ, що почався ЦЬОГО місяця" - окреме правило,
    _exclude_by_szch_started_this_month нижче."""
    return [row for row in rows_with_data if not _row_has_excluded_day_value(row, date_columns)]


_SZCH_VALUE = "СЗЧ"
# Нюанс, підтверджений користувачем: людина, чий СЗЧ триває з ПОПЕРЕДНЬОГО
# місяця (перенесено - див. _row_started_szch_this_month) і "повернулась" цього
# місяця, ВСЕ ОДНО НЕ отримує виплату, якщо статус, у який вона повернулась -
# САМЕ "ЗВРез" (порівнюється з СИРИМ текстом комірки, а не через
# resolve_status_category/MONEY_REPORT_CATEGORIES - "ЗВРез" наразі НЕ
# зареєстрована як категорія MONEY_REPORT_CATEGORIES, тож звичайне визначення
# категорії дня повернуло б None; це правило має спрацьовувати незалежно від
# того, чи (і коли) користувач сам додасть "ЗВРез" туди).
_SZCH_RETURN_EXCLUDED_VALUES = {"ЗВРЕЗ"}


def _row_started_szch_this_month(row, date_columns):
    """Визначає, чи людина "скоїла злочин ЦЬОГО місяця" (пішла в СЗЧ ПРОТЯГОМ
    поточного місяця) - підтверджено користувачем, на прикладах:

    1. 01-10 - будь-який статус (НЕ СЗЧ), 11 - СЗЧ (без повернення до кінця
       місяця) -> True (виключити).
    2. 01-10 - будь-який статус, 11 - СЗЧ, 17 - повернення (будь-який звичайний
       статус) -> ВСЕ ОДНО True (виключити) - раз СЗЧ ПОЧАВСЯ цього місяця,
       подальше повернення в ТОМУ Ж місяці нічого не змінює: "скоїв злочин
       цього місяця" - місяць уже "зіпсований" цілком.
    3. 01-10 - СЗЧ (з ПОПЕРЕДНЬОГО місяця - перший день місяця ВЖЕ був СЗЧ),
       11 - звичайний статус (10/30/100/70/170) -> False (НЕ виключати) - це
       "повернення", а не "втеча цього місяця".

    Тобто вирішує НЕ сам факт наявності "СЗЧ" десь у місяці, а те, чи ПЕРШИЙ
    КАЛЕНДАРНИЙ день місяця (date_columns[0] - буквально, а НЕ перший
    НЕПОРОЖНІЙ день, як було раніше) вже був "СЗЧ" - якщо так, СЗЧ перенесено
    з попереднього місяця (можливе повернення); якщо ні (перший день - будь-що
    ІНШЕ, включно з порожньою коміркою), а "СЗЧ" з'явився ПІЗНІШЕ - це відхід
    САМЕ цього місяця.

    РЕГРЕСІЯ (виявлено користувачем на реальному прикладі): раніше тут
    перевірявся ПЕРШИЙ НЕПОРОЖНІЙ день (порожні комірки просто пропускались) -
    для людини, чиї комірки ПОРОЖНІ на початку місяця,
    а "СЗЧ" з'являється ПІЗНІШЕ (напр. з 03-го числа, комірки 1-2-го - зовсім
    без значення, а не "звичайний статус"), це помилково читалось як "перший
    ДЕНЬ уже був СЗЧ" (сценарій 3 - перенесено з попереднього місяця), хоча
    порожня комірка НЕ доводить жодного статусу "до" СЗЧ - людина просто
    мовчки НЕ потрапляла під це правило (мала б, як і решта прикладів 1/2), і
    надалі, всупереч решті таких самих випадків, отримувала звичайну обробку
    "по днях".

    Нюанс "ЗВРез" (_SZCH_RETURN_EXCLUDED_VALUES) перевіряється тут же: у
    сценарії 3 (перший ДЕНЬ - буквально СЗЧ) - якщо ПЕРШИЙ день ПІСЛЯ
    СЗЧ-періоду (перший НЕ-СЗЧ день серед непорожніх) - "ЗВРез", повернення
    однаково НЕ рахується (виключити, True), навіть попри те, що це технічно
    не "втеча цього місяця" - "нюанс", підтверджений користувачем окремо від
    основного правила."""
    if not date_columns:
        return False

    values_in_order = []
    for col in date_columns:
        raw_value = row.get(col)
        if raw_value is None or raw_value == "" or pd.isna(raw_value):
            continue
        values_in_order.append(str(raw_value).strip().upper())

    if not values_in_order:
        return False

    first_day_raw_value = row.get(date_columns[0])
    first_day_is_szch = (
        first_day_raw_value is not None and first_day_raw_value != "" and not pd.isna(first_day_raw_value)
        and str(first_day_raw_value).strip().upper() == _SZCH_VALUE
    )

    if first_day_is_szch:
        first_value_after_szch = next((v for v in values_in_order if v != _SZCH_VALUE), None)
        return first_value_after_szch in _SZCH_RETURN_EXCLUDED_VALUES

    return _SZCH_VALUE in values_in_order


def _exclude_by_szch_started_this_month(rows_with_data, date_columns):
    """Прибирає з БР документів (process_generate_br.py - людину, що пішла в
    СЗЧ ЦЬОГО місяця, не можна залучати до бойових розпоряджень) в/сл, хто
    пішов у СЗЧ ПРОТЯГОМ поточного місяця (_row_started_szch_this_month) -
    людина взагалі не потрапляє в БР за цей місяць, незалежно від решти її
    днів (навіть якщо вона встигла повернутись до звичайного статусу до кінця
    місяця) - підтверджено користувачем: "якщо за поточний місяць він скоїв
    злочин [пішов в СЗЧ] то він не отримує кошти, а якщо повернувся з СЗЧ
    [СЗЧ триває з попереднього місяця] тоді отримує кошти" (з нюансом ЗВРез).

    Рапорт на додаткову винагороду (generate_report_for_get_money.py/
    generate_report_for_commander_money.py) БІЛЬШЕ не використовує цю функцію
    (раніше - через user_input.py, СПІЛЬНУ точку для обох генераторів) - там
    "не отримує кошти" ТЕПЕР означає лише виключення зі ЗВИЧАЙНИХ пунктів
    (30/70/100/170/10), а не зникнення людини з рапорту ЦІЛКОМ: сам пункт
    "NOT_PAID" ("Не виплачувати...") існує САМЕ для документування, кому не
    платять - людина мала й далі з'являтись ТАМ за свої дні СЗЧ, а не ставати
    невидимою навіть у переліку тих, кому не платять - підтверджено
    користувачем на реальному прикладі (людина з десятками днів СЗЧ, що
    почався цього місяця, взагалі не з'являлась у рапорті). Див.
    _blank_non_szch_days_for_szch_started_this_month нижче - саме вона тепер
    виконує цю роль для рапорту на додаткову винагороду."""
    return [row for row in rows_with_data if not _row_started_szch_this_month(row, date_columns)]


def _blank_non_szch_days_for_szch_started_this_month(rows_with_data, date_columns):
    """Аналог _exclude_by_szch_started_this_month ДЛЯ РАПОРТУ на додаткову
    винагороду (а не для БР) - НЕ виключає рядок людини ЦІЛКОМ, а лише порожнить
    (до None) усі її дні ПОЗА буквальним "СЗЧ" - так людина НЕ отримує оплату за
    жоден звичайний пункт (30/70/100/170/10, навіть за дні ДО чи ПІСЛЯ СЗЧ
    того самого місяця - "місяць зіпсований цілком", той самий принцип, що й
    раніше), АЛЕ Й ДАЛІ з'являється в "NOT_PAID" за свої дні СЗЧ - підтверджено
    користувачем: сенс пункту "не виплачувати" саме в тому, щоб задокументувати,
    кому саме не платять, а не приховати людину навіть звідти.

    Застосовується ПІСЛЯ _exclude_by_excluded_day_value (ПРВД) у виклику -
    людина з ОБОМА "СЗЧ, що почався цього місяця" і "ПРВД" (переведення в
    інший підрозділ/частину) уже виключена звідти цілком і сюди не доходить -
    ПРВД має пріоритет (підтверджено користувачем на реальному прикладі:
    людина повернулась з СЗЧ, а тоді була переведена - НЕ з'являється в
    "NOT_PAID", на відміну від решти таких самих випадків)."""
    result = []
    for row in rows_with_data:
        if not _row_started_szch_this_month(row, date_columns):
            result.append(row)
            continue
        patched = dict(row)
        for col in date_columns:
            raw_value = patched.get(col)
            if raw_value is None or raw_value == "" or pd.isna(raw_value):
                continue
            if str(raw_value).strip().upper() != _SZCH_VALUE:
                patched[col] = None
        result.append(patched)
    return result


def _row_ignorable_not_paid_statuses(row, date_columns):
    """Підмножина ключів NOT_PAID_IGNORED_WHEN_CONFIRMED_BY (constants.py, У
    ВЕРХНЬОМУ РЕГІСТРІ - порівняння регістронезалежне з обох боків, бо сам
    NOT_PAID_IGNORED_WHEN_CONFIRMED_BY може містити мішаний регістр, як і
    решта статусів проєкту, напр. "Задув."/"Повер"), чиї "підтверджуючі"
    статуси дійсно трапляються десь серед date_columns цієї людини - САМЕ ЦІ
    статуси (а не всі можливі ключі словника) мають бути проігноровані ДЛЯ
    НЕЇ. Порядок появи в датах НЕ має значення - лише сам факт присутності
    ОБОХ значень (ігнорованого і хоч одного підтверджуючого) десь протягом
    місяця (реальний випадок, підтверджений користувачем - "РОЗП" не завжди
    лишається ОСТАННІМ хронологічно статусом, докстрінг
    NOT_PAID_IGNORED_WHEN_CONFIRMED_BY)."""
    values = set()
    for col in date_columns:
        raw_value = row.get(col)
        if raw_value is None or raw_value == "" or pd.isna(raw_value):
            continue
        values.add(str(raw_value).strip().upper())

    result = set()
    for ignored_status, confirming_statuses in NOT_PAID_IGNORED_WHEN_CONFIRMED_BY.items():
        normalized_ignored = ignored_status.strip().upper()
        normalized_confirming = {status.strip().upper() for status in confirming_statuses}
        if normalized_ignored in values and values & normalized_confirming:
            result.add(normalized_ignored)
    return result


def _blank_not_paid_days_confirmed_elsewhere(rows_with_data, date_columns):
    """Для КОЖНОЇ людини - замінює значення її комірок, що дорівнюють
    буквально ОДНОМУ з _row_ignorable_not_paid_statuses (НЕ всі дні людини й
    НЕ дні з "підтверджуючим" статусом) на None, ЩЕ ДО побудови категорій -
    "NOT_PAID" (build_category_rows) просто не бачить цих днів узагалі, тож
    для НИХ не з'являється ні рядок, ні попередження про відсутню підставу
    (generate_report_for_get_money._log_missing_legal_basis). Решта днів
    людини (зокрема самі "підтверджуючі" дні - вони й так зазвичай не
    належать НІЯКІЙ категорії - і будь-які ІНШІ, звичайні категорії)
    лишаються БЕЗ ЗМІН: ігнорується лише сама заміна конкретних статусів на
    None, а не людина цілком (на відміну від _exclude_by_szch_started_this_month
    вище, що прибирає рядок людини з рапорту ПОВНІСТЮ)."""
    result = []
    for row in rows_with_data:
        ignorable_statuses = _row_ignorable_not_paid_statuses(row, date_columns)
        if not ignorable_statuses:
            result.append(row)
            continue
        patched = dict(row)
        for col in date_columns:
            raw_value = patched.get(col)
            if raw_value is not None and not pd.isna(raw_value) and str(raw_value).strip().upper() in ignorable_statuses:
                patched[col] = None
        result.append(patched)
    return result


# Назви колонки ОБЛІК.xlsx з датою внесення в статус СПЕЦКОНТИНГЕНТ - обидва
# написання (те, що вже є у файлі, і те, як користувач описав його словами)
# підтримуються, той самий підхід, що й "ПІДСТАВА"/"ПІДСТАВИ" нижче. Спільна з
# checker_accounting.report_checker (звірка рапорту з ОБЛІК.xlsx) - визначена
# тут, а не там, бо build_category_rows ("Очк.БЗ" нижче) теж її потребує, а
# report_checker імпортує З цього модуля (не навпаки).
_DISAPPEARANCE_DATE_COLUMN_NAMES = ("ДАТА ЗНИКНЕННЯ", "ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ")


def oblik_disappearance_date(person_row):
    """Значення колонки "ДАТА ЗНИКНЕННЯ"/"ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ" (обидва
    написання) з рядка ОБЛІК.xlsx, чи None, якщо жодної з колонок немає взагалі."""
    for column_name in _DISAPPEARANCE_DATE_COLUMN_NAMES:
        if column_name in person_row:
            return person_row[column_name]
    return None


def _parsed_disappearance_date(person_row):
    """date() зі значення oblik_disappearance_date, чи None - колонки немає
    взагалі, значення порожнє, чи воно в форматі, якого to_date не розпізнає.
    ЄДИНЕ джерело "чи є в людини дата зникнення" - і для _disappearance_accrual_dates
    (дати участі "Очк.БЗ") нижче, і для build_category_rows (текст колонки
    "Дата зникнення" таблиці "100_СПЕЦКОНТИНГЕНТ") - обидва мають вважати
    "немає дати" рівно однаково, інакше одна людина могла б отримати рядок в
    ОДНІЙ з двох, але не в іншій, без жодної видимої причини."""
    raw_date = oblik_disappearance_date(person_row)
    if raw_date is None or (not isinstance(raw_date, str) and pd.isna(raw_date)):
        return None
    try:
        return to_date(raw_date)
    except ValueError:
        return None


def _disappearance_date_text(person_row):
    """Текст (ДД.ММ.РРРР) для колонки "Дата зникнення" таблиці
    "100_СПЕЦКОНТИНГЕНТ" - "", якщо _parsed_disappearance_date не знайшла
    дати (exclude_rows_without_disappearance_date нижче виключає такий рядок
    з рапорту цілком, з попередженням у термінал - див. її docstring)."""
    parsed = _parsed_disappearance_date(person_row)
    return parsed.strftime("%d.%m.%Y") if parsed else ""


def build_pidstavy_extra_grounds_by_person(rows_with_data, column_names=GENERAL_PIDSTAVY_COLUMN_NAMES):
    """{нормалізований ПІБ: [текст колонки підстави]} для КОЖНОЇ людини, у кого
    ХОЧ ОДНА з column_names є в ОБЛІК.xlsx і непорожня для неї (перша непорожня
    за порядком column_names - перевіряється РІВНО одна колонка на людину, а не
    об'єднуються всі) - передається як extra_grounds_by_person в
    build_category_rows, тож ця підстава завжди додається до тексту "Підстава
    для виплати"/"для не виплати" людини, поруч з усім іншим (БР, grounds,
    general) - саме так, як і для решти джерел extra_lines у build_legal_basis_text.

    column_names=GENERAL_PIDSTAVY_COLUMN_NAMES (constants.py) за замовчуванням -
    "ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ" (загальна, для БУДЬ-ЯКОГО пункту/категорії) з
    "ПІДСТАВИ"/"ПІДСТАВА" як запасні написання. Викликачі для BASIS_REQUIRED_POINTS
    (generate_report_for_get_money._build_categories, content.report_changes.
    _basis_appeared_add_rows) передають СВОЮ, ВЛАСНУ колонку за
    constants.BASIS_REQUIRED_POINT_COLUMN_NAMES[point] - підтверджено
    користувачем: "100_СПЕЦКОНТИНГЕНТ"/"100_ВПБП"/"100_БПШП" раніше помилково
    ділили ОДНУ спільну колонку "ПІДСТАВИ", тепер - кожен свою.

    Немає жодної з column_names у файлі, чи всі порожні для когось - людина
    просто відсутня в результаті (тобто до її підстави нічого не додається - не
    помилка й не порожній рядок серед інших рядків підстави)."""
    result = {}
    for row in rows_with_data:
        raw_value = None
        for column_name in column_names:
            candidate = row.get(column_name)
            if candidate is not None and not pd.isna(candidate):
                raw_value = candidate
                break
        if raw_value is None:
            continue
        text = str(raw_value).strip()
        if not text:
            continue
        result[normalize_name(row.get("ПІБ", ""))] = [text]
    return result


# BASIS_REQUIRED_POINTS - тепер у самому constants.py (не тут) - саме там користувач
# керує НАБОРОМ пунктів (додає/прибирає, без жодних змін коду), а тут (і в
# content.report_changes/generate_report_for_get_money.py, куди він реекспортується
# нижче через імпорт) лише ВИКОРИСТОВУЄТЬСЯ.


def _commander_and_tvo_pibs(rows_with_data, rows_with_tvo_data, date_columns, narrator_row):
    """{normalize_name(ПІБ), ...} - РІВНО ті, хто фігурує в рапорті КБ/ТВО
    (generate_report_for_commander_money.py._collect_commander_rows) цього
    місяця - спільне обчислення для _exclude_commander_and_tvo (звичайний
    рапорт - без НИХ) і _keep_only_commander_and_tvo (рапорт КБ/ТВО - ЛИШЕ
    ВОНИ, для розділу "Прошу внести зміни..." саме цього рапорту).

    narrator_row - хто ПІДПИСУЄ рапорт КБ/ТВО (find_higher_commander на
    "сьогодні", рахує викликач) - підтверджено користувачем: якщо підписує
    штатний командир (без позначки ТВО) - рапорт КБ/ТВО НЕ містить жодного
    ТВО (_collect_commander_rows), тож жоден ТВО й не виключається зі
    звичайного рапорту - лишається в ньому як звичайний підлеглий (навіть
    якщо реально був ТВО частину місяця, але зараз, "сьогодні", уже не
    підписує). Лише коли підписує ТВО - виключаються І він (за ввесь
    звітний період - людина, що взагалі БУЛА ТВО хоч частину місяця, не має
    лишитись частково видимою в обох рапортах одразу), І штатний командир."""
    regular_pibs = {
        normalize_name(row.get("ПІБ", "")) for row in rows_with_data if row.get("ПОСАДА") == HIGHER_COMMANDER_TITLE
    }
    if not narrator_row.get("ТВО", False):
        return regular_pibs
    tvo_commander_pibs = get_tvo_commander_pibs(rows_with_tvo_data, HIGHER_COMMANDER_TITLE, date_columns[0], date_columns[-1])
    return tvo_commander_pibs | regular_pibs


def _exclude_commander_and_tvo(rows_with_data, rows_with_tvo_data, date_columns, narrator_row):
    """Прибирає з рапорту РІВНО тих, хто фігурує в рапорті КБ/ТВО цього місяця
    (_commander_and_tvo_pibs, узгоджено з narrator_row - хто його підписує) -
    для них є окремий рапорт від першої особи (generate_report_for_commander_money.py)."""
    excluded_pibs = _commander_and_tvo_pibs(rows_with_data, rows_with_tvo_data, date_columns, narrator_row)
    return [row for row in rows_with_data if normalize_name(row.get("ПІБ", "")) not in excluded_pibs]


def _keep_only_commander_and_tvo(rows_with_data, rows_with_tvo_data, date_columns, narrator_row):
    """Обернене до _exclude_commander_and_tvo - лишає ЛИШЕ тих, хто фігурує в
    рапорті КБ/ТВО цього місяця (_commander_and_tvo_pibs). Для розділу "Прошу
    внести зміни..." рапорту КБ/ТВО (generate_report_for_commander_money.py) -
    підтверджено користувачем: КБ/ТВО не фігурують у секціях змін ГОЛОВНОГО рапорту
    (той пишеться від третьої особи про підлеглих), але якщо в них самих є зміни
    за попередній місяць - ці зміни мають потрапити в ЇХНІЙ власний рапорт."""
    included_pibs = _commander_and_tvo_pibs(rows_with_data, rows_with_tvo_data, date_columns, narrator_row)
    return [row for row in rows_with_data if normalize_name(row.get("ПІБ", "")) in included_pibs]


def get_date_columns(column_names_personel):
    return sorted(col for col in column_names_personel if isinstance(col, datetime))


def _resolve_category_point_and_name(raw_text):
    """Знаходить (точку, КАНОНІЧНУ назву категорії з MONEY_REPORT_CATEGORIES) для
    довільного тексту БЕЗ урахування регістру (напр. "ртгр" чи "РТГР" теж знайдуть
    "РТГр"; "задув" чи "задут" знайдуть "Задув.|задув|задут" - див.
    _category_name_matches) - або (None, None), якщо такої категорії немає в
    жодній точці."""
    normalized = str(raw_text).strip().upper()
    for point, point_categories in MONEY_REPORT_CATEGORIES.items():
        for category_name in point_categories:
            if category_name != "general" and _category_name_matches(category_name, normalized):
                return point, category_name
    return None, None


def resolve_day_value_and_category(raw_cell_value, base_status_category):
    """Визначає (значення дня, категорія дня) для ОДНІЄЇ комірки дати з ОБЛІК.xlsx.

    Звичайний випадок - число (30 чи 100, тобто один із ключів MONEY_REPORT_CATEGORIES):
    значення дня - воно саме, а категорія дня - базовий статус людини (з колонки
    СТАТУС, resolve_status_category) - працює так само, як і раніше.

    Якщо ж замість числа в комірці стоїть ТЕКСТ, що збігається (без урахування
    регістру) з назвою якоїсь категорії в MONEY_REPORT_CATEGORIES (напр. "РТГр") -
    цей ОДИН день примусово належить саме цій категорії (незалежно від колонки
    СТАТУС, яка описує решту місяця), а значення дня визначається автоматично -
    точка (30 чи 100), під якою ця категорія описана. Це дозволяє позначити окремі
    дні винятком без окремої таблиці періодів - так само, як дні відпустки вже
    позначають текстом (ВП/ВД/ВЛК) замість числа.

    Інакше (відпустка, порожньо, нерозпізнаний текст) - (None, None): день не
    рахується в жодній категорії."""
    if raw_cell_value in MONEY_REPORT_CATEGORIES:
        return raw_cell_value, base_status_category

    if isinstance(raw_cell_value, str):
        point, category_name = _resolve_category_point_and_name(raw_cell_value)
        if point is not None:
            return point, category_name

    return None, None


def build_period_text_and_days(dates):
    """З відсортованого списку datetime повертає (текст періоду, кількість днів),
    групуючи послідовні дні в діапазони "ДД.ММ.РРРР-ДД.ММ.РРРР", розділені "; "."""
    if not dates:
        return "", 0

    ranges = []
    start = prev = dates[0]
    for current in dates[1:]:
        if (current - prev).days == 1:
            prev = current
            continue
        ranges.append((start, prev))
        start = prev = current
    ranges.append((start, prev))

    period_text = "; ".join(
        f"{start.strftime('%d.%m.%Y')}-{end.strftime('%d.%m.%Y')}" for start, end in ranges
    )
    return period_text, len(dates)


def build_day_value_periods(row, date_columns):
    """Для ОДНОГО рядка ОБЛІК.xlsx групує дні за ТОЧНИМ значенням комірки - в
    {значення: (текст періоду, кількість днів)} - але лише те, що САМЕ ПО СОБІ
    НЕ розпізнається MONEY_REPORT_CATEGORIES (resolve_day_value_and_category):
    коди відпустки/відрядження (ВП/ВД/ВЛК/ПРВД тощо), нерозпізнаний текст. Числа-точки
    (зараз 30/100) і назви категорій (ЖИТТЄДІЯЛЬНІСТЬ/РТГр/БД(СЗ)/... - зараз) уже
    й так потрапляють у сам рапорт, тож тут не показуються - похідне виключно від
    MONEY_REPORT_CATEGORIES, тож нові точки (напр. 70/170) чи категорії, додані туди
    пізніше, автоматично перестануть тут з'являтися без зміни цієї функції.

    Суто діагностика (для логування) - жодної фільтрації за статусом чи участю тут
    немає, на відміну від build_category_rows. Порожні комірки (None, "", NaN - пусті
    клітинки Excel читаються як float('nan'), не None) пропускаються."""
    dates_by_value = {}
    for col in date_columns:
        raw_value = row.get(col)
        if raw_value is None or raw_value == "" or pd.isna(raw_value):
            continue
        recognized_value, _ = resolve_day_value_and_category(raw_value, None)
        if recognized_value is not None:
            continue
        dates_by_value.setdefault(raw_value, []).append(col)

    return {
        value: build_period_text_and_days(sorted(dates))
        for value, dates in dates_by_value.items()
    }


def _lines_for_period(references, period_start, period_end):
    """Повертає рядки з `references` (список {"start", "end", "lines", ...} - як
    constants.MONEY_REPORT_GENERAL_REFERENCES чи "grounds" категорії), чий діапазон
    дії перетинається з [period_start; period_end]."""
    # to_date() повертає date, а дати участі (колонки ОБЛІК.xlsx) - datetime,
    # тож зводимо обидва боки до date перед порівнянням.
    period_start = to_date(period_start)
    period_end = to_date(period_end)

    lines = []
    for ref in references:
        ref_start = to_date(ref["start"])
        ref_end = to_date(ref["end"]) if ref.get("end") else period_end
        if ref_start <= period_end and ref_end >= period_start:
            lines.extend(ref.get("lines", []))
    return lines


# "посилання_бат" зберігається як готовий текст "NNN від ДД.ММ.РРРР" - той самий
# батальйонний наказ, що для ІНШОЇ дати вже міг потрапити в 'бат' цього ж словника
# (наказ від дня X часто згадується як "посилання" в записі пізнішого дня Y).
# Витягуємо сам НОМЕР, щоб дедуплікувати його РАЗОМ з 'бат' (за номером, так само,
# як і дедуплікація в межах самого 'бат' - той самий номер може стояти в кількох
# датах поспіль, це один багатоденний наказ, а не кілька) - інакше той самий
# наказ друкується двічі різним форматуванням (напр. "№ 87 від 01.07.2026" і,
# окремо, "№87 від 01.07.2026").
_DOC_NUM_RE = re.compile(r"^\s*(\d+)\s*від\s*\d{2}\.\d{2}\.\d{4}\s*$")


def build_brs_chain_lines(dates):
    """'бат' + 'посилання_бат' + 'посилання_брг' з constants.NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK,
    у цьому пріоритеті, для набору дат участі (дедуплікація за номером наказу -
    'бат' і 'посилання_бат' можуть називати ОДИН і той самий наказ; хронологічний
    порядок появи). Поля 'бз_бат'/'бз_брг' (constants.NUMBER_OF_DOCUMENTS_BRS_SAVE)
    свідомо не використовуються - це номери щоденних розпоряджень з безпеки
    застосування військ, не підстава для виплати."""
    ref_bat_num_lines, ref_bat_lines, ref_brg_lines = [], [], []
    seen_bat_numbers, seen_brg = set(), set()

    for date in dates:
        date_str = date_to_str(date)
        entry = NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK.get(date_str, {})

        bat_num = entry.get("бат")
        if bat_num and bat_num not in seen_bat_numbers:
            seen_bat_numbers.add(bat_num)
            ref_bat_num_lines.append(f"БР {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №{bat_num} від {date_str}")

        posylannia_bat = entry.get("посилання_бат")
        if posylannia_bat:
            match = _DOC_NUM_RE.match(posylannia_bat)
            num = match.group(1) if match else posylannia_bat
            if num not in seen_bat_numbers:
                seen_bat_numbers.add(num)
                ref_bat_lines.append(f"БР {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №{posylannia_bat}")

        posylannia_brg = entry.get("посилання_брг")
        if posylannia_brg and posylannia_brg not in seen_brg:
            seen_brg.add(posylannia_brg)
            ref_brg_lines.append(f"БР {SHORT_UNIT_BRIGADE} №{posylannia_brg}")

    return ref_bat_num_lines + ref_bat_lines + ref_brg_lines


def _build_brs_chain_lines_from_selected_folder(dates):
    """Аналог build_brs_chain_lines, але з constants.NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY
    (щоденні номери БАТ/БРГ - заповнюються через "Зчитати обрану папку з документами
    (для номерів БАТ / БЗ)?" при генерації, а не тижневі NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK) -
    для категорій з "use_brs_from_selected_folder": True (напр. МЕДИК пункту 100: підстава -
    щоденні бойові розпорядження за КОЖЕН робочий день, а не один тижневий наказ на весь період).

    Якщо папку з документами не зчитували цього запуску (відповідь "Ні" на питання
    вище) - записи NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY лишаються порожніми ('бат'/'брг' -
    '' чи відсутні), і ця функція просто поверне порожній список - build_legal_basis_text
    тоді дасть порожню підставу, а _log_missing_legal_basis (generate_report_for_get_money.py)
    повідомить про це окремим, зрозумілішим повідомленням."""
    ref_bat_lines, ref_brg_lines = [], []
    seen_bat_numbers, seen_brg_numbers = set(), set()

    for date in dates:
        date_str = date_to_str(date)
        entry = NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY.get(date_str, {})

        bat_num = entry.get("бат")
        if bat_num and bat_num not in seen_bat_numbers:
            seen_bat_numbers.add(bat_num)
            ref_bat_lines.append(f"БР {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №{bat_num} від {date_str}")

        brg_num = entry.get("брг")
        if brg_num and brg_num not in seen_brg_numbers:
            seen_brg_numbers.add(brg_num)
            ref_brg_lines.append(f"БР {SHORT_UNIT_BRIGADE} №{brg_num} від {date_str}")

    return ref_bat_lines + ref_brg_lines


def resolve_category_config(point, category_name, categories=None):
    """Будує готовий category_config (як очікує build_legal_basis_text/build_category_rows:
    {"grounds", "use_brs", "use_brs_from_selected_folder", "general", "required_basis"}) для
    категорії category_name у точці point (30 чи 100) з categories (за замовчуванням -
    constants.MONEY_REPORT_CATEGORIES, для головного рапорту; звіт на командира/ТВО
    передає constants.COMMANDER_MONEY_REPORT_CATEGORIES, аби мати НЕЗАЛЕЖНІ від головного
    рапорту підстави) - підставляючи в "general" СПІЛЬНИЙ (для обох точок) список цієї ж
    categories, за винятком записів, чиї "id" перелічені в "exclude_general" цієї
    категорії ("all" - виключити геть усі загальні підстави).

    "required_basis": False (напр. у пункті 10к - "виконання обов'язків військової
    служби", без прив'язки до конкретного БР/наказу) означає, що для цієї категорії
    порожня "Підстава для виплати" - ОЧІКУВАНА, а не помилка: _log_missing_legal_basis
    не попереджає про такі рядки. За замовчуванням (якщо категорія не вказала явно) -
    True, тобто підстава очікується заповненою."""
    if categories is None:
        categories = MONEY_REPORT_CATEGORIES
    point_config = categories[point]
    category = point_config[category_name]
    general_list = point_config.get("general", [])
    exclude = category.get("exclude_general", [])

    if exclude == "all":
        resolved_general = []
    else:
        excluded_ids = set(exclude)
        resolved_general = [entry for entry in general_list if entry.get("id") not in excluded_ids]

    return {
        "grounds": category.get("grounds", []),
        "use_brs": category.get("use_brs", False),
        "use_brs_from_selected_folder": category.get("use_brs_from_selected_folder", False),
        "general": resolved_general,
        "required_basis": category.get("required_basis", True),
    }


def _merge_preserving_file_order(rows_with_data, *row_lists):
    """Об'єднує кілька списків рядків (кожен - результат ОКРЕМОГО виклику
    build_category_rows для однієї категорії, вже впорядкований за файлом ОБЛІК.xlsx
    у межах ЦІЄЇ категорії) в один спільний список, зберігаючи порядок ОБЛІК.xlsx
    для ВСІХ рядків разом - а не групуючи по категоріях одну за одною. Якщо одна й та
    сама людина потрапила кількома рядками (різні категорії різних днів - напр. була
    ЗВРЕЗ, потім БД(СЗ)), стабільне сортування лишає її рядки в тому порядку, в якому
    вони вже прийшли (без додаткового перевпорядкування між собою)."""
    file_index = {normalize_name(row.get("ПІБ", "")): i for i, row in enumerate(rows_with_data)}
    combined = [row for rows in row_lists for row in rows]
    combined.sort(key=lambda row: file_index.get(normalize_name(row["ПІБ"]), len(rows_with_data)))
    return combined


# Точки з поведінкою, яку НЕМОЖЛИВО вивести із самої структури categories (це не
# дані, а бізнес-правило) - винятки задаються тут явно, в коді. Додавання БУДЬ-ЯКОЇ
# ІНШОЇ, звичайної точки в MONEY_REPORT_CATEGORIES/COMMANDER_MONEY_REPORT_CATEGORIES
# (плюс відповідний ключ "{point}_SECTION"/"COMMANDER_{point}_SECTION" у STATIK JSON)
# підхоплюється _build_categories (і content/report_changes.py) АВТОМАТИЧНО, без
# жодних змін коду.
# "70_РТГр"/"170_РТГр" - НАЗВАНІ категорії ВСЕРЕДИНІ пунктів 70/170 (форсуються
# буквальним текстом у комірці дня - resolve_day_value_and_category), на відміну
# від "100_РТГр" - ОКРЕМОГО пункту (не категорії пункту 100). Через це "зайвий"
# рядок пункту 100 для 70_РТГр/170_РТГр-в/сл (_COMBINED_TARGET_CELL_VALUES[100]
# нижче - хто отримав 70/170, автоматично отримує рядок і в 100-й) досі падав у
# ЗАГАЛЬНИЙ catch-all "БД(СЗ)"/"МЕДИК" пункту 100, а не в секцію "100_РТГр" разом
# з рештою РТГр (де опиняється лише той, чия комірка буквально "100_РТГр") -
# підтверджено користувачем на реальному прикладі: в/сл, чиї дні - буквально
# "170_РТГр", має бути РАЗОМ з рештою (чиї дні - буквально "100_РТГр") у секції
# "100_РТГр", а не в загальній "100". Бізнес-правило (не структура даних), тож - явний виняток
# тут, як і _COMBINED_TARGET_CELL_VALUES нижче: (1) catch-all пункту 100 більше
# НЕ приймає ці дві категорії (resolve_reportable_categories), (2)
# rtgr_cross_tier_rows будує для них ОКРЕМИЙ додатковий рядок САМЕ в "100_РТГр".
_RTGR_CROSS_TIER_SOURCE_CATEGORIES = {"70_РТГр", "170_РТГр"}
_RTGR_CROSS_TIER_TARGET_POINT = "100_РТГр"
_RTGR_CROSS_TIER_CELL_VALUES = {70, 170}


_COMBINED_TARGET_CELL_VALUES = {
    # Пункт 100 рахує ще й дні 70/170 (додатковий рядок з власною, 100-ю підставою) -
    # хто отримав компенсацію за 70 чи 170, автоматично отримує ще й рядок у 100-й.
    100: {70, 100, 170},
}


# Пріоритет пунктів рапорту - НЕ за зростанням номера, а в цьому фіксованому
# порядку (бізнес-правило, не структура даних - тож задається явно, тут).
_POINT_PRIORITY_ORDER = [30, 100, 70, 170, 10]


def _ordered_points(categories, priority_order=_POINT_PRIORITY_ORDER):
    """Порядок пунктів рапорту: спершу ВІДОМІ пункти з priority_order (за замовчуванням
    _POINT_PRIORITY_ORDER - 30, 100, 70, 170, 10), у цьому фіксованому пріоритеті, -
    лише ті з них, що фактично є в categories (відсутній пункт просто пропускається,
    наступний за пріоритетом підіймається вище - перелік коротшає, а не ламається);
    далі - БУДЬ-ЯКІ ІНШІ пункти categories (напр. нова точка "50", додана лише в
    MONEY_REPORT_CATEGORIES без жодних змін коду), за зростанням номера. Це та сама
    гарантія "додати нову точку без жодних змін коду", що й раніше (_build_categories) -
    просто порядок відомих точок тепер підпорядкований priority_order, а не номеру.

    priority_order - окремо переданий список дозволяє розділу "Прошу внести зміни..."
    (content/report_changes.py) мати СВІЙ пріоритет нумерації додатків (100->50->30->...),
    відмінний від пріоритету звичайних пунктів рапорту (_POINT_PRIORITY_ORDER).

    Номер пункту зазвичай число (30/100/...), але може бути й довільним рядком
    (напр. "100_ШП" - окрема категорія, не прив'язана до числового пункту виплати) -
    сортування "інших" пунктів рахує ЧИСЛА одним блоком (за зростанням, як і
    раніше), а РЯДКИ - іншим, окремим блоком після них (за алфавітом), щоб не
    порівнювати int з str напряму (сама по собі TypeError)."""
    known = [point for point in priority_order if point in categories]
    other = sorted(
        (point for point in categories if point not in priority_order),
        key=lambda point: (0, point) if isinstance(point, int) else (1, str(point)),
    )
    return known + other


def resolve_reportable_categories(point, categories):
    """Повертає впорядкований список специфікацій категорій ОДНОГО пункту point, які
    МАЮТЬ потрапити в рапорт (тобто "include_to_report" - за замовчуванням True) - для
    кожної: {"category_name", "status_filter", "exclude_status"} - готові аргументи для
    build_category_rows (разом з category_config від resolve_category_config).

    Категорії пункту визначаються ДИНАМІЧНО за фактичними ключами categories[point], а
    не захардкодженими назвами: кожна названа категорія отримує status_filter=її назва;
    РІВНО ОДНА категорія пункту (позначена "default": True у своєму конфігу - напр.
    "ЖИТТЄДІЯЛЬНІСТЬ" у пункті 30к, "БД(СЗ)"/"ОБОРОНА" у 70к/100к/170к, залежно від
    того, як САМЕ названо catch-all у ПЕРЕДАНОМУ categories) - це catch-all ("усе
    інше"): status_filter=None, exclude_status=усі НАЗВАНІ категорії пункту (включно з
    тими, що самі "include_to_report": False - їхні дні вже НЕ "усе інше", попри те, що
    не друкуються окремо в самому рапорті). Якщо жодна категорія пункту не позначена
    "default": True - у пункту просто немає catch-all запису в результаті.

    Якщо самого пункту point немає в categories (напр. точку 70 видалено цілком) -
    повертає порожній список: жодна категорія цього пункту НІКУДИ не потрапляє.

    Виняток пункту 100 (_RTGR_CROSS_TIER_SOURCE_CATEGORIES): catch-all пункту 100
    додатково НЕ приймає дні, форсовано позначені "70_РТГр"/"170_РТГр" - для них
    окремий "зайвий" рядок будує rtgr_cross_tier_rows, САМЕ в секції "100_РТГр"
    (див. коментар над _RTGR_CROSS_TIER_SOURCE_CATEGORIES вище)."""
    if point not in categories:
        return []
    point_categories = categories[point]
    catchall_category_name = next(
        (name for name, config in point_categories.items() if name != "general" and isinstance(config, dict) and config.get("default")),
        None,
    )
    named_categories = [name for name in point_categories if name not in ("general", catchall_category_name)]

    specs = [
        {"category_name": name, "status_filter": name, "exclude_status": None}
        for name in named_categories if point_categories[name].get("include_to_report", True)
    ]
    if catchall_category_name is not None and point_categories[catchall_category_name].get("include_to_report", True):
        exclude_status = set(named_categories)
        if point == 100:
            # Лише ті з _RTGR_CROSS_TIER_SOURCE_CATEGORIES, що дійсно існують ЯК
            # НАЗВАНА категорія ІНШОГО пункту цього Ж categories (напр. "70_РТГр" у
            # MONEY_REPORT_CATEGORIES[70]) - щоб виняток не "протікав" у синтетичні/
            # тестові набори categories, які цієї бізнес-логіки взагалі не моделюють.
            exclude_status |= {
                name for name in _RTGR_CROSS_TIER_SOURCE_CATEGORIES
                if any(name in other_categories for other_point, other_categories in categories.items() if other_point != point)
            }
        specs.append({"category_name": catchall_category_name, "status_filter": None, "exclude_status": exclude_status})
    return specs


def rtgr_cross_tier_rows(rows_with_data, date_columns, status_lookup, categories, extra_grounds_by_person=None):
    """"Зайвий" рядок секції "100_РТГр" для в/сл, чий день форсовано позначений
    "70_РТГр"/"170_РТГр" (категорії пунктів 70/170), а не буквально "100_РТГр" -
    той самий принцип, що й _COMBINED_TARGET_CELL_VALUES[100] для звичайної "100"
    (хто отримав 70/170, автоматично отримує рядок і в 100-й), але тут - для
    "100_РТГр" конкретно, щоб такі в/сл лишались РАЗОМ з рештою РТГр, а не в
    загальній "100"/"БД(СЗ)" (звідки їх, відповідно, вже виключає
    resolve_reportable_categories - див. _RTGR_CROSS_TIER_SOURCE_CATEGORIES).

    Повертає [] якщо в categories взагалі немає пункту "100_РТГр" (напр.
    COMMANDER_MONEY_REPORT_CATEGORIES, де РТГр-категорій немає)."""
    if _RTGR_CROSS_TIER_TARGET_POINT not in categories:
        return []
    specs = resolve_reportable_categories(_RTGR_CROSS_TIER_TARGET_POINT, categories)
    if not specs:
        return []
    catchall_spec = next((spec for spec in specs if spec["status_filter"] is None), specs[0])
    category_name = catchall_spec["category_name"]
    return build_category_rows(
        rows_with_data, date_columns, _RTGR_CROSS_TIER_CELL_VALUES, status_lookup,
        status_filter=_RTGR_CROSS_TIER_SOURCE_CATEGORIES,
        category_config=resolve_category_config(_RTGR_CROSS_TIER_TARGET_POINT, category_name, categories),
        category_name=category_name, extra_grounds_by_person=extra_grounds_by_person, categories=categories,
    )


def build_legal_basis_text(dates, category_config, extra_lines=None):
    """Будує текст колонки "Підстава для виплати" для набору дат участі за конфігом
    категорії з constants.MONEY_REPORT_CATEGORIES ({"grounds", "use_brs",
    "use_brs_from_selected_folder", "general"}), у пріоритеті:
    1. (якщо use_brs=True) 'бат'+'посилання_бат'+'посилання_брг' з тижневого
       NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK;
    1a. (якщо use_brs_from_selected_folder=True) те саме, але зі ЩОДЕННОГО
        NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY (_build_brs_chain_lines_from_selected_folder),
        заповненого лише якщо на питання "Зчитати обрану папку з документами (для
        номерів БАТ / БЗ)?" відповіли "Так" - для категорій, кому підстава: щоденні, а
        не тижневі бойові розпорядження (напр. МЕДИК пункту 100). Якщо відповіли "Ні",
        відповідні записи лишаються порожніми і ця підстава виходить порожньою -
        _log_missing_legal_basis (generate_report_for_get_money.py) тоді повідомляє про
        це окремим повідомленням;
    2. "general" - вже відфільтрований (resolve_category_config) плаский список записів
       constants.MONEY_REPORT_GENERAL_REFERENCES для цієї категорії;
    3. "grounds" - вручну вказані підстави САМЕ цієї категорії, в кінці, поверх усього іншого;
    4. extra_lines - ПЕРСОНАЛЬНІ рядки цієї КОНКРЕТНОЇ людини (не категорії) - напр.
       content.document_content_search.find_person_document_references (пункт 100
       розділу змін - номери документів, де реально згадане ПІБ саме цієї людини)."""
    if not dates:
        return ""

    lines = []
    if category_config.get("use_brs"):
        lines.extend(build_brs_chain_lines(dates))
    if category_config.get("use_brs_from_selected_folder"):
        lines.extend(_build_brs_chain_lines_from_selected_folder(dates))

    lines.extend(_lines_for_period(category_config.get("general", []), dates[0], dates[-1]))
    lines.extend(_lines_for_period(category_config.get("grounds", []), dates[0], dates[-1]))
    if extra_lines:
        lines.extend(extra_lines)

    # Дедуплікація (зі збереженням порядку першої появи): той самий рядок може
    # трапитись двічі - або через випадковий дубль у самому "lines" одного запису
    # constants.py, або коли ДВА різні (перетинаючись за періодом) записи
    # general/grounds містять один і той самий текст. build_brs_chain_lines вже
    # дедуплікує сама (за номером наказу) - тут дедуплікується решта (general/grounds).
    lines = list(dict.fromkeys(lines))

    return "\n".join(lines)


def _matches_filter(value, filter_spec):
    """filter_spec може бути None (немає фільтра), одним значенням, або колекцією
    значень (список/кортеж/множина) - у такому разі перевіряється належність."""
    if isinstance(filter_spec, (list, tuple, set, frozenset)):
        return value in filter_spec
    return value == filter_spec


# "Очк.БЗ" (constants.MONEY_REPORT_CATEGORIES["100_СПЕЦКОНТИНГЕНТ"]) - єдина
# категорія, чиї дні НЕМОЖЛИВО вивести зі сканування комірок дат ОБЛІК.xlsx (як
# решта категорій/пунктів) - для таких людей комірки дат просто не заповнюють.
# Це бізнес-правило (не структура даних), тож, як і _COMBINED_TARGET_CELL_VALUES/
# _POINT_PRIORITY_ORDER вище, задається явно тут, у коді - підтверджено користувачем.
_DISAPPEARANCE_ACCRUAL_CATEGORY = "Очк.БЗ"


def _disappearance_accrual_dates(person_row, date_columns):
    """Дати участі для _DISAPPEARANCE_ACCRUAL_CATEGORY - від дати зникнення
    (oblik_disappearance_date) до ОСТАННЬОЇ дати з date_columns включно (кінець
    обраного місяця генерації), а не лише дні самого date_columns - людина
    рахується СПЕЦКОНТИНГЕНТОМ безперервно від моменту зникнення, тож період
    охоплює й усі попередні місяці (підтверджено користувачем).

    [] (людина не потрапляє в рапорт цією категорією), якщо: дати зникнення
    немає взагалі чи вона порожня; вона в форматі, який to_date не розпізнає
    (обидва - _parsed_disappearance_date); вона ПІЗНІШЕ останньої дати
    date_columns; чи сам date_columns порожній."""
    if not date_columns:
        return []
    start_date = _parsed_disappearance_date(person_row)
    if start_date is None:
        return []

    end = max(date_columns)
    current = datetime(start_date.year, start_date.month, start_date.day)
    if current > end:
        return []

    dates = []
    while current <= end:
        dates.append(current)
        current += timedelta(days=1)
    return dates


def _resolve_matching_dates(person_row, date_columns, target_cell_values, base_status_category, status_filter, exclude_status):
    """Дати ОДНІЄЇ людини, що відповідають одній категорії - той самий фільтр
    (target_cell_values/status_filter/exclude_status, через resolve_day_value_and_category),
    що й build_category_rows (звідки цей код і винесено) - спільна основа й для
    category_dates_by_person нижче (без побудови рядка таблиці)."""
    dates = []
    for col in date_columns:
        day_value, day_category = resolve_day_value_and_category(person_row.get(col), base_status_category)
        if day_value not in target_cell_values:
            continue
        if exclude_status is not None and _matches_filter(day_category, exclude_status):
            continue
        if status_filter is not None and not _matches_filter(day_category, status_filter):
            continue
        dates.append(col)
    return dates


def category_dates_by_person(rows_with_data, date_columns, target_cell_values, status_lookup, status_filter, exclude_status=None, categories=None):
    """Як build_category_rows, але замість готових рядків таблиці (ПЕРІОД/ДНІ/
    ПІДСТАВА) повертає {normalize_name(ПІБ): [дата, ...]} - лише дні, що дійсно
    відповідають цій категорії, без побудови тексту підстави чи рядка взагалі.

    Для випадків, де потрібні САМЕ дати (не готовий текст): content.report_changes -
    дельта точки 30 для "Викласти в новій редакції" і дати участі точок 100/170
    для content.document_content_search.warn_missing_document_coverage (перевірка,
    що на кожен такий день дійсно є документ-підстава).

    НЕ підтримує category_name == _DISAPPEARANCE_ACCRUAL_CATEGORY ("Очк.БЗ") - її
    дати рахуються інакше (_disappearance_accrual_dates); жоден із цих викликачів
    цю категорію не використовує, тож тут навіть немає відповідного параметра."""
    result = {}
    for person_row in rows_with_data:
        base_status_category = resolve_status_category(
            status_lookup.get(normalize_name(person_row.get("ПІБ", ""))), person_row.get("ПІДРОЗДІЛ"), categories,
        )
        dates = _resolve_matching_dates(person_row, date_columns, target_cell_values, base_status_category, status_filter, exclude_status)
        if dates:
            result[normalize_name(person_row.get("ПІБ", ""))] = dates
    return result


def build_category_rows(rows_with_data, date_columns, target_cell_values, status_lookup, status_filter, category_config, exclude_status=None, category_name=None, extra_grounds_by_person=None, categories=None):
    """Формує рядки таблиці для однієї категорії з constants.MONEY_REPORT_CATEGORIES
    (напр. MONEY_REPORT_CATEGORIES[30]["ЗВРез"] чи MONEY_REPORT_CATEGORIES[100]["МЕДИК"]).

    categories - передається напряму в resolve_status_category (те саме джерело, з якого
    викликач узяв category_config/spec - за замовчуванням constants.MONEY_REPORT_CATEGORIES,
    для звіту на командира/ТВО - constants.COMMANDER_MONEY_REPORT_CATEGORIES) - день людини
    інакше міг би НЕ розпізнатись як належний до категорії, яка існує лише в переданому
    наборі, а не в глобальному MONEY_REPORT_CATEGORIES.

    target_cell_values - множина значень комірки ОБЛІК.xlsx, які рахуються як "день участі"
    для цієї категорії: {30} для категорій пункту 1, {100} для більшості категорій пункту 2,
    {30, 100} для медиків - для медиків рахується будь-який робочий день (30 чи 100), не
    рахуються лише дні з відпустками/відрядженнями тощо (ВП, ВД, ШП і подібні нечислові коди).
    status_filter=None -> без фільтра за статусом. Інакше - одна назва категорії ("МЕДИК")
    або колекція назв (наприклад {"МЕДИК", "ЗВРез"} для "усі, крім цих двох" разом з
    exclude_status) - лише ті ДНІ, чия категорія відповідає.
    exclude_status - так само, одне значення чи колекція - які дні НЕ повинні потрапити
    (напр. медики виключаються з пункту "ЖИТТЄДІЯЛЬНІСТЬ", бо їхні дні вже враховані
    єдиним блоком у категорії "МЕДИК" пункту 2).

    Категорія кожного ДНЯ (а не людини загалом) визначається resolve_day_value_and_category:
    для звичайних чисел (30/100) - це базовий статус людини з колонки СТАТУС (як і раніше);
    але якщо в комірці конкретного дня стоїть текст-назва категорії (напр. "РТГр") - саме
    ЦЕЙ день належить цій категорії, незалежно від колонки СТАТУС. Так одна людина може
    потрапити різними днями в різні категорії/пункти (напр. була ЗВРЕЗ, з певної дати -
    БД(СЗ), ще пізніше кілька днів - РТГр), без окремої таблиці періодів.
    category_config - див. build_legal_basis_text.

    Порядок рядків результату - точно як rows_with_data (тобто як в ОБЛІК.xlsx):
    жодного додаткового сортування (напр. за підрозділом) тут немає.

    Штатний командир батальйону (і ТВО на цю посаду) сюди НЕ жорстко зашиті - хто
    саме потрапляє в rows_with_data, вирішує викликач (generate_report_for_get_money.py
    прибирає їх зі свого списку заздалегідь, а generate_report_for_commander_money.py -
    навпаки, будує список САМЕ з них, для окремого рапорту від першої особи).

    extra_grounds_by_person - опційно, {normalize_name(ПІБ): [рядок, ...]} - ПЕРСОНАЛЬНІ
    (не категорійні) додаткові рядки підстави для конкретних людей (передається в
    build_legal_basis_text як extra_lines). Два джерела: content.report_changes
    використовує це для пункту 100 розділу змін (content.document_content_search);
    generate_report_for_get_money._build_combined_section передає сюди
    build_pidstavy_extra_grounds_by_person (колонка "ПІДСТАВИ" з ОБЛІК.xlsx) для
    КОЖНОГО пункту/категорії - обидва джерела просто ДОДАЮТЬ рядок(и) до вже
    зібраної підстави, незалежно одне від одного.

    Виняток: category_name == _DISAPPEARANCE_ACCRUAL_CATEGORY ("Очк.БЗ") - дати
    участі НЕ зі сканування date_columns (для таких людей комірки дат просто не
    заповнюють), а напряму від дати зникнення до останньої дати date_columns
    (_disappearance_accrual_dates) - лише для людей, чий БАЗОВИЙ статус (колонка
    СТАТУС) саме "Очк.БЗ"; підстава й решта полів рядка будуються так само, як
    і для решти категорій."""
    rows = []
    for person_row in rows_with_data:
        base_status_category = resolve_status_category(
            status_lookup.get(normalize_name(person_row.get("ПІБ", ""))), person_row.get("ПІДРОЗДІЛ"), categories,
        )

        if category_name == _DISAPPEARANCE_ACCRUAL_CATEGORY:
            if base_status_category != _DISAPPEARANCE_ACCRUAL_CATEGORY:
                continue
            dates = _disappearance_accrual_dates(person_row, date_columns)
        else:
            dates = _resolve_matching_dates(person_row, date_columns, target_cell_values, base_status_category, status_filter, exclude_status)

        if not dates:
            continue

        period_text, days_count = build_period_text_and_days(dates)
        normalized_pib = normalize_name(person_row.get("ПІБ", ""))
        extra_lines = (extra_grounds_by_person or {}).get(normalized_pib)

        rows.append({
            "ПІДРОЗДІЛ": person_row.get("ПІДРОЗДІЛ", ""),
            "ПОСАДА": person_row.get("ПОСАДА", ""),
            "ЗВАННЯ": person_row.get("ЗВАННЯ", ""),
            "ПІБ": person_row.get("ПІБ", ""),
            "ПЕРІОД": period_text,
            "ДНІ": days_count,
            # Лише для "100_СПЕЦКОНТИНГЕНТ" (exclude_rows_without_disappearance_date
            # нижче/generate_report_for_get_money) - "" для решти пунктів/категорій,
            # де ця колонка взагалі не рендериться, обчислюється однаково для
            # КОЖНОГО рядка (дешево - один прямий доступ до вже наявного person_row),
            # а не лише для category_name-ів "100_СПЕЦКОНТИНГЕНТ", щоб не плутати
            # виклик з тим, ЯКИЙ САМЕ набір category_name належить цьому пункту.
            "ДАТА_ЗНИКНЕННЯ": _disappearance_date_text(person_row),
            "ПІДСТАВА": build_legal_basis_text(dates, category_config, extra_lines=extra_lines),
            # Внутрішні прапорці - НЕ виводяться в саму таблицю (add_category_section
            # явно перелічує лише потрібні для документа поля) - лише для
            # _log_missing_legal_basis, щоб не попереджати про категорії, де порожня
            # підстава - очікувана (див. resolve_category_config/"required_basis"), і щоб
            # для категорій з "use_brs_from_selected_folder" (напр. МЕДИК пункту 100) при
            # порожній підставі показати не загальне попередження, а конкретне повідомлення
            # про те, що обрану папку з документами не було зчитано (чи в ній не було
            # потрібних документів на ці дати).
            "_ПІДСТАВА_ОБОВ'ЯЗКОВА": category_config.get("required_basis", True),
            "_ВИКОРИСТОВУЄ_ОБРАНУ_ПАПКУ": category_config.get("use_brs_from_selected_folder", False),
            "_КАТЕГОРІЯ": category_name,
        })

    return rows


# Той самий рядок-літерал, що й content.report_document_reader.SPETSKONTYNGENT_VALUE -
# НЕ імпортується звідти (report_document_reader САМ імпортує normalize_name
# ЗВІДСИ, тож зворотний імпорт створив би циклічну залежність), тому дублюється
# тут одним константним рядком, а не рядковим літералом у кожному місці нижче.
_SPETSKONTYNGENT_POINT = "100_СПЕЦКОНТИНГЕНТ"


def exclude_rows_without_disappearance_date(rows, point):
    """Для point == _SPETSKONTYNGENT_POINT - прибирає з rows кожного, чия
    "ДАТА_ЗНИКНЕННЯ" порожня (build_category_rows уже проставила її з
    ОБЛІК.xlsx для КОЖНОГО рядка САМЕ цього пункту, незалежно від того, з якої
    саме category_name він прийшов - "полон"/"безвісті"/"інд"/"Очк.БЗ"), з
    попередженням у термінал про кожного виключеного - підтверджено
    користувачем: без задокументованої дати (ОБЛІК.xlsx, аркуш "Табель",
    колонка "ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ"/"ДАТА ЗНИКНЕННЯ") людина в цю
    таблицю не потрапляє взагалі.

    Живе тут (не в generators/generate_report_for_get_money.py), бо
    "100_СПЕЦКОНТИНГЕНТ" може з'явитись і в ЗВИЧАЙНОМУ рапорті поточного
    місяця (generate_report_for_get_money._build_categories), і в
    РЕТРОАКТИВНИХ додаваннях за попередні місяці
    (content.report_changes._basis_appeared_extra_points_for_month, теж
    будує рядки через build_category_rows) - ОБИДВА викликачі мають
    застосовувати цю саму перевірку, тож вона - тут, в ОДНОМУ спільному
    місці, а не дублюється в кожному з них.

    Пункти ПОЗА _SPETSKONTYNGENT_POINT - rows повертається БЕЗ ЖОДНОЇ зміни."""
    if point != _SPETSKONTYNGENT_POINT:
        return rows
    kept_rows, excluded_rows = [], []
    for row in rows:
        (kept_rows if row.get("ДАТА_ЗНИКНЕННЯ") else excluded_rows).append(row)
    for row in excluded_rows:
        print_red(f"⚠ Немає ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ для {row['ПІБ']} - виключено з рапорту.")
    return kept_rows
