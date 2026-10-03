import calendar
from collections import Counter
from datetime import datetime
from functools import lru_cache

from constants import (
    MONEY_REPORT_CATEGORIES,
    MONEY_REPORT_CHANGES_ORDER_REFERENCES,
    MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR,
    MONEY_REPORT_GENERAL_REFERENCES_BN,
    GENERAL_PIDSTAVY_COLUMN_NAMES,
    BASIS_REQUIRED_POINT_COLUMN_NAMES,
    MONTH,
    YEAR,
)
from checker_accounting.checker import _read_check_file
from utils.date_utils import to_date
from content.document_content_search import find_person_document_references, collect_missing_document_coverage_warnings
from content.money_report_helpers import (
    build_category_rows,
    build_legal_basis_text,
    build_period_text_and_days,
    build_pidstavy_extra_grounds_by_person,
    category_dates_by_person,
    resolve_category_config,
    resolve_reportable_categories,
    exclude_rows_without_disappearance_date,
    _DISAPPEARANCE_DATE_COLUMN_NAMES,
    normalize_name,
    month_nominative_upper,
    _merge_preserving_file_order,
    _exclude_by_excluded_day_value,
    _exclude_commander_and_tvo,
    _keep_only_commander_and_tvo,
    _COMBINED_TARGET_CELL_VALUES,
    BASIS_REQUIRED_POINTS,
)

# "Прошу внести зміни в наказ ..." (додаток до рапорту за ПОПЕРЕДНІ місяці) - уся
# ця логіка працює з prev_MM_ОБЛІК/actual_MM_ОБЛІК напряму (НЕ з готовим
# changes_MM_ОБЛІК.xlsx, який sync/changes_folder.py вже зберіг лише як
# аудиторський артефакт для користувача - його кольори не несуть достатньо
# інформації, щоб АВТОМАТИЧНО побудувати "виключити"/"додати" по кожній категорії).

# Нумерація додатків у ЦЬОМУ розділі має СВІЙ пріоритет - НЕ той самий, що й звичайні
# пункти рапорту (_POINT_PRIORITY_ORDER = 30, 100, 70, 170, 10) - підтверджено
# користувачем: Додаток 1 - ЗАВЖДИ 100, Додаток 2 - ЗАВЖДИ 50, Додаток 3 - ЗАВЖДИ
# 30 (перелік поки неповний - нові пункти дописувати в кінець за потреби). Номер -
# ФІКСОВАНА позиція точки в цьому списку (index+1), а НЕ "стиснута" позиція серед
# лише присутніх цього місяця пунктів: якщо точка 50 цього місяця без змін (як
# майже завжди - її ще навіть не введено в MONEY_REPORT_CATEGORIES) - "Додаток 2"
# просто НЕ З'ЯВЛЯЄТЬСЯ В ДОКУМЕНТІ ВЗАГАЛІ, а точка 30 однаково лишається
# "Додатком 3" (а НЕ "зсувається" на 2) - підтверджено користувачем ДВІЧІ:
# "є тільки 1 додаток це 100 та 3 додаток це 30, іншого не дано" і повторно
# "зроби так щоб додаток 2 взагалі не з'являвся" (раніше тут була "стиснута"
# нумерація - позиція серед ЛИШЕ фактично змінених пунктів - через яку точка 30
# помилково ставала "Додатком 2", коли 50 без змін; замінено на фіксовану).
#
# 170 (і 70) СВІДОМО відсутні тут окремо - як і в звичайному щомісячному
# рапорті (_COMBINED_TARGET_CELL_VALUES нижче), їхні дні вже рахуються
# ДОДАТКОВИМ рядком усередині пункту 100 ("хто отримав компенсацію за 70 чи
# 170, автоматично отримує ще й рядок у 100-й") - додавання 170 сюди ЯК
# ОКРЕМИЙ пункт дублювало б кожну людину (і в Додатку 1, і в зайвому додатку).
_CHANGES_POINT_PRIORITY_ORDER = [100, 50, 30]

# Підстава з обраної папки документів - ЛИШЕ через find_person_document_references
# (ПІБ, реально згадане в тексті ЩОДЕННА/ЗАВДАННЯ файлу - extra_grounds_by_person)
# і ЛИШЕ для пункту 100 - підтверджено користувачем; НЕ через use_brs/
# use_brs_from_selected_folder (той механізм примусово вимкнено для ВСІХ пунктів
# у _changes_category_config, навіть для 100 - див. її docstring). ТОЙ САМИЙ
# пункт - єдиний, для якого перевіряється ПОКРИТТЯ документами по КОЖНОМУ дню
# участі (content.document_content_search.warn_missing_document_coverage) - і
# лише в цьому розділі (рапорт на поправки в наказі за попередні місяці), не в
# звичайному щомісячному рапорті. 170 (і 70) окремо тут НЕ згадуються - їхні дні
# вже входять до target_cell_values пункту 100 через _COMBINED_TARGET_CELL_VALUES
# (money_report_helpers.py), тож обидва ці механізми (підстава з папки, перевірка
# покриття) АВТОМАТИЧНО поширюються й на них, без окремого запису тут.
_FOLDER_BASED_GROUNDS_POINTS = {100}

# "Виключити пункти в Додатку N" НЕ існує в ЖОДНОМУ пункті цього розділу
# (_delta_rows_for_point нижче) - підтверджено користувачем: "якщо людина вже
# отримала 100 тис, не треба знімати її з 100 тис, щоб знову щось заплатити за
# цей день - це працює лише в один бік" (30->100 - нова участь, документується
# як "доповнити"; 100->[щось інше] - вона вже давно отримала свої 100 тис, тож
# НЕ "виключається" нізвідки). Кожен пункт показує ЛИШЕ НОВІ/змінені дні (ті,
# що з'явились в actual, але яких НЕ було в тій самій категорії в prev) - дні,
# що вже й раніше рахувались цією категорією, НЕ повторюються в результаті,
# навіть якщо загальний період людини змінився.
#
# RESTATE_POINTS - лише ВИБІР ПІДПИСУ підпункту, не окрема обчислювальна логіка
# (обчислення - _delta_rows_for_point, ОДНАКОВЕ для всіх пунктів): 30 підписує
# свій підпункт як "Викласти в новій редакції в Додатку N", решта пунктів
# (100, 50, ...) - як звичайне "Додаток N доповнити наступними пунктами".
# Публічна (без "_") - generators.generate_report_for_get_money._add_changes_sections
# звіряється з нею, щоб обрати підпис.
RESTATE_POINTS = {30}

# Усі назви колонок підстави, що можуть трапитись у prev/actual ОБЛІК.xlsx -
# загальна GENERAL_PIDSTAVY_COLUMN_NAMES + ВЛАСНА колонка КОЖНОГО пункту
# BASIS_REQUIRED_POINT_COLUMN_NAMES (constants.py), + обидва написання колонки
# дати зникнення (_DISAPPEARANCE_DATE_COLUMN_NAMES, money_report_helpers.py) -
# "100_СПЕЦКОНТИНГЕНТ" (exclude_rows_without_disappearance_date) потребує ЇЇ
# так само, як і саму "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ" - _read_check_file читає лише
# колонки, чиї назви явно перелічені (extra_label_columns), тож усі мають бути
# тут, а не лише ті, що використовує ЦЕЙ конкретний запуск.
_PIDSTAVY_LABEL_COLUMNS = GENERAL_PIDSTAVY_COLUMN_NAMES + tuple(
    name for names in BASIS_REQUIRED_POINT_COLUMN_NAMES.values() for name in names
) + _DISAPPEARANCE_DATE_COLUMN_NAMES


@lru_cache(maxsize=None)
def _rows_from_check_file(file_path):
    """Адаптує checker_accounting.checker._read_check_file (динамічне визначення
    колонок ПІДРОЗДІЛ/ПОСАДА/ЗВАННЯ/ПІБ за назвою заголовка й дат за типом -
    придатне для довільного, не обов'язково ідентичного за структурою файлу,
    на відміну від статичних PERSONEL_LIST_COLUMNS_LETTERS) у ту саму форму
    (rows_with_data, date_columns), яку очікують build_category_rows/build_legal_basis_text:
    rows_with_data - список {"ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", <datetime>: значення, ...};
    date_columns - відсортований список datetime (НЕ date - _read_check_file повертає
    date, але build_legal_basis_text/_lines_for_period через to_date() приймають лише
    datetime чи рядок, тож дати конвертуються тут же, один раз).

    Кешується за file_path (@lru_cache) - той самий файл (prev_MM_ОБЛІК/actual_MM_ОБЛІК)
    інакше перечитується (openpyxl) до 3 разів за один запуск (describe_eligible_changes_months
    -> _eligible_month_candidates, build_changes_entries -> _eligible_month_candidates
    ЗНОВУ, build_appendix_pairs_for_month) - файл за час одного запуску генератора не
    змінюється, тож повторний парсинг того самого шляху - чистий зайвий I/O. Викликачі
    НЕ мутують повернені rows_with_data/date_columns на місці (лише читають чи будують
    з них НОВІ списки/словники), тож спільний кешований об'єкт між викликами безпечний.

    extra_label_columns=_PIDSTAVY_LABEL_COLUMNS (загальна GENERAL_PIDSTAVY_COLUMN_NAMES
    + ВЛАСНА колонка кожного з BASIS_REQUIRED_POINT_COLUMN_NAMES, constants.py) - на
    відміну від звичайного виклику _read_check_file (checker_accounting.checker.
    _build_check_result) ці колонки тут ПОТРІБНІ: build_pidstavy_extra_grounds_by_person/
    _basis_appeared_add_rows читають їх з prev/actual rows_with_data (BASIS_REQUIRED_POINTS -
    ретроактивне додавання, коли підстава з'явилась лише в actual-файлі) - без цього
    колонка мовчки губилась би (не дата й не в базовому _LABEL_COLUMNS), і ця логіка
    НІКОЛИ не спрацьовувала б."""
    people, date_set = _read_check_file(file_path, extra_label_columns=_PIDSTAVY_LABEL_COLUMNS)
    date_columns = sorted(datetime.combine(d, datetime.min.time()) for d in date_set)
    datetime_by_date = {dt.date(): dt for dt in date_columns}

    rows_with_data = []
    for person in people:
        row = dict(person["label"])
        for plain_date, value in person["days"].items():
            row[datetime_by_date[plain_date]] = value
        rows_with_data.append(row)
    return rows_with_data, date_columns


_LABEL_FIELDS_TO_BACKFILL = ("ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ")


def _backfill_missing_labels(rows, other_rows):
    """Якщо запис людини (за нормалізованим ПІБ) у ЦЬОМУ файлі (prev чи actual) не
    має (чи має порожнім) щось із ПІДРОЗДІЛ/ПОСАДА/ЗВАННЯ - напр. попередній місяць
    зберегли скороченим експортом лише з ЗВАННЯ+ПІБ, без ПОСАДА - підставляємо
    значення з ІНШОГО файлу (other_rows) за тим самим ПІБ, якщо там воно є. Без
    цього людина в розділі "Прошу внести зміни..." лишалась би з порожньою
    колонкою "Посада" (чи "Підрозділ"/"Звання"), хоча ІНШИЙ файл цієї самої пари
    її знає."""
    other_by_pib = {normalize_name(row.get("ПІБ", "")): row for row in other_rows}
    result = []
    for row in rows:
        other_row = other_by_pib.get(normalize_name(row.get("ПІБ", "")))
        if other_row is None:
            result.append(row)
            continue
        patched = dict(row)
        for field in _LABEL_FIELDS_TO_BACKFILL:
            if not patched.get(field) and other_row.get(field):
                patched[field] = other_row[field]
        result.append(patched)
    return result


def _resolve_file_year_month(date_columns):
    """Визначає (рік, місяць) файлу за НАЙЧАСТІШЕ повторюваною парою (рік, місяць)
    серед його ж дат - а НЕ за 2-цифровим токеном у назві файлу (prev_06_ОБЛІК.xlsx):
    цей токен - лише спосіб зіставити prev з actual (sync/changes_folder.py), не
    джерело істини для хронологічного порівняння з обраним місяцем генерації.
    Порожній date_columns -> (None, None)."""
    if not date_columns:
        return None, None
    counts = Counter((d.year, d.month) for d in date_columns)
    return counts.most_common(1)[0][0]


def _months_before_selected(year, month):
    """Кількість місяців, на яку (рік, місяць) РАНІШЕ за обраний місяць генерації
    (constants.MONTH/YEAR) - додатне число означає, що файл дійсно описує ПОПЕРЕДНІЙ
    місяць; 0 чи від'ємне - файл описує ПОТОЧНИЙ чи МАЙБУТНІЙ (відносно обраного)
    місяць, тобто не підходить для розділу "зміни за попередні місяці". Порівняння
    ЗАВЖДИ через (рік*12 + місяць), НІКОЛИ напряму за номером місяця - інакше грудень
    ПОПЕРЕДНЬОГО року помилково виглядав би "пізнішим" за січень наступного (12 > 01)."""
    selected_index = int(YEAR) * 12 + int(MONTH)
    file_index = year * 12 + month
    return selected_index - file_index


def _changes_category_config(point, category_name, categories):
    """resolve_category_config, але з примусово ВИМКНЕНИМ use_brs/
    use_brs_from_selected_folder для БУДЬ-ЯКОГО пункту (навіть 100) -
    підтверджено користувачем: НІКОЛИ не тягнути дані з
    NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK/_DAY (use_brs/use_brs_from_selected_folder),
    навіть якщо вони щойно синхронізовані з ТІЄЇ САМОЇ обраної папки -
    sync.changes_folder.prompt_grounds_folders_for_changes - той синк
    призначений для ЗВИЧАЙНОГО рапорту.

    Додатково, ЛИШЕ для пунктів "Додаток N" (_CHANGES_POINT_PRIORITY_ORDER -
    100/50/30) - "general"/"grounds" ТЕЖ примусово перевизначені: підстава
    складається ЛИШЕ з MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR/_BN (за датою) і,
    для точки 100, окремого content_search (extra_grounds_by_person) - підтверджено
    користувачем (двічі): "має тільки бути MONEY_REPORT_GENERAL_REFERENCES_BN та
    MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR", НІКОЛИ "grounds"/своїм "general"
    самої категорії (напр. точка 30 "ЖИТТЄДІЯЛЬНІСТЬ" у ЗВИЧАЙНОМУ рапорті бере
    "grounds"/"general" з BR_HIGHT_UNIT - зовсім ІНШИЙ, не LOG_WAR/BN, набір
    посилань, який раніше мовчки протікав і сюди, у розділ змін).

    Пункти ПОЗА цим списком (BASIS_REQUIRED_POINTS - "100_СПЕЦКОНТИНГЕНТ" тощо, і
    70/170 у _extra_point_rows_for_month) лишають "general"/"grounds" САМОЇ
    категорії БЕЗ ЗМІН: BASIS_REQUIRED_POINTS у реальному constants.py й так
    завжди мають "general": []/"grounds": [] (їхня єдина підстава - конкретний
    наказ із колонки ПІДСТАВИ, а не загальні бойові посилання) - домішувати сюди
    LOG_WAR/BN було б новою помилкою, не тим, про що просив користувач; 70/170 у
    реальному constants.py й так мають "general": [*LOG_WAR, *BN] (той самий
    список - перевизначення тут нічого не змінило б у продакшн, лише зайве в
    тестах із синтетичними categories), а їхня категорія "70_РТГр"/"170_РТГр" має
    ВЛАСНІ, специфічні "grounds" (exclude_general="all") - обнулення тут
    зіпсувало б саме їх. Решта конфігу (required_basis тощо) - без змін."""
    config = resolve_category_config(point, category_name, categories)
    config = {**config, "use_brs": False, "use_brs_from_selected_folder": False}
    if point in _CHANGES_POINT_PRIORITY_ORDER:
        config["general"] = [*MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR, *MONEY_REPORT_GENERAL_REFERENCES_BN]
        config["grounds"] = []
    return config


def _delta_rows_for_point(prev_rows, prev_date_columns, actual_rows, actual_date_columns, point, categories, extra_grounds_by_person=None):
    """"Додаток N доповнити"/"Викласти в новій редакції в Додатку N" для
    БУДЬ-ЯКОГО пункту цього розділу - "Виключити" тут НІКОЛИ не рендериться
    (для ЖОДНОГО пункту) - підтверджено користувачем: "якщо людина вже
    отримала 100 тис, не треба знімати її з 100 тис, щоб знову щось заплатити
    за цей день - це працює лише в один бік (30->100 ок, а не навпаки)".

    ГЕЙТИНГ (кого взагалі включати) і ЗМІСТ рядка (які дати показувати) -
    РІЗНІ для звичайних пунктів і RESTATE_POINTS:

    - Звичайні пункти (напр. 100, "Додаток N доповнити") - гейтинг за ДЕЛЬТОЮ
      (є хоч один НОВИЙ день - category_dates_by_person, actual мінус prev),
      зміст рядка - ЛИШЕ ЦЯ дельта: підтверджено користувачем ("не треба
      знімати з 100 тис, щоб знову щось заплатити" - уже сплачені дні НЕ
      повторюються). Дні, що ЗНИКЛИ (були в prev, немає в actual -
      реклассифікація в іншу категорію чи повна втрата) - НЕ згадуються НІЯК.

    - RESTATE_POINTS (зараз - 30, "Викласти в новій редакції") - гейтинг за
      БУДЬ-ЯКОЮ різницею між actual і prev (не лише нові дні, а й дні, що
      ЗНИКЛИ) - підтверджено користувачем на прикладі АРТЬОМОВА: частина його
      30-днів (20.05-31.05) переїхала в 100 (це вже показано в Додатку 1), а
      РЕШТА (07.05-17.05, залишились НЕЗМІННИМИ) усе одно мала з'явитись
      власним записом у Додатку 3 - сам факт, що ЧАСТИНА старого запису
      зникла, означає, що ввесь запис змінився й має бути переоформлений.
      Зміст рядка - ПОВНИЙ актуальний період категорії (усі дні actual, а не
      лише дельта чи різниця) - "Викласти в новій редакції" замінює ВЕСЬ
      запис Додатка новим текстом, тож має описувати його ПОВНІСТЮ. Людина БЕЗ
      жодного дня цієї категорії в actual (повна втрата, нічого не лишилось) -
      відсутня в actual_dates_by_person узагалі (category_dates_by_person не
      повертає порожніх записів), тож автоматично не рендериться - як і для
      звичайних пунктів, "повна втрата" лишається мовчазною.

    "_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ" (bool, лише для RESTATE_POINTS - для решти
    пунктів завжди False) - чи МАЄ САМЕ ЦЕЙ рядок підписуватись "Викласти в
    новій редакції", а не звичайним "Додаток N доповнити" - True, лише якщо в
    ЦІЄЇ людини БУВ хоч один день цієї категорії в prev (є що "перевидавати");
    False, якщо prev_dates_set порожній (людина щойно долучилась - в неї
    НЕМАЄ ні 30, ні 100 в prev - Додаток тут не "перевидається", а вперше
    доповнюється) - підтверджено користувачем: "якщо в табелі з префіксом prev
    взагалі не фігурувала ні 30 ні 100, тоді має бути... Додаток 3 доповнити
    наступними пунктами", а не "Викласти в новій редакції". Той самий пункт
    (30) може дати ОБИДВА варіанти в одному місяці одночасно, для РІЗНИХ
    людей - generate_report_for_get_money._add_changes_sections рендерить їх
    ДВОМА окремими підпунктами під ТИМ САМИМ номером Додатка.

    extra_grounds_by_person - див. build_category_rows (content_search - ЛИШЕ
    для пунктів із _FOLDER_BASED_GROUNDS_POINTS, викликач сам вирішує, що
    передати).

    RESTATE_POINTS - "ПІДСТАВА" ЗАВЖДИ ПОРОЖНЯ (build_legal_basis_text тут
    НЕ викликається), і "_ПІДСТАВА_ОБОВ'ЯЗКОВА" завжди False (незалежно від
    "required_basis" самої категорії) - підтверджено користувачем: "для
    Викласти в новій редакції в Додатку 3 не потрібно відображати підстави
    для 30к" - раніше тут все одно підставлялись LOG_WAR/BN (_changes_category_config),
    і коли жоден із них не перетинав період запису - _log_missing_legal_basis
    (generate_report_for_get_money.py) хибно попереджала "Немає підстави для
    виплати" для рядка, якому підстава тут узагалі не потрібна.

    Повертає ГОТОВІ рядки (як build_category_rows), за порядком actual_rows -
    порожній список, якщо в ЖОДНОЇ людини немає жодного нового дня цієї
    категорії (тоді сам пункт просто не з'являється в рапорті)."""
    status_lookup = {}
    target_cell_values = _COMBINED_TARGET_CELL_VALUES.get(point, {point})
    specs = resolve_reportable_categories(point, categories)
    actual_by_pib = {normalize_name(row.get("ПІБ", "")): row for row in actual_rows}

    delta_rows = []
    for spec in specs:
        category_config = _changes_category_config(point, spec["category_name"], categories)
        prev_dates_by_person = category_dates_by_person(
            prev_rows, prev_date_columns, target_cell_values, status_lookup,
            status_filter=spec["status_filter"], exclude_status=spec["exclude_status"], categories=categories,
        )
        actual_dates_by_person = category_dates_by_person(
            actual_rows, actual_date_columns, target_cell_values, status_lookup,
            status_filter=spec["status_filter"], exclude_status=spec["exclude_status"], categories=categories,
        )
        for pib, actual_dates in actual_dates_by_person.items():
            actual_dates_set = set(actual_dates)
            prev_dates_set = set(prev_dates_by_person.get(pib, []))
            restate_this_row = False
            if point in RESTATE_POINTS:
                # "Викласти в новій редакції" - гейтинг за БУДЬ-ЯКОЮ різницею
                # (не лише НОВІ дні, а й дні, що ЗНИКЛИ - напр. частина 30-днів
                # переїхала в 100, і РЕШТА 30-днів, що лишились, теж має бути
                # переоформлена як окремий запис) - підтверджено користувачем
                # на прикладі АРТЬОМОВА (30к 07.05-17.05 лишились НЕЗМІННИМИ,
                # але Додаток 3 усе одно мав з'явитись, бо ІНША частина
                # колишнього 30-денного запису (20.05-31.05) зникла).
                if actual_dates_set == prev_dates_set:
                    continue
                display_dates = sorted(actual_dates_set)
                # "Викласти в новій редакції" МАЄ СЕНС ЛИШЕ якщо в prev УЖЕ БУВ
                # запис цієї категорії, який тепер переоформлюється - людина БЕЗ
                # жодного дня цієї категорії в prev (ні 30, ні 100 - типовий
                # реальний випадок: щойно долучилась/повернулась) НІЧОГО не
                # "перевидає", вона просто вперше з'являється в Додатку - тож
                # рендериться як звичайне "Додаток N доповнити", а не "Викласти
                # в новій редакції" - підтверджено користувачем.
                restate_this_row = bool(prev_dates_set)
            else:
                new_dates = sorted(actual_dates_set - prev_dates_set)
                if not new_dates:
                    continue
                display_dates = new_dates
            person_row = actual_by_pib.get(pib)
            if person_row is None:  # pragma: no cover - pib тут ЗАВЖДИ прийшов з
                # category_dates_by_person(actual_rows, ...), яка ключує ТИМ САМИМ
                # normalize_name(ПІБ) із ТОГО САМОГО actual_rows, що й actual_by_pib -
                # цей pib фізично не може бути відсутнім у actual_by_pib; лишено як
                # захисна гілка на випадок майбутньої зміни джерела pib.
                continue
            period_text, days_count = build_period_text_and_days(display_dates)
            if point in RESTATE_POINTS:
                basis_text, basis_required = "", False
            else:
                extra_lines = (extra_grounds_by_person or {}).get(pib)
                basis_text = build_legal_basis_text(display_dates, category_config, extra_lines=extra_lines)
                basis_required = category_config.get("required_basis", True)
            delta_rows.append({
                "ПІДРОЗДІЛ": person_row.get("ПІДРОЗДІЛ", ""),
                "ПОСАДА": person_row.get("ПОСАДА", ""),
                "ЗВАННЯ": person_row.get("ЗВАННЯ", ""),
                "ПІБ": person_row.get("ПІБ", ""),
                "ПЕРІОД": period_text,
                "ДНІ": days_count,
                "ПІДСТАВА": basis_text,
                "_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ": restate_this_row,
                "_ПІДСТАВА_ОБОВ'ЯЗКОВА": basis_required,
                "_ВИКОРИСТОВУЄ_ОБРАНУ_ПАПКУ": category_config.get("use_brs_from_selected_folder", False),
                "_КАТЕГОРІЯ": spec["category_name"],
            })

    return _merge_preserving_file_order(actual_rows, delta_rows)


def _basis_appeared_add_rows(prev_rows, prev_date_columns, actual_rows, actual_date_columns, point, categories):
    """BASIS_REQUIRED_POINTS ('100_СПЕЦКОНТИНГЕНТ'/'100_БПШП'/'100_ВПБП', constants.py) -
    ЄДИНІ пункти, де людину без непорожньої ВЛАСНОЇ підстави (build_pidstavy_extra_grounds_by_person
    зі СВОЄЮ колонкою за constants.BASIS_REQUIRED_POINT_COLUMN_NAMES[point] - напр.
    "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ" для "100_СПЕЦКОНТИНГЕНТ", а НЕ спільна "ПІДСТАВИ
    ДЛЯ ІНШИХ СИТУАЦІЙ" - підтверджено користувачем: кожен статус має ВЛАСНУ
    колонку, щоб підстава одного не могла хибно "розблокувати" інший)
    виключає з рапорту ЦІЛКОМ generate_report_for_get_money._build_categories (звичайний
    щомісячний рапорт), а не _delta_rows_for_point вище (той дивиться лише на ДЕЛЬТУ
    днів участі - порожня "ПІДСТАВА" на неї не впливає). Тож людина, чий статус
    (полон/безвісти/...) лишався НЕЗМІННИМ між prev і actual, але за prev-місяць
    підстави ще не було (і саме тому вона НЕ потрапила в тодішній рапорт), а
    тепер (actual) підстава задокументована - для звичайної дельти виглядала б
    як "без змін" (був відсутній в build_category_rows ОБИДВА рази -
    _BASIS_REQUIRED_POINTS-фільтрація тут узагалі не застосовується, щоб
    побачити, хто МАВ БИ бути там за самою участю).

    Підтверджено користувачем: тоді цю людину все одно треба додати - РЕТРОАКТИВНО,
    Додатком за ТОЙ САМИЙ (попередній) місяць, а не поточний, з ПІДСТАВОЮ, що
    щойно з'явилась (actual), а не пустою, як була б за prev. Повертає ГОТОВІ
    рядки (ПІДРОЗДІЛ/ПОСАДА/ЗВАННЯ/ПІБ/ПЕРІОД/ДНІ/ПІДСТАВА, як і _delta_rows_for_point) -
    period/days рахуються з prev (коли людина реально брала участь), а ПІДСТАВА
    підставляється з actual_extra_grounds (build_category_rows/build_legal_basis_text
    робить це напряму, без окремого патчингу полів)."""
    column_names = BASIS_REQUIRED_POINT_COLUMN_NAMES[point]
    prev_extra_grounds = build_pidstavy_extra_grounds_by_person(prev_rows, column_names)
    actual_extra_grounds = build_pidstavy_extra_grounds_by_person(actual_rows, column_names)

    target_cell_values = _COMBINED_TARGET_CELL_VALUES.get(point, {point})
    specs = resolve_reportable_categories(point, categories)

    add_lists = []
    for spec in specs:
        category_config = _changes_category_config(point, spec["category_name"], categories)
        # extra_grounds_by_person=actual_extra_grounds (НЕ prev) - рядок будується за
        # участю PREV-місяця, але з ПІДСТАВОЮ, яка щойно з'явилась в actual.
        prev_rows_with_actual_basis = build_category_rows(
            prev_rows, prev_date_columns, target_cell_values, {},
            status_filter=spec["status_filter"], category_config=category_config,
            exclude_status=spec["exclude_status"], category_name=spec["category_name"],
            extra_grounds_by_person=actual_extra_grounds, categories=categories,
        )
        add_lists.append([
            row for row in prev_rows_with_actual_basis
            if normalize_name(row["ПІБ"]) not in prev_extra_grounds
            and normalize_name(row["ПІБ"]) in actual_extra_grounds
        ])

    return _merge_preserving_file_order(prev_rows, *add_lists)


def build_appendix_pairs_for_month(prev_path, actual_path, rows_with_tvo_data, categories=None, content_search=None, commander_and_tvo_only=False, narrator_row=None, missing_coverage_warnings=None):
    """Для ОДНІЄЇ пари prev/actual файлів - повертає список (point, appendix_number,
    exclude_rows, add_rows) лише для пунктів, де є ХОЧ ОДНА зміна (add - exclude
    завжди порожній, _delta_rows_for_point) - пункт без жодної зміни цього разу
    просто відсутній у результаті (а не показаний порожнім, і НЕ "зсуває" номери
    інших пунктів - див. _CHANGES_POINT_PRIORITY_ORDER). appendix_number -
    ФІКСОВАНА позиція точки в _CHANGES_POINT_PRIORITY_ORDER (100->Додаток 1,
    50->Додаток 2, 30->Додаток 3, ЗАВЖДИ, незалежно від того, які саме пункти
    присутні цього місяця) - підтверджено користувачем.

    НА ВІДМІНУ від звичайних пунктів рапорту (generate_report_for_get_money._build_categories,
    де нову точку можна додати без жодних змін коду) - тут перебираються РІВНО пункти
    _CHANGES_POINT_PRIORITY_ORDER, а не всі top-level ключі categories: псевдо-пункти
    на кшталт "100_ШП"/"NOT_PAID" (потрібні лише звичайному щомісячному рапорту, не
    цьому розділу) НЕ повинні самі собою ставати новим додатком тут лише через те,
    що вони є ключами MONEY_REPORT_CATEGORIES - підтверджено користувачем (без цього
    людина з категорії "100_ШП" помилково опинялась у зайвому "Додатку 3", хоча
    реальних додатків мало бути лише два - 100 і 30). Новий РЕАЛЬНИЙ пункт (напр.
    майбутня "50") має бути дописаний у сам _CHANGES_POINT_PRIORITY_ORDER явно.
    BASIS_REQUIRED_POINTS ("100_СПЕЦКОНТИНГЕНТ"/"100_БПШП"/"100_ВПБП") сюди
    (у "Додаток N") УЗАГАЛІ НЕ входять - рендеряться ОКРЕМО, "як у звичайному
    рапорті" (_basis_appeared_extra_points_for_month нижче, той самий принцип,
    що й 70/170 - SECTIONS у resources/data.json, ПОЗА межами "Прошу внести
    зміни..."), а не як "Додаток N доповнити" - підтверджено користувачем.

    content_search - опційно, {"folder", "scope"} (sync.changes_folder - папка й
    режим "усі документи"/"лише ЗАВДАННЯ", обрані користувачем для цього місяця) -
    якщо задано, для КОЖНОЇ людини з prev+actual шукаються персональні згадки в
    ЩОДЕННА/ЗАВДАННЯ файлах цієї папки (content.document_content_search) і
    підставляються як extra_grounds_by_person ЛИШЕ для точки 100
    (_FOLDER_BASED_GROUNDS_POINTS) - решта пунктів цей пошук не бачить узагалі, як
    і раніше; 170 (і 70) сюди окремо не входять, але автоматично охоплені тим
    самим пошуком, бо їхні дні вже рахуються в межах точки 100
    (_COMBINED_TARGET_CELL_VALUES). Для тієї самої точки 100 додатково
    перевіряється ПОКРИТТЯ: для кожного дня фактичної участі (actual-файл) - якщо
    жоден файл ЩОДЕННА/ЗАВДАННЯ, датований САМЕ цим днем, не згадує ПІБ людини -
    готовий рядок попередження ДОДАЄТЬСЯ в missing_coverage_warnings (якщо
    передано - список, що викликач сам збирає й пізніше рендерить у САМ
    WORD-документ, окремим списком у кінці - підтверджено користувачем: "не
    потрібно щоб в термінал, а у файл word виводились відсутні").

    commander_and_tvo_only=False (за замовчуванням, головний рапорт) - прибирає
    з prev/actual РІВНО тих, хто фігурує в рапорті КБ/ТВО цього місяця
    (_exclude_commander_and_tvo, узгоджено з narrator_row - див. нижче): вони
    пишуть/фігурують у СВОЄМУ рапорті, не в головному. True (рапорт КБ/ТВО,
    generate_report_for_commander_money.py) - навпаки, лишає ЛИШЕ їх
    (_keep_only_commander_and_tvo): якщо в командира/ТВО самого є зміни за
    попередній місяць - вони мають потрапити в його ВЛАСНИЙ рапорт, а не
    залишитись невидимими лише тому, що головний рапорт їх не показує -
    підтверджено користувачем.

    narrator_row=None (типово) - хто підписує рапорт КБ/ТВО (find_higher_commander
    на "сьогодні" - справжні виклики з generate_report_for_get_money.py/
    generate_report_for_commander_money.py передають його явно); None
    трактується як {"ТВО": True} - стара, безумовна поведінка ("виключити
    БУДЬ-ЯКОГО, хто взагалі БУВ ТВО хоч частину місяця, незалежно від того, хто
    підписує сьогодні") для викликів (переважно тестів), яким сам нюанс "хто
    підписує" не потрібен."""
    if categories is None:
        categories = MONEY_REPORT_CATEGORIES
    if narrator_row is None:
        narrator_row = {"ТВО": True}
    prev_rows, prev_date_columns = _rows_from_check_file(prev_path)
    actual_rows, actual_date_columns = _rows_from_check_file(actual_path)

    prev_rows = _backfill_missing_labels(prev_rows, actual_rows)
    actual_rows = _backfill_missing_labels(actual_rows, prev_rows)

    commander_filter = _keep_only_commander_and_tvo if commander_and_tvo_only else _exclude_commander_and_tvo
    if prev_date_columns:
        prev_rows = commander_filter(prev_rows, rows_with_tvo_data, prev_date_columns, narrator_row)
        prev_rows = _exclude_by_excluded_day_value(prev_rows, prev_date_columns)
    if actual_date_columns:
        actual_rows = commander_filter(actual_rows, rows_with_tvo_data, actual_date_columns, narrator_row)
        actual_rows = _exclude_by_excluded_day_value(actual_rows, actual_date_columns)

    extra_grounds_by_person = {}
    if content_search:
        roster_names = {normalize_name(row.get("ПІБ", "")) for row in prev_rows + actual_rows}
        extra_grounds_by_person = find_person_document_references(
            content_search["folder"], roster_names, content_search["scope"],
        )
        if actual_date_columns:
            person_dates = {}
            for point in (p for p in _FOLDER_BASED_GROUNDS_POINTS if p in categories):
                target_cell_values = _COMBINED_TARGET_CELL_VALUES.get(point, {point})
                for spec in resolve_reportable_categories(point, categories):
                    dates_by_person = category_dates_by_person(
                        actual_rows, actual_date_columns, target_cell_values, {},
                        status_filter=spec["status_filter"], exclude_status=spec["exclude_status"], categories=categories,
                    )
                    for pib, dates in dates_by_person.items():
                        person_dates.setdefault(pib, set()).update(dates)
            if person_dates and missing_coverage_warnings is not None:
                rank_by_pib = {normalize_name(row.get("ПІБ", "")): row.get("ЗВАННЯ", "") for row in actual_rows}
                subdivision_by_pib = {normalize_name(row.get("ПІБ", "")): row.get("ПІДРОЗДІЛ", "") for row in actual_rows}
                missing_coverage_warnings.extend(
                    collect_missing_document_coverage_warnings(
                        content_search["folder"], person_dates, content_search["scope"],
                        rank_by_pib=rank_by_pib, subdivision_by_pib=subdivision_by_pib,
                    )
                )

    per_point = {}
    for point in (p for p in _CHANGES_POINT_PRIORITY_ORDER if p in categories):
        add_rows = _delta_rows_for_point(
            prev_rows, prev_date_columns, actual_rows, actual_date_columns, point, categories,
            extra_grounds_by_person=extra_grounds_by_person if point in _FOLDER_BASED_GROUNDS_POINTS else None,
        )
        if add_rows:
            per_point[point] = ([], add_rows)

    # BASIS_REQUIRED_POINTS ("100_СПЕЦКОНТИНГЕНТ"/"100_БПШП"/"100_ВПБП") тут
    # БІЛЬШЕ НЕ рендеряться - винесено в _basis_appeared_extra_points_for_month
    # (рендериться "як у звичайному рапорті", SECTIONS у resources/data.json,
    # ПОЗА межами "Прошу внести зміни...", той самий принцип, що й 70/170
    # нижче) - підтверджено користувачем: "ми просто відображаємо секції які в
    # resources/data.json SECTIONS 100_БПШП 100_ВПБП, 100_СПЕЦКОНТИНГЕНТ".

    # appendix_number - ФІКСОВАНА позиція в _CHANGES_POINT_PRIORITY_ORDER
    # (100->1, 50->2, 30->3, ЗАВЖДИ - підтверджено користувачем: відсутність
    # точки цього місяця просто пропускає її номер, а НЕ зсуває решту).
    appendix_pairs = [
        (point, _CHANGES_POINT_PRIORITY_ORDER.index(point) + 1, exclude_rows, add_rows)
        for point, (exclude_rows, add_rows) in per_point.items()
    ]
    return appendix_pairs


def _eligible_month_candidates(changes_file_pairs):
    """Пари prev/actual, чий actual-файл описує місяць, що дійсно передує ОБРАНОМУ
    місяцю генерації (constants.MONTH/YEAR) - відсортовані від найближчого місяця
    до найдальшого. (months_before, year, month, pair) - спільна логіка для
    build_changes_entries (які пари включати в рапорт) і
    sync.changes_folder.describe_eligible_changes_months (за які місяці питати
    папку з підставами), щоб критерій "підходить/не підходить" жив в одному місці."""
    candidates = []
    for pair in changes_file_pairs:
        _, actual_date_columns = _rows_from_check_file(pair["actual_path"])
        year, month = _resolve_file_year_month(actual_date_columns)
        if year is None:
            continue
        months_before = _months_before_selected(year, month)
        if months_before <= 0:
            continue
        candidates.append((months_before, year, month, pair))

    candidates.sort(key=lambda item: item[0])
    return candidates


def describe_eligible_changes_months(changes_file_pairs):
    """(pair, "MM (назва місяця)") для КОЖНОЇ елігібельної пари (_eligible_month_candidates),
    найближчий місяць першим - для запиту "Оберіть папку з документами (підставами)
    за {label}?" по кожному місяцю окремо."""
    return [
        (pair, f"{month:02d} ({month_nominative_upper(month).lower()})")
        for _months_before, _year, month, pair in _eligible_month_candidates(changes_file_pairs)
    ]


def _month_date_range(year, month):
    """Перший і останній день (рік, місяць) як datetime - для зіставлення з
    "start"/"end" MONEY_REPORT_CHANGES_ORDER_REFERENCES (весь амендований місяць,
    а не лише фактично відпрацьовані дні - наказ діє на календарний період)."""
    last_day = calendar.monthrange(year, month)[1]
    return datetime(year, month, 1), datetime(year, month, last_day)


def _resolve_admin_order_reference(period_start, period_end, references=None):
    """Перший рядок "lines" першого запису MONEY_REPORT_CHANGES_ORDER_REFERENCES,
    чий "start"/"end" перетинається з [period_start; period_end] (амендований
    місяць) - той самий формат {"start", "end", "lines"}, що й
    MONEY_REPORT_GENERAL_REFERENCES_BN/_LOG_WAR (готовий текст на кшталт "№1657
    від 08.07.2026", а не окремі число/дата) - для підстановки в
    STATIK["CHANGES_INTRO"]. None, якщо жоден запис не підходить (лишається
    порожнім місцем для ручного заповнення - як і до появи цього механізму)."""
    if references is None:
        references = MONEY_REPORT_CHANGES_ORDER_REFERENCES
    period_start, period_end = to_date(period_start), to_date(period_end)
    for ref in references:
        ref_start, ref_end = to_date(ref["start"]), to_date(ref["end"])
        if ref_start <= period_end and ref_end >= period_start and ref.get("lines"):
            return ref["lines"][0]
    return None


# 70/170 "як у звичайному рапорті" (не Додаток розділу "Прошу внести зміни...",
# а ЗВИЧАЙНИЙ пункт рапорту, SECTIONS у resources/data.json) - підтверджено
# користувачем: "є тільки 1 додаток це 100 та 3 додаток це 30, іншого не дано,
# просто добавляєш пункт поза межами 'прошу внести зміну'". На відміну від
# Додатку 100 (де 70/170-дні рахуються ДОДАТКОВИМ рядком, через
# _COMBINED_TARGET_CELL_VALUES) - тут СВОЯ, НЕ об'єднана участь: лише дні,
# буквально позначені САМЕ цим числом (100 десь в людини - не рахується сюди,
# він уже показаний у Додатку 100).
_EXTRA_STANDALONE_POINTS = (70, 170)


def _extra_point_rows_for_month(prev_path, actual_path, rows_with_tvo_data, categories, commander_and_tvo_only, narrator_row, content_search):
    """Для ОДНІЄЇ пари prev/actual файлів місяця - рядки пунктів
    _EXTRA_STANDALONE_POINTS (70/170), кожен - СВОЯ участь (target_cell_values=
    {point}, БЕЗ об'єднання з 100), точно як для цих самих точок у звичайному
    щомісячному рапорті.

    Гейтинг - ЛИШЕ НОВІ дні (actual мінус prev, той самий принцип, що й для
    "Додатку N доповнити" - _delta_rows_for_point, безпосередньо перевикористаний
    тут) - підтверджено користувачем: "якщо в файлі 170 не змінювалось, то не
    потрібно добавляти цей пункт" - людина, чиї 70/170-дні ІДЕНТИЧНІ між prev і
    actual (жодного НОВОГО дня), НЕ рендериться взагалі, навіть якщо вона й
    далі має ці дні в actual - раніше (до цього виправлення) ці два пункти
    рахувались ЛИШЕ з actual-файлу, без жодного порівняння з prev, тож
    з'являлись КОЖЕН місяць, навіть без жодної реальної зміни.

    ТОЙ САМИЙ, обраний користувачем для цього місяця, content_search (ПІБ-пошук
    у ЩОДЕННА/ЗАВДАННЯ файлах), що й для Додатку 100 - підтверджено
    користувачем. ПОКРИТТЯ документами тут ОКРЕМО не перевіряється -
    build_appendix_pairs_for_month вже перевіряє його для точки 100 з
    ОБ'ЄДНАНИМИ target_cell_values (100/70/170, _COMBINED_TARGET_CELL_VALUES) -
    той самий виклик автоматично охоплює й дати цих двох "власних" пунктів,
    повторна перевірка тут дублювала б попередження.

    Повертає [{"point", "rows", "period_start_day", "period_end_day"}, ...]
    лише для точок, де є хоч один рядок."""
    prev_rows, prev_date_columns = _rows_from_check_file(prev_path)
    actual_rows, actual_date_columns = _rows_from_check_file(actual_path)
    if not prev_date_columns or not actual_date_columns:
        return []
    if categories is None:
        categories = MONEY_REPORT_CATEGORIES
    if narrator_row is None:
        narrator_row = {"ТВО": True}

    prev_rows = _backfill_missing_labels(prev_rows, actual_rows)
    actual_rows = _backfill_missing_labels(actual_rows, prev_rows)

    commander_filter = _keep_only_commander_and_tvo if commander_and_tvo_only else _exclude_commander_and_tvo
    prev_rows = commander_filter(prev_rows, rows_with_tvo_data, prev_date_columns, narrator_row)
    prev_rows = _exclude_by_excluded_day_value(prev_rows, prev_date_columns)
    actual_rows = commander_filter(actual_rows, rows_with_tvo_data, actual_date_columns, narrator_row)
    actual_rows = _exclude_by_excluded_day_value(actual_rows, actual_date_columns)

    extra_grounds_by_person = {}
    if content_search:
        roster_names = {normalize_name(row.get("ПІБ", "")) for row in prev_rows + actual_rows}
        extra_grounds_by_person = find_person_document_references(
            content_search["folder"], roster_names, content_search["scope"],
        )

    sections = []
    for point in (p for p in _EXTRA_STANDALONE_POINTS if p in categories):
        rows = _delta_rows_for_point(
            prev_rows, prev_date_columns, actual_rows, actual_date_columns, point, categories,
            extra_grounds_by_person=extra_grounds_by_person,
        )
        if rows:
            sections.append({
                "point": point, "rows": rows,
                "period_start_day": actual_date_columns[0].day, "period_end_day": actual_date_columns[-1].day,
            })

    return sections


def _new_participation_rows_for_basis_required_point(prev_rows, prev_date_columns, actual_rows, actual_date_columns, point, categories):
    """BASIS_REQUIRED_POINTS - людина БЕЗ жодного дня цієї категорії в prev
    ВЗАГАЛІ (на відміну від _basis_appeared_add_rows, що дивиться на ПОРІВНЯННЯ
    ПІДСТАВИ за ТІ САМІ дні prev - тут дні цієї категорії в prev відсутні
    ЦІЛКОМ), у якої з'явились НОВІ дні в actual - РЕАЛЬНИЙ ВИПАДОК, підтверджений
    користувачем: "170"/"ВЛК" у prev, "100_ШПБП"
    (=100_БПШП, з альтернативним написанням) і "100_ВПБП" - ЛИШЕ в actual.
    _basis_appeared_add_rows тут НЕ спрацьовує - вона будує рядки ЗІ САМОГО
    prev_rows (build_category_rows(prev_rows, ...)), тож людина без жодного дня
    цієї категорії в prev не потрапляє в її результат НІКОЛИ, незалежно від
    підстави.

    Гейтинг і зміст рядка - ТОЙ САМИЙ принцип, що й _delta_rows_for_point для
    звичайних пунктів (ЛИШЕ дельта - нові дні, не весь період - "не треба
    знімати з уже сплаченого, щоб знову щось заплатити" стосується й цих
    пунктів так само). ПІДСТАВА, однак, - зі СВОЄЇ, ВЛАСНОЇ колонки цього
    статусу (BASIS_REQUIRED_POINT_COLUMN_NAMES, constants.py), а НЕ з
    LOG_WAR/BN (_changes_category_config НЕ форсує "general"/"grounds" для
    точок ПОЗА _CHANGES_POINT_PRIORITY_ORDER - лишає порожні "general"/"grounds"
    самої категорії, як і в реальному constants.py, тож build_legal_basis_text
    складається ЛИШЕ з extra_lines). Людина БЕЗ непорожньої ВЛАСНОЇ підстави -
    ВИКЛЮЧАЄТЬСЯ ЦІЛКОМ (той самий "required_basis" принцип, що й у звичайному
    щомісячному рапорті - generate_report_for_get_money._build_categories) - не
    рендериться з порожньою "Підстава", а просто відсутня в результаті."""
    column_names = BASIS_REQUIRED_POINT_COLUMN_NAMES[point]
    actual_extra_grounds = build_pidstavy_extra_grounds_by_person(actual_rows, column_names)

    rows = _delta_rows_for_point(
        prev_rows, prev_date_columns, actual_rows, actual_date_columns, point, categories,
        extra_grounds_by_person=actual_extra_grounds,
    )
    return [row for row in rows if row["ПІДСТАВА"]]


def _basis_appeared_extra_points_for_month(prev_path, actual_path, rows_with_tvo_data, categories, commander_and_tvo_only, narrator_row):
    """BASIS_REQUIRED_POINTS ("100_СПЕЦКОНТИНГЕНТ"/"100_БПШП"/"100_ВПБП",
    constants.py) - рендеряться "як у звичайному рапорті" (SECTIONS у
    resources/data.json - ті самі ключі-рядки, що й для звичайного щомісячного
    рапорту), ПОЗА межами "Прошу внести зміни..." - той самий принцип, що й
    _EXTRA_STANDALONE_POINTS (70/170) вище - підтверджено користувачем: "коли є
    цей статус тоді ми просто відображаємо секції які в resources/data.json
    SECTIONS 100_БПШП 100_ВПБП, 100_СПЕЦКОНТИНГЕНТ" - а НЕ "Додаток N
    доповнити" (ранішу поведінку, коли вони отримували "зайві" номери додатків
    ПІСЛЯ фіксованих 100/50/30, замінено).

    ДВІ РІЗНІ, ДОПОВНЮЮЧІ ОДНА ОДНУ причини появи людини в результаті (обидві
    підтверджені реальними випадками користувача):
    1. _basis_appeared_add_rows - статус (і дні цієї категорії) НЕЗМІННИЙ між
       prev і actual, але сама ПІДСТАВА (текст у ВЛАСНІЙ колонці) з'явилась
       лише в actual - людина вже БУЛА тут увесь час, просто документ підстави
       додали пізніше.
    2. _new_participation_rows_for_basis_required_point - ЗОВСІМ НОВІ дні цієї
       категорії, яких у prev не було ВЗАГАЛІ (людина щойно набула цього
       статусу цього місяця) - реальний випадок, підтверджений користувачем
       (170/ВЛК у prev, "100_ШПБП"/"100_ВПБП" - лише в actual).
    Результати ОБОХ об'єднуються (_merge_preserving_file_order) - людина, що
    (малоймовірно, але теоретично) підходить під ОБИДВА критерії одночасно,
    отримає два окремі рядки, а не один об'єднаний (як і для звичайних
    категорій - build_category_rows/_merge_preserving_file_order).

    prev_rows/actual_rows перечитуються тут ЖЕ (через _rows_from_check_file -
    @lru_cache, тож файл НЕ читається з диска повторно, лише повертається вже
    кешований результат) і фільтруються ТИМ САМИМ шляхом (commander_filter,
    _exclude_by_excluded_day_value), що й у build_appendix_pairs_for_month -
    обидві функції мають бачити ОДНАКОВИЙ набір людей, інакше запис міг би
    з'явитись тут для людини, вже вилученої (напр. ПРВД) з "Прошу внести
    зміни..."/звичайного рапорту.

    period_start_day/period_end_day - з actual_date_columns (а не prev - на
    відміну від попередньої версії: тепер рядки МОЖУТЬ описувати НОВІ,
    actual-дні (сценарій 2 вище), тож заголовок секції має покривати ввесь
    діапазон дат ЦЬОГО місяця, а не лише prev).

    Повертає [{"point", "rows", "period_start_day", "period_end_day"}, ...] у
    ТОМУ САМОМУ форматі, що й _extra_point_rows_for_month - build_changes_entries
    об'єднує обидва в ОДИН спільний "extra_points" список, рендериться ТІЄЮ
    САМОЮ generate_report_for_get_money._add_extra_point_sections без жодних
    змін коду рендеру."""
    prev_rows, prev_date_columns = _rows_from_check_file(prev_path)
    actual_rows, actual_date_columns = _rows_from_check_file(actual_path)
    if not prev_date_columns or not actual_date_columns:
        return []

    prev_rows = _backfill_missing_labels(prev_rows, actual_rows)
    actual_rows = _backfill_missing_labels(actual_rows, prev_rows)

    if categories is None:
        categories = MONEY_REPORT_CATEGORIES
    if narrator_row is None:
        narrator_row = {"ТВО": True}

    commander_filter = _keep_only_commander_and_tvo if commander_and_tvo_only else _exclude_commander_and_tvo
    prev_rows = commander_filter(prev_rows, rows_with_tvo_data, prev_date_columns, narrator_row)
    prev_rows = _exclude_by_excluded_day_value(prev_rows, prev_date_columns)
    actual_rows = commander_filter(actual_rows, rows_with_tvo_data, actual_date_columns, narrator_row)
    actual_rows = _exclude_by_excluded_day_value(actual_rows, actual_date_columns)

    sections = []
    for point in (p for p in BASIS_REQUIRED_POINTS if p in categories):
        retroactive_rows = _basis_appeared_add_rows(prev_rows, prev_date_columns, actual_rows, actual_date_columns, point, categories)
        new_rows = _new_participation_rows_for_basis_required_point(prev_rows, prev_date_columns, actual_rows, actual_date_columns, point, categories)
        rows = _merge_preserving_file_order(actual_rows, retroactive_rows, new_rows)
        # "100_СПЕЦКОНТИНГЕНТ" - той самий вимог, що й у звичайному щомісячному
        # рапорті (generate_report_for_get_money._build_categories): без дати
        # зникнення (ОБЛІК.xlsx) людина в цю таблицю не потрапляє взагалі,
        # навіть ретроактивно - функція сама пропускає решту пунктів без змін.
        rows = exclude_rows_without_disappearance_date(rows, point)
        if rows:
            sections.append({
                "point": point, "rows": rows,
                "period_start_day": actual_date_columns[0].day, "period_end_day": actual_date_columns[-1].day,
            })

    return sections


def build_changes_entries(rows_with_tvo_data, categories=None, changes_file_pairs=(), commander_and_tvo_only=False, narrator_row=None):
    """Повертає список записів "Прошу внести зміни..." - один на кожну елігібельну
    пару prev/actual з changes_file_pairs (_eligible_month_candidates), найближчий
    до обраного місяць першим (послідовна генерація по місяцях, від обраного до
    попередніх). Пара без жодної фактичної зміни (усі пункти незмінні) НЕ потрапляє
    в результат. "order_reference" - з MONEY_REPORT_CHANGES_ORDER_REFERENCES
    (None, якщо для цього місяця немає відповідного запису).

    commander_and_tvo_only, narrator_row - див. build_appendix_pairs_for_month.

    "extra_points" - об'єднання _extra_point_rows_for_month (70/170, ЛИШЕ НОВІ
    дні, як і "Додаток N доповнити" - "як у звичайному рапорті", але не без
    порівняння з prev) і _basis_appeared_extra_points_for_month
    (BASIS_REQUIRED_POINTS - "100_СПЕЦКОНТИНГЕНТ"/"100_БПШП"/"100_ВПБП",
    ретроактивна підстава - той самий принцип рендеру, підтверджено
    користувачем), в ОДИН список, ПОЗА межами "Прошу внести зміни..."; "month" -
    номер місяця (для рендеру їхнього заголовка - month_genitive_lower). Запис
    потрапляє в результат, якщо є ХОЧ ОДНЕ з двох (appendix_pairs чи
    extra_points) - навіть якщо appendix_pairs порожній, самі ці пункти однаково
    мають бути показані.

    "missing_coverage_warnings" - готові рядки попередження (content.document_content_search.
    collect_missing_document_coverage_warnings, через build_appendix_pairs_for_month)
    про дні точки 100 (і 70/170, автоматично охоплені - _COMBINED_TARGET_CELL_VALUES)
    БЕЗ жодного файлу-підстави в обраній папці - НЕ друкуються в термінал,
    підтверджено користувачем: рендеряться окремим списком у САМОМУ WORD-документі
    (generators.generate_report_for_get_money._add_missing_coverage_section)."""
    entries = []
    for _months_before, year, month, pair in _eligible_month_candidates(changes_file_pairs):
        content_search = None
        if pair.get("content_search_folder"):
            content_search = {"folder": pair["content_search_folder"], "scope": pair.get("content_search_scope")}
        missing_coverage_warnings = []
        appendix_pairs = build_appendix_pairs_for_month(
            pair["prev_path"], pair["actual_path"], rows_with_tvo_data, categories, content_search,
            commander_and_tvo_only=commander_and_tvo_only, narrator_row=narrator_row,
            missing_coverage_warnings=missing_coverage_warnings,
        )
        extra_points = _extra_point_rows_for_month(
            pair["prev_path"], pair["actual_path"], rows_with_tvo_data, categories, commander_and_tvo_only, narrator_row, content_search,
        ) + _basis_appeared_extra_points_for_month(
            pair["prev_path"], pair["actual_path"], rows_with_tvo_data, categories, commander_and_tvo_only, narrator_row,
        )
        if appendix_pairs or extra_points:
            period_start, period_end = _month_date_range(year, month)
            order_reference = _resolve_admin_order_reference(period_start, period_end)
            entries.append({
                "appendix_pairs": appendix_pairs, "order_reference": order_reference,
                "extra_points": extra_points, "month": month,
                "missing_coverage_warnings": missing_coverage_warnings,
            })
    return entries
