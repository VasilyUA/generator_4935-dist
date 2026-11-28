"""Книга налаштувань generator_br_and_report_for_money (resources/НАЛАШТУВАННЯ.xlsx).

Усе, що користувач налаштовує під свій батальйон і місяць (назви, номери БР,
підстави, пункти й категорії рапорту, тексти документів), - у цій книзі, а не
в constants.py чи data.json (чому - див. utils/settings_book.py).

- read_settings: книга -> словник налаштувань (аркуш-таблиця - список рядків);
- build_constants/build_number_dicts: з цього словника - ТІ САМІ структури,
  що колись жили в constants.py (NUMBER_OF_DOCUMENTS_BRS_*,
  MONEY_REPORT_CATEGORIES, BR_HIGHT_UNIT ...), тож решта коду працює з ними як
  і раніше;
- write_settings: словник -> нова книга (зразок, перенесення колишніх
  constants.py/data.json - utils/settings_migration.py);
- write_number_fields/write_numbers_workbook/read_numbers_workbook: номери
  документів, які програма сама зчитує з папки з документами, - назад у книгу
  (колись - прямо в рядки constants.py) і в окрему книгу папки етапу.

Словник налаштувань не залежить від того, як дані розкладено по аркушах:
номери трьох видів документів - один аркуш, а в словнику - три таблиці;
підстави - чотири таблиці (ЖБД, БН, вищого штабу, інші), а в словнику - один
список; пункт і його категорії - один блок таблиці, а в словнику - "points" і
"categories".

Аркушів у книзі кілька, на більшості - кілька таблиць одна під одною, кожна
зі смугою "■ назва" (settings_book.write_sections). Назви таблиць - колишні
назви аркушів (*_SHEET нижче), тож книга, де кожна таблиця мала власний аркуш,
читається так само, а при першому запуску один раз переписується в теперішній
вигляд (_upgrade_layout).

Без залежностей від constants.py: його імпортує сам constants.py."""
import calendar
import os
import re
import shutil
from collections import Counter
from dataclasses import replace
from types import SimpleNamespace

from utils import settings_book as sb
from utils.logging_utils import print_green, print_red

# Аркуші книги.
NUMBERS_SHEET = "Номери документів"  # аркуш з однією таблицею (його пише й зчитування папки з документами)
REFERENCES_BOOK_SHEET = "Підстави"
POINTS_BOOK_SHEET = "Пункти рапорту"
TEXTS_BOOK_SHEET = "Тексти рапорту"
TASKS_BOOK_SHEET = "БР і ЗАВДАННЯ"
GENERAL_BOOK_SHEET = "Загальні"

# Таблиці (смуга "■ назва" на своєму аркуші - PLACE).
GENERAL_SHEET = "Загальні"
LOG_WAR_SHEET = "ЖБД"
BN_SHEET = "БН"
HIGHER_REFERENCES_SHEET = "Підстави вищого штабу"
OTHER_REFERENCES_SHEET = "Інші підстави"
CHANGES_SHEET = "Накази виплат"
POINTS_SHEET = "Рапорт на особовий склад"
COMMANDER_POINTS_SHEET = "Рапорт командира/ТВО"
NOT_PAID_EXCEPTIONS_SHEET = "Винятки невиплати"
SECTION_TEXTS_SHEET = "Тексти пунктів"
REPORT_TEXTS_SHEET = "Загальні тексти рапорту"
SUBDIVISIONS_SHEET = "Підрозділи"
SUBDIVISION_TEXTS_SHEET = "Тексти підрозділів"
TASK_TEXTS_SHEET = "Тексти ЗАВДАННЯ"
TASK_VARIANTS_SHEET = "Варіанти ЗАВДАННЯ"

# Таблиця -> аркуш, на якому вона лежить.
PLACE = {
    LOG_WAR_SHEET: REFERENCES_BOOK_SHEET,
    BN_SHEET: REFERENCES_BOOK_SHEET,
    HIGHER_REFERENCES_SHEET: REFERENCES_BOOK_SHEET,
    OTHER_REFERENCES_SHEET: REFERENCES_BOOK_SHEET,
    CHANGES_SHEET: REFERENCES_BOOK_SHEET,
    POINTS_SHEET: POINTS_BOOK_SHEET,
    COMMANDER_POINTS_SHEET: POINTS_BOOK_SHEET,
    NOT_PAID_EXCEPTIONS_SHEET: POINTS_BOOK_SHEET,
    SECTION_TEXTS_SHEET: TEXTS_BOOK_SHEET,
    REPORT_TEXTS_SHEET: TEXTS_BOOK_SHEET,
    SUBDIVISION_TEXTS_SHEET: TASKS_BOOK_SHEET,
    TASK_TEXTS_SHEET: TASKS_BOOK_SHEET,
    TASK_VARIANTS_SHEET: TASKS_BOOK_SHEET,
    GENERAL_SHEET: GENERAL_BOOK_SHEET,
    SUBDIVISIONS_SHEET: GENERAL_BOOK_SHEET,
}
# Таблиці, чий колишній окремий аркуш звався інакше, ніж таблиця тепер.
LEGACY_SHEETS = {
    POINTS_SHEET: "Пункти рапорту",
    COMMANDER_POINTS_SHEET: "Пункти рапорту командира",
    REPORT_TEXTS_SHEET: "Тексти рапорту",
}


def _where(table):
    """(аркуш, таблиця чи None, колишній аркуш) - де в книзі таблиця table (для
    sb.read_rows); аркуш з однією таблицею (номери документів) - (аркуш, None, None)."""
    if table in PLACE:
        return PLACE[table], table, LEGACY_SHEETS.get(table)
    return table, None, None


def _label(table):
    """'Пункти рапорту", таблиця "Рапорт на особовий склад' - для sb.location."""
    return sb.table_label(PLACE[table], table) if table in PLACE else table

# Особливі групи підстав: ЖБД і БН батальйону код використовує й окремо від
# пунктів рапорту (шапка витягу з ЖБД, рапорт змін за минулі місяці), тож у
# кожної - власна таблиця.
LOG_WAR_GROUP = "ЖБД"
BN_GROUP = "БН"

UNIT_KEYS = ("SHORT_UNIT_BATTALION", "SHORT_UNIT_BRIGADE", "FULL_UNIT_BUT", "FULL_MILITARY_UNIT")

GENERAL_SETTINGS = (
    sb.Setting("SHORT_UNIT_BATTALION", label="Батальйон (скорочено)", note='Напр. "1 бмп" - у текстах БР, рапортів, ЖБД.'),
    sb.Setting("SHORT_UNIT_BRIGADE", label="Бригада (скорочено)", note='Напр. "40 обрмп".'),
    sb.Setting("FULL_UNIT_BUT", label="Батальйон (повна назва, родовий відмінок)", note='Напр. "1 батальйону морської піхоти".'),
    sb.Setting("FULL_MILITARY_UNIT", label="Військова частина", note='"військової частини А0000".'),
    sb.Setting("HIGHER_COMMANDER_TITLE", label="Посада командира батальйону",
               note="Точно так, як у колонці ПОСАДА ОБЛІК.xlsx: від його імені пишеться рапорт, він підписує БР."),
    sb.Setting("COMMANDER_TITLE", label="Посада начальника штабу",
               note="Точно так, як у колонці ПОСАДА ОБЛІК.xlsx (другий підпис БР)."),
    sb.Setting("SEQUENCE_HIGHER_COMMANDER", sb.LIST, required=False, label="Адресати БР (заступники)",
               note='Скорочення посад заступників через кому, у цьому порядку - рядок після "КОМАНДИРУ ..." у БР.'),
)

# ---------------------------------------------------------------- номери документів

DAILY = "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY"
WEEKLY = "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK"
SAFETY = "NUMBER_OF_DOCUMENTS_BRS_SAVE"

DAY_COLUMN = sb.Column(
    "day", "День", sb.DAY, nonempty=True, width=13,
    note='Число місяця 1-31: номери підставляються в місяць, обраний при запуску. Рядок "попередній" - останній '
         "день попереднього місяця. Дні, яких немає в обраному місяці (напр. 31 у листопаді), пропускаються.",
)

_DAILY_GROUP = "Щоденні БР"
_WEEKLY_GROUP = "ЗАВДАННЯ (тижневі)"
_SAFETY_GROUP = "Безпека (БЗ)"

# Ім'я словника в constants.py -> ((поле запису словника, колонка аркуша
# "Номери документів"), ...). Ключі колонок різні (поле "бат" є і в щоденних
# БР, і в ЗАВДАННЯ), заголовки - ті самі, що й на колишніх окремих аркушах.
NUMBER_TABLES = {
    DAILY: (
        ("бат", sb.Column("daily_bat", "БР бат №", sb.NUMBER_TEXT, width=11, center=True, group=_DAILY_GROUP,
                          note="Номер щоденного бойового розпорядження батальйону (ЩОДЕННА).")),
        ("брг", sb.Column("daily_brg", "БР брг №", sb.NUMBER_TEXT, width=11, center=True, group=_DAILY_GROUP,
                          note="Номер бригадного розпорядження, що продовжує оборону, на цей день.")),
        ("general_br_bat", sb.Column("daily_general", "Загальне БР бат", width=21, group=_DAILY_GROUP,
                                     note='Чинне тижневе ЗАВДАННЯ батальйону: "№ від дд.мм.рррр", напр. "349 від '
                                          '01.10.2026". Під час зчитування папки з документами рахується саме.')),
    ),
    WEEKLY: (
        ("бат", sb.Column("weekly_bat", "ЗАВДАННЯ бат №", sb.NUMBER_TEXT, width=12, center=True, group=_WEEKLY_GROUP,
                          note="Номер тижневого бойового розпорядження (ЗАВДАННЯ), виданого цього дня.")),
        ("посилання_брг", sb.Column("weekly_ref_brg", "Посилання на БР брг", width=21, group=_WEEKLY_GROUP,
                                    note='"№ від дд.мм.рррр", напр. "3071 від 19.08.2026".')),
        ("посилання_бат", sb.Column("weekly_ref_bat", "Посилання на БР бат", width=21, group=_WEEKLY_GROUP,
                                    note='"№ від дд.мм.рррр".')),
    ),
    SAFETY: (
        ("бз_бат", sb.Column("safety_bat", "БЗ бат №", sb.NUMBER_TEXT, width=11, center=True, group=_SAFETY_GROUP,
                             note="Номер розпорядження батальйону з безпеки застосування військ.")),
        ("бз_брг", sb.Column("safety_brg", "БЗ брг №", sb.NUMBER_TEXT, width=11, center=True, group=_SAFETY_GROUP,
                             note="Номер розпорядження бригади з безпеки застосування військ.")),
    ),
}
NUMBER_COLUMNS = (DAY_COLUMN, *(column for pairs in NUMBER_TABLES.values() for _field, column in pairs))

# Колишні окремі аркуші номерів: книга номерів папки етапу, записана до появи
# спільного аркуша, читається й так (ті самі заголовки колонок).
LEGACY_NUMBER_SHEETS = {DAILY: "Номери БР щоденні", WEEKLY: "Номери ЗАВДАННЯ", SAFETY: "Номери БЗ"}

# ---------------------------------------------------------------- підстави

_REFERENCE_SHEETS_TEXT = (
    f'таблиці "{LOG_WAR_SHEET}", "{BN_SHEET}", "{HIGHER_REFERENCES_SHEET}", "{OTHER_REFERENCES_SHEET}" '
    f'аркуша "{REFERENCES_BOOK_SHEET}"'
)

# Таблиці аркуша "Підстави" мають спільні колонки аркуша, тож порядок колонок
# у них однаковий: Початок, Кінець, Текст, а далі - Група чи ID.
_DATE_COLUMNS = (
    sb.Column("start", "Початок", sb.DATE, nonempty=True, width=13),
    sb.Column("end", "Кінець", sb.DATE, width=13, note="Порожньо - діє досі (по сьогодні)."),
    sb.Column("lines", "Текст підстави", sb.LINES, nonempty=True, width=72,
              note="Текст підстави так, як він піде в рапорт. Кілька рядків в одній клітинці (Alt+Enter) - кілька підстав."),
)
_ID_COLUMN = sb.Column(
    "id", "ID", required=False, width=10,
    note=f'Необов\'язково. Позначка запису для колонки "Без загальних підстав" таблиць аркуша "{POINTS_BOOK_SHEET}".',
)
_OWN_REFERENCE_COLUMNS = (*_DATE_COLUMNS, _ID_COLUMN)
GROUP_REFERENCE_COLUMNS = (
    *_DATE_COLUMNS,
    sb.Column("group", "Група", nonempty=True, width=22,
              note=f'Назва групи підстав - на неї посилається колонка "Підстави" таблиць аркуша "{POINTS_BOOK_SHEET}". '
                   "Рядки з однаковою групою - записи однієї групи."),
    _ID_COLUMN,
)

# (аркуш, група - None: з колонки "Група", "Вищий штаб").
REFERENCE_SHEETS = (
    (LOG_WAR_SHEET, LOG_WAR_GROUP, False),
    (BN_SHEET, BN_GROUP, False),
    (HIGHER_REFERENCES_SHEET, None, True),
    (OTHER_REFERENCES_SHEET, None, False),
)

CHANGES_COLUMNS = (
    sb.Column("start", "Початок", sb.DATE, nonempty=True, width=13, note="Перший день місяця, за який виплачено цим наказом."),
    sb.Column("end", "Кінець", sb.DATE, width=13, note="Останній день того місяця."),
    sb.Column("lines", "Наказ", sb.LINES, nonempty=True, width=40,
              note='Напр. "№1319 від 07.06.2026" - підставляється в "Прошу внести зміни в наказ ..." за цей місяць.'),
    sb.Column("id", "ID", required=False, width=10),
)

# ---------------------------------------------------------------- пункти й категорії рапорту

# Рядок пункту (у колонці "Категорія"): його "Підстави" й "Колонка підстав"
# діють для всіх категорій пункту.
POINT_ROW_MARKER = "весь пункт"

_POINT_NOTE = (
    "Пункт рапорту: категорія виплати числом (10, 30, 50, 70, 100, 170 - як у колонках дат ОБЛІК.xlsx) "
    f'або назвою (напр. NOT_PAID, 100_ВПБП). Той самий ключ - у таблиці "{SECTION_TEXTS_SHEET}" аркуша "{TEXTS_BOOK_SHEET}". '
    "У рядках категорій можна не повторювати - береться пункт з рядка вище."
)
_POINT_GROUP = "Пункт і категорія"
_GROUNDS_GROUP = "Підстави для виплати"
_BEHAVIOUR_GROUP = "Як обробляти"

POINT_BLOCK_COLUMNS = (
    sb.Column("point", "Пункт", sb.NUMBER_TEXT, width=23, group=_POINT_GROUP, note=_POINT_NOTE),
    sb.Column("category", "Категорія", width=30, wrap=True, group=_POINT_GROUP,
              note=f'"{POINT_ROW_MARKER}" - рядок самого пункту: його підстави діють для всіх категорій. Інакше - статус '
                   'з ОБЛІК.xlsx (колонка СТАТУС), що належить до цієї категорії; кілька написань - через "|", напр. '
                   "Задув.|задув|задут. Перша категорія, під яку підпадає статус, - його категорія. Статуси табеля - "
                   "кодами: перейменування статусу в табелі (аркуш «Статуси» його книги) підхоплюється само."),
    sb.Column("grounds", "Підстави", sb.LIST, width=24, wrap=True, group=_GROUNDS_GROUP,
              note=f"Групи підстав через кому ({_REFERENCE_SHEETS_TEXT}). У рядку "
                   f'"{POINT_ROW_MARKER}" - для всіх категорій пункту (напр. ЖБД, БН), у рядку категорії - лише для неї.'),
    sb.Column("column", "Колонка підстав в ОБЛІК", sb.LIST, width=22, wrap=True, group=_GROUNDS_GROUP,
              note="Колонка ОБЛІК.xlsx з підставами (напр. ПІДСТАВИ 10К). У рядку категорії порожньо - колонка "
                   f'з рядка "{POINT_ROW_MARKER}" (якщо є).'),
    sb.Column("exclude_general", "Без загальних підстав", width=12, center=True, choices=("всі",), group=_GROUNDS_GROUP,
              note=f'Порожньо - підстави рядка "{POINT_ROW_MARKER}" діють і для цієї категорії; "всі" - жодна з них; '
                   "або ID записів підстав через кому - крім цих."),
    sb.Column("use_brs", "Додати ЗАВДАННЯ", sb.BOOL, width=12, group=_GROUNDS_GROUP,
              note=f'Так - додавати в підставу тижневі ЗАВДАННЯ з посиланнями (аркуш "{NUMBERS_SHEET}"). Порожньо - Ні.'),
    sb.Column("use_brs_from_selected_folder", "Додати щоденні БР", sb.BOOL, width=12, group=_GROUNDS_GROUP,
              note=f'Так - додавати в підставу щоденні БР (аркуш "{NUMBERS_SHEET}"), зчитані з папки з документами. '
                   "Порожньо - Ні."),
    sb.Column("required_basis", "Підстава обов'язкова", sb.BOOL, width=13, group=_GROUNDS_GROUP,
              note='Ні - порожня "Підстава для виплати" для цієї категорії не помилка (напр. пункт 10). Порожньо - Так.'),
    sb.Column("default", "За замовчуванням", sb.BOOL, width=17, group=_BEHAVIOUR_GROUP,
              note="Так - категорія для людей пункту, чий статус не підпадає під жодну іншу категорію пункту. "
                   "Порожньо - Ні."),
    sb.Column("generate_br", "Генерувати БР", sb.BOOL, width=12, group=_BEHAVIOUR_GROUP,
              note="Ні - людям цієї категорії БР не генерується (напр. дії на території противника). Порожньо - Так."),
    sb.Column("include_to_report", "Включати в рапорт", sb.BOOL, width=12, group=_BEHAVIOUR_GROUP,
              note="Ні - категорія не потрапляє в рапорт. Порожньо - Так."),
)

# Прапорці категорії та їхні значення для порожньої клітинки.
_CATEGORY_FLAGS = {
    "use_brs": False,
    "use_brs_from_selected_folder": False,
    "required_basis": True,
    "default": False,
    "generate_br": True,
    "include_to_report": True,
}
# Колонки, які в рядку пункту не діють (там вони мають бути порожні).
_CATEGORY_ONLY_KEYS = ("exclude_general", *_CATEGORY_FLAGS)

SECTION_TEXT_COLUMNS = (
    sb.Column("point", "Пункт", sb.NUMBER_TEXT, nonempty=True, width=22, note=_POINT_NOTE),
    sb.Column("report", "Рапорт", width=80, wrap=True,
              note="Текст пункту в рапорті на особовий склад. Написи у фігурних дужках ({start_date_day}, {manth}, {year} ...) програма заміняє сама - не чіпайте їх."),
    sb.Column("commander", "Рапорт командира", width=80, wrap=True, note="Текст того самого пункту в рапорті командира/ТВО."),
)

# Статус невиплати не враховується, якщо в місяці в людини є й "підтверджуючий"
# статус. Початкове значення (ним заповнюється аркуш, коли його ще немає;
# далі - лише книга): реальний випадок, підтверджений користувачем, - "СЗЧ"
# ігнорується, якщо людину ОФІЦІЙНО оголошено в розшуку ("РОЗП"): сам цей факт
# уже достатня підстава. Порядок появи в датах не має значення (СЗЧ, потім РОЗП
# без повернення; СЗЧ, РОЗП і знову СЗЧ на межі місяця).
DEFAULT_NOT_PAID_EXCEPTIONS = ({"status": "СЗЧ", "confirmed_by": ["РОЗП"]},)

NOT_PAID_EXCEPTION_COLUMNS = (
    sb.Column("status", "Статус невиплати", nonempty=True, width=22,
              note="Статус з ОБЛІК.xlsx (як у категоріях пункту NOT_PAID), напр. СЗЧ."),
    sb.Column("confirmed_by", "Не враховувати, якщо в місяці є", sb.LIST, nonempty=True, width=34,
              note="Статуси через кому: якщо хоч один з них є в людини протягом місяця, статус ліворуч для неї не "
                   "враховується. Регістр не важить."),
)


# ---------------------------------------------------------------- підрозділи й тексти БР

# Перелік підрозділів і абзаци їхніх завдань - окремі аркуші: порядок абзаців
# у БР не обов'язково той самий, що й порядок у рядку "КОМАНДИРУ ...", і абзаци
# бувають і для тих, кого в переліку немає (гбс, говп, "упр, штаб").
SUBDIVISION_COLUMNS = (
    sb.Column("name", "Підрозділ", nonempty=True, width=26,
              note='Як у колонці ПІДРОЗДІЛ ОБЛІК.xlsx (напр. "1 рмп"), у тому порядку, у якому підрозділи стоять у рядку "КОМАНДИРУ ..." БР.'),
)

SUBDIVISION_TEXT_COLUMNS = (
    sb.Column("name", "Підрозділ", nonempty=True, width=22,
              note='Як у колонці ПІДРОЗДІЛ ОБЛІК.xlsx (напр. "1 рмп"). Кілька підрозділів зі спільним абзацом - через кому (напр. "упр, штаб"). '
                   "Порядок рядків - порядок абзаців у БР."),
    sb.Column("br_text", "Текст у БР", width=75, wrap=True,
              note="Абзац завдання підрозділу в щоденному БР. {battalion}, {unit}, {city}, {coordinates} програма заміняє сама. Порожньо - абзацу немає."),
    sb.Column("extract_text", "Текст у витязі з ЖБД", width=75, wrap=True, note="Той самий абзац у минулому часі - для витягу з ЖБД."),
)

TASK_VARIANT_COLUMNS = (
    sb.Column("variant", "Варіант", sb.INT, nonempty=True, width=10,
              note="Номер варіанта переліку завдань. Варіанти чергуються по одному на кожне БР бат. Рядок лише з номером - порожній варіант."),
    sb.Column("title", "Підрозділ", width=40, wrap=True, note='Заголовок завдання, напр. "1 рмп (оборона РОП «Назва»)".'),
    sb.Column("text", "Завдання", width=75, wrap=True),
)

# Відомі тексти таблиць "Загальні тексти рапорту"/"Тексти ЗАВДАННЯ" - лише
# порядок, назви й пояснення для запису книги. Читаються ці таблиці цілком
# (sb.read_texts): набір параметрів довільний, кожен текст бере вже той код,
# що його друкує.
REPORT_TEXT_SETTINGS = (
    sb.Setting("REPORT_TITLE", required=False, label="Заголовок рапорту", note="Напр. РАПОРТ."),
    sb.Setting("HIGH_COMMANDER", required=False, label="Кому адресовано", note="Праворуч угорі."),
    sb.Setting("START_SECTION", required=False, label="Вступ рапорту", note="Вступний абзац рапорту на особовий склад."),
    sb.Setting("SECTION_INFORMATION", required=False, label="Прохання про включення до наказу",
               note="Підставляються {month}, {year}."),
    sb.Setting("COMMANDER_INTRO", required=False, label="Вступ рапорту командира/ТВО",
               note="Підставляються {rank}, {pib}, {posada}, {month}, {year}."),
    sb.Setting("CHANGES_INTRO", required=False, label="Вступ рапорту змін",
               note='"Прошу внести зміни в наказ ... {order_reference}" - рапорт змін за минулі місяці.'),
    sb.Setting("CHANGES_EXCLUDE_LABEL", required=False, label="Підзаголовок «виключити»",
               note="Виключення пунктів додатка {appendix_number}. Зараз у рапорті змін не друкується - лишено на майбутнє."),
    sb.Setting("CHANGES_ADD_LABEL", required=False, label="Підзаголовок «доповнити»",
               note="Доповнення додатка {appendix_number}."),
    sb.Setting("CHANGES_RESTATE_LABEL", required=False, label="Підзаголовок «нова редакція»",
               note="Пункти додатка {appendix_number}, викладені в новій редакції."),
    sb.Setting("MISSING_COVERAGE_HEADING", required=False, label="Заголовок «відсутні підстави»",
               note="Заголовок переліку відсутніх підстав у рапорті змін."),
)
TASK_TEXT_SETTINGS = (
    sb.Setting("BR_ENEMY_ASSESSMENT_TEXT", required=False, label="Оцінка противника", note="Підставляється {brigade}."),
    sb.Setting("BR_FIRE_SUPPORT_INTRO_TEMPLATE", required=False, label="Вогнева підтримка"),
    sb.Setting("BR_TASK_SECTION_TEMPLATE", required=False, label="Завдання батальйону",
               note="Підставляються {brigade}, {brg_ref}, {unit}, {bat_ref}."),
    sb.Setting("BR_SITUATION_UPDATE_TEMPLATE", required=False, label="Зосередження основних зусиль"),
    sb.Setting("BR_DEFENSE_TASKS_TEXT", required=False, label="Завдання з оборони"),
    sb.Setting("BR_DEFENSE_MEASURES_TEXT", required=False, label="Заходи з оборони"),
)

# ---------------------------------------------------------------- оформлення аркушів

_EDIT_HINT = "Редагуйте жовті клітинки «Значення». Колонку «Параметр» не змінюйте: за нею програма знаходить рядок."
_LINES_HINT = "«Кінець» порожньо - діє досі; кілька підстав в одній клітинці - з нового рядка (Alt+Enter)."
_BRACES_HINT = "Написи у фігурних дужках програма заміняє сама - не чіпайте їх."
_STATUS_CODES_HINT = "Статуси табеля тут - коди: перейменування статусу в табелі (аркуш «Статуси» його книги) підхоплюється само."

NUMBERS_TITLE = "Номери документів - за днями обраного місяця"
NUMBERS_HINT = (
    "Рядок - день місяця, обраного при запуску («попередній» - останній день попереднього місяця). Програма "
    "заповнює номери сама, якщо при запуску відповісти «Так» на «Зчитати обрану папку з документами?» (книга має "
    "бути закрита в Excel). Вписувати вручну теж можна."
)

# Підказка під смугою кожної таблиці (перше речення - ще й опис таблиці на
# аркуші "Інструкція").
TABLE_HINTS = {
    LOG_WAR_SHEET: f"Журнали бойових дій батальйону: журнал - рядок з періодом дії. {_LINES_HINT} Група «{LOG_WAR_GROUP}» "
                   "у пунктах рапорту; також - шапка витягу з ЖБД.",
    BN_SHEET: f"Бойові накази батальйону: наказ - рядок з періодом дії. {_LINES_HINT} Група «{BN_GROUP}» у пунктах рапорту.",
    HIGHER_REFERENCES_SHEET: "Документи бригади, корпусу: рядок - документ і група, до якої він належить (на назву групи "
                             "посилаються пункти рапорту). Звірка рапорту з файлами ці документи не перевіряє. " + _LINES_HINT,
    OTHER_REFERENCES_SHEET: "Групи підстав для окремих категорій (РТГр, ЗРАДН ...): рядок - документ і його група (на назву "
                            "групи посилаються пункти рапорту). " + _LINES_HINT,
    CHANGES_SHEET: "Накази, якими вже виплачено за минулі місяці (перший і останній день місяця), - для «Прошу внести "
                   "зміни в наказ ...» у рапорті змін.",
    POINTS_SHEET: f"Пункти й категорії рапорту на особовий склад. Кожен пункт - помаранчевий рядок «{POINT_ROW_MARKER}» (його "
                  "«Підстави» діють для всіх категорій пункту) і під ним категорії - статуси з колонки СТАТУС ОБЛІК.xlsx. "
                  f"Нова категорія - новий рядок під своїм пунктом. Групи підстав - з аркуша «{REFERENCES_BOOK_SHEET}». "
                  + _STATUS_CODES_HINT,
    COMMANDER_POINTS_SHEET: f"Пункти й категорії рапорту командира/ТВО. Налаштовуються окремо, так само: рядок "
                            f"«{POINT_ROW_MARKER}» і під ним категорії пункту.",
    NOT_PAID_EXCEPTIONS_SHEET: "Коли статус невиплати не враховувати. Статус невиплати (пункт NOT_PAID) не враховується для "
                               "людини, якщо протягом місяця в неї є й один зі статусів праворуч (напр. СЗЧ, коли людину "
                               "офіційно оголошено в розшук - РОЗП). Порядок у датах не важить. " + _STATUS_CODES_HINT,
    SECTION_TEXTS_SHEET: "Текст кожного пункту в рапорті на особовий склад і в рапорті командира. Пункт - так само, як у "
                         f"таблицях аркуша «{POINTS_BOOK_SHEET}». " + _BRACES_HINT,
    REPORT_TEXTS_SHEET: f"Шапка, вступ, прохання й підзаголовки рапорту. {_EDIT_HINT} {_BRACES_HINT}",
    SUBDIVISION_TEXTS_SHEET: "Абзаци завдань підрозділів у щоденному БР і ті самі абзаци в минулому часі - для витягу з ЖБД. "
                             "Порядок рядків - порядок абзаців у БР. {battalion}, {unit}, {city}, {coordinates} програма "
                             "заміняє сама.",
    TASK_TEXTS_SHEET: f"Тексти тижневого бойового розпорядження (ЗАВДАННЯ). {_EDIT_HINT} {_BRACES_HINT}",
    TASK_VARIANTS_SHEET: "Варіанти переліку завдань у ЗАВДАННЯ. Варіанти чергуються по одному на кожне БР бат. Рядок - "
                         "завдання підрозділу у варіанті з цим номером; рядок лише з номером - порожній варіант.",
    GENERAL_SHEET: f"Назви батальйону, посади командира й начальника штабу, адресати БР. {_EDIT_HINT}",
    SUBDIVISIONS_SHEET: "Підрозділи батальйону: як у колонці ПІДРОЗДІЛ ОБЛІК.xlsx (напр. «1 рмп»), у тому порядку, у якому "
                        "вони стоять у рядку «КОМАНДИРУ ...» БР.",
}

_ROWS_HINT = (
    "Новий запис - у порожній рядок під таблицею; бракує рядків - вставте рядок усередині таблиці (правою кнопкою на "
    "номері рядка -> «Вставити»). Назви таблиць у смугах «■» і заголовки колонок не змінюйте."
)
# (аркуш, назва над таблицями, підказка, колір ярлика, таблиці згори донизу,
# скільки колонок зліва не гортаються) - аркуші з кількома таблицями.
BOOK_SHEETS = (
    (REFERENCES_BOOK_SHEET, "Підстави - документи, на які посилається рапорт",
     "ЖБД і БН батальйону, документи вищого штабу, групи підстав для окремих категорій і накази виплат за минулі "
     "місяці. " + _ROWS_HINT, sb.TAB_REFERENCES,
     (LOG_WAR_SHEET, BN_SHEET, HIGHER_REFERENCES_SHEET, OTHER_REFERENCES_SHEET, CHANGES_SHEET), 0),
    (POINTS_BOOK_SHEET, "Пункти рапорту - категорії виплат і їхні підстави",
     "Пункти й категорії обох рапортів і винятки невиплати. " + _ROWS_HINT, sb.TAB_REPORT,
     (POINTS_SHEET, COMMANDER_POINTS_SHEET, NOT_PAID_EXCEPTIONS_SHEET), 2),
    (TEXTS_BOOK_SHEET, "Тексти рапорту",
     "Тексти пунктів і решта текстів рапорту. Налаштовується рідко. " + _ROWS_HINT, sb.TAB_TEXTS,
     (SECTION_TEXTS_SHEET, REPORT_TEXTS_SHEET), 0),
    (TASKS_BOOK_SHEET, "БР і ЗАВДАННЯ - тексти бойових розпоряджень",
     "Абзаци завдань підрозділів у БР, тексти й варіанти тижневого ЗАВДАННЯ. Налаштовується рідко. " + _ROWS_HINT,
     sb.TAB_TEXTS, (SUBDIVISION_TEXTS_SHEET, TASK_TEXTS_SHEET, TASK_VARIANTS_SHEET), 0),
    (GENERAL_BOOK_SHEET, "Загальні - назви батальйону, посади, підрозділи",
     "Налаштовується один раз. " + _ROWS_HINT, sb.TAB_GENERAL, (GENERAL_SHEET, SUBDIVISIONS_SHEET), 0),
)
# [(аркуш, таблиця чи None), ...] - усі таблиці теперішнього вигляду книги.
BOOK_TABLES = (
    (NUMBERS_SHEET, None),
    *((sheet_name, table) for sheet_name, _title, _hint, _color, tables, _frozen in BOOK_SHEETS for table in tables),
)

INSTRUCTIONS_TITLE = "НАЛАШТУВАННЯ: БР і рапорти на додаткову винагороду"
INSTRUCTIONS_INTRO = (
    "Усе, що налаштовується під ваш батальйон, - у цій книзі (замість колишніх змінних constants.py і "
    "resources/data.json): кілька аркушів, на більшості - кілька таблиць зі смугою «■ назва». Назви таблиць і "
    "заголовки колонок не змінюйте; пояснення колонки - у примітці до її заголовка (червоний кутик). Помилку в книзі "
    "програма покаже при запуску: аркуш, таблицю, рядок і колонку."
)


def _at(table):
    """(аркуш, таблиця чи None) - куди веде посилання на таблицю з "Інструкції"."""
    sheet_name, section, _legacy = _where(table)
    return sheet_name, section


INSTRUCTIONS_TASKS = (
    ("Новий місяць: номери щоденних БР, ЗАВДАННЯ і БЗ (або програма зчитає їх з папки з документами сама)",
     (_at(NUMBERS_SHEET),)),
    ("Новий ЖБД чи БН - новий рядок (у попереднього впишіть «Кінець»)", (_at(LOG_WAR_SHEET), _at(BN_SHEET))),
    ("Новий документ бригади чи корпусу", (_at(HIGHER_REFERENCES_SHEET),)),
    ("Нова підстава для РТГр, ЗРАДН та інших окремих категорій", (_at(OTHER_REFERENCES_SHEET),)),
    ("Наказ, яким виплачено за минулий місяць", (_at(CHANGES_SHEET),)),
    (f"Нова категорія - рядок під своїм пунктом; новий пункт - рядок «{POINT_ROW_MARKER}» з категоріями",
     (_at(POINTS_SHEET),)),
    ("Те саме для рапорту командира/ТВО", (_at(COMMANDER_POINTS_SHEET),)),
    ("Коли статус невиплати не враховувати (напр. СЗЧ при оголошеному розшуку)", (_at(NOT_PAID_EXCEPTIONS_SHEET),)),
    ("Текст нового пункту рапорту чи зміна текстів рапорту", (_at(SECTION_TEXTS_SHEET), _at(REPORT_TEXTS_SHEET))),
    ("Тексти завдань підрозділів у БР, тексти й варіанти тижневого ЗАВДАННЯ",
     (_at(SUBDIVISION_TEXTS_SHEET), _at(TASK_TEXTS_SHEET), _at(TASK_VARIANTS_SHEET))),
    ("Назва батальйону, посади командира й начальника штабу, адресати БР", (_at(GENERAL_SHEET),)),
    ("Підрозділи батальйону (порядок у рядку «КОМАНДИРУ ...» БР)", (_at(SUBDIVISIONS_SHEET),)),
)
INSTRUCTIONS_TIPS = (
    "Дати - дд.мм.рррр або звичайна дата Excel. «Так»/«Ні» - зі списку в клітинці (зелене - Так, сіре - Ні). Кілька "
    "значень в одній клітинці - через кому; кілька рядків тексту - Alt+Enter.",
    "Новий запис - у порожній рядок під таблицею; бракує - вставте рядок усередині таблиці (правою кнопкою на номері "
    "рядка -> «Вставити»). Порожні рядки програма пропускає. Власні колонки праворуч таблиці - можна, програма їх не читає.",
)


# ---------------------------------------------------------------- читання


def _rows(workbook, table, columns, optional_sheet=False):
    """Рядки таблиці table (sb.read_rows), де номер рядка книги лежить у самому
    рядку під ключем "_row" (для повідомлень про помилку вже після читання)."""
    sheet_name, section, legacy = _where(table)
    rows = sb.read_rows(workbook, sheet_name, columns, optional_sheet=optional_sheet, section=section, legacy_sheet=legacy)
    return [{**row, "_row": row_number} for row_number, row in rows]


def _check_unique(rows, key, table, header):
    seen = {}
    for row in rows:
        value = row[key]
        if value in seen:
            raise sb.SettingsError(f"{sb.location(_label(table), row.get('_row'), header)}: {value!r} уже є в рядку {seen[value]}.")
        seen[value] = row.get("_row")


def _day_order(row):
    return (0, 0) if row["day"] == sb.PREVIOUS_MONTH_DAY else (1, row["day"])


def _number_sheet_name(workbook, name):
    """Аркуш з номерами словника name: спільний "Номери документів", а в книзі
    номерів етапу, записаній до його появи, - колишній окремий аркуш."""
    if sb.find_sheet(workbook, NUMBERS_SHEET) is None and sb.find_sheet(workbook, LEGACY_NUMBER_SHEETS[name]) is not None:
        return LEGACY_NUMBER_SHEETS[name]
    return NUMBERS_SHEET


def _read_number_table(workbook, name):
    sheet_name = _number_sheet_name(workbook, name)
    pairs = NUMBER_TABLES[name]
    rows = []
    for row in _rows(workbook, sheet_name, (DAY_COLUMN, *(column for _field, column in pairs))):
        values = {field: row[column.key] for field, column in pairs}
        # Рядок "попередній" (перехідний день) завжди є в щоденних БР, а в
        # ЗАВДАННЯ й БЗ - лише якщо там щось вписано.
        if row["day"] == sb.PREVIOUS_MONTH_DAY and name != DAILY and not any(values.values()):
            continue
        rows.append({"day": row["day"], **values, "_row": row["_row"]})
    _check_unique(rows, "day", sheet_name, DAY_COLUMN.header)
    return sorted(rows, key=_day_order)


def read_number_tables(workbook):
    """{ім'я словника NUMBER_OF_DOCUMENTS_BRS_*: [рядок, ...]} у порядку днів."""
    return {name: _read_number_table(workbook, name) for name in NUMBER_TABLES}


def read_references(workbook):
    """Рядки чотирьох таблиць підстав -> один список {group, start, end, lines,
    higher_hq, id}: у таблицях ЖБД і БН група - назва таблиці, у таблиці
    вищого штабу "Вищий штаб" = Так."""
    rows = []
    for sheet_name, group, higher_hq in REFERENCE_SHEETS:
        columns = _OWN_REFERENCE_COLUMNS if group else GROUP_REFERENCE_COLUMNS
        for row in _rows(workbook, sheet_name, columns):
            rows.append({
                "group": group or row["group"], "start": row["start"], "end": row["end"], "lines": row["lines"],
                "higher_hq": higher_hq, "id": row["id"], "_row": row["_row"],
            })
    return rows


def _is_point_row(row):
    words = re.sub(r"[\W_]+", " ", row["category"]).strip().casefold()
    return not words or words == POINT_ROW_MARKER


def read_point_blocks(workbook, sheet_name):
    """Таблиця пунктів sheet_name -> (рядки пунктів {point, general, column}, рядки
    категорій {point, category, grounds, exclude_general, прапорці, column}).
    Рядок "весь пункт" (чи без категорії) - сам пункт, решта - його категорії;
    порожній "Пункт" - пункт з рядка вище."""
    points, categories, current = [], [], ""
    for row in _rows(workbook, sheet_name, POINT_BLOCK_COLUMNS):
        point = row["point"] or current
        if not point:
            raise sb.SettingsError(f"{sb.location(_label(sheet_name), row['_row'], 'Пункт')}: не вказано пункт (а вище немає рядка пункту).")
        current = point
        if _is_point_row(row):
            used = next((column for column in POINT_BLOCK_COLUMNS
                         if column.key in _CATEGORY_ONLY_KEYS and row[column.key] not in (None, "", [])), None)
            if used is not None:
                raise sb.SettingsError(
                    f'{sb.location(_label(sheet_name), row["_row"], used.header)}: у рядку "{POINT_ROW_MARKER}" ця колонка не діє - '
                    "задайте її в рядках категорій пункту."
                )
            points.append({"point": point, "general": row["grounds"], "column": row["column"], "_row": row["_row"]})
        else:
            categories.append({
                "point": point, "category": row["category"], "grounds": row["grounds"],
                "exclude_general": row["exclude_general"],
                **{key: default if row[key] is None else row[key] for key, default in _CATEGORY_FLAGS.items()},
                "column": row["column"], "_row": row["_row"],
            })
    return points, categories


def _read_not_paid_exceptions(workbook):
    """Рядки таблиці "Винятки невиплати" або None, якщо її в книзі ще немає
    (книга, створена до її появи: діють початкові винятки, а load_settings
    переписує книгу в теперішній вигляд - уже з цією таблицею)."""
    if not sb.has_table(workbook, *_where(NOT_PAID_EXCEPTIONS_SHEET)):
        return None
    rows, seen = [], {}
    for row in _rows(workbook, NOT_PAID_EXCEPTIONS_SHEET, NOT_PAID_EXCEPTION_COLUMNS):
        key = row["status"].strip().upper()
        if key in seen:
            raise sb.SettingsError(
                f'{sb.location(_label(NOT_PAID_EXCEPTIONS_SHEET), row["_row"], "Статус невиплати")}: статус уже є в рядку {seen[key]}.'
            )
        seen[key] = row["_row"]
        rows.append(row)
    return rows


def not_paid_exceptions(settings):
    """NOT_PAID_IGNORED_WHEN_CONFIRMED_BY: {статус: {підтверджуючі статуси}}."""
    rows = settings.get("not_paid_exceptions")
    if rows is None:
        rows = DEFAULT_NOT_PAID_EXCEPTIONS
    return {row["status"]: set(row["confirmed_by"]) for row in rows}


def _texts(workbook, table):
    sheet_name, section, legacy = _where(table)
    return sb.read_texts(workbook, sheet_name, section=section, legacy_sheet=legacy)


def read_settings(path):
    """Книга -> словник налаштувань."""
    general_sheet, general_section, general_legacy = _where(GENERAL_SHEET)
    with sb.open_for_reading(path) as workbook:
        points, categories = read_point_blocks(workbook, POINTS_SHEET)
        commander_points, commander_categories = read_point_blocks(workbook, COMMANDER_POINTS_SHEET)
        settings = {
            "general": sb.read_settings_sheet(workbook, general_sheet, GENERAL_SETTINGS, section=general_section,
                                              legacy_sheet=general_legacy),
            "numbers": read_number_tables(workbook),
            "references": read_references(workbook),
            "changes_orders": _rows(workbook, CHANGES_SHEET, CHANGES_COLUMNS, optional_sheet=True),
            "points": points,
            "categories": categories,
            "commander_points": commander_points,
            "commander_categories": commander_categories,
            "section_texts": _rows(workbook, SECTION_TEXTS_SHEET, SECTION_TEXT_COLUMNS),
            "report_texts": _texts(workbook, REPORT_TEXTS_SHEET),
            "subdivisions": _rows(workbook, SUBDIVISIONS_SHEET, SUBDIVISION_COLUMNS),
            "subdivision_texts": _rows(workbook, SUBDIVISION_TEXTS_SHEET, SUBDIVISION_TEXT_COLUMNS),
            "task_texts": _texts(workbook, TASK_TEXTS_SHEET),
            "task_variants": _rows(workbook, TASK_VARIANTS_SHEET, TASK_VARIANT_COLUMNS),
            "not_paid_exceptions": _read_not_paid_exceptions(workbook),
        }
    return settings


# ---------------------------------------------------------------- побудова структур constants.py


def point_key(text):
    """Ключ пункту: "30" -> 30 (як число в колонці дати ОБЛІК.xlsx), "NOT_PAID" -> "NOT_PAID"."""
    text = str(text).strip()
    return int(text) if text.isdigit() else text


def section_key(name):
    """Ключ PARAGRAPH_MAP для підрозділу name - так само, як generate_content_br_general
    приводить колонку ПІДРОЗДІЛ (без пробілів, нижній регістр); "упр, штаб" -
    складена секція кількох підрозділів (ключ лишається з ", ")."""
    parts = [part.replace(" ", "").lower() for part in str(name).split(",")]
    return ", ".join(part for part in parts if part)


def _reference_entry(row, today):
    entry = {"start": row["start"], "end": row["end"] or today, "lines": list(row["lines"])}
    if row.get("id"):
        entry["id"] = row["id"]
    return entry


def build_reference_groups(reference_rows, today):
    """{назва групи (нормалізована): (назва як у книзі, [запис, ...])} - записи
    {"start", "end", "lines"[, "id"]} у порядку рядків. Порожній "Кінець" -
    today (діє досі), як колись datetime.now() у constants.py."""
    groups = {}
    for row in reference_rows:
        _name, entries = groups.setdefault(sb.normalize_name(row["group"]), (row["group"], []))
        entries.append(_reference_entry(row, today))
    return groups


def _higher_hq_groups(reference_rows, today):
    """BR_HIGHT_UNIT: {група: [записи таблиці "Підстави вищого штабу"]} - лише
    ці підстави звірка рапорту (checker_accounting/report_checker.py) не
    перевіряє з файлами."""
    result = {}
    for row in reference_rows:
        if row["higher_hq"]:
            result.setdefault(row["group"], []).append(_reference_entry(row, today))
    return result


def _resolve_groups(names, groups, where):
    entries = []
    for name in names:
        group = groups.get(sb.normalize_name(name))
        if group is None:
            known = ", ".join(display for display, _entries in groups.values()) or "жодної"
            raise sb.SettingsError(
                f'{where}: немає групи підстав "{name}" (є: {known}). Групи підстав - {_REFERENCE_SHEETS_TEXT}.'
            )
        entries.extend(group[1])
    return entries


def _column_value(names):
    """Колонка(и) підстав: [] -> None, одна -> рядок, кілька -> кортеж (як у constants.py)."""
    if not names:
        return None
    return names[0] if len(names) == 1 else tuple(names)


_EXCLUDE_ALL_WORDS = {"всі", "все", "усі", "all", "так", "т"}
_EXCLUDE_NONE_WORDS = {"", "ні", "н", "no", "-"}


def _exclude_value(text):
    word = text.strip().casefold()
    if word in _EXCLUDE_NONE_WORDS:
        return []
    if word in _EXCLUDE_ALL_WORDS:
        return "all"
    return sb.split_list(text)


_RESERVED_CATEGORY_NAMES = {"general", "column_basis_from_tracking"}


def build_categories(point_rows, category_rows, groups, sheet_name):
    """Рядки пунктів і категорій таблиці sheet_name -> словник формату
    MONEY_REPORT_CATEGORIES: {пункт: {"general": [...],
    ["column_basis_from_tracking": ...,] категорія: {"grounds", "use_brs", ...,
    "column_basis_from_tracking"}, ...}}. Порядок пунктів - як рядків "весь
    пункт" (пункти без такого рядка - за першою категорією), категорій - як
    рядків: перший збіг статусу перемагає (resolve_status_category)."""
    categories = {}
    for row in point_rows:
        point = point_key(row["point"])
        if point in categories:
            raise sb.SettingsError(
                f'{sb.location(_label(sheet_name), row.get("_row"), "Пункт")}: пункт {point!r} уже є вище '
                f'(рядок "{POINT_ROW_MARKER}" - один на пункт).'
            )
        config = {"general": _resolve_groups(row["general"], groups, sb.location(_label(sheet_name), row.get("_row"), "Підстави"))}
        column = _column_value(row["column"])
        if column is not None:
            config["column_basis_from_tracking"] = column
        categories[point] = config

    for row in category_rows:
        point = point_key(row["point"])
        config = categories.setdefault(point, {"general": []})
        name = row["category"]
        where = sb.location(_label(sheet_name), row.get("_row"), "Категорія")
        if name in _RESERVED_CATEGORY_NAMES:
            raise sb.SettingsError(f"{where}: {name!r} - службова назва, оберіть іншу.")
        if name in config:
            raise sb.SettingsError(f"{where}: категорія {name!r} уже є в пункті {point!r}.")
        config[name] = {
            "grounds": _resolve_groups(row["grounds"], groups, sb.location(_label(sheet_name), row.get("_row"), "Підстави")),
            "use_brs": row["use_brs"],
            "use_brs_from_selected_folder": row["use_brs_from_selected_folder"],
            "exclude_general": _exclude_value(row["exclude_general"]),
            "default": row["default"],
            "generate_br": row["generate_br"],
            "include_to_report": row["include_to_report"],
            "required_basis": row["required_basis"],
            "column_basis_from_tracking": _column_value(row["column"]),
        }
    return categories


def _section_texts(rows, key):
    return {str(point_key(row["point"])): row[key] for row in rows if row[key]}


def build_report_static(settings):
    """MONEY_REPORT_STATIC (колишній data.json["data_statik_report_for_get_additional_money"]):
    тексти таблиці "Загальні тексти рапорту" + SECTIONS/COMMANDER_SECTIONS з "Тексти пунктів"."""
    _check_unique(settings["section_texts"], "point", SECTION_TEXTS_SHEET, "Пункт")
    return {
        **settings["report_texts"],
        "SECTIONS": _section_texts(settings["section_texts"], "report"),
        "COMMANDER_SECTIONS": _section_texts(settings["section_texts"], "commander"),
    }


def build_task_order_variants(rows):
    """TASK_ORDER_VARIANTS: [[(підрозділ, завдання), ...], ...] у порядку номерів
    варіантів; рядок лише з номером - порожній варіант (тимчасовий фолбек на
    варіант 1, generate_documents_br_weekly_task.py)."""
    variants = {}
    for row in rows:
        pairs = variants.setdefault(row["variant"], [])
        if row["title"] or row["text"]:
            if not (row["title"] and row["text"]):
                header = "Завдання" if row["title"] else "Підрозділ"
                raise sb.SettingsError(f"{sb.location(_label(TASK_VARIANTS_SHEET), row.get('_row'), header)}: потрібні і підрозділ, і завдання.")
            pairs.append((row["title"], row["text"]))
    return [variants[number] for number in sorted(variants)]


def build_company_sections(subdivision_rows, text_rows):
    """(SEQUENCE, PARAGRAPH_MAP, PARAGRAPH_MAP_FOR_EXTRACT) з таблиць
    "Підрозділи" і "Тексти підрозділів" (кожна - у власному порядку рядків)."""
    sequence = [row["name"] for row in subdivision_rows]
    paragraph_map, extract_map, seen = {}, {}, {}
    for row in text_rows:
        key = section_key(row["name"])
        if key in seen:
            raise sb.SettingsError(
                f'{sb.location(_label(SUBDIVISION_TEXTS_SHEET), row.get("_row"), "Підрозділ")}: підрозділ {row["name"]!r} уже є в рядку {seen[key]}.'
            )
        seen[key] = row.get("_row")
        if row["br_text"]:
            paragraph_map[key] = row["br_text"]
        if row["extract_text"]:
            extract_map[key] = row["extract_text"]
    return sequence, paragraph_map, extract_map


def build_constants(settings, today):
    """Усе, що constants.py бере з книги, крім номерів документів (ті залежать
    від обраного при запуску місяця - build_number_dicts). today -
    "дд.мм.рррр" для підстав без кінця."""
    general = settings["general"]
    groups = build_reference_groups(settings["references"], today)
    log_war = groups.get(sb.normalize_name(LOG_WAR_GROUP), (LOG_WAR_GROUP, []))[1]
    bn = groups.get(sb.normalize_name(BN_GROUP), (BN_GROUP, []))[1]
    sequence, paragraph_map, extract_map = build_company_sections(settings["subdivisions"], settings["subdivision_texts"])
    task_texts = settings["task_texts"]
    return SimpleNamespace(
        **{key: general[key] for key in UNIT_KEYS},
        HIGHER_COMMANDER_TITLE=general["HIGHER_COMMANDER_TITLE"],
        COMMANDER_TITLE=general["COMMANDER_TITLE"],
        SEQUENCE_HIGHER_COMMANDER=list(general["SEQUENCE_HIGHER_COMMANDER"]),
        MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR=list(log_war),
        MONEY_REPORT_GENERAL_REFERENCES_BN=list(bn),
        BR_HIGHT_UNIT=_higher_hq_groups(settings["references"], today),
        MONEY_REPORT_CHANGES_ORDER_REFERENCES=[_reference_entry(row, today) for row in settings["changes_orders"]],
        MONEY_REPORT_CATEGORIES=build_categories(settings["points"], settings["categories"], groups, POINTS_SHEET),
        COMMANDER_MONEY_REPORT_CATEGORIES=build_categories(
            settings["commander_points"], settings["commander_categories"], groups, COMMANDER_POINTS_SHEET,
        ),
        MONEY_REPORT_STATIC=build_report_static(settings),
        BR_TASK_ORDER_TEXTS=dict(task_texts),
        TASK_ORDER_VARIANTS=build_task_order_variants(settings["task_variants"]),
        SEQUENCE=sequence,
        PARAGRAPH_MAP=paragraph_map,
        PARAGRAPH_MAP_FOR_EXTRACT=extract_map,
        NOT_PAID_IGNORED_WHEN_CONFIRMED_BY=not_paid_exceptions(settings),
    )


# ---------------------------------------------------------------- номери документів


def _month_length(month, year):
    return calendar.monthrange(year, month)[1]


def _previous_month(month, year):
    return (12, year - 1) if month == 1 else (month - 1, year)


def date_key_for_day(day, month, year):
    """Рядок аркуша номерів (день чи "попередній") -> ключ словника "дд.мм.рррр"
    обраного місяця, або None для дня, якого в цьому місяці немає."""
    if day == sb.PREVIOUS_MONTH_DAY:
        prev_month, prev_year = _previous_month(month, year)
        return f"{_month_length(prev_month, prev_year):02d}.{prev_month:02d}.{prev_year}"
    if day > _month_length(month, year):
        return None
    return f"{day:02d}.{month:02d}.{year}"


def number_fields(name):
    """Поля запису словника номерів name (напр. ["бат", "брг", "general_br_bat"])."""
    return [field for field, _column in NUMBER_TABLES[name]]


def number_dict_from_rows(rows, name, month, year):
    """Рядки таблиці номерів name -> словник {"дд.мм.рррр": {поле: номер}} обраного місяця."""
    fields = number_fields(name)
    result = {}
    for row in rows:
        key = date_key_for_day(row["day"], month, year)
        if key is not None:
            result[key] = {field: row.get(field, "") for field in fields}
    return result


def build_number_dicts(number_tables, month, year):
    """(EVERY_DAY, EVERY_WEEK, SAVE) - словники NUMBER_OF_DOCUMENTS_BRS_* для
    місяця month (число) року year."""
    return tuple(number_dict_from_rows(number_tables[name], name, month, year) for name in NUMBER_TABLES)


def day_labels(keys):
    """{ключ "дд.мм.рррр": рядок аркуша номерів (день чи "попередній")} для
    ключів словника номерів: дні основного (найчастішого серед keys) місяця -
    числом, останній день місяця перед ним - "попередній". Розбирається текст
    ключа, а не календарна дата, тож і "31.11.2026" (такого дня немає) лягає в
    рядок 31."""
    parsed = {}
    for key in keys:
        parts = str(key).split(".")
        if len(parts) == 3 and all(part.isdigit() for part in parts):
            parsed[key] = tuple(int(part) for part in parts)
    if not parsed:
        return {}
    # Основний - найчастіший місяць; за нічиєї (мало ключів, напр. лише
    # "31.10" і "01.11") - пізніший: єдиний "чужий" ключ - це останній день
    # попереднього місяця.
    counts = Counter((month, year) for _day, month, year in parsed.values())
    main_month, main_year = max(counts, key=lambda month_year: (counts[month_year], month_year[1], month_year[0]))
    previous = _previous_month(main_month, main_year)
    labels = {}
    for key, (day, month, year) in parsed.items():
        if (month, year) == (main_month, main_year) and 1 <= day <= 31:
            labels[key] = day
        elif (month, year) == previous and day == _month_length(month, year):
            labels[key] = sb.PREVIOUS_MONTH_DAY
    return labels


def write_number_fields(path, fields):
    """Записує номери в аркуш "Номери документів" книги path (решта книги не
    змінюється). fields - [(ім'я словника, поле, словник номерів, ключі), ...]:
    для кожного ключа з "ключі" пишеться словник[ключ][поле]. Повертає [(ім'я
    словника, ключ), ...] - ключі, для яких рядка дня в аркуші немає (не
    записані)."""
    columns = {(name, field): column for name, pairs in NUMBER_TABLES.items() for field, column in pairs}
    updates, written = {}, {}
    for name, field, number_dict, keys in fields:
        labels = day_labels(keys)
        for key in keys:
            label = labels.get(key)
            if label is None or key not in number_dict:
                continue
            updates.setdefault(label, {})[columns[(name, field)].key] = number_dict[key].get(field, "")
            written.setdefault(label, {})[(name, key)] = None
    if not updates:
        return []
    absent = sb.update_table_cells(path, NUMBERS_SHEET, DAY_COLUMN, NUMBER_COLUMNS[1:], updates)
    return [item for label, items in written.items() if label in absent for item in items]


def number_rows(number_dict, name):
    """Словник номерів name -> рядки таблиці номерів ({"day": день чи
    "попередній", поле: номер}); ключі не з основного місяця й не перехідний
    день відкидаються (day_labels)."""
    labels = day_labels(number_dict.keys())
    return [
        {"day": labels[key], **{field: "" if entry.get(field) is None else str(entry.get(field)) for field in number_fields(name)}}
        for key, entry in number_dict.items() if key in labels
    ]


def _number_sheet_rows(tables):
    """{ім'я словника: [рядок таблиці номерів]} -> рядки аркуша "Номери
    документів": один рядок на день, поля трьох таблиць - у своїх колонках."""
    by_day = {}
    for name, pairs in NUMBER_TABLES.items():
        for row in tables.get(name, []):
            combined = by_day.setdefault(row["day"], {"day": row["day"]})
            for field, column in pairs:
                combined[column.key] = row.get(field, "")
    return sorted(by_day.values(), key=_day_order)


def _write_numbers_sheet(workbook, tables):
    sb.write_table(workbook, NUMBERS_SHEET, NUMBER_COLUMNS, _number_sheet_rows(tables),
                   title=NUMBERS_TITLE, hint=NUMBERS_HINT, tab_color=sb.TAB_MONTHLY, frozen_columns=1)


def write_numbers_workbook(path, number_dicts):
    """Окрема книга лише з аркушем "Номери документів" (книга папки етапу).
    number_dicts - {ім'я словника: словник номерів}."""
    workbook = sb.new_workbook()
    _write_numbers_sheet(workbook, {name: number_rows(number_dicts.get(name, {}), name) for name in NUMBER_TABLES})
    sb.save_workbook(workbook, path)


def read_numbers_workbook(path, keys_by_name):
    """Номери з книги path для ключів keys_by_name ({ім'я словника: [ключ, ...]}):
    {ім'я словника: {ключ: запис}} - лише ключі, рядок яких у книзі є."""
    with sb.open_for_reading(path) as workbook:
        tables = read_number_tables(workbook)
    result = {}
    for name, keys in keys_by_name.items():
        fields = number_fields(name)
        rows_by_day = {row["day"]: row for row in tables[name]}
        labels = day_labels(keys)
        result[name] = {
            key: {field: rows_by_day[labels[key]].get(field, "") for field in fields}
            for key in keys if labels.get(key) in rows_by_day
        }
    return result


# ---------------------------------------------------------------- запис


def _without_row_numbers(rows):
    return [{key: value for key, value in row.items() if key != "_row"} for row in rows]


def reference_sheet_rows(references):
    """Рядки підстав -> {таблиця: [рядок, ...]}: ЖБД і БН - у власні таблиці,
    решта - у таблицю вищого штабу чи інших підстав, рядки однієї групи поруч
    (групи - у порядку першої появи, записи групи - у своєму порядку)."""
    by_sheet = {sheet_name: [] for sheet_name, _group, _higher_hq in REFERENCE_SHEETS}
    for row in references:
        group = sb.normalize_name(row["group"])
        if row["higher_hq"]:
            sheet_name = HIGHER_REFERENCES_SHEET
        elif group == sb.normalize_name(LOG_WAR_GROUP):
            sheet_name = LOG_WAR_SHEET
        elif group == sb.normalize_name(BN_GROUP):
            sheet_name = BN_SHEET
        else:
            sheet_name = OTHER_REFERENCES_SHEET
        by_sheet[sheet_name].append(row)
    for sheet_name in (HIGHER_REFERENCES_SHEET, OTHER_REFERENCES_SHEET):
        order = {}
        for row in by_sheet[sheet_name]:
            order.setdefault(sb.normalize_name(row["group"]), len(order))
        by_sheet[sheet_name].sort(key=lambda row: order[sb.normalize_name(row["group"])])
    return by_sheet


def point_block_rows(point_rows, category_rows):
    """Рядки пунктів і категорій -> рядки таблиці пунктів: кожен пункт - рядок
    "весь пункт" і під ним його категорії (у своєму порядку)."""
    point_by_key, categories_by_key = {}, {}
    for row in point_rows:
        point_by_key.setdefault(str(point_key(row["point"])), row)
    for row in category_rows:
        categories_by_key.setdefault(str(point_key(row["point"])), []).append(row)
    rows = []
    for key in [*point_by_key, *(key for key in categories_by_key if key not in point_by_key)]:
        point_row = point_by_key.get(key, {})
        point = point_row.get("point", key)
        rows.append({
            "point": point, "category": POINT_ROW_MARKER, "grounds": point_row.get("general", []),
            "column": point_row.get("column", []), "_style": "point",
        })
        rows.extend({**category, "_dim": ("point",)} for category in categories_by_key.get(key, []))
    return rows


def _sections(settings):
    """{таблиця: sb.Section} - усі таблиці книги (крім номерів документів) з
    даними settings."""
    references = reference_sheet_rows(settings["references"])
    group_names = tuple(dict.fromkeys(row["group"] for row in settings["references"]))
    # У колонці "Підстави" - випадний список груп підстав.
    point_columns = tuple(
        replace(column, choices=group_names) if column.key == "grounds" else column for column in POINT_BLOCK_COLUMNS
    )
    not_paid = settings.get("not_paid_exceptions")

    def table(name, columns, rows, **options):
        return sb.Section(name, TABLE_HINTS[name], columns, _without_row_numbers(rows), **options)

    def texts(name, settings_list, values, widths):
        return sb.Section(name, TABLE_HINTS[name], settings=settings_list, values=values, widths=widths, wrap_values=True)

    return {
        LOG_WAR_SHEET: table(LOG_WAR_SHEET, _OWN_REFERENCE_COLUMNS, references[LOG_WAR_SHEET]),
        BN_SHEET: table(BN_SHEET, _OWN_REFERENCE_COLUMNS, references[BN_SHEET]),
        HIGHER_REFERENCES_SHEET: table(HIGHER_REFERENCES_SHEET, GROUP_REFERENCE_COLUMNS, references[HIGHER_REFERENCES_SHEET],
                                       band_by="group"),
        OTHER_REFERENCES_SHEET: table(OTHER_REFERENCES_SHEET, GROUP_REFERENCE_COLUMNS, references[OTHER_REFERENCES_SHEET],
                                      band_by="group"),
        CHANGES_SHEET: table(CHANGES_SHEET, CHANGES_COLUMNS, settings["changes_orders"]),
        POINTS_SHEET: table(POINTS_SHEET, point_columns, point_block_rows(settings["points"], settings["categories"])),
        COMMANDER_POINTS_SHEET: table(COMMANDER_POINTS_SHEET, point_columns,
                                      point_block_rows(settings["commander_points"], settings["commander_categories"])),
        NOT_PAID_EXCEPTIONS_SHEET: table(NOT_PAID_EXCEPTIONS_SHEET, NOT_PAID_EXCEPTION_COLUMNS,
                                         list(DEFAULT_NOT_PAID_EXCEPTIONS if not_paid is None else not_paid)),
        # Колонки аркуша "Тексти рапорту" спільні: "Назва" - під "Пункт",
        # "Значення" й "Пояснення" - під текстами двох рапортів.
        SECTION_TEXTS_SHEET: table(SECTION_TEXTS_SHEET, SECTION_TEXT_COLUMNS, settings["section_texts"]),
        REPORT_TEXTS_SHEET: texts(REPORT_TEXTS_SHEET, REPORT_TEXT_SETTINGS, settings["report_texts"], (26, 80, 80)),
        SUBDIVISION_TEXTS_SHEET: table(SUBDIVISION_TEXTS_SHEET, SUBDIVISION_TEXT_COLUMNS, settings["subdivision_texts"]),
        TASK_TEXTS_SHEET: texts(TASK_TEXTS_SHEET, TASK_TEXT_SETTINGS, settings["task_texts"], (24, 75, 75)),
        TASK_VARIANTS_SHEET: table(TASK_VARIANTS_SHEET, TASK_VARIANT_COLUMNS, settings["task_variants"], band_by="variant"),
        GENERAL_SHEET: sb.Section(GENERAL_SHEET, TABLE_HINTS[GENERAL_SHEET], settings=GENERAL_SETTINGS,
                                  values=settings["general"]),
        SUBDIVISIONS_SHEET: table(SUBDIVISIONS_SHEET, SUBDIVISION_COLUMNS, settings["subdivisions"]),
    }


def _first_sentence(hint):
    """Опис таблиці на аркуші "Інструкція" - початок її підказки (до першої
    крапки чи двокрапки)."""
    return hint.split(". ")[0].split(": ")[0].rstrip(".") + "."


def _write_instructions(workbook):
    contents = [(NUMBERS_SHEET, sb.TAB_MONTHLY, [(None, _first_sentence(NUMBERS_HINT))])]
    contents.extend(
        (sheet_name, color, [(table, _first_sentence(TABLE_HINTS[table])) for table in tables])
        for sheet_name, _title, _hint, color, tables, _frozen in BOOK_SHEETS
    )
    sb.write_instructions(workbook, INSTRUCTIONS_TITLE, INSTRUCTIONS_INTRO, INSTRUCTIONS_TASKS, contents, INSTRUCTIONS_TIPS)


def write_settings(path, settings):
    """Словник налаштувань -> нова книга path (наявний файл замінюється)."""
    sections = _sections(settings)
    workbook = sb.new_workbook()
    _write_numbers_sheet(workbook, settings["numbers"])
    for sheet_name, title, hint, color, tables, frozen in BOOK_SHEETS:
        sb.write_sections(workbook, sheet_name, [sections[table] for table in tables], title=title, hint=hint,
                          tab_color=color, frozen_columns=frozen)
    _write_instructions(workbook)
    sb.save_workbook(workbook, path)


# ---------------------------------------------------------------- для constants.py


def _strip_row_numbers(value):
    if isinstance(value, dict):
        return {key: _strip_row_numbers(item) for key, item in value.items() if key != "_row"}
    if isinstance(value, (list, tuple)):
        return [_strip_row_numbers(item) for item in value]
    return value


def _layout_view(settings):
    """settings без номерів рядків і в тому порядку, у якому write_settings
    розкладає підстави й пункти по таблицях: так порівнюються дані книги до й
    після переписування (номери рядків і порядок груп підстав змінюються)."""
    view = _strip_row_numbers(settings)
    view["references"] = reference_sheet_rows(view["references"])
    for points, categories in (("points", "categories"), ("commander_points", "commander_categories")):
        view[points] = point_block_rows(view.pop(points), view.pop(categories))
    if view.get("not_paid_exceptions") is None:
        view["not_paid_exceptions"] = _strip_row_numbers(list(DEFAULT_NOT_PAID_EXCEPTIONS))
    return view


def _upgrade_layout(settings_path, settings):
    """Книга не в теперішньому вигляді (кожна таблиця - на власному аркуші,
    приховані аркуші, немає таблиці "Винятки невиплати") один раз
    переписується в теперішній: ті самі дані, відсутні винятки невиплати -
    початкові. Копія колишньої книги лишається поруч ("... (до
    оновлення).xlsx"); якщо переписана книга, прочитана назад, у чомусь не
    збіглася - колишня повертається на місце. Книга відкрита в Excel - спроба
    при наступному запуску. Повертає налаштування (після переписування - з
    номерами рядків нової книги)."""
    if not sb.missing_tables(settings_path, BOOK_TABLES):
        return settings
    # Копія, що лишилась від попереднього оновлення, не перезаписується.
    backup = sb.free_backup_path(settings_path)
    try:
        shutil.copy2(settings_path, backup)
        # Запис атомарний (sb.save_workbook): не вдався - книга та сама.
        write_settings(settings_path, settings)
    except (sb.SettingsError, OSError) as error:
        print_red(f"Вигляд {settings_path} не оновлено ({error}) - спробую при наступному запуску.")
        return settings
    try:
        upgraded = read_settings(settings_path)
    except sb.SettingsError:
        upgraded = None
    if upgraded is None or _layout_view(upgraded) != _layout_view(settings):
        os.replace(backup, settings_path)
        print_red(f"{settings_path}: оновлена книга не збіглася з колишньою - колишню повернуто, вигляд не змінено.")
        return settings
    print_green(
        f"{settings_path}: книгу переведено на новий вигляд (кілька таблиць на аркуші, без прихованих аркушів; дані ті "
        f"самі). Копія колишньої - {backup}."
    )
    return upgraded


def load_settings(settings_path, legacy_json_path, project_dir):
    """Словник налаштувань для constants.py. Книги ще немає, а є колишній
    resources/data.json - спершу переносить його разом зі старими значеннями
    constants.py у книгу (utils/settings_migration.py); якщо книга збіглася з
    колишніми значеннями, data.json видаляється. Немає ні того, ні іншого, чи
    книга з помилкою - програма зупиняється з червоним поясненням, що зробити
    (sb.stop)."""
    if not os.path.isfile(settings_path):
        if legacy_json_path and os.path.isfile(legacy_json_path):
            from utils.settings_migration import migrate_legacy_settings

            if migrate_legacy_settings(settings_path, legacy_json_path, project_dir):
                print_green(
                    f"Налаштування перенесено з {legacy_json_path} і колишнього constants.py у {settings_path} без втрат, "
                    "data.json видалено. Далі редагуйте книгу Excel."
                )
        else:
            raise sb.stop(
                f"Немає файлу налаштувань {settings_path}.",
                ("Створіть зразок: python resources_example/generate_sample.py (з теки generator_br_and_report_for_money).",
                 "Впишіть у нього дані свого підрозділу.", "Запустіть програму ще раз."),
            )
    try:
        settings = read_settings(settings_path)
    except sb.SettingsError as error:
        raise sb.stop(f"У книзі налаштувань помилка:\n{error}", sb.fix_in_book_steps(settings_path)) from error
    settings = _upgrade_layout(settings_path, settings)
    if settings["not_paid_exceptions"] is None:
        # Книга, створена до появи цієї таблиці (колись - словник у
        # constants.py), яку не вдалося переписати: цього разу - початкові.
        settings["not_paid_exceptions"] = [dict(row) for row in DEFAULT_NOT_PAID_EXCEPTIONS]
    return settings


def build_constants_or_exit(settings, today, settings_path=sb.SETTINGS_FILE_NAME):
    """build_constants; помилка значення книги - червоне пояснення з кроками
    (sb.stop), а не трасування Python."""
    try:
        return build_constants(settings, today)
    except sb.SettingsError as error:
        raise sb.stop(f"У книзі налаштувань помилка:\n{error}", sb.fix_in_book_steps(settings_path)) from error
