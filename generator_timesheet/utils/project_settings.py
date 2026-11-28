"""Книга налаштувань табеля resources/НАЛАШТУВАННЯ.xlsx.

read_settings повертає словник ТІЄЇ САМОЇ форми, що й колишній
resources/data.json ("unit", "groups", "EXCLUDE_NAMES", "IGNORE_LISTS" ...),
тож constants.py перевіряє й розбирає його так само, як і раніше, а
write_settings записує такий словник у книгу (зразок, перенесення data.json).

Колишній data.json при першому запуску переноситься в книгу сам
(load_settings) і видаляється, якщо книга збігається з ним."""
import json
import os
import re
import shutil
from contextlib import contextmanager

from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from utils import default_statuses
from utils import settings_book as sb
from utils import status_rules as sr
from utils.logging_utils import print_green, print_red

GENERAL_SHEET = "Загальні"
SUBDIVISIONS_SHEET = "Підрозділи"
CATEGORY_100_SHEET = "Ліміт категорії 100"
REVIEW_SHEET = "Без ручної перевірки"
ABSENT_SHEET = "Не подаються в БЧС"
IGNORE_SHEET = "Ігнорування addons"

UNIT_KEYS = ("SHORT_UNIT_BATTALION", "SHORT_UNIT_BRIGADE", "FULL_UNIT_BUT", "FULL_MILITARY_UNIT")

GENERAL_SETTINGS = (
    sb.Setting("SHORT_UNIT_BATTALION", label="Батальйон (скорочено)",
               note='Напр. "1 бмп". Чотири назви підрозділу разом зі словом ролі ("командир", "начальник штабу") дають змогу впізнати рядок підпису наприкінці рапорту.'),
    sb.Setting("SHORT_UNIT_BRIGADE", label="Бригада (скорочено)", note='Напр. "40 обрмп".'),
    sb.Setting("FULL_UNIT_BUT", label="Батальйон (повна назва, родовий відмінок)",
               note='Напр. "1 батальйону морської піхоти" - заголовок Відомості РОП/ВОП.'),
    sb.Setting("FULL_MILITARY_UNIT", label="Військова частина",
               note='"військової частини А0000" - заголовок Відомості РОП/ВОП; цифри - номер власної в/ч (книга ЦВС, переведення).'),
    sb.Setting("SPECIAL_TASK_GROUP", label="Спеціальна група (як у рапортах)",
               note='Напр. "РТГР «Назва»" ("з ..." - повернення, "у/до ..." - ВД).'),
    sb.Setting("BCHS_SIGNAL_GROUP_NAME", label="Група Signal з файлами БЧС",
               note="Звідси завантажуються файли БЧС у resources/information_unit."),
    sb.Setting("REPORTS_SIGNAL_GROUP_NAME", required=False, label="Група Signal з рапортами",
               note="Щоденні рапорти й відомості ВОП-РОП. Порожньо - звіряння з Signal вимкнене."),
    sb.Setting("INFORMATION_UNIT_KEEP_FILE_NAME", label="Файл БЧС, який не видаляти",
               note="Файл resources/information_unit, який лишається перед завантаженням нових файлів БЧС (управління і штаб)."),
    sb.Setting("CONST_PAYMENT_CATEGORY_100_MAX_DAYS", sb.INT, label="Ліміт днів категорії 100",
               note=f'Скільки днів поспіль у місяці можна мати категорію виплати 100 (ціле число 0-31). Власний ліміт людини - аркуш "{CATEGORY_100_SHEET}".'),
    sb.Setting("BRIGADE_MILITARY_UNITS", sb.LIST, required=False, label="Військові частини бригади",
               note="Номери через кому, напр. А0000, А0001. Переведення до такої частини (чи до власної) - ПРВДБ, до будь-якої іншої - ПРВДВЧ."),
)

# Аркуш "Підрозділи" книги generator_br_and_report_for_money має той самий
# вигляд (колонка "Підрозділ"), тож табель читає перелік і з тієї книги
# (constants.RECOGNIZED_SUBDIVISIONS) тією самою функцією - як файл, без
# імпорту коду того проєкту.
SUBDIVISION_COLUMNS = (
    sb.Column("name", "Підрозділ", nonempty=True, width=28,
              note="Підрозділ так, як його пишуть у колонці ПІДРОЗДІЛ (напр. \"1 рмп\"). Від кожного щодня очікується файл БЧС, у цьому порядку."),
)

CATEGORY_100_COLUMNS = (
    sb.Column("name", "ПІБ", nonempty=True, width=40),
    sb.Column("days", "Днів", sb.INT, width=10,
              note=f'Скільки днів поспіль у місяці людині можна мати категорію 100 (0-31, 31 - увесь місяць). Порожньо - спільний ліміт з аркуша "{GENERAL_SHEET}".'),
    sb.Column("color", "Колір", width=14, choices=(sb.YES, sb.NO), center=True,
              note="Заливка клітинки з ПІБ в ОБЛІК.xlsx і Перевірка_категорії_100.xlsx: Ні чи порожньо - без кольору; Так - жовтий; або колір числами, напр. FFFF00. Людина з кольором є в Перевірка_категорії_100.xlsx навіть без порушення."),
)

_REVIEW_GROUP = "Не писати у файл ручної перевірки"
# Таблиці аркуша "Люди" мають спільні колонки аркуша, тож скрізь "ПІБ" -
# першою колонкою (однакова ширина), а вужчі - праворуч.
REVIEW_COLUMNS = (
    sb.Column("name", "ПІБ", nonempty=True, width=40),
    sb.Column("date", "Дата", sb.DATE, nonempty=True, width=13,
              note="Дата з квадратних дужок [дд.мм.рррр] у рядку файлу ручної перевірки."),
    sb.Column("review", "ОБЛІК (стройовий)", sb.BOOL, default=True, width=19, group=_REVIEW_GROUP,
              note="Так (чи порожньо) - не писати цей рядок у Потребує_ручної_перевірки.txt."),
    sb.Column("payments", "ОБЛІК для виплат", sb.BOOL, default=True, width=19, group=_REVIEW_GROUP,
              note="Так (чи порожньо) - не писати цей рядок у Потребує_ручної_перевірки_виплати.txt."),
)

ABSENT_COLUMNS = (
    sb.Column("name", "ПІБ", nonempty=True, width=40,
              note="Людина, якої підрозділи можуть не подавати у файлах БЧС (напр. довготривале лікування): для неї немає запису \"Немає в жодному файлі information_unit\"."),
)

# Ключ IGNORE_LISTS (constants.ADDON_IGNORE_LIST_KEYS) -> назва книги в аркуші
# (та сама, що й у звіті mis_match.xlsx - constants.ADDON_BOOK_KINDS).
BOOKS = (
    ("BOOK_61", "book61"),
    ("CVS", "ЦВС"),
    ("COMBAT_TRAINING_JOURNAL", "журнал обліку БП"),
    ("TREATMENT_JOURNAL", "журнал лікування"),
)

IGNORE_COLUMNS = (
    sb.Column("name", "ПІБ", nonempty=True, width=40),
    sb.Column("book", "Книга", nonempty=True, width=22, choices=tuple(label for _key, label in BOOKS),
              note="Книга resources/addons (впізнається за вмістом, назва файлу не важить): " + ", ".join(label for _key, label in BOOKS) + "."),
    sb.Column("start", "Початок", sb.DATE, nonempty=True, width=13),
    sb.Column("end", "Кінець", sb.DATE, width=13, note="Порожньо - лише день Початку. Обидві дати включно."),
)

# Статуси, їхні кольори й виплати (колишні STATUS_MEANINGS, STATUS_COLORS,
# PAYMENT_STATUS_* у constants.py) - три аркуші; стандартні значення -
# utils/default_statuses.py.
STATUSES_SHEET = "Статуси"
OVERRIDES_SHEET = "Заміни для виплат"
PRIORITY_SHEET = "Пріоритет статусів"
# Ключ словника налаштувань -> аркуш статусів (і правил статусів -
# utils/status_rules.py); значення None - аркуша в книзі немає.
STATUS_SHEETS = {
    "STATUSES": STATUSES_SHEET,
    "PAYMENT_STATUS_OVERRIDES": OVERRIDES_SHEET,
    "PAYMENT_STATUS_PRIORITY": PRIORITY_SHEET,
    **sr.RULE_SHEETS,
}
STATUS_KEYS = tuple(STATUS_SHEETS)

_PAYMENTS_GROUP = "ОБЛІК для виплат"
STATUS_COLUMNS = (
    sb.Column("status", "Статус", sb.NUMBER_TEXT, nonempty=True, width=18, group="ОБЛІК.xlsx",
              note="Назва статусу так, як її пише ОБЛІК.xlsx (напр. ВП, ШП), чи категорія виплати (10, 100, 100_БЗ). "
                   "Щоб перейменувати статус, змініть назву тут - більше ніде: решта аркушів, ОБЛІК.xlsx і проєкт БР "
                   "підхоплять нову назву самі (розуміють і нову назву, і код). Категорії-числа не перейменовуються."),
    sb.Column("meaning", "Що означає", width=62, wrap=True, group="ОБЛІК.xlsx", note="Лише довідка - програма це не читає."),
    sb.Column("color", "Колір", sb.COLOR, width=10, required=False, group="ОБЛІК.xlsx",
              note="Залийте клітинку потрібним кольором (Головна -> Колір заливки) - так буде залито клітинку з цим "
                   "статусом в ОБЛІК.xlsx. Без заливки - клітинка лишається білою."),
    sb.Column("payment", "Записати завжди", sb.NUMBER_TEXT, width=16, group=_PAYMENTS_GROUP,
              note="Що записується в ОБЛІК для виплат замість цього статусу, хоч би що подав підрозділ: категорія "
                   "числом (10, 30, 100 ...) чи текстом (10_МЕД, 100_БПШП). Порожньо - статус лишається (або "
                   f'замінюється за аркушем "{OVERRIDES_SHEET}").'),
    sb.Column("fallback", "Якщо підрозділ не уточнив", sb.NUMBER_TEXT, width=17, group=_PAYMENTS_GROUP,
              note=f'Що записується, якщо аркуш "{OVERRIDES_SHEET}" цей статус не замінив (напр. НОВ -> 10). '
                   "Порожньо - статус лишається."),
    sb.Column("code", "Код", sb.NUMBER_TEXT, required=False, width=14,
              note="Незмінна технічна назва статусу - за нею програма впізнає його у своїх правилах (розпізнавання "
                   "рапортів, звірка з книгами addons і файлами БЧС). Не змінюйте; для нового статусу - порожньо."),
)

OVERRIDE_COLUMNS = (
    sb.Column("status", "Статус в ОБЛІК", sb.NUMBER_TEXT, nonempty=True, width=18,
              note="Статус клітинки ОБЛІК для виплат. РВЗ - ще не визначено (так само - порожня клітинка)."),
    sb.Column("unit_status", "Подає підрозділ", sb.NUMBER_TEXT, nonempty=True, width=18,
              note="Що на цю дату одностайно подають файли БЧС підрозділу (resources/information_unit)."),
    sb.Column("value", "Записати", sb.NUMBER_TEXT, nonempty=True, width=18,
              note="Що записати в клітинку ОБЛІК для виплат: категорію (10, 100, 10_МЕД ...) чи статус."),
)

PRIORITY_COLUMNS = (
    sb.Column("first", "Статус", sb.NUMBER_TEXT, nonempty=True, width=18),
    sb.Column("second", "Інший статус", sb.NUMBER_TEXT, nonempty=True, width=18),
    sb.Column("winner", "Перемагає", sb.NUMBER_TEXT, nonempty=True, width=18, note="Один з двох статусів рядка."),
)

# Де в книзі живе кожен колишній ключ data.json - для повідомлень про помилку
# (constants.py перевіряє значення вже після читання книги).
# Аркуші книги: на кожному - кілька таблиць зі смугою "■ назва"
# (settings_book.write_sections, locate_table). Назви таблиць - колишні назви
# аркушів (*_SHEET вище), тож книга, де кожна таблиця мала власний аркуш,
# читається так само.
PEOPLE_BOOK_SHEET = "Люди"
STATUSES_BOOK_SHEET = sr.STATUSES_BOOK_SHEET
PAYMENTS_BOOK_SHEET = "Виплати"
GENERAL_BOOK_SHEET = "Загальні"

# Таблиця -> аркуш, на якому вона лежить.
PLACE = {
    CATEGORY_100_SHEET: PEOPLE_BOOK_SHEET,
    ABSENT_SHEET: PEOPLE_BOOK_SHEET,
    IGNORE_SHEET: PEOPLE_BOOK_SHEET,
    REVIEW_SHEET: PEOPLE_BOOK_SHEET,
    STATUSES_SHEET: STATUSES_BOOK_SHEET,
    OVERRIDES_SHEET: PAYMENTS_BOOK_SHEET,
    PRIORITY_SHEET: PAYMENTS_BOOK_SHEET,
    GENERAL_SHEET: GENERAL_BOOK_SHEET,
    SUBDIVISIONS_SHEET: GENERAL_BOOK_SHEET,
    **sr.PLACE,
}


def _label(table):
    """'Люди", таблиця "Ліміт категорії 100' - для sb.location."""
    return sb.table_label(PLACE[table], table)


SOURCES = {
    "CONST_PAYMENT_CATEGORY_100_MAX_DAYS": f'аркуш "{_label(GENERAL_SHEET)}"',
    "EXCLUDE_NAMES": f'аркуш "{_label(CATEGORY_100_SHEET)}"',
    "EXCLUDE_FROM_REVIEW": f'аркуш "{_label(REVIEW_SHEET)}", колонка "{REVIEW_COLUMNS[2].header}"',
    "EXCLUDE_FROM_PAYMENTS_REVIEW": f'аркуш "{_label(REVIEW_SHEET)}", колонка "{REVIEW_COLUMNS[3].header}"',
    "EXCLUDE_ABSENT_FROM_INFORMATION_UNIT": f'аркуш "{_label(ABSENT_SHEET)}"',
    "IGNORE_LISTS": f'аркуш "{_label(IGNORE_SHEET)}"',
}


def source_of(key):
    """'НАЛАШТУВАННЯ.xlsx, аркуш "..." (KEY)' - місце значення key для повідомлень."""
    place = SOURCES.get(key)
    return f"{sb.SETTINGS_FILE_NAME}, {place} ({key})" if place else f"{sb.SETTINGS_FILE_NAME} ({key})"


# Підказка під смугою кожної таблиці.
TABLE_HINTS = {
    REVIEW_SHEET: "Уже перевірені пари «дата + ПІБ»: їх не треба писати у файли output/Потребує_ручної_перевірки*.txt. "
                  "«Так» у колонці - не писати в цей файл (порожньо - теж «Так»). Нова пара - новий рядок.",
    IGNORE_SHEET: "Для кожної книги resources/addons - люди й дні, яких не показувати в output/mis_match.xlsx. "
                  "«Кінець» порожньо - лише день «Початку»; обидві дати включно.",
    CATEGORY_100_SHEET: f"«Днів» - власний ліміт днів категорії 100 поспіль (порожньо - спільний з таблиці «{GENERAL_SHEET}»). "
                        "«Колір» - заливка клітинки ПІБ в ОБЛІК.xlsx: Так - жовтий, Ні - без кольору, або код кольору "
                        "(напр. FFC000).",
    ABSENT_SHEET: "Люди, яких підрозділи можуть не подавати у файлах БЧС (напр. довготривале лікування): для них немає "
                  "запису «Немає в жодному файлі information_unit».",
    STATUSES_SHEET: "«Колір» - сама заливка клітинки: оберіть її в Excel (Головна -> Колір заливки), і тим самим кольором "
                    "буде залито клітинку з цим статусом в ОБЛІК.xlsx. Праворуч - що пише ОБЛІК для виплат замість "
                    "статусу. Перейменувати статус - лише тут: змініть «Статус», і нову назву самі підхоплять решта "
                    "таблиць, ОБЛІК.xlsx і проєкт БР (сірий «Код» не чіпайте - за ним статус знають правила). Новий "
                    "статус - новий рядок з порожнім «Код».",
    OVERRIDES_SHEET: "Якщо в ОБЛІК для виплат у клітинці «Статус в ОБЛІК», а файли БЧС підрозділу на цю дату одностайно "
                     "подають «Подає підрозділ», у клітинку пишеться «Записати». Рядки РВЗ (ще не визначено) - головні: "
                     "так з файлів підрозділу береться категорія виплати." + sr.CODES_HINT,
    PRIORITY_SHEET: "Коли на одну дату для людини трапились обидва статуси рядка (різні файли підрозділу чи ОБЛІК і "
                    "файл підрозділу), правильним вважається «Перемагає»." + sr.CODES_HINT,
    GENERAL_SHEET: "Редагуйте жовті клітинки «Значення». Колонку «Параметр» не змінюйте: за нею програма знаходить рядок.",
    SUBDIVISIONS_SHEET: "Підрозділи, від яких щодня очікується файл БЧС: як у колонці ПІДРОЗДІЛ (напр. «1 рмп»), у цьому "
                        "порядку. Новий підрозділ - новий рядок.",
    **{name: hint for name, (_title, hint) in sr.SHEET_TITLES.items()},
}

_ROWS_HINT = (
    "Новий запис - у порожній рядок під таблицею; бракує рядків - вставте рядок усередині таблиці (правою кнопкою на "
    "номері рядка -> «Вставити»). Назви таблиць у смугах «■» і заголовки колонок не змінюйте."
)
# (аркуш, назва над таблицями, підказка, колір ярлика, таблиці згори донизу).
BOOK_SHEETS = (
    (PEOPLE_BOOK_SHEET, "Люди - винятки для окремих людей",
     "Чотири списки людей, для яких програма робить виняток. " + _ROWS_HINT, sb.TAB_MONTHLY,
     (REVIEW_SHEET, IGNORE_SHEET, CATEGORY_100_SHEET, ABSENT_SHEET)),
    (STATUSES_BOOK_SHEET, "Статуси - назви, кольори, виплати й особливі статуси",
     "Перша таблиця - усі статуси (перейменовують статус лише тут), друга - які статуси програма вважає статусом за "
     "замовчуванням, ВЛК, відпусткою тощо. " + _ROWS_HINT, sb.TAB_REPORT,
     (STATUSES_SHEET, sr.GROUPS_SHEET)),
    (PAYMENTS_BOOK_SHEET, "Виплати - що писати в ОБЛІК для виплат",
     "Заміни статусу, коли файли БЧС підрозділу його уточнюють, і пріоритет двох статусів на одну дату. " + _ROWS_HINT,
     sb.TAB_REPORT, (OVERRIDES_SHEET, PRIORITY_SHEET)),
    (sr.REPORTS_BOOK_SHEET, "Рапорти - як програма розпізнає щоденні рапорти",
     "Заголовки таблиць і речення рапорту, з яких програма бере статус, і що робити, коли рапорти на одну дату дали два "
     "різні статуси. Змінюйте обережно. " + _ROWS_HINT, sb.TAB_TEXTS,
     (sr.HEADINGS_SHEET, sr.SENTENCES_SHEET, sr.REPORT_CONFLICTS_SHEET)),
    (sr.ADDONS_BOOK_SHEET, "Звірка з addons - статуси Табеля й записи книг",
     "Як Табель звіряється з книгами resources/addons (book61, ЦВС, журнали). Змінюйте обережно. " + _ROWS_HINT,
     sb.TAB_TEXTS, (sr.ADDONS_SHEET, sr.EQUIVALENT_SHEET, sr.TRAINING_SHEET)),
    (GENERAL_BOOK_SHEET, "Загальні - назви підрозділу, групи Signal, підрозділи",
     "Налаштовується один раз. " + _ROWS_HINT, sb.TAB_GENERAL, (GENERAL_SHEET, SUBDIVISIONS_SHEET)),
)
# [(аркуш, таблиця), ...] - усі таблиці теперішнього вигляду книги.
BOOK_TABLES = tuple((sheet_name, table) for sheet_name, _title, _hint, _color, tables in BOOK_SHEETS for table in tables)

INSTRUCTIONS_TITLE = "НАЛАШТУВАННЯ табеля"
INSTRUCTIONS_INTRO = (
    "Усе, що налаштовується під ваш підрозділ, - у цій книзі: кілька аркушів, на кожному - кілька таблиць зі смугою "
    "«■ назва». Назви таблиць і заголовки колонок не змінюйте; пояснення колонки - у примітці до її заголовка (червоний "
    "кутик). Помилку в книзі програма покаже при запуску: аркуш, таблицю, рядок і колонку."
)


def _at(table):
    return PLACE[table], table


INSTRUCTIONS_TASKS = (
    ("Людину вже перевірено вручну - не писати її в Потребує_ручної_перевірки*.txt", (_at(REVIEW_SHEET),)),
    ("Не показувати людину в mis_match.xlsx за певні дні (книги resources/addons)", (_at(IGNORE_SHEET),)),
    ("Окремий ліміт днів категорії 100 для людини чи колір її ПІБ", (_at(CATEGORY_100_SHEET),)),
    ("Людину підрозділ може не подавати у файлах БЧС", (_at(ABSENT_SHEET),)),
    ("Перейменувати статус, змінити його колір чи виплату, додати новий статус", (_at(STATUSES_SHEET),)),
    ("Що програма вважає відпусткою, статусом за замовчуванням, переведенням тощо", (_at(sr.GROUPS_SHEET),)),
    ("Що писати в ОБЛІК для виплат, коли підрозділ уточнює статус", (_at(OVERRIDES_SHEET), _at(PRIORITY_SHEET))),
    ("Програма не розпізнає новий заголовок чи речення рапорту", (_at(sr.HEADINGS_SHEET), _at(sr.SENTENCES_SHEET))),
    ("Правила звірки статусів з книгами addons", (_at(sr.ADDONS_SHEET), _at(sr.EQUIVALENT_SHEET), _at(sr.TRAINING_SHEET))),
    ("Назва підрозділу, групи Signal, спільний ліміт категорії 100", (_at(GENERAL_SHEET),)),
    ("Новий підрозділ, від якого щодня має бути файл БЧС", (_at(SUBDIVISIONS_SHEET),)),
)
INSTRUCTIONS_TIPS = (
    "Дати - дд.мм.рррр або звичайна дата Excel. «Так»/«Ні» - зі списку в клітинці (зелене - Так, сіре - Ні). Кілька "
    "значень в одній клітинці - через кому.",
    "Новий запис - у порожній рядок під таблицею; бракує - вставте рядок усередині таблиці (правою кнопкою на номері "
    "рядка -> «Вставити»). Порожні рядки програма пропускає. Власні колонки праворуч таблиці - можна, програма їх не читає.",
)


def _text_rows(rows, key):
    return [row[key] for _row_number, row in rows]


def _color_value(text):
    """Клітинка "Колір" -> значення "color" колишнього data.json."""
    if text.casefold() in ("", "ні", "н", "false", "0", "-"):
        return False
    if text.casefold() in ("так", "т", "true", "1", "+"):
        return True
    return text


def _book_key(label, row_number):
    wanted = sb.normalize_name(label)
    for key, book_label in BOOKS:
        if wanted in (sb.normalize_name(key), sb.normalize_name(book_label)):
            return key
    labels = ", ".join(book_label for _key, book_label in BOOKS)
    raise sb.SettingsError(f"{sb.location(_label(IGNORE_SHEET), row_number, 'Книга')}: невідома книга {label!r} - можна: {labels}.")


def read_subdivisions(workbook):
    """Підрозділи таблиці "Підрозділи" у порядку рядків - на будь-якому аркуші
    книги (так само табель читає й таблицю "Підрозділи" книги
    generator_br_and_report_for_money, не знаючи, на якому вона аркуші)."""
    rows = sb.read_rows(workbook, None, SUBDIVISION_COLUMNS, section=SUBDIVISIONS_SHEET)
    return [row["name"] for _row_number, row in rows]


def read_subdivisions_file(path):
    """read_subdivisions для книги за шляхом (книга сусіднього проєкту)."""
    with sb.open_for_reading(path) as workbook:
        return read_subdivisions(workbook)


status_value = sr.status_value
# Статус для порівняння (той самий, що й у content/oblik_timesheet.py: без
# урахування регістру й зайвих пробілів).
_status_norm = sr.status_norm


def _check_unique_statuses(rows, keys, sheet_name, header):
    seen = {}
    for row_number, row in rows:
        key = tuple(_status_norm(status_value(row[name])) for name in keys)
        if key in seen:
            raise sb.SettingsError(f"{sb.location(sheet_name, row_number, header)}: такий рядок уже є в рядку {seen[key]}.")
        seen[key] = row_number


def _check_status_names(rows):
    """Назва статусу (колонка "Статус") не може збігатися з кодом іншого
    статусу (тоді незрозуміло, про який мова), а категорія-число не
    перейменовується (так її пише ОБЛІК для виплат і читає
    generator_br_and_report_for_money)."""
    codes = {_status_norm(row["code"] or row["status"]): row_number for row_number, row in rows}
    for row_number, row in rows:
        code = row["code"] or row["status"]
        if _status_norm(row["status"]) == _status_norm(code):
            continue  # не перейменовано (повтори ловить _check_unique_statuses)
        if str(code).isdigit():
            raise sb.SettingsError(
                f'{sb.location(_label(STATUSES_SHEET), row_number, "Статус")}: категорію виплати {code} не перейменовують - '
                f'поверніть "{code}".'
            )
        other = codes.get(_status_norm(row["status"]))
        if other is not None and other != row_number:
            raise sb.SettingsError(
                f'{sb.location(_label(STATUSES_SHEET), row_number, "Статус")}: назва "{row["status"]}" - це код статусу з рядка '
                f"{other}; оберіть іншу назву."
            )


def read_statuses(workbook):
    """Три аркуші статусів -> {"STATUSES": [...], "PAYMENT_STATUS_OVERRIDES":
    [...], "PAYMENT_STATUS_PRIORITY": [...]}; аркуша, якого в книзі немає, -
    None (книга, створена до появи цих аркушів). Статус без коду (новий рядок
    користувача) - код дорівнює назві."""
    result = dict.fromkeys(STATUS_KEYS)
    if sb.has_table(workbook, *_at(STATUSES_SHEET)):
        rows = sb.read_rows(workbook, PLACE[STATUSES_SHEET], STATUS_COLUMNS, section=STATUSES_SHEET)
        _check_status_names(rows)
        _check_unique_statuses(rows, ("status",), _label(STATUSES_SHEET), "Статус")
        coded = [(row_number, {**row, "code": row["code"] or row["status"]}) for row_number, row in rows]
        _check_unique_statuses(coded, ("code",), _label(STATUSES_SHEET), "Код")
        result["STATUSES"] = [
            {"status": row["status"], "code": row["code"], "meaning": row["meaning"], "color": row["color"],
             "payment": status_value(row["payment"]), "fallback": status_value(row["fallback"])}
            for _row_number, row in coded
        ]
    if sb.has_table(workbook, *_at(OVERRIDES_SHEET)):
        rows = sb.read_rows(workbook, PLACE[OVERRIDES_SHEET], OVERRIDE_COLUMNS, section=OVERRIDES_SHEET)
        _check_unique_statuses(rows, ("status", "unit_status"), _label(OVERRIDES_SHEET), "Подає підрозділ")
        result["PAYMENT_STATUS_OVERRIDES"] = [
            {"status": row["status"], "unit_status": status_value(row["unit_status"]), "value": status_value(row["value"])}
            for _row_number, row in rows
        ]
    if sb.has_table(workbook, *_at(PRIORITY_SHEET)):
        rows = sb.read_rows(workbook, PLACE[PRIORITY_SHEET], PRIORITY_COLUMNS, section=PRIORITY_SHEET)
        priority = []
        for row_number, row in rows:
            first, second, winner = (status_value(row[key]) for key in ("first", "second", "winner"))
            if _status_norm(winner) not in (_status_norm(first), _status_norm(second)):
                raise sb.SettingsError(
                    f'{sb.location(_label(PRIORITY_SHEET), row_number, "Перемагає")}: має бути один з двох статусів рядка ({first} чи {second}).'
                )
            priority.append({"first": first, "second": second, "winner": winner})
        result["PAYMENT_STATUS_PRIORITY"] = priority
    result.update(sr.read_rules(workbook))
    return result


def with_default_statuses(data):
    """data, де аркушів статусів бракує (None чи немає ключа), - зі стандартними
    статусами (utils/default_statuses.py)."""
    defaults = default_statuses.default_status_data()
    return {**data, **{key: defaults[key] for key in STATUS_KEYS if data.get(key) is None}}


def _code_mapper(statuses):
    """code_of(значення): назва чи код статусу (без урахування регістру й зайвих
    пробілів) -> код; невідоме значення - як є (число - числом)."""
    codes = {}
    for row in statuses:
        code = row.get("code") or row["status"]
        codes[_status_norm(str(code))] = code
        codes[_status_norm(str(row["status"]))] = code

    def code_of(value):
        return status_value(codes.get(_status_norm(value), value)) if isinstance(value, str) else value

    return code_of


def status_rules(data, special_group_name, transfer_heading_marker):
    """Аркуші правил статусів (utils/status_rules.py) -> значення constants.py
    (DEFAULT_STATUS, LEAVE_STATUSES, STATUS_HEADING_RULES, ADDON_* ...) - за
    кодами статусів."""
    data = with_default_statuses(data)
    return sr.build_rules(data, _code_mapper(data["STATUSES"]), special_group_name, transfer_heading_marker)


def status_tables(data):
    """Аркуші статусів -> словники, з якими працює код табеля (constants.py):
    STATUS_MEANINGS, STATUS_COLORS, PAYMENT_STATUS_CONVERSIONS,
    PAYMENT_STATUS_OVERRIDES, PAYMENT_STATUS_FALLBACKS, PAYMENT_STATUS_PRIORITY -
    за кодами статусів (так на них посилаються правила програми), і
    STATUS_NAMES - {код: назва} перейменованих статусів (content/status_names.py).
    Статус з цифр - числом (так він стоїть у клітинці ОБЛІК.xlsx). На аркушах
    замін і пріоритету, як і в колонках виплат, статус можна писати і назвою,
    і кодом."""
    data = with_default_statuses(data)
    statuses = data["STATUSES"]
    code_of = _code_mapper(statuses)

    overrides = {}
    for row in data["PAYMENT_STATUS_OVERRIDES"]:
        overrides.setdefault(code_of(row["status"]), {})[code_of(row["unit_status"])] = code_of(row["value"])
    by_code = [(status_value(row.get("code") or row["status"]), row) for row in statuses]
    return {
        "STATUS_MEANINGS": {str(code): row["meaning"] for code, row in by_code if row["meaning"]},
        "STATUS_COLORS": {code: row["color"] for code, row in by_code if row["color"]},
        "PAYMENT_STATUS_CONVERSIONS": {code: code_of(row["payment"]) for code, row in by_code if row["payment"] is not None},
        "PAYMENT_STATUS_OVERRIDES": overrides,
        "PAYMENT_STATUS_FALLBACKS": {code: code_of(row["fallback"]) for code, row in by_code if row["fallback"] is not None},
        "PAYMENT_STATUS_PRIORITY": {
            frozenset({code_of(row["first"]), code_of(row["second"])}): code_of(row["winner"])
            for row in data["PAYMENT_STATUS_PRIORITY"]
        },
        "STATUS_NAMES": {
            str(code): str(row["status"]) for code, row in by_code if str(row["status"]) != str(code)
        },
    }


def read_settings(path):
    """Книга -> словник форми колишнього resources/data.json (і статуси -
    STATUS_KEYS, read_statuses)."""
    with sb.open_for_reading(path) as workbook:
        general = sb.read_settings_sheet(workbook, PLACE[GENERAL_SHEET], GENERAL_SETTINGS, section=GENERAL_SHEET)
        subdivisions = read_subdivisions(workbook)
        category_100, review, absent, ignore = (
            sb.read_rows(workbook, PLACE[table], columns, section=table, optional_sheet=True)
            for table, columns in (
                (CATEGORY_100_SHEET, CATEGORY_100_COLUMNS), (REVIEW_SHEET, REVIEW_COLUMNS),
                (ABSENT_SHEET, ABSENT_COLUMNS), (IGNORE_SHEET, IGNORE_COLUMNS),
            )
        )
        statuses = read_statuses(workbook)

    groups = {"SPECIAL_TASK_GROUP": general["SPECIAL_TASK_GROUP"], "BCHS_SIGNAL_GROUP_NAME": general["BCHS_SIGNAL_GROUP_NAME"]}
    if general["REPORTS_SIGNAL_GROUP_NAME"]:
        groups["REPORTS_SIGNAL_GROUP_NAME"] = general["REPORTS_SIGNAL_GROUP_NAME"]

    ignore_lists = {key: [] for key, _label in BOOKS}
    for row_number, row in ignore:
        start, end = row["start"], row["end"]
        period = start if not end or end == start else f"{start}-{end}"
        ignore_lists[_book_key(row["book"], row_number)].append({"period": period, "name": row["name"]})

    data = {
        "unit": {key: general[key] for key in UNIT_KEYS},
        "groups": groups,
        "INFORMATION_UNIT_KEEP_FILE_NAME": general["INFORMATION_UNIT_KEEP_FILE_NAME"],
        "company_sections": {"SEQUENCE": subdivisions},
        "CONST_PAYMENT_CATEGORY_100_MAX_DAYS": general["CONST_PAYMENT_CATEGORY_100_MAX_DAYS"],
        "EXCLUDE_NAMES": [
            {"name": row["name"], "days": row["days"], "color": _color_value(row["color"])} for _row_number, row in category_100
        ],
        "EXCLUDE_ABSENT_FROM_INFORMATION_UNIT": _text_rows(absent, "name"),
        "EXCLUDE_FROM_PAYMENTS_REVIEW": [
            {"date": row["date"], "name": row["name"]} for _row_number, row in review if row["payments"]
        ],
        "EXCLUDE_FROM_REVIEW": [
            {"date": row["date"], "name": row["name"]} for _row_number, row in review if row["review"]
        ],
        "IGNORE_LISTS": ignore_lists,
        **statuses,
    }
    if general["BRIGADE_MILITARY_UNITS"]:
        data["BRIGADE_MILITARY_UNITS"] = general["BRIGADE_MILITARY_UNITS"]
    return data


_PERIOD_RE = re.compile(r"\s*(\d{1,2}\.\d{1,2}\.\d{4})\s*(?:[-–—]\s*(\d{1,2}\.\d{1,2}\.\d{4})\s*)?")


def _period_bounds(period):
    """"дд.мм.рррр" / "дд.мм.рррр-дд.мм.рррр" -> (початок, кінець чи None).
    Нерозпізнаний текст лишається в "Початок" як є - тоді книга при читанні
    покаже помилку саме в цьому рядку, а не загубить запис."""
    match = _PERIOD_RE.fullmatch(period) if isinstance(period, str) else None
    if match is None:
        return (str(period) if period is not None else None), None
    return match.group(1), match.group(2)


def _as_list(value):
    if value is None:
        return []
    return [value] if isinstance(value, str) else list(value)


def _general_values(data):
    unit = data.get("unit", {})
    groups = data.get("groups", {})
    values = {key: unit.get(key) for key in UNIT_KEYS}
    values.update({key: groups.get(key) for key in ("SPECIAL_TASK_GROUP", "BCHS_SIGNAL_GROUP_NAME", "REPORTS_SIGNAL_GROUP_NAME")})
    values["INFORMATION_UNIT_KEEP_FILE_NAME"] = data.get("INFORMATION_UNIT_KEEP_FILE_NAME")
    values["CONST_PAYMENT_CATEGORY_100_MAX_DAYS"] = data.get("CONST_PAYMENT_CATEGORY_100_MAX_DAYS")
    values["BRIGADE_MILITARY_UNITS"] = _as_list(data.get("BRIGADE_MILITARY_UNITS"))
    return values


def _review_rows(data):
    """Обидва списки виключень ручної перевірки - в один аркуш: рядок на пару
    (дата, ПІБ), а колонки кажуть, з якого файлу її прибрати."""
    rows = {}
    for key, flag in (("EXCLUDE_FROM_REVIEW", "review"), ("EXCLUDE_FROM_PAYMENTS_REVIEW", "payments")):
        for entry in _as_list(data.get(key)):
            entry = entry if isinstance(entry, dict) else {"name": str(entry)}
            row = rows.setdefault((entry.get("date"), entry.get("name")), {
                "date": entry.get("date"), "name": entry.get("name"), "review": False, "payments": False,
            })
            row[flag] = True
    return list(rows.values())


def _color_cell(color):
    if color is True:
        return sb.YES
    if color in (False, None, ""):
        return sb.NO
    return str(color)


_HEX_COLOR_RE = re.compile(r"(?:[0-9A-Fa-f]{2})?([0-9A-Fa-f]{6})")


def _paint_color_column(sheet, first_row, rows):
    """Клітинки "Колір" таблиці "Ліміт категорії 100" - тим самим кольором, яким
    ПІБ буде залито в ОБЛІК.xlsx: "Так" - жовтий (умовне форматування, тож і
    для нових рядків), код кольору - самим цим кольором."""
    column = sb.column_position(CATEGORY_100_COLUMNS, "color")
    letter = get_column_letter(column)
    cells = f"{letter}{first_row}:{letter}{first_row + len(rows) + 4}"
    sheet.conditional_formatting.add(
        cells, CellIsRule(operator="equal", formula=[f'"{sb.YES}"'], fill=PatternFill("solid", start_color="FFFF00", end_color="FFFF00")),
    )
    sheet.conditional_formatting.add(cells, CellIsRule(operator="equal", formula=[f'"{sb.NO}"'], font=Font(color="8C8C8C")))
    for row_index, row in enumerate(rows, start=first_row):
        match = _HEX_COLOR_RE.fullmatch(str(row.get("color") or ""))
        if match:
            sheet.cell(row=row_index, column=column).fill = PatternFill("solid", start_color=match.group(1).upper())


def _sections(data):
    """{таблиця: sb.Section} - усі таблиці книги з даними data (статусів у data
    немає - стандартні)."""
    data = with_default_statuses(data)
    labels = dict(BOOKS)
    ignore_rows = []
    for key, entries in (data.get("IGNORE_LISTS") or {}).items():
        for entry in _as_list(entries):
            start, end = _period_bounds(entry.get("period"))
            ignore_rows.append({"book": labels.get(key, key), "start": start, "end": end, "name": entry.get("name")})
    category_rows = [
        {"name": entry.get("name"), "days": entry.get("days"), "color": _color_cell(entry.get("color"))}
        for entry in _as_list(data.get("EXCLUDE_NAMES")) if isinstance(entry, dict)
    ]
    sections = {
        REVIEW_SHEET: sb.Section(REVIEW_SHEET, TABLE_HINTS[REVIEW_SHEET], REVIEW_COLUMNS, _review_rows(data)),
        IGNORE_SHEET: sb.Section(IGNORE_SHEET, TABLE_HINTS[IGNORE_SHEET], IGNORE_COLUMNS, ignore_rows, band_by="book"),
        CATEGORY_100_SHEET: sb.Section(CATEGORY_100_SHEET, TABLE_HINTS[CATEGORY_100_SHEET], CATEGORY_100_COLUMNS,
                                       category_rows, paint=_paint_color_column),
        ABSENT_SHEET: sb.Section(ABSENT_SHEET, TABLE_HINTS[ABSENT_SHEET], ABSENT_COLUMNS, [
            {"name": name} for name in _as_list(data.get("EXCLUDE_ABSENT_FROM_INFORMATION_UNIT"))
        ]),
        STATUSES_SHEET: sb.Section(STATUSES_SHEET, TABLE_HINTS[STATUSES_SHEET], STATUS_COLUMNS, [
            {**row, "code": row.get("code") or row["status"], "_dim": ("code",)} for row in data["STATUSES"]
        ]),
        OVERRIDES_SHEET: sb.Section(OVERRIDES_SHEET, TABLE_HINTS[OVERRIDES_SHEET], OVERRIDE_COLUMNS,
                                    data["PAYMENT_STATUS_OVERRIDES"], band_by="status"),
        PRIORITY_SHEET: sb.Section(PRIORITY_SHEET, TABLE_HINTS[PRIORITY_SHEET], PRIORITY_COLUMNS, data["PAYMENT_STATUS_PRIORITY"]),
        GENERAL_SHEET: sb.Section(GENERAL_SHEET, TABLE_HINTS[GENERAL_SHEET], settings=GENERAL_SETTINGS,
                                  values=_general_values(data)),
        SUBDIVISIONS_SHEET: sb.Section(SUBDIVISIONS_SHEET, TABLE_HINTS[SUBDIVISIONS_SHEET], SUBDIVISION_COLUMNS, [
            {"name": name} for name in _as_list(data.get("company_sections", {}).get("SEQUENCE"))
        ]),
    }
    for key, table in sr.RULE_SHEETS.items():
        sections[table] = sr.rule_section(key, data[key])
    return sections


def _write_instructions(workbook):
    contents = [
        (sheet_name, color, [(table, TABLE_HINTS[table].split(". ")[0] + ".") for table in tables])
        for sheet_name, _title, _hint, color, tables in BOOK_SHEETS
    ]
    sb.write_instructions(workbook, INSTRUCTIONS_TITLE, INSTRUCTIONS_INTRO, INSTRUCTIONS_TASKS, contents, INSTRUCTIONS_TIPS)


def write_settings(path, data):
    """Словник форми data.json -> нова книга path (наявний файл замінюється).
    Статусів у data немає - стандартні (utils/default_statuses.py)."""
    sections = _sections(data)
    workbook = sb.new_workbook()
    for sheet_name, title, hint, color, tables in BOOK_SHEETS:
        sb.write_sections(workbook, sheet_name, [sections[table] for table in tables], title=title, hint=hint, tab_color=color)
    _write_instructions(workbook)
    sb.save_workbook(workbook, path)


def _normalized(value):
    """Значення для порівняння data.json з книгою: книга обрізає пробіли на
    краях тексту й зберігає перенос рядка як "\n"."""
    if isinstance(value, str):
        return value.replace("\r\n", "\n").strip()
    if isinstance(value, dict):
        return {key: _normalized(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_normalized(item) for item in value]
    return value


def _expected_from_legacy(data):
    """data.json у тій формі, яку дає read_settings: невикористаний
    company_sections.MANAGEMENT відкинуто, порожні необов'язкові параметри не
    записуються, "color": null - те саме, що false, статусів data.json не мав -
    у книзі стандартні."""
    expected = with_default_statuses(json.loads(json.dumps(data, ensure_ascii=False)))
    expected["company_sections"] = {"SEQUENCE": _as_list(expected.get("company_sections", {}).get("SEQUENCE"))}
    expected["groups"] = {key: value for key, value in expected.get("groups", {}).items() if value}
    expected["EXCLUDE_ABSENT_FROM_INFORMATION_UNIT"] = _as_list(expected.get("EXCLUDE_ABSENT_FROM_INFORMATION_UNIT"))
    for key in ("EXCLUDE_NAMES", "EXCLUDE_FROM_PAYMENTS_REVIEW", "EXCLUDE_FROM_REVIEW"):
        expected[key] = _as_list(expected.get(key))
    for entry in expected["EXCLUDE_NAMES"]:
        if isinstance(entry, dict):
            entry["color"] = entry.get("color") or False
            entry["days"] = None if entry.get("days") in (None, "") else entry.get("days")
    ignore_lists = expected.get("IGNORE_LISTS") or {}
    expected["IGNORE_LISTS"] = {
        key: [{**entry, "period": "".join(str(entry.get("period", "")).split())} for entry in _as_list(ignore_lists.get(key))]
        for key, _label in BOOKS
    }
    if not expected.get("BRIGADE_MILITARY_UNITS"):
        expected.pop("BRIGADE_MILITARY_UNITS", None)
    else:
        expected["BRIGADE_MILITARY_UNITS"] = _as_list(expected["BRIGADE_MILITARY_UNITS"])
    return expected


def migration_differences(data, settings):
    """Ключі data.json, значення яких у книзі (settings - read_settings) не
    збіглося. Порожній список - перенесено без втрат."""
    expected = _normalized(_expected_from_legacy(data))
    actual = _normalized(settings)
    return sorted(key for key in set(expected) | set(actual) if expected.get(key) != actual.get(key))


def migrate_legacy_json(json_path, settings_path):
    """Колишній resources/data.json -> книга settings_path. Повертає
    (словник data.json, ключі, що не збіглися після читання книги назад)."""
    with open(json_path, encoding="utf-8") as json_file:
        data = json.load(json_file)
    write_settings(settings_path, data)
    return data, migration_differences(data, read_settings(settings_path))


def _with_equivalent_notes(data):
    """Групи "Рівнозначні статуси" без довідки (книга, записана до її появи) -
    з довідкою стандартних груп (utils/default_statuses.py)."""
    notes = {
        frozenset(_status_norm(str(status)) for status in row["statuses"]): row["note"]
        for row in default_statuses.default_rule_data()["EQUIVALENT_STATUSES"]
    }
    rows = [
        {**row, "note": row.get("note") or notes.get(frozenset(_status_norm(str(status)) for status in row["statuses"]), "")}
        for row in data["EQUIVALENT_STATUSES"]
    ]
    return {**data, "EQUIVALENT_STATUSES": rows}


def _upgrade_layout(settings_path, data):
    """Книга не в теперішньому вигляді (кожна таблиця - на власному аркуші, чи
    якихось таблиць статусів ще немає) один раз переписується в теперішній:
    ті самі дані, відсутні таблиці статусів - стандартні. Копія колишньої книги
    лишається поруч ("... (до оновлення).xlsx"); якщо переписана книга, прочитана
    назад, у чомусь не збіглася - колишня повертається на місце. Книга відкрита
    в Excel - спроба при наступному запуску. Повертає дані для constants.py."""
    upgraded = _with_equivalent_notes(with_default_statuses(data))
    if not sb.missing_tables(settings_path, BOOK_TABLES):
        return upgraded
    # Копія, що лишилась від попереднього оновлення, не перезаписується.
    backup = sb.free_backup_path(settings_path)
    try:
        shutil.copy2(settings_path, backup)
        # Запис атомарний (sb.save_workbook): не вдався - книга та сама.
        write_settings(settings_path, upgraded)
    except (sb.SettingsError, OSError) as error:
        print_red(f"Вигляд {settings_path} не оновлено ({error}) - спробую при наступному запуску.")
        return upgraded
    try:
        read_back = read_settings(settings_path)
    except sb.SettingsError:
        read_back = None
    if read_back != upgraded:
        os.replace(backup, settings_path)
        print_red(f"{settings_path}: оновлена книга не збіглася з колишньою - колишню повернуто, вигляд не змінено.")
        return upgraded
    print_green(
        f"{settings_path}: книгу переведено на новий вигляд (кілька таблиць на аркуші, дані ті самі). "
        f"Копія колишньої - {backup}."
    )
    return upgraded


def _migrate_or_stop(legacy_json_path, settings_path):
    """migrate_legacy_json з поясненням для користувача (sb.stop), якщо щось
    не так. Реальний випадок: у старому data.json не було "Ліміт днів
    категорії 100" - книгу записано, а прочитати назад не вдалося, і
    користувач бачив лише червоне "Програма завершилась з помилкою"."""
    try:
        _data, differences = migrate_legacy_json(legacy_json_path, settings_path)
    except sb.SettingsError as error:
        if os.path.isfile(settings_path):
            raise sb.stop(
                f"Налаштування з {legacy_json_path} перенесено в {settings_path}, але одне значення треба вписати "
                f"вручну (у старому data.json його не було або воно неправильне):\n{error}",
                sb.fix_in_book_steps(settings_path),
            ) from error
        raise _legacy_json_stop(legacy_json_path, settings_path, error) from error
    except OSError as error:
        raise sb.stop(
            f"Не вдалося перенести {legacy_json_path} у {settings_path}: {error}",
            (f"Закрийте {sb.SETTINGS_FILE_NAME} і data.json, якщо вони відкриті.", "Запустіть програму ще раз."),
        ) from error
    except ValueError as error:
        raise _legacy_json_stop(legacy_json_path, settings_path, error) from error
    if differences:
        print_red(
            f"Налаштування перенесено з {legacy_json_path} у {settings_path}, але не все збіглося: "
            f"{', '.join(differences)}. data.json лишено - перевірте ці розділи книги."
        )
    else:
        os.remove(legacy_json_path)
        print_green(
            f"Налаштування перенесено з {legacy_json_path} у {settings_path} без втрат, data.json видалено. "
            "Далі редагуйте книгу Excel."
        )


def _legacy_json_stop(legacy_json_path, settings_path, error):
    return sb.stop(
        f"Не вдалося перенести {legacy_json_path} у {settings_path}: {error}",
        (f"Виправте {legacy_json_path} (відкривається Блокнотом) або видаліть його й створіть книгу-зразок: "
         "python resources_example/generate_sample.py (з теки generator_timesheet) і впишіть у неї свої дані.",
         "Запустіть програму ще раз."),
    )


def load_settings(settings_path, legacy_json_path=None):
    """Словник налаштувань для constants.py. Книги ще немає, а є колишній
    data.json - спершу переносить його в книгу (одноразово, _migrate_or_stop).
    Якщо книга, прочитана назад, збігається з data.json, data.json видаляється
    (дані вже в книзі); інакше лишається, а користувач бачить, які розділи не
    збіглися. Немає ні того, ні іншого, чи книга з помилкою - програма
    зупиняється з червоним поясненням, що зробити (sb.stop)."""
    if not os.path.isfile(settings_path):
        if legacy_json_path and os.path.isfile(legacy_json_path):
            _migrate_or_stop(legacy_json_path, settings_path)
        else:
            raise sb.stop(
                f"Немає файлу налаштувань {settings_path}.",
                ("Створіть зразок: python resources_example/generate_sample.py (з теки generator_timesheet).",
                 "Впишіть у нього дані свого підрозділу.", "Запустіть програму ще раз."),
            )
    try:
        data = read_settings(settings_path)
    except sb.SettingsError as error:
        raise sb.stop(f"У книзі налаштувань помилка:\n{error}", sb.fix_in_book_steps(settings_path)) from error
    return _upgrade_layout(settings_path, data)


@contextmanager
def stop_on_invalid_settings(settings_path):
    """Помилка значення книги, яку constants.py знаходить під час імпорту
    (ValueError, текст починається з назви книги - source_of), - червоне
    пояснення з кроками (sb.stop) замість трасування Python. Інші помилки -
    як були."""
    try:
        yield
    except ValueError as error:
        if not str(error).startswith(sb.SETTINGS_FILE_NAME):
            raise
        raise sb.stop(f"У книзі налаштувань помилка:\n{error}", sb.fix_in_book_steps(settings_path)) from error

