import json
import os
import re
from datetime import datetime

from utils import default_statuses, project_settings
from utils.logging_utils import print_red
from utils.settings_book import SETTINGS_FILE_NAME as _SETTINGS_BASENAME, SettingsError

# Шляхи прив'язані до розташування цього файлу (__file__), а не до поточної
# робочої теки, як колись (RESOURCES_DIR = "resources"). Під час збирання
# тестів (кореневий conftest.py для кількох проєктів) cwd на момент імпорту
# цього модуля не обов'язково generator_timesheet/, і відносний шлях міг
# мовчки відкрити налаштування іншого проєкту. Абсолютний шлях працює за
# будь-якого cwd.
_MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
RESOURCES_DIR = os.path.join(_MODULE_DIR, "resources")

# Усе, що налаштовується під підрозділ (назви, групи Signal, списки людей), -
# у книзі Excel resources/НАЛАШТУВАННЯ.xlsx, а не тут: код оновлюється з
# дзеркала через "git reset --hard" (scripts/common.ps1), а resources/ git не
# чіпає. utils/project_settings.py читає книгу в словник тієї самої форми,
# що й колишній resources/data.json, тож _DATA нижче розбирається як і
# раніше. Колишній data.json при першому запуску переноситься в книгу сам.
SETTINGS_FILE_NAME = os.path.join(RESOURCES_DIR, _SETTINGS_BASENAME)
LEGACY_DATA_FILE_NAME = os.path.join(RESOURCES_DIR, "data.json")
_DATA = project_settings.load_settings(SETTINGS_FILE_NAME, LEGACY_DATA_FILE_NAME)

REPORT_DIR = os.path.join(RESOURCES_DIR, "report")
PERSONEL_LIST_FILE_NAME = os.path.join(RESOURCES_DIR, "ОБЛІК.xlsx")
PERSONEL_LIST_SHEET_NAME = "Табель"

OUTPUT_DIR = "output"
OUTPUT_FILE_NAME = os.path.join(OUTPUT_DIR, "ОБЛІК.xlsx")
REVIEW_FILE_NAME = os.path.join(OUTPUT_DIR, "Потребує_ручної_перевірки.txt")

# Аркуш "ТАБЕЛЬ"/"Табель" (без урахування регістру) у кожному файлі цієї теки -
# власна чернетка підрозділу з категорією виплати (число) чи статусом (текст)
# на кожен день. Джерело для
# generate_payments_timesheet (content/information_unit_reader.py).
INFORMATION_UNIT_DIR = os.path.join(RESOURCES_DIR, "information_unit")
INFORMATION_UNIT_SHEET_NAME = "табель"

# Перед звичайним запуском index.py можна окремо завантажити файли БЧС за
# обрану дату з групи Signal BCHS_SIGNAL_GROUP_NAME просто в INFORMATION_UNIT_DIR
# (content/signal_bchs.py). Тека спершу очищається, крім
# INFORMATION_UNIT_KEEP_FILE_NAME: цей файл не має щоденної дати й через Signal
# не приходить. Назва групи й файлу-винятку - справжні, тож, як і
# SPECIAL_TASK_GROUP/підрозділи, беруться з НАЛАШТУВАННЯ.xlsx (аркуш "Загальні",
# не під git), а не пишуться тут. "БЧС" - не назва підрозділу, а загальний тип
# донесення (як і "ТАБЕЛЬ"), тож лишається в коді.
BCHS_SIGNAL_GROUP_NAME = _DATA["groups"]["BCHS_SIGNAL_GROUP_NAME"]
BCHS_FILENAME_MARKER = "БЧС"
INFORMATION_UNIT_KEEP_FILE_NAME = _DATA["INFORMATION_UNIT_KEEP_FILE_NAME"]
# Підрозділи, від яких щодня чекаємо файл БЧС у групі Signal (НАЛАШТУВАННЯ.xlsx,
# таблиця "Підрозділи", у цьому порядку). Тих, чийого файлу немає, після
# завантаження друкуємо червоним. Управління і штаб - окремий файл
# INFORMATION_UNIT_KEEP_FILE_NAME, не з Signal.
BCHS_EXPECTED_SUBDIVISIONS = tuple(value.strip() for value in _DATA["company_sections"]["SEQUENCE"])

# Підтверджена вручну Відомість РОП/ВОП (того самого формату, що генерує
# content/rop_vop_statement.py), яку подають щодня протягом місяця. З нею
# звіряється обчислена з information_unit версія перед збереженням
# (content/schedule_reader.py).
SCHEDULE_DIR = os.path.join(RESOURCES_DIR, "schedule")

# Щоденні рапорти (REPORT_DIR) і відомості ВОП-РОП (SCHEDULE_DIR) можна звірити
# з групою Signal на старті index.py (content/signal_reports.py). Назва групи -
# справжня, тож з НАЛАШТУВАННЯ.xlsx; None - звіряння не налаштовано.
REPORTS_SIGNAL_GROUP_NAME = _DATA["groups"].get("REPORTS_SIGNAL_GROUP_NAME")
# Підтека для попередніх версій файлів, замінених новішими з Signal. Назва
# починається з крапки, тож glob("*") генераторів її не бачить.
SIGNAL_REPLACED_VERSIONS_DIR_NAME = ".попередні версії"

# Тека сторонніх книг для звірки ("Перевірка: звірка Облік.xlsx з addons" в
# index.py). Звіряються всі книги Excel цієї теки, без діалогу вибору. Тип
# книги (book61 - content/addon_reader.py, ЦВС - content/cvs_reader.py, журнал
# обліку БП - content/training_journal_reader.py) визначається за аркушами й
# заголовками, а не за іменем файлу (generators/check_addon_mismatches.py).
ADDONS_DIR = os.path.join(RESOURCES_DIR, "addons")

# "воп"/"тмп"/"тот" - різні слова, якими resources/schedule позначає ту саму
# категорію виплати 170 (ВОП), тоді як generator_timesheet пише лише "воп"
# (content/rop_vop_statement.py, _VOP_CATEGORY). Перелік веде сам користувач,
# як і PAYMENT_STATUS_PRIORITY: якщо відомість каже одне з цих слів, а
# обчислено "воп", це не розбіжність (_schedule_statuses_equivalent).
SCHEDULE_VOP_ALIASES = frozenset({"воп", "тмп", "тот"})
# Той самий файл "ОБЛІК.xlsx", що й OUTPUT_FILE_NAME (а не "ОБЛІК для
# виплат.xlsx"). Конфлікту імен немає: generate_timesheet і
# generate_payments_timesheet завжди запускаються окремо, а index.py очищає
# OUTPUT_DIR перед кожним прогоном. ("Згенерувати все" дає стройовому ОБЛІК
# власну назву.)
PAYMENTS_OUTPUT_FILE_NAME = OUTPUT_FILE_NAME
PAYMENTS_REVIEW_FILE_NAME = os.path.join(OUTPUT_DIR, "Потребує_ручної_перевірки_виплати.txt")

# Звіт розбіжностей (content/payment_mismatch_checker.py): для кожної людини й
# дати ОБЛІК для виплат звіряється з кожним файлом INFORMATION_UNIT_DIR окремо.
# Пишеться автоматично в межах тієї самої генерації
# (generate_payments_timesheet), а не окремим пунктом меню.
PAYMENT_MISMATCH_REPORT_FILE_NAME = os.path.join(OUTPUT_DIR, "error_mis_statuses.xlsx")
# Разом з error_mis_statuses.xlsx і error_mis_statuses_Відомість.xlsx - один
# файл з готовим текстом повідомлення командирам підрозділів, що виправити у
# своїх файлах БЧС (content/commander_message.py; за відомістю - 70 замість
# РОП, 170 замість ВОП).
COMMANDERS_MESSAGE_FILE_NAME = os.path.join(OUTPUT_DIR, "Повідомлення_командирам.txt")
CATEGORY_100_CHECK_FILE_NAME = os.path.join(OUTPUT_DIR, "Перевірка_категорії_100.xlsx")

# Звіт розбіжностей ОБЛІК.xlsx з книгами resources/addons (ADDONS_DIR,
# content/addon_mismatch_checker.py) - окремий пункт меню index.py чи питання
# після генерації.
ADDON_MISMATCH_REPORT_FILE_NAME = os.path.join(OUTPUT_DIR, "mis_match.xlsx")
# Звірка книг resources/addons між собою (без Табеля) - окреме перше питання
# index.py; "Так" - звірка, і на цьому все.
ADDONS_CROSS_CHECK_REPORT_FILE_NAME = os.path.join(OUTPUT_DIR, "addons_mis_match.xlsx")

# Типи книг resources/addons (впізнаються за вмістом) -> назва в повідомленнях.
# Книги будь-якого типу може не бути: тоді звіряються решта.
ADDON_TREATMENT_JOURNAL_KIND = "treatment"
ADDON_BOOK_KINDS = {
    "book61": "book61", "cvs": "ЦВС", "journal": "журнал обліку БП", ADDON_TREATMENT_JOURNAL_KIND: "журнал лікування",
}

# Дата в книзі addons пізніше за сьогодні більш ніж на стільки днів - майже
# напевно помилка в році (реальні випадки: лікування з 30.06.2026 з випискою
# "05.07.2027", відпустка до "03.12.2028" - через них були хибні розбіжності
# на місяці вперед). Законні майбутні кінці (відпустка, навчання) - до ~40
# днів. Така дата мовчки не виправляється, а лише показується у звіті
# addons_mis_match.xlsx і в консолі - виправити треба саму книгу.
ADDON_IMPLAUSIBLE_FUTURE_DAYS = 180

# Журнал обліку БП (resources/addons, content/training_journal_reader.py): це
# значення колонки "НЦ" у таблиці "Журнал обліку БП" книги налаштувань означає
# будь-який запис з номером власної в/ч (OWN_MILITARY_UNIT_CODES, з
# НАЛАШТУВАННЯ.xlsx). Аркуш журналу -> статус - ADDON_TRAINING_JOURNAL_SHEETS нижче.
ADDON_JOURNAL_OWN_UNIT = "власна в/ч"

# Готовий "ОБЛІК для виплат" (PAYMENTS_OUTPUT_FILE_NAME) можна записати ще
# кудись - зазвичай у сусідній проєкт generator_br_and_report_for_money, що
# генерує рапорти на додаткову винагороду (export_for_money_project; index.py
# пропонує це лише після повністю чистого прогону). Теку користувач обирає
# діалогом (utils/folder_picker.py), а це лише початкова тека діалогу, не
# фіксоване місце призначення. Шлях абсолютний (від _MODULE_DIR), як і
# RESOURCES_DIR.
MONEY_PROJECT_RESOURCES_DIR = os.path.join(_MODULE_DIR, "..", "generator_br_and_report_for_money", "resources")

# Звірка днів РОП/ВОП (find_schedule_mismatches): бойові роти, батареї й взводи
# зі спільного з generator_br_and_report_for_money переліку (таблиця "Підрозділи"
# НАЛАШТУВАННЯ.xlsx того проєкту - той самий вигляд, що й власний), без штабу й
# управління. Людина поза ними звіряється, лише якщо вона є і в ОБЛІК.xlsx, і у
# відомості. Перелік читається напряму з книги того проєкту, а не з його
# Python-коду (пакети двох проєктів навмисно розділені, див.
# MONTH_NAMES_NOMINATIVE_UPPER), щоб не тримати дві копії, які треба
# синхронізувати вручну. Поки той проєкт не запускався після переходу на книгу,
# перелік береться з його колишнього resources/data.json.
# resources/ того проєкту не під git і може бути не налаштований (напр., на
# машині, де генерують лише табель) - тоді перелік береться з власного аркуша
# "Підрозділи" (той самий, що й для BCHS_EXPECTED_SUBDIVISIONS вище), а не
# падає імпорт усього index.py.
def _money_project_subdivisions():
    settings_path = os.path.join(MONEY_PROJECT_RESOURCES_DIR, _SETTINGS_BASENAME)
    if os.path.isfile(settings_path):
        try:
            return project_settings.read_subdivisions_file(settings_path)
        except SettingsError as error:
            print_red(f"Перелік підрозділів з {settings_path} не прочитано ({error}) - беру власний аркуш \"Підрозділи\".")
            return None
    legacy_path = os.path.join(MONEY_PROJECT_RESOURCES_DIR, "data.json")
    if os.path.isfile(legacy_path):
        with open(legacy_path, encoding="utf-8") as legacy_file:
            return json.load(legacy_file)["company_sections"]["SEQUENCE"]
    return None


RECOGNIZED_SUBDIVISIONS = frozenset(
    value.strip().lower() for value in (_money_project_subdivisions() or _DATA["company_sections"]["SEQUENCE"])
)
# Той самий словник назв місяців (називний відмінок, великі літери), що й
# generator_br_and_report_for_money/content/money_report_helpers.py::
# MONTH_NAMES_NOMINATIVE_UPPER. Звідти не імпортується - у кожного проєкту свої
# пакети з однаковими назвами (див. кореневий conftest.py), - тож продубльовано
# тим самим написанням.
MONTH_NAMES_NOMINATIVE_UPPER = {
    1: "СІЧЕНЬ", 2: "ЛЮТИЙ", 3: "БЕРЕЗЕНЬ", 4: "КВІТЕНЬ", 5: "ТРАВЕНЬ", 6: "ЧЕРВЕНЬ",
    7: "ЛИПЕНЬ", 8: "СЕРПЕНЬ", 9: "ВЕРЕСЕНЬ", 10: "ЖОВТЕНЬ", 11: "ЛИСТОПАД", 12: "ГРУДЕНЬ",
}

# Точна назва аркуша, яку generator_br_and_report_for_money/utils/
# excel_reader.py::_resolve_sheet_name шукає першою (PERSONEL_LIST_SHEET_NAME у
# тому проєкті); інакше він бере запасний аркуш і читає не ті дані. Той проєкт
# тримає в цьому аркуші дані поточного місяця, тож назва - назва поточного
# місяця. Колись тут стояло фіксоване "ТРАВЕНЬ" (місяць, коли писався
# експорт), тепер назва обчислюється від поточної дати.
MONEY_PROJECT_SHEET_NAME = MONTH_NAMES_NOMINATIVE_UPPER[datetime.now().month]

# НАЛАШТУВАННЯ.xlsx (аркуш "Загальні") - для заголовка й імені файлу Відомості РОП/ВОП
# (content/rop_vop_statement.py). Текст зразка (реальний
# resources/{SHORT_UNIT_BATTALION без пробілу}ВОП-РОП_ЛИПЕНЬ_.xlsx) збігається з
# цими полями: "{FULL_UNIT_BUT} {FULL_MILITARY_UNIT} за МІСЯЦЬ РІК року" у
# заголовку, "{SHORT_UNIT_BATTALION без пробілу}ВОП-РОП_МІСЯЦЬ_.xlsx" в імені.
FULL_UNIT_BUT = _DATA["unit"]["FULL_UNIT_BUT"]
FULL_MILITARY_UNIT = _DATA["unit"]["FULL_MILITARY_UNIT"]
SHORT_UNIT_BATTALION = _DATA["unit"]["SHORT_UNIT_BATTALION"]
# Цифри номера власної військової частини (з FULL_MILITARY_UNIT, а не в коді).
# Книга ЦВС: лікувальний заклад власної частини - лікування в медроті (content/cvs_reader.py).
OWN_MILITARY_UNIT_CODES = frozenset(re.findall(r"[АA]\s*(\d{4})", FULL_MILITARY_UNIT))

# Маркер заголовка "Переведені:" (STATUS_HEADING_RULES). Справжній статус рядка
# залежить від того, куди переведено (колонка "Куди вибув",
# content/daily_report_reader.py::_transfer_status).
TRANSFER_HEADING_MARKER = "<переведення>"
# Номери військових частин у складі бригади (НАЛАШТУВАННЯ.xlsx, аркуш
# "Загальні", BRIGADE_MILITARY_UNITS, напр. "А0000, А0001" - позначення, тож не в
# коді). Переведення "до військової частини А..." з таким номером (чи з номером
# власної частини) - ПРВДБ, з будь-яким іншим - ПРВДВЧ.
BRIGADE_MILITARY_UNIT_CODES = OWN_MILITARY_UNIT_CODES | frozenset(
    code for value in _DATA.get("BRIGADE_MILITARY_UNITS", []) for code in re.findall(r"[АA]\s*(\d{4})", str(value))
)

# Іменовані (не-дата) колонки аркуша "Табель" визначаються за назвою заголовка,
# а не за позицією: схожі на ОБЛІК.xlsx файли можуть мати їх в іншому порядку
# чи взагалі без ПІДРОЗДІЛ (як у зразку тут).
LABEL_COLUMN_NAMES = ("ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ")
PIB_COLUMN_NAME = "ПІБ"

# Назва спеціальної групи (напр. розвідувально-тактичної) - з НАЛАШТУВАННЯ.xlsx
# (як і SIGNATORY_UNIT_NAMES), а не в коді: це чутлива оперативна інформація,
# а не стала частина логіки. lower() одразу тут, бо STATUS_HEADING_RULES
# звіряється з уже приведеним до нижнього регістру текстом заголовка
# (_status_for_heading).
_SPECIAL_GROUP_NAME = _DATA["groups"]["SPECIAL_TASK_GROUP"].lower()

# Рядок під заголовком повернення (за STATUS_HEADING_RULES "з ..." - повернення
# в район виконання завдань, DEFAULT_STATUS) насправді може описувати
# повернення в пункт постійної дислокації: колонка "Куди прибув"/"Тип
# відпустки, куди прибув" тоді прямо каже "...до пункту постійної
# дислокації...". Текст рядка (_table_events) має перевагу над заголовком -
# як і заголовок "до пункту постійної дислокації", лише джерело - колонка
# "куди", а не заголовок таблиці.
#
# Береться "до пункту постійної дислокації", а не просто "постійної
# дислокації", бо напрямок важливий: рядок під "В РТГР:" з "Куди вибув" -
# "Вибув з пункту постійної дислокації в РТГР" (людина залишає ППД, а не
# прибуває) має лишитись статус за заголовком.
#
# "в пункт постійної дислокації" - друга форма того самого прибуття: рядок під
# "З проходження військово-лікарської комісії:" (зазвичай DEFAULT_STATUS) з
# "Прибув в пункт постійної дислокації ... з військової частини ..." - людина
# прибула в пункт постійної дислокації. "з пункту постійної дислокації" (вибуття звідти)
# сюди не потрапляє: "в пункт" і "з пункту" не містять одне одного.
#
# "до пункту постійно дислокації" (без "ї") - реальна описка: рядок під "З
# відпустки:" - "Прибув до пункту постійно дислокації з щорічної відпустки".
# Той самий напрямок прибуття (як і варіант "постійно"/"постійної" в
# ARRIVAL_LIST_HEADING_KEYWORDS, див.
# test_arrival_list_heading_accepts_the_adjectival_wording_too).
PPD_DESTINATION_KEYWORDS = (
    "до пункту постійної дислокації",
    "в пункт постійної дислокації",
    "до пункту постійно дислокації",
)

# Статуси, їхні значення, кольори, виплати й правила - лише з книги
# налаштувань (аркуші "Статуси", "Заміни для виплат", "Пріоритет статусів" і
# аркуші правил - utils/status_rules.py): їх змінює сам користувач в Excel, у
# коді жодного статусу немає. Стандартні значення й пояснення, чому вони саме такі, -
# utils/default_statuses.py (ними заповнюється книга, де цих аркушів ще немає).
# Тести логіки табеля працюють саме зі стандартними статусами, а не з тими, що
# налаштовано в книзі (TIMESHEET_STATUSES_FROM_CODE, tests/conftest.py).
#
# - STATUS_MEANINGS: {статус: що означає} - лише довідка;
# - STATUS_COLORS: {статус: "FFRRGGBB"} - заливка клітинки ОБЛІК.xlsx
#   (content/oblik_timesheet.py); категорії виплати - числами (100, а не "100");
#   статус без кольору лишається без заливки - це не помилка;
# - PAYMENT_STATUS_CONVERSIONS: {статус: значення} - що завжди пишеться в ОБЛІК
#   для виплат замість статусу (apply_payment_values), хоч би що казали файли БЧС;
# - PAYMENT_STATUS_OVERRIDES: {статус в ОБЛІК: {що одностайно подає файл
#   information_unit: що записати}} - заміна, коли підрозділ уточнює причину;
# - PAYMENT_STATUS_FALLBACKS: {статус: значення} - що записується, якщо
#   PAYMENT_STATUS_OVERRIDES статус не замінив (новоприбулі -> 10);
# - PAYMENT_STATUS_PRIORITY: {frozenset({статус1, статус2}): переможець} - хто
#   перемагає, коли на одну дату трапились обидва.
#
# Ключі й значення - коди статусів (колонка "Код" аркуша "Статуси"): на них
# посилаються й правила нижче (ADDON_*, LEAVE_STATUSES ...), тож перейменування
# статусу в книзі цих правил не ламає.
_STATUS_SOURCE = default_statuses.default_status_data() if os.environ.get("TIMESHEET_STATUSES_FROM_CODE") == "1" else _DATA
_STATUS_TABLES = project_settings.status_tables(_STATUS_SOURCE)
STATUS_MEANINGS = _STATUS_TABLES["STATUS_MEANINGS"]
STATUS_COLORS = _STATUS_TABLES["STATUS_COLORS"]
PAYMENT_STATUS_CONVERSIONS = _STATUS_TABLES["PAYMENT_STATUS_CONVERSIONS"]
PAYMENT_STATUS_OVERRIDES = _STATUS_TABLES["PAYMENT_STATUS_OVERRIDES"]
PAYMENT_STATUS_FALLBACKS = _STATUS_TABLES["PAYMENT_STATUS_FALLBACKS"]
PAYMENT_STATUS_PRIORITY = _STATUS_TABLES["PAYMENT_STATUS_PRIORITY"]
# {код: назва} статусів, які користувач перейменував на аркуші "Статуси":
# правила вище й увесь код працюють з кодами, а назва - лише в клітинках
# ОБЛІК.xlsx і звітах (content/status_names.py).
STATUS_NAMES = _STATUS_TABLES["STATUS_NAMES"]

# Правила статусів (utils/status_rules.py, пояснення стандартних значень -
# utils/default_statuses.py). Аркуш книги -> назви нижче:
# - "Групи статусів": DEFAULT_STATUS (у районі виконання завдань - статус дня
#   без подій), AFTER_HOURS_ARRIVAL_STATUS (прибув після 17:00), VLK_STATUS,
#   ROZP_STATUS, ADDON_LONG_TREATMENT_STATUS, TRANSFER_TO_BRIGADE_STATUS /
#   TRANSFER_TO_OTHER_UNIT_STATUS (переведення: куди - за BRIGADE_MILITARY_UNIT_CODES),
#   LEGACY_TRANSFER_STATUS, переліки LEAVE_STATUSES, STATUSES_OVER_ROZP,
#   PAYMENT_DUTY_CATEGORIES, PAYMENT_MISMATCH_COMPARABLE_STATUSES,
#   PAYMENT_ABSENT_FROM_FILES_EXEMPT_STATUSES, ADDON_NOT_RETURNED_AFTER_SZCH_STATUSES,
#   ADDON_ARRIVAL_DAY_STATUSES / ADDON_ARRIVAL_DAY_SHEETS, "будь-який з"
#   (ADDON_ONE_OF_STATUSES і його ключі ADDON_*_STATUS), STATUS_CONFLICT_PAID_BY_DEFAULT,
#   STATUS_CONFLICT_PRIORITY_EXEMPT;
# - "Заголовки рапорту": STATUS_HEADING_RULES (заголовок таблиці -> статус;
#   перший збіг перемагає, порядок рядків важливий) і IGNORED_TABLE_HEADING_KEYWORDS;
# - "Речення рапорту": FREEFORM_STATUS_PATTERNS;
# - "Конфлікт у рапортах": STATUS_CONFLICT_OVERRIDES;
# - "Звірка з addons": ADDON_REQUIRED_BOOK61_SHEETS, ADDON_TRACKED_STATUSES,
#   ADDON_TABLE_PRIORITY_STATUSES, ADDON_NESTED_STATUSES;
# - "Рівнозначні статуси": ADDON_EQUIVALENT_STATUS_GROUPS;
# - "Журнал обліку БП": ADDON_TRAINING_JOURNAL_SHEETS.
_RULES = project_settings.status_rules(_STATUS_SOURCE, _SPECIAL_GROUP_NAME, TRANSFER_HEADING_MARKER)
DEFAULT_STATUS = _RULES["DEFAULT_STATUS"]
AFTER_HOURS_ARRIVAL_STATUS = _RULES["AFTER_HOURS_ARRIVAL_STATUS"]
VLK_STATUS = _RULES["VLK_STATUS"]
ROZP_STATUS = _RULES["ROZP_STATUS"]
ADDON_LONG_TREATMENT_STATUS = _RULES["ADDON_LONG_TREATMENT_STATUS"]
TRANSFER_TO_BRIGADE_STATUS = _RULES["TRANSFER_TO_BRIGADE_STATUS"]
TRANSFER_TO_OTHER_UNIT_STATUS = _RULES["TRANSFER_TO_OTHER_UNIT_STATUS"]
TRANSFER_STATUSES = _RULES["TRANSFER_STATUSES"]
LEGACY_TRANSFER_STATUS = _RULES["LEGACY_TRANSFER_STATUS"]
LEAVE_STATUSES = _RULES["LEAVE_STATUSES"]
STATUSES_OVER_ROZP = _RULES["STATUSES_OVER_ROZP"]
PAYMENT_DUTY_CATEGORIES = _RULES["PAYMENT_DUTY_CATEGORIES"]
PAYMENT_MISMATCH_COMPARABLE_STATUSES = _RULES["PAYMENT_MISMATCH_COMPARABLE_STATUSES"]
PAYMENT_ABSENT_FROM_FILES_EXEMPT_STATUSES = _RULES["PAYMENT_ABSENT_FROM_FILES_EXEMPT_STATUSES"]
ADDON_NOT_RETURNED_AFTER_SZCH_STATUSES = _RULES["ADDON_NOT_RETURNED_AFTER_SZCH_STATUSES"]
ADDON_ARRIVAL_DAY_STATUSES = _RULES["ADDON_ARRIVAL_DAY_STATUSES"]
ADDON_ARRIVAL_DAY_SHEETS = _RULES["ADDON_ARRIVAL_DAY_SHEETS"]
ADDON_NON_COMBAT_TREATMENT_STATUS = _RULES["ADDON_NON_COMBAT_TREATMENT_STATUS"]
ADDON_KILLED_WITHOUT_CERTIFICATE_STATUS = _RULES["ADDON_KILLED_WITHOUT_CERTIFICATE_STATUS"]
ADDON_TREATMENT_JOURNAL_STATUS = _RULES["ADDON_TREATMENT_JOURNAL_STATUS"]
ADDON_ONE_OF_STATUSES = _RULES["ADDON_ONE_OF_STATUSES"]
STATUS_CONFLICT_PAID_BY_DEFAULT = _RULES["STATUS_CONFLICT_PAID_BY_DEFAULT"]
STATUS_CONFLICT_PRIORITY_EXEMPT = _RULES["STATUS_CONFLICT_PRIORITY_EXEMPT"]
STATUS_HEADING_RULES = _RULES["STATUS_HEADING_RULES"]
IGNORED_TABLE_HEADING_KEYWORDS = _RULES["IGNORED_TABLE_HEADING_KEYWORDS"]
FREEFORM_STATUS_PATTERNS = _RULES["FREEFORM_STATUS_PATTERNS"]
STATUS_CONFLICT_OVERRIDES = _RULES["STATUS_CONFLICT_OVERRIDES"]
ADDON_REQUIRED_BOOK61_SHEETS = _RULES["ADDON_REQUIRED_BOOK61_SHEETS"]
ADDON_TRACKED_STATUSES = _RULES["ADDON_TRACKED_STATUSES"]
ADDON_TABLE_PRIORITY_STATUSES = _RULES["ADDON_TABLE_PRIORITY_STATUSES"]
ADDON_NESTED_STATUSES = _RULES["ADDON_NESTED_STATUSES"]
ADDON_EQUIVALENT_STATUS_GROUPS = _RULES["ADDON_EQUIVALENT_STATUS_GROUPS"]
ADDON_TRAINING_JOURNAL_SHEETS = _RULES["ADDON_TRAINING_JOURNAL_SHEETS"]

# Колонки, різниця в яких між ОБЛІК і файлом information_unit сама по собі
# вносить людину в error_mis_statuses.xlsx, навіть коли статуси на всі дати
# збігаються (реальний випадок: у файлі БЧС підрозділу в людини порожня посада
# й інше звання). "ПІДРОЗДІЛ" - ні: ту саму назву підрозділи пишуть по-різному
# ("втз"/"вт заб" - так на реальних даних з усіма 14 відмінностями).
PAYMENT_MISMATCH_LABEL_COLUMNS = ("ПОСАДА", "ЗВАННЯ")

def _name_list(key):
    """Одиничний рядок замість списку (пропущена кома в JSON) стає кортежем з
    одного елемента, а не розбивається на символи."""
    raw = _DATA.get(key, [])
    return tuple([raw] if isinstance(raw, str) else raw)


def _date_name_pairs(key):
    """[{"date": "дд.мм.рррр", "name": "ПІБ"}, ...] з налаштувань (_DATA) ->
    кортежі (date, name). ПІБ лишається як є - нормалізація в generators/, бо
    constants.py не імпортує content/. Запис, який не вдається прочитати (не
    об'єкт, немає ПІБ, дата не у форматі дд.мм.рррр, напр. "06.092026"), - це
    помилка одразу: інакше виняток для одного запису тихо не спрацював би, і
    запис лишився б у файлі для ручної перевірки без жодного пояснення."""
    raw = _DATA.get(key, [])
    if not isinstance(raw, list):
        raise ValueError(f"{project_settings.source_of(key)}: має бути списком об'єктів {{\"date\", \"name\"}}.")
    pairs = []
    for entry in raw:
        if not isinstance(entry, dict) or not isinstance(entry.get("name"), str) or not entry["name"].strip():
            raise ValueError(f"{project_settings.source_of(key)}: запис {entry!r} - очікується {{\"date\": \"дд.мм.рррр\", \"name\": \"ПІБ\"}}.")
        date_text = entry.get("date")
        if not isinstance(date_text, str):
            raise ValueError(f"{project_settings.source_of(key)}: дата {date_text!r} не у форматі дд.мм.рррр.")
        try:
            day = datetime.strptime(date_text, "%d.%m.%Y").date()
        except ValueError:
            raise ValueError(f"{project_settings.source_of(key)}: дата {date_text!r} не у форматі дд.мм.рррр.") from None
        pairs.append((day, entry["name"]))
    return tuple(pairs)


def _is_days_count(value):
    return type(value) is int and 0 <= value <= 31


def _category_100_max_days():
    """Спільний ліміт з НАЛАШТУВАННЯ.xlsx (аркуш "Загальні"): ціле число днів від 0 до 31.
    Інакше помилка одразу при запуску, а не посеред генерації."""
    value = _DATA.get("CONST_PAYMENT_CATEGORY_100_MAX_DAYS")
    if not _is_days_count(value):
        raise ValueError(
            f"{project_settings.source_of('CONST_PAYMENT_CATEGORY_100_MAX_DAYS')}: значення {value!r} - потрібне ціле число днів від 0 до 31."
        )
    return value


# Колір клітинки з ПІБ для "color": true у записі EXCLUDE_NAMES і в Табелі для
# людей з порушенням правила категорії 100 (рішення користувача).
PAYMENT_CATEGORY_100_HIGHLIGHT_COLOR = "FFFFFF00"
_CATEGORY_100_PERSON_FORMAT = '{"name": "ПІБ", "days": число днів або null, "color": false, true чи "FFFF00"}'
_HEX_COLOR_RE = re.compile(r"#?((?:[0-9A-Fa-f]{2})?[0-9A-Fa-f]{6})")


def _category_100_color(key, entry):
    """false, null чи порожньо - без кольору (None); true - жовтий; колір
    числами "FFFF00", "#FFFF00" чи "FFFFFF00" - цей колір. Повертає ARGB."""
    color = entry["color"]
    if color is None or color is False or color == "":
        return None
    if color is True:
        return PAYMENT_CATEGORY_100_HIGHLIGHT_COLOR
    match = _HEX_COLOR_RE.fullmatch(color) if isinstance(color, str) else None
    if match is None:
        raise ValueError(
            f"{project_settings.source_of(key)}: для \"{entry['name']}\" колір {color!r} - "
            "очікується false, true або колір числами, напр. \"FFFF00\"."
        )
    return match.group(1).upper().rjust(8, "F")


def _category_100_people(key):
    """[{"name": "ПІБ", "days": ..., "color": ...}, ...] з налаштувань (_DATA) ->
    кортеж словників з тими самими ключами: days - число або None (null чи
    порожньо, тоді діє CONST_PAYMENT_CATEGORY_100_MAX_DAYS), color - ARGB або
    None. ПІБ лишається як є - нормалізація в content/, бо constants.py не
    імпортує content/. Усі три поля обов'язкові. Запис, який не вдається
    прочитати (зокрема старий запис рядком), - помилка одразу, як і в
    _date_name_pairs."""
    raw = _DATA.get(key, [])
    if not isinstance(raw, list):
        raise ValueError(f"{project_settings.source_of(key)}: має бути списком об'єктів {_CATEGORY_100_PERSON_FORMAT}.")
    people = []
    for entry in raw:
        if (not isinstance(entry, dict) or not {"name", "days", "color"} <= entry.keys()
                or not isinstance(entry["name"], str) or not entry["name"].strip()):
            raise ValueError(f"{project_settings.source_of(key)}: запис {entry!r} - очікується {_CATEGORY_100_PERSON_FORMAT}.")
        days = None if entry["days"] in (None, "") else entry["days"]
        if days is not None and not _is_days_count(days):
            raise ValueError(
                f"{project_settings.source_of(key)}: для \"{entry['name']}\" кількість днів {days!r} - "
                "очікується ціле число від 0 до 31 або null."
            )
        people.append({"name": entry["name"], "days": days, "color": _category_100_color(key, entry)})
    return tuple(people)


# Скільки днів поспіль у місяці можна мати категорію виплати 100 - число
# задаєте самі в НАЛАШТУВАННЯ.xlsx (рішення користувача), один ліміт на всі
# місяці. Серія рахується з 1-го числа місяця: колонка останнього дня
# попереднього місяця (30.09 у жовтневому Табелі) до неї не входить. Порушення -
# серія, довша за ліміт (рядок у Перевірка_категорії_100.xlsx); будь-який інший
# статус чи пропущена дата серію перериває.
CONST_PAYMENT_CATEGORY_100_MAX_DAYS = _category_100_max_days()

# Люди з таблиці "Ліміт категорії 100" НАЛАШТУВАННЯ.xlsx (там справжні ПІБ, тож у коді їх
# немає). "days" важливіше за спільний ліміт: 31 - можна увесь місяць, null -
# діє CONST_PAYMENT_CATEGORY_100_MAX_DAYS. "color" заливає клітинку з ПІБ в
# ОБЛІК.xlsx і Перевірка_категорії_100.xlsx, а людина з кольором є в цій таблиці
# навіть без порушення (рішення користувача).
PAYMENT_CATEGORY_100_PEOPLE = _category_100_people("EXCLUDE_NAMES")

# Люди, яких підрозділи можуть не подавати у файлах information_unit (напр.
# довготривале лікування): для них немає запису "Немає в жодному файлі
# information_unit" в error_mis_statuses.xlsx. Список ПІБ -
# таблиця "Не подаються в БЧС" НАЛАШТУВАННЯ.xlsx (справжні ПІБ, у
# коді їх немає). Інші розбіжності цієї людини (файл подає інший статус)
# показуються, як і раніше.
PAYMENT_ABSENT_FROM_FILES_EXCLUDE_NAMES = _name_list("EXCLUDE_ABSENT_FROM_INFORMATION_UNIT")

# Пари (дата, ПІБ) з таблиці "Без ручної перевірки" НАЛАШТУВАННЯ.xlsx,
# для яких не треба писати запис на ручну перевірку. Та сама людина на іншу дату
# лишається у файлі. Перший список - лише для "Потребує_ручної_перевірки_виплати.txt",
# другий - лише для "Потребує_ручної_перевірки.txt". Дата - та, що в дужках
# [дд.мм.рррр] у рядку файлу. Статуси в табелі від цього не змінюються, а записи
# без дати (напр. серія категорії 100 - перевірка за діапазоном днів) цими
# парами не прибираються.
PAYMENTS_REVIEW_EXCLUDE_PAIRS = _date_name_pairs("EXCLUDE_FROM_PAYMENTS_REVIEW")
REVIEW_EXCLUDE_PAIRS = _date_name_pairs("EXCLUDE_FROM_REVIEW")

# Ключі IGNORE_LISTS (колонка "Книга" таблиці "Ігнорування addons") - сталі назви типів
# книг addons (ADDON_BOOK_KINDS), а не імена файлів: тип книги впізнається за
# вмістом, тож список не залежить від назви файлу (рік, позначення підрозділу).
ADDON_IGNORE_LIST_KEYS = {
    "BOOK_61": "book61",
    "CVS": "cvs",
    "COMBAT_TRAINING_JOURNAL": "journal",
    "TREATMENT_JOURNAL": ADDON_TREATMENT_JOURNAL_KIND,
}
_IGNORE_LIST_FORMAT = '{"period": "дд.мм.рррр" або "дд.мм.рррр-дд.мм.рррр", "name": "ПІБ"}'
_IGNORE_PERIOD_RE = re.compile(r"(\d{2}\.\d{2}\.\d{4})(?:\s*[-–—]\s*(\d{2}\.\d{2}\.\d{4}))?")


def _ignore_period(key, period_text):
    """"дд.мм.рррр" -> (дата, дата); "дд.мм.рррр-дд.мм.рррр" -> (початок,
    кінець), обидва включно."""
    error = ValueError(
        f"{project_settings.source_of('IGNORE_LISTS')}, книга \"{key}\": період {period_text!r} - "
        "очікується \"дд.мм.рррр\" або \"дд.мм.рррр-дд.мм.рррр\"."
    )
    match = _IGNORE_PERIOD_RE.fullmatch(period_text.strip()) if isinstance(period_text, str) else None
    if match is None:
        raise error
    start_text, end_text = match.group(1), match.group(2) or match.group(1)
    try:
        start, end = (datetime.strptime(text, "%d.%m.%Y").date() for text in (start_text, end_text))
    except ValueError:
        raise error from None
    if start > end:
        raise ValueError(f"{project_settings.source_of('IGNORE_LISTS')}, книга \"{key}\": період {period_text!r} - початок пізніше за кінець.")
    return start, end


def _addon_ignore_list():
    """{"BOOK_61"/"CVS"/... (ADDON_IGNORE_LIST_KEYS): [{"period": ..., "name":
    "ПІБ"}, ...]} з налаштувань (_DATA["IGNORE_LISTS"]) -> кортеж (тип книги
    з ADDON_BOOK_KINDS, ПІБ, початок періоду, кінець періоду). ПІБ лишається
    як є - нормалізація в content/. Невідомий ключ чи запис, який не вдається
    прочитати, - помилка одразу, як і в _date_name_pairs: інакше людина тихо
    лишилась би в mis_match.xlsx."""
    raw = _DATA.get("IGNORE_LISTS", {})
    keys = ", ".join(ADDON_IGNORE_LIST_KEYS)
    if not isinstance(raw, dict):
        raise ValueError(f"{project_settings.source_of('IGNORE_LISTS')}: має бути об'єктом зі списками {keys}.")
    entries = []
    for key, items in raw.items():
        if key not in ADDON_IGNORE_LIST_KEYS:
            raise ValueError(f"{project_settings.source_of('IGNORE_LISTS')}: невідомий список \"{key}\" - можна: {keys}.")
        if not isinstance(items, list):
            raise ValueError(f"{project_settings.source_of('IGNORE_LISTS')}, книга \"{key}\": має бути списком об'єктів {_IGNORE_LIST_FORMAT}.")
        for entry in items:
            if not isinstance(entry, dict) or not isinstance(entry.get("name"), str) or not entry["name"].strip():
                raise ValueError(
                    f"{project_settings.source_of('IGNORE_LISTS')}, книга \"{key}\": запис {entry!r} - очікується {_IGNORE_LIST_FORMAT}."
                )
            entries.append((ADDON_IGNORE_LIST_KEYS[key], entry["name"], *_ignore_period(key, entry.get("period"))))
    return tuple(entries)


# Що не показувати в mis_match.xlsx (таблиця "Ігнорування addons" НАЛАШТУВАННЯ.xlsx,
# справжні ПІБ, у коді їх немає): людину на аркуші книги цього типу за дату
# чи період включно. Інші дати й інші книги addons звіряються, як і раніше.
ADDON_MISMATCH_IGNORE_LIST = _addon_ignore_list()

# Заголовок групи "Поза межами складу сил та засобів ...:". Людей там згадано
# реченнями, без таблиці й без сталої структури, тож надійно визначити статус і
# дату автоматично не вийде (content/daily_report_reader.py). За замовчуванням
# такі згадки лише збираються для ручної перевірки (REVIEW_FILE_NAME).
FREEFORM_GROUP_KEYWORD = "поза межами"

# Заголовок-список (не таблиця) "ПРИБУЛИ до пункту постійно(ї) дислокації
# військової частини ...:". Під ним кілька людей, кожна окремим абзацом
# "звання ПІБ" (стиль "List Paragraph", як і в заголовка), без таблиці й без
# дати. На відміну від вільного тексту "Поза межами...", це простий список
# імен з однозначним сенсом ("прибули" - повернулись у район виконання
# завдань), тож він застосовується автоматично (DEFAULT_STATUS, дата
# рапорту). У тексті мають бути обидва ключові слова, а не точна фраза: буває
# "постійно дислокації" без закінчення "-ої".
ARRIVAL_LIST_HEADING_KEYWORDS = ("прибули", "дислокації")

# Назви й скорочення власного підрозділу - з НАЛАШТУВАННЯ.xlsx (як і в
# generator_br_and_report_for_money), а не в коді: це оперативна інформація,
# яка може змінитися (перейменування, нова абревіатура). Разом із
# SIGNATORY_ROLE_KEYWORDS (сталі слова ролі) вони дають змогу впізнати рядок
# підпису наприкінці рапорту (_is_signatory_paragraph), щоб не вважати його
# нерозпізнаним військовослужбовцем у "Поза межами...".
SIGNATORY_UNIT_NAMES = tuple(_DATA["unit"].values())
SIGNATORY_ROLE_KEYWORDS = ("командир", "начальник штабу")

# Дата в тексті: назва файлу рапорту ("01.08.2026 - щоденний рапорт.docx") чи
# клітинка "Дата прибуття/вибуття", можливо з припискою на кшталт "після 17:00".
DATE_RE = re.compile(r"\d{2}\.\d{2}\.\d{4}")

# Формат клітинок-дат у звітах (заголовки error_mis_statuses.xlsx,
# mis_match.xlsx, звіту розбіжностей РОП/ВОП). Excel регістру не розрізняє, а
# інші програми перегляду розрізняють: "DD" там - день року, "mm" - хвилини.
# Реальний випадок: файл на іншому пристрої показував 30.09.2026 як
# "273.09.2026". "dd.MM.yyyy" однаково читають і Excel, і такі програми.
DATE_NUMBER_FORMAT = "dd.MM.yyyy"