"""Правила статусів табеля на аркушах книги resources/НАЛАШТУВАННЯ.xlsx.

Колись усі ці правила жили в constants.py, і змінити статус у них можна було
лише в коді. Тепер:
- "Групи статусів" - особливі статуси (за замовчуванням, ВЛК, переведення ...)
  і переліки статусів з однаковою роллю (відпустки, звіряти з БЧС ...);
- "Заголовки рапорту" - слова заголовка таблиці рапорту -> статус;
- "Речення рапорту" - шаблон речення з розділу "Поза межами" -> статус;
- "Конфлікт у рапортах" - хто перемагає, коли рапорти на одну дату дали два статуси;
- "Звірка з addons" - де в книгах resources/addons має бути запис статусу, що
  в Табелі важливіше за запис книги, що може бути вкладене;
- "Рівнозначні статуси" - групи статусів, однакових для звірки з addons;
- "Журнал обліку БП" - аркуш журналу (і значення колонки "НЦ") -> статус.

read_rules читає аркуші в словник налаштувань (аркуша немає - None, і книгу
доповнює project_settings.load_settings), write_rule_sheets пише їх,
build_rules будує з них значення, які constants.py віддає решті коду під
колишніми назвами (DEFAULT_STATUS, LEAVE_STATUSES, STATUS_HEADING_RULES ...).

Статуси на цих аркушах можна писати і назвою (колонка "Статус" аркуша
"Статуси"), і кодом: build_rules переводить усе в коди (code_of)."""
import re

from utils import settings_book as sb

GROUPS_SHEET = "Групи статусів"
HEADINGS_SHEET = "Заголовки рапорту"
SENTENCES_SHEET = "Речення рапорту"
REPORT_CONFLICTS_SHEET = "Конфлікт у рапортах"
ADDONS_SHEET = "Звірка з addons"
EQUIVALENT_SHEET = "Рівнозначні статуси"
TRAINING_SHEET = "Журнал обліку БП"

# Аркуші книги, на яких лежать ці таблиці (кілька таблиць на аркуші: таблиця
# знаходиться за смугою "■ назва" - settings_book.locate_table).
STATUSES_BOOK_SHEET = "Статуси"
REPORTS_BOOK_SHEET = "Рапорти"
ADDONS_BOOK_SHEET = "Звірка з addons"
PLACE = {
    GROUPS_SHEET: STATUSES_BOOK_SHEET,
    HEADINGS_SHEET: REPORTS_BOOK_SHEET,
    SENTENCES_SHEET: REPORTS_BOOK_SHEET,
    REPORT_CONFLICTS_SHEET: REPORTS_BOOK_SHEET,
    ADDONS_SHEET: ADDONS_BOOK_SHEET,
    EQUIVALENT_SHEET: ADDONS_BOOK_SHEET,
    TRAINING_SHEET: ADDONS_BOOK_SHEET,
}


def _label(table):
    return sb.table_label(PLACE[table], table)


# Ключ словника налаштувань -> таблиця (назва в смузі "■ ...").
RULE_SHEETS = {
    "STATUS_GROUPS": GROUPS_SHEET,
    "HEADING_RULES": HEADINGS_SHEET,
    "FREEFORM_PATTERNS": SENTENCES_SHEET,
    "REPORT_CONFLICTS": REPORT_CONFLICTS_SHEET,
    "ADDON_STATUSES": ADDONS_SHEET,
    "EQUIVALENT_STATUSES": EQUIVALENT_SHEET,
    "TRAINING_JOURNAL": TRAINING_SHEET,
}

# Позначки таблиці "Заголовки рапорту" (ті самі, що й у utils/default_statuses.py).
TRANSFER_HEADING = "(переведення)"
IGNORED_HEADING = "(не відстежувати)"
SPECIAL_GROUP_PLACEHOLDER = "{спецгрупа}"

# Вид книги resources/addons (constants.ADDON_BOOK_KINDS) -> назва на аркуші.
ADDON_BOOK_LABELS = {"book61": "book61", "cvs": "ЦВС", "journal": "журнал обліку БП", "treatment": "журнал лікування"}


def status_value(text):
    """Статус чи категорія з клітинки: "100" -> 100 (як число в клітинці
    ОБЛІК.xlsx), "100_БЗ" -> "100_БЗ", порожньо -> None."""
    text = sb.cell_text(text)
    if not text:
        return None
    return int(text) if text.isdigit() else text


def status_norm(value):
    """Статус для порівняння: без урахування регістру й зайвих пробілів."""
    return " ".join(value.split()).upper() if isinstance(value, str) else value


# Однакова примітка аркушів, де статуси згадуються: назва статусу живе в
# одному місці - колонка "Статус" аркуша "Статуси".
CODES_HINT = (
    " Статуси тут - коди (сіра колонка «Код» аркуша «Статуси»): назву статусу змінюють лише на аркуші «Статуси», "
    "тут після цього нічого міняти не треба."
)

# ---------------------------------------------------------------- "Групи статусів"

_ONE = " Один статус."
_ANY = " У Табелі підходить будь-який із цих статусів."
# (ключ, назва, пояснення, один статус?, це статуси (а не назви аркушів)?)
_GROUPS = (
    ("DEFAULT_STATUS", "Статус за замовчуванням",
     "У районі виконання завдань: людина повернулась з будь-якої відсутності; з ним програма заповнює дні без подій." + _ONE, True),
    ("AFTER_HOURS_ARRIVAL_STATUS", "Прибув після 17:00",
     "Рапорт складається станом на 17 годину: хто прибув «після 17.00», того дня має цей статус, а з наступного - "
     "статус за замовчуванням." + _ONE, True),
    ("VLK_STATUS", "Військово-лікарська комісія", "ВЛК усередині відпустки - відпустка (окремо не оплачується)." + _ONE, True),
    ("ROZP_STATUS", "Розпорядження",
     "Адміністративна підстава: проти статусів «Важливіші за розпорядження» - не розбіжність." + _ONE, True),
    ("ADDON_LONG_TREATMENT_STATUS", "Довготривале лікування",
     "Під час звірки з addons рівнозначне лікуванню, ВЛК і відпусткам за станом здоров'я." + _ONE, True),
    ("TRANSFER_TO_BRIGADE_STATUS", "Переведення в підрозділи бригади",
     "Переведення до власної в/ч чи в/ч бригади (аркуш «Загальні»); виплату здійснює підрозділ, куди переведено." + _ONE, True),
    ("TRANSFER_TO_OTHER_UNIT_STATUS", "Переведення в іншу в/ч", "Переведення до будь-якої іншої в/ч." + _ONE, True),
    ("LEGACY_TRANSFER_STATUS", "Колишнє переведення без уточнення",
     "Ще трапляється в уже заповненому ОБЛІК (не видно, куди переведено)." + _ONE, True),
    ("LEAVE_STATUSES", "Відпустки", "Усі види відпустки: ВЛК усередині них - відпустка; відпустка важливіша за ВЛК.", False),
    ("STATUSES_OVER_ROZP", "Важливіші за розпорядження",
     "Лікування й відпустки за станом здоров'я: «розпорядження» проти них - не розбіжність.", False),
    ("PAYMENT_DUTY_CATEGORIES", "Категорії виплати людини на службі",
     "Числа з файлів БЧС для людини в районі виконання завдань - вид оплати, а не статус (звіти показують статус).", False),
    ("PAYMENT_MISMATCH_COMPARABLE_STATUSES", "Звіряти з файлами БЧС",
     "Статуси ОБЛІК для виплат, розбіжність яких з файлом підрозділу йде в error_mis_statuses.xlsx (числа - завжди).", False),
    ("PAYMENT_ABSENT_FROM_FILES_EXEMPT_STATUSES", "Можуть бути відсутні у файлах БЧС",
     "Людина з таким останнім статусом уже не на обліку підрозділу: її відсутність у файлах БЧС - не розбіжність.", False),
    ("ADDON_NOT_RETURNED_AFTER_SZCH_STATUSES", "Не повернувся після СЗЧ",
     "Після «Повер» у book61 у Табелі підходить будь-який статус, крім цих (вони кажуть, що людина не служить).", False),
    ("ADDON_ARRIVAL_DAY_STATUSES", "Статуси дня прибуття",
     "У день прибуття з аркушів «Аркуші книг: день прибуття» ці статуси Табеля - не розбіжність.", False),
    ("ADDON_ARRIVAL_DAY_SHEETS", "Аркуші книг: день прибуття",
     "Аркуші книг addons (назви аркушів, не статуси), останній день запису яких - день прибуття.", False),
    ("ADDON_NON_COMBAT_TREATMENT", "ЦВС: небойове лікування",
     "Книга ЦВС, аркуш лікування (соматика, травмування, небойове)." + _ANY, False),
    ("ADDON_KILLED_WITHOUT_CERTIFICATE", "book61: загибель без свідоцтва",
     "Аркуш «Алф2», загибель без свідоцтва про смерть (до дати свідоцтва)." + _ANY, False),
    ("ADDON_TREATMENT_JOURNAL", "Журнал лікування", "Журнал лікування поза частиною статусу не має." + _ANY, False),
    ("STATUS_CONFLICT_PAID_BY_DEFAULT", "Конфлікт у рапортах: оплачувані",
     "Два статуси на одну дату без рядка на аркуші «Конфлікт у рапортах»: оплачуваний звідси перемагає "
     "неоплачуваний; обидва чи жоден - ручна перевірка.", False),
    ("STATUS_CONFLICT_PRIORITY_EXEMPT", "Конфлікт у рапортах: без автоматики",
     "Адміністративні, дисциплінарні, остаточні статуси: суперечка з ними завжди йде на ручну перевірку.", False),
)
GROUP_SETTINGS = tuple(sb.Setting(key, sb.LIST, required=False, note=note, label=label) for key, label, note, _one in _GROUPS)
_SINGLE_KEYS = {key for key, _label, _note, one in _GROUPS if one}
_NOT_STATUS_KEYS = {"ADDON_ARRIVAL_DAY_SHEETS"}

# ---------------------------------------------------------------- таблиці
# Таблиці одного аркуша мають спільні колонки аркуша, тож широка колонка
# займає кілька колонок аркуша (span): так вона не розтягує вузьку колонку
# таблиці над чи під нею.

HEADING_COLUMNS = (
    sb.Column("keyword", "Слова в заголовку", nonempty=True, width=44, span=2,
              note="Шматок заголовка таблиці рапорту (без урахування регістру), напр. «у відпустку». "
                   f"{SPECIAL_GROUP_PLACEHOLDER} - назва спеціальної групи з аркуша «Загальні»."),
    sb.Column("status", "Статус", sb.NUMBER_TEXT, nonempty=True, width=20,
              note=f"Статус людей з таблиці під таким заголовком. {TRANSFER_HEADING} - статус дає колонка «Куди вибув» "
                   f"(переведення в бригаду чи іншу в/ч); {IGNORED_HEADING} - рядки під заголовком не відстежуються."),
)

SENTENCE_COLUMNS = (
    sb.Column("pattern", "Шаблон речення", nonempty=True, width=86, wrap=True, span=4,
              note="Регулярний вираз (без урахування регістру), напр. вибув\\s+у\\s+СЗЧ. Змінюйте обережно: помилку "
                   "в шаблоні програма покаже при запуску."),
    sb.Column("status", "Статус", sb.NUMBER_TEXT, nonempty=True, width=14),
    sb.Column("date_prefix", "Дата на початку", sb.BOOL, default=True, width=13,
              note="Так - перед реченням може стояти дата «дд.мм.рррр (року)»: з неї й діє статус (без неї - дата "
                   "рапорту). Ні - шаблон сам ловить дату групою (?P<date>...)."),
)

CONFLICT_COLUMNS = (
    sb.Column("first", "Статус", sb.NUMBER_TEXT, nonempty=True, width=22),
    sb.Column("second", "Інший статус", sb.NUMBER_TEXT, nonempty=True, width=22),
    sb.Column("winner", "Перемагає", sb.NUMBER_TEXT, nonempty=True, width=18, note="Один з двох статусів рядка."),
)

_ADDON_GROUP = "Де в книгах має бути запис"
ADDON_COLUMNS = (
    sb.Column("status", "Статус", sb.NUMBER_TEXT, nonempty=True, width=22,
              note="Статус Табеля (чи статус запису книги - для двох колонок праворуч)."),
    sb.Column("book", "Книга", width=18, group=_ADDON_GROUP, choices=tuple(ADDON_BOOK_LABELS.values()),
              note="Книга resources/addons, де має бути запис на дні з цим статусом у Табелі. Порожньо - не перевіряється."),
    sb.Column("sheet", "Аркуш книги", width=20, group=_ADDON_GROUP),
    sb.Column("required", "Завжди", sb.BOOL, default=False, width=10, group=_ADDON_GROUP,
              note="Так - запис має бути завжди (book61, напр. відпустки на аркуші «ВП»); Ні - лише в дні, коли жодна "
                   "книга нічого не каже."),
    sb.Column("table_priority", "Табель важливіший, якщо там", sb.LIST, width=26, wrap=True,
              note="Книга очікує цей статус, а в Табелі - один із цих: Табель має пріоритет, це не розбіжність."),
    sb.Column("nested", "Може бути всередині", sb.LIST, width=25, wrap=True,
              note="Звірка книг між собою: запис із цими статусами, повністю вкладений у період запису з цим статусом "
                   "(чи з цього аркуша), - не розбіжність (напр. ШП посеред БЗВП)."),
)

EQUIVALENT_NOTE_HEADER = "Де трапляється (довідка)"
EQUIVALENT_COLUMNS = (
    sb.Column("statuses", "Статуси", sb.LIST, nonempty=True, width=30, span=2, wrap=True,
              note="Два чи більше статусів через кому, які під час звірки з addons - одне й те саме."),
    sb.Column("note", EQUIVALENT_NOTE_HEADER, width=81, wrap=True, required=False, span=4,
              note="Лише пояснення для людини: яка книга й аркуш пише один статус групи, а Табель - інший. Програма "
                   "цю колонку не читає."),
)

TRAINING_COLUMNS = (
    sb.Column("sheet", "Аркуш журналу", nonempty=True, width=22),
    sb.Column("value", "Значення колонки «НЦ»", width=24, span=2,
              note="Порожньо - статус для всього аркуша; значення - статус, коли в колонці «НЦ» саме воно "
                   "(«власна в/ч» - навчання у власній в/ч)."),
    sb.Column("status", "Статус", sb.NUMBER_TEXT, nonempty=True, width=14),
)

SHEET_TITLES = {
    GROUPS_SHEET: (
        "Групи статусів - особливі статуси й переліки",
        "Які статуси програма вважає статусом за замовчуванням, ВЛК, переведенням, відпусткою тощо. Кілька статусів - "
        "через кому. Колонку «Параметр» не змінюйте." + CODES_HINT,
    ),
    HEADINGS_SHEET: (
        "Заголовки рапорту - який статус дає таблиця",
        "Заголовки таблиць щоденного рапорту перевіряються по черзі згори вниз, перший збіг перемагає: «з відпустки» "
        "має стояти вище за «відпустк». (переведення) - статус з колонки «Куди вибув»; (не відстежувати) - таблиця "
        "ігнорується." + CODES_HINT,
    ),
    SENTENCES_SHEET: (
        "Речення рапорту - зміна статусу у вільному тексті",
        "Речення з розділу «Поза межами...», які впізнаються як зміна статусу, хоч би який базовий статус людини. "
        "Шаблони - регулярні вирази: змінюйте обережно." + CODES_HINT,
    ),
    REPORT_CONFLICTS_SHEET: (
        "Конфлікт у рапортах - який статус точніший",
        "Два статуси для однієї людини на ту саму дату: якщо пара є тут, одразу береться «Перемагає» (той самий факт з "
        "уточненою причиною), інакше - правило груп «Конфлікт у рапортах» чи ручна перевірка." + CODES_HINT,
    ),
    ADDONS_SHEET: (
        "Звірка з addons - статуси Табеля й записи книг",
        "Для статусу Табеля - книга й аркуш, де має бути запис (mis_match.xlsx), і коли запис книги не суперечить "
        "Табелю. Статуси, яких не веде жодна книга, - без книги." + CODES_HINT,
    ),
    EQUIVALENT_SHEET: (
        "Рівнозначні статуси - однакові для звірки з addons",
        "Кожен рядок - група статусів, однакових під час звірки з resources/addons: книга пише один статус групи, а "
        "Табель - інший, і це не розбіжність. Діє для всіх книг addons; де саме таке трапляється - колонка праворуч."
        + CODES_HINT,
    ),
    TRAINING_SHEET: (
        "Журнал обліку БП - аркуш журналу -> статус",
        "Статус людей з аркуша журналу обліку бойової підготовки; рядок із «Значення колонки «НЦ»» - інший статус для "
        "цього значення." + CODES_HINT,
    ),
}

# ---------------------------------------------------------------- читання


def _read_groups(workbook):
    raw_values = sb.read_key_values(workbook, PLACE[GROUPS_SHEET], section=GROUPS_SHEET)
    groups = {}
    for setting in GROUP_SETTINGS:
        if setting.key not in raw_values:
            raise sb.SettingsError(
                f'{sb.location(_label(GROUPS_SHEET))}: немає рядка з параметром "{setting.label}" ({setting.key}).'
            )
        row_number, raw = raw_values[setting.key]
        items = sb.split_list(sb.cell_text(raw))
        values = items if setting.key in _NOT_STATUS_KEYS else [status_value(item) for item in items]
        if setting.key in _SINGLE_KEYS and len(values) != 1:
            raise sb.SettingsError(
                f'{sb.location(_label(GROUPS_SHEET), row_number, sb.VALUE_HEADER)}: "{setting.label}" - має бути рівно один статус.'
            )
        groups[setting.key] = values
    return groups


def _pair(row_number, row, sheet_name):
    first, second, winner = (status_value(row[key]) for key in ("first", "second", "winner"))
    if status_norm(winner) not in (status_norm(first), status_norm(second)):
        raise sb.SettingsError(
            f'{sb.location(sheet_name, row_number, "Перемагає")}: має бути один з двох статусів рядка ({first} чи {second}).'
        )
    return {"first": first, "second": second, "winner": winner}


def freeform_pattern(anchor):
    """Шаблон речення з необов'язковою датою "дд.мм.рррр (року)" перед ним."""
    return re.compile(rf"(?:(?P<date>\d{{2}}\.\d{{2}}\.\d{{4}})\s*(?:року)?\s*)?{anchor}", re.IGNORECASE)


def compile_sentence(pattern, date_prefix):
    return freeform_pattern(pattern) if date_prefix else re.compile(pattern, re.IGNORECASE)


def _read_sentences(workbook):
    rows = []
    for row_number, row in sb.read_rows(workbook, PLACE[SENTENCES_SHEET], SENTENCE_COLUMNS, section=SENTENCES_SHEET):
        try:
            compile_sentence(row["pattern"], row["date_prefix"])
        except re.error as error:
            raise sb.SettingsError(
                f'{sb.location(_label(SENTENCES_SHEET), row_number, "Шаблон речення")}: помилка в шаблоні ({error}).'
            ) from error
        rows.append({"pattern": row["pattern"], "status": status_value(row["status"]), "date_prefix": row["date_prefix"]})
    return rows


def _read_headings(workbook):
    rows, seen = [], {}
    for row_number, row in sb.read_rows(workbook, PLACE[HEADINGS_SHEET], HEADING_COLUMNS, section=HEADINGS_SHEET):
        key = sb.normalize_name(row["keyword"])
        if key in seen:
            raise sb.SettingsError(
                f'{sb.location(_label(HEADINGS_SHEET), row_number, "Слова в заголовку")}: такі слова вже є в рядку {seen[key]}.'
            )
        seen[key] = row_number
        rows.append({"keyword": row["keyword"], "status": status_value(row["status"])})
    return rows


def _book_kind(label, row_number):
    wanted = sb.normalize_name(label)
    for kind, book_label in ADDON_BOOK_LABELS.items():
        if wanted in (sb.normalize_name(kind), sb.normalize_name(book_label)):
            return kind
    raise sb.SettingsError(
        f'{sb.location(_label(ADDONS_SHEET), row_number, "Книга")}: невідома книга {label!r} - можна: {", ".join(ADDON_BOOK_LABELS.values())}.'
    )


def _read_addons(workbook):
    rows, seen = [], {}
    for row_number, row in sb.read_rows(workbook, PLACE[ADDONS_SHEET], ADDON_COLUMNS, section=ADDONS_SHEET):
        status = status_value(row["status"])
        if status_norm(status) in seen:
            raise sb.SettingsError(f'{sb.location(_label(ADDONS_SHEET), row_number, "Статус")}: статус уже є в рядку {seen[status_norm(status)]}.')
        seen[status_norm(status)] = row_number
        book = _book_kind(row["book"], row_number) if row["book"] else None
        if (book is None) != (not row["sheet"]):
            raise sb.SettingsError(
                f'{sb.location(_label(ADDONS_SHEET), row_number, "Аркуш книги" if book else "Книга")}: книга й аркуш - лише разом.'
            )
        if row["required"] and book != "book61":
            raise sb.SettingsError(f'{sb.location(_label(ADDONS_SHEET), row_number, "Завжди")}: "Так" - лише для книги book61.')
        rows.append({
            "status": status, "book": book, "sheet": row["sheet"] or None, "required": bool(row["required"]),
            "table_priority": [status_value(item) for item in row["table_priority"]],
            "nested": [status_value(item) for item in row["nested"]],
        })
    return rows


def _read_equivalent(workbook):
    rows = []
    for row_number, row in sb.read_rows(workbook, PLACE[EQUIVALENT_SHEET], EQUIVALENT_COLUMNS, section=EQUIVALENT_SHEET):
        if len(row["statuses"]) < 2:
            raise sb.SettingsError(f'{sb.location(_label(EQUIVALENT_SHEET), row_number, "Статуси")}: потрібні щонайменше два статуси.')
        rows.append({"statuses": [status_value(item) for item in row["statuses"]], "note": row["note"]})
    return rows


def _read_training(workbook):
    rows, seen = [], {}
    for row_number, row in sb.read_rows(workbook, PLACE[TRAINING_SHEET], TRAINING_COLUMNS, section=TRAINING_SHEET):
        key = (sb.normalize_name(row["sheet"]), sb.normalize_name(row["value"]))
        if key in seen:
            raise sb.SettingsError(f'{sb.location(_label(TRAINING_SHEET), row_number, "Аркуш журналу")}: такий рядок уже є в рядку {seen[key]}.')
        seen[key] = row_number
        rows.append({"sheet": row["sheet"], "value": row["value"], "status": status_value(row["status"])})
    sheets_with_default = {sb.normalize_name(row["sheet"]) for row in rows if not row["value"]}
    for row in rows:
        if sb.normalize_name(row["sheet"]) not in sheets_with_default:
            raise sb.SettingsError(
                f'{sb.location(_label(TRAINING_SHEET))}: для аркуша "{row["sheet"]}" немає рядка з порожнім «Значення колонки «НЦ»» '
                "(статус для всього аркуша)."
            )
    return rows


_READERS = {
    "STATUS_GROUPS": _read_groups,
    "HEADING_RULES": _read_headings,
    "FREEFORM_PATTERNS": _read_sentences,
    "REPORT_CONFLICTS": lambda workbook: [
        _pair(row_number, row, _label(REPORT_CONFLICTS_SHEET)) for row_number, row in sb.read_rows(workbook, PLACE[REPORT_CONFLICTS_SHEET], CONFLICT_COLUMNS, section=REPORT_CONFLICTS_SHEET)
    ],
    "ADDON_STATUSES": _read_addons,
    "EQUIVALENT_STATUSES": _read_equivalent,
    "TRAINING_JOURNAL": _read_training,
}


def read_rules(workbook):
    """{ключ RULE_SHEETS: дані таблиці чи None, якщо таблиці в книзі немає}."""
    return {
        key: (_READERS[key](workbook) if sb.has_table(workbook, PLACE[sheet_name], sheet_name) else None)
        for key, sheet_name in RULE_SHEETS.items()
    }


# ---------------------------------------------------------------- запис


def _shown(value):
    return str(value) if value is not None else None


def rule_section(key, value):
    """Таблиця правил key (RULE_SHEETS) зі значенням value - для
    settings_book.write_sections (аркуш - PLACE)."""
    name = RULE_SHEETS[key]
    hint = SHEET_TITLES[name][1]
    if key == "STATUS_GROUPS":
        values = {setting_key: ", ".join(str(item) for item in items) for setting_key, items in value.items()}
        # Колонки аркуша "Статуси" - спільні з таблицею статусів над цією:
        # "Назва" - під "Статус", "Значення" - під "Що означає", "Пояснення" -
        # на чотири вузькі колонки праворуч (Колір ... Код).
        return sb.Section(name, hint, settings=GROUP_SETTINGS, values=values, widths=(24, 62, 57), spans=(1, 1, 4, 1),
                          wrap_values=True)
    if key == "HEADING_RULES":
        return sb.Section(name, hint, HEADING_COLUMNS, [{"keyword": row["keyword"], "status": _shown(row["status"])} for row in value])
    if key == "FREEFORM_PATTERNS":
        return sb.Section(name, hint, SENTENCE_COLUMNS, value)
    if key == "REPORT_CONFLICTS":
        return sb.Section(name, hint, CONFLICT_COLUMNS, value)
    if key == "ADDON_STATUSES":
        return sb.Section(name, hint, ADDON_COLUMNS, [
            {**row, "book": ADDON_BOOK_LABELS.get(row["book"], row["book"]) if row["book"] else None} for row in value
        ])
    if key == "EQUIVALENT_STATUSES":
        return sb.Section(name, hint, EQUIVALENT_COLUMNS, value)
    return sb.Section(name, hint, TRAINING_COLUMNS, value, band_by="sheet")


# ---------------------------------------------------------------- побудова значень constants.py


def build_rules(data, code_of, special_group_name, transfer_heading_marker):
    """Аркуші правил -> {назва в constants.py: значення} того самого вигляду,
    що й колись у constants.py. code_of - назва чи код статусу -> код;
    special_group_name - назва спеціальної групи (нижній регістр) замість
    {спецгрупа}; transfer_heading_marker - constants.TRANSFER_HEADING_MARKER
    замість (переведення)."""
    groups = data["STATUS_GROUPS"]

    def codes(items):
        return tuple(code_of(item) for item in items)

    def group(key):
        return codes(groups[key])

    def composite(key):
        statuses = group(key)
        return "/".join(str(status) for status in statuses), statuses

    non_combat, non_combat_statuses = composite("ADDON_NON_COMBAT_TREATMENT")
    killed, killed_statuses = composite("ADDON_KILLED_WITHOUT_CERTIFICATE")
    treatment, treatment_statuses = composite("ADDON_TREATMENT_JOURNAL")

    heading_rules, ignored_headings = {}, []
    for row in data["HEADING_RULES"]:
        keyword = row["keyword"].replace(SPECIAL_GROUP_PLACEHOLDER, special_group_name).lower()
        status = row["status"]
        if status_norm(status) == status_norm(IGNORED_HEADING):
            ignored_headings.append(keyword)
        elif status_norm(status) == status_norm(TRANSFER_HEADING):
            heading_rules[keyword] = transfer_heading_marker
        else:
            heading_rules[keyword] = code_of(status)

    addons = data["ADDON_STATUSES"]
    training = {}
    for row in data["TRAINING_JOURNAL"]:
        if not row["value"]:
            training.setdefault(row["sheet"], [None, {}])[0] = code_of(row["status"])
    for row in data["TRAINING_JOURNAL"]:
        if row["value"]:
            training[row["sheet"]][1][row["value"]] = code_of(row["status"])

    to_brigade, to_other = group("TRANSFER_TO_BRIGADE_STATUS")[0], group("TRANSFER_TO_OTHER_UNIT_STATUS")[0]
    return {
        "DEFAULT_STATUS": group("DEFAULT_STATUS")[0],
        "AFTER_HOURS_ARRIVAL_STATUS": group("AFTER_HOURS_ARRIVAL_STATUS")[0],
        "VLK_STATUS": group("VLK_STATUS")[0],
        "ROZP_STATUS": group("ROZP_STATUS")[0],
        "ADDON_LONG_TREATMENT_STATUS": group("ADDON_LONG_TREATMENT_STATUS")[0],
        "TRANSFER_TO_BRIGADE_STATUS": to_brigade,
        "TRANSFER_TO_OTHER_UNIT_STATUS": to_other,
        "TRANSFER_STATUSES": (to_brigade, to_other),
        "LEGACY_TRANSFER_STATUS": group("LEGACY_TRANSFER_STATUS")[0],
        "LEAVE_STATUSES": group("LEAVE_STATUSES"),
        "STATUSES_OVER_ROZP": group("STATUSES_OVER_ROZP"),
        "PAYMENT_DUTY_CATEGORIES": group("PAYMENT_DUTY_CATEGORIES"),
        "PAYMENT_MISMATCH_COMPARABLE_STATUSES": group("PAYMENT_MISMATCH_COMPARABLE_STATUSES"),
        "PAYMENT_ABSENT_FROM_FILES_EXEMPT_STATUSES": group("PAYMENT_ABSENT_FROM_FILES_EXEMPT_STATUSES"),
        "ADDON_NOT_RETURNED_AFTER_SZCH_STATUSES": group("ADDON_NOT_RETURNED_AFTER_SZCH_STATUSES"),
        "ADDON_ARRIVAL_DAY_STATUSES": group("ADDON_ARRIVAL_DAY_STATUSES"),
        "ADDON_ARRIVAL_DAY_SHEETS": tuple(groups["ADDON_ARRIVAL_DAY_SHEETS"]),
        "ADDON_NON_COMBAT_TREATMENT_STATUS": non_combat,
        "ADDON_KILLED_WITHOUT_CERTIFICATE_STATUS": killed,
        "ADDON_TREATMENT_JOURNAL_STATUS": treatment,
        "ADDON_ONE_OF_STATUSES": {non_combat: non_combat_statuses, killed: killed_statuses, treatment: treatment_statuses},
        "STATUS_CONFLICT_PAID_BY_DEFAULT": group("STATUS_CONFLICT_PAID_BY_DEFAULT"),
        "STATUS_CONFLICT_PRIORITY_EXEMPT": group("STATUS_CONFLICT_PRIORITY_EXEMPT"),
        "STATUS_HEADING_RULES": heading_rules,
        "IGNORED_TABLE_HEADING_KEYWORDS": tuple(ignored_headings),
        "FREEFORM_STATUS_PATTERNS": tuple(
            (compile_sentence(row["pattern"], row["date_prefix"]), code_of(row["status"])) for row in data["FREEFORM_PATTERNS"]
        ),
        "STATUS_CONFLICT_OVERRIDES": {
            frozenset({code_of(row["first"]), code_of(row["second"])}): code_of(row["winner"]) for row in data["REPORT_CONFLICTS"]
        },
        "ADDON_REQUIRED_BOOK61_SHEETS": {code_of(row["status"]): row["sheet"] for row in addons if row["required"]},
        "ADDON_TRACKED_STATUSES": {
            code_of(row["status"]): (row["book"], row["sheet"]) for row in addons if row["book"] and not row["required"]
        },
        "ADDON_TABLE_PRIORITY_STATUSES": {code_of(row["status"]): codes(row["table_priority"]) for row in addons if row["table_priority"]},
        "ADDON_NESTED_STATUSES": {code_of(row["status"]): codes(row["nested"]) for row in addons if row["nested"]},
        "ADDON_EQUIVALENT_STATUS_GROUPS": tuple(frozenset(codes(row["statuses"])) for row in data["EQUIVALENT_STATUSES"]),
        "ADDON_TRAINING_JOURNAL_SHEETS": {sheet: (status, by_value) for sheet, (status, by_value) in training.items()},
    }

