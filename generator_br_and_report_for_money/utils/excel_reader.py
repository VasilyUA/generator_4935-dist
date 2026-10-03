import os
import pandas as pd
from datetime import datetime
from openpyxl.utils import column_index_from_string, get_column_letter

from utils.date_utils import date_to_str
# SUPPORTED_EXTENSIONS/OPTIONAL_PERSONEL_COLUMN_NAMES живуть у constants.py (єдине
# джерело), але constants.py, у свою чергу, викликає detect_personel_list_columns_letters
# (нижче) - імпорт "from constants import ..." тут НАГОРІ файлу створив би циклічний
# імпорт (constants.py ще не встиг би визначити ці значення, коли сам довантажується
# ЧЕРЕЗ цей модуль) - тож обидва значення імпортуються ЛОКАЛЬНО, у місці використання.


# Запасні назви аркуша, якщо очікуваного немає у файлі — перевіряються по
# черзі, у цьому порядку. Навмисно лише НЕЙТРАЛЬНІ дефолтні назви Excel
# (безпечні для БУДЬ-ЯКОГО пошуку аркуша, у будь-якому файлі/викликові) -
# ЗМІСТОВНІ назви (напр. "Табель", назва місяця) сюди НЕ входять, бо вони
# можуть існувати у файлі як СТОРОННІЙ аркуш іншого призначення (напр. файл,
# де шукають "ТВО", може містити аркуш "Табель" з геть іншими даними) - якщо
# додати їх у цей загальний список, пошук "ТВО" міг би помилково "знайти" й
# прочитати звичайний аркуш особового складу як ТВО-дані. Такі, ЗМІСТОВНІ
# запасні назви передаються ОКРЕМО, явно, лише для викликів, що дійсно шукають
# САМЕ аркуш особового складу - див. PERSONEL_LIST_FALLBACK_SHEET_NAMES нижче
# і параметр extra_fallback_sheet_names.
_FALLBACK_SHEET_NAMES = [
    'Аркуш1', 'Аркуш 1', 'Лист1', 'Sheet1', 'Лист 1', 'Sheet 1', 'Лист_1', 'Sheet_1',
    'Лист', 'Sheet', 'Лист 2', 'Sheet 2', 'Лист_2', 'Sheet_2',
]

# Назви місяців (той самий словник, що й content/money_report_helpers.py::
# MONTH_NAMES_NOMINATIVE_UPPER, продубльований тут через ту саму циклічний-
# імпорт причину, що описана вище для SUPPORTED_EXTENSIONS) - PERSONEL_LIST_SHEET_NAME
# завжди називається поточним місяцем (напр. "ТРАВЕНЬ"), тож усі 12 назв - законні
# запасні варіанти САМЕ для аркуша особового складу (не для будь-якого пошуку).
_UKRAINIAN_MONTH_NAMES_UPPER = [
    "СІЧЕНЬ", "ЛЮТИЙ", "БЕРЕЗЕНЬ", "КВІТЕНЬ", "ТРАВЕНЬ", "ЧЕРВЕНЬ",
    "ЛИПЕНЬ", "СЕРПЕНЬ", "ВЕРЕСЕНЬ", "ЖОВТЕНЬ", "ЛИСТОПАД", "ГРУДЕНЬ",
]

# Передається як extra_fallback_sheet_names ЛИШЕ для викликів, що шукають
# аркуш особового складу (PERSONEL_LIST_SHEET_NAME) - "Табель" додано, бо
# саме так називає свій єдиний аркуш ОБЛІК.xlsx сусідній проєкт
# generator_timesheet, ЯКЩО файл скопійований туди вручну, поза штатним
# "Зберегти кудись іще?" діалогом (штатний шлях уже перейменовує аркуш на
# поточний місяць - generator_timesheet/constants.py::MONEY_PROJECT_SHEET_NAME -
# тож зазвичай сюди й не дійде, це лише запасний варіант).
PERSONEL_LIST_FALLBACK_SHEET_NAMES = ['Табель', *_UKRAINIAN_MONTH_NAMES_UPPER]

# Передається як extra_fallback_sheet_names ЛИШЕ для викликів, що шукають
# аркуш ТВО (TVO_LIST_SHEET_NAME) - підтверджено користувачем: окремого
# аркуша "ТВО" вже немає, дані про періоди Start/End/ПОСАДА/ПІБ/ТВО (хто
# виконує обов'язки замість кого й коли - get_tvo_commander_periods,
# content/money_report_helpers.py) тепер ведуться прямо в аркуші ПОТОЧНОГО
# місяця (напр. "СЕРПЕНЬ"), окремою таблицею від самого особового складу.
# НАВМИСНО без "Табель" (на відміну від PERSONEL_LIST_FALLBACK_SHEET_NAMES
# вище) - "Табель" тепер аркуш ОСОБОВОГО СКЛАДУ (ПІДРОЗДІЛ/ПОСАДА/ЗВАННЯ/ПІБ +
# дати), структурно ІНШИЙ за ТВО-таблицю (ПОСАДА/Start/End/ЗВАННЯ/ПІБ/ТВО) -
# якби він теж був тут запасним варіантом, TVO_LIST_COLUMNS_LETTERS (літери
# A-G) прочитали б із "Табель" геть не ті колонки (напр. "ЗВАННЯ" замість
# "Start"), а не впали б чистою помилкою.
TVO_LIST_FALLBACK_SHEET_NAMES = list(_UKRAINIAN_MONTH_NAMES_UPPER)


def _normalize_header_text(value):
    """Схлопує БУДЬ-ЯКИЙ пробільний символ (включно з переносом рядка ВСЕРЕДИНІ
    самого заголовка комірки Excel - реальний ОБЛІК.xlsx переносить довгі назви
    колонок на 2 рядки, напр. комірка буквально містить "ПІДСТАВИ\nДЛЯ ІНШИХ
    СИТУАЦІЙ") в ОДИН пробіл - без цього ані пряме порівняння з очікуваною
    назвою (OPTIONAL_PERSONEL_COLUMN_NAMES/_LABEL_COLUMNS), ані прямий доступ
    за ключем (row.get("ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ")) не спрацював би - колонка
    мовчки не підхоплювалась би, хоча вона Є у файлі. Не-рядкові значення (дати
    тощо) повертаються без змін."""
    return " ".join(value.split()) if isinstance(value, str) else value


def find_file_with_any_extension(file_path):
    """Якщо файл не знайдено — пробує інші розширення."""
    if os.path.isfile(file_path):
        return file_path

    from constants import SUPPORTED_EXTENSIONS

    base = os.path.splitext(file_path)[0]
    for ext in SUPPORTED_EXTENSIONS:
        candidate = base + ext
        if os.path.isfile(candidate):
            print(f"Файл {file_path} не знайдено, використовую {candidate}")
            return candidate

    raise FileNotFoundError(f"Файл {file_path} не знайдено (також перевірено: {', '.join(base + e for e in SUPPORTED_EXTENSIONS)})")


def _resolve_sheet_name(file_path, preferred_sheet_name, engine, extra_fallback_sheet_names=None):
    """Кандидати перевіряються по черзі: preferred_sheet_name, потім
    extra_fallback_sheet_names (ЗМІСТОВНІ, специфічні для конкретного пошуку
    назви - напр. "Табель"/назва місяця для аркуша особового складу - той, хто
    викликає, явно знає, що це має бути ЗА ЗМІСТОМ), і ЛИШЕ ПОТІМ - нейтральні
    Excel-дефолти _FALLBACK_SHEET_NAMES ("Аркуш1" тощо).

    РЕАЛЬНИЙ БАГ (виправлено): раніше _FALLBACK_SHEET_NAMES перевірявся ПЕРЕД
    extra_fallback_sheet_names - якщо у файлі, ОКРІМ очікуваного аркуша (напр.
    перейменованого на "Табель"), випадково існував і сторонній аркуш на кшталт
    "Аркуш1" (без жодного стосунку до особового складу - лишився від копіювання
    тексту абощо), пошук помилково "знаходив" САМЕ його - "Аркуш1" збігався
    РАНІШЕ, ніж дійшла черга до "Табель", хоча останній - ЯВНО правильний
    (містить колонки з датами, тоді як "Аркуш1" - ні). Змістовна, надана
    викликачем підказка - завжди сильніший сигнал, ніж збіг з нейтральною
    типовою назвою Excel."""
    available_sheets = pd.ExcelFile(file_path, engine=engine).sheet_names
    ordered_fallbacks = list(extra_fallback_sheet_names or []) + _FALLBACK_SHEET_NAMES
    candidates = [preferred_sheet_name] + [s for s in ordered_fallbacks if s != preferred_sheet_name]

    for candidate in candidates:
        if candidate in available_sheets:
            return candidate

    raise ValueError(
        f"У файлі {file_path} не знайдено жодного з очікуваних аркушів ({', '.join(str(c) for c in candidates)}). "
        f"Наявні аркуші: {', '.join(str(s) for s in available_sheets)}"
    )


def excel_col_to_index(col):
    """Перетворює Excel-літеру (наприклад, 'A', 'Z', 'AA', 'AD') у індекс (0-відлік)."""
    return column_index_from_string(col.upper()) - 1


def index_to_excel_col(index):
    """Перетворює індекс (0-відлік) у Excel-літеру (наприклад, 0 -> 'A', 26 -> 'AA')."""
    return get_column_letter(index + 1)


# Іменовані колонки, які підхоплюються, ЯКЩО вони є в ОБЛІК.xlsx (десь після
# fixed_columns_count, незалежно від позиції відносно дат) - за замовчуванням
# ПІДСТАВА/ПІДСТАВИ (content.money_report_helpers.build_pidstavy_extra_grounds_by_person:
# персональна підстава для конкретної людини, додається до "Підстава для
# виплати"/"для не виплати" для БУДЬ-ЯКОГО пункту рапорту). Обидва написання
# (однина й множина) підтримуються, бо реальний ОБЛІК.xlsx використовує
# однину ("ПІДСТАВА"). Немає жодної з цих колонок у файлі - просто не
# підхоплюється, без жодної помилки. Значення живе в constants.py - тут лише
# default=None (не саму константу, щоб уникнути циклічного імпорту, див.
# коментар вгорі файлу), реальне значення підставляється нижче в тілі функції.


def _detect_columns_letters_for_resolved_sheet(file_path, resolved_sheet_name, engine, month, year, fixed_columns_count, optional_column_names):
    """Спільна серцевина для detect_personel_list_columns_letters/
    detect_optional_sheet_columns_letters - рахує літери колонок ДЛЯ ВЖЕ
    визначеного (наявного) аркуша resolved_sheet_name: перші
    fixed_columns_count колонок + дати обраного місяця/року + будь-яка з
    optional_column_names. Повертає None, якщо колонок з датами не знайдено -
    що з цим робити (кинути помилку чи мовчки пропустити) вирішує ВИКЛИКАЧ,
    бо це залежить від того, ОБОВ'ЯЗКОВИЙ цей аркуш чи ні."""
    header = pd.read_excel(file_path, sheet_name=resolved_sheet_name, engine=engine, nrows=0)
    columns = list(header.columns)

    fixed_letters = [index_to_excel_col(i) for i in range(min(fixed_columns_count, len(columns)))]

    month_int = int(month)
    year_int = int(year)

    date_columns = [
        (col.day, index_to_excel_col(i))
        for i, col in enumerate(columns)
        if i >= fixed_columns_count
        and isinstance(col, (pd.Timestamp, datetime))
        and col.month == month_int
        and col.year == year_int
    ]
    date_columns.sort(key=lambda item: item[0])
    date_letters = [letter for _, letter in date_columns]

    if not date_letters:
        return None

    optional_letters = [
        index_to_excel_col(i)
        for i, col in enumerate(columns)
        if i >= fixed_columns_count and isinstance(col, str) and _normalize_header_text(col) in optional_column_names
    ]

    return fixed_letters + date_letters + optional_letters


def detect_personel_list_columns_letters(file_path, sheet_name, month, year, fixed_columns_count=4, optional_column_names=None):
    """
    Автоматично визначає літери колонок у файлі ОБЛІК.xlsx:
    перші `fixed_columns_count` колонок (ПІДРОЗДІЛ, ПОСАДА, звання фактичне, ПІБ)
    + колонки з датами обраного місяця/року, відсортовані за днем
    + будь-яка з optional_column_names (за назвою заголовка, у будь-якій позиції
    після fixed_columns_count), якщо вона є у файлі. optional_column_names=None
    (за замовчуванням) -> constants.OPTIONAL_PERSONEL_COLUMN_NAMES.

    Завжди шукає САМЕ аркуш особового складу, тож додає
    PERSONEL_LIST_FALLBACK_SHEET_NAMES (Табель/назви місяців) до запасних
    варіантів - на відміну від read_datafile/read_optional_datafile, де це
    вирішує викликач (не кожен виклик читає саме особовий склад).
    """
    if optional_column_names is None:
        from constants import OPTIONAL_PERSONEL_COLUMN_NAMES
        optional_column_names = OPTIONAL_PERSONEL_COLUMN_NAMES

    file_path = find_file_with_any_extension(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'
    resolved_sheet_name = _resolve_sheet_name(file_path, sheet_name, engine, PERSONEL_LIST_FALLBACK_SHEET_NAMES)

    letters = _detect_columns_letters_for_resolved_sheet(
        file_path, resolved_sheet_name, engine, month, year, fixed_columns_count, optional_column_names,
    )
    if letters is None:
        month_int, year_int = int(month), int(year)
        raise ValueError(
            f"У файлі {file_path} (аркуш {resolved_sheet_name}) не знайдено колонок з датами "
            f"за {month_int:02d}.{year_int}."
        )
    return letters


def read_optional_pridani_sheet(file_path, sheet_name, month, year):
    """Читає аркуш "ПРИДАНІ" (приданий особовий склад - constants.PRIDANI_SHEET_NAME)
    У САМОМУ ОБЛІК.xlsx - НАВМИСНО НЕ за позицією колонок, як "Табель"
    (перші 4 фіксовані + дати) - реальний файл довів, що структура ЗОВСІМ
    інша: "№ з/п"/"ПІДРОЗДІЛ" (походження людини)/"ПРИДАНИЙ ДО" (до якого
    підрозділу ЦЬОГО батальйону прикріплено)/"ВІЙСЬКОВЕ ЗВАННЯ"/"ПІБ"/дати
    прибуття-вибуття/НАЯВНІСТЬ/БЧС/... і ЛИШЕ ПІСЛЯ всього цього - дні місяця
    (ті самі значення 30/70/100/170 тощо, що й "Табель") - "ПОСАДА" немає
    ВЗАГАЛІ. Спроба читати за позицією (як спочатку) давала рядки з хибними
    ключами (перші 4 колонки замість ПІДРОЗДІЛ/ПОСАДА/ЗВАННЯ/ПІБ) і падала з
    KeyError на 'ПОСАДА' у content/br_general_catalog.py - виявлено
    користувачем на реальному прикладі. Тому колонки шукаються ЗА НАЗВОЮ.

    Повертає [] (а не None/помилку), якщо аркуша немає, немає колонки "ПІБ",
    чи немає колонок з датами обраного місяця - приданий особовий склад
    просто відсутній, як і при відсутності будь-якого іншого опційного аркуша.

    Рядки результату мають ТІ САМІ ключі, що й звичайний особовий склад
    ("ПІДРОЗДІЛ"/"ПОСАДА"/"ЗВАННЯ"/"ПІБ" + дати), щоб content/br_general_catalog.py
    працював з ними без жодних змін:
    - "ПІДРОЗДІЛ" - САМЕ "ПРИДАНИЙ ДО" (до якого підрозділу ЦЬОГО батальйону
      прикріплено, напр. "2РМП"/"ДШР"/"МП" - ті самі, впізнавані коди, що й у
      звичайного особового складу, включно з "МП" -> МЕДИК, resolve_status_category)
      - для групування в БР за секціями САМЕ цього батальйону; власне
      походження людини (перша колонка "ПІДРОЗДІЛ" аркуша) для БР значення
      не має.
    - "ЗВАННЯ" - "ВІЙСЬКОВЕ ЗВАННЯ".
    - "ПОСАДА" - порожній рядок (приданий особовий склад посади в ЦЬОМУ
      батальйоні не має - лише звання й підрозділ, до якого прикріплений)."""
    file_path = find_file_with_any_extension(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'

    available_sheets = pd.ExcelFile(file_path, engine=engine).sheet_names
    if sheet_name not in available_sheets:
        return []

    header = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine, nrows=0)
    normalized_columns = [_normalize_header_text(col) if isinstance(col, str) else col for col in header.columns]

    def _letter_for(name):
        index = next((i for i, col in enumerate(normalized_columns) if col == name), None)
        return index_to_excel_col(index) if index is not None else None

    pib_letter = _letter_for("ПІБ")
    if pib_letter is None:
        return []

    month_int, year_int = int(month), int(year)
    date_letters = [
        index_to_excel_col(i) for i, col in enumerate(normalized_columns)
        if isinstance(col, (pd.Timestamp, datetime)) and col.month == month_int and col.year == year_int
    ]
    if not date_letters:
        return []

    pidrozdil_letter = _letter_for("ПРИДАНИЙ ДО")
    zvannya_letter = _letter_for("ВІЙСЬКОВЕ ЗВАННЯ")
    use_letters = [letter for letter in (pidrozdil_letter, zvannya_letter, pib_letter) if letter is not None] + date_letters

    raw_rows = read_datafile(file_path, sheet_name, use_letters).get("rows", [])

    date_column_names = [normalized_columns[excel_col_to_index(letter)] for letter in date_letters]

    def _blank_if_missing(value):
        # pd.isna(value) на не-float значеннях (напр. рядку) сама по собі
        # безпечна (повертає False), тож окрема isinstance-перевірка тут не
        # потрібна - на відміну від "ПІБ" нижче, де None/NaN означає "пропустити
        # рядок цілком", а не просто "порожнє значення поля".
        return "" if value is None or pd.isna(value) else value

    result = []
    for raw_row in raw_rows:
        pib = raw_row.get("ПІБ")
        if pib is None or (isinstance(pib, float) and pd.isna(pib)):
            continue
        # str+strip (не лише для перевірки на порожність, а й для самого
        # значення, що йде в result) - реальний файл довів: випадковий зайвий
        # пробіл у клітинці ПІБ (людська помилка вводу) робить ідентичність
        # людини НЕСУМІСНОЮ з тим самим ПІБ, витягнутим з тексту вже
        # згенерованого документа (там пробіл перед ";" завжди відкидається
        # regex-ом) - людина мовчки не проходила жодну звірку "хто ще
        # валідний" (test_no_unexpected_pib_in_daily_br), хоч насправді мала.
        pib = str(pib).strip()
        if not pib:
            continue
        row = {
            "ПІДРОЗДІЛ": _blank_if_missing(raw_row.get("ПРИДАНИЙ ДО")),
            "ПОСАДА": "",
            "ЗВАННЯ": _blank_if_missing(raw_row.get("ВІЙСЬКОВЕ ЗВАННЯ")),
            "ПІБ": pib,
        }
        for date_column_name in date_column_names:
            row[date_column_name] = raw_row.get(date_column_name)
        result.append(row)
    return result


def read_datafile(file_path, sheet_name, use_cols_letters, date_can_be_empty=False, extra_fallback_sheet_names=None):
    file_path = find_file_with_any_extension(file_path)

    ext = os.path.splitext(file_path)[1].lower()
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'

    resolved_sheet_name = _resolve_sheet_name(file_path, sheet_name, engine, extra_fallback_sheet_names)

    print(f"Read document {file_path}, {resolved_sheet_name}...")
    data = pd.read_excel(file_path, sheet_name=resolved_sheet_name, engine=engine)
    # Нормалізує заголовки колонок (_normalize_header_text) ПРЯМО на DataFrame,
    # ОДИН раз, тут - інакше рядок з перенесеним заголовком (див. її docstring)
    # ставав би ключем результату з "живим" переносом рядка всередині, і жоден
    # прямий row.get("ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ") ніде далі його не знайшов би.
    data.columns = [_normalize_header_text(col) for col in data.columns]

    columns_data = [excel_col_to_index(letter) for letter in use_cols_letters]
    column_names = [data.columns[i] for i in columns_data]

    return get_data(data, column_names, date_can_be_empty)


def read_optional_datafile(file_path, sheet_name, use_cols_letters, date_can_be_empty=False, sheet_can_be_missing=False, extra_fallback_sheet_names=None):
    """Як read_datafile, але якщо файл відсутній — не падає, а повертає порожній
    список рядків. sheet_can_be_missing=True — так само, якщо файл є, але в
    ньому немає ані sheet_name, ані жодного із запасних (_resolve_sheet_name
    підіймає ValueError) - для аркушів, чия відсутність САМА ПО СОБІ означає
    "немає таких даних", а не помилку файлу (напр. аркуш "ТВО" у файлі,
    збереженому сусіднім проєктом generator_timesheet, який про ТВО не знає
    взагалі). За замовчуванням False, щоб типовий виклик (напр. ПРИДАНІ.xlsx)
    і далі падав голосно на помилку в назві аркуша, а не мовчки її ігнорував."""
    try:
        return read_datafile(file_path, sheet_name, use_cols_letters, date_can_be_empty, extra_fallback_sheet_names).get("rows", [])
    except FileNotFoundError:
        print(f"File {file_path} not found. Skipping.")
        return []
    except ValueError:
        if sheet_can_be_missing:
            print(f"Аркуш {sheet_name} не знайдено у файлі {file_path}. Продовжую без цих даних.")
            return []
        raise


def get_data(transfer_data, columns_transfer, date_can_be_empty=False):
    if date_can_be_empty:
        rows = [
            {
                col: (date_to_str(val) if not pd.isna(val) else None)
                if isinstance(val := row[col], (pd.Timestamp, datetime)) else val
                for col in columns_transfer
            }
            for _, row in transfer_data[columns_transfer].iterrows()
        ]
    else:
        # Порожня клітинка в колонці, де хоч ОДНЕ інше значення - дата (напр.
        # "ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ", заповнена лише для частини людей),
        # приходить від pandas як pd.NaT, а не NaN - і pd.NaT ТЕЖ проходить
        # isinstance(..., datetime) (сумісність pandas), тож без цієї перевірки
        # date_to_str(NaT) падає з "NaTType does not support strftime". Порожня
        # дата тут - це не помилка (сама колонка необов'язкова), тому теж -> None.
        rows = [
            {
                col: (date_to_str(val) if not pd.isna(val) else None)
                if isinstance(val := row[col], (pd.Timestamp, datetime)) else val
                for col in columns_transfer
            }
            for _, row in transfer_data[columns_transfer].iterrows()
        ]

    return {"rows": rows, "columns": columns_transfer}
