"""Книга налаштувань resources/НАЛАШТУВАННЯ.xlsx: читання й запис аркушів.

Усе, що користувач налаштовує під свій підрозділ і місяць (назви, номери
документів, підстави, списки людей), живе в цій книзі, а не в constants.py
чи data.json. Код оновлюється з публічного дзеркала через "git reset --hard"
(scripts/common.ps1), тож правка в constants.py зникла б при наступному
запуску START.bat, а теку resources/ git не чіпає. Excel до того ж звичніший
за JSON: немає ризику зламати файл зайвою комою.

Модуль однаковий у generator_br_and_report_for_money і generator_timesheet
(копія в кожному проєкті: пакети двох проєктів навмисно розділені, див.
кореневий conftest.py). Які саме аркуші й колонки має книга проєкту - у його
utils/project_settings.py, тут лише спільні правила:

- аркуш-таблиця: над таблицею - назва аркуша (темна смуга), підказка й, за
  потреби, кольорові групи колонок; далі - рядок заголовків і дані. Рядок
  заголовків при читанні шукається сам (серед перших рядків аркуша), тож
  користувач може змінити підказку чи висоту рядків, нічого не зламавши;
- колонки шукаються за назвою (без урахування регістру й зайвих пробілів),
  тож порядок колонок довільний, а зайві колонки (власні примітки
  користувача) ігноруються;
- рядок, у якому порожні всі відомі колонки, пропускається;
- помилка називає аркуш, рядок і колонку, щоб було ясно, що виправити.

Без залежностей від constants.py (його імпортує сам constants.py) - лише
openpyxl і colorama (червоне пояснення помилки, stop)."""
import colorsys
import math
import os
import re
import tempfile
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime
from xml.etree import ElementTree

from colorama import Fore, Style
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

SETTINGS_FILE_NAME = "НАЛАШТУВАННЯ.xlsx"

# Види значень колонки/параметра.
TEXT = "text"  # рядок (число з клітинки теж стає рядком: 350 -> "350")
NUMBER_TEXT = "number_text"  # рядок, який записується числом, якщо це лише цифри (номер документа)
LINES = "lines"  # кілька рядків в одній клітинці (Alt+Enter) -> список рядків
LIST = "list"  # значення через кому -> список
INT = "int"  # ціле число чи None
BOOL = "bool"  # Так/Ні
DATE = "date"  # дата -> рядок "дд.мм.рррр"
DAY = "day"  # день місяця 1-31 чи PREVIOUS_MONTH_DAY
# Колір - сама заливка клітинки (користувач обирає її в Excel кнопкою "Колір
# заливки") -> "FFRRGGBB" чи None. Текст клітинки береться, лише якщо заливки
# немає, а вписано код кольору (напр. FFC000).
COLOR = "color"

YES = "Так"
NO = "Ні"
# Рядок "попередній" аркуша номерів документів - останній день місяця перед
# обраним (той самий "перехідний" день, що й колись у constants.py).
PREVIOUS_MONTH_DAY = "попередній"
DATE_FORMAT = "%d.%m.%Y"
# "dd.MM.yyyy", а не "dd.mm.yyyy": деякі програми перегляду читають "mm" як
# хвилини (див. generator_timesheet/constants.py::DATE_NUMBER_FORMAT).
EXCEL_DATE_FORMAT = "dd.MM.yyyy"

# Кольори ярликів аркушів (і розділів змісту на аркуші "Інструкція") - за
# тим, як часто й навіщо аркуш відкривають.
TAB_INFO = "1F3864"  # інструкція
TAB_MONTHLY = "548235"  # заповнюється щомісяця
TAB_REFERENCES = "2F5597"  # підстави, списки людей
TAB_REPORT = "C55A11"  # пункти й тексти рапорту
TAB_TEXTS = "7030A0"  # тексти документів
TAB_GENERAL = "7F7F7F"  # загальні назви

# Кольори груп колонок (смуга над заголовками), по черзі.
GROUP_COLORS = ("548235", "2F5597", "C55A11", "7030A0", "806000", "1F7A7A")

# Рядки аркуша-таблиці: 1 - назва, 2 - підказка, 3 - групи колонок,
# 4 - заголовки колонок, з 5 - дані.
TITLE_ROW = 1
HINT_ROW = 2
GROUP_ROW = 3
HEADER_ROW = 4
FIRST_DATA_ROW = 5
# Серед скількох перших рядків аркуша шукати рядок заголовків при читанні.
_HEADER_SEARCH_ROWS = 15
# Скільки рядків під заголовком покривають випадний список Так/Ні й підсвітка
# - із запасом на рядки, які користувач допише сам.
_VALIDATION_ROWS = 1000
# Скільки порожніх рядків з рамками лишати під даними - видно, куди дописувати.
_EMPTY_ROWS = 3
# Найменша ширина смуги з назвою й підказкою (у символах): вузька таблиця
# (одна-дві колонки) інакше дала б підказку стовпчиком у кілька слів.
_MIN_BANNER_WIDTH = 90
_DEFAULT_COLUMN_WIDTH = 8.43

_TITLE_FILL = PatternFill("solid", start_color="1F3864")
_HEADER_FILL = PatternFill("solid", start_color="D9E2F3")
_STRIPE_FILL = PatternFill("solid", start_color="F3F6FB")
_EDIT_FILL = PatternFill("solid", start_color="FFF8E1")
_LABEL_FILL = PatternFill("solid", start_color="EEF3FA")
_POINT_FILL = PatternFill("solid", start_color="FBE5D6")
_MUTED_FILL = PatternFill("solid", start_color="EDEDED")
_YES_FILL = PatternFill("solid", start_color="D8F0D2", end_color="D8F0D2")
_NO_FILL = PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
_THIN = Side(style="thin", color="C9D1DC")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HEADER_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=Side(style="medium", color="2F5597"))
_POINT_BORDER = Border(left=_THIN, right=_THIN, top=Side(style="medium", color="C55A11"), bottom=_THIN)

_TITLE_FONT = Font(bold=True, size=14, color="FFFFFF")
_HINT_FONT = Font(italic=True, size=10, color="595959")
_HEADER_FONT = Font(bold=True, color="1F3864")
_GROUP_FONT = Font(bold=True, color="FFFFFF")
_BOLD = Font(bold=True)
_POINT_FONT = Font(bold=True, color="843C0C")
_MUTED_FONT = Font(italic=True, color="7F7F7F")
_DIM_FONT = Font(color="A6A6A6")
_NOTE_FONT = Font(italic=True, size=10, color="595959")
_KEY_FONT = Font(size=9, color="8C8C8C")
_LINK_FONT = Font(color="0563C1", underline="single", bold=True)

_WRAP = Alignment(wrap_text=True, vertical="top")
_TOP = Alignment(vertical="top")
_LEFT = Alignment(horizontal="left", vertical="top")
_CENTER = Alignment(horizontal="center", vertical="top", wrap_text=True)
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

_TRUE_WORDS = {"так", "т", "true", "yes", "y", "1", "+", "x", "х"}
_FALSE_WORDS = {"ні", "н", "false", "no", "n", "0", "-"}
_LIST_SEPARATORS_RE = re.compile(r"[,;\n]")
_DATE_TEXT_RE = re.compile(r"\d{1,2}\.\d{1,2}\.\d{4}")


class SettingsError(ValueError):
    """Помилка в книзі налаштувань. Текст - для користувача: де саме й що не так."""


class SettingsStop(SystemExit):
    """Зупинка програми через налаштування (stop). Пояснення вже надруковано
    червоним, тож Python його вдруге не друкує, а лише завершується з кодом 1;
    str() - те саме пояснення (тести)."""

    def __init__(self, message):
        super().__init__(1)
        self.message = message

    def __str__(self):
        return self.message


def stop(problem, steps=()):
    """SettingsStop (raise stop(...)) з поясненням, надрукованим червоним: що
    не так (problem) і пронумеровані кроки, що зробити (steps). Колись
    SystemExit друкував текст білим, а червоним лаунчер
    (scripts/launcher.ps1) писав лише "Програма завершилась з помилкою", -
    і користувач, який читає тільки червоне, не знав, що робити."""
    lines = ["ПОМИЛКА В НАЛАШТУВАННЯХ - програму зупинено.", problem]
    if steps:
        lines.append("Що зробити:")
        lines.extend(f"  {number}. {step}" for number, step in enumerate(steps, start=1))
    message = "\n".join(lines)
    print(Fore.RED + message + Style.RESET_ALL)
    return SettingsStop(message)


def fix_in_book_steps(settings_path):
    """Кроки для помилки в самій книзі: текст помилки вже називає аркуш,
    рядок і колонку."""
    return (
        f"Відкрийте {settings_path} в Excel.",
        "Знайдіть аркуш, рядок і колонку, названі вище, і впишіть правильне значення.",
        "Збережіть файл, закрийте Excel і запустіть програму ще раз.",
    )


@dataclass(frozen=True)
class Column:
    """Колонка аркуша-таблиці. key - ключ у словнику рядка, header - заголовок
    у книзі. required - колонка мусить бути в аркуші; nonempty - клітинка
    мусить бути заповнена в кожному непорожньому рядку. default - значення
    порожньої клітинки BOOL-колонки. note - підказка (примітка до заголовка).
    choices - випадний список у клітинках (вписати інше теж можна). group -
    назва групи колонок (кольорова смуга над заголовками); center -
    вирівнювання по центру. span - скільки колонок аркуша займає колонка
    (об'єднані клітинки): так широка колонка однієї таблиці не розтягує вузьку
    колонку іншої таблиці того самого аркуша; width тоді - її загальна ширина."""
    key: str
    header: str
    kind: str = TEXT
    required: bool = True
    nonempty: bool = False
    default: object = None
    width: int = 18
    wrap: bool = False
    note: str = ""
    choices: tuple = ()
    group: str = ""
    center: bool = False
    span: int = 1


@dataclass(frozen=True)
class Setting:
    """Параметр аркуша "Назва | Значення | Пояснення | Параметр". key -
    незмінна технічна назва (за нею параметр і шукається), label - зрозуміла
    назва для людини. required - параметр мусить бути в аркуші й мати значення."""
    key: str
    kind: str = TEXT
    required: bool = True
    note: str = ""
    label: str = ""


def location(sheet_name, row=None, header=None):
    """'НАЛАШТУВАННЯ.xlsx, аркуш "Х", рядок 5, колонка "У"' - для повідомлень."""
    parts = [SETTINGS_FILE_NAME, f'аркуш "{sheet_name}"']
    if row is not None:
        parts.append(f"рядок {row}")
    if header is not None:
        parts.append(f'колонка "{header}"')
    return ", ".join(parts)


def _normalize_header(value):
    return " ".join(str(value).split()).casefold() if value is not None else ""


def normalize_name(value):
    """Назва для порівняння (групи підстав, книги тощо): регістр і зайві
    пробіли не важать."""
    return " ".join(str(value).split()).casefold()


def is_blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def cell_text(value):
    """Текст клітинки: None -> "", 350 -> "350", 350.0 -> "350", дата ->
    "дд.мм.рррр", рядок - без зайвих пробілів на краях і з переносом рядка
    "\\n" (Excel на Windows інколи зберігає "\\r\\n")."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return YES if value else NO
    if isinstance(value, datetime):
        return value.strftime(DATE_FORMAT)
    if isinstance(value, date):
        return value.strftime(DATE_FORMAT)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (int, float)):
        return str(value)
    return str(value).replace("\r\n", "\n").replace("\r", "\n").strip()


def split_list(text):
    """"А, Б; В" -> ["А", "Б", "В"] (кома, крапка з комою чи новий рядок)."""
    return [item.strip() for item in _LIST_SEPARATORS_RE.split(text or "") if item.strip()]


def parse_value(raw, kind, where, default=None):
    """Значення клітинки raw як kind. where - місце для повідомлення про помилку."""
    if kind in (TEXT, NUMBER_TEXT):
        return cell_text(raw)
    if kind == LINES:
        return [line.strip() for line in cell_text(raw).split("\n") if line.strip()]
    if kind == LIST:
        return split_list(cell_text(raw))
    if kind == INT:
        return _parse_int(raw, where)
    if kind == BOOL:
        return _parse_bool(raw, where, default)
    if kind == DATE:
        return _parse_date(raw, where)
    if kind == DAY:
        return _parse_day(raw, where)
    if kind == COLOR:
        return raw or None
    raise ValueError(f"Невідомий вид значення {kind!r}")


def _parse_int(raw, where):
    if is_blank(raw):
        return None
    if isinstance(raw, bool):
        raise SettingsError(f"{where}: очікується ціле число, а не {cell_text(raw)!r}.")
    if isinstance(raw, int):
        return raw
    if isinstance(raw, float) and raw.is_integer():
        return int(raw)
    text = cell_text(raw)
    if re.fullmatch(r"[+-]?\d+", text):
        return int(text)
    raise SettingsError(f"{where}: очікується ціле число, а не {text!r}.")


def _parse_bool(raw, where, default):
    if is_blank(raw):
        return default
    if isinstance(raw, bool):
        return raw
    word = cell_text(raw).casefold()
    if word in _TRUE_WORDS:
        return True
    if word in _FALSE_WORDS:
        return False
    raise SettingsError(f'{where}: очікується "{YES}" або "{NO}", а не {cell_text(raw)!r}.')


def _parse_date(raw, where):
    if is_blank(raw):
        return None
    if isinstance(raw, (datetime, date)):
        return raw.strftime(DATE_FORMAT)
    text = cell_text(raw)
    if _DATE_TEXT_RE.fullmatch(text):
        try:
            return datetime.strptime(text, DATE_FORMAT).strftime(DATE_FORMAT)
        except ValueError:
            pass
    raise SettingsError(f"{where}: очікується дата дд.мм.рррр, а не {text!r}.")


def _parse_day(raw, where):
    if is_blank(raw):
        return None
    text = cell_text(raw)
    if text.casefold().startswith("попер"):
        return PREVIOUS_MONTH_DAY
    day = _parse_int(raw, where) if re.fullmatch(r"\d+", text) else None
    if day == 0:
        return PREVIOUS_MONTH_DAY
    if day is not None and 1 <= day <= 31:
        return day
    raise SettingsError(f'{where}: очікується день місяця 1-31 або "{PREVIOUS_MONTH_DAY}", а не {text!r}.')


# ---------------------------------------------------------------- читання


@contextmanager
def open_for_reading(path):
    """Книга лише для читання (значення формул - останні збережені Excel).
    Закривається після блоку with: у read-only режимі openpyxl тримає файл
    відкритим, і на Windows його тоді не можна було б перезаписати."""
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except PermissionError as error:
        raise SettingsError(f"Не вдалося відкрити {path}: файл зайнятий іншою програмою.") from error
    except Exception as error:
        raise SettingsError(f"Не вдалося відкрити {path} як книгу Excel: {error}") from error
    try:
        yield workbook
    finally:
        workbook.close()


def find_sheet(workbook, sheet_name):
    """Аркуш за назвою без урахування регістру й зайвих пробілів, або None."""
    wanted = normalize_name(sheet_name)
    for name in workbook.sheetnames:
        if normalize_name(name) == wanted:
            return workbook[name]
    return None


def _require_sheet(workbook, sheet_name):
    sheet = find_sheet(workbook, sheet_name)
    if sheet is None:
        raise SettingsError(
            f'{SETTINGS_FILE_NAME}: немає аркуша "{sheet_name}". Наявні аркуші: {", ".join(workbook.sheetnames)}.'
        )
    return sheet


def _column_positions(sheet_name, header_values, columns):
    normalized = [_normalize_header(value) for value in header_values]
    positions = {}
    for column in columns:
        index = next((i for i, value in enumerate(normalized) if value == _normalize_header(column.header)), None)
        if index is None:
            if column.required:
                raise SettingsError(f'{location(sheet_name)}: немає колонки "{column.header}" у рядку заголовків.')
            continue
        positions[column.key] = index
    return positions


def _find_header(sheet, label, columns, first_row=1, last_row=None):
    """(номер рядка заголовків, {key колонки: індекс}) - перший серед перших
    рядків таблиці (з first_row), де є заголовки всіх обов'язкових колонок (над
    таблицею - назва й підказка). Такого рядка немає - помилка з назвою колонки,
    якої бракує в найповнішому з рядків-кандидатів."""
    wanted = {_normalize_header(column.header) for column in columns if column.required}
    best_values, best_found = (), -1
    search_last = first_row + _HEADER_SEARCH_ROWS - 1
    if last_row is not None:
        search_last = min(search_last, last_row)
    # Рядки - одразу списком: недочитаний генератор книги read-only тримав би
    # файл відкритим до прибирання сміття, і Windows не дала б його перезаписати.
    rows = list(sheet.iter_rows(min_row=first_row, max_row=search_last, values_only=True))
    for row_number, values in enumerate(rows, start=first_row):
        present = {_normalize_header(value) for value in values}
        if wanted <= present:
            return row_number, _column_positions(label, values, columns)
        if len(wanted & present) > best_found:
            best_values, best_found = values, len(wanted & present)
    _column_positions(label, best_values, columns)
    raise SettingsError(f"{location(label)}: не знайдено рядка із заголовками колонок.")


# ---------------------------------------------------------------- кілька таблиць на аркуші

# Таблиця на аркуші з кількома таблицями починається з кольорової смуги
# "■ Назва таблиці" у колонці A (далі - підказка, заголовки, рядки) і триває до
# смуги наступної таблиці. За назвою в смузі програма таблицю й шукає, тож її,
# як і заголовки колонок, не змінюють.
SECTION_MARK = "■"


def section_band_text(name):
    return f"{SECTION_MARK} {name}"


def _band_name(value):
    if isinstance(value, str):
        text = value.strip()
        if text.startswith(SECTION_MARK):
            return normalize_name(text[len(SECTION_MARK):])
    return None


def _bands(sheet):
    """{назва таблиці (нормалізована): номер рядка її смуги} на аркуші."""
    bands = {}
    rows = list(sheet.iter_rows(min_col=1, max_col=1, values_only=True))
    for row_number, values in enumerate(rows, start=1):
        name = _band_name(values[0] if values else None)
        if name is not None and name not in bands:
            bands[name] = row_number
    return bands


def table_label(sheet_name, section=None):
    """Аркуш (і таблиця на ньому) для повідомлень: location(table_label(...))
    дає 'аркуш "Люди", таблиця "Ліміт категорії 100"'."""
    return f'{sheet_name}", таблиця "{section}' if section else sheet_name


def locate_table(workbook, sheet_name, section=None, legacy_sheet=None):
    """(аркуш, підпис для повідомлень, перший рядок, останній рядок чи None) -
    де в книзі таблиця, або None, якщо її немає.

    section=None - уся таблиця аркуша sheet_name (аркуш з однією таблицею).
    Інакше - таблиця зі смугою "■ section" на аркуші sheet_name (sheet_name=None -
    на будь-якому аркуші: так один проєкт читає таблицю книги іншого, не
    знаючи, на якому вона аркуші); немає - аркуш legacy_sheet (типово -
    аркуш з назвою section) без смуг: книга, записана, коли кожна таблиця
    мала власний аркуш."""
    if section is None:
        sheet = find_sheet(workbook, sheet_name)
        return (sheet, sheet_name, 1, None) if sheet is not None else None
    candidates = [find_sheet(workbook, sheet_name)] if sheet_name else list(workbook.worksheets)
    wanted = normalize_name(section)
    for sheet in candidates:
        if sheet is None:
            continue
        bands = _bands(sheet)
        if wanted in bands:
            start = bands[wanted]
            following = [row for row in bands.values() if row > start]
            return sheet, table_label(sheet.title, section), start + 1, (min(following) - 1 if following else None)
    legacy = find_sheet(workbook, legacy_sheet or section)
    if legacy is not None and not _bands(legacy):
        return legacy, legacy.title, 1, None
    return None


def has_table(workbook, sheet_name, section=None, legacy_sheet=None):
    return locate_table(workbook, sheet_name, section, legacy_sheet) is not None


def _require_table(workbook, sheet_name, section, legacy_sheet):
    located = locate_table(workbook, sheet_name, section, legacy_sheet)
    if located is None:
        if section is None:
            _require_sheet(workbook, sheet_name)
        raise SettingsError(
            f'{SETTINGS_FILE_NAME}: немає таблиці "{section}" (смуга "{section_band_text(section)}" на аркуші '
            f'"{sheet_name}").'
        )
    return located


def missing_tables(path, tables):
    """Таблиці tables ([(аркуш, назва таблиці чи None), ...]), яких у книзі path
    немає саме в теперішньому вигляді (на своєму аркуші зі смугою)."""
    with open_for_reading(path) as workbook:
        missing = []
        for sheet_name, section in tables:
            sheet = find_sheet(workbook, sheet_name)
            if sheet is None or (section is not None and normalize_name(section) not in _bands(sheet)):
                missing.append((sheet_name, section))
    return missing


def read_rows(workbook, sheet_name, columns, *, optional_sheet=False, section=None, legacy_sheet=None):
    """[(номер рядка, {key: значення}), ...] - непорожні рядки таблиці (аркуша
    sheet_name чи, з section, таблиці "■ section" на ньому - locate_table).
    Колонки, якої немає в таблиці (required=False), у словнику рядка мають
    значення порожньої клітинки. COLOR-колонка - колір заливки клітинки (рядок,
    де заповнено лише її, все одно порожній). Таблиці немає: optional_sheet -
    [], інакше помилка."""
    if optional_sheet and not has_table(workbook, sheet_name, section, legacy_sheet):
        return []
    sheet, label, first_row, last_row = _require_table(workbook, sheet_name, section, legacy_sheet)
    header_row, positions = _find_header(sheet, label, columns, first_row, last_row)
    color_keys = [column.key for column in columns if column.key in positions and column.kind == COLOR]
    result = []
    rows = list(sheet.iter_rows(min_row=header_row + 1, max_row=last_row, values_only=not color_keys))
    for row_number, cells in enumerate(rows, start=header_row + 1):
        values = tuple(getattr(cell, "value", None) for cell in cells) if color_keys else cells
        raw = {key: (values[index] if index < len(values) else None) for key, index in positions.items()}
        if all(is_blank(value) for value in raw.values()):
            continue
        for key in color_keys:
            index = positions[key]
            raw[key] = cell_fill_color(cells[index] if index < len(cells) else None, workbook) or _hex_color(raw[key])
        parsed = {}
        for column in columns:
            where = location(label, row_number, column.header)
            value = parse_value(raw.get(column.key), column.kind, where, column.default)
            if column.nonempty and value in (None, "", []):
                raise SettingsError(f"{where}: клітинка не може бути порожньою.")
            parsed[column.key] = value
        result.append((row_number, parsed))
    return result


_HEX_RE = re.compile(r"#?(?:[0-9A-Fa-f]{2})?([0-9A-Fa-f]{6})")
# Порядок кольорів теми, як на них посилається клітинка (індекс): світлий і
# темний основні - навпаки, ніж у самому файлі теми.
_THEME_ORDER = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
                "hlink", "folHlink")
_THEME_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def _hex_color(text):
    """Код кольору, вписаний текстом (FFC000, #FFC000, FFFFC000) -> "FFRRGGBB"."""
    match = _HEX_RE.fullmatch(cell_text(text))
    return f"FF{match.group(1).upper()}" if match else None


def _theme_palette(workbook):
    """Кольори теми книги (Excel пропонує їх верхніми рядками палітри "Колір
    заливки") - 6-значні коди за індексом теми."""
    cached = getattr(workbook, "_settings_theme_palette", None)
    if cached is not None:
        return cached
    palette = []
    try:
        root = ElementTree.fromstring(workbook.loaded_theme)
        scheme = root.find(".//a:clrScheme", _THEME_NS)
        for name in _THEME_ORDER:
            element = scheme.find(f"a:{name}", _THEME_NS)
            value = None
            if element is not None and len(element):
                value = element[0].get("val") if element[0].tag.endswith("srgbClr") else element[0].get("lastClr")
            palette.append(value or "000000")
    except (AttributeError, TypeError, ElementTree.ParseError):
        palette = []
    try:
        workbook._settings_theme_palette = palette
    except AttributeError:
        pass
    return palette


def _apply_tint(hex_color, tint):
    """Відтінок кольору теми (світліше/темніше), як рахує Excel: tint < 0 -
    темніше, > 0 - світліше."""
    red, green, blue = (int(hex_color[index:index + 2], 16) / 255 for index in (0, 2, 4))
    hue, lightness, saturation = colorsys.rgb_to_hls(red, green, blue)
    lightness = lightness * (1 + tint) if tint < 0 else lightness * (1 - tint) + tint
    return "".join(f"{round(part * 255):02X}" for part in colorsys.hls_to_rgb(hue, lightness, saturation))


def cell_fill_color(cell, workbook=None):
    """Колір заливки клітинки "FFRRGGBB" чи None (без заливки). Розуміє всі
    способи, якими Excel зберігає колір: код RGB, колір теми з відтінком і
    колір зі старої палітри (indexed)."""
    fill = getattr(cell, "fill", None)
    if fill is None or getattr(fill, "fill_type", None) in (None, "none"):
        return None
    color = fill.fgColor
    if color.type == "rgb" and isinstance(color.rgb, str) and len(color.rgb) >= 6:
        return f"FF{color.rgb[-6:].upper()}"
    if color.type == "indexed" and isinstance(color.indexed, int) and 0 <= color.indexed < len(COLOR_INDEX):
        return f"FF{COLOR_INDEX[color.indexed][-6:].upper()}"
    if color.type == "theme" and workbook is not None:
        palette = _theme_palette(workbook)
        if isinstance(color.theme, int) and 0 <= color.theme < len(palette):
            return f"FF{_apply_tint(palette[color.theme], color.tint or 0)}"
    return None


KEY_HEADER = "Параметр"
VALUE_HEADER = "Значення"
LABEL_HEADER = "Назва"
NOTE_HEADER = "Пояснення"


def read_settings_sheet(workbook, sheet_name, settings, *, section=None, legacy_sheet=None):
    """{key: значення} таблиці параметрів за списком settings. Невідомі
    параметри ігноруються; відсутній чи порожній обов'язковий - помилка;
    відсутній необов'язковий - як порожня клітинка."""
    label, raw_values = _key_values(workbook, sheet_name, section, legacy_sheet)
    result = {}
    for setting in settings:
        row_number, raw = raw_values.get(setting.key, (None, None))
        where = location(label, row_number, VALUE_HEADER) if row_number else location(label)
        named = f'"{setting.label}" ({setting.key})' if setting.label else f'"{setting.key}"'
        if setting.required and row_number is None:
            raise SettingsError(f"{location(label)}: немає рядка з параметром {named}.")
        value = parse_value(raw, setting.kind, f"{where} ({setting.key})")
        if setting.required and value in (None, "", []):
            hint = f" Що вписати: {setting.note}" if setting.note else ""
            raise SettingsError(f"{where}: параметр {named} не може бути порожнім.{hint}")
        result[setting.key] = value
    return result


def _key_values(workbook, sheet_name, section=None, legacy_sheet=None):
    """(підпис для повідомлень, {параметр: (номер рядка, сире значення)}) -
    усі рядки таблиці параметрів (колонки "Параметр" і "Значення"; повторений
    параметр - помилка)."""
    sheet, label, first_row, last_row = _require_table(workbook, sheet_name, section, legacy_sheet)
    key_column = Column("key", KEY_HEADER)
    value_column = Column("value", VALUE_HEADER)
    header_row, positions = _find_header(sheet, label, (key_column, value_column), first_row, last_row)
    result = {}
    rows = list(sheet.iter_rows(min_row=header_row + 1, max_row=last_row, values_only=True))
    for row_number, values in enumerate(rows, start=header_row + 1):
        key = cell_text(values[positions["key"]] if positions["key"] < len(values) else None)
        raw = values[positions["value"]] if positions["value"] < len(values) else None
        if not key:
            if not is_blank(raw):
                raise SettingsError(f"{location(label, row_number)}: значення без назви параметра (колонка \"{KEY_HEADER}\").")
            continue
        if key in result:
            raise SettingsError(f'{location(label, row_number)}: параметр "{key}" уже є в рядку {result[key][0]}.')
        result[key] = (row_number, raw)
    return label, result


def read_key_values(workbook, sheet_name, *, section=None, legacy_sheet=None):
    """{параметр: (номер рядка, сире значення)} таблиці параметрів."""
    return _key_values(workbook, sheet_name, section, legacy_sheet)[1]


def read_texts(workbook, sheet_name, *, section=None, legacy_sheet=None):
    """{параметр: текст} - усі параметри таблиці текстів (тексти документів:
    набір ключів довільний, перевіряє вже сам код, що їх використовує)."""
    raw_values = read_key_values(workbook, sheet_name, section=section, legacy_sheet=legacy_sheet)
    return {key: cell_text(raw) for key, (_row, raw) in raw_values.items()}


# ---------------------------------------------------------------- запис


def new_workbook():
    workbook = Workbook()
    workbook.remove(workbook.active)
    return workbook


def _set_cell(cell, value, kind):
    if value is None or value == [] or value == "" or kind == COLOR:
        return
    if kind == BOOL:
        cell.value = YES if value else NO
    elif kind == LINES:
        _set_text(cell, "\n".join(value) if isinstance(value, (list, tuple)) else str(value))
    elif kind == LIST:
        _set_text(cell, ", ".join(str(item) for item in value) if isinstance(value, (list, tuple)) else str(value))
    elif kind == DATE:
        try:
            cell.value = datetime.strptime(str(value), DATE_FORMAT)
            cell.number_format = EXCEL_DATE_FORMAT
        except ValueError:
            _set_text(cell, str(value))
    elif kind == DAY:
        cell.value = value
    elif kind == NUMBER_TEXT and isinstance(value, str) and value.isdigit() and str(int(value)) == value:
        cell.value = int(value)
    elif kind in (INT, NUMBER_TEXT) and not isinstance(value, str):
        cell.value = value
    else:
        _set_text(cell, str(value))


def _set_text(cell, text):
    cell.value = text
    # Текст, що починається з "=", openpyxl інакше записав би як формулу.
    if text.startswith("="):
        cell.data_type = "s"


def _new_sheet(workbook, sheet_name, tab_color, index=None):
    sheet = workbook.create_sheet(sheet_name, index)
    if tab_color:
        sheet.sheet_properties.tabColor = tab_color
    sheet.sheet_view.showGridLines = False
    sheet.sheet_view.zoomScale = 110
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    return sheet


def _text_height(text, width_chars, line_height=15):
    """Орієнтовна висота рядка для тексту з переносами в клітинці шириною
    width_chars символів (об'єднані клітинки Excel сам не підганяє)."""
    lines = sum(max(1, math.ceil(len(part) / max(width_chars, 10))) for part in str(text).split("\n"))
    return max(line_height + 4, lines * line_height + 6)


def _banner_columns(column_count, total_width):
    """Скільки колонок займає смуга назви й підказки: уся таблиця, а вузька -
    ще кілька порожніх колонок праворуч (до _MIN_BANNER_WIDTH символів)."""
    extra = max(0, math.ceil((_MIN_BANNER_WIDTH - total_width) / _DEFAULT_COLUMN_WIDTH))
    return max(column_count, 1) + extra, max(total_width, _MIN_BANNER_WIDTH)


def _merged_text_row(sheet, row, text, span, width, *, font, fill=None, height=None, line_height=14):
    if span > 1:
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    cell = sheet.cell(row=row, column=1)
    _set_text(cell, text)
    cell.font = font
    cell.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
    if fill is not None:
        for column in range(1, span + 1):
            sheet.cell(row=row, column=column).fill = fill
    sheet.row_dimensions[row].height = height or _text_height(text, width * 1.05, line_height=line_height)


def _write_banner(sheet, title, hint, column_count, total_width):
    """Рядки 1-2: назва аркуша (темна смуга) і підказка (курсив) на всю ширину таблиці."""
    span, width = _banner_columns(column_count, total_width)
    _merged_text_row(sheet, TITLE_ROW, title, span, width, font=_TITLE_FONT, fill=_TITLE_FILL, height=30)
    if hint:
        _merged_text_row(sheet, HINT_ROW, hint, span, width, font=_HINT_FONT)
    else:
        sheet.row_dimensions[HINT_ROW].height = 6


def _starts(columns):
    """Номер першої колонки аркуша кожної колонки таблиці (з урахуванням span)."""
    starts, position = [], 1
    for column in columns:
        starts.append(position)
        position += column.span
    return starts


def _total_span(columns):
    return sum(column.span for column in columns)


def _merge_span(sheet, row, start, column):
    if column.span > 1:
        sheet.merge_cells(start_row=row, start_column=start, end_row=row, end_column=start + column.span - 1)


def _write_groups(sheet, columns, row=GROUP_ROW):
    """Рядок кольорових груп колонок (сусідні колонки з тією самою group -
    одна об'єднана смуга). Без груп - вузький порожній рядок-відступ.
    Повертає True, якщо групи є."""
    if not any(column.group for column in columns):
        sheet.row_dimensions[row].height = 6
        return False
    starts = _starts(columns)
    colors, start = {}, 0
    while start < len(columns):
        group = columns[start].group
        end = start
        while end + 1 < len(columns) and columns[end + 1].group == group:
            end += 1
        if group:
            fill = PatternFill("solid", start_color=colors.setdefault(group, GROUP_COLORS[len(colors) % len(GROUP_COLORS)]))
            first_column, last_column = starts[start], starts[end] + columns[end].span - 1
            if last_column > first_column:
                sheet.merge_cells(start_row=row, start_column=first_column, end_row=row, end_column=last_column)
            cell = sheet.cell(row=row, column=first_column, value=group)
            cell.font = _GROUP_FONT
            cell.alignment = _HEADER_ALIGN
            for index in range(first_column, last_column + 1):
                sheet.cell(row=row, column=index).fill = fill
        start = end + 1
    sheet.row_dimensions[row].height = 22
    return True


def _write_header(sheet, columns, row=HEADER_ROW):
    for start, column in zip(_starts(columns), columns):
        _merge_span(sheet, row, start, column)
        for index in range(start, start + column.span):
            sheet.cell(row=row, column=index).fill = _HEADER_FILL
            sheet.cell(row=row, column=index).border = _HEADER_BORDER
        cell = sheet.cell(row=row, column=start, value=column.header)
        cell.font = _HEADER_FONT
        cell.alignment = _HEADER_ALIGN
        if column.note:
            comment = Comment(column.note, "НАЛАШТУВАННЯ")
            comment.width = 360
            comment.height = 160
            cell.comment = comment
    lines = max((math.ceil(len(column.header) * 1.1 / max(column.width - 1, 4)) for column in columns), default=1)
    sheet.row_dimensions[row].height = max(32, lines * 15 + 8)


def column_position(columns, key):
    """Номер колонки аркуша (1 - A), з якої починається колонка key таблиці
    columns (з урахуванням span попередніх колонок)."""
    return next(start for start, column in zip(_starts(columns), columns) if column.key == key)


def _set_widths(sheet, tables):
    """Ширини колонок аркуша для всіх його таблиць tables ([columns, ...]): у
    таблиць аркуша спільні колонки, тож ширина - найбільша з колонок, що
    займають одну колонку аркуша. Колонці на кілька колонок (span) бракує
    ширини - її додає остання з її колонок аркуша, ширини якої не задає жодна
    інша таблиця (а немає такої - просто остання)."""
    widths = {}
    for columns in tables:
        for start, column in zip(_starts(columns), columns):
            if column.span == 1:
                widths[start] = max(widths.get(start, 0), column.width)
    fixed = set(widths)
    for columns in tables:
        for start, column in zip(_starts(columns), columns):
            if column.span == 1:
                continue
            covered = range(start, start + column.span)
            missing = column.width - sum(widths.get(index, _DEFAULT_COLUMN_WIDTH) for index in covered)
            if missing > 0:
                free = [index for index in covered if index not in fixed]
                target = free[-1] if free else covered[-1]
                widths[target] = widths.get(target, _DEFAULT_COLUMN_WIDTH) + missing
    for index, width in widths.items():
        sheet.column_dimensions[get_column_letter(index)].width = width


def _list_formula(options):
    """Формула випадного списку або None, якщо варіанти туди не вміщаються
    (Excel приймає не більше 255 символів)."""
    formula = '"' + ",".join(str(option).replace('"', '""') for option in options) + '"'
    return formula if len(formula) <= 255 else None


def _add_validations(sheet, columns, first_row, last_row):
    """Випадні списки (BOOL - "Так/Ні", choices) і кольори "Так/Ні" для
    рядків first_row..last_row таблиці."""
    for index, column in zip(_starts(columns), columns):
        options = (YES, NO) if column.kind == BOOL else column.choices
        letter = get_column_letter(index)
        cells = f"{letter}{first_row}:{letter}{last_row}"
        formula = _list_formula(options) if options else None
        if formula is not None:
            validation = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=False)
            sheet.add_data_validation(validation)
            validation.add(cells)
        if column.kind == BOOL:
            # "Так" - зелена клітинка, "Ні" - сіра (умовне форматування: колір
            # міняється сам, коли користувач обирає інше значення).
            sheet.conditional_formatting.add(
                cells, CellIsRule(operator="equal", formula=[f'"{YES}"'], fill=_YES_FILL, font=Font(bold=True, color="2E6B24")),
            )
            sheet.conditional_formatting.add(
                cells, CellIsRule(operator="equal", formula=[f'"{NO}"'], fill=_NO_FILL, font=Font(color="8C8C8C")),
            )


def _cell_alignment(column):
    if column.center or column.kind in (BOOL, DATE, DAY, INT, COLOR):
        return _CENTER
    if column.wrap or column.kind in (LINES, LIST):
        return _WRAP
    # Номер чи пункт ліворуч, як і текст: 30 і NOT_PAID в одній колонці.
    return _LEFT if column.kind == NUMBER_TEXT else _TOP


_NO_BAND = object()


def _write_rows(sheet, first_row, columns, rows, band_by=None, empty_rows=_EMPTY_ROWS):
    """Рядки rows ({key: значення}) з first_row і empty_rows порожніх рядків з
    рамками під ними (видно, куди дописувати). Повертає номер останнього рядка.

    Оформлення рядка: "_style": "point" - рядок-заголовок блоку (помаранчевий,
    жирний), "muted" - приглушений; "_dim": ключі колонок, текст яких сірий
    (напр. повтор значення з рядка-заголовка). Чергування заливки: кожен
    рядок, або (band_by - ключ колонки) щоразу, коли змінюється значення цієї
    колонки (напр. група підстав)."""
    row_index, stripe, previous_band = first_row, True, _NO_BAND
    for row in rows:
        style = row.get("_style")
        dim = set(row.get("_dim", ()))
        if band_by is None:
            stripe = not stripe
        else:
            band_value = normalize_name(row.get(band_by) or "")
            if band_value != previous_band:
                stripe = previous_band is not _NO_BAND and not stripe
                previous_band = band_value
        height = None
        for start, column in zip(_starts(columns), columns):
            _merge_span(sheet, row_index, start, column)
            cell = sheet.cell(row=row_index, column=start)
            _set_cell(cell, row.get(column.key), column.kind)
            cell.alignment = _cell_alignment(column)
            if column.span > 1 and cell.value is not None and (column.wrap or column.kind in (LINES, LIST)):
                # Об'єднані клітинки Excel сам за висотою тексту не підганяє.
                height = max(height or 0, _text_height(cell_text(cell.value), column.width * 1.05))
            for index in range(start, start + column.span):
                part = sheet.cell(row=row_index, column=index)
                if style == "point":
                    part.fill, part.font, part.border = _POINT_FILL, _POINT_FONT, _POINT_BORDER
                    continue
                part.border = _BORDER
                if style == "muted":
                    part.fill, part.font = _MUTED_FILL, _MUTED_FONT
                elif stripe:
                    part.fill = _STRIPE_FILL
                if column.key in dim:
                    part.font = _DIM_FONT
            if column.kind == COLOR and style != "point":
                color = row.get(column.key)
                cell.fill = PatternFill("solid", start_color=color, end_color=color) if color else PatternFill(fill_type=None)
        if height:
            sheet.row_dimensions[row_index].height = height
        row_index += 1
    for _ in range(empty_rows):
        for start, column in zip(_starts(columns), columns):
            _merge_span(sheet, row_index, start, column)
            for index in range(start, start + column.span):
                sheet.cell(row=row_index, column=index).border = _BORDER
        row_index += 1
    return row_index - 1


def write_table(workbook, sheet_name, columns, rows, *, title=None, hint="", tab_color=None, band_by=None,
                frozen_columns=0):
    """Новий аркуш з однією таблицею: назва (title, типово - назва аркуша) й
    підказка над таблицею, кольорові групи колонок, заголовки (пояснення
    колонки - у примітці до заголовка) і рядки rows (оформлення - _write_rows).
    BOOL-колонки - випадний список "Так/Ні" і колір значення. Повертає аркуш."""
    sheet = _new_sheet(workbook, sheet_name, tab_color)
    _write_banner(sheet, title or sheet_name, hint, len(columns), sum(column.width for column in columns))
    _write_groups(sheet, columns)
    _write_header(sheet, columns)
    _set_widths(sheet, [columns])
    sheet.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
    _write_rows(sheet, FIRST_DATA_ROW, columns, rows, band_by)
    _add_validations(sheet, columns, FIRST_DATA_ROW, FIRST_DATA_ROW + _VALIDATION_ROWS)
    sheet.freeze_panes = sheet.cell(row=FIRST_DATA_ROW, column=frozen_columns + 1)
    return sheet


def _settings_columns(all_settings, label_width, value_width, note_width, wrap_values, spans=(1, 1, 1, 1)):
    # Технічні назви - дрібним шрифтом, тож на символ треба менше ширини.
    key_width = max(18, min(40, math.ceil(max((len(setting.key) for setting in all_settings), default=0) * 0.95) + 2))
    label_span, value_span, note_span, key_span = spans
    return (
        Column("label", LABEL_HEADER, width=label_width, span=label_span),
        Column("value", VALUE_HEADER, width=value_width, wrap=wrap_values, span=value_span),
        Column("note", NOTE_HEADER, width=note_width, wrap=True, span=note_span),
        Column("key", KEY_HEADER, width=key_width, span=key_span),
    )


def _all_settings(settings, values):
    """settings і параметри з values, яких у settings немає (тексти документів з
    довільним набором ключів), - наприкінці."""
    known = {setting.key for setting in settings}
    return (*settings, *(Setting(key, TEXT, False) for key in values if key not in known))


def _write_settings_rows(sheet, first_row, all_settings, values, wrap_values, columns=None):
    """Рядки параметрів "Назва | Значення | Пояснення | Параметр". Клітинки
    "Значення" підсвічені - саме їх редагують; "Параметр" - незмінна технічна
    назва (за нею програма шукає параметр). Повертає номер останнього рядка."""
    columns = columns or _settings_columns(all_settings, 34, 45, 58, wrap_values)
    starts = dict(zip((column.key for column in columns), _starts(columns)))
    row_index = first_row - 1
    for row_index, setting in enumerate(all_settings, start=first_row):
        for start, column in zip(_starts(columns), columns):
            _merge_span(sheet, row_index, start, column)
            for index in range(start, start + column.span):
                sheet.cell(row=row_index, column=index).border = _BORDER
        label_cell = sheet.cell(row=row_index, column=starts["label"], value=setting.label or setting.key)
        label_cell.font = _BOLD
        label_cell.alignment = _WRAP
        value_cell = sheet.cell(row=row_index, column=starts["value"])
        _set_cell(value_cell, values.get(setting.key), setting.kind)
        value_cell.alignment = _WRAP if wrap_values or setting.kind in (LINES, LIST) else _TOP
        if setting.kind == INT:
            value_cell.alignment = Alignment(horizontal="left", vertical="top")
        note_cell = sheet.cell(row=row_index, column=starts["note"])
        if setting.note:
            _set_text(note_cell, setting.note)
        note_cell.font = _NOTE_FONT
        note_cell.alignment = _WRAP
        key_cell = sheet.cell(row=row_index, column=starts["key"], value=setting.key)
        key_cell.font = _KEY_FONT
        key_cell.alignment = _TOP
        for key, fill in (("label", _LABEL_FILL), ("value", _EDIT_FILL)):
            column = next(column for column in columns if column.key == key)
            for index in range(starts[key], starts[key] + column.span):
                sheet.cell(row=row_index, column=index).fill = fill
        if any(column.span > 1 for column in columns):
            # Об'єднані клітинки Excel сам за висотою тексту не підганяє.
            texts = ((cell_text(values.get(setting.key)), columns[1].width), (setting.note or "", columns[2].width * 1.25),
                     (setting.label or setting.key, columns[0].width))
            sheet.row_dimensions[row_index].height = max(_text_height(text, width * 1.05) for text, width in texts)
    return row_index


def write_settings_sheet(workbook, sheet_name, settings, values, *, title=None, hint="", tab_color=None,
                         label_width=34, value_width=45, note_width=58, wrap_values=False):
    """Новий аркуш з однією таблицею параметрів "Назва | Значення | Пояснення |
    Параметр": по рядку на setting (і на параметри з values поза settings)."""
    all_settings = _all_settings(settings, values)
    columns = _settings_columns(all_settings, label_width, value_width, note_width, wrap_values)
    sheet = _new_sheet(workbook, sheet_name, tab_color)
    _write_banner(sheet, title or sheet_name, hint, len(columns), sum(column.width for column in columns))
    _write_groups(sheet, columns)
    _write_header(sheet, columns)
    _set_widths(sheet, [columns])
    _write_settings_rows(sheet, FIRST_DATA_ROW, all_settings, values, wrap_values, columns)
    sheet.freeze_panes = sheet.cell(row=FIRST_DATA_ROW, column=1)
    return sheet


@dataclass
class Section:
    """Таблиця на аркуші з кількома таблицями (write_sections).

    Звичайна таблиця - columns і rows (band_by - як у _write_rows); таблиця
    параметрів "Назва | Значення | Пояснення | Параметр" - settings і values
    (widths - ширини назви, значення й пояснення; spans - скільки колонок
    аркуша займає назва, значення, пояснення й параметр). paint(sheet, перший
    рядок даних, рядки) - додаткове оформлення проєкту (напр. колір у клітинці)."""
    name: str
    hint: str = ""
    columns: tuple = ()
    rows: list = ()
    band_by: str = None
    settings: tuple = None
    values: dict = None
    widths: tuple = (34, 45, 58)
    spans: tuple = (1, 1, 1, 1)
    wrap_values: bool = False
    paint: object = None


# Порожніх рядків з рамками під кожною таблицею аркуша з кількома таблицями -
# місце для нових записів (бракує - вставляють рядок усередині таблиці).
_SECTION_EMPTY_ROWS = 5


def write_sections(workbook, sheet_name, sections, *, title, hint="", tab_color=None, frozen_columns=0):
    """Новий аркуш з кількома таблицями одна під одною: зверху назва й
    підказка аркуша, далі кожна таблиця - кольорова смуга "■ назва" (за нею
    програма таблицю й знаходить, locate_table), підказка, заголовки, рядки й
    кілька порожніх рядків. Рядок смуги кожної таблиці запам'ятовується в
    книзі - на нього ведуть посилання аркуша "Інструкція". frozen_columns -
    скільки колонок зліва лишаються на місці, коли широкий аркуш гортають
    праворуч. Повертає аркуш."""
    sheet = _new_sheet(workbook, sheet_name, tab_color)
    band_fill = PatternFill("solid", start_color=tab_color or TAB_INFO)
    prepared = []
    for section in sections:
        if section.settings is not None:
            all_settings = _all_settings(section.settings, section.values or {})
            columns = _settings_columns(all_settings, *section.widths, section.wrap_values, section.spans)
            prepared.append((section, columns, all_settings))
        else:
            prepared.append((section, tuple(section.columns), None))
    widest = max((sum(column.width for column in columns) for _section, columns, _settings in prepared), default=0)
    most = max((_total_span(columns) for _section, columns, _settings in prepared), default=1)
    _write_banner(sheet, title, hint, most, widest)
    # Смуги й підказки всіх таблиць аркуша - однакової ширини, як і назва аркуша.
    span, width = _banner_columns(most, widest)
    sheet.row_dimensions[GROUP_ROW].height = 8
    _set_widths(sheet, [columns for _section, columns, _settings in prepared])
    registry = getattr(workbook, "_settings_sections", None)
    if registry is None:
        registry = {}
        workbook._settings_sections = registry
    row = GROUP_ROW + 1
    for section, columns, all_settings in prepared:
        _merged_text_row(sheet, row, section_band_text(section.name), span, width,
                         font=Font(bold=True, size=12, color="FFFFFF"), fill=band_fill, height=24)
        registry[(normalize_name(sheet_name), normalize_name(section.name))] = row
        row += 1
        if section.hint:
            _merged_text_row(sheet, row, section.hint, span, width, font=_HINT_FONT)
            row += 1
        if _write_groups(sheet, columns, row):
            row += 1
        _write_header(sheet, columns, row)
        first_data = row + 1
        if all_settings is not None:
            last = _write_settings_rows(sheet, first_data, all_settings, section.values or {}, section.wrap_values, columns)
        else:
            last = _write_rows(sheet, first_data, columns, section.rows, section.band_by, _SECTION_EMPTY_ROWS)
            _add_validations(sheet, columns, first_data, last)
            if section.paint is not None:
                section.paint(sheet, first_data, list(section.rows))
        row = last + 2
        sheet.row_dimensions[last + 1].height = 14
    if frozen_columns:
        sheet.freeze_panes = f"{get_column_letter(frozen_columns + 1)}1"
    return sheet


INSTRUCTIONS_SHEET = "Інструкція"
# Заголовок блоку "що де змінити" на аркуші "Інструкція".
TASKS_HEADING = "Що потрібно зробити - де це змінити"


def _link_target(workbook, sheet_name, section):
    row = getattr(workbook, "_settings_sections", {}).get((normalize_name(sheet_name), normalize_name(section or "")), 1)
    return f"'{sheet_name}'!A{row}"


def _heading_row(sheet, row_index, heading, color, width_count):
    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=width_count)
    heading_cell = sheet.cell(row=row_index, column=1, value=heading)
    heading_cell.font = _GROUP_FONT
    heading_cell.alignment = Alignment(vertical="center", indent=1)
    for column in range(1, width_count + 1):
        sheet.cell(row=row_index, column=column).fill = PatternFill("solid", start_color=color)
    sheet.row_dimensions[row_index].height = 22


def write_instructions(workbook, title, intro, tasks=(), contents=(), tips=()):
    """Перший аркуш "Інструкція" (код його не читає): вступ; "Що потрібно
    зробити - де це змінити" - tasks = [(що зробити, [(аркуш, таблиця чи None),
    ...]), ...]; "Що де лежить" - contents = [(аркуш, колір, [(таблиця чи None,
    що там), ...]), ...]; поради наприкінці. Назва таблиці - посилання прямо на
    її смугу на аркуші."""
    sheet = _new_sheet(workbook, INSTRUCTIONS_SHEET, TAB_INFO, index=0)
    widths = (34, 96)
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    _write_banner(sheet, title, intro, len(widths), sum(widths))
    row_index = HEADER_ROW

    def item_row(places, description):
        nonlocal row_index
        link = sheet.cell(row=row_index, column=1)
        _set_text(link, " / ".join(f"{sheet_name} > {section}" if section else sheet_name for sheet_name, section in places))
        target_sheet, target_section = places[0]
        link.hyperlink = Hyperlink(ref=link.coordinate, location=_link_target(workbook, target_sheet, target_section),
                                   display=link.value)
        link.font = _LINK_FONT
        link.alignment = _WRAP
        text = sheet.cell(row=row_index, column=2)
        _set_text(text, description)
        text.alignment = _WRAP
        for column in range(1, len(widths) + 1):
            sheet.cell(row=row_index, column=column).border = _BORDER
        sheet.row_dimensions[row_index].height = max(
            _text_height(description, widths[1] * 0.9), _text_height(link.value, widths[0] * 0.85),
        )
        row_index += 1

    if tasks:
        _heading_row(sheet, row_index, TASKS_HEADING, TAB_INFO, len(widths))
        row_index += 1
        for description, places in tasks:
            item_row(list(places), description)
        row_index += 1
    for sheet_name, color, items in contents:
        _heading_row(sheet, row_index, f"Аркуш «{sheet_name}»", color, len(widths))
        row_index += 1
        for section, description in items:
            item_row([(sheet_name, section)], description)
        row_index += 1
    for tip in tips:
        sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=len(widths))
        tip_cell = sheet.cell(row=row_index, column=1)
        _set_text(tip_cell, tip)
        tip_cell.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
        tip_cell.font = _HINT_FONT
        sheet.row_dimensions[row_index].height = _text_height(tip, sum(widths) * 1.0, line_height=14)
        row_index += 1
    workbook.active = 0
    for other in workbook.worksheets[1:]:
        other.sheet_view.tabSelected = False
    return sheet


def save_workbook(workbook, path):
    """Зберігає книгу через тимчасовий файл у тій самій теці й os.replace:
    якщо запис зірветься, наявний файл лишиться цілим. Файл відкритий в Excel
    (Windows не дає його замінити) - SettingsError з проханням закрити його."""
    directory = os.path.dirname(os.path.abspath(path))
    os.makedirs(directory, exist_ok=True)
    handle, temp_path = tempfile.mkstemp(prefix=".~", suffix=".xlsx", dir=directory)
    os.close(handle)
    try:
        workbook.save(temp_path)
        os.replace(temp_path, path)
    except PermissionError as error:
        raise SettingsError(
            f"Не вдалося записати {path}: файл відкритий в Excel чи іншій програмі. Закрийте його й спробуйте ще раз."
        ) from error
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def free_backup_path(path, label="до оновлення"):
    """Шлях для копії книги path перед тим, як програма її перепише:
    "НАЛАШТУВАННЯ (до оновлення).xlsx", а якщо така копія вже є (від
    попереднього оновлення) - "НАЛАШТУВАННЯ (до оновлення 2).xlsx" і т. д.:
    наявну копію ніколи не перезаписуємо."""
    stem, extension = os.path.splitext(path)
    candidate, number = f"{stem} ({label}){extension}", 1
    while os.path.exists(candidate):
        number += 1
        candidate = f"{stem} ({label} {number}){extension}"
    return candidate


def load_editable(path):
    """Книга для зміни й повторного збереження. Відкривається через власний
    дескриптор файлу: openpyxl, якому дали шлях, тримає файл відкритим до
    збирання сміття, і Windows тоді не дає save_workbook його замінити."""
    try:
        with open(path, "rb") as handle:
            return load_workbook(handle)
    except PermissionError as error:
        raise SettingsError(f"Не вдалося відкрити {path}: файл зайнятий іншою програмою.") from error


def update_table_cells(path, sheet_name, key_column, columns, updates):
    """Записує значення в наявні рядки аркуша-таблиці книги path, не чіпаючи
    решти книги (оформлення, примітки, інші аркуші).

    key_column - колонка, за значенням якої шукається рядок; updates -
    {значення key_column: {key колонки з columns: нове значення}}. Повертає
    множину ключів updates, для яких рядка не знайшлося (їх не записано)."""
    workbook = load_editable(path)
    sheet = _require_sheet(workbook, sheet_name)
    header_row, positions = _find_header(sheet, sheet_name, (key_column, *columns))
    kinds = {column.key: column.kind for column in columns}
    missing = set(updates)
    for row_number in range(header_row + 1, sheet.max_row + 1):
        raw_key = sheet.cell(row=row_number, column=positions[key_column.key] + 1).value
        if is_blank(raw_key):
            continue
        key = parse_value(raw_key, key_column.kind, location(sheet_name, row_number, key_column.header))
        if key not in updates:
            continue
        missing.discard(key)
        for field, value in updates[key].items():
            if field not in positions:
                continue
            cell = sheet.cell(row=row_number, column=positions[field] + 1)
            cell.value = None
            _set_cell(cell, value, kinds[field])
    save_workbook(workbook, path)
    return missing
