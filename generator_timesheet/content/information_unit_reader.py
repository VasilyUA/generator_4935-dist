import glob
import os
import re
from datetime import date, datetime
from zipfile import BadZipFile

import xlrd
from openpyxl import load_workbook
from openpyxl.drawing.text import Font as _DrawingFont
from xlrd import xldate

from constants import DEFAULT_STATUS, INFORMATION_UNIT_SHEET_NAME, PIB_COLUMN_NAME
from content.oblik_timesheet import normalize_name

# Реальний випадок (файл цієї теки за 04.09.2026): деякі файли мають
# вбудований малюнок/штамп, чий текстовий стиль записує ПОРОЖНЄ значення
# атрибута "panose" (замість валідного hex-рядка чи повної відсутності
# атрибута) - openpyxl.drawing.text.Font.panose (HexBinary, патерн
# "[0-9a-fA-F]+$", вимагає ХОЧА Б ОДИН hex-символ) підіймає ValueError
# ("Unable to read workbook: ... invalid XML") ЩЕ ДО того, як стають
# доступні дані аркушів ТАБЕЛЬ/БЧС - раніше (_safe_load_workbook нижче)
# такий файл лише ПІЙМАВСЯ й пропускався ЦІЛКОМ (жодних даних підрозділу не
# враховувалось), хоча самі малюнки нас не цікавлять узагалі. Послаблюється
# ТОЧКОВО - лише СКОМПІЛЬОВАНИЙ pattern УЖЕ СТВОРЕНОГО дескриптора
# (Font.panose): заміна самого КЛАСУ HexBinary.pattern не подіяла б, бо
# test_pattern компілюється ОДИН РАЗ при створенні дескриптора (імпорт
# openpyxl.drawing.text), задовго до виклику _load_workbook.
_DrawingFont.panose.pattern = "[0-9a-fA-F]*$"
_DrawingFont.panose.test_pattern = re.compile(_DrawingFont.panose.pattern, re.VERBOSE)

# .xls (старий бінарний формат) - openpyxl.load_workbook НЕ вміє його читати
# напряму (лише .xlsx/.xlsm), тож для нього окремо читається через xlrd
# (_load_workbook нижче) - реальний випадок: один файл цієї теки був саме
# .xls і мовчки пропускався (дані підрозділу взагалі не враховувались).
_XLS_EXTENSION = ".xls"
_SUPPORTED_EXTENSIONS = (".xlsx", ".xlsm", _XLS_EXTENSION)

# Реальні .xls файли цієї теки записані в кодуванні cp1251 (стандартна
# кодова сторінка кирилиці на Windows для старого бінарного формату) - без
# цього xlrd читає назви аркушів/текст як "тарабарщину" (неправильна
# кодова сторінка визначається автоматично невірно для цих конкретних
# файлів).
_XLS_ENCODING = "cp1251"


def _is_office_lock_file(path):
    """Тимчасовий lock-файл Excel/Office ("~$Назва.xlsx") - створюється,
    поки файл ВІДКРИТИЙ в Excel, має ТЕ САМЕ розширення (.xlsx), тож інакше
    потрапив би в _SUPPORTED_EXTENSIONS і спричинив PermissionError/крах
    читання (сам lock-файл - не справжня книга, лише маленька позначка) -
    реальний випадок: користувач тримає файл цієї теки відкритим в Excel
    під час генерації."""
    return os.path.basename(path).startswith("~$")


class _XlsCell:
    """Той самий інтерфейс, що й одна клітинка openpyxl (.value) - навколо
    вже СКОНВЕРТОВАНОГО (_xls_cell_value) значення xlrd."""

    __slots__ = ("value",)

    def __init__(self, value):
        self.value = value


def _xls_cell_value(sheet, book, row, col):
    """Значення однієї клітинки xlrd (.xls) - у ТОМУ САМОМУ вигляді, що й
    openpyxl (Python-типи, а не сирі дані xlrd): дата - datetime (xlrd сам
    по собі дає лише число серійної дати - _tabel_date_row_and_columns/
    _parse_cell_date очікують datetime.date, як від openpyxl), ціле число,
    записане без дробової частини, - int, а не float (xlrd завжди дає float
    для чисел, незалежно від того, як воно виглядало у файлі - інакше
    категорія виплати "100" показувалась би як "100.0"), порожня клітинка -
    None (не порожній рядок чи 0.0)."""
    cell = sheet.cell(row, col)
    if cell.ctype == xlrd.XL_CELL_DATE:
        return xldate.xldate_as_datetime(cell.value, book.datemode)
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        value = cell.value
        return int(value) if value == int(value) else value
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return bool(cell.value)
    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return None
    return cell.value or None


class _XlsSheetAdapter:
    """Мінімальна "маска" аркуша xlrd (0-індексований sheet.cell(row, col))
    під той самий інтерфейс, що й openpyxl Worksheet (1-індексований
    ws.cell(row=, column=).value, max_row, max_column, title) - щоб УСІ
    наявні helper-и (_pib_column, _tabel_date_row_and_columns, _bchs_columns
    тощо, написані під openpyxl) працювали для .xls файлів БЕЗ ЖОДНОЇ зміни."""

    def __init__(self, sheet, book):
        self._sheet = sheet
        self._book = book
        self.title = sheet.name
        self.max_row = sheet.nrows
        self.max_column = sheet.ncols

    def cell(self, row, column):
        r, c = row - 1, column - 1
        if r < 0 or c < 0 or r >= self.max_row or c >= self.max_column:
            return _XlsCell(None)
        return _XlsCell(_xls_cell_value(self._sheet, self._book, r, c))


class _XlsWorkbookAdapter:
    """Той самий інтерфейс, що й openpyxl Workbook (sheetnames, wb[title]) -
    саме стільки, скільки потребує _find_sheet нижче, більше нічого з
    "книги" не використовується."""

    def __init__(self, path):
        book = xlrd.open_workbook(path, encoding_override=_XLS_ENCODING)
        self._book = book
        self.sheetnames = book.sheet_names()

    def __getitem__(self, title):
        return _XlsSheetAdapter(self._book.sheet_by_name(title), self._book)


def _load_workbook(path):
    """openpyxl (data_only=True) для .xlsx/.xlsm, _XlsWorkbookAdapter (xlrd)
    для .xls - ОДНАКОВИЙ інтерфейс для обох (sheetnames, wb[title],
    ws.cell(row=, column=).value, ws.max_row/max_column/title), тож усі
    наявні helper-и й читачі (тут і в content/payment_mismatch_checker.py)
    працюють однаково незалежно від формату файлу."""
    if os.path.splitext(path)[1].lower() == _XLS_EXTENSION:
        return _XlsWorkbookAdapter(path)
    return load_workbook(path, data_only=True)


def _safe_load_workbook(path):
    """_load_workbook(path), АЛЕ толерує пошкоджений файл - реальний випадок:
    один файл цієї теки мав пошкоджений внутрішній XML (значення "panose"
    вбудованого шрифту в малюнку не відповідало власній перевірці формату
    openpyxl - ValueError "Unable to read workbook: ... invalid XML") - без
    цього ОДНА така книга зупиняла ВЕСЬ прогін (жоден файл цієї теки взагалі
    не встигав прочитатись, включно з рештою, цілком справних). Порожній
    "panose" (найпоширеніший реальний варіант ЦІЄЇ конкретної помилки) тепер
    ПОПЕРЕДЖУЄТЬСЯ ще на рівні імпорту модуля (див. патч Font.panose вище) -
    цей try/except лишається як ЗАПАСНИЙ варіант для БУДЬ-ЯКОЇ ІНШОЇ помилки
    розбору внутрішнього XML, яку так заздалегідь не передбачиш. ValueError -
    сам openpyxl явно згортає в нього ЛЮБУ помилку розбору XML під час
    читання аркушів; KeyError/BadZipFile - той самий клас "файл існує, має
    підтримуваний формат, але вміст не вдається розібрати як книгу" для
    пошкодженого архіву .xlsx/відсутньої частини всередині нього.

    Повертає (wb, None) при успіху, (None, причина) - інакше, той самий
    формат {"reason"}, що й read_tabel/bchs_sheet нижче - викликач додає її
    у skipped і читає ДАЛІ, а не падає."""
    try:
        return _load_workbook(path), None
    except (ValueError, KeyError, BadZipFile) as e:
        return None, f"Не вдалось прочитати файл {path}: {e}"

# Рядок з датами аркуша ТАБЕЛЬ - РІЗНИТЬСЯ між реальними файлами цієї теки:
# більшість мають груповий заголовок "ДАТА" в рядку 1 (одна клітинка) і самі
# дати в рядку 2 (дані - з рядка 3), але деякі файли мають дати ОДРАЗУ в
# рядку 1, одразу після ПІБ, без групового заголовка (дані - з рядка 2) -
# рядок з датами визначається ДИНАМІЧНО (_tabel_date_row), а не фіксується
# наперед, інакше другий варіант мовчки не читається взагалі.
_TABEL_DATE_ROW_CANDIDATES = (1, 2)

# Аркуш "БЧС" - ДРУГЕ (окрім "ТАБЕЛЬ") джерело даних у КОЖНОМУ файлі цієї теки -
# знімок ОДНОГО конкретного дня (а не грід із колонкою на кожен день, як
# "ТАБЕЛЬ"): рядок 1 - назва підрозділу (можливо, з датою в довільному
# форматі), рядок 2 - заголовки колонок ("ПРІЗВИЩЕ ім'я по батькові" замість
# "ПІБ", "БЧС" - статус, "Виплата додактової винагороди" - число категорії
# виплати), дані - з рядка 3. На РЕАЛЬНИХ даних "ТАБЕЛЬ" у частині файлів
# містить лише ЗАСТАРІЛІ дати (напр. увесь липень, замість потрібного серпня) -
# аркуш "БЧС" тоді лишається ЄДИНИМ джерелом актуальних даних на цю дату.
_BCHS_SHEET_NAME = "бчс"
_BCHS_HEADER_ROW = 2
_DATE_IN_TEXT_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{2,4})")


def _find_sheet(wb, name):
    return next((wb[title] for title in wb.sheetnames if title.strip().lower() == name), None)


# Реальний випадок: клітинка-дата аркуша ТАБЕЛЬ (рядок дат) була записана
# як текст "06.092026" - друкарська помилка, пропущена крапка МІЖ місяцем
# і роком (мало бути "06.09.2026") - жоден із двох форматів strptime нижче
# цього не розпізнає, тож УСЯ колонка цього дня мовчки не вважалась
# "колонкою-датою" взагалі (_date_columns) - навіть попри те, що самі
# значення виплат у ній є. Анкероване (^...$) - лише РІВНО ця форма (день,
# крапка, місяць, рік БЕЗ крапки) - не чіпає жоден із двох форматів вище.
_MISSING_DOT_DATE_RE = re.compile(r"^(?P<day>\d{2})\.(?P<month>\d{2})(?P<year>\d{4})$")


def _parse_cell_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        stripped = value.strip()
        for fmt in ("%d.%m.%Y", "%d.%m.%y"):
            try:
                return datetime.strptime(stripped, fmt).date()
            except ValueError:
                continue
        match = _MISSING_DOT_DATE_RE.match(stripped)
        if match:
            try:
                return date(int(match.group("year")), int(match.group("month")), int(match.group("day")))
            except ValueError:
                return None
    return None


def _pib_column(ws):
    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    return next((idx for idx, value in enumerate(header, start=1) if isinstance(value, str) and value.strip().upper() == PIB_COLUMN_NAME), None)


def _date_columns(ws, row):
    columns = ((c, _parse_cell_date(ws.cell(row=row, column=c).value)) for c in range(1, ws.max_column + 1))
    return [(c, parsed) for c, parsed in columns if parsed is not None]


def _tabel_date_row_and_columns(ws):
    """(date_row, date_cols) - пробує кожен рядок з _TABEL_DATE_ROW_CANDIDATES
    ПО ЧЕРЗІ (спершу 1, потім 2) і бере ПЕРШИЙ, де знайшлась хоча б одна
    дата. Рядок 1 у "старому" форматі (груповий заголовок "ДАТА", текст, а не
    дата) ніколи не дає жодної дати, тож для таких файлів природно
    підбирається рядок 2 - зворотної сумісності не порушує."""
    for row in _TABEL_DATE_ROW_CANDIDATES:
        date_cols = _date_columns(ws, row)
        if date_cols:
            return row, date_cols
    return None, []


def _search_date(text):
    """Шукає ПЕРШУ дату (DD.MM.YY чи DD.MM.YYYY) БУДЬ-ДЕ в тексті (назва файлу
    чи заголовок аркуша) - на відміну від _parse_cell_date, який очікує, що
    ВЕСЬ рядок - це ЛИШЕ дата."""
    match = _DATE_IN_TEXT_RE.search(text or "")
    if not match:
        return None
    day, month, year = match.groups()
    if len(year) == 2:
        year = f"20{year}"
    try:
        return date(int(year), int(month), int(day))
    except ValueError:
        return None


def _extract_bchs_date(path, ws):
    """Дата аркуша БЧС - спершу з ІМЕНІ ФАЙЛУ (як для щоденних рапортів,
    content/daily_report_reader.report_date_from_filename), запасний варіант -
    із заголовка (рядок 1) самого аркуша: деякі імена файлів не мають року
    (напр. "...08.08_1_3.xlsx"), деякі заголовки не мають дати взагалі -
    жодне з двох джерел не гарантовано, тож пробуються обидва."""
    return _search_date(os.path.basename(path)) or _search_date(str(ws.cell(row=1, column=1).value or ""))


def _bchs_columns(ws):
    """(pib_col, pidrozdil_col, posada_col, zvannya_col, status_col,
    payment_col) з рядка _BCHS_HEADER_ROW - жодна позиція НЕ фіксована
    (різниться між файлами), визначається за ТЕКСТОМ заголовка: ПІБ -
    substring "прізвищ" (тут заголовок - "ПРІЗВИЩЕ ім'я по батькові", а не
    просто "ПІБ"), Підрозділ/Посада/Звання - точний збіг (pidrozdil_col/
    posada_col/zvannya_col використовуються лише
    content/payment_mismatch_checker.py - _read_bchs_sheet нижче їх
    ігнорує), статус - точний "БЧС", виплата - substring "виплат" ("Виплата
    додактової винагороди" - саме так, з друкарською помилкою, у реальних
    файлах)."""
    pib_col = pidrozdil_col = posada_col = zvannya_col = status_col = payment_col = None
    for c in range(1, ws.max_column + 1):
        value = ws.cell(row=_BCHS_HEADER_ROW, column=c).value
        if not isinstance(value, str):
            continue
        lowered = value.strip().lower()
        if "прізвищ" in lowered:
            pib_col = c
        elif lowered == "підрозділ":
            pidrozdil_col = c
        elif lowered == "посада":
            posada_col = c
        elif lowered == "звання":
            zvannya_col = c
        elif lowered == "бчс":
            status_col = c
        elif "виплат" in lowered:
            payment_col = c
    return pib_col, pidrozdil_col, posada_col, zvannya_col, status_col, payment_col


def _clean_cell_value(value):
    """Значення клітинки ТАБЕЛЬ/БЧС - ГОТОВЕ для порівняння/збереження: (1)
    текстові значення - обрізані з країв (реальний випадок: зайвий пробіл у
    значенні на кшталт "БПШП " мовчки "губив" збіг з PAYMENT_STATUS_
    CONVERSIONS/PAYMENT_STATUS_OVERRIDES, content/oblik_timesheet.py, - БЕЗ
    жодного повідомлення про помилку, статус просто лишався незміненим); (2)
    рядок, що ПОВНІСТЮ складається з цифр (можливо, з "-" на початку) -
    ПЕРЕТВОРЮЄТЬСЯ на ЧИСЛО (реальний ризик: категорія виплати "100",
    збережена як ТЕКСТ - напр. клітинка відформатована як "Текст" чи
    скопійована з іншої системи - інакше сприймалась би apply_payment_values
    як "нечислове значення", хоча вона однозначно являє собою число). Інші
    типи (число, дата, None) - без змін."""
    if isinstance(value, str):
        value = value.strip()
        if value.lstrip("-").isdigit():
            return int(value)
    return value


def _read_tabel_sheet(wb, path, values):
    """Читає аркуш ТАБЕЛЬ (грід - колонка на кожен день) у values. Рядок з
    датами (і, відповідно, перший рядок даних одразу під ним) визначається
    ДИНАМІЧНО (_tabel_date_row_and_columns) - не всі реальні файли мають
    однаковий формат. Повертає None (успіх) або текст причини, чому не
    вдалось."""
    ws = _find_sheet(wb, INFORMATION_UNIT_SHEET_NAME)
    if ws is None:
        return f"У файлі {path} не знайдено аркуш ТАБЕЛЬ."

    pib_col = _pib_column(ws)
    date_row, date_cols = _tabel_date_row_and_columns(ws)
    if pib_col is None or not date_cols:
        return f"У файлі {path} (аркуш {ws.title}) не знайдено {'колонку ПІБ' if pib_col is None else 'жодної колонки-дати'}."

    for row in range(date_row + 1, ws.max_row + 1):
        pib_raw = ws.cell(row=row, column=pib_col).value
        if not pib_raw:
            continue
        normalized = normalize_name(pib_raw)
        for col, day in date_cols:
            value = _clean_cell_value(ws.cell(row=row, column=col).value)
            if value not in (None, ""):
                values.setdefault((normalized, day), []).append(value)

    return None


def _read_bchs_sheet(wb, path, values):
    """Читає аркуш БЧС (знімок ОДНОГО дня) у values - для кожної людини:
    статус "БЧС" ІНШИЙ за DEFAULT_STATUS ("РВЗ") записується як є (текст,
    той самий підхід, що й нечислове значення в ТАБЕЛЬ - apply_payment_values
    трактує це як "підрозділ каже щось ІНШЕ"), інакше записується число з
    колонки виплати (якщо є). Повертає None (успіх) або текст причини."""
    ws = _find_sheet(wb, _BCHS_SHEET_NAME)
    if ws is None:
        return f"У файлі {path} не знайдено аркуш БЧС."

    day = _extract_bchs_date(path, ws)
    if day is None:
        return f"У файлі {path} (аркуш {ws.title}) не вдалось визначити дату (ні з назви файлу, ні із заголовка аркуша)."

    pib_col, _pidrozdil_col, _posada_col, _zvannya_col, status_col, payment_col = _bchs_columns(ws)
    if pib_col is None:
        return f"У файлі {path} (аркуш {ws.title}) не знайдено колонку ПІБ."

    for row in range(_BCHS_HEADER_ROW + 1, ws.max_row + 1):
        pib_raw = ws.cell(row=row, column=pib_col).value
        if not pib_raw:
            continue
        normalized = normalize_name(pib_raw)

        status = _clean_cell_value(ws.cell(row=row, column=status_col).value if status_col else None)
        if isinstance(status, str) and status and status.upper() != DEFAULT_STATUS:
            values.setdefault((normalized, day), []).append(status)
            continue

        payment = _clean_cell_value(ws.cell(row=row, column=payment_col).value if payment_col else None)
        if payment not in (None, ""):
            values.setdefault((normalized, day), []).append(payment)

    return None


def read_payment_values(directory):
    """Читає всі файли directory - ДВА можливі джерела на файл, аркуші
    "ТАБЕЛЬ"/"Табель" (грід - колонка на кожен день) і "БЧС" (знімок ОДНОГО
    дня, дата - з імені файлу чи заголовка аркуша) - без урахування регістру
    назви аркуша. ТАБЕЛЬ - ПРІОРИТЕТНЕ джерело: якщо його вдалось прочитати,
    БЧС ТОГО САМОГО файлу взагалі НЕ читається (за прямою вказівкою
    користувача - обидва аркуші одного файлу можуть суперечити один одному,
    напр. друкарська помилка в ПІБ саме на одному з них, тож читання ОБОХ
    водночас лише множить розбіжності замість того, щоб мати ОДНЕ джерело
    правди). БЧС читається лише як ЗАПАСНИЙ варіант, коли ТАБЕЛЬ не вдалось
    прочитати взагалі. Повертає (values, skipped).

    values - {(normalize_name(ПІБ), date): [значення, ...]} - значення ЯК Є
    (число - категорія виплати; текст - статус на кшталт "ВП"/"ШП") з
    ПРІОРИТЕТНОГО джерела (аркуша) цього файлу, де знайдено цю людину й дату
    (список, а не одне значення - людина може траплятись У РІЗНИХ ФАЙЛАХ
    одразу, напр. якщо її "придано" іншому підрозділу - apply_payment_values
    (content/oblik_timesheet.py) вирішує, що робити із суперечливими
    значеннями).

    skipped - список {"reason"} для файлів, де НІ ТАБЕЛЬ, НІ БЧС не вдалось
    прочитати (непідтримуваний формат, пошкоджений вміст - _safe_load_workbook,
    жоден з двох аркушів/потрібних колонок не знайдено)."""
    values = {}
    skipped = []

    for path in sorted(glob.glob(os.path.join(directory, "*"))):
        if _is_office_lock_file(path):
            continue
        if os.path.splitext(path)[1].lower() not in _SUPPORTED_EXTENSIONS:
            skipped.append({"reason": f"Формат файлу не підтримується: {path}.", "report_date": None})
            continue

        wb, load_reason = _safe_load_workbook(path)
        if wb is None:
            skipped.append({"reason": load_reason, "report_date": None})
            continue
        tabel_reason = _read_tabel_sheet(wb, path, values)
        if tabel_reason is None:
            continue

        bchs_reason = _read_bchs_sheet(wb, path, values)
        if bchs_reason is not None:
            skipped.append({"reason": tabel_reason, "report_date": None})

    return values, skipped
