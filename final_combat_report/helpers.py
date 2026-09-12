import json, os, re, shutil, time, pandas as pd

from colorama import Fore, Style
from datetime import datetime, timedelta
from typing import Union, List
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_COLOR_INDEX
from docx.oxml.ns import qn

from constants import REGEX_REPLACEMENTS, REPLACEMENTS, TABLE_5_1_ROWS, ATTACHED_UNITS, OVT_CATEGORIES, UNIT_BATTALION


def get_signal_data(file_path):
    """Зчитує JSON і повертає оброблений список повідомлень."""
    if not os.path.exists(file_path):
        print(f"Файл {file_path} не знайдено.")
        return []

    try:
        data = []
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                if line.strip(): # ігноруємо порожні рядки
                    data.append(json.loads(line))
            return data
    except json.JSONDecodeError as e:
        print(f"Помилка JSON у файлі {file_path}: {e}")
        return []
    

def convert_to_short_name(full_name, id_position=None):
    if not isinstance(full_name, str):
        return full_name

    # 1. Виправляємо апостроф + пробіл -> зліплюємо назад
    cleaned = re.sub(r"(В|в)'[\s]+([а-яА-ЯІіЇїЄєҐґ])", r"\1'\2", full_name.strip())

    # 2. Прибираємо зайві пробіли
    cleaned = ' '.join(cleaned.split())

    # 3. Розбити на слова
    words = cleaned.split()

    # 4. Якщо менше ніж 3 — пробуємо розклеїти друге слово (ім’я+по батькові)
    if len(words) == 2:
        # Пробуємо вставити пробіл між ім’ям і по батькові
        possible_fix = re.sub(r'([а-яґєіїʼ’]+)([А-ЯІЇЄҐ])', r'\1 \2', words[1])
        fixed_words = [words[0]] + possible_fix.split()
        if len(fixed_words) == 3:
            words = fixed_words

    # 5. Перевірка
    if len(words) != 3:
        raise ValueError(f"❌ Некоректний ПІБ для {id_position}: \"{full_name}\" → після обробки: \"{' '.join(words)}\"")

    # 6. Повертаємо скорочене ім’я
    return f"{words[1]} {words[0]}"

def get_signature_officer(rows_with_data, posada):
    """Знаходить у СПИСКУ (rows_with_data) офіцера за посадою і повертає
    (звання, скорочене ПІБ) - для підпису в кінці звіту. Звання фактичне
    пріоритетне над штатним (людина підписується тим званням, яке має зараз)."""
    row = next((r for r in rows_with_data if r.get('Посада') == posada), None)
    if not row:
        return "", ""
    # реальні заголовки СПИСКУ мають несумісний регістр між собою: "Звання
    # фактичне" (з великої) і "звання за штатом" (з малої) - саме так, як є
    rank = next(
        (v.strip() for v in (row.get('Звання фактичне'), row.get('звання за штатом')) if isinstance(v, str) and v.strip()),
        ""
    )
    pib = row.get('П.І.Б.')
    name = convert_to_short_name(pib, posada) if isinstance(pib, str) else ""
    return rank, name

def set_margins(doc, top, bottom, left, right):
    sections = doc.sections
    for section in sections:
        section.page_height = Inches(11.69)  # 297 мм
        section.page_width = Inches(8.27)    # 210 мм
        section.top_margin = top
        section.bottom_margin = bottom
        section.left_margin = left
        section.right_margin = right

def set_line_spacing(doc, spacing):
    for paragraph in doc.paragraphs:
        paragraph.paragraph_format.line_spacing = Pt(spacing)

def add_text_for_header_and_footer_document(doc):
    section = doc.sections[0]

    # 🔧 Увімкнути різний колонтитул для першої сторінки
    section.different_first_page_header_footer = True

    # --- Верхній колонтитул (з’явиться ТІЛЬКИ починаючи з другої сторінки)
    header = section.header
    header_para = header.paragraphs[0]
    header_para.text = "ДЛЯ СЛУЖБОВОГО КОРИСТУВАННЯ"
    header_para.alignment = 1  # по центру
    header_para.style.font.size = Pt(14)
    header_para.style.font.name = "Times New Roman"

    # --- Нижній колонтитул (є на всіх сторінках)
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.text = "ДЛЯ СЛУЖБОВОГО КОРИСТУВАННЯ"
    footer_para.alignment = 1  # по центру
    footer_para.style.font.size = Pt(14)
    footer_para.style.font.name = "Times New Roman"

    # --- Нижній колонтитул ДЛЯ ПЕРШОЇ сторінки
    first_footer = section.first_page_footer
    first_footer_paragraph = first_footer.paragraphs[0]
    first_footer_paragraph.text = "ДЛЯ СЛУЖБОВОГО КОРИСТУВАННЯ"
    first_footer_paragraph.alignment = 1  # по центру
    first_footer_paragraph.style.font.size = Pt(14)
    first_footer_paragraph.style.font.name = "Times New Roman"


def add_paragraph_with_style(doc, text, font_name='Times New Roman', font_size=14, bold: Union[bool, str, List[str]] = False, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent: float = 0, format_tabs=False):
    paragraph = doc.add_paragraph()

    # 1. Налаштування абзацу
    paragraph.alignment = alignment
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    if first_line_indent > 0:
        paragraph.paragraph_format.first_line_indent = Cm(first_line_indent)

    # 2. Логіка розбиття тексту на частини (звичайні та жирні)
    # Якщо bold це рядок, перетворюємо його на список для універсальності
    words_to_bold = [bold] if isinstance(bold, str) else (bold if isinstance(bold, list) else [])
    
    
    if words_to_bold:
        # Створюємо регулярний вираз для пошуку всіх потрібних слів
        pattern = f"({'|'.join(map(re.escape, words_to_bold))})"
        parts = re.split(pattern, text)
    else:
        parts = [text]

    # 3. Додавання тексту частинами (runs)
    for part in parts:
        if not part: continue
        run = paragraph.add_run(part)
        
        # Налаштування шрифту для кожної частини
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor(0, 0, 0)
        
        # Перевірка: чи має ця частина бути жирною
        if (isinstance(bold, bool) and bold) or (part in words_to_bold):
            run.bold = True

    # 4. Налаштування табуляції
    if format_tabs:
        section = doc.sections[0]
        width = Inches(section.page_width.inches - (section.left_margin.inches + section.right_margin.inches))
        paragraph.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)

    return paragraph

def add_custom_heading(doc, text, font_name='Times New Roman', heading=1, color_rgb=RGBColor(0, 0, 0), font_size=14, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent: float = 0):
    paragraph = doc.add_paragraph()
    paragraph.style = f"Heading {heading}"

    paragraph.alignment = alignment
    paragraph.paragraph_format.line_spacing = 1.0  # Одинарний міжрядковий
    paragraph.paragraph_format.space_before = Pt(0)  # Взагалі нуль перед абзацом
    paragraph.paragraph_format.space_after = Pt(0)  # Взагалі нуль після абзацу
    paragraph.paragraph_format.first_line_indent = Cm(first_line_indent)

    run = paragraph.add_run(text)
    run.bold = True
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.color.rgb = color_rgb
    run.font.size = Pt(font_size)

    return paragraph

def mark_no_auto_highlight(doc, paragraph):
    """Виключає абзац (чи paragraph комірки таблиці) із загального порівняння
    з учорашнім звітом (highlight_changes_from_previous_report) - бо підсвічування
    для нього вже явно вирішене в місці формування (add_paragraph_with_parts) або
    текст завжди статичний (0-заглушки без джерела реальних даних, заголовки).
    Зберігаємо сам елемент _p (oxml), а не paragraph-обгортку: python-docx
    створює нову Paragraph-обгортку щоразу при зверненні до doc.paragraphs.
    Важливо зберігати саме об'єкт _p (не id()!) - lxml перевикористовує проксі-
    обгортки лише поки на них є жива посилка; без неї id() швидко переприсвоюється
    іншому елементу, і set() на основі id() ловить хибні збіги."""
    if not hasattr(doc, '_no_auto_highlight'):
        doc._no_auto_highlight = set()
    doc._no_auto_highlight.add(paragraph._p)
    return paragraph

def _add_formatted_run(paragraph, text, highlight, font_name='Times New Roman', font_size=14, bold=False, color_rgb=RGBColor(0, 0, 0)):
    run = paragraph.add_run(text)
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(font_size)
    run.font.color.rgb = color_rgb
    if bold:
        run.bold = True
    if highlight:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW

def add_paragraph_with_parts(doc, parts, font_name='Times New Roman', font_size=14, bold=False, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent: float = 0):
    """Абзац, зібраний з частин (text, highlight) - для рядків, де лише конкретні
    токени (дата, номер донесення, лічильник) мають підсвічуватись жовтим, а решта
    шаблонного тексту - ні. Підсвічування прив'язане до токена в момент його
    формування, а не до факту відмінності всього абзацу від учорашнього звіту -
    тому абзац автоматично виключається з тієї загальної перевірки."""
    paragraph = doc.add_paragraph()
    paragraph.alignment = alignment
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    if first_line_indent > 0:
        paragraph.paragraph_format.first_line_indent = Cm(first_line_indent)

    for text, highlight in parts:
        if not text:
            continue
        _add_formatted_run(paragraph, text, highlight, font_name, font_size, bold)

    return mark_no_auto_highlight(doc, paragraph)

def add_custom_heading_with_parts(doc, parts, font_name='Times New Roman', heading=1, color_rgb=RGBColor(0, 0, 0), font_size=14, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent: float = 0):
    paragraph = doc.add_paragraph()
    paragraph.style = f"Heading {heading}"
    paragraph.alignment = alignment
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.first_line_indent = Cm(first_line_indent)

    for text, highlight in parts:
        if not text:
            continue
        _add_formatted_run(paragraph, text, highlight, font_name, font_size, bold=True, color_rgb=color_rgb)

    return mark_no_auto_highlight(doc, paragraph)

# таблиці розділу 5 (5.1, 5.2) - 10pt, на відміну від 14pt в решті документа
TABLE_FONT_SIZE = 10

def _table_cell_text(cell, text, bold=False, font_size=TABLE_FONT_SIZE, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    para = cell.paragraphs[0]
    para.alignment = align
    run = para.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(font_size)
    run.font.name = "Times New Roman"

def _table_cell_parts(cell, parts, bold=False, font_size=TABLE_FONT_SIZE, align=WD_ALIGN_PARAGRAPH.CENTER):
    """Як _table_cell_text, але з частин (text, highlight) - для заголовків
    таблиць 5.1/5.2, де підсвічувати треба лише дати звітного періоду."""
    cell.text = ""
    para = cell.paragraphs[0]
    para.alignment = align
    for text, highlight in parts:
        if not text:
            continue
        run = para.add_run(str(text))
        run.bold = bold
        run.font.size = Pt(font_size)
        run.font.name = "Times New Roman"
        if highlight:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return para

def _table_cell_valign(cell, val="center"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), val)
    tcPr.append(vAlign)

def _table_set_col_width(cell, width_dxa):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(width_dxa))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)

def _apply_table_col_widths(table, col_widths, default_width=900):
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            _table_set_col_width(cell, col_widths[i] if i < len(col_widths) else default_width)
            _table_cell_valign(cell, "center")

def paragraph_five_dot_one_table(doc, report_date_start="??", report_date_end="??", hour_of_report=20, units_data=None):
    """
    Генерує таблицю втрат о/с для розділу 5.1.

    Параметри:
        doc - Word документ
        report_date_start - дата початку звітного періоду, наприклад "08.06.2026"
        report_date_end   - дата кінця звітного періоду, наприклад "09.06.2026"
        units_data        - словник з даними втрат по підрозділам (ключ - назва
        підрозділу без пробілу, напр. TABLE_5_1_ROWS[0]):
            {
                "<підрозділ>": {
                    "безповоротні": 0,      # звітний період
                    "санітарні": 0,
                    "зниклі": 0,
                    "полонені": 0,
                    "безповоротні_заг": 2,  # з початку включення
                    "санітарні_заг": 0,
                    "зниклі_заг": 0,
                    "полонені_заг": 0,
                },
                ...
            }
        Якщо units_data=None — всі клітинки порожні.
    """
    if units_data is None:
        units_data = {}

    PERIOD_KEYS = ["безповоротні", "санітарні", "зниклі", "полонені"]
    TOTAL_KEYS  = ["безповоротні_заг", "санітарні_заг", "зниклі_заг", "полонені_заг"]
    NUM_COLS = 9  # підрозділ + 4 звітних + 4 загальних

    cell_text, cell_valign = _table_cell_text, _table_cell_valign

    def get_val(unit, key):
        return units_data.get(unit, {}).get(key) or ""

    # Ширини колонок в DXA (1 cm ≈ 567 DXA)
    COL_WIDTHS = [1700, 900, 900, 900, 900, 900, 900, 900, 900]

    # --- створення таблиці ---
    table = doc.add_table(rows=0, cols=NUM_COLS)
    table.style = "Table Grid"

    # --- Рядок 1: великі заголовки ---
    r0 = table.add_row()

    cell_text(r0.cells[0], "Підрозділ")
    mark_no_auto_highlight(doc, r0.cells[0].paragraphs[0])

    c_period = r0.cells[1]
    c_period.merge(r0.cells[4])
    _table_cell_parts(
        c_period,
        [
            (f"Втрати особового складу за звітний період з {hour_of_report}:00 ", False),
            (report_date_start, True),
            (f" по {hour_of_report}.00 ", False),
            (report_date_end, True),
        ],
        bold=True,
    )
    mark_no_auto_highlight(doc, c_period.paragraphs[0])

    c_total = r0.cells[5]
    c_total.merge(r0.cells[8])
    cell_text(
        c_total,
        'Втрати з початком включення в склад сил і засобів 9 АК УВ (с) “Схід”',
        bold=True,
    )
    mark_no_auto_highlight(doc, c_total.paragraphs[0])

    # --- Рядок 2: підзаголовки ---
    r1 = table.add_row()
    subheaders = [
        "Підрозділ", "Безповоротні", "Санітарні", "Зниклі безвісті", "Полон",
        "Безповоротні", "Санітарні", "Зниклі безвісті", "Полон",
    ]
    for i, h in enumerate(subheaders):
        cell_text(r1.cells[i], h, bold=True)
        cell_valign(r1.cells[i], "center")
        mark_no_auto_highlight(doc, r1.cells[i].paragraphs[0])

    # --- Рядки підрозділів ---
    for unit in TABLE_5_1_ROWS:
        r = table.add_row()
        cell_text(r.cells[0], unit, align=WD_ALIGN_PARAGRAPH.LEFT)
        for i, key in enumerate(PERIOD_KEYS + TOTAL_KEYS):
            cell_text(r.cells[i + 1], get_val(unit, key))

    # --- Придані підрозділи ---
    r_attached_header = table.add_row()
    c_attached = r_attached_header.cells[0]
    c_attached.merge(r_attached_header.cells[NUM_COLS - 1])
    cell_text(c_attached, "Придані підрозділи", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    mark_no_auto_highlight(doc, c_attached.paragraphs[0])

    for unit in ATTACHED_UNITS:
        r = table.add_row()
        cell_text(r.cells[0], unit, align=WD_ALIGN_PARAGRAPH.LEFT)
        for i, key in enumerate(PERIOD_KEYS + TOTAL_KEYS):
            cell_text(r.cells[i + 1], get_val(unit, key))

    # --- Підсумковий рядок ---
    all_units = TABLE_5_1_ROWS + ATTACHED_UNITS
    r_sum = table.add_row()
    cell_text(r_sum.cells[0], f"Всього в {UNIT_BATTALION}", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    for i, key in enumerate(PERIOD_KEYS + TOTAL_KEYS):
        total = sum(
            (units_data.get(u, {}).get(key) or 0)
            for u in all_units
        )
        cell_text(r_sum.cells[i + 1], str(total) if total else "", bold=True)

    _apply_table_col_widths(table, COL_WIDTHS)


def paragraph_five_dot_two_table(doc, report_date_start="??", report_date_end="??", hour_of_report=19, ovt_data=None):
    """
    Генерує таблицю втрат ОВТ для розділу 5.2.

    ovt_data - словник з даними по категоріях ОВТ:
        {
            "ББМ": {
                "знищено": 0,       # звітний період
                "пошкоджено": 0,
                "всього": 0,
                "знищено_заг": 0,   # з початку включення
                "пошкоджено_заг": 6,
                "всього_заг": 6,
            },
            ...
            "Всього ОВТ за {UNIT_BATTALION}": {...}  # опційний рядок підсумку (переноситься з попереднього звіту)
        }
    Якщо ovt_data=None - всі клітинки порожні.
    """
    if ovt_data is None:
        ovt_data = {}

    PERIOD_KEYS = ["знищено", "пошкоджено", "всього"]
    TOTAL_KEYS = ["знищено_заг", "пошкоджено_заг", "всього_заг"]
    NUM_COLS = 7  # категорія + 3 звітних + 3 загальних

    cell_text, cell_valign = _table_cell_text, _table_cell_valign

    def get_val(category, key):
        return ovt_data.get(category, {}).get(key) or ""

    COL_WIDTHS = [2600, 900, 900, 900, 900, 900, 900]

    table = doc.add_table(rows=0, cols=NUM_COLS)
    table.style = "Table Grid"

    # --- Рядок 1: великі заголовки ---
    r0 = table.add_row()
    cell_text(r0.cells[0], "Підрозділ")
    mark_no_auto_highlight(doc, r0.cells[0].paragraphs[0])

    c_period = r0.cells[1]
    c_period.merge(r0.cells[3])
    _table_cell_parts(
        c_period,
        [
            (f"Втрати ОВТ за звітний період з {hour_of_report}:00 ", False),
            (report_date_start, True),
            (f" по {hour_of_report}.00 ", False),
            (report_date_end, True),
        ],
        bold=True,
    )
    mark_no_auto_highlight(doc, c_period.paragraphs[0])

    c_total = r0.cells[4]
    c_total.merge(r0.cells[6])
    cell_text(
        c_total,
        ' Втрати з початком включення в склад сил і засобів 9 АК УВ (с) “Схід”',
        bold=True,
    )
    mark_no_auto_highlight(doc, c_total.paragraphs[0])

    # --- Рядок 2: підзаголовки ---
    r1 = table.add_row()
    subheaders = ["Підрозділ", "Знищено", "Пошкоджено", "Всього", "Знищено", "Пошкоджено", "Всього"]
    for i, h in enumerate(subheaders):
        cell_text(r1.cells[i], h, bold=True)
        cell_valign(r1.cells[i], "center")
        mark_no_auto_highlight(doc, r1.cells[i].paragraphs[0])

    # --- Рядки категорій ОВТ ---
    for category in OVT_CATEGORIES:
        r = table.add_row()
        cell_text(r.cells[0], category, align=WD_ALIGN_PARAGRAPH.LEFT)
        for i, key in enumerate(PERIOD_KEYS + TOTAL_KEYS):
            cell_text(r.cells[i + 1], get_val(category, key))

    # --- Підсумковий рядок (переносимо з попереднього звіту, якщо є; інакше рахуємо суму) ---
    r_sum = table.add_row()
    ovt_summary_label = f"Всього ОВТ за {UNIT_BATTALION}"
    cell_text(r_sum.cells[0], ovt_summary_label, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    carried_total = ovt_data.get(ovt_summary_label)
    for i, key in enumerate(PERIOD_KEYS + TOTAL_KEYS):
        if carried_total is not None:
            value = carried_total.get(key) or ""
        else:
            total = sum((ovt_data.get(c, {}).get(key) or 0) for c in OVT_CATEGORIES)
            value = str(total) if total else ""
        cell_text(r_sum.cells[i + 1], value, bold=True)

    _apply_table_col_widths(table, COL_WIDTHS)

def close_open_output_documents(directory):
    """Закриває (без збереження) відкриті у Word документи з папки output,
    щоб shutil.rmtree/повторна генерація не впала через заблокований файл."""
    if not os.path.exists(directory):
        return

    abs_dir = os.path.abspath(directory)

    try:
        import win32com.client
        import pythoncom
    except ImportError:
        return

    try:
        pythoncom.CoInitialize()
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception:
        return

    try:
        for document in list(word.Documents):
            try:
                # normcase - Windows-шляхи регістронезалежні (C:\ vs c:\)
                doc_path = os.path.normcase(os.path.abspath(document.FullName))
                if doc_path.startswith(os.path.normcase(abs_dir)):
                    document.Close(SaveChanges=False)
                    print_green(f"Закрито відкритий файл: {document.Name}")
            except Exception:
                continue
    finally:
        pythoncom.CoUninitialize()


def create_or_clear_output_directory(directory):
    close_open_output_documents(directory)

    if os.path.exists(directory):
        # закриття документа через COM іноді реєструється із затримкою -
        # даємо декілька спроб, перш ніж здатися
        attempts = 5
        for attempt in range(1, attempts + 1):
            try:
                shutil.rmtree(directory)
                break
            except PermissionError:
                if attempt == attempts:
                    raise
                close_open_output_documents(directory)
                time.sleep(0.5)

    os.makedirs(directory)


def excel_col_to_index(col):
    """Перетворює Excel-літеру (наприклад, 'A', 'Z', 'AA', 'AD') у індекс (0-відлік)."""
    col = col.upper()
    index = 0
    for c in col:
        index = index * 26 + (ord(c) - ord('A') + 1)
    return index - 1

def get_rows(transfer_data, columns_transfer, date_can_be_empty=False):
    if date_can_be_empty:
        return [
            {
                col: (date_to_str(val) if not pd.isna(val) else None)
                if isinstance(val := row[col], (pd.Timestamp, datetime)) else val
                for col in columns_transfer
            }
            for _, row in transfer_data[columns_transfer].iterrows()
        ]

    return [
        {
            col: date_to_str(row[col]) if isinstance(row[col], (pd.Timestamp, datetime)) else row[col]
            for col in columns_transfer
        }
        for _, row in transfer_data[columns_transfer].iterrows()
    ]

def date_to_str(date_str, action='-', days=0):
    return (to_date(date_str) + timedelta(days)).strftime('%d.%m.%Y') if action == '+' else (to_date(date_str) - timedelta(days)).strftime('%d.%m.%Y')

def to_date(x):
    if isinstance(x, datetime):
        return x.date()
    elif isinstance(x, str):
        for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
            try:
                return datetime.strptime(x, fmt).date()
            except ValueError:
                continue
    raise ValueError(f"Невідомий формат або тип дати: {x}")


def print_green(s):
    print(Fore.GREEN + s + Style.RESET_ALL)

def parse_any_date(date_val):
    if isinstance(date_val, datetime):
        return date_val
    
    date_val = str(date_val).strip()
    
    # Спроба розібрати різні формати
    formats = [
        "%H:%M %d.%m.%y",     # 07:25 21.04.26
        "%H:%M %d.%m.%Y",     # 17:15 21.04.2026
        "%Y-%m-%d %H:%M:%S",  # 2026-04-21 16:54:00
        "%d.%m.%Y",           # 21.04.2026
        "%d.%m.%y"            # 21.04.26
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(date_val, fmt)
        except ValueError:
            continue
            
    # Якщо нічого не підійшло, повертаємо мінімальну дату, щоб не було помилки TypeError
    return datetime.min

def merge_and_sort_reports(*lists):
    combined = [item for lst in lists for item in lst]
    
    # Приводимо все до datetime перед сортуванням
    for item in combined:
        # Перевіряємо ключ 'date', якщо його немає — 'Дата'
        key = 'date' if 'date' in item else 'Дата'
        item['date_obj'] = parse_any_date(item.get(key))

    # Сортуємо по новоствореному об'єкту
    combined.sort(key=lambda x: x['date_obj'])
    
    # Видаляємо тимчасовий об'єкт, щоб не засмічувати дані (опціонально)
    for item in combined:
        del item['date_obj']
        
    return combined

_CALLSIGN_QUOTE_RE = re.compile(r'(?<=[A-Za-zА-Яа-яІіЇїЄєҐґ])\s*"\s*(?=[A-Za-zА-Яа-яІіЇїЄєҐґ])')

def _fix_callsign_quote_spacing(text):
    """Позивні в лапках: перед ВІДКРИВАЮЧОЮ лапкою - пробіл, після - без пробілу
    ("ТЗ"Газда" -> ТЗ "Газда"). Закриваючу лапку не чіпаємо взагалі - лишаємо
    пробіл після неї як є (число+лапка типу 10" сюди й не потрапляє, бо перед
    нею вимагається літера, не цифра).

    Відкриваючу й закриваючу лапку розрізняємо за ПАРНІСТЮ входження в тексті
    (1-ше, 3-тє... - відкриваюча; 2-ге, 4-те... - закриваюча), а не за сусідніми
    символами: і відкриваюча, і закриваюча лапка в звичайному '"Ім'я"' межують
    з літерою з ОБОХ боків, тому розпізнавання за сусіднім символом плутало їх
    і псувало текст, коли за закриваючою лапкою одразу йшло продовження речення
    (напр. 'Каспер" працював' ставало 'Каспер "працював', ковтаючи пробіл)."""
    count = [0]

    def repl(m):
        count[0] += 1
        return ' "' if count[0] % 2 == 1 else m.group(0)

    return _CALLSIGN_QUOTE_RE.sub(repl, text)

def clean_text(text: str) -> str:
    if not isinstance(text, str):
        return text

    for old, new in REPLACEMENTS.items():
        text = text.replace(old, new)

    for pattern, new in REGEX_REPLACEMENTS:
        text = re.sub(pattern, new, text)

    text = _fix_callsign_quote_spacing(text)

    return text.strip()

def parse_date_key(item):
    date_str = item.get("date", "") or ""
    
    # прибираємо "None " на початку
    date_str = re.sub(r'^None\s*', '', date_str).strip()
    # додаємо пробіл якщо час і дата злипись: "05:5813.07.26" -> "05:58 13.07.26"
    date_str = re.sub(r'(\d{2}:\d{2})(\d{2}\.\d{2}\.)', r'\1 \2', date_str)
    # якщо немає часу — додаємо 00:00
    if re.match(r'^\d{2}\.\d{2}\.', date_str):
        date_str = "00:00 " + date_str

    for fmt in ("%H:%M %d.%m.%Y", "%H:%M %d.%m.%y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return datetime.min

def _highlight_paragraph_yellow(paragraph):
    for run in paragraph.runs:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW

def highlight_changes_from_previous_report(doc, previous_doc):
    """Підсвічує жовтим абзаци й комірки таблиць, яких не було в учорашньому звіті
    (previous_doc) - новий чи змінений текст. Порівняння точне (весь абзац/комірка
    цілком): якщо текст слово-в-слово збігається з учорашнім - підсвічування немає;
    якщо ні (новий абзац чи хоч якась відмінність) - підсвічується цілком, оскільки
    для звітних записів (події, обстріли тощо) немає надійного способу відрізнити
    "цей самий запис із правкою" від "новий запис, що просто схожий на старий за
    шаблоном" (обидва мають однакову службову обгортку тексту). Якщо previous_doc
    немає (першого звіту чи вчорашнього файлу немає в resources/) - нічого не робить.

    Абзаци й комірки, зареєстровані через mark_no_auto_highlight (статичні
    заголовки/заглушки без джерела реальних даних, або рядки з власним точковим
    підсвічуванням через add_paragraph_with_parts), ця перевірка пропускає -
    інакше щоденна зміна дати/номера в шаблонному рядку підсвітила б увесь
    рядок цілком."""
    if previous_doc is None:
        return

    exclude_elements = getattr(doc, '_no_auto_highlight', set())
    prev_texts = {p.text for p in previous_doc.paragraphs if p.text.strip()}
    for paragraph in doc.paragraphs:
        if paragraph._p in exclude_elements:
            continue
        if paragraph.text.strip() and paragraph.text not in prev_texts:
            _highlight_paragraph_yellow(paragraph)

    # таблиці мають фіксовану структуру - порівнюємо клітинки позиційно
    for t_idx, today_table in enumerate(doc.tables):
        if t_idx >= len(previous_doc.tables):
            continue
        prev_table = previous_doc.tables[t_idx]
        for r_idx, today_row in enumerate(today_table.rows):
            if r_idx >= len(prev_table.rows):
                continue
            prev_row = prev_table.rows[r_idx]
            for c_idx, today_cell in enumerate(today_row.cells):
                if c_idx >= len(prev_row.cells):
                    continue
                cell_paragraph = today_cell.paragraphs[0]
                if cell_paragraph._p in exclude_elements:
                    continue
                if today_cell.text.strip() and today_cell.text != prev_row.cells[c_idx].text:
                    _highlight_paragraph_yellow(cell_paragraph)