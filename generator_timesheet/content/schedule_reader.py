import glob
import os
import re
from datetime import timedelta

import content.information_unit_reader as iur
from content.daily_report_reader import report_date_from_filename
from content.oblik_timesheet import normalize_name

# Заголовок дня - рівно дві цифри текстом ("01".."31"): так пише і реальний
# файл, і content/rop_vop_statement.py::write_statement.
_DAY_HEADER_RE = re.compile(r"^\d{2}$")

# Деякі ПІБ у resources/schedule мають наприкінці дописку в дужках (назва
# іншого підрозділу, звідки людину прикріплено), якої немає в
# resources/ОБЛІК.xlsx. Без неї normalize_name дає той самий ключ, що й у
# Табелі. Текст у дужках зберігається: звіт розбіжностей
# (write_schedule_mismatch_report) показує його як "Підрозділ" для людей, яких
# не вдалося зіставити з Табелем.
_ATTACHMENT_SUFFIX_RE = re.compile(r"\s*\(([^)]*)\)\s*$")

# Скільки рядків від початку аркуша шукати заголовок: у реальному файлі він у
# рядку 6 (4 рядки титулу й порожній рядок), тож 20 - із запасом.
_HEADER_SEARCH_ROWS = 20


def latest_schedule_file(schedule_dir, year, month):
    """Шлях до найпізнішого файлу schedule_dir, чия фактична дата припадає на
    year/month, або None (тоді звіряти нема з чим). Lock-файли Excel ("~$...")
    і файли без дати в назві пропускаються.

    Дата в назві файлу (report_date_from_filename, як і в рапортів) - це день
    подання, наступного ранку, а не день, за який відомість. Напр. файл
    "...ВОП-РОП(СЕРПЕНЬ) 01.09.2026.xlsx" - відомість за 31.08.2026, подана
    вранці 01.09.2026. Тому фактична дата = дата з назви мінус один день."""
    candidates = []
    for path in sorted(glob.glob(os.path.join(schedule_dir, "*"))):
        if iur._is_office_lock_file(path):
            continue
        file_date = report_date_from_filename(path)
        if file_date is None:
            continue
        actual_date = file_date - timedelta(days=1)
        if actual_date.year == year and actual_date.month == month:
            candidates.append((actual_date, path))
    if not candidates:
        return None
    return max(candidates)[1]


def _header_row_and_columns(ws):
    """(header_row, pib_col, day_columns, posada_col, zvannya_col), де
    day_columns = {номер дня: колонка}. Колонки шукаються за текстом
    заголовка, а не за позицією: у реальному файлі перед видимими даними 8
    зайвих колонок. posada_col і zvannya_col можуть бути None. Повертає None,
    якщо в перших _HEADER_SEARCH_ROWS рядках немає рядка з колонкою ПІБ і хоча
    б однією колонкою-днем."""
    for row in range(1, min(ws.max_row, _HEADER_SEARCH_ROWS) + 1):
        pib_col = posada_col = zvannya_col = None
        day_columns = {}
        for c in range(1, ws.max_column + 1):
            value = ws.cell(row=row, column=c).value
            if not isinstance(value, str):
                continue
            stripped = value.strip()
            lowered = stripped.lower()
            if "прізвищ" in lowered:
                pib_col = c
            elif "посад" in lowered:
                posada_col = c
            elif "звання" in lowered:
                zvannya_col = c
            elif _DAY_HEADER_RE.match(stripped):
                day_columns[int(stripped)] = c
        if pib_col is not None and day_columns:
            return row, pib_col, day_columns, posada_col, zvannya_col
    return None


def read_schedule_marks(path):
    """Читає один файл resources/schedule. Повертає (marks, people):
    - marks - {(normalize_name(ПІБ), день): "роп"/"воп"/...} лише для
      заповнених клітинок. Порожня клітинка означає "ще не заповнено", а не
      "цього дня не на РОП/ВОП", тож не звіряється.
    - people - {normalize_name(ПІБ): {"pib_raw", "posada", "zvannya",
      "pidrozdil"}} - дані людини з цього файлу (posada/zvannya - None, якщо
      такої колонки немає; pidrozdil - текст дописки в дужках або None).
      Потрібні звіту розбіжностей (write_schedule_mismatch_report), щоб
      показати справжні дані людини з відомості, а не просто "Не знайдено в
      ОБЛІК.xlsx".

    Люди читаються від рядка під заголовком до першого порожнього ПІБ: одразу
    під ними, без порожнього рядка, іде підписний блок. Якщо заголовка
    (колонки ПІБ і хоча б одного дня) немає, обидва словники порожні."""
    marks = {}
    people = {}
    ws = iur._load_workbook(path).active
    header = _header_row_and_columns(ws) if ws is not None else None
    if ws is None or header is None:
        return marks, people
    header_row, pib_col, day_columns, posada_col, zvannya_col = header

    for row in range(header_row + 1, ws.max_row + 1):
        pib_raw = ws.cell(row=row, column=pib_col).value
        if not pib_raw:
            break

        suffix_match = _ATTACHMENT_SUFFIX_RE.search(pib_raw) if isinstance(pib_raw, str) else None
        pidrozdil = suffix_match.group(1).strip() if suffix_match else None
        cleaned_pib = _ATTACHMENT_SUFFIX_RE.sub("", pib_raw).strip() if isinstance(pib_raw, str) else pib_raw
        normalized = normalize_name(cleaned_pib)

        people[normalized] = {
            "pib_raw": cleaned_pib,
            "posada": ws.cell(row=row, column=posada_col).value if posada_col else None,
            "zvannya": ws.cell(row=row, column=zvannya_col).value if zvannya_col else None,
            "pidrozdil": pidrozdil,
        }

        for day, col_idx in day_columns.items():
            value = ws.cell(row=row, column=col_idx).value
            if isinstance(value, str) and value.strip():
                marks[(normalized, day)] = value.strip().lower()

    return marks, people
