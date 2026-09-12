import glob
import os
import re
from datetime import timedelta

import content.information_unit_reader as iur
from content.daily_report_reader import report_date_from_filename
from content.oblik_timesheet import normalize_name

# Заголовок дня - точно 2-значний текст ("01".."31", як і сам генерований
# файл - content/rop_vop_statement.py::write_statement пише day_headers у
# ТОМУ САМОМУ форматі) - не число (реальний файл пише його як ТЕКСТ).
_DAY_HEADER_RE = re.compile(r"^\d{2}$")

# Реальний випадок: деякі ПІБ у resources/schedule мають дописку в дужках
# наприкінці (напр. "... (назва стороннього підрозділу)") - позначення
# прикріплення з ІНШОГО підрозділу, відсутнє в самому роcтері
# (resources/ОБЛІК.xlsx) - без відкидання цього суфіксу normalize_name дав
# би ІНШИЙ ключ, ніж роcтеровий, і людину не вдалось би зіставити взагалі.
# Група, що ЗАХОПЛЮЄ текст дужки (без самих дужок) - за прямою вказівкою
# користувача, ЦЕЙ текст показується як "Підрозділ" у звіті розбіжностей
# (write_schedule_mismatch_report, content/rop_vop_statement.py) для людей,
# яких не вдалось зіставити з роcтером - раніше просто відкидався.
_ATTACHMENT_SUFFIX_RE = re.compile(r"\s*\(([^)]*)\)\s*$")

# Скільки рядків від початку аркуша шукати заголовок - реальний файл має
# заголовок у рядку 6 (4 рядки титулу + порожній рядок), запас із головою.
_HEADER_SEARCH_ROWS = 20


def latest_schedule_file(schedule_dir, year, month):
    """Шлях НАЙПІЗНІШОГО файлу schedule_dir/*, чия ФАКТИЧНА дата належить
    САМЕ ЦЬОМУ year/month - lock-файли Excel ("~$...") і файли БЕЗ
    розпізнаваної дати чи ІНШОГО місяця/року пропускаються. None, якщо
    жодного такого файлу немає (тека порожня чи немає файлів на цей місяць)
    - викликач тоді просто нічого не звіряє.

    За прямою вказівкою користувача - дата в ІМЕНІ файлу (report_date_from_
    filename, ТОЙ САМИЙ принцип, що й resources/report/*.docx) - це дата
    "коли подано" (наступного ранку), а НЕ дата, ЗА ЯКУ сама відомість -
    реальний випадок: "{SHORT_UNIT_BATTALION без пробілу}ВОП-РОП(СЕРПЕНЬ) 01.09.2026.xlsx" - аркуш
    (титул) усередині каже "Серпень", хоча ім'я файлу - вже 01.09.2026
    (відомість ЗА 31.08.2026, подана ЗРАНКУ 01.09.2026) - ФАКТИЧНА дата
    файлу = дата з імені МІНУС один день."""
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
    """(header_row, pib_col, day_columns, posada_col, zvannya_col) -
    day_columns = {номер_дня: колонка}, УСІ знайдені ЗА ТЕКСТОМ заголовка
    (не фіксованою позицією - реальний файл має 8 ЗАЙВИХ колонок перед
    видимими даними, той самий принцип, що й _bchs_columns/
    _tabel_label_columns, content/information_unit_reader.py/
    content/payment_mismatch_checker.py). posada_col/zvannya_col - як і в
    _tabel_label_columns - НЕ обов'язкові (None, якщо немає такої колонки).
    (None, None, {}, None, None), якщо в межах _HEADER_SEARCH_ROWS не
    знайдено рядка з ОБОМА обов'язковими (колонка ПІБ і хоча б одна
    колонка-день)."""
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
    return None, None, {}, None, None


def read_schedule_marks(path):
    """Читає ОДИН файл resources/schedule/*.xlsx. Повертає (marks, people):
    - marks - {(normalize_name(ПІБ), день): "роп"/"воп"/...} ЛИШЕ для
      НЕПОРОЖНІХ клітинок-днів (порожня клітинка - "ще не заповнено" - за
      прямою вказівкою користувача, НЕ вважається твердженням "не на
      РОП/ВОП цього дня", тож не звіряється взагалі).
    - people - {normalize_name(ПІБ): {"pib_raw", "posada", "zvannya",
      "pidrozdil"}} - "знімок" ЦЬОГО файлу для КОЖНОЇ людини (posada/
      zvannya - None, якщо в файлі немає відповідної колонки; pidrozdil -
      текст дужкової дописки прикріплення, без дужок, або None, якщо
      дописки немає) - за прямою вказівкою користувача, потрібне
      write_schedule_mismatch_report (content/rop_vop_statement.py), щоб
      показати РЕАЛЬНІ дані людини з schedule, а не загальний напис "Не
      знайдено в ОБЛІК.xlsx" на обидві клітинки одразу.

    Рядки людей - від рядка одразу під заголовком, до ПЕРШОГО порожнього
    ПІБ (підписний блок нижче - без окремого порожнього рядка-роздільника
    перед ним у реальних файлах).

    Обидва словники порожні, якщо заголовок (колонка ПІБ + хоча б одна
    колонка-день) не вдалось знайти взагалі."""
    wb = iur._load_workbook(path)
    ws = wb.active
    header_row, pib_col, day_columns, posada_col, zvannya_col = _header_row_and_columns(ws)
    marks = {}
    people = {}
    if header_row is None:
        return marks, people

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
