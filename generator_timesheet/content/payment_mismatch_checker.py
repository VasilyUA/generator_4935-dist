import glob
import os
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import content.information_unit_reader as iur
from constants import (
    DEFAULT_STATUS,
    INFORMATION_UNIT_SHEET_NAME,
    PAYMENT_MISMATCH_COMPARABLE_STATUSES,
    PAYMENT_STATUS_CONVERSIONS,
    PAYMENT_STATUS_OVERRIDES,
    PAYMENT_STATUS_PRIORITY,
    PIB_COLUMN_NAME,
)
from content.oblik_timesheet import normalize_name

# Підпис рядка "офіційного" (за щоденними рапортами) джерела в звіті - на
# відміну від рядка-файлу information_unit, тут немає ОДНОГО конкретного
# файлу-джерела (значення обчислене з УСІХ щоденних рапортів + вже
# застосованих категорій виплат).
_REPORT_SOURCE_LABEL = "ОБЛІК для виплат (за рапортами)"

# Значення "офіційної" сторони для людини, якої ВЗАГАЛІ немає в ОБЛІК для
# виплат.xlsx (не в ростері) - застосовується до КОЖНОЇ порівнюваної колонки
# (ПІБ/Підрозділ/Посада/Звання/Статус), а не лише до дати, щоб роcтерова
# сторона ніде не лишалась порожньою клітинкою (користувач не повинен бачити
# пусті комірки - незрозуміло порожня клітинка й клітинка з явним поясненням
# "чому тут пусто" - не одне й те саме).
_NOT_IN_ROSTER_LABEL = "Не знайдено в ОБЛІК.xlsx"

# Заповнювач для БУДЬ-ЯКОГО іншого порожнього значення (напр. роcтер узагалі
# не має колонки ПІДРОЗДІЛ, або file information_unit не подає Посаду) - щоб
# візуально не лишалось "незрозуміло чому білих" клітинок серед кольорової
# гами; сама логіка збігу/розбіжності (_normalize_for_compare) працює з
# ОРИГІНАЛЬНИМ значенням (None), заповнювач - лише для відображення.
_EMPTY_PLACEHOLDER = "—"

# Стиль файлу (за прямою вказівкою користувача) - Times New Roman 14, зелена
# заливка там, де роcтер і файл information_unit збігаються, червона - де ні
# (той самий відтінок, що й STATUS_COLORS/checker_accounting/checker.py - для
# візуальної узгодженості з рештою проєкту). Межі - по всіх клітинках (і
# заголовку, і даних), щоб таблицю було легко читати оком навіть без кольору.
# wrap_text=False (+ щедрі ширини колонок нижче) - навмисно, замість переносу
# тексту в кілька рядків: за прямою вказівкою користувача перенесений текст
# ставав нечитабельним/накладався на сусідній рядок при висоті рядка 20.
_FONT = Font(name="Times New Roman", size=14)
_MATCH_FILL = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")
_MISMATCH_FILL = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)
_THIN_SIDE = Side(style="thin")
_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)

# Заголовок - той самий вигляд, що й заголовок аркуша "Табель" у самому
# ОБЛІК.xlsx (resources/ОБЛІК.xlsx: Bahnschrift Light SemiCondensed 14 жирний
# білий текст на темно-бірюзовій заливці FF035C6B) - за прямою вказівкою
# користувача, "вигляд як в ОБЛІК файлі", а не довільний стиль з нуля.
_HEADER_FONT = Font(name="Bahnschrift Light SemiCondensed", size=14, bold=True, color="FFFFFFFF")
_HEADER_FILL = PatternFill(start_color="FF035C6B", end_color="FF035C6B", fill_type="solid")
_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _normalize_for_compare(value):
    """Порожня клітинка (None) і рядок зводяться до одного вигляду (будь-які
    пробіли - в один, обрізані з країв, верхній регістр) - трохи ширше, ніж у
    generator_br_and_report_for_money/checker_accounting/checker.py (там лише
    strip()), бо ПІБ з різних джерел іноді відрізняється ВНУТРІШНІМИ
    подвійними пробілами, а не лише скраю - інакше такі суто косметичні
    відмінності хибно позначались би як розбіжність."""
    if value is None:
        return ""
    if isinstance(value, str):
        return " ".join(value.split()).upper()
    return value


def _display_value(value):
    """Значення для запису в клітинку - як є, окрім None/порожнього рядка
    (замінюється на _EMPTY_PLACEHOLDER, щоб не лишати "непояснено білих"
    клітинок). НЕ впливає на _normalize_for_compare - порівняння й далі
    працює з ОРИГІНАЛЬНИМ значенням."""
    if value is None or value == "":
        return _EMPTY_PLACEHOLDER
    return value


def _tabel_label_columns(ws):
    """(pidrozdil_col, posada_col, zvannya_col) з рядка 1 аркуша ТАБЕЛЬ - той
    самий підхід, що й information_unit_reader._pib_column (точний
    заголовок, а не фіксована позиція). Підрозділ - рідкісна колонка (не всі
    файли її мають), None - так само нормально, як і для Посада/Звання."""
    pidrozdil_col = posada_col = zvannya_col = None
    for c in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=c).value
        if not isinstance(value, str):
            continue
        upper = value.strip().upper()
        if upper == "ПІДРОЗДІЛ":
            pidrozdil_col = c
        elif upper == "ПОСАДА":
            posada_col = c
        elif upper == "ЗВАННЯ":
            zvannya_col = c
    return pidrozdil_col, posada_col, zvannya_col


def _read_tabel_snapshot(wb, path, file_name, snapshot):
    """Як information_unit_reader._read_tabel_sheet, але кожен запис - кортеж
    (file_name, pib_raw, pidrozdil, posada, zvannya, value), а не саме лише
    value - потрібна атрибуція джерела для звіту розбіжностей."""
    ws = iur._find_sheet(wb, INFORMATION_UNIT_SHEET_NAME)
    if ws is None:
        return f"У файлі {path} не знайдено аркуш ТАБЕЛЬ."

    pib_col = iur._pib_column(ws)
    date_row, date_cols = iur._tabel_date_row_and_columns(ws)
    if pib_col is None or not date_cols:
        return f"У файлі {path} (аркуш {ws.title}) не знайдено {'колонку ПІБ' if pib_col is None else 'жодної колонки-дати'}."

    pidrozdil_col, posada_col, zvannya_col = _tabel_label_columns(ws)

    for row in range(date_row + 1, ws.max_row + 1):
        pib_raw = ws.cell(row=row, column=pib_col).value
        if not pib_raw:
            continue
        normalized = normalize_name(pib_raw)
        pidrozdil = ws.cell(row=row, column=pidrozdil_col).value if pidrozdil_col else None
        posada = ws.cell(row=row, column=posada_col).value if posada_col else None
        zvannya = ws.cell(row=row, column=zvannya_col).value if zvannya_col else None
        for col, day in date_cols:
            value = iur._clean_cell_value(ws.cell(row=row, column=col).value)
            if value not in (None, ""):
                snapshot.setdefault((normalized, day), []).append((file_name, pib_raw, pidrozdil, posada, zvannya, value))

    return None


def _read_bchs_snapshot(wb, path, file_name, snapshot):
    """Як information_unit_reader._read_bchs_sheet, але з атрибуцією джерела
    (file_name, pidrozdil, posada, zvannya) на кожен запис - див.
    _read_tabel_snapshot."""
    ws = iur._find_sheet(wb, iur._BCHS_SHEET_NAME)
    if ws is None:
        return f"У файлі {path} не знайдено аркуш БЧС."

    day = iur._extract_bchs_date(path, ws)
    if day is None:
        return f"У файлі {path} (аркуш {ws.title}) не вдалось визначити дату (ні з назви файлу, ні із заголовка аркуша)."

    pib_col, pidrozdil_col, posada_col, zvannya_col, status_col, payment_col = iur._bchs_columns(ws)
    if pib_col is None:
        return f"У файлі {path} (аркуш {ws.title}) не знайдено колонку ПІБ."

    for row in range(iur._BCHS_HEADER_ROW + 1, ws.max_row + 1):
        pib_raw = ws.cell(row=row, column=pib_col).value
        if not pib_raw:
            continue
        normalized = normalize_name(pib_raw)
        pidrozdil = ws.cell(row=row, column=pidrozdil_col).value if pidrozdil_col else None
        posada = ws.cell(row=row, column=posada_col).value if posada_col else None
        zvannya = ws.cell(row=row, column=zvannya_col).value if zvannya_col else None

        status = iur._clean_cell_value(ws.cell(row=row, column=status_col).value if status_col else None)
        if isinstance(status, str) and status and status.upper() != DEFAULT_STATUS:
            snapshot.setdefault((normalized, day), []).append((file_name, pib_raw, pidrozdil, posada, zvannya, status))
            continue

        payment = iur._clean_cell_value(ws.cell(row=row, column=payment_col).value if payment_col else None)
        if payment not in (None, ""):
            snapshot.setdefault((normalized, day), []).append((file_name, pib_raw, pidrozdil, posada, zvannya, payment))

    return None


def read_information_unit_snapshot(directory):
    """Як information_unit_reader.read_payment_values (ТІ САМІ два джерела на
    файл - ТАБЕЛЬ і БЧС, ТАБЕЛЬ - ПРІОРИТЕТНЕ: якщо вдалось прочитати, БЧС
    того самого файлу взагалі не читається), але зберігає АТРИБУЦІЮ джерела
    для кожного запису - потрібно для write_mismatch_report (колонка "Назва
    файлу").

    Повертає (snapshot, skipped): snapshot - {(normalize_name(ПІБ), date):
    [(file_name, pib_raw, pidrozdil, posada, zvannya, value), ...]}; skipped -
    той самий формат {"reason"}, що й read_payment_values."""
    snapshot = {}
    skipped = []

    for path in sorted(glob.glob(os.path.join(directory, "*"))):
        if iur._is_office_lock_file(path):
            continue
        if os.path.splitext(path)[1].lower() not in iur._SUPPORTED_EXTENSIONS:
            skipped.append({"reason": f"Формат файлу не підтримується: {path}.", "report_date": None})
            continue

        file_name = os.path.basename(path)
        wb, load_reason = iur._safe_load_workbook(path)
        if wb is None:
            skipped.append({"reason": load_reason, "report_date": None})
            continue
        tabel_reason = _read_tabel_snapshot(wb, path, file_name, snapshot)
        if tabel_reason is None:
            continue

        bchs_reason = _read_bchs_snapshot(wb, path, file_name, snapshot)
        if bchs_reason is not None:
            skipped.append({"reason": tabel_reason, "report_date": None})

    return snapshot, skipped


def _surname_firstname_key(normalized):
    """Перші ДВА "слова" нормалізованого ПІБ (Прізвище Ім'я) - БЕЗ
    по-батькові. НЕ fuzzy-збіг (жодної відстані редагування чи схожості) -
    точний збіг лише перших двох слів; решта (по-батькові) навмисно
    ігнорується - реальний випадок: один файл information_unit написав
    по-батькові з друкарською помилкою (відрізняється на одну-дві літери від
    роcтерового написання), через що normalize_name дає ІНШИЙ ключ, ніж
    роcтеровий, і цю людину раніше НЕ вдавалось зіставити з роcтером
    узагалі."""
    parts = normalized.split(" ")
    return " ".join(parts[:2]) if len(parts) >= 2 else normalized


def _candidate_roster_person(person_rows, file_normalized):
    """row_idx ЄДИНОГО роcтерового рядка, чиї Прізвище+Ім'я збігаються з
    file_normalized (див. _surname_firstname_key), коли САМ file_normalized
    НЕ Є точним ключем у person_rows (інакше це вже не "кандидат", а точний
    збіг). Кандидатів 0 чи 2+ (неоднозначно - напр. однофамільці/тезки) -
    None: краще показати "не знайдено", ніж вгадати неправильну людину."""
    prefix = _surname_firstname_key(file_normalized)
    candidates = [row_idx for normalized, row_idx in person_rows.items() if _surname_firstname_key(normalized) == prefix]
    return candidates[0] if len(candidates) == 1 else None


def _missing_from_roster_records(timesheet, person_rows, snapshot):
    """Записи для людей, які Є в snapshot (information_unit), але їхній
    normalize_name НЕ збігається з ЖОДНИМ роcтеровим рядком - або тому, що
    людини ВЗАГАЛІ немає в ОБЛІК для виплат.xlsx (не в ростері), або тому, що
    ПІБ написане ІНАКШЕ, ніж у роcтері (типова причина - друкарська помилка в
    по-батькові). Без цього вони мовчки зникали б зі звіту, бо основний
    прохід find_mismatches іде ПО РОСТЕРУ, а не по snapshot.

    Якщо серед роcтера є РІВНО ОДИН кандидат із таким самим Прізвищем+Ім'ям
    (_candidate_roster_person) - роcтерова сторона показує РЕАЛЬНЕ ПІБ/
    Підрозділ/Посада/Звання/поточний статус цього кандидата (щоб розбіжність
    у написанні ПІБ було видно ПОРЯД, а не лише загальний напис "не
    знайдено") - інакше (0 чи 2+ кандидатів, неоднозначно) роcтерова сторона -
    _NOT_IN_ROSTER_LABEL для КОЖНОЇ з цих колонок (не лише дати), а не None -
    інакше write_mismatch_report лишав би роcтерову сторону порожньою
    клітинкою без пояснення."""
    pib_col = timesheet.label_columns[PIB_COLUMN_NAME]
    pidrozdil_col = timesheet.label_columns.get("ПІДРОЗДІЛ")
    posada_col = timesheet.label_columns.get("ПОСАДА")
    zvannya_col = timesheet.label_columns.get("ЗВАННЯ")
    date_col_by_date = {d: idx for idx, d in timesheet.date_columns}

    candidate_by_normalized = {}
    by_person_file = {}
    for (normalized, date_value), entries in snapshot.items():
        if normalized in person_rows:
            continue
        if normalized not in candidate_by_normalized:
            candidate_by_normalized[normalized] = _candidate_roster_person(person_rows, normalized)
        candidate_row = candidate_by_normalized[normalized]

        if candidate_row is not None and date_value in date_col_by_date:
            roster_value = timesheet.ws.cell(row=candidate_row, column=date_col_by_date[date_value]).value
        else:
            roster_value = _NOT_IN_ROSTER_LABEL

        for file_name, pib_raw, file_pidrozdil, file_posada, file_zvannya, value in entries:
            entry = by_person_file.setdefault(
                (normalized, file_name),
                {
                    "pib_raw": pib_raw, "pidrozdil": file_pidrozdil, "posada": file_posada, "zvannya": file_zvannya,
                    "dates": {}, "candidate_row": candidate_row,
                },
            )
            entry["dates"][date_value] = (roster_value, value)

    # Дедуплікація ФАЙЛІВ з ідентичним вмістом (_dedupe_files_with_identical_
    # content, той самий принцип, що й у find_mismatches) - ОКРЕМО для КОЖНОЇ
    # людини (by_person_file змішує кількох людей в одному словнику, тож
    # групуємо за normalized перед дедуплікацією, щоб не сплутати ДВОХ РІЗНИХ
    # людей, чиї дані випадково збіглись).
    by_normalized = {}
    for (normalized, file_name), info in by_person_file.items():
        by_normalized.setdefault(normalized, {})[file_name] = info
    deduped_items = [
        ((normalized, file_name), info)
        for normalized, files in by_normalized.items()
        for file_name, info in _dedupe_files_with_identical_content(files).items()
    ]

    records = []
    for (_normalized, file_name), info in sorted(deduped_items, key=lambda item: (item[0][1], item[1]["pib_raw"] or "")):
        candidate_row = info["candidate_row"]
        if candidate_row is not None:
            roster_pib_raw = timesheet.ws.cell(row=candidate_row, column=pib_col).value
            roster_pidrozdil = timesheet.ws.cell(row=candidate_row, column=pidrozdil_col).value if pidrozdil_col else None
            roster_posada = timesheet.ws.cell(row=candidate_row, column=posada_col).value if posada_col else None
            roster_zvannya = timesheet.ws.cell(row=candidate_row, column=zvannya_col).value if zvannya_col else None
        else:
            roster_pib_raw = roster_pidrozdil = roster_posada = roster_zvannya = _NOT_IN_ROSTER_LABEL

        records.append({
            "file_name": file_name,
            "roster_pib_raw": roster_pib_raw,
            "roster_pidrozdil": roster_pidrozdil,
            "roster_posada": roster_posada,
            "roster_zvannya": roster_zvannya,
            "file_pib_raw": info["pib_raw"],
            "file_pidrozdil": info["pidrozdil"],
            "file_posada": info["posada"],
            "file_zvannya": info["zvannya"],
            "dates": info["dates"],
        })
    return records


def _is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _is_comparable_status(value):
    """Порівняння з information_unit має сенс лише поки людина досі "на
    обліку" підрозділу для виплат: число (вже застосована категорія виплати)
    чи один зі статусів PAYMENT_MISMATCH_COMPARABLE_STATUSES (constants.py -
    за замовчуванням DEFAULT_STATUS "РВЗ" і "ВД"). Для БУДЬ-ЯКОГО ІНШОГО
    статусу (відпустка/шпиталь/СЗЧ/ВЛК/адміністративний) - подальше
    звітування підрозділу НЕ суперечить цьому статусу (людина або й далі
    отримує виплату на загальних підставах, або дані підрозділу вже застаріли
    й не стосуються справи) - порівняння НЕ проводиться, щоб не створювати
    хибних розбіжностей (реальні випадки: людина в СЗЧ, людина у відпустці за
    станом здоров'я - жоден із цих статусів не повинен породжувати запис у
    error_mis_statuses.xlsx, незалежно від того, що каже файл
    information_unit)."""
    return value in PAYMENT_MISMATCH_COMPARABLE_STATUSES or _is_number(value)


def _statuses_equivalent(a, b):
    """Дві сторони "порівнюваної" клітинки вважаються ЕКВІВАЛЕНТНИМИ (не
    розбіжністю), якщо або їхні нормалізовані значення буквально збігаються
    (_normalize_for_compare), або вони - ДВА РІЗНІ ЗАПИСИ ОДНІЄЇ й ТІЄЇ Ж
    реальної ситуації через PAYMENT_STATUS_CONVERSIONS чи PAYMENT_STATUS_
    OVERRIDES (constants.py, обидва - "статус з інформ_unit -> значення в
    ОБЛІК для виплат" мапи, застосовані apply_payment_values,
    content/oblik_timesheet.py, ще ДО того, як цей звіт узагалі щось
    порівнює). Реальні випадки: apply_payment_values ВЖЕ конвертував
    базовий статус "ППД" у фіксоване число 10 (PAYMENT_STATUS_CONVERSIONS),
    а файл information_unit - свій ВЛАСНИЙ, незалежний словник статусів - і
    далі каже текстом "ППД" для тієї самої дати; так само "РВЗ" ВЖЕ
    замінено на 10 через "БЗВП"/"Адап" (PAYMENT_STATUS_OVERRIDES[DEFAULT_
    STATUS]), а файл і далі каже текстом "БЗВП"/"Адап". Це та сама реальна
    ситуація, записана по-різному двома різними джерелами, а НЕ розбіжність
    - без цього "10" і "ППД"/"БЗВП"/"Адап" ніколи не могли б збігтись (різні
    типи значень), тож масово хибно потрапляли б у error_mis_statuses.xlsx.

    ТАКОЖ еквівалентна (не розбіжність) - пара з PAYMENT_STATUS_PRIORITY
    (constants.py, той самий об'єкт, керований користувачем самостійно, що
    й apply_payment_values._resolve_value_priority_conflict) - НЕЗАЛЕЖНО
    від того, з якого боку (роcтерове ФІНАЛЬНЕ значення чи сире значення
    information_unit) прийшов кожен статус пари: якщо для цієї пари ВЖЕ Є
    визначений переможець - подальше звітування підрозділу ІНШИМ статусом
    пари теж НЕ вважається помилкою, а лише підтвердженням того самого
    факту з нижчим пріоритетом. Реальний випадок, підтверджений
    користувачем: "100_БЗ" (роcтер, вже конвертоване "БЗ" - зниклий
    безвісті) і "Розп" (розпорядження, information_unit) - "БЗ" переважає,
    тож ЦЕ не розбіжність."""
    if _normalize_for_compare(a) == _normalize_for_compare(b):
        return True
    pair = {_normalize_for_compare(a), _normalize_for_compare(b)}
    for status_text, converted in PAYMENT_STATUS_CONVERSIONS.items():
        if pair == {_normalize_for_compare(status_text), _normalize_for_compare(converted)}:
            return True
    for override_map in PAYMENT_STATUS_OVERRIDES.values():
        for info_unit_value, replacement in override_map.items():
            if pair == {_normalize_for_compare(info_unit_value), _normalize_for_compare(replacement)}:
                return True
    for status_pair in PAYMENT_STATUS_PRIORITY:
        if pair == {_normalize_for_compare(status) for status in status_pair}:
            return True
    return False


def _dedupe_files_with_identical_content(by_file):
    """Якщо ДВА (чи більше) файли information_unit дають АБСОЛЮТНО ІДЕНТИЧНІ
    дані для ОДНІЄЇ людини (той самий ПІБ/Підрозділ/Посада/Звання файлу І
    той самий набір дата -> значення) - реальний випадок: це буквально ОДИН файл, що
    існує на диску ДВІЧІ під ледь різними іменами (напр. одна зайва пробіл у
    назві - типова причина, випадкове дублювання файлу під час копіювання).
    Лишає лише файл з АЛФАВІТНО ПЕРШОЮ назвою - інакше користувач бачив би
    той самий рядок розбіжності двічі в error_mis_statuses.xlsx."""
    seen_signatures = set()
    deduped = {}
    for file_name in sorted(by_file):
        info = by_file[file_name]
        signature = (
            info["pib_raw"], info["pidrozdil"], info["posada"], info["zvannya"],
            tuple(sorted(info["dates"].items())),
        )
        if signature in seen_signatures:
            continue
        seen_signatures.add(signature)
        deduped[file_name] = info
    return deduped


def find_mismatches(timesheet, snapshot):
    """Для КОЖНОЇ людини й КОЖНОЇ дати, де snapshot (read_information_unit_
    snapshot) має запис - звіряє ФІНАЛЬНЕ значення в timesheet (ОБЛІК для
    виплат, уже ПІСЛЯ apply_payment_values) із тим, що каже КОЖЕН файл
    ОКРЕМО, і ЗАВЖДИ зберігає пару (final_value, file_value) у dates (навіть
    коли final_value - "непорівнюваний" статус, напр. "ВПСЗ"/"СЗЧ", чи коли
    значення збігаються) - щоб рядок звіту показував ЦІЛУ "історію" по
    кожній даті цього файлу (збіг чи ні), а не лише розбіжні дати з
    порожніми клітинками поміж них (реальні випадки: АНДРЕЇШИН/ТРОФІМОВ -
    дати з "ВПСЗ"/"СЗЧ", де файл ПОГОДЖУЄТЬСЯ з роcтером, раніше повністю
    пропускались, навіть коли рядок ВЖЕ показувався через справжню
    розбіжність на ІНШУ дату).

    Розбіжність (has_mismatch = True, визначає, чи файл узагалі потрапляє в
    результат) фіксується ЛИШЕ коли ОБИДВІ умови виконані: (1) final_value -
    "порівнюваний" статус (_is_comparable_status: PAYMENT_MISMATCH_
    COMPARABLE_STATUSES з constants.py, за замовчуванням РВЗ/ВД, чи будь-яке
    число - реальний випадок: людина вже "ВД" (відрядження) за щоденними
    рапортами, а підрозділ через кілька днів подає число - теж РЕАЛЬНА
    розбіжність), і (2) значення НЕ еквівалентні (_statuses_equivalent -
    враховує PAYMENT_STATUS_CONVERSIONS: "10" з боку ОБЛІК і текстове "ППД"
    з боку файлу - ТА САМА реальна ситуація, а не розбіжність). "Непорівнюваний"
    статус (СЗЧ/ВПСЗ/ВЛК/...) НЕ створює розбіжність САМ ПО СОБІ, незалежно
    від того, погоджується файл чи ні - але ЯКЩО рядок вже показується через
    ІНШУ, справжню розбіжність - ця дата теж відображається (зелена, якщо
    погоджується, червона, якщо ні), для повної прозорості.

    Файли з АБСОЛЮТНО ІДЕНТИЧНИМ набором даних для однієї людини (типова
    причина - файл випадково задубльовано на диску під трохи різним іменем)
    - зводяться до ОДНОГО (_dedupe_files_with_identical_content), щоб той
    самий рядок розбіжності не показувався двічі.

    ОКРЕМО (_missing_from_roster_records) - люди, чий normalize_name НЕ
    збігається з ЖОДНИМ роcтеровим рядком (відсутні в ростері АБО написані
    інакше, напр. друкарська помилка в по-батькові) - інакше вони мовчки
    зникали б зі звіту, бо основний прохід іде ПО РОСТЕРУ.

    Повертає список записів (по одному на кожну пару людина+файл із хоча б
    ОДНІЄЮ розбіжною датою) - {"file_name", "roster_pib_raw", "roster_pidrozdil",
    "roster_posada", "roster_zvannya", "file_pib_raw", "file_pidrozdil",
    "file_posada", "file_zvannya", "dates": {date: (final_value, file_value)}}."""
    pib_col = timesheet.label_columns[PIB_COLUMN_NAME]
    pidrozdil_col = timesheet.label_columns.get("ПІДРОЗДІЛ")
    posada_col = timesheet.label_columns.get("ПОСАДА")
    zvannya_col = timesheet.label_columns.get("ЗВАННЯ")

    person_rows = timesheet.person_rows()
    records = []
    for normalized, row_idx in person_rows.items():
        pib_raw = timesheet.ws.cell(row=row_idx, column=pib_col).value
        roster_pidrozdil = timesheet.ws.cell(row=row_idx, column=pidrozdil_col).value if pidrozdil_col else None
        roster_posada = timesheet.ws.cell(row=row_idx, column=posada_col).value if posada_col else None
        roster_zvannya = timesheet.ws.cell(row=row_idx, column=zvannya_col).value if zvannya_col else None

        by_file = {}
        for col_idx, date_value in timesheet.date_columns:
            final_value = timesheet.ws.cell(row=row_idx, column=col_idx).value
            for file_name, file_pib_raw, file_pidrozdil, file_posada, file_zvannya, value in snapshot.get((normalized, date_value), []):
                entry = by_file.setdefault(
                    file_name,
                    {"pib_raw": file_pib_raw, "pidrozdil": file_pidrozdil, "posada": file_posada, "zvannya": file_zvannya, "dates": {}, "has_mismatch": False},
                )
                entry["dates"][date_value] = (final_value, value)
                if _is_comparable_status(final_value) and not _statuses_equivalent(value, final_value):
                    entry["has_mismatch"] = True

        by_file = _dedupe_files_with_identical_content(by_file)
        for file_name, info in sorted(by_file.items()):
            if not info["has_mismatch"]:
                continue
            records.append({
                "file_name": file_name,
                "roster_pib_raw": pib_raw,
                "roster_pidrozdil": roster_pidrozdil,
                "roster_posada": roster_posada,
                "roster_zvannya": roster_zvannya,
                "file_pib_raw": info["pib_raw"],
                "file_pidrozdil": info["pidrozdil"],
                "file_posada": info["posada"],
                "file_zvannya": info["zvannya"],
                "dates": info["dates"],
            })

    records.extend(_missing_from_roster_records(timesheet, person_rows, snapshot))
    return records


def _write_compared_cell(ws, row_a, row_b, col_idx, roster_value, file_value, match=None):
    """Одна "порівнювана" колонка (ПІБ/Підрозділ/Посада/Звання/Статус): якщо
    роcтер і файл дають ОДНАКОВЕ значення - ОДНА об'єднана клітинка, зелена
    заливка (_MATCH_FILL); інакше - 2 окремих значення (по одному на рядок),
    червона заливка на ОБОХ (_MISMATCH_FILL) - "весь рядочок" зеленим/червоним
    за прямою вказівкою користувача. Порожнє значення (None/"") відображається
    як _EMPTY_PLACEHOLDER (_display_value). За замовчуванням "збіг" - за
    ОРИГІНАЛЬНим значенням (_normalize_for_compare); колонки-дати (статус)
    передають власний, ШИРШИЙ критерій через `match` (_statuses_equivalent -
    враховує PAYMENT_STATUS_CONVERSIONS, напр. "10" і "ППД" - той самий
    статус, записаний по-різному), щоб не позначати ЕКВІВАЛЕНТНІ, а не лише
    буквально ІДЕНТИЧНІ значення як розбіжність."""
    is_match = match if match is not None else _normalize_for_compare(roster_value) == _normalize_for_compare(file_value)
    cell_a = ws.cell(row=row_a, column=col_idx, value=_display_value(roster_value))
    if is_match:
        ws.merge_cells(start_row=row_a, start_column=col_idx, end_row=row_b, end_column=col_idx)
        cell_a.fill = _MATCH_FILL
    else:
        cell_a.fill = _MISMATCH_FILL
        ws.cell(row=row_b, column=col_idx, value=_display_value(file_value)).fill = _MISMATCH_FILL


def _write_manual_entry_cell(ws, row_a, row_b, col_idx):
    """Одна колонка для РУЧНОГО заповнення (ПІДСТАВИ/ДАТА В СТАТУС
    СПЕЦКОНТИНГЕНТУ) - жодних даних для цього немає в жодному наявному
    джерелі (роcтер/рапорти/information_unit), за прямою вказівкою
    користувача - лишається ПОРОЖНЬОЮ (об'єднана клітинка, без заливки, без
    _EMPTY_PLACEHOLDER - людина заповнює вручну, "—" лише заважав би)."""
    ws.merge_cells(start_row=row_a, start_column=col_idx, end_row=row_b, end_column=col_idx)


# Порядок колонок - за прямою вказівкою користувача, той самий, що й у
# офіційному шаблоні звіту: Назва файлу, Підрозділ, Посада, Звання, ПІБ,
# дати, і, наприкінці, дві колонки для РУЧНОГО заповнення.
_LABEL_HEADERS = ["Назва файлу", "Підрозділ", "Посада", "Звання", "ПІБ"]
_LABEL_COLUMN_WIDTHS = (50, 26, 30, 24, 30)
_DATE_COLUMN_WIDTH = 24
_MANUAL_ENTRY_HEADERS = ["ПІДСТАВИ", "ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ"]
_MANUAL_ENTRY_COLUMN_WIDTHS = (30, 30)


def write_mismatch_report(records, output_path):
    """Будує output_path (.xlsx) - за прямою вказівкою користувача, СТРУКТУРА
    колонок - як в офіційному шаблоні звіту: Назва файлу, Підрозділ, Посада,
    Звання, ПІБ, потім КОЖНА дата - ОКРЕМА колонка (а не одна спільна
    колонка "Дата", заголовок - сама дата, як в самому ОБЛІК.xlsx), і
    наприкінці - дві колонки для РУЧНОГО заповнення (ПІДСТАВИ, ДАТА В
    СТАТУС СПЕЦКОНТИНГЕНТУ - немає джерела даних для них у проєкті, завжди
    лишаються порожніми). Для КОЖНОГО запису (людина + файл-джерело з
    розбіжністю) - ДВА рядки: перший - _REPORT_SOURCE_LABEL, другий -
    конкретний файл information_unit; ОДНА людина+файл із кількома
    розбіжними датами - ОДНА пара рядків із кількома заповненими
    колонками-датами (решта колонок-дат для цього запису - порожні, межі
    клітинок все одно є). "Підрозділ"/"ПІБ"/"Посада"/"Звання" і кожна дата -
    ОДНА об'єднана клітинка (зелена заливка), якщо роcтер і файл дають
    ОДНАКОВЕ значення, інакше - 2 окремих значення (по одному на рядок,
    червона заливка) - за прямою вказівкою користувача, "Підрозділ" ТЕЖ
    звіряється (раніше - ні: файл information_unit часто законно каже щось
    інше, напр. відрядження/прикріплення, тож розбіжність не обов'язково
    помилка - але користувач попросив звіряти й показувати її так само, як
    решту колонок, а не приховувати). "Назва файлу" (сама лише ідентифікація
    джерела рядка, а не порівнюване значення) заливкою не позначається. Заголовок -
    той самий вигляд, що й в ОБЛІК.xlsx (Bahnschrift Light SemiCondensed 14
    жирний білий на темно-бірюзовому); дані - Times New Roman 14. Тонкі межі
    по КОЖНІЙ клітинці (і заголовку, і даних), текст БЕЗ переносу
    (wrap_text=False) - щедрі ширини колонок нижче замість переносу в кілька
    рядків.

    Повертає output_path, АБО None, якщо records порожній - за прямою
    вказівкою користувача, немає сенсу створювати файл лише із заголовком,
    коли розповідати нема про що."""
    if not records:
        return None

    all_dates = sorted({date_value for record in records for date_value in record["dates"]})

    wb = Workbook()
    ws = wb.active
    ws.title = "Розбіжності"

    date_headers = [datetime.combine(d, datetime.min.time()) for d in all_dates]
    headers = _LABEL_HEADERS + date_headers + _MANUAL_ENTRY_HEADERS
    for col_idx, value in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=value)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGNMENT
        cell.border = _BORDER
        if isinstance(value, datetime):
            cell.number_format = "DD.MM.YYYY"
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 45

    date_col_by_date = {d: idx for idx, d in enumerate(all_dates, start=6)}
    manual_col_start = 6 + len(all_dates)

    current_row = 2
    for record in records:
        row_a, row_b = current_row, current_row + 1

        ws.cell(row=row_a, column=1, value=_REPORT_SOURCE_LABEL)
        ws.cell(row=row_b, column=1, value=record["file_name"])

        for col_idx, roster_value, file_value in (
            (2, record["roster_pidrozdil"], record["file_pidrozdil"]),
            (3, record["roster_posada"], record["file_posada"]),
            (4, record["roster_zvannya"], record["file_zvannya"]),
            (5, record["roster_pib_raw"], record["file_pib_raw"]),
        ):
            _write_compared_cell(ws, row_a, row_b, col_idx, roster_value, file_value)

        for date_value, (final_value, file_value) in record["dates"].items():
            _write_compared_cell(
                ws, row_a, row_b, date_col_by_date[date_value], final_value, file_value,
                match=_statuses_equivalent(final_value, file_value),
            )

        for col_idx in range(manual_col_start, manual_col_start + len(_MANUAL_ENTRY_HEADERS)):
            _write_manual_entry_cell(ws, row_a, row_b, col_idx)

        for row in (row_a, row_b):
            ws.row_dimensions[row].height = 20
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row, column=col_idx)
                cell.font = _FONT
                cell.alignment = _CENTER
                cell.border = _BORDER

        current_row += 2

    for col_idx in range(1, len(headers) + 1):
        letter = get_column_letter(col_idx)
        if col_idx <= 5:
            width = _LABEL_COLUMN_WIDTHS[col_idx - 1]
        elif col_idx < manual_col_start:
            width = _DATE_COLUMN_WIDTH
        else:
            width = _MANUAL_ENTRY_COLUMN_WIDTHS[col_idx - manual_col_start]
        ws.column_dimensions[letter].width = width

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    try:
        wb.save(output_path)
    except PermissionError as e:
        raise PermissionError(
            f"Не вдалось зберегти '{output_path}': файл зараз відкритий в іншій програмі "
            "(напр. Excel). Закрийте його та запустіть генерацію знову."
        ) from e
    return output_path
