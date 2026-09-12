import os
import re
import sys
from datetime import date as date_type, datetime
from glob import glob
from pathlib import Path

# Додаємо корінь проєкту в sys.path
sys.path.append(str(Path(__file__).resolve().parent.parent))

import pandas as pd
import pytest

from constants import (
    PERSONEL_LIST_FILE_NAME,
    PERSONEL_LIST_SHEET_NAME,
    PERSONEL_LIST_COLUMNS_LETTERS,
    TVO_LIST_SHEET_NAME,
    TVO_LIST_COLUMNS_LETTERS,
    DOWRIES_LIST_FILE_NAME,
    DOWRIES_LIST_SHEET_NAME,
    DOWRIES_LIST_COLUMNS_LETTERS,
    PRIDANI_SHEET_NAME,
    OUTPUT_DIR_BR,
    HIGHER_COMMANDER_TITLE,
    MONEY_REPORT_CATEGORIES,
    MONTH,
    YEAR,
)
from helpers import excel_col_to_index, get_data, read_datafile, read_optional_datafile, find_higher_commander, date_to_str, to_date
from utils.excel_reader import PERSONEL_LIST_FALLBACK_SHEET_NAMES, read_optional_pridani_sheet
from helpers_for_TEST_WORK import read_docx_text

_FILENAME_DATE_RE = re.compile(r"ЩОДЕННА (\d{2}\.\d{2}\.\d{4})\.docx$")

# Особовий склад у документі завжди записаний як "ПРІЗВИЩЕ Ім'я По-батькові;"
# (прізвище ВЕЛИКИМИ літерами, ім'я/по-батькові — Titlecase, одразу перед ";").
# Це дозволяє надійно витягнути саме ПІБ з тексту документа, не плутаючи його
# з званням (завжди мало написане) чи короткими підписами (без ";").
_PIB_ENTRY_RE = re.compile(
    r"([А-ЯЁЇЄІҐ]{2,}(?:['ʼ-][А-ЯЁЇЄІҐ]+)?\s+[А-ЯЁЇЄІҐ][а-яёїєіґ'ʼ-]+(?:\s+[А-ЯЁЇЄІҐ][а-яёїєіґ'ʼ-]+)?)\s*;"
)


def _existing_daily_br_files():
    if not os.path.isdir(OUTPUT_DIR_BR):
        return []
    found = []
    for path in glob(os.path.join(OUTPUT_DIR_BR, "*ЩОДЕННА *.docx")):
        if os.path.basename(path).startswith("~$"):
            continue  # тимчасовий lock-файл Word (документ відкритий), не справжній docx
        match = _FILENAME_DATE_RE.search(os.path.basename(path))
        if match:
            found.append((path, match.group(1)))
    return found


def _read_dowries_data():
    if not os.path.isfile(DOWRIES_LIST_FILE_NAME):
        return []
    data = pd.read_excel(DOWRIES_LIST_FILE_NAME, sheet_name=DOWRIES_LIST_SHEET_NAME)
    columns = [data.columns[excel_col_to_index(letter)] for letter in DOWRIES_LIST_COLUMNS_LETTERS]
    return get_data(data, columns, date_can_be_empty=True).get("rows", [])


def _dowry_pib_valid_for_date(col_name):
    current = to_date(col_name)
    valid = set()
    for row in _ROWS_WITH_DOWRIES_DATA:
        pib = row.get('Прізвище та ініціали')
        if not pib:
            continue
        row_from = to_date(row['ДАТА ПРИБУВ (З)']) if row.get('ДАТА ПРИБУВ (З)') else None
        row_to = to_date(row['ДАТА ВІДБУТТЯ (ПО)']) if row.get('ДАТА ВІДБУТТЯ (ПО)') else date_type.max
        if row_from and row_from < current < row_to:
            valid.add(pib)
    return valid


def _pridani_sheet_pib_valid_for_date(col_name):
    """Аркуш "ПРИДАНІ" У САМОМУ ОБЛІК.xlsx (на відміну від _dowry_pib_valid_for_date -
    ОКРЕМИЙ файл ПРИДАНІ.xlsx) - people тут можуть з'явитись у щоденному БР
    ЛИШЕ якщо мають бойовий статус (100/70/170) САМЕ за цю дату (як і звичайний
    особовий склад) - process_generate_br.py об'єднує ці рядки з rows_with_data
    ще ДО побудови БР, тож фільтр той самий (_COMBAT_CELL_VALUES)."""
    return {
        row['ПІБ'] for row in _ROWS_WITH_PRIDANI_SHEET_DATA
        if row.get(col_name) in _COMBAT_CELL_VALUES
    }


_DAILY_BR_FILES = _existing_daily_br_files()
_PERSONEL_DATA = read_datafile(
    PERSONEL_LIST_FILE_NAME, PERSONEL_LIST_SHEET_NAME, PERSONEL_LIST_COLUMNS_LETTERS,
    extra_fallback_sheet_names=PERSONEL_LIST_FALLBACK_SHEET_NAMES,
)
_ROWS_WITH_DATA = _PERSONEL_DATA.get("rows", [])
_COLUMNS = _PERSONEL_DATA.get("columns", [])
_ROWS_WITH_TVO_DATA = read_optional_datafile(PERSONEL_LIST_FILE_NAME, TVO_LIST_SHEET_NAME, TVO_LIST_COLUMNS_LETTERS, sheet_can_be_missing=True)
_ROWS_WITH_DOWRIES_DATA = _read_dowries_data()
_ROWS_WITH_PRIDANI_SHEET_DATA = read_optional_pridani_sheet(PERSONEL_LIST_FILE_NAME, PRIDANI_SHEET_NAME, MONTH, YEAR)
# Значення комірки ОБЛІК.xlsx, що рахуються як "бойовий" день для щоденного БР -
# усі точки MONEY_REPORT_CATEGORIES, крім 30 і 10 (той самий набір, що й у br_general_catalog.py).
_COMBAT_CELL_VALUES = set(MONEY_REPORT_CATEGORIES) - {30, 10}


def test_no_unexpected_pib_in_daily_br():
    """Зворотна перевірка: кожен ПІБ, що фактично присутній у документі ЩОДЕННА,
    має або мати бойовий статус (100/70/170) в ОБЛІК.xlsx за дату документа, або
    бути дійсним приданим на цю дату (ПРИДАНІ.xlsx, ОКРЕМИЙ файл, чи аркуш
    "ПРИДАНІ" У САМОМУ ОБЛІК.xlsx), або значитись у ТВО (ОБЛІК.xlsx, аркуш "ТВО").
    Якщо ні — це "зайвий" ПІБ: або він є в ОБЛІК.xlsx, але без бойового статусу
    (помилка-попередження), або його там взагалі немає (жорсткий фейл)."""
    if not _DAILY_BR_FILES:
        pytest.skip(f"Немає жодного згенерованого документа ЩОДЕННА в {OUTPUT_DIR_BR} — спочатку запустіть python index.py")

    main_roster = {row['ПІБ']: row for row in _ROWS_WITH_DATA if row.get('ПІБ')}
    tvo_pib = {row['ПІБ'] for row in _ROWS_WITH_TVO_DATA if row.get('ПІБ')}

    for path_to_file, date_str in _DAILY_BR_FILES:
        col_name = next((c for c in _COLUMNS if isinstance(c, datetime) and date_to_str(c) == date_str), None)
        if col_name is None:
            # Файл за місяць, який ЗАРАЗ не завантажено в ОБЛІК.xlsx (напр.
            # застарілий документ у OUTPUT_DIR_BR з ПОПЕРЕДНЬОГО запуску,
            # коли був обраний інший місяць) - ця перевірка стосується ЛИШЕ
            # файлів ПОТОЧНОГО завантаженого місяця, тож старий файл просто
            # пропускається, а не вважається помилкою (ОБЛІК.xlsx природно
            # не має колонки за місяць, який уже позаду).
            continue

        higher_commander_data = find_higher_commander(_ROWS_WITH_TVO_DATA, _ROWS_WITH_DATA, HIGHER_COMMANDER_TITLE, col_name)
        expected_pib = {
            row['ПІБ'] for row in _ROWS_WITH_DATA
            if row.get(col_name) in _COMBAT_CELL_VALUES and row.get('ПОСАДА') != higher_commander_data.get('ПОСАДА')
        }
        valid_dowry_pib = _dowry_pib_valid_for_date(col_name)
        valid_pridani_sheet_pib = _pridani_sheet_pib_valid_for_date(col_name)

        generated_text = read_docx_text(path_to_file)
        candidates = {m.strip() for m in _PIB_ENTRY_RE.findall(generated_text)}

        extra_wrong_status = []
        unknown_pib = []
        for pib in candidates:
            if pib in expected_pib or pib in valid_dowry_pib or pib in valid_pridani_sheet_pib or pib in tvo_pib:
                continue
            if pib in main_roster:
                actual_status = main_roster[pib].get(col_name)
                extra_wrong_status.append(f"{pib} (статус за {date_str}: {actual_status!r}, не бойовий)")
            else:
                unknown_pib.append(pib)

        assert not extra_wrong_status, (
            f"У документі {path_to_file} є зайві ПІБ без бойового статусу за {date_str}: " + "; ".join(extra_wrong_status)
        )
        assert not unknown_pib, (
            f"У документі {path_to_file} є ПІБ, яких немає в {PERSONEL_LIST_FILE_NAME}: " + "; ".join(unknown_pib)
        )
