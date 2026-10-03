import os
import sys
from datetime import datetime
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
    OUTPUT_DIR,
    OUTPUT_DIR_BR,
    OUTPUT_DIR_EXTRACT_BR,
    OUTPUT_DIR_BR_SAVE,
    OUTPUT_DIR_COMBAT_LOG_EXTRACT_WAR,
    NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY,
    HIGHER_COMMANDER_TITLE,
)
from helpers import excel_col_to_index, get_data, read_datafile, read_optional_datafile, find_higher_commander, date_to_str
from process_generate_br import process_generate_br
from helpers_for_TEST_WORK import read_docx_text
from utils.excel_reader import PERSONEL_LIST_FALLBACK_SHEET_NAMES

CITY = "КИЇВ"
COORDINATES = "36T TT 12345 67890"


def _read_dowries_data():
    if not os.path.isfile(DOWRIES_LIST_FILE_NAME):
        return []
    data = pd.read_excel(DOWRIES_LIST_FILE_NAME, sheet_name=DOWRIES_LIST_SHEET_NAME)
    columns = [data.columns[excel_col_to_index(letter)] for letter in DOWRIES_LIST_COLUMNS_LETTERS]
    return get_data(data, columns, date_can_be_empty=True).get("rows", [])


@pytest.fixture(scope="module")
def generated_output():
    """Той самий пайплайн, що й python index.py, але на реальних resources/ файлах."""
    personel_data = read_datafile(
        PERSONEL_LIST_FILE_NAME, PERSONEL_LIST_SHEET_NAME, PERSONEL_LIST_COLUMNS_LETTERS,
        extra_fallback_sheet_names=PERSONEL_LIST_FALLBACK_SHEET_NAMES,
    )
    tvo_rows = read_optional_datafile(PERSONEL_LIST_FILE_NAME, TVO_LIST_SHEET_NAME, TVO_LIST_COLUMNS_LETTERS, sheet_can_be_missing=True)
    dowries_rows = _read_dowries_data()

    columns = personel_data.get("columns", [])
    rows = personel_data.get("rows", [])

    try:
        result = process_generate_br(
            column_names_personel=columns,
            rows_with_data=rows,
            rows_with_tvo_data=tvo_rows,
            rows_with_dowries_data=dowries_rows,
            city=CITY,
            coordinates=COORDINATES,
        )
    except PermissionError as e:
        pytest.skip(f"Не вдалось перегенерувати {OUTPUT_DIR}: файл відкритий в іншій програмі (напр. Word) — закрийте його і повторіть. ({e})")

    return {"result": result, "rows": rows, "columns": columns, "tvo_rows": tvo_rows}


def _configured_date_columns(columns):
    return [c for c in columns if isinstance(c, datetime) and date_to_str(c) in NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY]


def test_main_runs_end_to_end_on_real_data(generated_output):
    assert generated_output["result"] is True
    for output_dir in (OUTPUT_DIR, OUTPUT_DIR_BR, OUTPUT_DIR_EXTRACT_BR, OUTPUT_DIR_BR_SAVE, OUTPUT_DIR_COMBAT_LOG_EXTRACT_WAR):
        assert os.path.isdir(output_dir), f"Директорія {output_dir} не була створена"


def test_daily_br_generated_for_every_configured_date(generated_output):
    date_columns = _configured_date_columns(generated_output["columns"])
    assert date_columns, f"У {PERSONEL_LIST_FILE_NAME} немає жодної дати, налаштованої в NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY"

    existing_files = {f for f in os.listdir(OUTPUT_DIR_BR) if not f.startswith("~$")}
    for col in date_columns:
        date_str = date_to_str(col)
        assert any(f.endswith(f"ЩОДЕННА {date_str}.docx") for f in existing_files), \
            f"Не знайдено документ ЩОДЕННА за {date_str} у {OUTPUT_DIR_BR}"


def test_daily_br_content_reflects_real_personnel_status(generated_output):
    """Гнучка перевірка (без фіксованої кількості символів): бере одну реальну дату
    з ОБЛІК.xlsx і перевіряє, що людина зі статусом 100 за цю дату потрапила
    в текст відповідного документа ЩОДЕННА, а сам ІКСП/координати теж на місці."""
    rows = generated_output["rows"]
    columns = generated_output["columns"]
    tvo_rows = generated_output["tvo_rows"]

    date_columns = _configured_date_columns(columns)
    assert date_columns

    col = date_columns[0]
    date_str = date_to_str(col)
    matches = [f for f in os.listdir(OUTPUT_DIR_BR) if not f.startswith("~$") and f.endswith(f"ЩОДЕННА {date_str}.docx")]
    assert matches, f"Не знайдено документ ЩОДЕННА за {date_str}"

    text = read_docx_text(os.path.join(OUTPUT_DIR_BR, matches[0]))
    assert CITY in text
    assert COORDINATES in text

    higher_commander_data = find_higher_commander(tvo_rows, rows, HIGHER_COMMANDER_TITLE, col)
    person_with_status_100 = next(
        (row for row in rows if row.get(col) == 100 and row.get('ПОСАДА') != higher_commander_data.get('ПОСАДА') and row.get('ПІБ')),
        None,
    )
    assert person_with_status_100 is not None, f"У {PERSONEL_LIST_FILE_NAME} немає жодної людини зі статусом 100 за {date_str}"
    assert person_with_status_100['ПІБ'] in text
