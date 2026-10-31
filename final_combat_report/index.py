import io, os, sys, pandas as pd

# консоль Windows часто працює у cp1251/cp866 - без цього print() з кирилицею/емодзі падає.
# isinstance-перевірка звужує тип до TextIOWrapper, де reconfigure() точно є (задовольняє і Pyright).
if isinstance(sys.stdout, io.TextIOWrapper) and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if isinstance(sys.stderr, io.TextIOWrapper) and sys.stderr.encoding.lower() != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from document_final_combat_report import document_final_combat_report
from get_normalized_data.get_UAV_data import get_UAV_data
from get_normalized_data.get_active_data import get_active_data
from get_normalized_data.get_basketball_data import get_basketball_data
from get_normalized_data.get_paintball_data import get_paintball_data
from get_normalized_data.get_data_commission import get_data_commission
from get_data import get_actual_data_json, get_date, run_signal_export, get_hour_of_report
from constants import SIGNAL_GROUPS, OUTPUT_DIR, PERSONEL_LIST_FILE_NAME, PERSONEL_LIST_SHEET_NAME, PERSONEL_LIST_COLUMNS_LETTERS
from helpers import create_or_clear_output_directory, excel_col_to_index, get_rows

# ОБЛІК.xlsm
if not os.path.isfile(PERSONEL_LIST_FILE_NAME):
    raise FileNotFoundError(f"Файл {PERSONEL_LIST_FILE_NAME} не знайдено.")
print(f"Read document {PERSONEL_LIST_FILE_NAME}, {PERSONEL_LIST_SHEET_NAME}...")
personel_data = pd.read_excel(PERSONEL_LIST_FILE_NAME, sheet_name=PERSONEL_LIST_SHEET_NAME)
columns_personel_data = [excel_col_to_index(letter) for letter in PERSONEL_LIST_COLUMNS_LETTERS]
column_names_personel = [personel_data.columns[i] for i in columns_personel_data] # header wit data
rows_with_data = get_rows(personel_data, column_names_personel)

def index():
    if get_actual_data_json():
        run_signal_export(*SIGNAL_GROUPS)
    create_or_clear_output_directory(OUTPUT_DIR)
    hour_of_report = get_hour_of_report()
    selected_date = get_date(hour_of_report)

    data_paintball = get_paintball_data(hour_of_report, selected_date)
    data_basketball = get_basketball_data(hour_of_report, selected_date)
    data_uav = get_UAV_data(hour_of_report, selected_date, rows_with_data)
    data_active = get_active_data(hour_of_report, selected_date)
    data_commission = get_data_commission(hour_of_report, selected_date)

    document_final_combat_report(selected_date, hour_of_report, data_paintball, data_basketball, data_uav, data_active, data_commission, rows_with_data)

if __name__ == "__main__":
    index()