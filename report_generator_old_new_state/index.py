import pandas as pd
import warnings

from helpers import (
    create_or_clear_output_directory,
    excel_col_to_index,
    get_rows,
    input_date
)
from constants import (
    OUTPUT_DIR, PERSONEL_LIST_FILE_NAME_OLD, TASK_LIST_FILE_NAME, TASK_LIST_SHEET_NAME, PERSONEL_LIST_FILE_NAME_NEW,
    TRANSFER_LIST_FILE_NAME, PERSONEL_LIST_SHEET_NAME,
    TRANSFER_LIST_SHEET_NAME, PERSONEL_LIST_COLUMNS_LETTERS_OLD,
    TRANSFER_LIST_FILE_COLUMNS_LETTERS, TRANSFER_LIST_FILE_COLUMNS_LETTERS, PERSONEL_LIST_COLUMNS_LETTERS_NEW,
    TASK_LIST_FILE_COLUMNS_LETTERS, TVO_LIST_SHEET_NAME, TVO_LIST_COLUMNS_LETTERS
)
from generate import generate

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", category=FutureWarning)



def main():
    d = input_date()

    create_or_clear_output_directory(OUTPUT_DIR)

    print("Read personel data old!")
    personel_data_old = pd.read_excel(PERSONEL_LIST_FILE_NAME_OLD, sheet_name=PERSONEL_LIST_SHEET_NAME)
    print("Read personel data new!")
    personel_data_new = pd.read_excel(PERSONEL_LIST_FILE_NAME_NEW, sheet_name=PERSONEL_LIST_SHEET_NAME)
    print("Read transfer new!")
    transfer_data = pd.read_excel(TRANSFER_LIST_FILE_NAME, sheet_name=TRANSFER_LIST_SHEET_NAME)
    print("Read task data!")
    task_data = pd.read_excel(TASK_LIST_FILE_NAME, sheet_name=TASK_LIST_SHEET_NAME)
    print("Read TVO data!")
    tvo_data = pd.read_excel(TASK_LIST_FILE_NAME, sheet_name=TVO_LIST_SHEET_NAME)

    # Отримуємо індекси колонок за літерами
    columns_personel_data_old = [excel_col_to_index(letter) for letter in PERSONEL_LIST_COLUMNS_LETTERS_OLD]
    columns_personel_data_new = [excel_col_to_index(letter) for letter in PERSONEL_LIST_COLUMNS_LETTERS_NEW]
    columns_transfer_data = [ord(letter.upper()) - ord('A') for letter in TRANSFER_LIST_FILE_COLUMNS_LETTERS]
    columns_task_data = [ord(letter.upper()) - ord('A') for letter in TASK_LIST_FILE_COLUMNS_LETTERS]
    columns_tvo_data = [ord(letter.upper()) - ord('A') for letter in TVO_LIST_COLUMNS_LETTERS]

    # Отримуємо назви колонок за літерами
    column_names_personel_old = [personel_data_old.columns[i] for i in columns_personel_data_old]
    column_names_personel_new = [personel_data_new.columns[i] for i in columns_personel_data_new]
    column_names_transfer = [transfer_data.columns[i] for i in columns_transfer_data]
    column_names_task = [task_data.columns[i] for i in columns_task_data]
    column_names_tvo = [tvo_data.columns[i] for i in columns_tvo_data]

    rows_personel_old = get_rows(personel_data_old, column_names_personel_old)
    rows_personel_new = get_rows(personel_data_new, column_names_personel_new)
    rows_move = get_rows(transfer_data, column_names_transfer)
    rows_task = get_rows(task_data, column_names_task)
    rows_tvo = get_rows(tvo_data, column_names_tvo)
    
    generate(rows_move, rows_task, rows_personel_old, rows_personel_new, rows_tvo, d)


main()