import pandas as pd
import warnings
import helpers
import constants as c
import template
from movement_generate import mass_generate

# Пригнічуємо зайві попередження
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", category=FutureWarning)

def main():
    helpers.create_or_clear_output_directory(c.OUTPUT_DIR)

    print("Reading data...")
    rows_personel = helpers.read_excel_rows(c.PERSONEL_LIST_FILE_NAME, c.PERSONEL_LIST_SHEET_NAME, c.PERSONEL_LIST_COLUMNS_LETTERS)
    rows_task = helpers.read_excel_rows(c.TASK_LIST_FILE_NAME, c.TASK_LIST_SHEET_NAME, c.TASK_LIST_FILE_COLUMNS_LETTERS)
    action, args = helpers.get_value_data(rows_personel, rows_task)

    # Карта дій: ключ — константа, значення — функція
    handlers = {
        c.VALUE_RESIGNATION: template.handed_position_report,
        c.VALUE_NEW_POSITION: template.generate_accepted_position_report,
        c.VALUE_MILITARY_ASSAULT_COURSE: template.military_assault_course,
        c.VALUE_SANITATION: template.sanitation,
        c.VALUE_VACATION: template.vacation,
    }

    if action == c.VALUE_MASS_MOVEMENT:
        print("Processing mass movement...")
        rows_move = helpers.read_excel_rows(c.TRANSFER_LIST_FILE_NAME, c.TRANSFER_LIST_SHEET_NAME, c.TRANSFER_LIST_FILE_COLUMNS_LETTERS)
        mass_generate(rows_move, rows_task, rows_personel)
    elif action in handlers:
        handlers[action](*args)
    else:
        print(f"Unknown action: {action}")

if __name__ == "__main__":
    main()
