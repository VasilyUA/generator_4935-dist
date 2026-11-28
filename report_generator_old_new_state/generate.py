from template.change_position import change_position_report
from template.handed_position import handed_position_report
from template.accepted_position import generate_accepted_position_report
from helpers import (
    print_green, 
    get_position_code,
    print_red, 
    get_subordinate,
)
from validator import DataValidation

def generate(rows_move, rows_task, rows_personel_old, rows_personel_new, rows_tvo, date):
    # # #
    # # # Document 1
    # # #
    doc_name_1 = change_position_report(rows_move, rows_personel_old, rows_personel_new, rows_task, rows_tvo, date)
    print_green(f"- {doc_name_1}")

    for row in rows_move:
        position_codes = get_position_code(row)

        validator = DataValidation({
            "position_codes": position_codes,
            "rows_personel_old": rows_personel_old,
            "rows_personel_new": rows_personel_new,
            "rows_task": rows_task,
            "rows_tvo": rows_tvo
        })
        is_valid = validator.is_valid()

        if not is_valid:
            print_red(f"Валідація не пройшла для даних: {position_codes}")
            continue

        # структурування даних
        data_for_generation = get_subordinate(position_codes, rows_personel_old, rows_personel_new, rows_task, rows_tvo, date)

        # # #
        # # # Document 2 здав посаду
        # # #
        doc_name_2 = handed_position_report(data_for_generation, position_codes['commander_position_code_in_old_state'], rows_tvo)
        print_green(f"- {doc_name_2}")

        # # #
        # # # Document 3 прийняв посаду
        # # #
        doc_name_3 = generate_accepted_position_report(data_for_generation, position_codes['commander_position_code_in_new_state'], rows_tvo)
        print_green(f"- {doc_name_3}")
    pass
    

    

    

    
 
    
    