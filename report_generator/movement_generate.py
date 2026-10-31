from docx import Document
from template.change_position import change_position_report
from template.handed_position import handed_position_report
from template.accepted_position import generate_accepted_position_report
from helpers import get_date_for_military

def mass_generate(rows_move, rows_task, rows_personel):
    doc = Document()

    for index, row in enumerate(rows_move, 1):
        len_row=len(rows_move)
        data_for_soldier_from = get_date_for_military(rows_personel, rows_task, row.get('НОМЕР ПОСАДИ З ЯКОЇ ПЕРЕМІЩУЄТЬСЯ ОСОБА'))
        data_for_soldier_to = get_date_for_military(rows_personel, rows_task, row.get('НОМЕР ПОСАДИ НА ЯКУ ПЕРЕМІЩУЄТЬСЯ ОСОБА'))
        data_for_commander = get_date_for_military(rows_personel, rows_task, row.get('КОМАНДИР ВЗВОДУ ЧИ РОТИ АБО ЇХ ТВО ЯКІ КЛОПОЧУТЬ КОМАНДИРУ БАТАЛЬЙОНУ'))
        data_for_higher_commander = get_date_for_military(rows_personel, rows_task, row.get('КОМАНДИР БАТАЛЬЙОНУ АБО ЙОГО ТВО'))
    
        #
        # Document 1
        #
        change_position_report(doc, data_for_soldier_from, data_for_soldier_to, data_for_higher_commander, index, len_row)

        # #
        # # Document 2 здав посаду
        # #
        handed_position_report(data_for_soldier_from, data_for_commander, data_for_higher_commander)

        # #
        # # Document 3 прийняв посаду
        # #
        data_for_soldier_to.update({ k: data_for_soldier_from[k] for k in ['rank_fact_nominative', 'rank_fact_genitive', 'name_nominative', 'name_genitive'] })
        generate_accepted_position_report(data_for_soldier_to, data_for_commander, data_for_higher_commander)

    



    