import copy
from utils.date_utils import date_to_str, get_prev_general
from content.br_general_catalog import get_catalog_log_general_extract_log_war, get_content_log_general_extract_log_war
from content.br_helpers import get_number_br
from formatting.docx_utils import add_combat_log_extract_table, save_combat_log_extract_war, setup_combat_log_extract_document
from constants import SHORT_UNIT_BRIGADE, PARAGRAPH_MAP, PARAGRAPH_MAP_FOR_EXTRACT, NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY, SHORT_UNIT_BATTALION

def generate_extract_log_war_general_br_every_day(doc, col_name, commander_data, section_lists, previous_section_lists, city="", coordinates=""):
    safe_col_name = date_to_str(col_name)
    setup_combat_log_extract_document(doc, safe_col_name)
    catalog = get_content_log_general_extract_log_war(section_lists, copy.deepcopy(PARAGRAPH_MAP), city=city, coordinates=coordinates)
    content_log = get_catalog_log_general_extract_log_war(previous_section_lists, copy.deepcopy(PARAGRAPH_MAP_FOR_EXTRACT), city=city, coordinates=coordinates)
    generate_main_combat_log_extract_war_every_day(doc, safe_col_name, catalog, content_log)
    save_combat_log_extract_war(doc, safe_col_name, commander_data, NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY, " ЩОДЕННА")

def generate_main_combat_log_extract_war_every_day(doc, safe_col_name, content, content_log):
    """Пункт 5 - СЬОГОДНІШНЄ щоденне бойове розпорядження (num_bat за safe_col_name) -
    те, що ще належить виконати. Пункт 6 - звіт про виконання ПОПЕРЕДНЬОГО дня
    (get_prev_general - справді попередній запис NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY,
    СТРОГО раніший за safe_col_name, а не сам safe_col_name), уже у виконаному часі:
    для 01.07 (перший день обліку) get_prev_general поверне None (попереднього
    щоденного розпорядження просто не існує) - пункт 6 тоді лишається без
    посилання на конкретне розпорядження ("..."); з 02.07 і далі - завжди
    посилається на РЕАЛЬНИЙ вчорашній запис (за 01.07 для 02.07, за 02.07 для
    03.07 і т.д.)."""
    date_for_prev_general = get_prev_general(safe_col_name, NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY)
    num_bat = get_number_br(safe_col_name, NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY)

    if date_for_prev_general is None:
        execution_report = "...\n"
    else:
        prev_num_bat = get_number_br(date_for_prev_general, NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY)
        execution_report = (
            f"{date_to_str(safe_col_name)} відповідно до бойового розпорядження командира {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} "
            f"№{prev_num_bat.get('бат', '')} від {date_for_prev_general} підрозділи батальйону здійснили:\n"
            + content_log
            + "\n..."
        )

    text = (
        "...\n"
        "4. Коротке викладення отриманого бойового завдання (розпорядження) (у тому числі його номер та дата).\n"
        f"На виконання бойового розпорядження командира {SHORT_UNIT_BRIGADE} №{num_bat.get('брг')} від {date_to_str(safe_col_name, action='-', days=1)} продовжити ведення стійкої позиційної оборони батальйонного району оборони, із завданням не допустити висадки противника, стійко обороняти займані позиції, нанести йому вогневе ураження та перешкодити просуванню в глибину оборони."
        "\n...\n"
        f"5. Рішення командира ({SHORT_UNIT_BATTALION}) військової частини та бойові завдання, поставлені підпорядкованим частинам (підрозділам).\n"
        "...\n"
        f"На виконання бойового розпорядження командира {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №{num_bat.get('бат')} від {safe_col_name}, здійснити:\n"
        + content
        + "\n...\n"
        "6. Відомості про виконання бойового (спеціального) завдання, ведення бойових дій (бою), у тому числі за придані (підтримуючі) підрозділи.\n"
        "...\n"
        + execution_report
    )
    add_combat_log_extract_table(doc, safe_col_name, text)
