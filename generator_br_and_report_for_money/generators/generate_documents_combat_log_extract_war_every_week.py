from constants import SHORT_UNIT_BATTALION, SHORT_UNIT_BRIGADE, NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK
from utils.date_utils import date_to_str, get_bat_period_and_variant, get_prev_general
from content.br_helpers import get_number_br
from content.br_task_order_text import BR_TASK_SECTION_TEMPLATE, TASK_ORDER_VARIANTS
from formatting.docx_utils import add_combat_log_extract_table, save_combat_log_extract_war, setup_combat_log_extract_document

def generate_documents_combat_log_extract_war_every_week(doc, col_name, commander_data):
    safe_col_name = date_to_str(col_name)
    setup_combat_log_extract_document(doc, safe_col_name)
    generate_main_combat_log_extract_war_every_week(doc, safe_col_name)
    save_combat_log_extract_war(doc, safe_col_name, commander_data, NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK, " ЩОТИЖНЕВА")

def _tasks_text_for(variant_index):
    tasks = TASK_ORDER_VARIANTS[variant_index] if variant_index < len(TASK_ORDER_VARIANTS) else []
    if not tasks:
        tasks = TASK_ORDER_VARIANTS[0]
    return "\n".join(f"{role}. {text}" for role, text in tasks)

def generate_main_combat_log_extract_war_every_week(doc, safe_col_name):
    """Уся інформація пункту 5 береться з того самого джерела, що й документ
    ЗАВДАННЯ (generate_documents_br_weekly_task.py): пункт 3 (BR_TASK_SECTION_TEMPLATE)
    і перелік завдань підрозділам після НАКАЗАВ (TASK_ORDER_VARIANTS), без ПІБ
    особового складу. Період і варіант переліку - за ФАКТИЧНИМИ датами БР бат
    (див. get_bat_period_and_variant); safe_col_name МАЄ мати непустий 'бат'
    (process_generate_br.py сам перевіряє цю умову перед викликом).

    Пункт 6 (відомості про виконання) - так само, як і в щоденному витягу
    (generate_extract_log_war_general_br_every_day.py) - звітує про ВЖЕ
    ЗАВЕРШЕНЕ ПОПЕРЕДНЄ тижневе розпорядження (get_prev_general - справді
    попередній запис NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK, СТРОГО раніший за
    safe_col_name), а НЕ те саме розпорядження, що й пункт 5: для 01.07 (першого
    тижневого БР місяця) попереднього просто не існує - пункт 6 лишається без
    посилання ("..."); коли з'являється НАСТУПНИЙ тижневий БР (напр. 07.07),
    його документ пункт 6 звітує про виконання розпорядження від 01.07."""
    unit = f"{SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE}"
    num_bat = get_number_br(safe_col_name)

    order_entry = NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK.get(safe_col_name, {})
    order_refs = {
        'brg_ref': order_entry.get('посилання_брг') or '___',
        'bat_ref': order_entry.get('посилання_бат') or '___',
    }

    period_text, variant_index = get_bat_period_and_variant(safe_col_name, NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK)
    tasks_text = _tasks_text_for(variant_index)

    date_for_prev_bat = get_prev_general(safe_col_name, NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK)
    if date_for_prev_bat is None:
        execution_report = "...\n"
    else:
        prev_num_bat = get_number_br(date_for_prev_bat)
        prev_period_text, prev_variant_index = get_bat_period_and_variant(date_for_prev_bat, NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK)
        execution_report = (
            f"{safe_col_name} відповідно до бойового розпорядження командира {unit} №{prev_num_bat.get('бат')} від {date_for_prev_bat}, "
            f"за період {prev_period_text}, підрозділи здійснили:\n"
            + _tasks_text_for(prev_variant_index)
            + "\n..."
        )

    text = (
        "...\n"
        "4. Коротке викладення отриманого бойового завдання (розпорядження) (у тому числі його номер та дата).\n"
        f"{BR_TASK_SECTION_TEMPLATE.format(unit=unit, brigade=SHORT_UNIT_BRIGADE, **order_refs)}"
        "\n...\n"
        f"5. Рішення командира військової частини ({unit}) та бойові завдання, поставлені підпорядкованим частинам (підрозділам).\n"
        "...\n"
        f"На виконання бойового розпорядження командира {unit} №{num_bat.get('бат')} від {safe_col_name}, для виконання завдань в період {period_text}, здійснити:\n"
        + tasks_text
        + "\n...\n"
        "6. Відомості про виконання бойового (спеціального) завдання, ведення бойових дій (бою), у тому числі за придані (підтримуючі) підрозділи.\n"
        "...\n"
        + execution_report
    )
    add_combat_log_extract_table(doc, safe_col_name, text)
