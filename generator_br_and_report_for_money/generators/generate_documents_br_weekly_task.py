import os

from constants import SHORT_UNIT_BATTALION, SHORT_UNIT_BRIGADE, NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK
from content.br_task_order_text import (
    BR_ENEMY_ASSESSMENT_TEXT, BR_FIRE_SUPPORT_INTRO_TEMPLATE,
    BR_TASK_SECTION_TEMPLATE, BR_SITUATION_UPDATE_TEMPLATE, BR_DEFENSE_TASKS_TEXT,
    BR_DEFENSE_MEASURES_TEXT, TASK_ORDER_VARIANTS,
)
from utils.logging_utils import print_green, print_red
from utils.date_utils import date_to_str, get_bat_period_and_variant
from content.br_helpers import show_coordinates
from formatting.docx_utils import (
    add_blank_paragraphs,
    add_body_paragraph,
    add_paragraph_with_style,
    add_recipients_preamble,
    add_section_heading,
    add_signature_block,
    add_titled_section,
    check_number_file_is_exist,
    generate_head_documents_br,
)


def generate_documents_br(doc, col_name, higher_commander_data, city, coordinates, output_dir_br, is_extract_br=False):
    date_str = date_to_str(col_name)
    num__doc = NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK.get(date_str, {}).get('бат') or None

    generate_head_documents_br(doc)
    generate_basic_header_br_documents(doc, col_name, num__doc, is_extract_br, city, coordinates)
    generate_order_section(doc, col_name)
    generate_footer_br(doc, higher_commander_data, output_dir_br, num__doc, date_str)


def generate_basic_header_br_documents(doc, col_name, num__doc, is_extract_br, city, coordinates):
    date_str = date_to_str(col_name)
    unit = f"{SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE}"
    prefix = check_number_file_is_exist(num__doc)

    # Посилання на попередні розпорядження (брг/бат) для розділу 3 — заповнюються вручну
    # в constants.py (поля 'посилання_брг'/'посилання_бат' у NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK),
    # одразу в готовому вигляді "номер від дата", як і 'general_br_bat'.
    order_entry = NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK.get(date_str, {})
    order_refs = {
        'brg_ref': order_entry.get('посилання_брг') or '___',
        'bat_ref': order_entry.get('посилання_бат') or '___',
    }

    add_recipients_preamble(doc)

    doc_type = "ВИТЯГ з БОЙОВОГО" if is_extract_br else "БОЙОВЕ"
    add_body_paragraph(
        doc,
        f"{doc_type} РОЗПОРЯДЖЕННЯ КОМАНДИРА {unit} {prefix}КСП {city} "
        f"({show_coordinates(is_extract_br, coordinates)}), 06.00 {date_str}. Карта 1:25000 видання 2024 року."
    )
    add_blank_paragraphs(doc)

    add_titled_section(doc, "1. ВИСНОВКИ З ОЦІНЮВАННЯ ПРОТИВНИКА ТА ЙМОВІРНИЙ ХАРАКТЕР ЙОГО ДІЙ", BR_ENEMY_ASSESSMENT_TEXT.format(brigade=SHORT_UNIT_BRIGADE))

    add_section_heading(doc, "2. ЗАВДАННЯ, ЩО ВИКОНУЮТЬСЯ СИЛАМИ І ЗАСОБАМИ СТАРШОГО КОМАНДИРА НА НАПРЯМКУ ДІЙ БАТАЛЬЙОНУ")
    add_body_paragraph(doc, BR_FIRE_SUPPORT_INTRO_TEMPLATE.format(unit=unit))
    add_blank_paragraphs(doc)

    add_section_heading(doc, "3. БОЙОВЕ ЗАВДАННЯ.")
    add_body_paragraph(doc, BR_TASK_SECTION_TEMPLATE.format(unit=unit, brigade=SHORT_UNIT_BRIGADE, **order_refs))
    add_body_paragraph(doc, BR_SITUATION_UPDATE_TEMPLATE)
    add_body_paragraph(doc, BR_DEFENSE_TASKS_TEXT)
    add_body_paragraph(doc, BR_DEFENSE_MEASURES_TEXT)


def generate_order_section(doc, col_name):
    """Розділ 'НАКАЗАВ:' — перелік завдань підрозділам. Період і те, який із трьох
    варіантів переліку (TASK_ORDER_VARIANTS) використати, визначаються за ФАКТИЧНИМИ
    датами БР бат (NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK, див. get_bat_period_and_variant) - col_name
    МАЄ мати непустий 'бат' (генерація "ЗАВДАННЯ" відбувається лише для таких дат,
    process_generate_br.py сам перевіряє цю умову перед викликом)."""
    period_text, variant_index = get_bat_period_and_variant(col_name, NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK)

    add_body_paragraph(doc, f"Для виконання вищевказаних завдань в період {period_text}, НАКАЗАВ:")
    add_blank_paragraphs(doc)

    tasks = TASK_ORDER_VARIANTS[variant_index] if variant_index < len(TASK_ORDER_VARIANTS) else []
    if not tasks:
        print(
            f"Варіант завдань №{variant_index + 1} (TASK_ORDER_VARIANTS[{variant_index}]) ще не заповнено "
            f"в constants.py — тимчасово використано варіант №1."
        )
        tasks = TASK_ORDER_VARIANTS[0]

    for role, text in tasks:
        add_paragraph_with_style(doc, f"{role}. {text}", first_line_indent=34, bold=[role])
        add_blank_paragraphs(doc)
    
    add_titled_section(doc, "4. ОРГАНІЗАЦІЯ УПРАВЛІННЯ", "Без змін.")
    add_titled_section(doc, "5. ЧАС ГОТОВНОСТІ ДО ВИКОНАННЯ ЗАВДАНЬ", "З отриманням розпорядження.")
    add_titled_section(doc, "6. ТЕРМІН ВИКОНАННЯ ЗАВДАННЯ", "До окремого розпорядження.", blanks=2)


def generate_footer_br(doc, higher_commander_data, output_dir_br, num__doc, date_str):
    role = "ТВО командира" if higher_commander_data.get('ТВО', "") else "Командир"
    add_signature_block(doc, role, higher_commander_data)

    prefix = check_number_file_is_exist(num__doc)
    file_name = f"{prefix}ЗАВДАННЯ {date_str}.docx"
    doc.save(os.path.join(output_dir_br, file_name))
    print_green(file_name)
