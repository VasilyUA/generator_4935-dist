import os
from constants import OUTPUT_DIR_BR, OUTPUT_DIR_EXTRACT_BR, SHORT_UNIT_BRIGADE, NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY, PARAGRAPH_MAP, SHORT_UNIT_BATTALION
from utils.logging_utils import print_green
from utils.date_utils import date_to_str
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

def generate_documents_br_general(doc, col_name, rows_with_data, section_lists, higher_commander_data, city, coordinates, commander_data={}, is_extract_br=False):
    num__doc = NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY.get(date_to_str(col_name), {})
    num_bat = num__doc.get('бат', '___')
    num_brg = num__doc.get('брг', '___')
    num_general_brg = num__doc.get('general_br_bat', '___')
    if not is_extract_br:
        generate_head_documents_br(doc)
    generate_basic_header_br_documents(doc, col_name, num_bat, num_brg, num_general_brg, city, coordinates, is_extract_br)
    for section, data in section_lists.items():
        if data: generate_content_br(doc, section, data, city, coordinates)
    generate_footer_br(doc, col_name, higher_commander_data, num_bat, commander_data, is_extract_br)

def generate_basic_header_br_documents(doc, col_name, num__doc='___', num_brg='___', general_br_bat='___', city="", coordinates="", is_extract_br=False):
    date_str = date_to_str(col_name)
    add_recipients_preamble(doc)
    doc_type = "ВИТЯГ З БОЙОВОГО" if is_extract_br else "БОЙОВЕ"
    add_body_paragraph(doc, f"{doc_type} РОЗПОРЯДЖЕННЯ КОМАНДИРА {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №{num__doc} КСП {city} ({coordinates}), 06.00 {date_str}. Карта 25000 видання 2024 року.")
    add_blank_paragraphs(doc)
    add_titled_section(doc, "1. ВИСНОВКИ З ОЦІНЮВАННЯ ПРОТИВНИКА.", "Згідно розвідувальних відомостей, розвідувальна інформація, та розвідувальних даних що надходять.")
    add_titled_section(doc, "2. ЗАВДАННЯ, ЩО ВИКОНУЮТЬСЯ СИЛАМИ І ЗАСОБАМИ СТАРШОГО КОМАНДИРА НА НАПРЯМКУ ДІЙ БАТАЛЬЙОНУ.", "За викликом засобами старшого начальника уражаються цілі відповідно до таблиці вогню артилерії згідно плану.")
    add_section_heading(doc, "3. БОЙОВЕ ЗАВДАННЯ БАТАЛЬЙОНУ")
    add_body_paragraph(doc, f"На виконання бойового розпорядження {SHORT_UNIT_BRIGADE} №{num_brg} від {date_to_str(col_name, action='-', days=1)} та продовження виконання бойового розпорядження командира {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №{general_br_bat} з метою недопущення просочування чи вклинення противника в оборонні порядки, а також підтримання готовності підрозділів до відбиття можливих атак, забезпечення своєчасної готовністю до реагування на зміни в оперативній обстановці, для нарощування сил і засобів на ПУ, ВП, СП, ВОП, РОП у межах БРО та підвищення ефективності виконання бойових (спеціальних) завдань, НАКАЗАВ:")

def generate_content_br(doc, section, data, city="", coordinates=""):
    section_key = ", ".join(section) if isinstance(section, (list, tuple)) else section
    if section_key not in PARAGRAPH_MAP:
        return
    text = PARAGRAPH_MAP[section_key].format(city=city, coordinates=coordinates, battalion=SHORT_UNIT_BATTALION, unit=f"{SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE}")
    add_paragraph_with_style(doc, text, font_size=14, bold=[section_key], first_line_indent=34)
    for row in data: add_body_paragraph(doc, f"{row.get('ЗВАННЯ', '')} {row.get('ПІБ', '')};")
    add_blank_paragraphs(doc)

def generate_footer_br(doc, col_name, higher_commander_data, num__doc='___', commander_data={}, is_extract_br=False):
    add_titled_section(doc, "4. ОРГАНІЗАЦІЯ УПРАВЛІННЯ", "Без змін.")
    add_titled_section(doc, "5. ЧАС ГОТОВНОСТІ ДО ДІЙ", "З отримання даного розпорядження.")
    add_titled_section(doc, "6. ТЕРМІН ВИКОНАННЯ ЗАВДАННЯ", "До окремого розпорядження.", blanks=2)
    role = "ТВО командира" if higher_commander_data.get('ТВО', "") else "Командир"
    add_signature_block(doc, role, higher_commander_data)
    if is_extract_br:
        add_blank_paragraphs(doc, 2)
        add_paragraph_with_style(doc, "Згідно з оригіналом:")
        role = "ТВО начальника штабу - заступник командира" if commander_data.get('ТВО', "") else "Начальник штабу - заступник командира"
        add_signature_block(doc, role, commander_data)
    prefix = check_number_file_is_exist(num__doc)
    file_name = f"{'Витяг ' if is_extract_br else ''}{prefix}ЩОДЕННА {date_to_str(col_name)}.docx"
    br_dir = OUTPUT_DIR_EXTRACT_BR if is_extract_br else OUTPUT_DIR_BR
    doc.save(os.path.join(br_dir, file_name))
    print_green(file_name)
