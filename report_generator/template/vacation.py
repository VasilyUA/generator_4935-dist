import pandas as pd, os

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from helpers import convert_to_short_name, set_margins, set_line_spacing, add_paragraph_with_style, is_empty, print_green
from constants import OUTPUT_DIR, FULL_MILITARY_UNIT

def vacation(data_for_soldier, data_for_commander, data_for_higher_commander, extra_fields):
    doc = Document()
    set_margins(doc, Inches(0.786), Inches(0.786), Inches(1.18), Inches(0.395))
    set_line_spacing(doc, 14)  # встановлення міжрядкового інтервалу
    if pd.notna(data_for_commander.get('position_id')) and not is_empty(f"{data_for_commander.get('position_id')}"):
        add_paragraph_with_style(doc, data_for_commander.get('position_short_dative'), first_line_indent=250)
    else:
        add_paragraph_with_style(doc, "Командиру батальйону", first_line_indent=250)
    add_paragraph_with_style(doc, "")
    add_paragraph_with_style(doc, "")
    add_paragraph_with_style(doc, "")
    add_paragraph_with_style(doc, "РАПОРТ", alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_style(doc, "")
    add_paragraph_with_style(doc, "")
    address = f"{extra_fields.get('region_for_vacation', '')}, {extra_fields.get('district_for_vacation', '')}, {extra_fields.get('settlement_for_vacation', '')}, {extra_fields.get('street_for_vacation', '')}, {extra_fields.get('house_number_for_vacation', '')}".replace(", ,", ",")
    add_paragraph_with_style(doc, f"Прошу Вашого клопотання перед вищим командуванням про надання мені {extra_fields.get('part_vacation', '')} частини щорічної основної відпустки за {extra_fields.get('selected_year', '')} рік з {extra_fields.get('selected_day', '')} {extra_fields.get('selected_month', '')} {extra_fields.get('selected_year', '')} року терміном на {extra_fields.get('selected_days_for_vacation', '')} днів, із збереженням грошового забезпечення, у відповідності до Закону України №2822 від 01.12.2022 року «Про соціальний і правовий захист військовослужбовців та членів їх сімей».", first_line_indent=34)
    add_paragraph_with_style(doc, f"Відпустку буду проводити за адресою:{address}; мій номер телефону: {extra_fields.get('my_number_phone', '')}, телефон {extra_fields.get('relative_for_military', '')}: {extra_fields.get('relative_number_phone', '')}.", first_line_indent=34)
    add_paragraph_with_style(doc, "По заходам безпеки поводження громадських місцях проінструктований.", first_line_indent=34)
    add_paragraph_with_style(doc, "")
    add_paragraph_with_style(doc, f"{data_for_soldier.get('position_full_nominative', '')} {FULL_MILITARY_UNIT}", space_after=False)
    add_paragraph_with_style(doc, f"{data_for_soldier.get('rank_fact_nominative', '')}\t{convert_to_short_name(data_for_soldier.get('name_nominative', ''))}", format_tabs=True)
    add_paragraph_with_style(doc, "")
    add_paragraph_with_style(doc, "")
    if pd.notna(data_for_commander.get('position_id')) and not is_empty(f"{data_for_commander.get('position_id')}"):
        add_paragraph_with_style(doc, "Командиру батальйону", first_line_indent=250)
        add_paragraph_with_style(doc, "")
        add_paragraph_with_style(doc, f"Доповідаю по суті рапорту {data_for_soldier.get('rank_fact_genitive', '')} {convert_to_short_name(data_for_soldier.get('name_genitive', ''))}", first_line_indent=34)
        add_paragraph_with_style(doc, "")
        add_paragraph_with_style(doc, f"{data_for_commander.get('position_full_nominative', '')} {FULL_MILITARY_UNIT}", space_after=False)
        add_paragraph_with_style(doc, f"{data_for_commander.get('rank_fact_nominative', '')}\t{convert_to_short_name(data_for_commander.get('name_nominative', ''))}", format_tabs=True)
        add_paragraph_with_style(doc, "")
        add_paragraph_with_style(doc, "")
        add_paragraph_with_style(doc, f"Командиру {FULL_MILITARY_UNIT}", first_line_indent=250)
        add_paragraph_with_style(doc, "")
        add_paragraph_with_style(doc, f"Доповідаю по суті рапорту {data_for_commander.get('rank_fact_genitive', '')} {convert_to_short_name(data_for_commander.get('name_genitive', ''))}", first_line_indent=34)
    else:
        add_paragraph_with_style(doc, f"Командиру {FULL_MILITARY_UNIT}", first_line_indent=250)
        add_paragraph_with_style(doc, "")
        add_paragraph_with_style(doc, f"Доповідаю по суті рапорту {data_for_soldier.get('rank_fact_genitive', '')} {convert_to_short_name(data_for_soldier.get('name_genitive', ''))}", first_line_indent=34)
    add_paragraph_with_style(doc, "")
    add_paragraph_with_style(doc, "")
    add_paragraph_with_style(doc, f"{data_for_higher_commander.get('position_full_nominative', '')} {FULL_MILITARY_UNIT}", space_after=False)
    add_paragraph_with_style(doc, f"{data_for_higher_commander.get('rank_fact_nominative', '')}\t{convert_to_short_name(data_for_higher_commander.get('name_nominative', ''))}", format_tabs=True)
   
    doc_name = os.path.join(OUTPUT_DIR, f"{data_for_soldier.get('name_nominative')} здав посаду.docx")
    doc.save(doc_name)
    print_green(f"- {doc_name}")
    return doc_name
        