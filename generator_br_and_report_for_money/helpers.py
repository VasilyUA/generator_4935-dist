# Цей модуль лишається для зворотної сумісності (стара точка входу
# `from helpers import ...` / `import helpers`, якою користуються тести).
# Сама логіка тепер живе в окремих пакетах за відповідальністю:
#   utils/       — логування, дати, читання Excel, валідація вводу, вибір папки
#   formatting/  — низькорівневе форматування python-docx
#   content/     — генерація тексту БР/журналу бойових дій
#   generators/  — побудова конкретних Word-документів
#   sync/        — синхронізація номерів БАТ/БЗ у constants.py
from utils.logging_utils import print_red, print_green

from utils.date_utils import (
    to_date,
    date_to_str,
    get_bat_period_and_variant,
    get_prev_general,
)

from constants import SUPPORTED_EXTENSIONS

from utils.excel_reader import (
    find_file_with_any_extension,
    read_datafile,
    read_optional_datafile,
    get_data,
    excel_col_to_index,
)

from utils.validation import (
    validate_battalion_commander_exists,
    get_validated_ksp_data,
)

from content.br_helpers import (
    convert_to_short_name,
    find_higher_commander,
    show_coordinates,
    get_number_br,
)

from content.br_general_catalog import (
    generate_content_br_general,
    get_unique_sections_list,
    get_content_log_general_extract_log_war,
    get_catalog_log_general_extract_log_war,
)

from formatting.docx_utils import (
    set_line_spacing,
    create_or_clear_output_directory,
    add_paragraph_with_style,
    add_custom_heading,
    add_body_paragraph,
    add_section_heading,
    add_blank_paragraphs,
    add_signature_block,
    add_titled_section,
    generate_combat_log_extract_title,
    check_number_file_is_exist,
    save_combat_log_extract_war,
    set_margins,
    add_text_for_header_and_footer_document_br,
    add_page_number_header,
    switch_landscape,
    create_table,
)
