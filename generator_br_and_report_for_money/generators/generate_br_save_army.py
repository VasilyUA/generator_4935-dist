import os
from constants import NUMBER_OF_DOCUMENTS_BRS_SAVE, SEQUENCE, SHORT_UNIT_BATTALION, SHORT_UNIT_BRIGADE, SEQUENCE_HIGHER_COMMANDER
from utils.logging_utils import print_green
from utils.date_utils import date_to_str
from formatting.docx_utils import (
    add_blank_paragraphs,
    add_body_paragraph,
    add_recipients_preamble,
    add_section_heading,
    add_signature_block,
    add_titled_section,
    check_number_file_is_exist,
    generate_head_documents_br,
)

HIGHER_COMMANDERS = ', '.join(SEQUENCE_HIGHER_COMMANDER)
UNIT_COMMANDERS = ', '.join(SEQUENCE)

SAFETY_MEASURES = [
    "використання систем розпізнавання (паролів «свій-чужий»);",
    "дотримання режимів радіомовчання;",
    "недопущення скупчення особового складу та техніки в районах відновлення та відпочинку;",
    "здійснювати передачу інформації виключно з криптографічно захищених радіостанцій, закритих цифрових каналів передачі даних;",
    "заборона передачі координат, наказів та персональних даних через незахищені месенджери або відкритий радіоефір;",
    "ротації, підвіз БК та евакуацію здійснювати в темну пору доби або під час складних погодних умов (туман, сильний дощ);",
    "під час пересування пішим порядком відстань між військовослужбовцями має становити не менше 10 м;",
    "при переміщенні до позицій використовувати закриті зеленкою або рельєфом маршрути, уникати пересування відкритими дорогами та стежками, які проглядаються з повітря;",
    "райони видачі провізії не розміщувати біля орієнтирів, уникати одиночних будинків, перехресть, розв'язок, натомість використовувати підвали, заглиблені споруди;",
    "підрозділи прибувають строго по графіку для отримання провізії, БК, ПММ з часовим розривом у 20-30 хв;",
    "оповіщення підрозділів (у тому числі сил безпеки оборони та органів державної влади, військових адміністрацій, органів місцевого самоврядування) у визначених районах, про наміри противника нанесення ураження;",
    "забезпечити маскування районів розгортання (зосередження) військ (сил), усунення демаскуючих ознак в місцях розгортання пунктів управління, позиційних районах підрозділів;",
    "здійснювати здійснення спільно з підрозділами військової контррозвідки Служби безпеки України та Національної поліції України заходів щодо виявлення та закриття каналів витоку інформації стосовно стану, складу, положення та ймовірного характеру дій підрозділів підпорядкованих підрозділів;",
    "посилити контроль за виконанням заходів безпеки застосування підпорядкованих підрозділів.",
]

def generate_documents_br_save(doc, col_name, higher_commander_data, city, coordinates, output_dir_br):
    date_str = date_to_str(col_name)
    num__doc = NUMBER_OF_DOCUMENTS_BRS_SAVE.get(date_str, {})
    num_bat = num__doc.get('бз_бат', None)
    num_brg = num__doc.get('бз_брг', None)

    if not num_bat or not num_brg:
        print(f"Відсутній номер бойового розпорядження застосування безпеки на дату {date_str}. Пропускаємо генерацію документа.")
        return

    generate_head_documents_br(doc)
    generate_basic_header_br_documents(doc, col_name, num_bat, num_brg, city, coordinates)
    generate_main_content(doc)
    generate_footer_br(doc, col_name, higher_commander_data, output_dir_br, num_bat)

def generate_basic_header_br_documents(doc, col_name, num__doc='___', num_brg='___', city="", coordinates=""):
    date_str = date_to_str(col_name)

    add_recipients_preamble(doc)
    add_body_paragraph(doc, f"РОЗПОРЯДЖЕННЯ з безпеки застосування військ (сил) {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №{num__doc} КСП {city} ({coordinates}), 06.00 {date_str}. Карта 25000 видання 2024 року.")
    add_blank_paragraphs(doc)
    add_titled_section(doc, "1. ВИСНОВКИ З ОЦІНЮВАННЯ ПРОТИВНИКА ТА ЙМОВІРНИЙ ХАРАКТЕР ЙОГО ДІЙ", "Згідно розвідувальних відомостей, розвідувальна інформація, та розвідувальних даних що надходять.")
    add_titled_section(doc, "2. ЗАВДАННЯ, ЩО ВИКОНУЮТЬСЯ СИЛАМИ І ЗАСОБАМИ СТАРШОГО КОМАНДИРА НА НАПРЯМКУ ДІЙ БАТАЛЬЙОНУ", "За викликом засобами старшого начальника уражаються цілі відповідно до таблиці вогню артилерії згідно плану.")
    add_section_heading(doc, "3. ЗАВДАННЯ БЕЗПЕКИ ЗАСТОСУВАННЯ ВІЙСЬК (СИЛ) ПІДРОЗДІЛАМ.")
    add_body_paragraph(doc, f"На виконання розпорядження з безпеки застосування підрозділів командира {SHORT_UNIT_BRIGADE} №{num_brg} з метою забезпечення належного морально-психологічного і фізичного стану кожного військовослужбовця. Збереження життя та здоров'я. Висування і переміщення особового складу дозволено виключно в темну пору доби або в умовах недостатньої видимості, з обов’язковим дотриманням заходів прихованості, безпеки застосування сил і засобів, а також скритого управління військами, дотримуватись порядку використання засобів індивідуального захисту НАКАЗУЮ:")

def generate_main_content(doc):
    add_body_paragraph(doc, HIGHER_COMMANDERS)
    add_body_paragraph(doc, f"командиру {UNIT_COMMANDERS} вжити невідкладні заходи щодо забезпечення живучості військ (сил), військових об'єктів, збереження функціонування критичних об'єктів інфраструктур, здійснювати та забезпечити:")

    for measure in SAFETY_MEASURES:
        add_body_paragraph(doc, measure)

def generate_footer_br(doc, col_name, higher_commander_data, output_dir_br, num__doc='___'):
    add_blank_paragraphs(doc)
    add_titled_section(doc, "4. ЧАС ГОТОВНОСТІ ДО ДІЙ", "Готовність до виконання завдань – з отримання даного розпорядження.", blanks=2)

    title_role = "ТВО командира" if higher_commander_data.get('ТВО', "") else "Командир"
    add_signature_block(doc, title_role, higher_commander_data)

    file_prefix = check_number_file_is_exist(num__doc)
    file_name = f"{file_prefix}застосування безпеки {date_to_str(col_name)}.docx"
    doc.save(os.path.join(output_dir_br, file_name))
    print_green(file_name)
