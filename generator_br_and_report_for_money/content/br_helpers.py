import re
from datetime import datetime

from constants import NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK
from utils.date_utils import to_date, date_to_str
from content.money_report_helpers import _clean_pib_spacing


def convert_to_short_name(full_name, id_position=None):
    if not isinstance(full_name, str):
        return full_name

    # 1-2. Виправляємо апостроф + пробіл (В' Ярош -> В'Ярош) і зайві пробіли.
    cleaned = _clean_pib_spacing(full_name)

    # 3. Розбити на слова
    words = cleaned.split()

    # 4. Якщо менше ніж 3 — пробуємо розклеїти друге слово (ім’я+по батькові)
    if len(words) == 2:
        # Пробуємо вставити пробіл між ім’ям і по батькові
        possible_fix = re.sub(r'([а-яґєіїʼ’]+)([А-ЯІЇЄҐ])', r'\1 \2', words[1])
        fixed_words = [words[0]] + possible_fix.split()
        if len(fixed_words) == 3:
            words = fixed_words

    # 5. Перевірка
    if len(words) != 3:
        raise ValueError(f"❌ Некоректний ПІБ для {id_position}: \"{full_name}\" → після обробки: \"{' '.join(words)}\"")

    # 6. Повертаємо скорочене ім’я
    return f"{words[1]} {words[0]}"


def find_higher_commander(rows_tvo, rows, pos_title, col_name):
    date_col = to_date(col_name)

    commander = next(
        (r for r in rows_tvo
         if r['ПОСАДА'] == pos_title and r.get('ТВО') is True
         and to_date(r['Start']) <= date_col <= to_date(r['End'])),
        None
    )

    if not commander:
        commander = next(
            (r for r in rows if r['ПОСАДА'] == pos_title),
            {}
        )

    return commander


def show_coordinates(is_extract_br, coordinates):
    return re.sub(r'(36T TT )(\d{2})\d{3}\s(\d{2})\d{3}', r'\1\2*** \3***', coordinates) if is_extract_br else coordinates


def get_number_br(col_name, number_objects=NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK):
    if isinstance(col_name, str):
        # пробуємо розпарсити рядок як дату
        try:
            col_name = datetime.strptime(col_name, "%d.%m.%Y")
        except ValueError:
            # fallback: якщо формат інший, повертаємо пусте
            return {'number_documents': {}, 'брг': '___', 'бат': '___'}

    num__doc = number_objects.get(date_to_str(col_name), {})
    num_brg = num__doc.get('брг', '___')
    num_bat = num__doc.get('бат', '___')

    return {'number_documents': num__doc, 'брг': num_brg, 'бат': num_bat}
