import os
import re
import tempfile
import zipfile
from collections import Counter
from datetime import datetime, timedelta

from utils.logging_utils import print_red, print_green, print_progress
from utils.project_settings import (
    NUMBERS_SHEET, day_labels, read_numbers_workbook, write_number_fields, write_numbers_workbook,
)
from utils.settings_book import PREVIOUS_MONTH_DAY, SettingsError
# extract_document_text (content.document_content_search) імпортується ЛОКАЛЬНО,
# нижче, у місці виклику (sync_brigade_document_numbers_from_folder) - той самий
# модуль сам імпортує constants.py, а constants.py (через sync.stage_folder)
# імпортує ЦЕЙ модуль ще ДО того, як constants.py встигає визначити свої власні
# константи (SHORT_UNIT_BATTALION/BRIGADE) - імпорт тут, нагорі файлу, створив би
# циклічний імпорт (той самий випадок, що вже пояснено в самому constants.py).

# "№117 ЩОДЕННА 13.07.2026.docx" -> номер 117 підставляється в
# NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY[дата]['бат']
_DAILY_BR_FILENAME_RE = re.compile(r"(?:Витяг\s+)?№(\d+)\s*ЩОДЕННА\s+(\d{2}\.\d{2}\.\d{4})\.docx", re.IGNORECASE)
# "№91 застосування безпеки 02.07.2026.docx" -> номер 91 підставляється в
# NUMBER_OF_DOCUMENTS_BRS_SAVE[дата]['бз_бат']
_SAFETY_BR_FILENAME_RE = re.compile(r"№(\d+)\s*застосування безпеки\s+(\d{2}\.\d{2}\.\d{4})\.docx", re.IGNORECASE)
# "№87 ЗАВДАННЯ 01.07.2026.pdf" -> номер 87 підставляється в
# NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK[дата]['бат'] (те саме поле, що вже там було — просто досі не заповнювалось)
_TASK_ORDER_FILENAME_RE = re.compile(r"№(\d+)\s*ЗАВДАННЯ\s+(\d{2}\.\d{2}\.\d{4})\.(?:docx|pdf)", re.IGNORECASE)


def _derive_general_br_bat(day_keys, week_dict):
    """NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY[дата]['general_br_bat'] - те саме тижневе
    розпорядження ("ЗАВДАННЯ"), що діє на цю дату: номер+дата НАЙПІЗНІШОГО запису
    NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK['бат'], чия дата не пізніша за саму цю дату
    (тижневе розпорядження лишається чинним, поки не з'явиться наступне - тож
    ДЕНЬ МІЖ двома тижневими датами так само посилається на ОСТАННЄ видане, а не
    порожнє значення). '', якщо жодного заповненого тижневого запису ще не було
    (напр. початок місяця до першого "ЗАВДАННЯ").

    Раніше це поле ніде автоматично не рахувалось - лишалось тим, що людина
    востаннє вписала вручну, навіть коли sync_document_numbers_from_folder нижче
    оновлював саме NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK з нових файлів "ЗАВДАННЯ" в
    обраній папці (підтверджено користувачем: general_br_bat має оновлюватись
    РАЗОМ з ним, номери в обох - одні й ті самі).

    РЕГРЕСІЯ (реальний випадок користувача, вересень 2026): колись ключі
    словників номерів були рядками constants.py, які людина сама редагувала
    щомісяця, копіюючи структуру ПОПЕРЕДНЬОГО місяця як шаблон - місяць із 31
    днем, скопійований у місяць із 30 (чи 28/29 - лютий), лишав "зайвий"
    рядок-ключ на кшталт "31.09.2026", якого календарно не існує. strptime на
    такому ключі кидає ValueError - раніше це обривало ВЕСЬ
    sync_document_numbers_from_folder (і разом з ним - усю щойно відскановану
    папку з документами). Тепер словники будуються з книги налаштувань, і дні,
    яких в обраному місяці немає, туди не потрапляють
    (utils/project_settings.py::date_key_for_day), а такий ключ, якщо все ж
    трапиться, просто пропускається (з попередженням), а решта дат рахуються
    як зазвичай."""
    week_entries = []
    for date_str, entry in week_dict.items():
        if not entry.get("бат"):
            continue
        try:
            week_entries.append((datetime.strptime(date_str, "%d.%m.%Y"), entry["бат"]))
        except ValueError:
            print_red(f"Некоректна дата-ключ \"{date_str}\" у NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK (такого календарного дня не існує) - пропускаю її для general_br_bat.")
    week_entries.sort(key=lambda item: item[0])

    result = {}
    for date_str in day_keys:
        try:
            day_date = datetime.strptime(date_str, "%d.%m.%Y")
        except ValueError:
            print_red(f"Некоректна дата-ключ \"{date_str}\" у NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY (такого календарного дня не існує) - general_br_bat для неї не порахований.")
            continue
        latest = None
        for week_date, number in week_entries:
            if week_date > day_date:
                break
            latest = (week_date, number)
        result[date_str] = f"{latest[1]} від {latest[0].strftime('%d.%m.%Y')}" if latest else ""
    return result


# Колишній формат: номери документів жили прямо в constants.py, а папка етапу
# мала копію constants.py.bak. Рядок запису словника там виглядав так:
#     f'05.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '361', 'брг': '3956'},
# а перший рядок щоденного словника - f'{(... - relativedelta(days=1)).day}.{...}' -
# останній день попереднього місяця.
_LEGACY_DAY_LINE_RE = re.compile(r"^\s*f'(\d{1,2})\.")
_LEGACY_PREVIOUS_DAY_LINE_RE = re.compile(r"^\s*f'\{\(datetime")
_LEGACY_FIELD_RE = re.compile(r"'(\w+)':\s*'([^']*)'")


def _find_dict_block(lines, varname):
    """Повертає (start, end) - [start, end) рядки-записи словника varname. Якщо varname
    у файлі немає (напр. constants.py.bak лишився зі старої версії constants.py, де
    словник називався/був структурований інакше) - повертає (None, None), а не падає:
    викликачі (load_number_dicts_from_legacy_backup) самі вирішують, як це обробити."""
    header = f"{varname} = {{"
    start = next((i for i, line in enumerate(lines) if line.startswith(header)), None)
    if start is None:
        return None, None
    end = next((i for i in range(start + 1, len(lines)) if lines[i].rstrip("\n") == "}"), None)
    return (None, None) if end is None else (start + 1, end)


def load_number_dicts_from_legacy_backup(backup_path, keys_by_name):
    """Номери з constants.py.bak папки етапу колишнього формату - щоб перехід на
    книгу налаштувань не загубив номери вже наявних етапів. Файл читається як
    текст (виконувати його небезпечно - там той самий інтерактивний вибір
    місяця): день запису - з тексту його ключа, тож номери лягають на ключі
    keys_by_name ({ім'я словника: [ключ, ...]}) поточного місяця так само, як
    колись лягали на рядки constants.py.

    Повертає {ім'я словника: {ключ: запис}} або None, якщо файлу немає чи
    бракує одного зі словників (тоді краще ігнорувати бекап, ніж падати)."""
    if not os.path.isfile(backup_path):
        return None

    with open(backup_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    result = {}
    for varname, keys in keys_by_name.items():
        start, end = _find_dict_block(lines, varname)
        if start is None:
            print_red(f"{backup_path} застарілого формату (не знайдено {varname}) - ігноруємо його.")
            return None
        entries_by_day = {}
        for line in lines[start:end]:
            day_match = _LEGACY_DAY_LINE_RE.match(line)
            day = int(day_match.group(1)) if day_match else (PREVIOUS_MONTH_DAY if _LEGACY_PREVIOUS_DAY_LINE_RE.match(line) else None)
            if day is not None:
                entries_by_day[day] = dict(_LEGACY_FIELD_RE.findall(line))
        labels = day_labels(keys)
        result[varname] = {key: entries_by_day[labels[key]] for key in keys if labels.get(key) in entries_by_day}
    return result


def persist_number_dicts_to_settings_file(settings_file_path, fields):
    """Записує ЗАДАНІ поля номерів в аркуш "Номери документів" книги налаштувань на диску
    (resources/НАЛАШТУВАННЯ.xlsx; колись - прямо в рядки constants.py). Пишеться
    лише в наявні рядки днів - нові рядки не дописуються, а дні, рядків яких
    немає, називаються червоним.

    fields - [(varname, field_name, number_dict, keys), ...] - кожен запис визначає
    ОДНЕ поле ОДНОГО словника для перезапису; викликачі самі збирають список (БАТ/БЗ-
    синхронізація - 'бат'/'general_br_bat'/'бз_бат'; БРГ-синхронізація, нижче -
    'брг'/'бз_брг') - так кожен виклик чіпає ЛИШЕ ті поля, які він дійсно оновив,
    а не переписує ВСЕ щоразу.

    'general_br_bat' сюди приходить УЖЕ порахованим викликачем (_derive_general_br_bat) -
    ця функція лише записує те, що отримала, як і для решти полів.

    Книга відкрита в Excel (чи інша помилка запису) - номери лишаються лише в
    пам'яті цього запуску, а користувач бачить, що треба закрити файл."""
    try:
        missing = write_number_fields(settings_file_path, fields)
    except SettingsError as error:
        print_red(f"Номери документів НЕ записано в книгу налаштувань: {error} У цьому запуску вони діють, "
                  "але наступного разу папку доведеться зчитати ще раз.")
        return
    except Exception as error:
        print_red(f"Номери документів НЕ записано в {settings_file_path}: {error!r}")
        return

    for key in dict.fromkeys(key for _varname, key in missing):
        print_red(f"Аркуш \"{NUMBERS_SHEET}\": немає рядка дня {key} - номер не записано (додайте рядок за потреби).")
    print_green(f"Номери документів записано в {settings_file_path}.")


def _clear_field_for_missing_dates(number_dict, field_name, keys_in_order, matched_dates):
    """Очищає field_name для кожної дати з keys_in_order, якої немає серед matched_dates
    (документ цього виду в папці не знайдено). Повертає відсортований список дат, де
    значення справді було заповнене й тепер очищене - лише їх і показуємо користувачу."""
    cleared = []
    for date_str in keys_in_order:
        if date_str in matched_dates:
            continue
        entry = number_dict.setdefault(date_str, {})
        if entry.get(field_name):
            entry[field_name] = ''
            cleared.append(date_str)
    return sorted(cleared)


def _dates_from_other_month(keys_in_order):
    """Серед keys_in_order (dd.mm.yyyy) - дати, чий (місяць, рік) НЕ збігається
    з НАЙПОШИРЕНІШИМ серед УСІХ keys_in_order - тобто "перехідний" день з
    ПОПЕРЕДНЬОГО місяця (підтверджено користувачем, реальний приклад: перший
    рядок NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY/_WEEK (рядок "попередній" аркуша номерів) - останній
    день місяця ПЕРЕД обраним, доданий лише для general_br_bat/безперервності
    нумерації - решта ключів належать самому обраному місяцю генерації).

    Порожній список, якщо всі ключі одного місяця (напр. NUMBER_OF_DOCUMENTS_BRS_SAVE,
    де такого перехідного рядка взагалі немає) чи keys_in_order порожній
    (persist=False виклики - sync.changes_folder)."""
    parsed = []
    for date_str in keys_in_order:
        try:
            parsed.append((date_str, datetime.strptime(date_str, "%d.%m.%Y")))
        except ValueError:
            continue
    if not parsed:
        return []

    main_month = Counter((d.month, d.year) for _date_str, d in parsed).most_common(1)[0][0]
    return [date_str for date_str, d in parsed if (d.month, d.year) != main_month]


# Той самий словник (значення), що й content/money_report_helpers.py::
# MONTH_NAMES_NOMINATIVE_UPPER, лише нижнім регістром і ЛОКАЛЬНО продубльований
# тут - той самий прийом, що й utils/excel_reader.py::_UKRAINIAN_MONTH_NAMES_UPPER,
# з тієї самої причини (див. docstring _quarter_month_folder_path нижче).
_MONTH_NAMES_NOMINATIVE_LOWER = {
    1: "січень", 2: "лютий", 3: "березень", 4: "квітень", 5: "травень", 6: "червень",
    7: "липень", 8: "серпень", 9: "вересень", 10: "жовтень", 11: "листопад", 12: "грудень",
}


def _quarter_month_folder_path(base_folder, month_int, year_int):
    """Шлях у звичній структурі користувача для КОНКРЕТНОГО місяця:
    <base_folder>/<N квартал>/<назва місяця> (підтверджено користувачем,
    реальний приклад: ".../3 квартал/вересень") - N - номер календарного
    кварталу (1-4), назва місяця - українською, називний відмінок, нижній
    регістр. _MONTH_NAMES_NOMINATIVE_LOWER - ЛОКАЛЬНА копія (не імпорт
    content.money_report_helpers.MONTH_NAMES_NOMINATIVE_UPPER - той самий
    прийом, що й utils/excel_reader.py::_UKRAINIAN_MONTH_NAMES_UPPER):
    sync_stage_constants_backup (sync/stage_folder.py) викликає ЦЕЙ модуль
    НАПРЯМУ з ВЛАСНОГО top-level коду constants.py (режим етапу) - РЕГРЕСІЯ,
    реальний випадок користувача: лінивий імпорт з content.money_report_helpers
    (а той сам імпортує з constants.py) тоді однаково циклиться - ImportError
    "cannot import name ... from partially initialized module constants", бо
    constants.py ще не встиг визначити решту своїх значень (HIGHER_COMMANDER_TITLE
    тощо), коли сам опиняється посеред ланцюжка власного ж імпорту."""
    quarter = (month_int - 1) // 3 + 1
    month_name = _MONTH_NAMES_NOMINATIVE_LOWER[month_int]
    return os.path.join(base_folder, f"{quarter} квартал", month_name)


def _sibling_month_folder_two_levels_up(folder, month_int, year_int):
    """Для дати з ІНШОГО місяця, ніж той, що відповідає обраній папці folder -
    шлях до папки ЦЬОГО (іншого) місяця, побудований за ТІЄЮ САМОЮ структурою
    "<...>/<N квартал>/<назва місяця>", що й сама folder (підтверджено
    користувачем, реальний приклад: обрано ".../4 квартал/жовтень", а
    "перехідний" день - з вересня - потрібна ".../3 квартал/вересень" - ДВА
    рівні ВИЩЕ папки folder, тобто її "дід" - та сама базова папка, спільна
    для ВСІХ кварталів).

    None, якщо такої папки не існує - викликач тоді просто не знаходить
    номер за цю дату, як і раніше (жодної помилки чи попередження - папка
    попереднього місяця може бути ще не створена чи лежати геть інакше)."""
    base_folder = os.path.dirname(os.path.dirname(folder.rstrip("\\/")))
    candidate = _quarter_month_folder_path(base_folder, month_int, year_int)
    return candidate if os.path.isdir(candidate) else None


def _find_number_for_exact_date(folder, filename_regex, target_date_str):
    """Шукає в folder (рекурсивно) ПЕРШИЙ файл, що відповідає filename_regex і
    чия дата - САМЕ target_date_str - повертає знайдений номер (рядок) чи
    None. На відміну від основного сканування sync_document_numbers_from_folder
    нижче, НЕ збирає жодної іншої дати з цієї папки - викликається САМЕ для
    "перехідного" дня з іншого місяця (_dates_from_other_month), щоб не
    засмічувати new_daily_keys/new_task_keys попередженнями про геть не
    пов'язані дні сусідньої папки."""
    for dirpath, _dirnames, filenames in os.walk(folder):
        for filename in filenames:
            match = filename_regex.search(filename)
            if match and match.group(2) == target_date_str:
                return match.group(1)
    return None


def sync_document_numbers_from_folder(
    folder, settings_file_path,
    number_of_documents_brs_every_day, original_day_keys,
    number_of_documents_brs_every_week, original_week_keys,
    number_of_documents_brs_save, original_save_keys,
    persist=True,
):
    """persist=False (напр. підстави за ПОПЕРЕДНІЙ місяць - sync.changes_folder) -
    заповнює number_of_documents_brs_* лише В ПАМ'ЯТІ (для рендеру ЦЬОГО запуску),
    НЕ чіпаючи книгу налаштувань на диску взагалі: викликач тоді ЗАВЖДИ передає ПОРОЖНІ
    original_*_keys (а не поточні ключі словника), інакше "очищення" нижче (дати,
    яких більше немає серед файлів папки) помилково обнулило б УЖЕ заповнені номери
    ІНШОГО (поточного) місяця, чиї ключі випадково опинились у словнику раніше."""
    updated_daily, updated_safety, updated_tasks = 0, 0, 0
    new_daily_keys, new_safety_keys, new_task_keys = set(), set(), set()
    matched_daily_dates, matched_task_dates, matched_safety_dates = set(), set(), set()

    for dirpath, _dirnames, filenames in os.walk(folder):
        for filename in filenames:
            match = _DAILY_BR_FILENAME_RE.search(filename)
            if match:
                number, date_str = match.group(1), match.group(2)
                if date_str not in original_day_keys:
                    new_daily_keys.add(date_str)
                matched_daily_dates.add(date_str)
                number_of_documents_brs_every_day.setdefault(date_str, {'бат': '', 'брг': ''})['бат'] = number
                updated_daily += 1
                continue

            match = _SAFETY_BR_FILENAME_RE.search(filename)
            if match:
                number, date_str = match.group(1), match.group(2)
                if date_str not in original_save_keys:
                    new_safety_keys.add(date_str)
                matched_safety_dates.add(date_str)
                number_of_documents_brs_save.setdefault(date_str, {'бз_бат': '', 'бз_брг': ''})['бз_бат'] = number
                updated_safety += 1
                continue

            match = _TASK_ORDER_FILENAME_RE.search(filename)
            if match:
                number, date_str = match.group(1), match.group(2)
                if date_str not in original_week_keys:
                    new_task_keys.add(date_str)
                matched_task_dates.add(date_str)
                number_of_documents_brs_every_week.setdefault(date_str, {'бат': '', 'посилання_брг': '', 'посилання_бат': ''})['бат'] = number
                updated_tasks += 1

    # "Перехідний" день з ПОПЕРЕДНЬОГО місяця (останній рядок перед 01-м числом
    # обраного місяця - NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY/_WEEK/_SAVE) - обрана
    # папка folder (користувач завжди обирає папку САМЕ обраного місяця
    # генерації, напр. ".../4 квартал/жовтень") НІКОЛИ не містить файл за цю
    # дату - вона належить ПОПЕРЕДНЬОМУ місяцю (підтверджено користувачем,
    # реальний приклад). Якщо сусідня папка попереднього місяця існує (та
    # сама структура "<N> квартал/<місяць>", два рівні вище folder -
    # _sibling_month_folder_two_levels_up) - шукаємо в НІЙ, лише цю ОДНУ
    # дату (_find_number_for_exact_date - а не весь місяць, щоб не засмічувати
    # new_daily_keys/new_task_keys/new_safety_keys попередженнями про геть не
    # пов'язані дні). NUMBER_OF_DOCUMENTS_BRS_SAVE зараз такого перехідного
    # рядка не має (усі його ключі - ОДНОГО місяця, _dates_from_other_month
    # тоді просто повертає [] - нуль змін), але перевіряється тим самим
    # способом, що й решта двох словників - щоб лишатись однаковим для всіх
    # трьох, якщо такий рядок колись з'явиться і там.
    for label, number_dict, keys, matched_dates, filename_regex, field_name, counter_name in (
        ("ЩОДЕННА", number_of_documents_brs_every_day, original_day_keys, matched_daily_dates, _DAILY_BR_FILENAME_RE, 'бат', "daily"),
        ("ЗАВДАННЯ", number_of_documents_brs_every_week, original_week_keys, matched_task_dates, _TASK_ORDER_FILENAME_RE, 'бат', "tasks"),
        ("застосування безпеки", number_of_documents_brs_save, original_save_keys, matched_safety_dates, _SAFETY_BR_FILENAME_RE, 'бз_бат', "safety"),
    ):
        for date_str in _dates_from_other_month(keys):
            if date_str in matched_dates:
                continue
            day_int, month_int, year_int = (int(part) for part in date_str.split("."))
            sibling_folder = _sibling_month_folder_two_levels_up(folder, month_int, year_int)
            if sibling_folder is None:
                continue
            found_number = _find_number_for_exact_date(sibling_folder, filename_regex, date_str)
            if found_number is None:
                continue
            number_dict.setdefault(date_str, {})[field_name] = found_number
            matched_dates.add(date_str)
            if counter_name == "daily":
                updated_daily += 1
            elif counter_name == "tasks":
                updated_tasks += 1
            else:
                updated_safety += 1
            print_green(f"{label}: {date_str} (попередній місяць) знайдено в {sibling_folder}: №{found_number}.")

    if not (updated_daily or updated_safety or updated_tasks):
        print_red(f"У папці {folder} не знайдено файлів у форматі '№... ЩОДЕННА ...', '№... застосування безпеки ...' або '№... ЗАВДАННЯ ...'.")
        return

    print_green(f"Синхронізовано номери документів з {folder}: бат — {updated_daily}, бз_бат — {updated_safety}, завдання — {updated_tasks}.")

    # Документи, яких у папці більше немає, не мають лишати свій старий номер у книзі налаштувань
    # (раніше очищувався лише "бат" для ЩОДЕННА/ЗАВДАННЯ, а "бз_бат" - ніколи, тож видалене
    # застосування безпеки "жило" далі й потрапляло в br_save). Очищується ЛИШЕ той вид
    # документа, якого в папці знайдено хоча б один: папка з самими ЗАВДАННЯМИ не має стерти
    # БЗ-номери. Порожня папка сюди вже не доходить (ранній вихід вище).
    cleared_daily = _clear_field_for_missing_dates(
        number_of_documents_brs_every_day, 'бат', original_day_keys, matched_daily_dates,
    ) if updated_daily else []
    cleared_tasks = _clear_field_for_missing_dates(
        number_of_documents_brs_every_week, 'бат', original_week_keys, matched_task_dates,
    ) if updated_tasks else []
    cleared_safety = _clear_field_for_missing_dates(
        number_of_documents_brs_save, 'бз_бат', original_save_keys, matched_safety_dates,
    ) if updated_safety else []
    for label, cleared in (("ЩОДЕННА", cleared_daily), ("ЗАВДАННЯ", cleared_tasks), ("застосування безпеки", cleared_safety)):
        if cleared:
            print_red(f"{label}: очищено номери для дат, документів яких більше немає в папці: {', '.join(cleared)}")

    # general_br_bat кожного дня - те саме тижневе розпорядження (ЗАВДАННЯ), що й
    # щойно (можливо) оновлений NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK вище - рахується
    # ЩОРАЗУ (а не лише коли саме "завдання"-файли щось оновили), бо новий файл
    # "ЩОДЕННА" (updated_daily) для дати ПІСЛЯ вже наявного тижневого розпорядження
    # так само має отримати актуальний general_br_bat.
    general_br_bat_by_day = _derive_general_br_bat(number_of_documents_brs_every_day.keys(), number_of_documents_brs_every_week)
    for date_str, value in general_br_bat_by_day.items():
        number_of_documents_brs_every_day[date_str]["general_br_bat"] = value

    if not persist:
        return

    persist_number_dicts_to_settings_file(settings_file_path, [
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", "бат", number_of_documents_brs_every_day, original_day_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", "general_br_bat", number_of_documents_brs_every_day, original_day_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", "бат", number_of_documents_brs_every_week, original_week_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_SAVE", "бз_бат", number_of_documents_brs_save, original_save_keys),
    ])

    if new_daily_keys:
        print_red(
            f"ЩОДЕННА: ці дати - не дні обраного місяця на аркуші \"{NUMBERS_SHEET}\" книги налаштувань, "
            "тому НЕ записані (додайте рядок дня чи оберіть папку саме цього місяця): "
            + ", ".join(sorted(new_daily_keys))
        )

    if new_task_keys:
        print_red(
            f"ЗАВДАННЯ: ці дати - не дні обраного місяця на аркуші \"{NUMBERS_SHEET}\" книги налаштувань, "
            "тому НЕ записані (додайте рядок дня чи оберіть папку саме цього місяця): "
            + ", ".join(sorted(new_task_keys))
        )

    if new_safety_keys:
        print_red(
            f"Застосування безпеки: ці дати - не дні обраного місяця на аркуші \"{NUMBERS_SHEET}\" книги налаштувань, "
            "тому НЕ записані (додайте рядок дня чи оберіть папку саме цього місяця): "
            + ", ".join(sorted(new_safety_keys))
        )


def refresh_stage_backup_numbers(
    backup_path,
    number_of_documents_brs_every_day, number_of_documents_brs_every_week, number_of_documents_brs_save,
    day_keys=None, week_keys=None, save_keys=None,
):
    """Переписує книгу номерів етапу (sync/stage_folder.py - НОМЕРИ ДОКУМЕНТІВ.xlsx у
    папці етапу) номерами БАТ/БЗ/БРГ з ПАМ'ЯТІ (тим, що зараз у NUMBER_OF_DOCUMENTS_BRS_*).
    Потрібно після синхронізації БРГ у режимі етапу: при наступному запуску з цим
    етапом номери беруться саме з цієї книги (див. sync_stage_constants_backup), тож
    без цього щойно знайдені номери БРГ тихо губляться. У пам'яті тоді - номери
    з книги етапу плюс щойно знайдені; поле, якого для дати в пам'яті немає,
    лишається таким, як у книзі етапу.

    Якщо книги етапу немає - нічого не робиться. day_keys/week_keys/save_keys
    лишились від колишнього текстового бекапу й не потрібні."""
    if not backup_path or not os.path.isfile(backup_path):
        return
    in_memory = {
        "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY": number_of_documents_brs_every_day,
        "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK": number_of_documents_brs_every_week,
        "NUMBER_OF_DOCUMENTS_BRS_SAVE": number_of_documents_brs_save,
    }
    try:
        stored = read_numbers_workbook(backup_path, {name: list(numbers) for name, numbers in in_memory.items()})
        # Поле, якого для дати в пам'яті немає (None чи відсутнє), лишається таким, як у книзі етапу.
        write_numbers_workbook(backup_path, {
            name: {
                key: {**stored[name].get(key, {}), **{field: value for field, value in entry.items() if value is not None}}
                for key, entry in numbers.items()
            }
            for name, numbers in in_memory.items()
        })
    except SettingsError as error:
        print_red(f"Не вдалось оновити номери етапу: {error}")
        return
    print_green(f"Номери етапу оновлено номерами БАТ/БЗ/БРГ: {backup_path}.")


# Номери БРГ (на відміну від БАТ, вище) НЕ визначаються за іменем файлу - реальні
# документи надходять запакованими в .zip (часто ЩЕ РАЗ запакованими у вкладений
# .zip усередині), а сам номер і дата - лише ВСЕРЕДИНІ тексту документа
# (підтверджено користувачем, реальні зразки resources/вересень, resources/серпень).
#
# "продовжити обороняти батальйонний район" - документ, що ПРОДОВЖУЄ дію бойового
# розпорядження (номер підставляється в NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY[дата]['брг']).
_PHRASE_CONTINUE_DEFENSE = "продовжити обороняти батальйонний район"
# РЕГРЕСІЯ (виявлено на реальних даних серпня, документ №3071 від 19.08.2026 -
# бракувало саме 20.08 в NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY): та сама команда
# "продовжити оборону" реально зустрічається у ДВОХ різних формулюваннях
# залежно від автора конкретного документа - не лише _PHRASE_CONTINUE_DEFENSE
# вище, а й це, довше: "...продовжити ведення оборонного бою у складі І
# ешелону оборони 40 обрмп в батальйонному районі оборони...". Обидва
# формулювання рівнозначні для мети цього сканування (яка фраза саме
# трапилась - не важливо, важливо ЩО документ продовжує дію попереднього
# бойового розпорядження) - тому перевіряється НАЯВНІСТЬ БУДЬ-ЯКОЇ з двох.
_PHRASE_CONTINUE_DEFENSE_ALT = "продовжити ведення оборонного бою"
# "Про невиконання заходів безпеки застосування військ..." - розпорядження з
# безпеки застосування військ (номер підставляється в
# NUMBER_OF_DOCUMENTS_BRS_SAVE[дата]['бз_брг']).
_PHRASE_SAFETY_NONCOMPLIANCE = "про невиконання заходів безпеки застосування військ або їх зрив доповідати командиру"

# Перша дата (ДД.ММ.РРРР чи ДД.ММ.РР) у ТЕКСТІ документа - підтверджено
# користувачем: це дата САМОГО документа; будь-які дати ПІЗНІШЕ в тексті - це вже
# посилання на ПОПЕРЕДНІ розпорядження, чиєю дією цей документ є продовженням
# (напр. "на виконання розпорядження від 31.08.2026 №1031...").
_DATE_IN_TEXT_RE = re.compile(r"\b(\d{2})\.(\d{2})\.(\d{2,4})\b")
# Перше число з 3+ цифр після "№" - підтверджено користувачем: коротші (напр.
# "№ 1") - це номер ПУНКТУ самого документа, а не номер документа.
_NUMBER_IN_TEXT_RE = re.compile(r"№\s*(\d{3,6})\b")
# Лише перші символи тексту документа - і дата, і номер документа завжди в самій
# шапці; пошук по ВСЬОМУ тексту ризикував би підхопити випадкову дату/число з
# основного змісту документа (координати, посилання на інші накази тощо).
_HEADER_SEARCH_WINDOW = 1000


_MAX_ZIP_NESTING_DEPTH = 10


def _iter_document_texts_from_zip(zip_path, tmp_root, depth=0, ancestor_names=frozenset(), errors=None):
    """Рекурсивно розпаковує zip_path (і БУДЬ-ЯКІ вкладені в нього .zip - реальні
    зразки: документ надходить запакованим у .zip, а всередині ЩЕ РАЗ у .zip) у
    tmp_root і повертає пару (текст, розмір файлу в байтах) для КОЖНОГО знайденого .doc/.docx.
    Помилка розпакування чи читання ОКРЕМОГО файлу - попередження, пошук решти
    файлів триває (той самий підхід, що й extract_document_text).

    errors=None (типовий виклик напряму, напр. тести) - одразу друкує червоним.
    errors=список (sync_brigade_document_numbers_from_folder нижче) - лише
    додає туди (той самий принцип, що й content.document_content_search.
    _report_error, яка керує тим самим для extract_document_text - передається
    сюди й далі, у рекурсивні виклики ТА в extract_document_text): сканування
    йде під активним print_progress, і негайний друк ПОСЕРЕД нього
    "приклеївся" б до кінця смужки прогресу в тому самому рядку консолі -
    викликач друкує зібране ОДНИМ акуратним блоком вже ПІСЛЯ завершення
    сканування.

    ancestor_names - імена (basename) УСІХ .zip, вже відкритих у ЦЬОМУ
    ланцюжку рекурсії (не по всій папці загалом - лише "предки" САМЕ цього
    виклику). Реальні дані (серпень, архів "5_736_461дск.zip") довели: архів
    може містити ВКЛАДЕНИЙ .zip з ТИМ САМИМ іменем, що й у нього самого чи
    когось із власних предків (найімовірніше - помилка ПАКУВАННЯ самого
    архіву, напр. хтось випадково заархівував папку, в якій уже лежала копія
    цього ж .zip, - не помилка цього коду й не щось, що можна "прочитати
    правильно": там немає дна). Без цієї перевірки код розпаковував би той
    самий вміст знову й знову, аж до _MAX_ZIP_NESTING_DEPTH - повільно і з
    незрозумілим для користувача попередженням. Перевірка за ІМЕНЕМ (а не
    вмістом/хешем) - свідомий компроміс: дешево, і збігається з реальним
    випадком; двоє РІЗНИХ документів з випадково однаковим іменем на різних
    рівнях вкладеності теоретично дали б хибне спрацювання, але формат
    реальних назв (унікальний числовий ідентифікатор у самій назві) робить
    це малоймовірним.

    depth - запасний запобіжник (РЕГРЕСІЯ, виявлено користувачем, реальні
    дані серпня): та сама архівна "матрьошка" без спрацювання за іменем
    (напр. вкладені .zip із РІЗНИМИ іменами на кожному рівні) раніше давала
    СПРАВЖНЮ нескінченну рекурсію (Python-івський RecursionError) - зараз
    ловиться далеко ДО того, як вичерпає стек виклику."""
    def _report(message):
        if errors is None:
            print_red(message)
        else:
            errors.append(message)

    basename = os.path.basename(zip_path)
    if basename in ancestor_names:
        _report(
            f"Архів {basename} містить вкладений .zip із ТИМ САМИМ іменем ({zip_path}) - "
            f"це самопосилання в самому архіві (найімовірніше, помилка при його пакуванні, "
            f"а не програми) - розпаковувати глибше нема сенсу, пропускаю."
        )
        return
    if depth > _MAX_ZIP_NESTING_DEPTH:
        _report(f"Забагато вкладених .zip (>{_MAX_ZIP_NESTING_DEPTH}) - пропускаю {zip_path}.")
        return

    from content.document_content_search import extract_document_text

    # depth у назві теки розпакування - щоб рівні з ОДНАКОВИМ basename (але
    # РІЗНИМ вмістом - інакше вище вже впіймали б їх як самопосилання) не
    # розпаковувались в ОДНУ Й ТУ САМУ теку, переплутуючи/перезаписуючи вміст
    # різних рівнів вкладеності.
    dest_dir = os.path.join(tmp_root, f"{depth}_{basename}_x")
    try:
        with zipfile.ZipFile(zip_path) as archive:
            archive.extractall(dest_dir)
    except Exception as error:
        _report(f"Не вдалось розпакувати {zip_path}: {error}")
        return

    child_ancestor_names = ancestor_names | {basename}
    for dirpath, _dirnames, filenames in os.walk(dest_dir):
        for filename in filenames:
            full_path = os.path.join(dirpath, filename)
            ext = os.path.splitext(filename)[1].lower()
            if ext == ".zip":
                yield from _iter_document_texts_from_zip(full_path, tmp_root, depth + 1, child_ancestor_names, errors)
            elif ext in (".doc", ".docx"):
                text = extract_document_text(full_path, errors)
                if text:
                    yield text, os.path.getsize(full_path)


def _own_document_date(text):
    """Дата САМОГО документа (_DATE_IN_TEXT_RE, перша в шапці) - None, якщо в
    шапці взагалі немає дати такого формату."""
    match = _DATE_IN_TEXT_RE.search(text[:_HEADER_SEARCH_WINDOW])
    if not match:
        return None
    day, month, year = match.groups()
    if len(year) == 2:
        year = f"20{year}"
    try:
        return datetime(int(year), int(month), int(day))
    except ValueError:
        return None


def _own_document_number(text):
    """Номер САМОГО документа (_NUMBER_IN_TEXT_RE, перший у шапці) - None, якщо
    в шапці взагалі немає такого числа."""
    match = _NUMBER_IN_TEXT_RE.search(text[:_HEADER_SEARCH_WINDOW])
    return match.group(1) if match else None


def _carry_forward_field(number_dict, keys_in_order, field_name, found_by_date_str):
    """Заповнює field_name для КОЖНОГО ключа keys_in_order (у хронологічному
    порядку, як вони й ідуть в аркуші номерів книги налаштувань) - якщо для дати щось
    ЗНАЙДЕНО (found_by_date_str) - підставляє знайдене. Прогалина (нічого не
    знайдено для цієї дати) заповнюється повторенням ПОПЕРЕДНЬОГО знайденого
    значення (підтверджено користувачем) - АЛЕ ЛИШЕ якщо ПІСЛЯ цієї прогалини
    десь іще є ХОЧ ОДНЕ знайдене значення (тобто прогалина - інтерполяція МІЖ
    двома реальними знахідками, а не екстраполяція В НЕВІДОМЕ).

    РЕГРЕСІЯ (виявлено користувачем, реальні дані серпня): попередня версія
    носила "останнє знайдене" значення ВПЕРЕД без обмежень - через це
    ОСТАННІЙ день місяця (напр. 31-ше, для якого документа зазвичай іще
    просто немає - "продовжити оборону" на 1-ше число НАСТУПНОГО місяця в
    обраній папці не з'явиться) хибно ОТРИМУВАВ номер попереднього дня, хоча
    насправді мав ЛИШИТИСЬ ПОРОЖНІМ (те, що вже було в книзі). Та сама
    причина, чому 1-ше число місяця теж не мало б чіпатись, якщо для нього
    самого нічого не знайдено, - РАНІШЕ це забезпечувалось окремим "запасним"
    правилом (лишати незмінним, якщо взагалі нема last_known), але ТЕ САМЕ
    правило хибно дозволяло ІСНУЮЧОМУ значенню 1-го числа "просочитись" як
    last_known - і випадково перезаписати ВЖЕ ПРАВИЛЬНЕ значення 2-го числа,
    якщо для 2-го теж нічого не знайшлось. Тепер last_known ПОХОДИТЬ ЛИШЕ від
    дійсно ЗНАЙДЕНИХ цим скануванням значень - ключ, для якого нічого не
    знайдено, і ПЕРЕД, і ПІСЛЯ якого немає жодної знахідки в потрібному
    напрямку, - завжди лишається НЕЗМІННИМ.

    Повертає список (дата, дата-джерело) ДЛЯ КОЖНОЇ дати, заповненої саме
    ПОВТОРЕННЯМ (а не власною знахідкою) - підтверджено користувачем: без
    цього людина бачить лише агреговане число в підсумковому повідомленні
    ("N днів-прогалин заповнено") і дізнається, ЗА ЯКУ САМЕ дату документа
    насправді НЕ знайдено (а значення - лише повторення сусіднього дня), тільки
    порівнюючи номери вручну рядок за рядком - реальний випадок, що вже
    двічі збивав з пантелику (29.08.2026, де немає жодного документа в
    resources/серпень/28, і показане значення - просто повторення 28.08)."""
    last_found_index = max(
        (index for index, key in enumerate(keys_in_order) if key in found_by_date_str),
        default=-1,
    )
    filled_by_repetition = []
    last_known = None
    last_known_key = None
    for index, key in enumerate(keys_in_order):
        found = found_by_date_str.get(key)
        if found is not None:
            number_dict.setdefault(key, {})[field_name] = found
            last_known = found
            last_known_key = key
        elif last_known is not None and index <= last_found_index:
            number_dict.setdefault(key, {})[field_name] = last_known
            filled_by_repetition.append((key, last_known_key))
        # інакше - ані знайденого значення для ЦЬОГО ключа, ані знайденого
        # значення ПІЗНІШЕ (для інтерполяції) - лишаємо ключ НЕЗМІННИМ (те,
        # що вже було в книзі налаштувань до синхронізації).
    return filled_by_repetition


def _remember_candidate(candidates, date_str, number, size):
    """Запам'ятовує документ-кандидат на дату: для кожного номера - найбільший розмір
    його файлу (той самий номер у двох копіях - це один і той самий розпорядження)."""
    by_number = candidates.setdefault(date_str, {})
    by_number[number] = max(size, by_number.get(number, 0))


def _pick_largest_candidates(candidates):
    """Для кожної дати обирає номер з НАЙБІЛЬШИМ розміром файлу (за однакового розміру -
    менший номер, щоб результат не залежав від порядку файлів у папці). Повертає
    (found: дата -> номер, duplicates: дата -> [(номер, розмір), ...] від найбільшого до
    найменшого - лише для дат, де кандидатів більше одного)."""
    found, duplicates = {}, {}
    for date_str, by_number in candidates.items():
        ranked = sorted(by_number.items(), key=lambda item: (-item[1], int(item[0])))
        found[date_str] = ranked[0][0]
        if len(ranked) > 1:
            duplicates[date_str] = ranked
    return found, duplicates


def _kb(size_bytes):
    return f"{size_bytes / 1024:.1f} КБ".replace(".", ",")


def sync_brigade_document_numbers_from_folder(
    folder, settings_file_path,
    number_of_documents_brs_every_day, day_keys,
    number_of_documents_brs_save, save_keys,
    persist=True,
):
    """Сканує folder (рекурсивно, os.walk - для .zip файлів, включно з ВКЛАДЕНИМИ
    зразок resources/вересень, resources/серпень) - для КОЖНОГО .doc/.docx
    усередині шукає _PHRASE_CONTINUE_DEFENSE ('брг',
    NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY) та _PHRASE_SAFETY_NONCOMPLIANCE ('бз_брг',
    NUMBER_OF_DOCUMENTS_BRS_SAVE).

    Цільова дата - ДАТА САМОГО ДОКУМЕНТА (_own_document_date) + 1 ДЕНЬ
    (підтверджено користувачем, звірено з реальними даними серпня: документ,
    датований, напр., 07.08, стосується ДНЯ 08.08 - фізична папка, де файл
    физично лежав, для розрахунку дати НЕ використовується взагалі, лише текст
    самого документа). Дві дати, що дають ОДНАКОВУ цільову дату з РІЗНИМИ номерами -
    лишається номер документа з НАЙБІЛЬШИМ розміром файлу (підтверджено користувачем),
    а в червоному попередженні видно і обраний, і відхилені варіанти.

    Дні без жодного знайденого документа - _carry_forward_field повторює
    значення з попереднього дня; перший день місяця, для якого взагалі нічого
    не знайдено (і нема звідки повторити) - лишається НЕЗМІННИМ.

    persist=True (як і БАТ/БЗ-синхронізація вище) записує одразу в
    книгу налаштувань на диску, без окремого підтвердження - підтверджено
    користувачем: зайве питання "чи справді записати" лише плутає, людина не
    програміст."""
    candidates_brg, candidates_bz_brg = {}, {}

    # Спершу рахуємо ВСІ .zip (щоб знати total ДО початку) - саме сканування
    # (розпаковування + читання .doc/.docx, деякі - через Word COM) реально
    # довге, підтверджено користувачем: без відсотків незрозуміло, чи скрипт
    # ще працює, чи "завис" (той самий print_progress, що вже є в
    # report_log_war_checker.py для аналогічно довгого сканування ЖБД).
    zip_paths = [
        os.path.join(dirpath, filename)
        for dirpath, _dirnames, filenames in os.walk(folder)
        for filename in filenames
        if filename.lower().endswith(".zip")
    ]

    # scan_errors - ЗБИРАЄ (а не одразу друкує) попередження про файли/архіви, які
    # не вдалось прочитати ПІД ЧАС сканування (_iter_document_texts_from_zip,
    # content.document_content_search._report_error) - воно активне рівно в тому
    # самому циклі, що й print_progress нижче, тож негайний друк посеред
    # активного рядка "лоудингу" приклеївся б до кінця смужки прогресу (реальний
    # випадок, підтверджений користувачем). Друкується ОДНИМ акуратним блоком
    # нижче, вже ПІСЛЯ завершення сканування (100%).
    scan_errors = []
    with tempfile.TemporaryDirectory(prefix="brg_sync_") as tmp_root:
        for index, zip_path in enumerate(zip_paths, start=1):
            for text, size in _iter_document_texts_from_zip(zip_path, tmp_root, errors=scan_errors):
                normalized = " ".join(text.split()).lower()
                has_defense = _PHRASE_CONTINUE_DEFENSE in normalized or _PHRASE_CONTINUE_DEFENSE_ALT in normalized
                has_safety = _PHRASE_SAFETY_NONCOMPLIANCE in normalized
                if not (has_defense or has_safety):
                    continue

                own_date = _own_document_date(text)
                own_number = _own_document_number(text)
                if own_date is None or own_number is None:
                    continue
                target_date_str = (own_date + timedelta(days=1)).strftime("%d.%m.%Y")

                if has_defense:
                    _remember_candidate(candidates_brg, target_date_str, own_number, size)
                if has_safety:
                    _remember_candidate(candidates_bz_brg, target_date_str, own_number, size)

            print_progress(index, len(zip_paths), prefix="Сканування документів БРГ: ")

    if scan_errors:
        print_red(f"⚠ Під час сканування виникли проблеми з деякими файлами/архівами ({len(scan_errors)}):")
        for message in scan_errors:
            print_red(f"  - {message}")

    found_brg, duplicates_brg = _pick_largest_candidates(candidates_brg)
    found_bz_brg, duplicates_bz_brg = _pick_largest_candidates(candidates_bz_brg)
    for label, duplicates in (("брг", duplicates_brg), ("бз_брг", duplicates_bz_brg)):
        if duplicates:
            print_red(f"{label}: за ці дати знайдено КІЛЬКА документів з РІЗНИМИ номерами - взято найбільший за розміром файлу:")
            for date_str in sorted(duplicates):
                chosen, *rejected = duplicates[date_str]
                print_red(
                    f"  {date_str}: залишено {chosen[0]} ({_kb(chosen[1])}), відхилено: "
                    + ", ".join(f"{number} ({_kb(size_bytes)})" for number, size_bytes in rejected)
                )

    if not (found_brg or found_bz_brg):
        print_red(f"У папці {folder} не знайдено жодного документа з фразою про БРГ чи безпеку застосування військ.")
        return

    brg_gaps = _carry_forward_field(number_of_documents_brs_every_day, day_keys, "брг", found_brg)
    bz_brg_gaps = _carry_forward_field(number_of_documents_brs_save, save_keys, "бз_брг", found_bz_brg)

    print_green(
        f"Синхронізовано номери БРГ з {folder}: брг — {len(found_brg)}, бз_брг — {len(found_bz_brg)} "
        f"(знайдених дат; дні-прогалини заповнено повторенням попереднього значення)."
    )
    for label, gaps in (("брг", brg_gaps), ("бз_брг", bz_brg_gaps)):
        if gaps:
            print_red(
                f"{label}: за ці дати документ НЕ знайдено (немає джерела в обраній папці) - "
                f"значення лише ПОВТОРЮЄ попередній знайдений день, а не підтверджене реальним документом:"
            )
            for date, source_date in gaps:
                print_red(f"  - {date} (з {source_date})")

    if not persist:
        return

    persist_number_dicts_to_settings_file(settings_file_path, [
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", "брг", number_of_documents_brs_every_day, day_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_SAVE", "бз_брг", number_of_documents_brs_save, save_keys),
    ])
