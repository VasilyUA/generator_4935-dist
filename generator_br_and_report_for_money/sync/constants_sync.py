import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timedelta

from utils.logging_utils import print_red, print_green, print_progress
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

    РЕГРЕСІЯ (реальний випадок користувача, вересень 2026): ключі
    constants.py - рядки, які людина сама редагує щомісяця, копіюючи
    структуру ПОПЕРЕДНЬОГО місяця як шаблон - місяць із 31 днем, скопійований
    у місяць із 30 (чи 28/29 - лютий), лишає "зайвий" рядок-ключ на кшталт
    "31.09.2026", якого календарно не існує. strptime на такому ключі кидає
    ValueError - раніше це обривало ВЕСЬ sync_document_numbers_from_folder
    (і разом з ним - усю щойно відскановану папку з документами, навіть коли
    сканування вже завершилось успішно, до самого запису в constants.py).
    Тепер такий ключ просто пропускається (з попередженням), а решта дат
    рахуються як зазвичай - людина бачить проблему й може виправити САМЕ цей
    рядок constants.py, а не втрачає весь результат сканування через нього."""
    week_entries = []
    for date_str, entry in week_dict.items():
        if not entry.get("бат"):
            continue
        try:
            week_entries.append((datetime.strptime(date_str, "%d.%m.%Y"), entry["бат"]))
        except ValueError:
            print_red(f"Некоректна дата-ключ \"{date_str}\" у NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK constants.py (такого календарного дня не існує) - пропускаю її для general_br_bat.")
    week_entries.sort(key=lambda item: item[0])

    result = {}
    for date_str in day_keys:
        try:
            day_date = datetime.strptime(date_str, "%d.%m.%Y")
        except ValueError:
            print_red(f"Некоректна дата-ключ \"{date_str}\" у NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY constants.py (такого календарного дня не існує) - general_br_bat для неї не порахований.")
            continue
        latest = None
        for week_date, number in week_entries:
            if week_date > day_date:
                break
            latest = (week_date, number)
        result[date_str] = f"{latest[1]} від {latest[0].strftime('%d.%m.%Y')}" if latest else ""
    return result


def _find_dict_block(lines, varname):
    """Повертає (start, end) - [start, end) рядки-записи словника varname. Якщо varname
    у файлі немає (напр. constants.py.bak лишився зі старої версії constants.py, де
    словник називався/був структурований інакше) - повертає (None, None), а не падає:
    викликачі (load_number_dicts_from_backup) самі вирішують, як це обробити."""
    header = f"{varname} = {{"
    start = next((i for i, line in enumerate(lines) if line.startswith(header)), None)
    if start is None:
        return None, None
    end = next((i for i in range(start + 1, len(lines)) if lines[i].rstrip("\n") == "}"), None)
    return (None, None) if end is None else (start + 1, end)


def _apply_field_to_lines(lines, start, end, field_name, keys_in_order, values_by_key):
    field_re = re.compile(r"('" + re.escape(field_name) + r"':\s*')[^']*(')")
    for offset, i in enumerate(range(start, end)):
        if offset >= len(keys_in_order):
            break
        value = values_by_key.get(keys_in_order[offset])
        if value is None:
            continue
        lines[i] = field_re.sub(lambda m: m.group(1) + value + m.group(2), lines[i], count=1)


def load_number_dicts_from_backup(backup_path, day_keys, week_keys, save_keys):
    """Читає NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY, NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK та
    NUMBER_OF_DOCUMENTS_BRS_SAVE з constants.py.bak за позицією рядків (а не виконанням
    файлу як коду — бекап містить той самий інтерактивний вибір місяця, який небезпечно
    викликати повторно) і перепризначає значення на поточні ключі дат (day_keys/week_keys/save_keys).

    Повертає (day_dict, week_dict, save_dict) або (None, None, None), якщо файлу немає,
    АБО якщо бекап у форматі, якого зараз не існує (напр. constants.py.bak лишився з
    версії ДО розділення NUMBER_OF_DOCUMENTS_BRS на EVERY_WEEK/SAVE - один із трьох
    словників тоді просто відсутній у файлі) - у цьому разі краще повністю ігнорувати
    застарілий бекап (виклик нижче за кодом сам перезапише .bak свіжим), ніж падати."""
    if not os.path.isfile(backup_path):
        return None, None, None

    with open(backup_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    field_re = re.compile(r"'(\w+)':\s*'([^']*)'")

    def load_block(varname, keys):
        start, end = _find_dict_block(lines, varname)
        if start is None:
            return None
        return {key: dict(field_re.findall(lines[i])) for key, i in zip(keys, range(start, end))}

    day_dict = load_block("NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", day_keys)
    week_dict = load_block("NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", week_keys)
    save_dict = load_block("NUMBER_OF_DOCUMENTS_BRS_SAVE", save_keys)

    if day_dict is None or week_dict is None or save_dict is None:
        print_red(
            f"{backup_path} застарілого формату (не знайдено потрібного словника) - ігноруємо його. "
            f"Оберіть \"Так\" на питання про зчитування папки з документами, щоб перегенерувати бекап заново."
        )
        return None, None, None

    return day_dict, week_dict, save_dict


def persist_number_dicts_to_constants_file(constants_file_path, fields):
    """Переписує ЗАДАНІ поля прямо у constants.py на диску (лише для дат, які вже
    існували в файлі до синхронізації — нові дати туди не дописуються автоматично,
    щоб не ризикувати зламати f-string-ключі).

    fields - [(varname, field_name, number_dict, keys), ...] - кожен запис визначає
    ОДНЕ поле ОДНОГО словника для перезапису; викликачі самі збирають список (БАТ/БЗ-
    синхронізація - 'бат'/'general_br_bat'/'бз_бат'; БРГ-синхронізація, нижче -
    'брг'/'бз_брг') - так кожен виклик чіпає ЛИШЕ ті поля, які він дійсно оновив,
    а не переписує ВСЕ щоразу.

    'general_br_bat' сюди приходить УЖЕ порахованим викликачем (_derive_general_br_bat) -
    ця функція лише записує те, що отримала, як і для решти полів."""
    with open(constants_file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    for varname, field_name, number_dict, keys in fields:
        start, end = _find_dict_block(lines, varname)
        _apply_field_to_lines(lines, start, end, field_name, keys, {k: number_dict[k].get(field_name, '') for k in keys})

    backup_path = constants_file_path + ".bak"
    shutil.copyfile(constants_file_path, backup_path)
    with open(constants_file_path, "w", encoding="utf-8") as f:
        f.writelines(lines)

    try:
        import py_compile
        py_compile.compile(constants_file_path, doraise=True)
    except Exception as e:
        shutil.copyfile(backup_path, constants_file_path)
        print_red(f"Не вдалось безпечно записати constants.py, відновлено з резервної копії: {e}")
        return

    print_green(f"constants.py на диску оновлено (резервна копія: {backup_path}).")


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


def sync_document_numbers_from_folder(
    folder, constants_file_path,
    number_of_documents_brs_every_day, original_day_keys,
    number_of_documents_brs_every_week, original_week_keys,
    number_of_documents_brs_save, original_save_keys,
    persist=True,
):
    """persist=False (напр. підстави за ПОПЕРЕДНІЙ місяць - sync.changes_folder) -
    заповнює number_of_documents_brs_* лише В ПАМ'ЯТІ (для рендеру ЦЬОГО запуску),
    НЕ чіпаючи constants.py на диску взагалі: викликач тоді ЗАВЖДИ передає ПОРОЖНІ
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

    if not (updated_daily or updated_safety or updated_tasks):
        print_red(f"У папці {folder} не знайдено файлів у форматі '№... ЩОДЕННА ...', '№... застосування безпеки ...' або '№... ЗАВДАННЯ ...'.")
        return

    print_green(f"Синхронізовано номери документів з {folder}: бат — {updated_daily}, бз_бат — {updated_safety}, завдання — {updated_tasks}.")

    # Документи, яких у папці більше немає, не мають лишати свій старий номер у constants.py
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

    persist_number_dicts_to_constants_file(constants_file_path, [
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", "бат", number_of_documents_brs_every_day, original_day_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", "general_br_bat", number_of_documents_brs_every_day, original_day_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", "бат", number_of_documents_brs_every_week, original_week_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_SAVE", "бз_бат", number_of_documents_brs_save, original_save_keys),
    ])

    if new_daily_keys:
        print_red(
            "ЩОДЕННА: ці дати відсутні серед ключів NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY у "
            "constants.py, тому НЕ записані на диск (додайте вручну за потреби): "
            + ", ".join(sorted(new_daily_keys))
        )

    if new_task_keys:
        print_red(
            "ЗАВДАННЯ: ці дати відсутні серед ключів NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK у "
            "constants.py, тому НЕ записані на диск (додайте вручну за потреби): "
            + ", ".join(sorted(new_task_keys))
        )

    if new_safety_keys:
        print_red(
            "Застосування безпеки: ці дати відсутні серед ключів NUMBER_OF_DOCUMENTS_BRS_SAVE у "
            "constants.py, тому НЕ записані на диск (додайте вручну за потреби): "
            + ", ".join(sorted(new_safety_keys))
        )


def refresh_stage_backup_numbers(
    backup_path,
    number_of_documents_brs_every_day, number_of_documents_brs_every_week, number_of_documents_brs_save,
    day_keys, week_keys, save_keys,
):
    """Переписує в constants.py.bak етапу номери БАТ/БЗ/БРГ з ПАМ'ЯТІ (тим, що зараз у
    NUMBER_OF_DOCUMENTS_BRS_*) - лише ті поля, що синхронізуються з папки; решта полів
    бекапу не чіпається. Потрібно після синхронізації БРГ у режимі етапу: при наступному
    запуску з цим етапом бекап ЗАМІНЮЄ словники з constants.py (див. sync_stage_constants_backup),
    тож без цього щойно знайдені номери БРГ тихо губляться.

    Якщо бекапу немає - нічого не робиться. Якщо в бекапі немає потрібного словника
    (застарілий формат) або отриманий текст не є коректним Python - бекап лишається без змін."""
    if not os.path.isfile(backup_path):
        return

    with open(backup_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    fields = [
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", "бат", number_of_documents_brs_every_day, day_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", "брг", number_of_documents_brs_every_day, day_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", "general_br_bat", number_of_documents_brs_every_day, day_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", "бат", number_of_documents_brs_every_week, week_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_SAVE", "бз_бат", number_of_documents_brs_save, save_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_SAVE", "бз_брг", number_of_documents_brs_save, save_keys),
    ]
    for varname, field_name, number_dict, keys in fields:
        start, end = _find_dict_block(lines, varname)
        if start is None:
            print_red(f"{backup_path}: немає словника {varname} (застарілий формат) - бекап етапу не оновлено.")
            return
        # None (поля немає в пам'яті для цієї дати) - значення лишається таким, як було в бекапі
        _apply_field_to_lines(lines, start, end, field_name, keys, {k: number_dict[k].get(field_name) for k in keys})

    updated_text = "".join(lines)
    try:
        compile(updated_text, backup_path, "exec")
    except SyntaxError as e:
        print_red(f"Не вдалось оновити бекап етапу (некоректний текст номера): {e} - бекап лишено без змін.")
        return

    with open(backup_path, "w", encoding="utf-8") as f:
        f.write(updated_text)
    print_green(f"Бекап етапу оновлено номерами БАТ/БЗ/БРГ: {backup_path}.")


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


def _iter_document_texts_from_zip(zip_path, tmp_root, depth=0, ancestor_names=frozenset()):
    """Рекурсивно розпаковує zip_path (і БУДЬ-ЯКІ вкладені в нього .zip - реальні
    зразки: документ надходить запакованим у .zip, а всередині ЩЕ РАЗ у .zip) у
    tmp_root і повертає пару (текст, розмір файлу в байтах) для КОЖНОГО знайденого .doc/.docx.
    Помилка розпакування чи читання ОКРЕМОГО файлу - червоне попередження, пошук
    решти файлів триває (той самий підхід, що й extract_document_text).

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
    # "\n" на початку КОЖНОГО попередження тут - бо воно може трапитись
    # ПОСЕРЕД активного рядка "лоудингу" print_progress (той друкується лише
    # ПІСЛЯ повного опрацювання одного .zip з верхнього рівня, а ці помилки -
    # ВСЕРЕДИНІ, під час рекурсивного розпакування) - без цього текст
    # попередження "приклеївся" б до кінця смужки прогресу в тому самому
    # рядку консолі.
    basename = os.path.basename(zip_path)
    if basename in ancestor_names:
        print_red(
            f"\nАрхів {basename} містить вкладений .zip із ТИМ САМИМ іменем ({zip_path}) - "
            f"це самопосилання в самому архіві (найімовірніше, помилка при його пакуванні, "
            f"а не програми) - розпаковувати глибше нема сенсу, пропускаю."
        )
        return
    if depth > _MAX_ZIP_NESTING_DEPTH:
        print_red(f"\nЗабагато вкладених .zip (>{_MAX_ZIP_NESTING_DEPTH}) - пропускаю {zip_path}.")
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
        print_red(f"\nНе вдалось розпакувати {zip_path}: {error}")
        return

    child_ancestor_names = ancestor_names | {basename}
    for dirpath, _dirnames, filenames in os.walk(dest_dir):
        for filename in filenames:
            full_path = os.path.join(dirpath, filename)
            ext = os.path.splitext(filename)[1].lower()
            if ext == ".zip":
                yield from _iter_document_texts_from_zip(full_path, tmp_root, depth + 1, child_ancestor_names)
            elif ext in (".doc", ".docx"):
                text = extract_document_text(full_path)
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
    порядку, як вони й ідуть у самому constants.py) - якщо для дати щось
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
    насправді мав ЛИШИТИСЬ ПОРОЖНІМ (те, що вже було в constants.py). Та сама
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
    порівнюючи constants.py вручну рядок за рядком - реальний випадок, що вже
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
        # що вже було в constants.py до синхронізації).
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
    folder, constants_file_path,
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
    constants.py на диску, без окремого підтвердження - підтверджено
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

    with tempfile.TemporaryDirectory(prefix="brg_sync_") as tmp_root:
        for index, zip_path in enumerate(zip_paths, start=1):
            for text, size in _iter_document_texts_from_zip(zip_path, tmp_root):
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
                f"значення лише ПОВТОРЮЄ попередній знайдений день, а не підтверджене реальним документом: "
                + ", ".join(f"{date} (з {source_date})" for date, source_date in gaps)
            )

    if not persist:
        return

    persist_number_dicts_to_constants_file(constants_file_path, [
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", "брг", number_of_documents_brs_every_day, day_keys),
        ("NUMBER_OF_DOCUMENTS_BRS_SAVE", "бз_брг", number_of_documents_brs_save, save_keys),
    ])
