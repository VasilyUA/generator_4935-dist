import os
import re
from collections import Counter

from InquirerPy.prompts.list import ListPrompt

from utils.logging_utils import print_red
from utils.folder_picker import pick_folder
from utils.prompts import ask_yes_no as _ask_yes_no
from checker_accounting.checker import check_accounting, _read_check_file
from sync.constants_sync import sync_document_numbers_from_folder
from content.report_changes import describe_eligible_changes_months
from content.money_report_helpers import month_nominative_upper, month_genitive_lower
from constants import NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY, NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK, NUMBER_OF_DOCUMENTS_BRS_SAVE

# "prev_06_ОБЛІК.xlsx"/"actual_06_ОБЛІК.xlsm" - роль (prev/actual) + 2-цифровий
# місяць + розширення (.xlsx чи .xlsm).
_CHANGES_FILENAME_RE = re.compile(r'^(prev|actual)_(\d{2})_ОБЛІК\.(xlsx|xlsm)$', re.IGNORECASE)


def _ask_content_search_scope():
    """"Усі документи" (ЩОДЕННА + ЗАВДАННЯ) чи "лише ЗАВДАННЯ" (менша, швидша
    вибірка) - режим content.document_content_search.find_person_document_references
    для пункту 100 (єдиного пункту, де підстава будується з реальної згадки ПІБ
    людини в тексті файлу, а не з дати - див. content/report_changes.py)."""
    return ListPrompt(
        message="Шукати підстави за прізвищем: усі документи (ЩОДЕННА + ЗАВДАННЯ) чи лише ЗАВДАННЯ?",
        choices=[
            {"name": "Усі документи (ЩОДЕННА + ЗАВДАННЯ)", "value": "all"},
            {"name": "Лише ЗАВДАННЯ", "value": "only_task"},
        ],
        default="all",
    ).execute()


_MONTH_DIR_RE = re.compile(r'^\d{2}$')

# "1 change", "2 change", ... - той самий принцип, що й sync.stage_folder.
# _STAGE_FOLDER_RE/find_stage_folders для "1 етап"/"2 етап" (окремі "пакети"
# змін, кожен - зі СВОЄЮ повною структурою changes_dir/MM всередині) -
# підтверджено користувачем: коли changes/ має підпапки такого вигляду
# (замість підпапок-місяців НАПРЯМУ), перед звичайним пошуком MM-підпапок
# спершу питається, ЯКИЙ САМЕ пакет використати.
_CHANGE_BATCH_DIR_RE = re.compile(r'^(\d+)\s*change$', re.IGNORECASE)


def _find_change_batch_folders(changes_dir):
    """Повертає [(номер, назва, повний_шлях), ...] для підпапок виду '1 change',
    '2 change', ... з changes_dir, відсортовані від БІЛЬШОГО номера пакета до
    меншого (найновіший пакет - природний default вибору, той самий принцип,
    що й sync.stage_folder.find_stage_folders)."""
    batches = []
    for entry in sorted(os.listdir(changes_dir)):
        full_path = os.path.join(changes_dir, entry)
        if not os.path.isdir(full_path):
            continue
        match = _CHANGE_BATCH_DIR_RE.match(entry)
        if match:
            batches.append((int(match.group(1)), entry, full_path))

    batches.sort(key=lambda batch: batch[0], reverse=True)
    return batches


def _has_month_subfolders(changes_dir):
    """Чи є в changes_dir хоч ОДНА підпапка-місяць (_MONTH_DIR_RE, напр. '05')
    НАПРЯМУ - для перевірки конфлікту з _find_change_batch_folders (обидва типи
    підпапок одночасно - неоднозначно, яку структуру використовувати)."""
    return any(
        os.path.isdir(os.path.join(changes_dir, entry)) and _MONTH_DIR_RE.match(entry)
        for entry in os.listdir(changes_dir)
    )


def _resolve_changes_dir(changes_dir):
    """Якщо changes_dir має підпапки-місяці (_MONTH_DIR_RE) НАПРЯМУ - повертає
    changes_dir БЕЗ ЗМІН (звичайна, "пласка" структура - підтверджено
    користувачем: "працює так як є зараз"). Якщо НАТОМІСТЬ має підпапки-пакети
    (_find_change_batch_folders, '1 change'/'2 change'/...) - питає, який пакет
    використати (найновіший - за замовчуванням), і повертає шлях ДО НЬОГО (він
    сам, за задумом, має підпапки-місяці всередині - решта коду (_pair_changes_files)
    працює з ним так само, як зі звичайним changes_dir).

    ОБИДВА типи підпапок одночасно в ОДНІЙ changes_dir - неоднозначно, яку
    структуру застосовувати - підтверджено користувачем: "не може одночасно
    знаходитись" - явне червоне попередження, None (розділ "Прошу внести
    зміни..." пропускається цілком, як і при відсутній/порожній changes_dir).

    None також повертається, якщо є пакети, але жодного не обрано (скасовано
    вибір) - з окремим попередженням."""
    change_batches = _find_change_batch_folders(changes_dir)
    has_direct_months = _has_month_subfolders(changes_dir)

    if change_batches and has_direct_months:
        print_red(
            f"У папці '{changes_dir}' одночасно є підпапки-місяці (напр. '05') і "
            f"підпапки-пакети змін (напр. '1 change') - неоднозначно, яку структуру "
            f"використати. Лишіть у цій папці ЛИШЕ ОДИН тип підпапок: або місяці "
            f"напряму, або пакети змін (кожен - зі своїми місяцями всередині)."
        )
        return None

    if not change_batches:
        return changes_dir

    choice = ListPrompt(
        message="Оберіть папку зі змінами:",
        choices=[{"name": name, "value": full_path} for _num, name, full_path in change_batches],
        default=change_batches[0][2],
    ).execute()
    if not choice:
        print_red("Папку зі змінами не обрано - розділ 'Прошу внести зміни...' пропущено.")
        return None
    return choice


def _folder_name_matches_month(folder, month_token):
    """Токен місяця (2-цифровий, напр. "06") АБО назва місяця українською -
    називний відмінок ("Червень", month_nominative_upper) чи родовий ("червня",
    month_genitive_lower), БЕЗ урахування регістру - хоч ОДНЕ з цього має
    міститись у НАЗВІ САМОЇ обраної папки (basename, без решти шляху) - захист
    від випадкового вибору папки з документами ІНШОГО місяця (напр. переплутали
    папку, чи в діалозі лишився шлях з минулого разу) для питання "Оберіть
    папку з документами за {місяць}?" - підтверджено користувачем.

    РЕАЛЬНИЙ БАГ (виправлено): спершу перевірявся ЛИШЕ цифровий токен - папка,
    названа буквально словом місяця без жодної цифри (напр. просто "липень"),
    помилково відхилялась як "не той місяць", хоча вона й була саме тим
    місяцем. Перевіряється лише назва папки, а не вміст файлів усередині
    (їхні власні дати звіряються пізніше, при самому пошуку -
    content.document_content_search)."""
    basename = os.path.basename(folder.rstrip("\\/")).lower()
    month_int = int(month_token)
    candidates = (month_token, month_nominative_upper(month_int).lower(), month_genitive_lower(month_int))
    return any(candidate in basename for candidate in candidates)


def _pair_changes_files(changes_dir):
    """Кожен місяць - ОКРЕМА підпапка changes_dir/MM (MM - 2-цифровий номер
    місяця, напр. "05") - повертає {token: {"dir", "prev", "actual"}} лише для
    ПОВНИХ пар (є і prev_MM_ОБЛІК, і actual_MM_ОБЛІК у ЦІЙ підпапці). Підпапка,
    чия назва не є 2-цифровим числом, ігнорується мовчки (не місяць). Попереджає
    (червоним) і пропускає:
    - токен, для якого в його підпапці є лише prev або лише actual (без пари);
    - токен, для якого prev (чи actual) зустрічається одразу в ДВОХ файлах цієї
      підпапки (напр. і .xlsx, і .xlsm одночасно) - неоднозначно, який мається на увазі."""
    pairs = {}
    for entry in sorted(os.listdir(changes_dir)):
        month_dir = os.path.join(changes_dir, entry)
        if not os.path.isdir(month_dir) or not _MONTH_DIR_RE.match(entry):
            continue
        token = entry

        found = {}
        for filename in sorted(os.listdir(month_dir)):
            full_path = os.path.join(month_dir, filename)
            if not os.path.isfile(full_path):
                continue
            match = _CHANGES_FILENAME_RE.match(filename)
            if not match or match.group(2) != token:
                continue
            found.setdefault(match.group(1).lower(), []).append(full_path)

        ambiguous = False
        for role in ("prev", "actual"):
            if len(found.get(role, [])) > 1:
                print_red(
                    f"У папці {month_dir} для '{role}_{token}_ОБЛІК' знайдено кілька файлів "
                    f"(напр. .xlsx і .xlsm одночасно) - пропускаємо цю пару, лишіть лише один файл."
                )
                ambiguous = True
        if ambiguous:
            continue
        if "prev" not in found or "actual" not in found:
            missing_role = "actual" if "prev" in found else "prev"
            print_red(
                f"У папці {month_dir} немає файлу "
                f"'{missing_role}_{token}_ОБЛІК.xlsx' (чи .xlsm) - пара неповна, пропускаємо."
            )
            continue
        pairs[token] = {"dir": month_dir, "prev": found["prev"][0], "actual": found["actual"][0]}
    return pairs


def _warn_if_month_mismatch(token, file_path):
    """Попереджає (червоним), якщо ФАКТИЧНІ дати file_path (найчастіша пара
    рік-місяць серед його ж дат) НЕ відповідають токену місяця (MM) з назви
    його папки/файлу (напр. файл поклали не в ту підпапку 'changes/MM', чи
    дати всередині помилково належать іншому місяцю) - токен лишається єдиним
    джерелом "якого місяця ця пара" для решти коду (content.report_changes
    визначає рік/місяць лише з ФАКТИЧНИХ дат actual-файлу, не звіряючись із
    токеном), тож розбіжність варто показати користувачу явно, а не мовчки
    покладатись на те, що вони завжди збігаються.

    Будь-яка помилка читання файлу тут - мовчки пропускається (не наше діло:
    prepare_changes_files однаково пропустить непридатний файл далі, через
    check_accounting, з власним поясненням)."""
    try:
        _people, date_set = _read_check_file(file_path)
    except Exception:
        return
    if not date_set:
        return
    (year, month), _count = Counter((d.year, d.month) for d in date_set).most_common(1)[0]
    if f"{month:02d}" != token:
        print_red(
            f"⚠ У файлі {file_path} дати здебільшого належать до {month:02d}.{year}, "
            f"а не до місяця {token} (за назвою папки/файлу) - перевірте, чи не переплутано місяць."
        )


def prepare_changes_files(resources_dir):
    """Шукає в {resources_dir}/changes/MM/ (MM - підпапка на кожен 2-цифровий місяць)
    повні пари prev_MM_ОБЛІК/actual_MM_ОБЛІК (розширення .xlsx чи .xlsm), для кожної
    звіряє обидва файли (checker_accounting.check_accounting - той самий принцип, що
    й для resources/check) і зберігає результат як changes_MM_ОБЛІК.xlsx У ТІЙ САМІЙ
    підпапці MM (а не у плоскій changes/) - аудиторський артефакт для користувача
    (щоб одразу візуально бачити розбіжності), а НЕ сировина для самого рапорту: сам
    рапорт (content/report_changes.py) диференціює prev/actual напряму, за правилами
    MONEY_REPORT_CATEGORIES, а не за кольором цього файлу.

    Викликається лише коли обраний режим генерації дійсно потребує цього
    (user_input._CHANGES_CAPABLE_MODES) - саме цей вибір режиму й є "згодою"
    користувача, окремого питання "так/ні" тут більше немає.

    Для КОЖНОГО файлу пари - попереджає (червоним, _warn_if_month_mismatch),
    якщо його ФАКТИЧНІ дати не відповідають місяцю MM його папки/назви - напр.
    файл поклали не в ту підпапку, чи дати всередині помилково належать
    іншому місяцю - підтверджено користувачем.

    Повертає [{"month_token", "prev_path", "actual_path", "changes_path"}, ...] для
    КОЖНОЇ успішно звіреної пари, відсортовано за токеном місяця. Порожній список,
    якщо папки changes немає (з відповідним повідомленням користувачу) чи в ній
    немає жодної повної пари.

    Якщо {resources_dir}/changes/ має підпапки-пакети '1 change'/'2 change'/...
    замість підпапок-місяців НАПРЯМУ - _resolve_changes_dir спершу питає, який
    пакет використати (місяці шукаються вже ВСЕРЕДИНІ обраного пакета) -
    підтверджено користувачем."""
    changes_dir = os.path.join(resources_dir, "changes")
    if not os.path.isdir(changes_dir):
        print_red(
            f"Створіть папку 'changes' в '{resources_dir}', а в ній - підпапку на кожен "
            f"місяць (напр. '05'), і додайте туди два файли ОБЛІК - попередній і "
            f"актуальний, за якими потрібно внести зміни."
        )
        return []

    changes_dir = _resolve_changes_dir(changes_dir)
    if changes_dir is None:
        return []

    pairs = _pair_changes_files(changes_dir)
    results = []
    for token, paths in sorted(pairs.items()):
        _warn_if_month_mismatch(token, paths["prev"])
        _warn_if_month_mismatch(token, paths["actual"])

        changes_path = os.path.join(paths["dir"], f"changes_{token}_ОБЛІК.xlsx")
        result_path = check_accounting(
            check_dir=paths["dir"], output_path=changes_path, file_paths=[paths["prev"], paths["actual"]],
        )
        if result_path is None:
            continue
        results.append({
            "month_token": token, "prev_path": paths["prev"], "actual_path": paths["actual"], "changes_path": result_path,
        })
    return results


def prompt_grounds_folders_for_changes(changes_file_pairs):
    """Для КОЖНОГО місяця, за який справді будуються "зміни" в рапорті
    (content.report_changes.describe_eligible_changes_months - той самий критерій
    "елігібельності", що й сам рапорт використовує) - окреме питання "Оберіть папку
    з документами (підставами) за {місяць}?": номери БАТ/БЗ поточного місяця вже
    відомі (constants.py заповнює їх при виборі місяця генерації), а ось підстави
    для ПОПЕРЕДНЬОГО місяця, який амендується заднім числом, - ні, тому їх треба
    зчитати окремо, за потреби, з іншої папки документів на кожен такий місяць.

    persist=False, original_*_keys=[] (НЕ поточні ключі словника!): цей словник
    представляє ПОТОЧНИЙ обраний місяць генерації (напр. липень), а тут читається
    папка ЗОВСІМ ІНШОГО, попереднього місяця (напр. червень) - лише щоб заповнити
    номери БАТ/БЗ ЦИХ (червневих) дат у пам'яті для рендеру розділу "Прошу внести
    зміни...". Якщо тут передати ПОТОЧНІ ключі (як для поточного місяця), крок
    "очищення застарілих дат" у sync_document_numbers_from_folder помилково обнулив
    би ВЖЕ заповнені номери поточного (липневого) місяця - жодна липнева дата не
    знайшлася б серед файлів червневої папки. Так само constants.py на диску НЕ
    повинен переписуватись через сторонню, минулу папку - лише сам рендер поточного
    запуску має бачити ці номери.

    content_search_folder/content_search_scope записуються прямо в pair (той
    самий словник, що вже проходить через changes_file_pairs аж до
    content.report_changes.build_changes_entries) - для пунктів 100/170
    (content.report_changes._FOLDER_BASED_GROUNDS_POINTS, єдиних, де підстава -
    реальна згадка ПІБ людини в тексті файлу, а не дата).

    Якщо за якийсь місяць папку так і НЕ обрано (відповіли "Ні", чи відповіли
    "Так", але скасували вибір папки), АБО обрано папку, чия НАЗВА не містить
    ні токен цього місяця, ні його українську назву (_folder_name_matches_month -
    підтверджено користувачем: захист від випадкового вибору папки ІНШОГО
    місяця) - явне червоне
    попередження САМЕ з назвою цього місяця (а не загальне, без прив'язки до
    місяця, як раніше) і папка ВІДХИЛЯЄТЬСЯ (як і "не обрано" вище, а не
    приймається попри розбіжність) - підтверджено користувачем: підстава за цей
    місяць лишиться порожньою (крім MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR/_BN,
    якщо їхній період перетинається з цим місяцем - вони й так завжди
    підставляються, незалежно від обраної папки)."""
    for pair, label in describe_eligible_changes_months(changes_file_pairs):
        if not _ask_yes_no(f"Оберіть папку з документами (підставами) за {label}?", False):
            print_red(f"⚠ Не обрано підстав за {label} - підстава буде порожньою (окрім загальних посилань).")
            continue

        chosen_folder = pick_folder(f"Оберіть папку з документами за {label}")
        if not chosen_folder:
            print_red(f"⚠ Папку не обрано за {label} - синхронізація підстав пропущена, підстава буде порожньою (окрім загальних посилань).")
            continue

        if not _folder_name_matches_month(chosen_folder, pair["month_token"]):
            print_red(
                f"⚠ Назва папки '{chosen_folder}' не відповідає місяцю {label} "
                f"(немає ні токена '{pair['month_token']}', ні назви місяця в назві папки) - "
                f"папку відхилено, підстава буде порожньою (окрім загальних посилань)."
            )
            continue

        sync_document_numbers_from_folder(
            chosen_folder, None,
            NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY, [],
            NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK, [],
            NUMBER_OF_DOCUMENTS_BRS_SAVE, [],
            persist=False,
        )
        pair["content_search_folder"] = chosen_folder
        pair["content_search_scope"] = _ask_content_search_scope()
