import calendar
import os
import re
from collections import Counter
from datetime import timedelta
from functools import lru_cache

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from constants import OUTPUT_DIR
from content.log_war_document_reader import scan_log_war_folder, extract_log_war_sections
from content.money_report_helpers import normalize_name, month_nominative_upper
from content.report_document_reader import extract_report_entries, _date_subranges_from_period_text
from checker_accounting.checker import GREEN_FILL, RED_FILL
from utils.logging_utils import print_green, print_red, print_progress
from utils.excel_writer import save_workbook_safely

# "Перевірка рапорту та ЖБД" - звіряє ГОТОВИЙ рапорт на додаткову винагороду
# (ПІБ/пункт/ПЕРІОД - content.report_document_reader.extract_report_entries) з
# реальними журналами бойових дій (ЖБД, по одному .docx на день) - для КОЖНОГО
# дня періоду людини перевіряє, чи ДІЙСНО підтверджено її участь у наративі
# ЖБД, за принципом (підтверджено користувачем, на реальних зразках
# resources/Травень МП 2026): завдання, поставлене СЬОГОДНІ в пункті "5.
# Рішення командира...", підтверджується ЗАВТРА в пункті "6. Відомості про
# виконання..." -> "Хід виконання спланованих завдань". ПЕРШИЙ день КОЖНОГО
# під-діапазону періоду (розділеного "; " у самому тексті ПЕРІОДУ) - такий
# "перехідний" день; решта днів під-діапазону - "звичайні" (людина вже
# фігурує в наративі щодня, без потреби у свіжому наказі про переміщення).
#
# ФОРМАТ РЕЗУЛЬТАТУ - "довгий" (рядок на КОЖНУ окрему перевірку дата+пункт), а
# НЕ календарна сітка дат-колонок - підтверджено користувачем: календарна
# сітка не показувала явно, ЯКУ саме дату/файл/пункт (5 чи 6) саме треба
# виправити для конкретної людини, лише зафарбовану клітинку без напису.
# "Пункт рапорту" - ЛІТЕРАЛЬНИЙ номер пункту з тексту рапорту + сума в дужках
# (_point_label, напр. "1 (30к)") - підтверджено користувачем: сам номер без
# суми, а тим паче внутрішній код raw_value (30/70/100/170/"NOT_PAID"/"?"),
# був незрозумілий без відкриття самого рапорту.
# У ФАЙЛ потрапляють ЛИШЕ рядки з РЕАЛЬНОЮ помилкою (check_report_against_log_war
# фільтрує "Знайдено"/"ОК" ще ДО виклику _write_log_war_check_result) -
# підтверджено користувачем: файл має показувати ЛИШЕ те, що дійсно треба
# виправити, а не увесь масив перевірок.
# ОХОПЛЕННЯ - лише пункти, ЛІТЕРАЛЬНО пронумеровані "1"/"2" у самому рапорті
# (_CHECKED_HEADING_NUMBERS, нижче) - підтверджено користувачем: решта пунктів
# (навіть якщо теж 30к/100к, напр. ретроактивні поправки за минулі місяці під
# іншим номером) цією перевіркою взагалі не займається.
# Підкреслення замість пробілів у САМІЙ НАЗВІ файлу (підтверджено
# користувачем) - деякі програми/шляхи ненадійно працюють з пробілами в
# імені файлу (копіювання, посилання, командний рядок).
_RESULT_FILE_BASE_NAME = "Перевірка_рапорту_та_ЖБД"


def _result_file_name(report_year_month):
    """Ім'я файлу результату - З НАЗВОЮ МІСЯЦЯ перевірки (підтверджено
    користувачем, напр. "..._Травень.xlsx") - якщо відомий (див.
    _majority_year_month - більшість дат САМОГО РАПОРТУ, а не папки ЖБД: цей
    режим не використовує constants.MONTH, дати визначаються з обраних
    файлу/папки щоразу наново). report_year_month - None (рапорт узагалі без
    жодної розпізнаної дати - украй рідкісний випадок) - тоді просто базова
    назва, без місяця."""
    if report_year_month is None:
        return os.path.join(OUTPUT_DIR, f"{_RESULT_FILE_BASE_NAME}.xlsx")
    _year, month = report_year_month
    month_name = month_nominative_upper(month).capitalize()
    return os.path.join(OUTPUT_DIR, f"{_RESULT_FILE_BASE_NAME}_{month_name}.xlsx")

_IDENTITY_COLUMNS = ["Посада", "Звання", "ПІБ", "Пункт рапорту", "Період (з рапорту)"]
_ROW_COLUMNS = ["ДАТА ПЕРЕВІРКИ", "Пункт ЖБД", "Файл ЖБД", "Статус"]
_NO_FILE_LABEL = "(файл ЖБД за цю дату відсутній)"
# "Яскраво червоним" (підтверджено користувачем) - жирний ЧЕРВОНИЙ шрифт
# ПОВЕРХ червоної заливки для рядків з невиконаним пунктом 5/6, щоб такий
# рядок неможливо було проґавити серед решти.
_BRIGHT_RED_FONT = Font(bold=True, color="FF9C0006")
# Сірий - для рядків, "прощених" винятком пункту 30 нижче (ні зелений, ні
# червоний: ми НЕ підтвердили людину, але й не вважаємо це помилкою).
_NEUTRAL_FILL = PatternFill(start_color="FFF2F2F2", end_color="FFF2F2F2", fill_type="solid")
# Кольорова гама (підтверджено користувачем) - ЛИШЕ щоб візуально одразу
# розрізняти, які колонки походять із самого РАПОРТУ (_IDENTITY_COLUMNS -
# Посада/Звання/ПІБ/Пункт рапорту/Період), а які - з перевірки ПРОТИ ЖБД
# (Пункт ЖБД/Файл ЖБД) - "ДАТА ПЕРЕВІРКИ" - ОКРЕМИЙ, третій колір (підтверджено
# користувачем: найважливіша колонка для швидкого орієнтування по датах, тож
# має вирізнятись і від "рапортних", і від "жбд-шних" колонок). "Статус"
# свідомо БЕЗ групової заливки - у нього своя, змістовна (GREEN_FILL/
# RED_FILL/_NEUTRAL_FILL, вище) - групову заливку туди додавати не варто, щоб
# не забивати сенс кольору.
_REPORT_DATA_FILL = PatternFill(start_color="FFD9E1F2", end_color="FFD9E1F2", fill_type="solid")
_LOG_WAR_DATA_FILL = PatternFill(start_color="FFFCE4D6", end_color="FFFCE4D6", fill_type="solid")
_CHECK_DATE_FILL = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")

# Тонкі межі по всій табличці (підтверджено користувачем) - і в шапці, і в
# даних, включно з об'єднаними (merged) клітинками ідентичності запису. БЕЗ
# явного кольору (як і _FALLBACK_BORDER у generators/generate_timetable.py) -
# openpyxl/Excel тоді малює звичайну ЧОРНУ межу, помітну на будь-якому тлі
# (світло-сірий колір, використаний тут раніше, губився на пастельних
# заливках - виявлено користувачем: межі виглядали як відсутні).
_THIN_SIDE = Side(style="thin")
_TABLE_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)

# Виняток для пункту 30 (підтверджено користувачем): відсутність ПІБ у пункті
# 5/6 ЖБД для людини з ЦЬОГО пункту рапорту - ПОМИЛКА лише якщо в ІНШИХ
# абзацах ТОГО САМОГО пункту (5 чи 6) того самого дня є фраза "з батальйонного
# району оборони" (тобто прибатальйонний тиловий район цього дня в наративі
# ВЗАГАЛІ згадується - відсутність САМЕ цієї людини тоді підозріла). Якщо
# фрази немає взагалі - тиловий район цього дня просто НЕ деталізували в
# наративі (та сама причина, що й для решти пункту 30 - реальні зразки
# resources/Травень МП 2026 показують: з певного дня місяця "Хід виконання"
# перестає щодня перелічувати геть усіх, лише окремі події) - відсутність ЦІЄЇ
# конкретної людини тоді НЕ вважається помилкою.
_WAIVED_RAW_VALUE = 30
_BATTALION_DEFENSE_AREA_PHRASE = "З БАТАЛЬЙОННОГО РАЙОНУ ОБОРОНИ"

# Перевірка стосується ЛИШЕ пунктів, ЛІТЕРАЛЬНО пронумерованих "1" (30к) і "2"
# (100к) у самому тексті рапорту (entry["НОМЕР_ПУНКТУ"]) - підтверджено
# користувачем: НЕ будь-який пункт із сумою 30к/100к - в/сл, кому платять
# 30к/100к ЗА ІНШИМ пунктом (напр. ретроактивна поправка за минулий місяць,
# пронумерована "5"/"6"/"7") цією перевіркою НЕ охоплюється, навіть якщо сума
# ті самі 30к/100к (виявлено користувачем: побачив зайвий "6 пункт" у
# результаті).
_CHECKED_HEADING_NUMBERS = ("1", "2")

# Пункт "1" (30к, батальйонний район оборони) - ОСОБЛИВЕ ставлення в
# "новій"/"змішаній" методиці (підтверджено користувачем, ДВІЧІ уточнено на
# реальному прикладі за липень): "нова" методика (документи, де щодня
# заново narratively призначають/підтверджують саме 100к-ротацію) СТРУКТУРНО
# не описує статичне чергування в БРО (пункт 1) ЖОДНИМ чином - ні щоденно, ні
# через окремі "вхід"/"вихід" - тож для дат, де застосовується "нова" зона,
# пункт "1" узагалі НЕ ПЕРЕВІРЯЄТЬСЯ (жодного рядка, навіть не "прощеного" -
# перевірка для нього там просто неможлива, а не "не знайдено"). Для дат
# "старої" зони (чи коли обрано "Стара" для всього рапорту) пункт "1"
# перевіряється звичайним входом/виходом, як завжди
# (_checks_for_entry_old/_checks_for_entry_mixed_point1, нижче).
_POINT1_HEADING_NUMBER = "1"

# Три методики перевірки покриття (підтверджено користувачем, на реальному
# зразку resources/Травень МП 2026):
# - "old" ("стара"): перевіряються ЛИШЕ ВХІД і ВИХІД кожного під-діапазону
#   періоду - завдання на переміщення СТАВИТЬСЯ в пункті "5" ЗА ДЕНЬ ДО (входу
#   чи виходу), виконання підтверджується в пункті "6" САМЕ В ДЕНЬ. Дні МІЖ
#   входом і виходом узагалі НЕ перевіряються.
# - "new" ("нова"): КОЖЕН день періоду отримує СВІЖЕ завдання в пункті "5"
#   ЦЬОГО Ж дня, підтверджене в пункті "6" НАСТУПНОГО - для УСІХ днів без
#   винятку (а не лише для входу/виходу) - СТОСУЄТЬСЯ ЛИШЕ пункту "2" (100к) -
#   пункт "1" (30к, _POINT1_HEADING_NUMBER, вище) у "новій" зоні/методиці
#   ВЗАГАЛІ НЕ ПЕРЕВІРЯЄТЬСЯ (не просто "стара" - зовсім жодної перевірки).
# - "mixed" ("змішана"): дати до автоматично визначеної межі - за "old", дати
#   від іншої межі - за "new" (для пункту "2") чи взагалі не перевіряються
#   (для пункту "1" - див. вище), дати МІЖ двома межами (сам розрив стилю
#   ведення журналу) - не перевіряються взагалі (для обох пунктів).
#   Межа визначається _detect_style_transition АВТОМАТИЧНО (за вмістом самих
#   документів) - без потреби вводити чи пам'ятати конкретні дати
#   (підтверджено користувачем: пам'ятати дату переходу - зайва інформація,
#   скрипт має розпізнавати це сам).
LOG_WAR_LOGIC_OLD = "old"
LOG_WAR_LOGIC_NEW = "new"
LOG_WAR_LOGIC_MIXED = "mixed"
LOG_WAR_LOGIC_MODES = (LOG_WAR_LOGIC_OLD, LOG_WAR_LOGIC_NEW, LOG_WAR_LOGIC_MIXED)

# Поріг (символів, "5"+"6" разом) для розпізнавання "порожнього" (перехідного)
# дня в _detect_style_transition - реальний зразок (після виправлення
# content.log_war_document_reader._NUMBERED_HEADING_RE - див. коментар там:
# фальшиві "заголовки" виду "20.00 07.05.2026 ..." раніше обривали 07.05 на
# 82 символах замість реальних ~54 тис.): 08.05.2026 - лише ~100-200 символів
# шаблонного тексту без жодного ПІБ, тоді як звичайні дні (включно з 07.05) -
# від кількох тисяч до сотень тисяч символів.
_GAP_MAX_COMBINED_LENGTH = 300


def _normalize_text(text):
    return " ".join(text.split()).upper() if text else ""


# Виявлено користувачем (реальний зразок): наратив "Хід виконання спланованих
# завдань" - вільний прозовий текст, а НЕ список ("ПІБ;") - і часто відмінює
# ПІБ за відмінком речення (напр. родовий: "...з батальйонного району оборони
# [кого?] Катерини Олександрівни", а НЕ називний "Катерина Олександрівна", як
# записано в самому рапорті) - точний збіг усього ПІБ (normalize_name(pib) in
# text) тоді НЕ спрацьовує, хоча людина дійсно згадана. КОЖНЕ слово ПІБ - лише
# КОРІНЬ (без двох останніх літер, якщо слово досить довге) + довільний
# "хвіст" - закінчення відмінка типово міняє останні 1-2 літери, корінь
# стабільний.
#
# РЕГРЕСІЯ (виявлено користувачем, ДРУГИЙ реальний зразок): прізвище
# ПРИКМЕТНИКОВОГО типу ("-ИЙ"/"-СЬКИЙ" тощо, напр. "ЧЕКЕРСЬКИЙ") відмінюється
# ЯК ПРИКМЕТНИК ("ЧЕКЕРСЬКИЙ" (називний) -> "ЧЕКЕРСЬКОМУ" (давальний)) - НЕ
# лишається незмінним, як припускалось раніше (на основі іншого зразка,
# "МОЛИБОГ", де прізвище дійсно не відмінювалось) - тож прізвище (перше
# слово) ТЕПЕР толерантне до відмінка так само, як і решта ПІБ (раніше було
# винятком - точний збіг, БЕЗ обрізання кореня).
#
# РЕГРЕСІЯ (виявлено користувачем, ТРЕТІЙ реальний зразок): коротке ім'я
# РІВНО 5 літер (напр. "ПАВЛО") з ЗАМІННИМ (а не додатковим) закінченням
# відмінка - "ПАВЛО" (називний) -> "ПАВЛУ" (давальний) - ОСТАННЯ літера
# ЗАМІНЮЄТЬСЯ ("О"->"У"), а не додається зверху ("ПАВЛО"+"У"="ПАВЛОУ" - так
# НЕ буває). Поріг "> 5" не обрізав 5-літерні слова взагалі (застарілий
# компроміс на користь точності), тож "ПАВЛО" звірялось ТОЧНИМ підрядком і
# НЕ збігалось із "ПАВЛУ" (5-та літера різна). Поріг знижено до "> 3" - тепер
# обрізаються й короткі (4-5 літер) слова, корінь "ПАВЛ" стабільний для
# ВСІХ відмінків "ПАВЛО"/"ПАВЛА"/"ПАВЛУ"/"ПАВЛОМ"/"ПАВЛОВІ".
_DECLENSION_TOLERANT_MIN_LENGTH = 3
_DECLENSION_TRIM = 2


def _word_search_fragment(word):
    """re.escape(word)[:-2] + "\\S*" для достатньо довгого слова - "\\S*"
    покриває БУДЬ-яке відмінкове закінчення (в т.ч. "нульове" - збігається і з
    незміненим словом). Закоротке слово - лишається ПОВНИМ (без обрізання
    кореня), теж із "\\S*" про всяк випадок."""
    if len(word) > _DECLENSION_TOLERANT_MIN_LENGTH:
        word = word[:-_DECLENSION_TRIM]
    return re.escape(word) + r"\S*"


@lru_cache(maxsize=None)
def _name_search_pattern(normalized_pib):
    words = normalized_pib.split()
    if not words:
        return None
    return re.compile(r"\s+".join(_word_search_fragment(word) for word in words))


def _mentioned(pib, text):
    """Чи згадується normalize_name(pib) у text - ТОЛЕРАНТНО до відмінювання
    імені/по батькові (_name_search_pattern, вище) - на відміну від точного
    входження підрядка, яке досі використовується деінде в проєкті для
    ПОДІБНОГО пошуку (checker_accounting.report_checker._check_basis_documents,
    content.document_content_search.find_person_document_references) - там
    джерело тексту, як правило, СПИСОК ("ПІБ (Місце);"), де ПІБ записане так
    само, як у рапорті (називний відмінок), а не вільний наратив зі зміною
    відмінків, тож толерантність до відмінка там не була потрібна."""
    normalized_text = _normalize_text(text)
    if not normalized_text:
        return False
    pattern = _name_search_pattern(normalize_name(pib))
    return bool(pattern and pattern.search(normalized_text))


def _contains_battalion_defense_area_phrase(text):
    return _BATTALION_DEFENSE_AREA_PHRASE in _normalize_text(text)


def _majority_year_month(dates):
    """(рік, місяць), яким належить БІЛЬШІСТЬ dates - None, якщо dates
    порожній. Спільна основа для _warn_if_folder_month_mismatch (звірка
    рапорту з папкою) і назви файлу результату (_result_file_name, нижче) -
    один Counter-прохід замість двох окремих."""
    if not dates:
        return None
    (year, month), _count = Counter((d.year, d.month) for d in dates).most_common(1)[0]
    return year, month


def _warn_if_folder_month_mismatch(report_year_month, log_war_dates):
    """Попереджає (червоним, не блокує), якщо папка ЖБД (обрана користувачем
    ЩОРАЗУ наново - pick_folder, без фіксованого шляху) здебільшого належить до
    ІНШОГО місяця, ніж сам рапорт - той самий Counter-based підхід, що й
    sync.changes_folder._warn_if_month_mismatch, застосований до ДВОХ
    обчислених наборів дат (дат рапорту й дат назв файлів папки), а не до
    токена папки/файлу проти вмісту файлу."""
    folder_year_month = _majority_year_month(log_war_dates)
    if report_year_month is None or folder_year_month is None:
        return

    if report_year_month != folder_year_month:
        report_year, report_month = report_year_month
        folder_year, folder_month = folder_year_month
        print_red(
            f"⚠ Рапорт здебільшого стосується {report_month:02d}.{report_year}, а обрана папка ЖБД - "
            f"{folder_month:02d}.{folder_year} - перевірте, чи не переплутано місяць/папку."
        )


def _file_label(log_war_files, date):
    path = log_war_files.get(date)
    return os.path.basename(path) if path else f"{date.strftime('ЖБД %d.%m.%Y.docx')} {_NO_FILE_LABEL}"


def _covered_value(pib, text, is_waivable):
    """True/False/None ("covered" - для перевірки в _missing_days_count і
    відображення в _write_log_war_check_result: True = ПІБ знайдено; False =
    відсутнє й це ПОМИЛКА; None = відсутнє, але для пункту 30 - "прощено", бо
    фраза _BATTALION_DEFENSE_AREA_PHRASE відсутня в тексті цього ж абзацу
    (тобто прибатальйонний тиловий район цього дня в наративі взагалі не
    деталізували - відсутність САМЕ цієї людини тоді не підозріла)."""
    if _mentioned(pib, text):
        return True
    if is_waivable and not _contains_battalion_defense_area_phrase(text):
        return None
    return False


def _is_month_start(date):
    """Чи date - 1-ше число місяця - див. _waive_boundary_absence (вхід
    точно в ЦЕЙ день - рідкість, реальний вхід міг статись і в попередньому
    місяці, поза межами самого рапорту)."""
    return date.day == 1


def _is_month_end(date):
    """Чи date - останній день місяця, АБО ОДИН день до нього - див.
    _waive_boundary_absence. РЕГРЕСІЯ (уточнено користувачем на реальних
    даних): періоди в рапорті регулярно закінчуються НЕ на буквально
    останньому дні місяця (напр. "01.05.2026-30.05.2026", хоча травень має
    31 день) - реальний вихід так само міг статись пізніше, за межами
    рапорту, тож невеликий допуск в 1 день - навмисний, а не помилка."""
    last_day = calendar.monthrange(date.year, date.month)[1]
    return date.day >= last_day - 1


def _waive_boundary_absence(checks_by_key, boundary_keys, eligible):
    """Підтверджено користувачем (з конкретним обґрунтуванням, УТОЧНЕНО
    ВДРУГЕ): для "старої"/"змішаної" методики - якщо межа (ВХІД чи ВИХІД,
    КОЖНА судиться ОКРЕМО, а не разом - РЕГРЕСІЯ: раніше вимагалось, щоб
    ВЕСЬ під-діапазон був РІВНО одним повним календарним місяцем одночасно
    з обох боків, тому переважна більшість реальних періодів - напр.
    "08.05-31.05" (вихід на кінці місяця, але вхід НЕ 1-го числа) чи
    "01.05-10.05" (вхід 1-го числа, але вихід НЕ в кінці місяця) - взагалі
    не отримували жодного прощення) збігається з початком/кінцем
    календарного місяця (eligible - _is_month_start/_is_month_end) - реальний
    вхід/вихід такої людини, найімовірніше, стався ЗА МЕЖАМИ обраного місяця
    - до нього чи після (напр. пункт 30к: вийшла з БРО ще МИНУЛОГО місяця й
    ще НЕ заходила знов; пункт 100к: зайшла в БРО МИНУЛОГО місяця й ще НЕ
    виходила) - "стара" методика перевіряє ЛИШЕ жорстку пару (день-до,
    день-СЬОГОДНІ), а РЕАЛЬНИЙ вхід/вихід міг статись БУДЬ-ЯКОГО із сусідніх
    днів (скрипт не знає точної дати) - тож ЧАСТКОВА відсутність (не ОБИДВІ
    перевірки межі) теж "прощується" (covered False -> None). Лишається
    помилкою (нічого не прощується) ЛИШЕ якщо межа ПОВНІСТЮ підтверджена (усі
    covered == True) - тоді прощати нічого. Не застосовується до "нової"
    методики (їй користувач це не підтвердив - та й там немає такої пари
    "день-до/день-сьогодні" узагалі)."""
    if not eligible or not boundary_keys:
        return
    if all(checks_by_key[key]["covered"] is True for key in boundary_keys):
        return
    for key in boundary_keys:
        if checks_by_key[key]["covered"] is False:
            checks_by_key[key]["covered"] = None


def _checks_for_entry_old(entry, log_war_files, sections_by_date, person_dates=frozenset()):
    """"Стара" методика (підтверджено користувачем, на реальному прикладі:
    період 02-15.04 -> перевірки ЛИШЕ за 01.04(п.5)+02.04(п.6) - вхід - і
    14.04(п.5)+15.04(п.6) - вихід; дні 03.04-13.04 узагалі НЕ перевіряються):
    для КОЖНОГО під-діапазону періоду - завдання на переміщення СТАВИТЬСЯ в
    пункті "5" ЗА ДЕНЬ ДО дня входу/виходу, підтверджується в пункті "6" САМЕ
    В ДЕНЬ входу/виходу. Одноденний під-діапазон (start == end) - вхід і
    вихід збігаються, лише ОДНА пара перевірок (природно дедуплікується через
    checks_by_key).

    person_dates - УСІ дати ЦІЄЇ Ж людини з УСІХ її записів рапорту (і
    30к, і 100к, незалежно від пункту) - РЕГРЕСІЯ (виявлено користувачем,
    реальний зразок): людина без жодного розриву переходить з одного пункту
    в інший (напр. 100к до 04.05, 30к від 05.05 - без проміжку) - РЕАЛЬНЕ
    переміщення в ЖБД зафіксоване ОДИН раз (03.05 п.5 + 04.05 п.6 - як ВИХІД
    із 100к), а НЕ ДВІЧІ - тож "вхід" у 30к (04.05 п.5 + 05.05 п.6) шукав
    ДРУГЕ, неіснуюче підтвердження того самого руху. Якщо день ПЕРЕД входом
    (чи ПІСЛЯ виходу) вже покритий ІНШИМ записом ЦІЄЇ Ж людини - це не нова
    межа, а безшовне продовження - окрема перевірка тут НЕ потрібна.

    ВХІД, окремо, "прощується" (_waive_boundary_absence), якщо start_date -
    1-ше число місяця; ВИХІД, окремо, "прощується", якщо end_date - в
    останніх 2 днях свого місяця - КОЖНА межа судиться САМА ПО СОБІ, НЕ
    вимагаючи, щоб ОБИДВІ одразу відповідали (підтверджено користувачем:
    реальний вхід/вихід тоді, найімовірніше, стався поза межами обраного
    місяця).

    ДОДАТКОВО (уточнено користувачем, реальний приклад: вхід 03.04, вихід
    16.05, а завантажена папка ЖБД - лише за травень): якщо файлу ЖБД за
    ЦЮ КОНКРЕТНУ дату в обраній папці взагалі НЕМАЄ (log_war_files) - ця
    ОКРЕМА перевірка теж "covered=None" (не помилка) - ми фізично не
    можемо перевірити те, чого не завантажили, і це НЕ провина людини. Цей
    механізм ДІЄ ОДНОЧАСНО й НЕЗАЛЕЖНО від календарного пробачення вище
    (напр. вхід 03.04 - НЕ 1-ше число, тож календарне пробачення тут не
    спрацювало б, а це - спрацьовує). Якщо з ПАРИ дат лише ОДНА відсутня в
    log_war_files - ІНША (за яку файл ДІЙСНО є) перевіряється ЗВИЧАЙНО
    (реальний found/not-found), а не "прощається" лише через сусідство
    відсутньої дати."""
    pib = entry["ПІБ"]
    is_waivable = entry["raw_value"] == _WAIVED_RAW_VALUE
    empty_sections = {"section5": "", "section6_hid": ""}
    checks_by_key = {}

    def _add(date, point, text):
        # Файлу ЖБД за ЦЮ дату в обраній папці взагалі немає - фізично
        # неможливо перевірити, тож НЕ помилка (None), незалежно від
        # календарного пробачення (_waive_boundary_absence, нижче).
        covered = None if date not in log_war_files else _covered_value(pib, text, is_waivable)
        checks_by_key[(date, point)] = {
            "date": date, "point": point, "file_name": _file_label(log_war_files, date),
            "covered": covered,
        }

    def _add_boundary_if_fresh(boundary_date, adjacent_date):
        """Пара (boundary_date-1,"5")+(boundary_date,"6") - ЛИШЕ якщо
        adjacent_date (день ПЕРЕД входом чи ПІСЛЯ виходу) НЕ покритий іншим
        записом ЦІЄЇ Ж людини (person_dates) - інакше це не нова межа, а
        продовження вже задокументованого переміщення."""
        if adjacent_date in person_dates:
            return
        before_date = boundary_date - timedelta(days=1)
        before = sections_by_date.get(before_date, empty_sections)
        today = sections_by_date.get(boundary_date, empty_sections)
        _add(before_date, "5", before["section5"])
        _add(boundary_date, "6", today["section6_hid"])

    for start_date, dates in _date_subranges_from_period_text(entry["ПЕРІОД"]):
        end_date = dates[-1]

        keys_before = set(checks_by_key)
        _add_boundary_if_fresh(start_date, start_date - timedelta(days=1))
        entry_keys = set(checks_by_key) - keys_before
        entry_eligible = _is_month_start(start_date) or (end_date == start_date and _is_month_end(end_date))
        _waive_boundary_absence(checks_by_key, entry_keys, entry_eligible)

        if end_date != start_date:
            keys_before = set(checks_by_key)
            _add_boundary_if_fresh(end_date, end_date + timedelta(days=1))
            exit_keys = set(checks_by_key) - keys_before
            _waive_boundary_absence(checks_by_key, exit_keys, _is_month_end(end_date))

    return sorted(checks_by_key.values(), key=lambda c: (c["date"], c["point"]))


def _checks_for_entry_new(entry, log_war_files, sections_by_date, person_dates=frozenset()):
    """"Нова" методика (ВИПРАВЛЕНО ВТРЕТЄ й уточнено користувачем на реальних
    зразках за липень - ДЯКОВИЧ, МИРОНЕНКО, ТІЩЕНКО, САСЮК): реальні
    документи НЕ підтверджують кожен окремий день перебування в позиції
    (100к) власним завданням - перевіряється ЛИШЕ ОСТАННІЙ день під-діапазону,
    і ЛИШЕ присутність у пункті "6" (Хід виконання), а не пара з пунктом "5".
    Але "Хід виконання" ЗА ОСТАННІЙ день перебування МОЖЕ фігурувати як у
    ЦЬОГО Ж дня файлі (людина ще встигла бути записана туди - ДЯКОВИЧ), так і
    в НАСТУПНОГО дня файлі (звичайна конвенція "Хід виконання" - підтверджує
    ВЧОРАШНЄ виконання, напр. САСЮК: виконував завдання 16.07, підтверджено в
    17.07, хоча 17.07 він уже у відрядженні) - тож спершу перевіряється
    ОСТАННІЙ день, і ЛИШЕ якщо там НЕ знайдено - перевіряється НАСТУПНИЙ
    (запис помилки тоді - за НАСТУПНИМ днем, де перевірку справді завершено).
    Перевірка НАСТУПНОГО дня НЕ виконується, якщо вже знайдено в останньому
    дні (ТІЩЕНКО - вибув у шпиталь наступного дня, там і не було б чого
    шукати, але цього й не треба - він уже знайдений в останній день).

    ВИКЛИКАЄТЬСЯ ЛИШЕ для пункту "2" (100к) - диспетчер _checks_for_entry
    ЗАВЖДИ повертає ПОРОЖНІЙ список для пункту "1" (30к), коли застосовується
    "нова" методика (незалежно від того, обрано "Нова" для всього рапорту чи
    це "нова" зона "Змішаної") - підтверджено користувачем: "нові" документи
    структурно НЕ описують статичне чергування в БРО жодним чином (ні щодня,
    ні через вхід/вихід), тож перевірка для пункту "1" там просто неможлива,
    а не "не знайдено".

    person_dates - НЕ впливає на цю методику (лишений у сигнатурі лише для
    однакового виклику з диспетчера _checks_for_entry)."""
    pib = entry["ПІБ"]
    is_waivable = entry["raw_value"] == _WAIVED_RAW_VALUE
    empty_sections = {"section5": "", "section6_hid": ""}
    checks_by_key = {}

    for _start_date, dates in _date_subranges_from_period_text(entry["ПЕРІОД"]):
        end_date = dates[-1]
        today = sections_by_date.get(end_date, empty_sections)
        if _mentioned(pib, today["section6_hid"]):
            check_date, text = end_date, today["section6_hid"]
        else:
            next_date = end_date + timedelta(days=1)
            text = sections_by_date.get(next_date, empty_sections)["section6_hid"]
            check_date = next_date
        checks_by_key[(check_date, "6")] = {
            "date": check_date, "point": "6", "file_name": _file_label(log_war_files, check_date),
            "covered": _covered_value(pib, text, is_waivable),
        }

    return sorted(checks_by_key.values(), key=lambda c: (c["date"], c["point"]))


def _checks_for_entry_mixed(entry, log_war_files, sections_by_date, old_cutoff, new_cutoff, person_dates=frozenset()):
    """"Змішана" методика: під-діапазон, що ПОЧИНАЄТЬСЯ в "старій" зоні
    (start_date <= old_cutoff) - ВХІД перевіряється за "старою" методикою
    (день ДО - п.5, день входу - п.6); під-діапазон, що ЗАКІНЧУЄТЬСЯ в
    "новій" зоні (end_date >= new_cutoff) - ВИХІД перевіряється за "новою"
    (ВИПРАВЛЕНО ВТРЕТЄ - див. _checks_for_entry_new: присутність в пункті "6"
    САМОГО останнього дня під-діапазону, а якщо там не знайдено - в пункті
    "6" НАСТУПНОГО дня, БЕЗ пункту "5" узагалі); під-діапазон, що ПОВНІСТЮ
    лишається в "старій" зоні (не сягає "нової") - ВИХІД перевіряється за
    "старою" (день ДО останнього дня -
    п.5, останній день - п.6). Дні, що НЕ є ані початком (у "старій" зоні),
    ані кінцем під-діапазону - НЕ перевіряються взагалі, так само як і дні
    строго МІЖ межами (сам розрив стилю ведення журналу). old_cutoff/
    new_cutoff визначаються АВТОМАТИЧНО (_detect_style_transition) -
    користувачу не потрібно їх вводити чи пам'ятати (підтверджено
    користувачем).

    person_dates - те саме, що й у _checks_for_entry_old (див. докстрінг
    там) - стосується ЛИШЕ "старозонного" входу й "старостильного" виходу
    (обидва - пари на межі двох календарних днів, де можливе подвійне
    покриття із сусіднім записом ЦІЄЇ Ж людини); "новостильний" вихід від
    person_dates НЕ залежить - див. _checks_for_entry_new.

    Файл ЖБД за конкретну дату відсутній у log_war_files - "covered=None"
    (не помилка), незалежно від календарного пробачення - див. докстрінг
    _checks_for_entry_old."""
    pib = entry["ПІБ"]
    is_waivable = entry["raw_value"] == _WAIVED_RAW_VALUE
    empty_sections = {"section5": "", "section6_hid": ""}
    checks_by_key = {}

    def _add(date, point, text):
        covered = None if date not in log_war_files else _covered_value(pib, text, is_waivable)
        checks_by_key[(date, point)] = {
            "date": date, "point": point, "file_name": _file_label(log_war_files, date),
            "covered": covered,
        }

    def _add_boundary_if_fresh(boundary_date, adjacent_date):
        if adjacent_date in person_dates:
            return
        before_date = boundary_date - timedelta(days=1)
        before = sections_by_date.get(before_date, empty_sections)
        today = sections_by_date.get(boundary_date, empty_sections)
        _add(before_date, "5", before["section5"])
        _add(boundary_date, "6", today["section6_hid"])

    def _add_new_style_exit(end_date):
        today = sections_by_date.get(end_date, empty_sections)
        if _mentioned(pib, today["section6_hid"]):
            _add(end_date, "6", today["section6_hid"])
            return
        next_date = end_date + timedelta(days=1)
        tomorrow = sections_by_date.get(next_date, empty_sections)
        _add(next_date, "6", tomorrow["section6_hid"])

    for start_date, dates in _date_subranges_from_period_text(entry["ПЕРІОД"]):
        end_date = dates[-1]
        starts_in_old_zone = start_date <= old_cutoff
        ends_in_new_zone = end_date >= new_cutoff

        if starts_in_old_zone:
            keys_before = set(checks_by_key)
            _add_boundary_if_fresh(start_date, start_date - timedelta(days=1))
            entry_keys = set(checks_by_key) - keys_before
            entry_eligible = _is_month_start(start_date) or (
                end_date == start_date and not ends_in_new_zone and _is_month_end(end_date)
            )
            _waive_boundary_absence(checks_by_key, entry_keys, entry_eligible)

        if ends_in_new_zone:
            _add_new_style_exit(end_date)
        elif starts_in_old_zone and end_date != start_date:
            keys_before = set(checks_by_key)
            _add_boundary_if_fresh(end_date, end_date + timedelta(days=1))
            exit_keys = set(checks_by_key) - keys_before
            _waive_boundary_absence(checks_by_key, exit_keys, _is_month_end(end_date))

    return sorted(checks_by_key.values(), key=lambda c: (c["date"], c["point"]))


def _checks_for_entry_mixed_point1(entry, log_war_files, sections_by_date, old_cutoff, new_cutoff, person_dates=frozenset()):
    """Пункт "1" (30к) під "змішаною" методикою (підтверджено користувачем):
    дати "старої" зони (<= old_cutoff) - звичайний вхід/вихід, як
    _checks_for_entry_old (вихід - ЛИШЕ якщо весь під-діапазон лишається в
    "старій" зоні, інакше людина просто переходить у розрив/"нову" зону без
    штучного "виходу" на межі); дати "нової" зони (>= new_cutoff) і дати в
    розриві МІЖ зонами - ВЗАГАЛІ НЕ ПЕРЕВІРЯЮТЬСЯ (на відміну від пункту "2",
    _checks_for_entry_mixed, вище, де "нова" зона перевіряється щодня) -
    "нові" документи структурно не описують статичне чергування в БРО жодним
    чином, тож перевірка там неможлива, а не "не знайдено".

    Файл ЖБД за конкретну дату відсутній у log_war_files - "covered=None"
    (не помилка), незалежно від календарного пробачення - див. докстрінг
    _checks_for_entry_old."""
    pib = entry["ПІБ"]
    is_waivable = entry["raw_value"] == _WAIVED_RAW_VALUE
    empty_sections = {"section5": "", "section6_hid": ""}
    checks_by_key = {}

    def _add(date, point, text):
        covered = None if date not in log_war_files else _covered_value(pib, text, is_waivable)
        checks_by_key[(date, point)] = {
            "date": date, "point": point, "file_name": _file_label(log_war_files, date),
            "covered": covered,
        }

    def _add_boundary_if_fresh(boundary_date, adjacent_date):
        if adjacent_date in person_dates:
            return
        before_date = boundary_date - timedelta(days=1)
        before = sections_by_date.get(before_date, empty_sections)
        today = sections_by_date.get(boundary_date, empty_sections)
        _add(before_date, "5", before["section5"])
        _add(boundary_date, "6", today["section6_hid"])

    for start_date, dates in _date_subranges_from_period_text(entry["ПЕРІОД"]):
        end_date = dates[-1]
        if start_date > old_cutoff:
            continue

        keys_before = set(checks_by_key)
        _add_boundary_if_fresh(start_date, start_date - timedelta(days=1))
        entry_keys = set(checks_by_key) - keys_before
        ends_in_old_zone = end_date <= old_cutoff
        entry_eligible = _is_month_start(start_date) or (
            end_date == start_date and ends_in_old_zone and _is_month_end(end_date)
        )
        _waive_boundary_absence(checks_by_key, entry_keys, entry_eligible)

        if ends_in_old_zone and end_date != start_date:
            keys_before = set(checks_by_key)
            _add_boundary_if_fresh(end_date, end_date + timedelta(days=1))
            exit_keys = set(checks_by_key) - keys_before
            _waive_boundary_absence(checks_by_key, exit_keys, _is_month_end(end_date))

    return sorted(checks_by_key.values(), key=lambda c: (c["date"], c["point"]))


def _checks_for_entry(entry, log_war_files, sections_by_date, logic_mode, old_cutoff=None, new_cutoff=None, person_dates=frozenset()):
    """Список {"date", "point", "file_name", "covered"} - ОДИН запис на КОЖНУ
    окрему перевірку (а не {дата: покрито}) - підтверджено користувачем: так
    одразу видно ДАТУ, ФАЙЛ ЖБД і ПУНКТ (5 чи 6), а не лише зафарбовану
    клітинку в календарній сітці. "covered" - True/False/None, див.
    _covered_value. Диспетчер до однієї з трьох методик (LOG_WAR_LOGIC_MODES,
    вище) - "mixed" без визначеної межі (old_cutoff/new_cutoff - None, коли
    _detect_style_transition нічого не знайшла) тихо переходить на "new" (як і
    check_report_against_log_war, що вже надрукувала попередження про це).

    person_dates - див. докстрінги _checks_for_entry_old/_checks_for_entry_new/
    _checks_for_entry_mixed - усі три методики пропускають ПЕРШИЙ день
    під-діапазону, якщо день перед ним уже покритий ІНШИМ записом ЦІЄЇ Ж
    людини (безшовний перехід між пунктами 30к/100к, без розриву).

    Пункт "1" (30к, _POINT1_HEADING_NUMBER) - ЛИШЕ "стара" методика має сенс
    (звичайний вхід/вихід); там, де застосовується "нова" (logic_mode "new",
    чи "нова" зона/фолбек "змішаної") - пункт "1" ВЗАГАЛІ НЕ ПЕРЕВІРЯЄТЬСЯ
    (підтверджено користувачем: "нові" документи структурно не описують
    статичне чергування в БРО жодним чином - ні щодня, ні через вхід/вихід)."""
    is_point1 = entry["НОМЕР_ПУНКТУ"] == _POINT1_HEADING_NUMBER
    if logic_mode == LOG_WAR_LOGIC_OLD:
        return _checks_for_entry_old(entry, log_war_files, sections_by_date, person_dates)
    if logic_mode == LOG_WAR_LOGIC_MIXED and old_cutoff is not None and new_cutoff is not None:
        if is_point1:
            return _checks_for_entry_mixed_point1(entry, log_war_files, sections_by_date, old_cutoff, new_cutoff, person_dates)
        return _checks_for_entry_mixed(entry, log_war_files, sections_by_date, old_cutoff, new_cutoff, person_dates)
    # "new" (чи "mixed" без визначеної межі - тихо переходить на "new").
    if is_point1:
        return []
    return _checks_for_entry_new(entry, log_war_files, sections_by_date, person_dates)


def _detect_style_transition(log_war_files, sections_by_date):
    """Автоматично визначає межу переходу зі "старої" методики ведення ЖБД на
    "нову" (підтверджено користувачем: скрипт має сам розпізнавати це, а не
    вимагати вводити/пам'ятати дати) - шукає НАЙДОВШУ суцільну послідовність
    "порожніх" днів (розділи "5"+"6" разом коротші за _GAP_MAX_COMBINED_LENGTH)
    серед ВІДСОРТОВАНИХ дат папки - реальний зразок resources/Травень МП 2026:
    08.05 - лише шаблонний текст без жодного ПІБ (розрив), дні ДО (включно з
    07.05) - десятки тисяч символів (стара методика), дні ПІСЛЯ - кілька
    тисяч (нова).

    Повертає (остання_дата_ПЕРЕД_розривом, перша_дата_ПІСЛЯ_розриву), або
    None, якщо розриву не знайдено (замало дат; жодного "порожнього" дня;
    розрив упирається в сам початок чи кінець папки, тож немає дати "до" чи
    "після", яку можна взяти за межу)."""
    dates = sorted(log_war_files)
    if len(dates) < 3:
        return None

    def _is_blank(date):
        sections = sections_by_date.get(date)
        if not sections:
            return False
        return len(sections["section5"]) + len(sections["section6_hid"]) < _GAP_MAX_COMBINED_LENGTH

    best_run, run_start_index = None, None
    for index, date in enumerate(dates):
        if _is_blank(date):
            if run_start_index is None:
                run_start_index = index
            continue
        if run_start_index is not None:
            run = (run_start_index, index - 1)
            if best_run is None or (run[1] - run[0]) > (best_run[1] - best_run[0]):
                best_run = run
            run_start_index = None
    if run_start_index is not None:
        run = (run_start_index, len(dates) - 1)
        if best_run is None or (run[1] - run[0]) > (best_run[1] - best_run[0]):
            best_run = run

    if best_run is None:
        return None

    run_start_index, run_end_index = best_run
    if run_start_index == 0 or run_end_index == len(dates) - 1:
        return None

    return dates[run_start_index - 1], dates[run_end_index + 1]


def _print_style_detection_table(log_war_files, sections_by_date, old_cutoff, new_cutoff):
    """Друкує ПОДЕННУ таблицю (дата, довжина "5"+"6", позначку "порожній
    день"/старий/новий за визначеною межею) - підтверджено користувачем:
    один короткий підсумковий рядок ("стара методика по X, нова - з Y") без
    жодних доказів "виглядає непереконливо" ("бредове повідомлення") -
    користувач має бачити САМІ ДОВЖИНИ ТЕКСТУ, з яких вирахувано межу, щоб
    самому оцінити, чи це вирахувано правильно (реальний випадок: короткий
    "порожній" день посеред місяця - це НЕ обов'язково стиль ЖБД дійсно
    змінився назавжди, могла це бути лише тимчасова прогалина одного дня)."""
    dates = sorted(log_war_files)
    print_green("Автоматичне визначення стилю ЖБД за довжиною тексту пунктів 5+6 (поріг \"порожнього\" дня - "
                 f"{_GAP_MAX_COMBINED_LENGTH} символів):")
    for date in dates:
        sections = sections_by_date.get(date, {"section5": "", "section6_hid": ""})
        combined = len(sections["section5"]) + len(sections["section6_hid"])
        in_gap = bool(old_cutoff and new_cutoff and old_cutoff < date < new_cutoff)
        if combined < _GAP_MAX_COMBINED_LENGTH:
            label = "порожній день (розрив)" if in_gap else "порожній день"
        elif old_cutoff and date <= old_cutoff:
            label = "стара методика"
        elif new_cutoff and date >= new_cutoff:
            label = "нова методика"
        else:
            label = ""
        print(f"  {date.strftime('%d.%m.%Y')}: {combined:6d} символів - {label}")


def _missing_days_count(checks):
    """Кількість ДНІВ (а не окремих рядків 5/6) без підтвердження - день
    вважається "з помилкою", якщо ХОЧА Б ОДНА його перевірка (одна для
    звичайного дня, дві - "5" і "6" - для перехідного) має covered == False
    (СТРОГО - не None: None означає "прощено" пунктом 30 і не вважається
    помилкою, лише неповним підтвердженням)."""
    by_date = {}
    for check in checks:
        by_date.setdefault(check["date"], []).append(check["covered"])
    return sum(1 for values in by_date.values() if any(value is False for value in values))


def _point_label(entry):
    """"1 (30к)"/"2 (100к)" - літеральний номер пункту + сума (raw_value) в
    дужках (підтверджено користувачем) - охоплення й так обмежено
    _CHECKED_HEADING_NUMBERS ("1"/"2", завжди 30/100 відповідно), тож сума тут
    - для наочності, а не для розрізнення пунктів."""
    raw_value = entry["raw_value"]
    amount = f"{raw_value}к" if isinstance(raw_value, int) else str(raw_value)
    return f'{entry["НОМЕР_ПУНКТУ"]} ({amount})'


def _write_log_war_check_result(entries_with_checks, output_path):
    """Пише результат "довгим" форматом (підтверджено користувачем): для
    КОЖНОГО запису рапорту - Посада/Звання/ПІБ/Пункт рапорту/Період ОБ'ЄДНАНІ
    вертикально на всю групу рядків цього запису, а далі - ОКРЕМИЙ рядок на
    КОЖНУ перевірку (ДАТА ПЕРЕВІРКИ/Пункт ЖБД/Файл ЖБД/Статус). "Статус"
    зафарбований GREEN_FILL/RED_FILL/_NEUTRAL_FILL (ті самі, що й
    changes_MM_ОБЛІК.xlsx), з жирним яскраво-червоним шрифтом для відсутніх -
    щоб такий рядок неможливо було проґавити. РЕШТА колонок (і в шапці, і в
    даних) зафарбована ГРУПОВО (підтверджено користувачем) - _REPORT_DATA_FILL
    для колонок ІЗ САМОГО РАПОРТУ (Посада..Період), _CHECK_DATE_FILL для
    "ДАТА ПЕРЕВІРКИ" ОКРЕМО, _LOG_WAR_DATA_FILL для решти колонок перевірки
    ПРОТИ ЖБД (Пункт ЖБД/Файл ЖБД) - щоб одразу було видно, звідки яка
    колонка. Тонка сіра межа (_TABLE_BORDER) - по КОЖНІЙ клітинці всієї
    табличці (і об'єднаних теж - інакше межа об'єднаного блоку виглядала б
    розірваною)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Перевірка"

    headers = _IDENTITY_COLUMNS + _ROW_COLUMNS
    ws.append(headers)
    identity_col_count = len(_IDENTITY_COLUMNS)
    for col_index, cell in enumerate(ws[1], start=1):
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if col_index <= identity_col_count:
            cell.fill = _REPORT_DATA_FILL
        elif col_index == identity_col_count + 1:
            cell.fill = _CHECK_DATE_FILL
        elif col_index <= identity_col_count + 3:
            cell.fill = _LOG_WAR_DATA_FILL

    row_index = 2
    for item in entries_with_checks:
        entry, checks = item["entry"], item["checks"]
        first_row = row_index
        for check in checks:
            date_cell = ws.cell(row=row_index, column=identity_col_count + 1, value=check["date"].strftime("%d.%m.%Y"))
            date_cell.fill = _CHECK_DATE_FILL
            for col_offset, value in enumerate((check["point"], check["file_name"]), start=2):
                cell = ws.cell(row=row_index, column=identity_col_count + col_offset, value=value)
                cell.fill = _LOG_WAR_DATA_FILL

            status_cell = ws.cell(row=row_index, column=identity_col_count + 4)
            if check["covered"] is True:
                status_cell.value, status_cell.fill = "Знайдено", GREEN_FILL
            elif check["covered"] is False:
                status_cell.value, status_cell.fill = "ВІДСУТНЄ", RED_FILL
                status_cell.font = _BRIGHT_RED_FONT
            else:
                status_cell.value, status_cell.fill = "ОК (без деталізації)", _NEUTRAL_FILL
            row_index += 1

        last_row = row_index - 1
        identity_values = [entry["ПОСАДА"], entry["ЗВАННЯ"], entry["ПІБ"], _point_label(entry), entry["ПЕРІОД"]]
        for col_offset, value in enumerate(identity_values, start=1):
            cell = ws.cell(row=first_row, column=col_offset, value=value)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.fill = _REPORT_DATA_FILL
            if last_row > first_row:
                ws.merge_cells(start_row=first_row, end_row=last_row, start_column=col_offset, end_column=col_offset)

    # Межа - ОКРЕМИМ проходом по ВСІХ клітинках (включно з "порожніми"
    # клітинками під об'єднаннями) - openpyxl не домальовує межу автоматично
    # для об'єднаного блоку, лише за стилем видимої (верхньої) клітинки.
    last_data_row = row_index - 1
    for row in ws.iter_rows(min_row=1, max_row=max(last_data_row, 1), min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = _TABLE_BORDER

    for col_index, width in enumerate([20, 16, 30, 16, 22, 14, 10, 45, 14], start=1):
        ws.column_dimensions[get_column_letter(col_index)].width = width
    ws.freeze_panes = "A2"

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    save_workbook_safely(wb, output_path)


def check_report_against_log_war(report_path, log_war_folder, output_path=None, logic_mode=LOG_WAR_LOGIC_NEW):
    """Головна функція: рапорт report_path + папка log_war_folder (обрана
    користувачем, ЩОРАЗУ вручну - pick_folder) -> xlsx-звіт про покриття
    кожного дня участі кожної людини журналами бойових дій. Повертає шлях до
    збереженого файлу (завжди - навіть якщо все покрито, як і
    check_report_against_oblik).

    logic_mode - одна з LOG_WAR_LOGIC_MODES ("old"/"new"/"mixed", обирає користувач -
    user_input.py). "mixed" САМА визначає межу переходу стилю ведення ЖБД
    (_detect_style_transition, за вмістом документів) - якщо межу не знайдено,
    тихо (з попередженням) переходить на "new" для всього рапорту.

    У ФАЙЛ потрапляють ЛИШЕ рядки з РЕАЛЬНОЮ помилкою (covered is False) -
    "Знайдено"/"ОК" (covered True/None) НЕ записуються взагалі (підтверджено
    користувачем: файл має показувати ЛИШЕ те, що дійсно треба виправити, а не
    весь масив перевірок). Якщо в запису після цього фільтру не лишилось
    жодного рядка - весь запис (людина) пропускається, її не видно у файлі.

    Друкує "лоудинг" з відсотками (print_progress) для ДВОХ довгих кроків
    (зчитування .docx ЖБД і перевірка кожного запису рапорту) - підтверджено
    користувачем: без цього незрозуміло, чи скрипт ще працює.

    Охоплює ЛИШЕ пункти, ЛІТЕРАЛЬНО пронумеровані "1"/"2" у самому рапорті
    (_CHECKED_HEADING_NUMBERS) - підтверджено користувачем: решта пунктів
    (навіть якщо теж 30к/100к, напр. ретроактивні поправки за минулі місяці
    під іншим номером) цією перевіркою взагалі не займається."""
    entries = [
        entry for entry in extract_report_entries(report_path)
        if entry["НОМЕР_ПУНКТУ"] in _CHECKED_HEADING_NUMBERS
    ]
    report_dates = [date for entry in entries for _start, dates in _date_subranges_from_period_text(entry["ПЕРІОД"]) for date in dates]
    report_year_month = _majority_year_month(report_dates)
    output_path = output_path or _result_file_name(report_year_month)

    log_war_files = scan_log_war_folder(log_war_folder)
    _warn_if_folder_month_mismatch(report_year_month, log_war_files)

    sections_by_date = {}
    dates = sorted(log_war_files)
    for index, date in enumerate(dates, start=1):
        sections_by_date[date] = extract_log_war_sections(log_war_files[date])
        print_progress(index, len(dates), prefix="Зчитування ЖБД: ")

    old_cutoff = new_cutoff = None
    if logic_mode == LOG_WAR_LOGIC_MIXED:
        transition = _detect_style_transition(log_war_files, sections_by_date)
        old_cutoff, new_cutoff = transition if transition else (None, None)
        _print_style_detection_table(log_war_files, sections_by_date, old_cutoff, new_cutoff)
        if transition:
            print_green(
                f"Змішана логіка: автоматично визначено перехід стилю ЖБД - стара методика "
                f"по {old_cutoff.strftime('%d.%m.%Y')} включно, нова - з {new_cutoff.strftime('%d.%m.%Y')}."
            )
        else:
            print_red(
                "Змішана логіка: не вдалось автоматично визначити межу переходу стилю ЖБД "
                "(немає явного розриву в документах) - застосовано нову методику для всього рапорту."
            )

    # Усі дати КОЖНОЇ людини з УСІХ її записів (30к і 100к разом) -
    # підтверджено користувачем, реальний зразок: людина може перейти з
    # одного пункту в інший БЕЗ ЖОДНОГО розриву (напр. 100к до 04.05, 30к від
    # 05.05) - реальне переміщення в ЖБД задокументоване ОДИН раз на такій
    # межі, а не для КОЖНОГО із двох записів окремо (див. докстрінг
    # _checks_for_entry_old).
    dates_by_pib = {}
    for entry in entries:
        normalized_pib = normalize_name(entry["ПІБ"])
        entry_dates = {date for _start, dates in _date_subranges_from_period_text(entry["ПЕРІОД"]) for date in dates}
        dates_by_pib.setdefault(normalized_pib, set()).update(entry_dates)

    entries_with_checks = []
    total_missing_days = 0
    for index, entry in enumerate(entries, start=1):
        person_dates = dates_by_pib.get(normalize_name(entry["ПІБ"]), frozenset())
        checks = _checks_for_entry(entry, log_war_files, sections_by_date, logic_mode, old_cutoff, new_cutoff, person_dates)
        total_missing_days += _missing_days_count(checks)
        error_checks = [check for check in checks if check["covered"] is False]
        if error_checks:
            entries_with_checks.append({"entry": entry, "checks": error_checks})
        print_progress(index, len(entries), prefix="Перевірка записів рапорту: ")

    _write_log_war_check_result(entries_with_checks, output_path)

    if total_missing_days == 0:
        print_green(f"Перевірка рапорту та ЖБД: усі {len(entries)} записів покриті - {output_path}")
    else:
        print_red(f"Перевірка рапорту та ЖБД: {total_missing_days} дн. без підтвердження з {len(entries)} записів - {output_path}")
    return output_path
