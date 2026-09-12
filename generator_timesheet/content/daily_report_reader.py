import glob
import os
import re
from datetime import datetime, timedelta

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

import content.information_unit_reader as iur

from constants import (
    ARRIVAL_LIST_HEADING_KEYWORDS,
    DATE_RE,
    DEFAULT_STATUS,
    FREEFORM_GROUP_KEYWORD,
    FREEFORM_STATUS_PATTERNS,
    IGNORED_TABLE_HEADING_KEYWORDS,
    PPD_DESTINATION_KEYWORDS,
    SIGNATORY_ROLE_KEYWORDS,
    SIGNATORY_UNIT_NAMES,
    STATUS_HEADING_RULES,
)

# ПІБ у вільному тексті розділу "Поза межами..." - ПРІЗВИЩЕ (великими) + Ім'я +
# По-батькові (з великої літери) - той самий формат, що й у таблицях рапорту,
# лише без табличної структури навколо нього. Не претендує на 100% точність
# (вільний текст, а не дані) - лише допомога людині, що перевірятиме
# REVIEW_FILE_NAME вручну (див. generators/generate_timesheet.py).
_NAME_RE = re.compile(
    r"[А-ЯІЇЄҐ]{2,}(?:[-’'][А-ЯІЇЄҐ][а-яіїєґ’']+)?\s+"
    r"[А-ЯІЇЄҐ][а-яіїєґ’']+\s+[А-ЯІЇЄҐ][а-яіїєґ’']+"
)

# Складніший, БАГАТОПОДІЄВИЙ зворот вільного тексту - "виписаний ... ВЛК ...
# потребує відпустки [за станом здоров'я|для лікування] ..." - ОДНЕ речення
# описує ДВІ послідовні події (день ВЛК/виписки, і наступний день - вже
# відпустка), тож не вкладається в просту "один зворот -> один статус" схему
# FREEFORM_STATUS_PATTERNS вище (див. _discharge_vlk_leave_events нижче).
#
# _NEEDS_LEAVE_RE - ОБОВ'ЯЗКОВА умова (без неї нічого не застосовується
# взагалі) - підтверджує, що речення дійсно описує ЦЕЙ сценарій, а не просто
# випадково згадує ВЛК/виписку з іншого приводу. "\s+" (не буквальний пробіл)
# - у реальних рапортах трапляється подвійний пробіл ("потребує  відпустки").
# "відпустк\w*" (а не буквальне "відпустки") - реальний випадок: рапорт може
# писати як родовий відмінок ("потребує відпустки"), так і знахідний
# ("потребує відпустку") - обидва трапляються в реальних рапортах.
_NEEDS_LEAVE_RE = re.compile(r"потребує\s+відпустк\w*", re.IGNORECASE)

# Дата БІЛЯ згадки ВЛК - "довідки ВЛК №... від ДАТА", "ДАТА за висновком
# ВЛК", "закінчив проходження військово-лікарської комісії ДАТА" (реальний
# випадок - людина завершила саму комісію цього дня, а не отримала довідку за
# її висновком), чи "рішення/висновок військово-лікарської комісії від ДАТА"
# (реальний випадок, підтверджений користувачем: ПОВНА назва "військово-
# лікарської комісії", а не абревіатура "ВЛК" - той САМИЙ сенс, що й
# "ВЛК ... від ДАТА", лише іншим написанням) - ЛИШЕ ЦІ чотири форми (не
# будь-яка дата в реченні), інакше сплуталися б РІЗНІ дати того самого
# речення (напр. дата виписки ПОРЯД з датою самої ВЛК-довідки).
_VLK_DATE_RE = re.compile(
    r"(?:(?P<date_before>\d{2}\.\d{2}\.\d{4})\s*(?:року)?\s*за\s+висновком\s+ВЛК"
    r"|ВЛК\s*(?:№\S+)?\s*від\s*(?P<date_after>\d{2}\.\d{2}\.\d{4})"
    r"|закінчив\w*\s+проходженн\w*\s+військово-лікарс\w*\s+комісі\w*\s+(?P<date_after2>\d{2}\.\d{2}\.\d{4})"
    r"|військово-лікарс\w*\s+комісі\w*\s+від\s*(?P<date_after3>\d{2}\.\d{2}\.\d{4}))",
    re.IGNORECASE,
)
# Дата БІЛЯ "виписаний" - "ДАТА [року] виписаний..." чи "виписаний ... ДАТА"
# (до 40 символів між ними - "з лікувального закладу" тощо).
_DISCHARGE_DATE_RE = re.compile(
    r"(?:(?P<date_before>\d{2}\.\d{2}\.\d{4})\s*(?:року)?\s*виписан"
    r"|виписан.{0,40}?(?P<date_after>\d{2}\.\d{2}\.\d{4}))",
    re.IGNORECASE,
)

# Реальний випадок: "ДАТА виписаний із [заклад]. Після виписки ... не
# прибув ... самовільно залишив військову частину (СЗЧ)." - ДВА окремих
# речення (крапка МІЖ ними), а не один зворот - людину виписано з лікарні,
# вона НЕ прибула до частини й вважається такою, що самовільно залишила
# частину. Пошук - без обмеження довжини проміжку (на відміну від інших
# "виписаний...щось" зворотів вище) - ЄДИНА дата в реченні БІЛЯ "виписан"
# (_DISCHARGE_DATE_RE, ТОЙ САМИЙ regex, що й _discharge_vlk_leave_events),
# тож немає ризику зловити ІНШУ, не ту дату далі в тексті. "вибув у СЗЧ" САМ
# ПО СОБІ вже розпізнається окремо, ЩЕ РАНІШЕ в ланцюжку - _freeform_event
# (FREEFORM_STATUS_PATTERNS вище), тож тут навмисно НЕ дублюється.
_AWOL_AFTER_DISCHARGE_RE = re.compile(r"самовільно\s+залишив\w*\s+військову\s+частину", re.IGNORECASE)

# Реальні випадки, підтверджені користувачем (3 різних рапорти): "ДАТА
# [звання ПІБ] виписаний ... ТА ПРОХОДИТЬ ВЛК" чи "виписаний ... та
# приступив до проходження ВЛК" - людина ЩОЙНО РОЗПОЧАЛА комісію (теперішній/
# доконаний час "проходить"/"приступив", а НЕ завершений факт із документом,
# як _VLK_DATE_RE вище - "закінчив"/"згідно висновку"/"довідки... від ДАТА")
# - день виписки МАЄ бути "ВЛК" (людина вже НЕ просто "ще лікується десь" -
# ШП, а конкретно на комісії), НЕЗАЛЕЖНО від того, як саме знайдено дату
# цього дня - через _DISCHARGE_DATE_RE (дата біля "виписан") чи через дату
# ПЕРЕД самим ПІБ (коли дата на початку абзацу, а "виписан" далі за ПІБ -
# _DISCHARGE_DATE_RE тоді закороткий, той самий прийом, що й
# _discharge_then_awol_event). Без цього застосовувалась хибна "ШП" (якщо
# дата ВСЕ Ж знайшлась через _DISCHARGE_DATE_RE) або сценарій узагалі
# лишався нерозпізнаним (якщо дата - лише перед ПІБ) - обидва реальні
# випадки, підтверджені користувачем.
_VLK_IN_PROGRESS_RE = re.compile(r"проходить\s+ВЛК|приступив\w*\s+до\s+проходженн\w*\s+ВЛК", re.IGNORECASE)

# Якщо ТЕ САМЕ речення десь ще й прямо згадує "відпустки/відпустку за станом
# здоров'я" (не обов'язково зворотом "вибув у ..." як у FREEFORM_STATUS_PATTERNS
# вище - тут ця фраза часто описує стан ДО/ПІД ЧАС самої події, напр. "під час
# відпустки за станом здоров'я ... потребує відпустки для лікування") - парна
# подія "відпустка" (наступний день, у _discharge_vlk_leave_events нижче) це
# "ВПСЗ", а не просто "ВП": підтверджено користувачем на реальному випадку, де
# справжній статус наступного дня (підтверджений ОКРЕМИМ реченням наступного
# рапорту) виявився саме "ВПСЗ". Якорна подія (сам день ВЛК/виписки) на це НЕ
# зважає - вона лишається "ВЛК"/"ШП" незалежно.
_HEALTH_LEAVE_MENTION_RE = re.compile(r"відпуст\w*\s+за\s+станом\s+здоров['’]?я", re.IGNORECASE)

# Реальний випадок, підтверджений користувачем: "потребує відпустки ДЛЯ
# ЛІКУВАННЯ ПІСЛЯ ПОРАНЕННЯ на 30 календарних днів" - ЩЕ ОДНЕ, конкретніше
# формулювання причини відпустки (на відміну від "за станом здоров'я" вище) -
# збігається з уже підтримуваним статусом "ВПБП" ("Відпустка для лікування
# після тяжкого поранення", STATUS_MEANINGS, constants.py) - "тяжкого" НЕОБОВ'
# ЯЗКОВЕ слово (реальний випадок - без нього), той самий сенс коротшим
# формулюванням. Перевіряється РАНІШЕ за _HEALTH_LEAVE_MENTION_RE у
# _discharge_vlk_leave_events нижче (специфічніший статус перемагає) - без
# ЦЬОГО парна подія "відпустка" мовчки лишалась загальним "ВП", хоча речення
# прямо називає причину.
_WOUND_TREATMENT_LEAVE_MENTION_RE = re.compile(r"лікуванн\w*\s+після\s+(?:тяжкого\s+)?поранення", re.IGNORECASE)

# Реальний випадок: НАСТУПНОГО дня після рапорту "потребує відпустки за
# висновком ВЛК ..." (без причини - див. _discharge_vlk_leave_events вище,
# leave_status тоді - "ВП", бо ЦЕ речення саме по собі не згадує "за станом
# здоров'я") - ОКРЕМИЙ рапорт ПІДТВЕРДЖУЄ причину ОКРЕМИМ реченням - "З ДАТА
# відпустка за станом здоров'я ..." АБО "відпустка за станом здоров'я З ДАТА"
# (обидва порядки слів трапляються в реальних рапортах) - без "потребує" (той
# стан уже позаду) і без "вибув у відпустку" (не той зворот, що в
# FREEFORM_STATUS_PATTERNS). Кожна гілка альтернативи (|) ВИМАГАЄ власну дату
# - навмисно, щоб ГОЛЕ "потребує відпустки за станом здоров'я" (без жодної
# дати - test_freeform_needs_leave_without_discharge_or_vlk_date_is_not_auto_
# applied) і далі НЕ підхоплювалось тут (лишається на ручну перевірку, як і
# раніше). _discharge_vlk_leave_events НЕ бачить ЦЕЙ, ІНШИЙ рапорт (кожен
# рапорт розбирається окремо) - подія з ЦЬОГО речення ("ВПСЗ") ОКРЕМО
# потрапляє в apply_events, де вона АВТОМАТИЧНО (без ручної перевірки)
# переможе вже наявну "ВП" на ТУ САМУ дату - той самий override, що й
# STATUS_CONFLICT_OVERRIDES (constants.py).
_CONFIRMED_HEALTH_LEAVE_RE = re.compile(
    r"з\s+(?P<date1>\d{2}\.\d{2}\.\d{4})\s+відпуст\w*\s+за\s+станом\s+здоров['’]?я"
    r"|відпуст\w*\s+за\s+станом\s+здоров['’]?я\s+з\s+(?P<date2>\d{2}\.\d{2}\.\d{4})",
    re.IGNORECASE,
)

# Реальний випадок: "з ДАТА1 [по ДАТА_кінця] у відпустці за станом здоров'я,
# з ДАТА2 [року] направлений на проходження військово-лікарської комісії" -
# ОДНЕ речення описує ДВІ послідовні події з ОБОМА датами ПОЧАТКУ явно
# вказаними: людина на ВПСЗ (health leave) з ДАТА1, а з ДАТА2 - викликана на
# ВЛК (комісія визначає подальший статус). "[\s\S]{0,25}?" - короткий (до 25
# символів, включно з можливим переносом рядка всередині абзацу - "\s\S", а
# не звичайний "." - реальний випадок переносу МІЖ "з ДАТА1" і "по ДАТА_кінця"
# в самому рапорті) НЕЖАДІБНИЙ проміжок між датою й фразою - НАВМИСНО
# ОБМЕЖЕНИЙ довжиною (а не "(?:\S+\s+)?"/необмежений ".*?", як в інших
# зворотах цього файлу): без обмеження ДАТА1 (значно РАНІШЕ в реченні) сама
# по собі "дотяглася" б і до фрази ВЛК (де насправді потрібна ДАТА2, ближча
# до неї) - 25 символів свідомо достатньо для "по ДАТА_кінця"/"року
# направлений", але замало, щоб перестрибнути ЦІЛУ проміжну фразу
# відпустки/ВЛК і зловити НЕ ТУ дату.
_HEALTH_LEAVE_DATED_RE = re.compile(
    r"(?P<date>\d{2}\.\d{2}\.\d{4})[\s\S]{0,25}?відпуст\w*\s+за\s+станом\s+здоров['’]?я",
    re.IGNORECASE,
)
_VLK_REFERRAL_DATED_RE = re.compile(
    r"(?P<date>\d{2}\.\d{2}\.\d{4})[\s\S]{0,25}?на\s+проходженн\w*\s+військово-лікарс\w*\s+комісі\w*",
    re.IGNORECASE,
)

# Реальний випадок: "ДАТА1 звання ПІБ завершив/закінчив проходження
# військово-лікарської комісії. З ДАТА2 направляється у відпустку ..." -
# ТОЙ САМИЙ факт, що й "закінчив ВЛК ДАТА, потребує відпустки за станом
# здоров'я ..." (_discharge_vlk_leave_events, date_after2 форма), лише
# переказаний у НАСТУПНОМУ рапорті ІНШИМИ словами - без "потребує"
# (_NEEDS_LEAVE_RE тут не спрацює) і без "за станом здоров'я"
# (_HEALTH_LEAVE_MENTION_RE тут теж не спрацює - тут "відпустку ДЛЯ
# ЛІКУВАННЯ"), тож жоден наявний механізм цього не бачить.
_VLK_COMPLETED_MENTION_RE = re.compile(
    r"(?:завершив|закінчив)\w*\s+проходженн\w*\s+військово-лікарс\w*\s+комісі\w*", re.IGNORECASE
)
# Реальний випадок: "ДАТА1 звання ПІБ виписаний із [заклад], з ДАТА2
# направляється у відпустку для лікування ..." - ТОЙ САМИЙ "З ДАТА
# направляється у відпустку" зворот, лише медичний контекст - "виписаний" (з
# лікарні), а не "завершив ВЛК" - ТОЙ САМИЙ принцип, лише ІНШЕ джерело
# контексту. Бере лише СТЕМ "виписан" (присутність, а не власна дата) - сама
# дата відпустки бере з _DATED_TREATMENT_LEAVE_RE нижче (не з
# _DISCHARGE_DATE_RE - той очікує дату ЩІЛЬНО біля "виписан"/перед ПІБ, а тут
# дата "виписаний" відсутня зовсім чи заплутана переліком ІНШИХ дат/установ
# у тому самому реченні).
_DISCHARGE_MENTION_RE = re.compile(r"виписан\w*", re.IGNORECASE)
# ОБИДВІ (медичний контекст: _VLK_COMPLETED_MENTION_RE АБО
# _DISCHARGE_MENTION_RE) І ЦЯ фрази МАЮТЬ бути присутні РАЗОМ (див.
# _dated_treatment_leave_event) - навмисно вужче, ніж будь-яка сама по собі
# (голе "З ДАТА направляється у відпустку" саме по собі занадто неоднозначне
# - могла бути звичайна, немедична відпустка). Друга альтернатива - "ДАТА
# вибув у відпустку" (реальний випадок, підтверджений користувачем: "...
# закінчив проходження військово-лікарської комісії, ДАТА вибув у
# відпустку ...") - ІНШЕ дієслово ("вибув", а не "направля-") і дата ПЕРЕД
# дієсловом, БЕЗ прийменника "з" - той самий сенс, окрема named-група
# "date2" (щоб не плутати з "date" першої альтернативи).
_DATED_TREATMENT_LEAVE_RE = re.compile(
    r"з\s+(?P<date>\d{2}\.\d{2}\.\d{4})\s+направля\w*\s+у\s+відпустку"
    r"|(?P<date2>\d{2}\.\d{2}\.\d{4})\s+вибув\w*\s+у\s+відпустку",
    re.IGNORECASE,
)


def _iter_block_items(doc):
    """Абзаци й таблиці документа в порядку появи (спільний прохід по
    doc.element.body - стандартний рецепт python-docx, doc.paragraphs/doc.tables
    окремо не зберігають взаємний порядок)."""
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def _parse_date(text):
    match = DATE_RE.search(text or "")
    return datetime.strptime(match.group(), "%d.%m.%Y").date() if match else None


def report_date_from_filename(file_path):
    return _parse_date(os.path.basename(file_path))


def latest_report_date(report_dir):
    """Дата НАЙПІЗНІШОГО .docx рапорту в report_dir (за датою в ІМЕНІ
    файлу - report_date_from_filename) - lock-файли Word ("~$...", той
    самий принцип, що й generators/generate_timesheet.py::_build_timesheet)
    і файли БЕЗ розпізнаваної дати пропускаються. None, якщо report_dir не
    містить жодного такого файлу.

    За прямою вказівкою користувача - content/rop_vop_statement.py-подібні
    генератори (де ПОТРІБЕН ОДИН конкретний "цільовий" місяць, на відміну
    від generate_timesheet/generate_payments_timesheet, які просто
    накопичують УСІ наявні дати як окремі колонки) мають орієнтуватись на
    ЦЮ дату (найпізніший рапорт, що РЕАЛЬНО є), а НЕ на date.today() -
    реальний випадок: сьогодні вже вересень, але жодних рапортів/файлів
    information_unit/schedule за вересень ще не подано (типова затримка
    подачі на початку місяця) - генерація "за поточний місяць" (вересень)
    дала б ПОРОЖНІЙ результат ("жодна людина не має днів РОП/ВОП"), хоча
    дані за СЕРПЕНЬ (останній місяць, за який дані РЕАЛЬНО є) цілком готові
    для генерації."""
    dates = []
    for path in glob.glob(os.path.join(report_dir, "*.docx")):
        if iur._is_office_lock_file(path):
            continue
        report_date = report_date_from_filename(path)
        if report_date is not None:
            dates.append(report_date)
    return max(dates) if dates else None


def _status_for_heading(heading_text):
    lowered = (heading_text or "").lower()
    for keyword, status in STATUS_HEADING_RULES.items():
        if keyword in lowered:
            return status
    return None


def _is_ignored_table_heading(heading_text):
    """True, якщо heading_text містить одне з IGNORED_TABLE_HEADING_
    KEYWORDS (constants.py) - за прямою вказівкою користувача, ВЕСЬ пункт
    рапорту з ЦИМ заголовком МАЄ ігноруватись ПОВНІСТЮ (_table_events -
    жодної події, жодного запису на ручну перевірку), а НЕ отримувати
    статус (STATUS_HEADING_RULES) чи потрапляти в unresolved як
    "невідомий заголовок"."""
    lowered = (heading_text or "").lower()
    return any(keyword in lowered for keyword in IGNORED_TABLE_HEADING_KEYWORDS)


def _is_arrival_list_heading(text):
    """Заголовок-СПИСОК "ПРИБУЛИ до пункту постійно(ї) дислокації військової
    частини ...:" (ARRIVAL_LIST_HEADING_KEYWORDS, constants.py) - за НИМ,
    без таблиці, йдуть окремі абзаци "звання ПІБ" (extract_status_events
    нижче застосовує DEFAULT_STATUS кожному з них автоматично)."""
    lowered = text.lower()
    return all(keyword in lowered for keyword in ARRIVAL_LIST_HEADING_KEYWORDS)


def _is_bare_status_heading(text):
    """Текст - ЛИШЕ ключове слово підзаголовка (STATUS_HEADING_RULES),
    можливо з кінцевою двокрапкою, БЕЗ жодного додаткового змісту (реальний
    випадок: розділ вільного тексту "Поза межами..." має міні-підзаголовок
    "СЗЧ:" ПЕРЕД реченням про конкретну людину - сам підзаголовок описує НЕ
    подію конкретної людини, а групу, тож не повинен потрапляти на ручну
    перевірку окремо, так само як підзаголовки таблиць не перевіряються
    окремо від самої таблиці). ТОЧНИЙ (а не частковий/substring, як в
    _status_for_heading - той шукає ключове слово ВСЕРЕДИНІ довшого тексту
    таблиці) збіг - навмисно строгіше, щоб не сховати від людини СПРАВЖНІ
    речення, які лише ЗГАДУЮТЬ ключове слово мимохідь."""
    normalized = text.strip().rstrip(":").strip().lower()
    return normalized in STATUS_HEADING_RULES


def _unresolved(reason, pib_raw, report_date, **extra):
    return {"reason": reason, "pib_raw": pib_raw, "report_date": report_date, **extra}


def _freeform_event(text, name_match, report_date):
    """Подія зі СТІЙКОГО звороту вільного тексту розділу "Поза межами..."
    (FREEFORM_STATUS_PATTERNS) - None, якщо ПІБ не розпізнано чи жоден
    візерунок не підійшов (лишається для ручної перевірки - див.
    extract_status_events)."""
    if name_match is None:
        return None
    for pattern, status in FREEFORM_STATUS_PATTERNS:
        match = pattern.search(text)
        if match:
            effective_date = _parse_date(match.group("date")) or report_date
            return {
                "pib_raw": name_match.group(), "status": status, "effective_date": effective_date,
                "heading": FREEFORM_GROUP_KEYWORD, "report_date": report_date,
            }
    return None


def _confirmed_health_leave_event(text, name_match, report_date):
    """Подія зі "підтверджувального" речення НАСТУПНОГО рапорту - "З ДАТА
    відпустка за станом здоров'я ..." чи "відпустка за станом здоров'я З
    ДАТА" (_CONFIRMED_HEALTH_LEAVE_RE, див. докстрінг вище константи) -
    "ВПСЗ" на дату, вказану ПРЯМО в реченні (не report_date - тут вона
    завжди явна). None, якщо ПІБ не розпізнано чи зворот не підійшов
    (лишається на _discharge_vlk_leave_events/ручну перевірку - див.
    extract_status_events)."""
    if name_match is None:
        return None
    match = _CONFIRMED_HEALTH_LEAVE_RE.search(text)
    if match is None:
        return None
    date_text = match.group("date1") or match.group("date2")
    return {
        "pib_raw": name_match.group(), "status": "ВПСЗ", "effective_date": _parse_date(date_text) or report_date,
        "heading": FREEFORM_GROUP_KEYWORD, "report_date": report_date,
    }


def _dated_treatment_leave_event(text, name_match, report_date):
    """Подія зі "підтверджувального" речення - "ДАТА1 [звання ПІБ]
    завершив/закінчив проходження військово-лікарської комісії" АБО "ДАТА1
    [звання ПІБ] виписаний із [заклад]" - ТА, В ТОМУ Ж РЕЧЕННІ, "з ДАТА2
    направляється у відпустку ..." АБО "ДАТА2 вибув у відпустку ..."
    (_VLK_COMPLETED_MENTION_RE чи _DISCHARGE_MENTION_RE РАЗОМ із
    _DATED_TREATMENT_LEAVE_RE, див. докстрінг вище констант) - "ВПСЗ" (за
    прямою вказівкою користувача - той самий статус, що й у "завершив ВЛК
    ДАТА, потребує відпустки за станом здоров'я..." формі
    _discharge_vlk_leave_events, лише підтверджений ІНШИМИ словами - ВЛК
    завершено чи виписка з лікарні, в ОБОХ випадках людина вже вийшла з
    медичного процесу й переходить на відпустку) на ДАТУ2, вказану ПРЯМО в
    реченні (не report_date - тут вона завжди явна). Медичний контекст
    (ВЛК чи виписка) І "ДАТА2 [направляється/вибув] у відпустку" МАЮТЬ
    бути присутні РАЗОМ - навмисно вужче, ніж "ДАТА направляється/вибув у
    відпустку" саме по собі (могла бути звичайна, немедична відпустка).
    None, якщо ПІБ не розпізнано чи зворот не підійшов (лишається на
    ручну перевірку - див. extract_status_events)."""
    if name_match is None or not (_VLK_COMPLETED_MENTION_RE.search(text) or _DISCHARGE_MENTION_RE.search(text)):
        return None
    match = _DATED_TREATMENT_LEAVE_RE.search(text)
    if match is None:
        return None
    date_text = match.group("date") or match.group("date2")
    return {
        "pib_raw": name_match.group(), "status": "ВПСЗ",
        "effective_date": _parse_date(date_text) or report_date,
        "heading": FREEFORM_GROUP_KEYWORD, "report_date": report_date,
    }


def _vlk_in_progress_anchor_event(text, name_match, report_date, leave_date):
    """Реальні випадки, підтверджені користувачем: "ДАТА1 [звання ПІБ]
    виписаний ... та ПРОХОДИТЬ ВЛК/приступив до проходження ВЛК, з ДАТА2
    відпустка/направляється у відпустку ..." - на відміну від
    _dated_treatment_leave_event/_confirmed_health_leave_event (припускають,
    що людина ВЖЕ вийшла з медичного процесу, тож дають ЛИШЕ подію
    "відпустка" на ДАТУ2), явна згадка "проходить"/"приступив до
    проходження" ВЛК (_VLK_IN_PROGRESS_RE, теперішній/недавній час, а НЕ
    завершений факт - "закінчив"/"завершив" ВЛК, _VLK_COMPLETED_MENTION_RE,
    навмисно НЕ входить сюди: підтверджено користувачем на РЕАЛЬНОМУ
    випадку test_completed_vlk_then_dated_leave_sentence_produces_vpsz_
    event - для "завершеної" ВЛК ОКРЕМИЙ день "ВЛК" НЕ потрібен, лише
    "ВПСЗ" на дату відпустки) означає, що ДАТА1 (день виписки/початку
    комісії) сама по собі МАЄ стати ОКРЕМОЮ подією "ВЛК" - інакше вона
    мовчки лишається попереднім статусом (не "ВПСЗ" і не "ВЛК", просто те,
    що було раніше).

    ДАТА1 шукається ТИМ САМИМ способом, що й у _discharge_vlk_leave_events:
    спершу _DISCHARGE_DATE_RE (дата біля "виписан"), запасний варіант - дата
    ПЕРЕД самим ПІБ (коли дата на початку абзацу, а "виписан" - вже ПІСЛЯ
    ПІБ, задалеко для _DISCHARGE_DATE_RE).

    Повертає подію "ВЛК" на ДАТУ1, або None - якщо немає згадки "проходить
    ВЛК", немає жодної придатної дати, чи ДАТА1 збігається з leave_date (уже
    ОДНА дата в реченні - друга подія на ТУ САМУ дату була б суперечністю)."""
    if not _VLK_IN_PROGRESS_RE.search(text):
        return None

    discharge_match = _DISCHARGE_DATE_RE.search(text)
    if discharge_match is not None:
        date_text = discharge_match.group("date_before") or discharge_match.group("date_after")
        anchor_date = _parse_date(date_text) or report_date
    else:
        leading_dates = re.findall(r"\d{2}\.\d{2}\.\d{4}", text[:name_match.start()])
        if not leading_dates:
            return None
        anchor_date = _parse_date(leading_dates[-1]) or report_date

    if anchor_date == leave_date:
        return None
    return {
        "pib_raw": name_match.group(), "status": "ВЛК", "effective_date": anchor_date,
        "heading": FREEFORM_GROUP_KEYWORD, "report_date": report_date,
    }


def _health_leave_then_vlk_events(text, name_match, report_date):
    """(events) - 0 чи 2 події зі складеного речення "з ДАТА1 у відпустці за
    станом здоров'я та ДАТА2 на проходження військово-лікарської комісії"
    (реальний випадок, підтверджено користувачем) - людина на "ВПСЗ" з
    ДАТА1, а з ДАТА2 - "ВЛК" (комісія визначає подальше). ОБИДВІ фрази (і
    ОБИДВІ власні дати - _HEALTH_LEAVE_DATED_RE/_VLK_REFERRAL_DATED_RE) МАЮТЬ
    бути присутні РАЗОМ - навмисно вужче, ніж будь-яка сама по собі (звичайна
    відпустка чи ВЛК з ІНШОЇ причини були б занадто неоднозначні), тож без
    обох - лишається на ручну перевірку, як і раніше."""
    if name_match is None:
        return []
    leave_match = _HEALTH_LEAVE_DATED_RE.search(text)
    vlk_match = _VLK_REFERRAL_DATED_RE.search(text)
    if leave_match is None or vlk_match is None:
        return []
    pib_raw = name_match.group()
    return [
        {
            "pib_raw": pib_raw, "status": "ВПСЗ",
            "effective_date": _parse_date(leave_match.group("date")) or report_date,
            "heading": FREEFORM_GROUP_KEYWORD, "report_date": report_date,
        },
        {
            "pib_raw": pib_raw, "status": "ВЛК",
            "effective_date": _parse_date(vlk_match.group("date")) or report_date,
            "heading": FREEFORM_GROUP_KEYWORD, "report_date": report_date,
        },
    ]


def _discharge_vlk_leave_events(text, name_match, report_date):
    """(events) - 0, 1 чи 2 події зі складнішого, БАГАТОПОДІЄВОГО звороту
    "виписаний ... ВЛК ... потребує відпустки ..." (підтверджено користувачем
    на реальних випадках з resources/report/*.docx, розділ "Поза межами").

    Без _NEEDS_LEAVE_RE ("потребує відпустки") - ЖОДНОЇ події (надто
    неоднозначно, щоб застосовувати автоматично лише за згадкою ВЛК/виписки
    самих по собі). З нею:

    - якщо є дата БІЛЯ ЗГАДКИ ВЛК ЯК ДОКУМЕНТА/РІШЕННЯ (_VLK_DATE_RE,
      "довідки ВЛК №... від ДАТА", "ДАТА за висновком ВЛК", чи "рішення/
      висновок військово-лікарської комісії від ДАТА" (повна назва, не
      абревіатура) - date_before/date_after/date_after3) - ДВІ події:
      якорна цього дня - статус "ВЛК" (комісія ще "в процесі"
      адміністративно того дня), а відпустка - НАСТУПНОГО дня (підтверджено
      користувачем на реальних випадках 04.08.2026/07.08.2026/31.08.2026);
    - якщо ж дата - від ЗАВЕРШЕННЯ самого проходження комісії ("закінчив/
      завершив проходження ... комісії ДАТА" - date_after2) - ОДНА подія:
      відпустка починається ТОГО САМОГО дня, без "дня ВЛК" (реальний випадок,
      підтверджений користувачем: комісію вже ПРОЙДЕНО ("завершив") того ж
      дня, тож "день ВЛК" уже позаду - ставити тут окремий статус "ВЛК"
      НЕПРАВИЛЬНО, на відміну від date_before/date_after форм вище, де ДАТА -
      лише дата ДОКУМЕНТА, а не факт завершення самого проходження);
    - інакше, якщо є дата БІЛЯ "виписаний" (_DISCHARGE_DATE_RE) - якорна
      подія цього дня - статус "ШП" (людина того дня ще не покинула медичну
      систему - лише виписана з ОДНОГО закладу, конкретики про ВЛК ще нема),
      АБО "ВЛК", якщо речення ТАКОЖ прямо каже, що людина ЩОЙНО РОЗПОЧАЛА
      комісію того ж дня ("проходить ВЛК"/"приступив до проходження ВЛК" -
      _VLK_IN_PROGRESS_RE, теперішній/недавній час, а НЕ завершений факт із
      документом, як _VLK_DATE_RE вище) - відпустка в ОБОХ випадках -
      наступного дня;
    - інакше, якщо речення прямо каже "проходить ВЛК"/"приступив до
      проходження ВЛК" (_VLK_IN_PROGRESS_RE) І є дата ПЕРЕД самим ПІБ (той
      самий прийом, що й _discharge_then_awol_event - "ДАТА звання ПІБ
      виписаний...", коли дата стоїть на початку абзацу, а "виписан" - вже
      ПІСЛЯ ПІБ, задалеко для _DISCHARGE_DATE_RE) - якорна подія цього дня -
      статус "ВЛК", відпустка - наступного дня;
    - немає ЖОДНОЇ з цих дат - подій немає взагалі (лишається на ручну
      перевірку, як і раніше).

    Подія "відпустка" - "ВПСЗ", якщо речення десь ще й прямо згадує
    "відпустки за станом здоров'я" (_HEALTH_LEAVE_MENTION_RE) - ПЕРЕВІРЯЄТЬСЯ
    ПЕРШИМ, навіть якщо речення ТАКОЖ згадує "лікування після поранення" -
    реальний випадок, підтверджений користувачем: речення МОЖЕ згадувати
    ОБИДВІ фрази одразу ("... під час відпустки ЗА СТАНОМ ЗДОРОВ'Я ПО
    ПОРАНЕННЮ, ... потребує відпустки ДЛЯ ЛІКУВАННЯ ПІСЛЯ ПОРАНЕННЯ ..."),
    і саме "ВПСЗ" - той статус, що його ПІЗНІШЕ підтверджує ОКРЕМЕ речення
    НАСТУПНОГО рапорту ("вибув у відпустку за станом здоров'я") - "ВПБП"
    (_WOUND_TREATMENT_LEAVE_MENTION_RE) застосовується ЛИШЕ якщо "за станом
    здоров'я" У РЕЧЕННІ ВЗАГАЛІ НЕМАЄ (інакше довелось би на кожну пару
    "ВПБП"/"ВПСЗ" писати запис у STATUS_CONFLICT_OVERRIDES замість того, щоб
    просто не створювати фальшивої суперечності одразу); інакше - "ВП"."""
    if name_match is None or not _NEEDS_LEAVE_RE.search(text):
        return []

    pib_raw = name_match.group()
    if _HEALTH_LEAVE_MENTION_RE.search(text):
        leave_status = "ВПСЗ"
    elif _WOUND_TREATMENT_LEAVE_MENTION_RE.search(text):
        leave_status = "ВПБП"
    else:
        leave_status = "ВП"
    in_progress_vlk = _VLK_IN_PROGRESS_RE.search(text)

    vlk_match = _VLK_DATE_RE.search(text)
    if vlk_match:
        date_text = (
            vlk_match.group("date_before") or vlk_match.group("date_after")
            or vlk_match.group("date_after2") or vlk_match.group("date_after3")
        )
        anchor_date = _parse_date(date_text) or report_date
        if vlk_match.group("date_after2"):
            return [{
                "pib_raw": pib_raw, "status": leave_status, "effective_date": anchor_date,
                "heading": FREEFORM_GROUP_KEYWORD, "report_date": report_date,
            }]
        anchor_status = "ВЛК"
    else:
        discharge_match = _DISCHARGE_DATE_RE.search(text)
        if discharge_match is not None:
            anchor_status = "ВЛК" if in_progress_vlk else "ШП"
            date_text = discharge_match.group("date_before") or discharge_match.group("date_after")
            anchor_date = _parse_date(date_text) or report_date
        elif in_progress_vlk:
            leading_dates = re.findall(r"\d{2}\.\d{2}\.\d{4}", text[:name_match.start()])
            if not leading_dates:
                return []
            anchor_status = "ВЛК"
            anchor_date = _parse_date(leading_dates[-1]) or report_date
        else:
            return []

    return [
        {
            "pib_raw": pib_raw, "status": anchor_status, "effective_date": anchor_date,
            "heading": FREEFORM_GROUP_KEYWORD, "report_date": report_date,
        },
        {
            "pib_raw": pib_raw, "status": leave_status, "effective_date": anchor_date + timedelta(days=1),
            "heading": FREEFORM_GROUP_KEYWORD, "report_date": report_date,
        },
    ]


def _discharge_then_awol_event(text, name_match, report_date):
    """(event чи None) - зворот "ДАТА звання ПІБ виписаний із [заклад]. Після
    виписки ... не прибув ... самовільно залишив військову частину (СЗЧ)."
    (реальний випадок, підтверджено користувачем) - ДВА окремих речення
    (людину виписано з лікарні, а до частини вона не прибула - самовільно
    залишила частину), АЛЕ ОДНА подія: статус "СЗЧ" на дату ВИПИСКИ.

    Дата тут стоїть ПЕРЕД ПІБ (на початку абзацу - "ДАТА звання ПІБ
    виписаний..."), а не одразу біля самого "виписан" - тож _DISCHARGE_DATE_RE
    (розрахований на дату ЩІЛЬНО біля "виписан") тут не підходить: пробували
    б замінити на нього бонований проміжок, який довелось би підбирати "на
    око" під довжину ПІБ. Натомість - дата шукається в частині тексту ПЕРЕД
    самим ПІБ (name_match.start()) - той самий принцип, що й
    _DISCHARGE_DATE_RE, лише прив'язаний до позиції ПІБ, а не слова
    "виписан" (перевірений реальний формат рапортів - абзац завжди починається
    "ДАТА звання ПІБ ..."). Якщо перед ПІБ дат кілька - береться ОСТАННЯ
    (найближча до ПІБ). Якщо перед ПІБ дати немає взагалі - як запасний
    варіант, _DISCHARGE_DATE_RE (дата щільно біля "виписан") - охоплює й
    протилежний порядок ("виписаний ... ДАТА").

    На відміну від _discharge_vlk_leave_events, тут НЕМАЄ "потребує
    відпустки", тож той механізм цього звороту не бачить, а сама по собі
    згадка "виписаний" без AWOL-фрази (_AWOL_AFTER_DISCHARGE_RE) - надто
    неоднозначна, щоб застосовувати статус "СЗЧ" автоматично лише за нею.
    None, якщо ПІБ не розпізнано, AWOL-фрази немає чи дати виписки немає
    (лишається на ручну перевірку - див. extract_status_events)."""
    if name_match is None or not _AWOL_AFTER_DISCHARGE_RE.search(text):
        return None
    leading_dates = re.findall(r"\d{2}\.\d{2}\.\d{4}", text[:name_match.start()])
    if leading_dates:
        date_text = leading_dates[-1]
    else:
        discharge_match = _DISCHARGE_DATE_RE.search(text)
        if discharge_match is None:
            return None
        date_text = discharge_match.group("date_before") or discharge_match.group("date_after")
    return {
        "pib_raw": name_match.group(), "status": "СЗЧ", "effective_date": _parse_date(date_text) or report_date,
        "heading": FREEFORM_GROUP_KEYWORD, "report_date": report_date,
    }


def _is_signatory_paragraph(block, text):
    """Рядок підпису наприкінці рапорту ("Тимчасово виконуючий обов'язки
    командира ...", "Командир ..." + звання/ім'я) - не стосується жодної
    конкретної людини зі складу, тож не повинен потрапляти в unresolved
    розділу "Поза межами..." як "текст без розпізнаного ПІБ". Перевіряється
    для КОЖНОГО абзаца рапорту (не лише в розділі "Поза межами...") - інакше
    рапорт БЕЗ жодної згадки в "Поза межами" взагалі не мав би нагоди
    розпізнати підпис.

    Розпізнається, ЯКЩО абзац має стиль "No Spacing" (у реальних рапортах -
    завжди САМЕ рядок підпису, підтверджено на 8 файлах) АБО текст одночасно
    згадує роль підписанта (SIGNATORY_ROLE_KEYWORDS - стабільні слова, не
    специфічні до підрозділу) і назву/скорочення ВЛАСНОГО підрозділу
    (SIGNATORY_UNIT_NAMES, resources/data.json). Обидві ознаки потрібні
    РАЗОМ (а не сама назва підрозділу) - інакше законна згадка ВЛАСНОГО
    підрозділу в реченні про людину (напр. "прибув з ... [назва підрозділу]")
    хибно зникла б із результату.

    Виняток - "Командиру ..." (АДРЕСАТ рапорту, ЗАВЖДИ перший рядок, давальний
    відмінок) - НЕ підпис, хоч і згадує ту саму роль ("командир") і підрозділ:
    "командир" - підрядок і "командира" (родовий, підпис), і "командиру"
    (давальний, адресат), тож без цього винятку перший рядок рапорту сам
    хибно зупинив би розбір усього іншого."""
    if block.style is not None and block.style.name == "No Spacing":
        return True
    lowered = text.lower()
    if lowered.startswith("командиру"):
        return False
    has_role = any(keyword in lowered for keyword in SIGNATORY_ROLE_KEYWORDS)
    has_unit = any(name.lower() in lowered for name in SIGNATORY_UNIT_NAMES)
    return has_role and has_unit


def _detect_pib_and_date_columns(header_row):
    """Позиції колонок ПІБ і дати (прибуття/вибуття) у ЦІЙ конкретній таблиці -
    за назвою заголовка (як і в generator_br_and_report_for_money), бо набір і
    порядок колонок відрізняється між пунктами рапорту ("Дата прибуття" проти
    "Дата вибуття", "Тип відпустки, куди вибув" проти просто "Звідки вибув"
    тощо). date_col - None, якщо в таблиці взагалі немає колонки з датою
    (тоді ефективною датою стає дата самого рапорту - див. _table_events)."""
    header_texts = [cell.text.strip().lower() for cell in header_row.cells]
    pib_col = next((i for i, text in enumerate(header_texts) if "прізвищ" in text), None)
    date_col = next((i for i, text in enumerate(header_texts) if "дата" in text), None)
    return pib_col, date_col


def _destination_column(header_row):
    """Позиція колонки призначення ("Куди прибув"/"Куди вибув"/"Тип
    відпустки, куди прибув" тощо) - substring "куди" в тексті заголовка, той
    самий принцип, що й pib_col/date_col у _detect_pib_and_date_columns.
    None, якщо в таблиці немає такої колонки."""
    header_texts = [cell.text.strip().lower() for cell in header_row.cells]
    return next((i for i, text in enumerate(header_texts) if "куди" in text), None)


def _is_ppd_destination(text):
    """True, якщо текст колонки "Куди ..." згадує ПОСТІЙНИЙ пункт
    дислокації (PPD_DESTINATION_KEYWORDS, constants.py) - substring-пошук,
    без урахування регістру, той самий принцип, що й _status_for_heading/
    _is_ignored_table_heading."""
    lowered = text.lower()
    return any(keyword in lowered for keyword in PPD_DESTINATION_KEYWORDS)


def _table_events(table, heading_text, report_date):
    """(events, unresolved) для ОДНІЄЇ таблиці рапорту.

    За прямою вказівкою користувача - heading_text із
    IGNORED_TABLE_HEADING_KEYWORDS (constants.py) МАЄ ігноруватись
    ПОВНІСТЮ - ([], []) одразу, БЕЗ жодної події й БЕЗ жодного запису на
    ручну перевірку (перевіряється РАНІШЕ за все інше нижче).

    Інакше: events - людина отримує статус, визначений за heading_text
    (_status_for_heading) - ОКРІМ коли колонка "Куди ..." ЦЬОГО рядка прямо
    згадує постійний пункт дислокації (PPD_DESTINATION_KEYWORDS,
    constants.py) - тоді статус ЦЬОГО рядка "ППД", незалежно від
    заголовка-замовчування (реальний випадок, підтверджений користувачем:
    рядок під "З відпустки:" - звичайно DEFAULT_STATUS - чия колонка "Куди
    прибув" каже "...до пункту постійної дислокації..."). unresolved - той
    самий рядок, якщо: (1) heading_text не відповідає жодному з
    STATUS_HEADING_RULES (новий/незнайомий підзаголовок - статус невідомий,
    рядок не втрачається, а лишається для ручної перевірки); (2) колонку
    ПІБ у таблиці взагалі не вдалось визначити."""
    if _is_ignored_table_heading(heading_text):
        return [], []

    status = _status_for_heading(heading_text)
    pib_col, date_col = _detect_pib_and_date_columns(table.rows[0])
    destination_col = _destination_column(table.rows[0])
    events, unresolved = [], []

    if pib_col is None:
        unresolved.append(_unresolved(f"Не вдалося визначити колонку ПІБ у таблиці під заголовком \"{heading_text.strip()}\".", None, report_date))
        return events, unresolved

    for row in table.rows[1:]:
        cells = row.cells
        pib_raw = cells[pib_col].text.strip()
        if not pib_raw:
            continue

        if status is None:
            unresolved.append(_unresolved(f"Невідомий заголовок пункту рапорту: \"{heading_text.strip()}\".", pib_raw, report_date))
            continue

        destination_text = cells[destination_col].text.strip() if destination_col is not None else ""
        row_status = "ППД" if _is_ppd_destination(destination_text) else status

        date_text = cells[date_col].text.strip() if date_col is not None else ""
        events.append({
            "pib_raw": pib_raw, "status": row_status, "effective_date": _parse_date(date_text) or report_date,
            "heading": heading_text.strip(), "report_date": report_date,
        })

    return events, unresolved


def extract_status_events(file_path):
    """Читає ОДИН щоденний рапорт (.docx) і повертає (events, unresolved).

    events - список {"pib_raw", "status", "effective_date", "heading",
    "report_date"} - по одному на кожен рядок таблиці ПРИБУЛИ/ВИБУЛИ пункту
    рапорту, чий підзаголовок розпізнано (STATUS_HEADING_RULES).
    effective_date - з колонки "Дата прибуття"/"Дата вибуття" ЦІЄЇ таблиці
    (може відрізнятись від дати самого рапорту: зміна станом "після 17:00"
    потрапляє в НАСТУПНИЙ рапорт, але зі СПРАВЖНЬОЮ датою в комірці) - із
    запасним варіантом (дата рапорту), якщо комірку не вдалось розпізнати.

    unresolved - усе, що НЕ вдалось однозначно перетворити на подію: рядки
    таблиці з нерозпізнаним підзаголовком чи без визначеної колонки ПІБ, і
    кожен абзац розділу "Поза межами складу сил та засобів ...:" (вільний
    текст без таблиці - див. FREEFORM_GROUP_KEYWORD в constants.py), КРІМ
    абзаців зі стійким зворотом із FREEFORM_STATUS_PATTERNS (_freeform_event,
    одна подія), складнішим "виписаний ... ВЛК ... потребує відпустки ..."
    (_discharge_vlk_leave_events, ДВІ послідовні події з одного речення),
    "виписаний ... самовільно залишив військову частину (СЗЧ)"
    (_discharge_then_awol_event, ОДНА подія "СЗЧ" на дату виписки),
    "підтверджувальним" "З ДАТА відпустка за станом здоров'я ..."
    (_confirmed_health_leave_event, ВПСЗ на явну дату з речення) чи
    "завершив ВЛК"/"виписаний ...", з ДАТА направляється у відпустку ..."
    (_dated_treatment_leave_event, теж ВПСЗ на явну дату з речення) - вони
    теж стають подіями - і КРІМ самого рядка підпису наприкінці рапорту
    (_is_signatory_paragraph), усього ПІСЛЯ нього (взагалі не розглядається),
    та "голого" міні-підзаголовку без ПІБ (_is_bare_status_heading - напр.
    "СЗЧ:" перед реченням про конкретну людину - сам підзаголовок описує НЕ
    подію, а групу). Решта вільного тексту описує зміни без стабільної
    структури для надійного автоматичного розбору - навмисно НЕ парситься на
    статус/дату, а лишається для ручної перевірки людиною.

    ОКРЕМО - заголовок-СПИСОК "ПРИБУЛИ до пункту постійно(ї) дислокації
    військової частини ...:" (_is_arrival_list_heading,
    ARRIVAL_LIST_HEADING_KEYWORDS у constants.py) - КОЖЕН наступний абзац
    "звання ПІБ" (той самий стиль "List Paragraph", що й сам заголовок, БЕЗ
    таблиці) стає подією DEFAULT_STATUS ("РВЗ") на дату рапорту, АВТОМАТИЧНО
    (без ручної перевірки) - на відміну від розділу "Поза межами..." вище,
    цей формат - простий список імен з однозначним сенсом. Список
    завершується, щойно трапляється абзац ТОГО САМОГО стилю, у якому НЕ
    знайдено ПІБ (_NAME_RE) - тоді він обробляється як звичайний заголовок
    (напр. наступний розділ)."""
    doc = Document(file_path)
    report_date = report_date_from_filename(file_path)

    events, unresolved = [], []
    current_heading = ""
    in_freeform_section = False
    in_arrival_section = False
    arrival_heading_text = ""
    past_signature = False

    for block in _iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text or past_signature:
                continue
            if _is_signatory_paragraph(block, text):
                past_signature = True
                continue

            is_group_heading = block.style is not None and block.style.name == "List Paragraph"

            if is_group_heading and in_arrival_section and not _is_arrival_list_heading(text):
                name_match = _NAME_RE.search(text)
                if name_match is not None:
                    events.append({
                        "pib_raw": name_match.group(), "status": DEFAULT_STATUS, "effective_date": report_date,
                        "heading": arrival_heading_text, "report_date": report_date,
                    })
                    continue
                in_arrival_section = False

            if is_group_heading:
                in_freeform_section = FREEFORM_GROUP_KEYWORD in text.lower()
                in_arrival_section = _is_arrival_list_heading(text)
                if in_arrival_section:
                    arrival_heading_text = text
            elif in_freeform_section:
                name_match = _NAME_RE.search(text)
                event = _freeform_event(text, name_match, report_date)
                if event is not None:
                    events.append(event)
                else:
                    extra_events = _discharge_vlk_leave_events(text, name_match, report_date)
                    if not extra_events:
                        extra_events = _health_leave_then_vlk_events(text, name_match, report_date)
                    if extra_events:
                        events.extend(extra_events)
                    else:
                        awol_event = _discharge_then_awol_event(text, name_match, report_date)
                        if awol_event is not None:
                            events.append(awol_event)
                        else:
                            confirmed_event = (
                                _confirmed_health_leave_event(text, name_match, report_date)
                                or _dated_treatment_leave_event(text, name_match, report_date)
                            )
                            if confirmed_event is not None:
                                vlk_anchor_event = _vlk_in_progress_anchor_event(
                                    text, name_match, report_date, confirmed_event["effective_date"],
                                )
                                if vlk_anchor_event is not None:
                                    events.append(vlk_anchor_event)
                                events.append(confirmed_event)
                            elif not (name_match is None and _is_bare_status_heading(text)):
                                unresolved.append(_unresolved(
                                    "Вільний текст розділу \"Поза межами...\" - перевірити вручну.",
                                    name_match.group() if name_match else None, report_date, raw_text=text,
                                ))

            current_heading = text
            continue

        if past_signature:
            continue
        table_events, table_unresolved = _table_events(block, current_heading, report_date)
        events.extend(table_events)
        unresolved.extend(table_unresolved)

    return events, unresolved
