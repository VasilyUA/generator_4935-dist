import glob
import os
import re
from datetime import date, datetime, timedelta
from typing import TypeGuard

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
    BRIGADE_MILITARY_UNIT_CODES,
    TRANSFER_HEADING_MARKER,
    TRANSFER_TO_BRIGADE_STATUS,
    TRANSFER_TO_OTHER_UNIT_STATUS,
)

# ПІБ у вільному тексті розділу "Поза межами...": ПРІЗВИЩЕ великими, Ім'я й
# По батькові з великої - той самий формат, що й у таблицях рапорту. Це не
# стовідсоткове розпізнавання (вільний текст, а не дані), а лише допомога
# тому, хто перевірятиме REVIEW_FILE_NAME вручну.
_NAME_RE = re.compile(
    r"[А-ЯІЇЄҐ]{2,}(?:[-’'][А-ЯІЇЄҐ][а-яіїєґ’']+)?\s+"
    r"[А-ЯІЇЄҐ][а-яіїєґ’']+\s+[А-ЯІЇЄҐ][а-яіїєґ’']+"
)

# Складніший зворот вільного тексту: "виписаний ... ВЛК ... потребує відпустки
# [за станом здоров'я|для лікування] ...". Одне речення описує дві послідовні
# події (день ВЛК чи виписки, а наступного дня - відпустка), тож у просту
# схему "один зворот - один статус" (FREEFORM_STATUS_PATTERNS) воно не
# вкладається (див. _discharge_vlk_leave_events).
#
# _NEEDS_LEAVE_RE - обов'язкова умова: без неї нічого не застосовується. Вона
# підтверджує, що речення справді про цей сценарій, а не просто згадує ВЛК чи
# виписку з іншого приводу. "\s+", бо в рапортах трапляється подвійний пробіл
# ("потребує  відпустки"); "відпустк\w*", бо пишуть і "потребує відпустки", і
# "потребує відпустку".
_NEEDS_LEAVE_RE = re.compile(r"потребує\s+відпустк\w*", re.IGNORECASE)

# Дата біля згадки ВЛК - лише в цих формах: "довідки ВЛК №... від ДАТА", "ДАТА
# за висновком ВЛК", "закінчив проходження військово-лікарської комісії ДАТА"
# (людина цього дня завершила саму комісію, а не отримала довідку) і
# "рішення/висновок військово-лікарської комісії від ДАТА" (повна назва
# замість абревіатури - той самий сенс). Не будь-яка дата в реченні: інакше
# сплутались би різні дати, напр. дата виписки поруч із датою довідки ВЛК.
# "№\s*\S+", бо буває пробіл після "№" ("згідно довідки ВЛК № 2026-... від
# 18.09.2026"); без цього "від ДАТА" не знаходилось, і день довідки хибно
# ставав "ШП" (за датою виписки) замість "ВЛК".
_VLK_DATE_RE = re.compile(
    r"(?:(?P<date_before>\d{2}\.\d{2}\.\d{4})\s*(?:року)?\s*за\s+висновком\s+ВЛК"
    r"|ВЛК\s*(?:№\s*\S+)?\s*від\s*(?P<date_after>\d{2}\.\d{2}\.\d{4})"
    r"|закінчив\w*\s+проходженн\w*\s+військово-лікарс\w*\s+комісі\w*\s+(?P<date_after2>\d{2}\.\d{2}\.\d{4})"
    r"|військово-лікарс\w*\s+комісі\w*\s+від\s*(?P<date_after3>\d{2}\.\d{2}\.\d{4}))",
    re.IGNORECASE,
)
# Дата біля "виписаний": "ДАТА [року] виписаний..." чи "виписаний ... ДАТА"
# (до 40 символів між ними, напр. "з лікувального закладу").
_DISCHARGE_DATE_RE = re.compile(
    r"(?:(?P<date_before>\d{2}\.\d{2}\.\d{4})\s*(?:року)?\s*виписан"
    r"|виписан.{0,40}?(?P<date_after>\d{2}\.\d{2}\.\d{4}))",
    re.IGNORECASE,
)

# Реальний випадок: "ДАТА виписаний із [заклад]. Після виписки ... не прибув
# ... самовільно залишив військову частину (СЗЧ)." Це два речення: людину
# виписали з лікарні, до частини вона не прибула й вважається такою, що
# самовільно залишила частину. Довжина проміжку тут не обмежена (на відміну
# від інших "виписаний ... щось"), бо дата шукається лише біля "виписан"
# (той самий _DISCHARGE_DATE_RE, що й у _discharge_vlk_leave_events), тож
# іншу дату далі в тексті не зачепить. "вибув у СЗЧ" розпізнається окремо й
# раніше (_freeform_event за FREEFORM_STATUS_PATTERNS), тож тут не дублюється.
_AWOL_AFTER_DISCHARGE_RE = re.compile(r"самовільно\s+залишив\w*\s+військову\s+частину", re.IGNORECASE)

# Реальні випадки (три різні рапорти): "ДАТА [звання ПІБ] виписаний ... та
# проходить ВЛК" чи "виписаний ... та приступив до проходження ВЛК" - людина
# щойно почала комісію ("проходить"/"приступив", а не завершений факт з
# документом, як у _VLK_DATE_RE: "закінчив", "згідно висновку", "довідки ...
# від ДАТА"). День виписки тоді має бути "ВЛК" (людина вже на комісії, а не
# просто ще лікується - ШП), хоч би як знайшлась дата цього дня: через
# _DISCHARGE_DATE_RE (біля "виписан") чи як дата перед самим ПІБ (коли дата
# на початку абзацу, а "виписан" далеко за ПІБ - той самий прийом, що й у
# _discharge_then_awol_event). Без цього ставилось хибне "ШП" або сценарій
# узагалі лишався нерозпізнаним.
_VLK_IN_PROGRESS_RE = re.compile(r"проходить\s+ВЛК|приступив\w*\s+до\s+проходженн\w*\s+ВЛК", re.IGNORECASE)

# Якщо те саме речення ще й прямо згадує "відпустки/відпустку за станом
# здоров'я" (не обов'язково зворотом "вибув у ...", як у
# FREEFORM_STATUS_PATTERNS: тут фраза часто описує стан до чи під час події,
# напр. "під час відпустки за станом здоров'я ... потребує відпустки для
# лікування"), то парна подія "відпустка" наступного дня
# (_discharge_vlk_leave_events) - це "ВПСЗ", а не просто "ВП". Так показав
# реальний випадок: окреме речення наступного рапорту підтвердило саме
# "ВПСЗ". На якорну подію (сам день ВЛК чи виписки) це не впливає - вона
# лишається "ВЛК"/"ШП".
_HEALTH_LEAVE_MENTION_RE = re.compile(r"відпуст\w*\s+за\s+станом\s+здоров['’]?я", re.IGNORECASE)

# Напрямок важливий (як і в PPD_DESTINATION_KEYWORDS): колонка "Куди прибув"
# буває "Прибув ... з відпустки за станом здоров'я" - людина повертається, це
# звичайний РВЗ, а не "ВПСЗ". "ВПСЗ" - лише коли "Вибув ... у відпустку за
# станом здоров'я". Тому тут "у відпустк\w*" (напрямок "туди"), а не просто
# "відпуст\w*", як у _HEALTH_LEAVE_MENTION_RE: та - для вільного тексту, де
# напрямок уже видно з інших слів звороту (напр. "вибув у відпустку").
_HEALTH_LEAVE_DEPARTURE_RE = re.compile(r"у\s+відпустк\w*\s+за\s+станом\s+здоров['’]?я", re.IGNORECASE)

# Те саме для відпустки за сімейними обставинами - "ВПС". Реальний випадок:
# рядок під "У відпустку:" - "Вибув з району виконання завдань у відпустку за
# сімейними обставинами", а в Табелі було загальне "ВП", і звірка з addons
# (book61, "Тип відпустки") показувала розбіжність. Напрямок той самий: лише
# "у/в відпустку", а "з відпустки за сімейними обставинами" - повернення.
_FAMILY_LEAVE_DEPARTURE_RE = re.compile(r"\b[ув]\s+відпустк\w*\s+за\s+сімейн\w*\s+обставин", re.IGNORECASE)

# Ще одне формулювання причини відпустки: "потребує відпустки для лікування
# після поранення на 30 календарних днів" - той самий медичний сенс, що й
# "за станом здоров'я", тож парна подія - теж "ВПСЗ", а не загальне "ВП". Не
# "ВПБП" ("Відпустка для лікування після тяжкого поранення"): "ВПБП" і
# "100_ВПБП" беруться лише з resources/information_unit
# (PAYMENT_STATUS_OVERRIDES), а не з тексту рапорту (рішення користувача).
_WOUND_TREATMENT_LEAVE_MENTION_RE = re.compile(r"лікуванн\w*\s+після\s+(?:тяжкого\s+)?поранення", re.IGNORECASE)

# Реальний випадок: наступного дня після "потребує відпустки за висновком ВЛК
# ..." без причини (у _discharge_vlk_leave_events тоді виходить "ВП", бо
# речення не згадує "за станом здоров'я") окремий рапорт підтверджує причину
# окремим реченням: "З ДАТА відпустка за станом здоров'я ..." або "відпустка
# за станом здоров'я з ДАТА" (трапляються обидва порядки слів) - вже без
# "потребує" і без "вибув у відпустку". Кожна гілка шаблону вимагає власну
# дату, щоб голе "потребує відпустки за станом здоров'я" без дати тут не
# підхоплювалось і лишалось на ручну перевірку
# (test_freeform_needs_leave_without_discharge_or_vlk_date_is_not_auto_applied).
# Кожен рапорт розбирається окремо, тож подія "ВПСЗ" з цього речення окремо
# потрапляє в apply_events і там без ручної перевірки перемагає вже наявну
# "ВП" на ту саму дату (STATUS_CONFLICT_OVERRIDES).
_CONFIRMED_HEALTH_LEAVE_RE = re.compile(
    r"з\s+(?P<date1>\d{2}\.\d{2}\.\d{4})\s+відпуст\w*\s+за\s+станом\s+здоров['’]?я"
    r"|відпуст\w*\s+за\s+станом\s+здоров['’]?я\s+з\s+(?P<date2>\d{2}\.\d{2}\.\d{4})",
    re.IGNORECASE,
)

# Реальний випадок: "з ДАТА1 [по ДАТА_кінця] у відпустці за станом здоров'я, з
# ДАТА2 [року] направлений на проходження військово-лікарської комісії". Одне
# речення - дві послідовні події, і обидві дати початку вказано явно: ВПСЗ з
# ДАТА1, ВЛК з ДАТА2 (комісія визначає подальший статус). Між датою й фразою
# - короткий нежадібний проміжок "[\s\S]{0,25}?" ("\s\S", бо всередині абзацу
# буває перенос рядка, напр. між "з ДАТА1" і "по ДАТА_кінця"). Довжину
# обмежено навмисно: інакше ДАТА1 з початку речення "дотяглася" б до фрази
# ВЛК, де потрібна ближча ДАТА2. 25 символів вистачає на "по ДАТА_кінця" чи
# "року направлений", але не на цілу проміжну фразу.
_HEALTH_LEAVE_DATED_RE = re.compile(
    r"(?P<date>\d{2}\.\d{2}\.\d{4})[\s\S]{0,25}?відпуст\w*\s+за\s+станом\s+здоров['’]?я",
    re.IGNORECASE,
)
_VLK_REFERRAL_DATED_RE = re.compile(
    r"(?P<date>\d{2}\.\d{2}\.\d{4})[\s\S]{0,25}?на\s+проходженн\w*\s+військово-лікарс\w*\s+комісі\w*",
    re.IGNORECASE,
)

# Реальний випадок: "ДАТА1 звання ПІБ завершив/закінчив проходження
# військово-лікарської комісії. З ДАТА2 направляється у відпустку ..." - той
# самий факт, що й "закінчив ВЛК ДАТА, потребує відпустки за станом здоров'я
# ..." (_discharge_vlk_leave_events, форма date_after2), лише переказаний у
# наступному рапорті іншими словами: без "потребує" (_NEEDS_LEAVE_RE не
# спрацює) і без "за станом здоров'я" (тут "відпустку для лікування"), тож
# жоден інший механізм цього не бачить.
_VLK_COMPLETED_MENTION_RE = re.compile(
    r"(?:завершив|закінчив)\w*\s+проходженн\w*\s+військово-лікарс\w*\s+комісі\w*", re.IGNORECASE
)
# Реальний випадок: "ДАТА1 звання ПІБ виписаний із [заклад], з ДАТА2
# направляється у відпустку для лікування ..." - той самий зворот "з ДАТА
# направляється у відпустку", але медичний контекст тут - виписка з лікарні,
# а не завершена ВЛК. Потрібна лише присутність "виписан", без власної дати:
# дата відпустки береться з _DATED_TREATMENT_LEAVE_RE, а не з
# _DISCHARGE_DATE_RE (той чекає дату впритул до "виписан" чи перед ПІБ, а
# тут дати виписки немає або її заплутують інші дати й установи в реченні).
_DISCHARGE_MENTION_RE = re.compile(r"виписан\w*", re.IGNORECASE)
# Медичний контекст (_VLK_COMPLETED_MENTION_RE або _DISCHARGE_MENTION_RE) і ця
# фраза мають бути разом (_dated_treatment_leave_event): саме по собі "з ДАТА
# направляється у відпустку" занадто неоднозначне - це може бути звичайна,
# немедична відпустка. Друга альтернатива - "ДАТА вибув у відпустку"
# (реальний випадок: "... закінчив проходження військово-лікарської комісії,
# ДАТА вибув у відпустку ...") - інше дієслово, дата перед ним і без "з", тож
# окрема група "date2", щоб не плутати з "date".
_DATED_TREATMENT_LEAVE_RE = re.compile(
    r"з\s+(?P<date>\d{2}\.\d{2}\.\d{4})\s+направля\w*\s+у\s+відпустку"
    r"|(?P<date2>\d{2}\.\d{2}\.\d{4})\s+вибув\w*\s+у\s+відпустку",
    re.IGNORECASE,
)


def _iter_block_items(doc):
    """Абзаци й таблиці документа в порядку появи. Стандартний рецепт
    python-docx: прохід по doc.element.body, бо doc.paragraphs і doc.tables
    окремо взаємного порядку не зберігають."""
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
    """Дата найпізнішого .docx рапорту в report_dir за датою в назві
    (report_date_from_filename) або None, якщо таких рапортів немає.
    Lock-файли Word ("~$...") і файли без дати в назві пропускаються.

    Генератори, яким потрібен один цільовий місяць (Відомість РОП/ВОП), беруть
    цю дату, а не date.today(): на початку місяця рапортів і файлів за нього
    ще немає (звичайна затримка), і "поточний місяць" дав би порожній
    результат, хоча дані за попередній місяць готові."""
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
    """Чи містить heading_text одне з IGNORED_TABLE_HEADING_KEYWORDS - тоді весь
    пункт рапорту ігнорується повністю (_table_events: ні подій, ні записів на
    ручну перевірку), а не отримує статус і не йде в unresolved як "невідомий
    заголовок"."""
    lowered = (heading_text or "").lower()
    return any(keyword in lowered for keyword in IGNORED_TABLE_HEADING_KEYWORDS)


def _is_arrival_list_heading(text):
    """Заголовок-список "ПРИБУЛИ до пункту постійно(ї) дислокації військової
    частини ...:" (ARRIVAL_LIST_HEADING_KEYWORDS): за ним без таблиці йдуть
    абзаци "звання ПІБ", і кожен автоматично отримує DEFAULT_STATUS
    (extract_status_events)."""
    lowered = text.lower()
    return all(keyword in lowered for keyword in ARRIVAL_LIST_HEADING_KEYWORDS)


def _is_bare_status_heading(text):
    """Чи це підзаголовок пункту рапорту всередині розділу "Поза межами..."
    (STATUS_HEADING_RULES). Такий підзаголовок описує не подію конкретної
    людини, а групу (чи стоїть перед таблицею), тож окремо на ручну перевірку
    не йде - як і підзаголовки таблиць.

    Обов'язкова ознака - двокрапка в кінці: так закінчуються всі підзаголовки
    рапорту ("З відпустки:", "У відрядження:", "СЗЧ:", "Безповоротні втрати:",
    "Повідомлення про факт смерті ...:"), а речення про конкретну людину
    завжди закінчуються крапкою. Без неї пошук підрядка в _status_for_heading
    міг би зачепити звичайне речення, що мимохідь згадує ключове слово.

    Пошук - за підрядком, а не точним збігом: реальний підзаголовок
    "Повідомлення про факт смерті військовослужбовця ЗСУ, який знаходився на
    лікуванні у закладах Міністерства охорони здоров'я ...:"
    (STATUS_HEADING_RULES["факт смерті"]) значно довший за "СЗЧ:", і точна
    перевірка його не впізнавала - він хибно йшов на ручну перевірку, хоча
    таблицю під ним уже правильно розбирав _table_events."""
    stripped = text.strip()
    if not stripped.endswith(":"):
        return False
    return _status_for_heading(stripped) is not None


def _unresolved(reason, pib_raw, report_date, **extra):
    return {"reason": reason, "pib_raw": pib_raw, "report_date": report_date, **extra}


def _freeform_event(text, name_match, report_date):
    """Подія зі сталого звороту вільного тексту "Поза межами..."
    (FREEFORM_STATUS_PATTERNS) або None, якщо ПІБ не розпізнано чи жоден
    шаблон не підійшов (тоді - ручна перевірка, див. extract_status_events)."""
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
    """Подія з підтверджувального речення наступного рапорту: "З ДАТА відпустка
    за станом здоров'я ..." чи "відпустка за станом здоров'я з ДАТА"
    (_CONFIRMED_HEALTH_LEAVE_RE) - "ВПСЗ" на дату з самого речення (тут вона
    завжди явна, а не report_date). None, якщо ПІБ не розпізнано чи зворот не
    підійшов."""
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
    """Подія з підтверджувального речення: "ДАТА1 [звання ПІБ] завершив/закінчив
    проходження військово-лікарської комісії" або "ДАТА1 [звання ПІБ]
    виписаний із [заклад]", а в тому самому реченні - "з ДАТА2 направляється у
    відпустку ..." чи "ДАТА2 вибув у відпустку ..." (_VLK_COMPLETED_MENTION_RE
    або _DISCHARGE_MENTION_RE разом із _DATED_TREATMENT_LEAVE_RE). Результат -
    "ВПСЗ" на ДАТУ2 з речення (рішення користувача): той самий статус, що й у
    формі "завершив ВЛК ДАТА, потребує відпустки за станом здоров'я ..."
    (_discharge_vlk_leave_events) - людина вийшла з медичного процесу й іде у
    відпустку. Медичний контекст і "ДАТА2 ... у відпустку" мають бути разом:
    саме по собі друге може бути звичайною відпусткою. None, якщо ПІБ не
    розпізнано чи зворот не підійшов."""
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
    """Реальні випадки: "ДАТА1 [звання ПІБ] виписаний ... та проходить
    ВЛК/приступив до проходження ВЛК, з ДАТА2 відпустка/направляється у
    відпустку ...". _dated_treatment_leave_event і
    _confirmed_health_leave_event вважають, що людина вже вийшла з медичного
    процесу, і дають лише подію "відпустка" на ДАТУ2. Явне "проходить" чи
    "приступив до проходження" ВЛК (_VLK_IN_PROGRESS_RE) означає, що ДАТА1
    (день виписки й початку комісії) має бути окремою подією "ВЛК" - інакше
    цей день мовчки лишається попереднім статусом. Завершена ВЛК ("закінчив",
    "завершив" - _VLK_COMPLETED_MENTION_RE) сюди навмисно не входить: для неї
    окремий день "ВЛК" не потрібен, лише "ВПСЗ" на дату відпустки (як
    підтвердив користувач, test_completed_vlk_then_dated_leave_sentence_produces_vpsz_event).

    ДАТА1 шукається так само, як у _discharge_vlk_leave_events: спершу
    _DISCHARGE_DATE_RE (дата біля "виписан"), а запасний варіант - дата перед
    самим ПІБ (коли дата на початку абзацу, а "виписан" далеко за ПІБ).

    Повертає подію "ВЛК" на ДАТУ1 або None: немає "проходить ВЛК", немає
    придатної дати чи ДАТА1 збігається з leave_date (друга подія на ту саму
    дату була б суперечністю)."""
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
    """0 чи 2 події зі складеного речення "з ДАТА1 у відпустці за станом
    здоров'я та ДАТА2 на проходження військово-лікарської комісії": "ВПСЗ" з
    ДАТА1 і "ВЛК" з ДАТА2 (комісія визначає подальше). Обидві фрази з власними
    датами (_HEALTH_LEAVE_DATED_RE і _VLK_REFERRAL_DATED_RE) мають бути разом:
    кожна окремо занадто неоднозначна (звичайна відпустка чи ВЛК з іншої
    причини), тож без обох - ручна перевірка."""
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
    """0, 1 чи 2 події зі складного звороту "виписаний ... ВЛК ... потребує
    відпустки ..." (реальні випадки з resources/report/*.docx, розділ "Поза
    межами").

    Без _NEEDS_LEAVE_RE ("потребує відпустки") подій немає: лише за згадкою
    ВЛК чи виписки застосовувати щось автоматично занадто ризиковано. З нею:

    - Є дата біля ВЛК як документа чи рішення (_VLK_DATE_RE: "довідки ВЛК
      №... від ДАТА", "ДАТА за висновком ВЛК", "рішення/висновок
      військово-лікарської комісії від ДАТА" - форми date_before, date_after,
      date_after3) - дві події: цього дня "ВЛК" (комісія того дня ще триває
      адміністративно), а відпустка - з наступного дня (підтверджено на
      реальних випадках 04.08.2026, 07.08.2026, 31.08.2026).
    - Дата - завершення самого проходження ("закінчив/завершив проходження
      ... комісії ДАТА", date_after2) - одна подія: відпустка з того самого
      дня. Комісію вже пройдено, тож окремий день "ВЛК" був би помилкою (на
      відміну від форм вище, де ДАТА - лише дата документа).
    - Інакше, є дата біля "виписаний" (_DISCHARGE_DATE_RE) - цього дня "ШП"
      (людину лише виписали з одного закладу, про ВЛК ще нічого конкретного)
      або "ВЛК", якщо речення ще й каже, що людина того дня почала комісію
      ("проходить ВЛК"/"приступив до проходження ВЛК", _VLK_IN_PROGRESS_RE).
      Відпустка в обох випадках - з наступного дня.
    - Інакше, якщо речення каже "проходить ВЛК"/"приступив до проходження
      ВЛК" і є дата перед самим ПІБ (як у _discharge_then_awol_event: дата на
      початку абзацу, а "виписан" далеко за ПІБ) - цього дня "ВЛК", відпустка
      - з наступного.
    - Жодної з цих дат немає - подій немає (ручна перевірка).

    Відпустка - "ВПСЗ", якщо речення ще й згадує "відпустки за станом
    здоров'я" (_HEALTH_LEAVE_MENTION_RE) чи "лікування після [тяжкого]
    поранення" (_WOUND_TREATMENT_LEAVE_MENTION_RE) - обидва про лікування;
    інакше - "ВП". Ніколи не "ВПБП": його беруть лише з
    resources/information_unit (рішення користувача).

    Подія "відпустка" має "is_forecast": True. Рапорт не каже, коли відпустка
    справді почалась - дату лише обчислено як "наступний день" після ВЛК. Якщо
    пізніший рапорт на ту саму дату прямо повідомляє інше (напр. людина
    повернулась у ППД, а відпустка почалась на кілька днів пізніше),
    apply_events (content/oblik_timesheet.py) віддає перевагу цьому факту, а
    не вважає це суперечністю для ручної перевірки."""
    if name_match is None or not _NEEDS_LEAVE_RE.search(text):
        return []

    pib_raw = name_match.group()
    if _HEALTH_LEAVE_MENTION_RE.search(text) or _WOUND_TREATMENT_LEAVE_MENTION_RE.search(text):
        leave_status = "ВПСЗ"
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
            "heading": FREEFORM_GROUP_KEYWORD, "report_date": report_date, "is_forecast": True,
        },
    ]


def _discharge_then_awol_event(text, name_match, report_date):
    """Подія чи None зі звороту "ДАТА звання ПІБ виписаний із [заклад]. Після
    виписки ... не прибув ... самовільно залишив військову частину (СЗЧ)." Це
    два речення (людину виписали, а до частини вона не прибула), але одна
    подія: "СЗЧ" на дату виписки.

    Дата тут стоїть перед ПІБ, на початку абзацу, а не впритул до "виписан",
    тож _DISCHARGE_DATE_RE не підходить - довелося б підбирати довжину
    проміжку "на око" під довжину ПІБ. Натомість дата шукається в тексті
    перед ПІБ (до name_match.start()): реальні абзаци завжди починаються
    "ДАТА звання ПІБ ...". Якщо дат перед ПІБ кілька - береться остання
    (найближча до ПІБ); якщо жодної - запасний варіант _DISCHARGE_DATE_RE
    (охоплює й порядок "виписаний ... ДАТА").

    Тут немає "потребує відпустки", тож _discharge_vlk_leave_events цього не
    бачить, а сама згадка "виписаний" без фрази про самовільне залишення
    (_AWOL_AFTER_DISCHARGE_RE) - занадто мало, щоб ставити "СЗЧ". None, якщо
    ПІБ не розпізнано, цієї фрази немає чи немає дати виписки (ручна
    перевірка)."""
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


# Назви місяців у родовому відмінку - лише для _EXTENDED_BUSINESS_TRIP_RE, де
# дату записано словами ("ДЕНЬ МІСЯЦЬ РІК року"). Решта дат рапорту завжди
# числові (дд.мм.рррр, DATE_RE).
_UKRAINIAN_MONTHS_GENITIVE = {
    "січня": 1, "лютого": 2, "березня": 3, "квітня": 4, "травня": 5, "червня": 6,
    "липня": 7, "серпня": 8, "вересня": 9, "жовтня": 10, "листопада": 11, "грудня": 12,
}

# Розділ "Поза межами..." може мати речення-заголовок без жодного ПІБ:
# "Продовжено відрядження з ДЕНЬ МІСЯЦЬ РІК року терміном на N днів до
# [установа] з метою [причина]:", а під ним - простий список "звання ПІБ"
# (стиль "Normal", без таблиці й без "List Paragraph" - на відміну від
# _is_arrival_list_heading, де і заголовок, і список - "List Paragraph").
# Дату записано словами, звідси _UKRAINIAN_MONTHS_GENITIVE. "з", а не "від",
# як у записі FREEFORM_STATUS_PATTERNS "продовжен... відрядженн... від ДАТА":
# той розрахований на одне речення з ПІБ і статусом (_freeform_event вимагає
# name_match), а цей - на заголовок без ПІБ перед списком
# (pending_group_status у extract_status_events).
_EXTENDED_BUSINESS_TRIP_RE = re.compile(
    r"продовжен\w*\s+відрядженн\w*\s+з\s+(?P<day>\d{1,2})\s+(?P<month>"
    + "|".join(_UKRAINIAN_MONTHS_GENITIVE) + r")\s+(?P<year>\d{4})",
    re.IGNORECASE,
)


def _extended_business_trip_date(text):
    """Дата зі звороту "Продовжено відрядження з ДЕНЬ МІСЯЦЬ РІК року ..."
    (_EXTENDED_BUSINESS_TRIP_RE) або None, якщо зворот не підійшов чи дата
    неможлива (напр. 31 лютого)."""
    match = _EXTENDED_BUSINESS_TRIP_RE.search(text)
    if match is None:
        return None
    month = _UKRAINIAN_MONTHS_GENITIVE[match.group("month").lower()]
    try:
        return date(int(match.group("year")), month, int(match.group("day")))
    except ValueError:
        return None


def _is_bare_name_paragraph(name_match, text) -> TypeGuard[re.Match[str]]:
    """Чи рядок складається лише з ПІБ (можливо, зі званням перед ним, напр.
    "старший лейтенант ПЕРШИЙ Перший Перший") - тобто це учасник списку під
    _EXTENDED_BUSINESS_TRIP_RE, а не речення з додатковим змістом про людину
    (таке має пройти звичайним ланцюжком _freeform_event і далі)."""
    return name_match is not None and name_match.end() == len(text)


# Скільки перших слів абзацу перевіряти на SIGNATORY_ROLE_KEYWORDS. Рядок
# підпису завжди починається з ролі чи звання, тож довга фраза ("Прошу ...
# доручити заступнику командира батальйону ...") не має вважатися підписом
# лише тому, що згадує те саме слово десь посередині.
_SIGNATORY_ROLE_LEAD_WORDS = 5


def _is_signatory_paragraph(block, text):
    """Рядок підпису наприкінці рапорту ("Тимчасово виконуючий обов'язки
    командира ...", "Командир ..." і звання з ім'ям). Він не стосується
    жодної людини зі складу, тож не має потрапляти в unresolved як "текст без
    розпізнаного ПІБ". Перевіряється кожен абзац рапорту, а не лише розділ
    "Поза межами...", інакше в рапорті без цього розділу підпис не знайшовся
    б.

    Підпис - якщо абзац має стиль "No Spacing" (у реальних рапортах це завжди
    саме рядок підпису, перевірено на 8 файлах) або текст згадує і роль
    підписанта (SIGNATORY_ROLE_KEYWORDS - сталі слова, не прив'язані до
    підрозділу), і назву власного підрозділу (SIGNATORY_UNIT_NAMES,
    resources/data.json). Потрібні обидві ознаки: сама назва підрозділу буває
    й у звичайному реченні про людину ("прибув з ... [назва підрозділу]").

    Виняток - "Командиру ..." (адресат рапорту, завжди перший рядок,
    давальний відмінок): "командир" - підрядок і "командира" (підпис), і
    "командиру" (адресат), тож без винятку перший рядок рапорту зупиняв би
    розбір усього іншого."""
    if block.style is not None and block.style.name == "No Spacing":
        return True
    lowered = text.lower()
    if lowered.startswith("командиру"):
        return False
    # Реальний випадок (08.10.2026): звичайне речення посередині рапорту
    # ("Прошу за даним фактом призначити службове розслідування, проведення
    # якого доручити заступнику командира батальйону з психологічної
    # підтримки персоналу [повна назва підрозділу]") згадує і роль, і
    # підрозділ. Рядок підпису ж завжди починається з ролі чи звання. Без
    # цього обмеження past_signature спрацьовував на 10 абзаців раніше і мовчки
    # "проковтував" решту рапорту (тут - цілу таблицю "У відпустку:" з трьома
    # людьми), без жодного запису на ручну перевірку.
    lead_words = " ".join(lowered.split()[:_SIGNATORY_ROLE_LEAD_WORDS])
    has_role = any(keyword in lead_words for keyword in SIGNATORY_ROLE_KEYWORDS)
    has_unit = any(name.lower() in lowered for name in SIGNATORY_UNIT_NAMES)
    return has_role and has_unit


def _detect_pib_and_date_columns(header_row):
    """Позиції колонок ПІБ і дати (прибуття чи вибуття) у цій таблиці - за
    назвою заголовка (як і в generator_br_and_report_for_money), бо набір і
    порядок колонок у пунктах рапорту різні ("Дата прибуття" і "Дата
    вибуття", "Тип відпустки, куди вибув" і "Звідки вибув" тощо). date_col -
    None, якщо колонки з датою немає (тоді діє дата самого рапорту, див.
    _table_events).

    Таблиця під "Повідомлення про факт смерті...:" (STATUS_HEADING_RULES["факт
    смерті"]) має дві колонки з "дата" - "Дата госпіталізації" і "Дата
    смерті", - і перша за порядком тут не підходить. Тому спершу шукається
    колонка, де є і "дата", і "смерті"; якщо такої немає - перша колонка з
    "дата" в назві."""
    header_texts = [cell.text.strip().lower() for cell in header_row.cells]
    pib_col = next((i for i, text in enumerate(header_texts) if "прізвищ" in text), None)
    date_col = next((i for i, text in enumerate(header_texts) if "дата" in text and "смерті" in text), None)
    if date_col is None:
        date_col = next((i for i, text in enumerate(header_texts) if "дата" in text), None)
    return pib_col, date_col


def _destination_column(header_row, words=("куди",)):
    """Позиція колонки призначення ("Куди прибув", "Куди вибув", "Тип
    відпустки, куди прибув" тощо) - за першим зі слів words у заголовку, як і
    pib_col/date_col у _detect_pib_and_date_columns. None, якщо такої колонки
    немає."""
    header_texts = [cell.text.strip().lower() for cell in header_row.cells]
    for word in words:
        found = next((i for i, text in enumerate(header_texts) if word in text), None)
        if found is not None:
            return found
    return None


# Без колонки "Куди ..." текст тієї самої форми буває в колонці "Звідки
# вибув". Реальний випадок (рапорт 29.10): під "У відпустку:" - "Звідки
# вибув": "Вибув з району виконання завдань у відпустку за сімейними
# обставинами", і тип відпустки губився. Для переведень цієї заміни немає:
# там важливий підрозділ, куди переведено, а не звідки.
_DESCRIPTION_COLUMN_WORDS = ("куди", "звідки")


def _is_ppd_destination(text):
    """Чи згадує текст колонки "Куди ..." пункт постійної дислокації
    (PPD_DESTINATION_KEYWORDS) - пошук підрядка без урахування регістру, як у
    _status_for_heading та _is_ignored_table_heading."""
    lowered = text.lower()
    return any(keyword in lowered for keyword in PPD_DESTINATION_KEYWORDS)


def _is_health_leave_destination(text):
    """Чи каже текст колонки "Куди .../Тип відпустки, куди ..." прямо "у
    відпустку за станом здоров'я" (_HEALTH_LEAVE_DEPARTURE_RE). Реальний
    випадок: рядок під "У відпустку:" (за заголовком завжди загальне "ВП"), у
    колонці якого "...вибув ... у відпустку за станом здоров'я" - точніший
    статус "ВПСЗ" має переважати заголовок, як і в _is_ppd_destination.
    Напрямок важливий: "прибув ... з відпустки за станом здоров'я" (людина
    повертається, звичайний РВЗ) сюди не підпадає - шаблон вимагає саме "у
    відпустку"."""
    return bool(_HEALTH_LEAVE_DEPARTURE_RE.search(text))


def _is_family_leave_destination(text):
    """Чи каже текст колонки "Куди ..." прямо "у відпустку за сімейними
    обставинами" (_FAMILY_LEAVE_DEPARTURE_RE) - тоді "ВПС", а не загальне
    "ВП" заголовка "У відпустку:"."""
    return bool(_FAMILY_LEAVE_DEPARTURE_RE.search(text))


# Рядок під "У відрядження:" (STATUS_HEADING_RULES["у відрядженн"] - загальне
# "ВД"), у колонці "Тип відрядження, куди вибув" якого причина чи пункт
# призначення - навчання: абревіатура навчального центру ("Вибув з району
# виконання завдань до НЦ") або просто слово "навчання" ("Вибув з району
# виконання завдань на навчання до Національної академії сухопутних військ
# ..."). Такий рядок одразу отримує точніший статус "Навч"
# (STATUS_MEANINGS["Навч"] - "навчання по відрядженню"), а не загальне "ВД" -
# той самий принцип, що й у _is_ppd_destination/_is_health_leave_destination.
# "\bнц\b" - з межами слова, інакше спрацювало б і на "нц" всередині слів
# (напр. "конференції"); "навчанн\w*" межі не потребує - інших коренів з
# таким початком немає.
_TRAINING_DESTINATION_RE = re.compile(r"\bнц\b|навчанн\w*", re.IGNORECASE)


def _is_training_destination(text):
    """Чи називає колонка "Тип відрядження, куди вибув" причиною чи пунктом
    призначення навчання (_TRAINING_DESTINATION_RE). Викликається лише тоді,
    коли заголовок уже дав "ВД" (_table_events), а всі заголовки, що ведуть до
    "ВД" ("у відрядженн", "у {SPECIAL_TASK_GROUP}", "до {SPECIAL_TASK_GROUP}"),
    - лише про вибуття (повернення "з відрядженн" дає DEFAULT_STATUS). Тож, на
    відміну від _HEALTH_LEAVE_DEPARTURE_RE, напрямок тут окремо перевіряти не
    треба."""
    return bool(_TRAINING_DESTINATION_RE.search(text))


# Переведення (заголовок "Переведені:", TRANSFER_HEADING_MARKER): статус рядка
# визначає колонка "Куди вибув" (_transfer_status), напр. "Переведений до
# військової частини А0000".
_MILITARY_UNIT_CODE_RE = re.compile(r"[АA]\s*(\d{4})")
_OTHER_MILITARY_UNIT_RE = re.compile(r"військов\w*\s+частин\w*|\bв/ч\b", re.IGNORECASE)
_BRIGADE_SUBUNIT_RE = re.compile(r"бригад\w*|батальйон\w*|\bрот\w*|батаре\w*|взвод\w*|підрозділ\w*|дивізіон\w*", re.IGNORECASE)


def _transfer_status(destination_text):
    """Переведений у підрозділи бригади - "ПРВДБ", в іншу військову частину -
    "ПРВДВЧ" (рішення користувача). Номер частини в тексті дає ПРВДБ, лише
    якщо це власна частина чи частина з data.json ["BRIGADE_MILITARY_UNITS"]
    (BRIGADE_MILITARY_UNIT_CODES), інакше ПРВДВЧ. Без номера: "військова
    частина"/"в/ч" - ПРВДВЧ, підрозділ бригади (бригада, батальйон, рота,
    батарея, взвод, підрозділ) - ПРВДБ. Якщо не видно - None (ручна
    перевірка, без вгадування)."""
    codes = set(_MILITARY_UNIT_CODE_RE.findall(destination_text))
    if codes:
        return TRANSFER_TO_BRIGADE_STATUS if codes <= BRIGADE_MILITARY_UNIT_CODES else TRANSFER_TO_OTHER_UNIT_STATUS
    if _OTHER_MILITARY_UNIT_RE.search(destination_text):
        return TRANSFER_TO_OTHER_UNIT_STATUS
    if _BRIGADE_SUBUNIT_RE.search(destination_text):
        return TRANSFER_TO_BRIGADE_STATUS
    return None


# "20.10.2026 після 17.00" чи "після 17:00" у клітинці дати - подія сталась
# увечері. Якщо на цей день у resources/ОБЛІК.xlsx уже вручну вписано інший
# статус, подія діє з наступного дня (oblik_timesheet.apply_events).
_AFTER_HOURS_RE = re.compile(r"після\s+\d{1,2}[:.]\d{2}", re.IGNORECASE)


def _table_events(table, heading_text, report_date):
    """(events, unresolved) для однієї таблиці рапорту.

    Заголовок з IGNORED_TABLE_HEADING_KEYWORDS ігнорується повністю - одразу
    ([], []), без подій і без записів на ручну перевірку (ця перевірка - перша).

    Інакше кожна людина отримує статус за заголовком (_status_for_heading),
    крім випадків, коли колонка "Куди ..." цього рядка каже точніше:
    - згадує пункт постійної дислокації (PPD_DESTINATION_KEYWORDS) - "ППД",
      хоч би що казав заголовок (реальний випадок: рядок під "З відпустки:",
      зазвичай DEFAULT_STATUS, а в "Куди прибув" - "...до пункту постійної
      дислокації...");
    - згадує "у відпустку за станом здоров'я" (_is_health_leave_destination)
      - "ВПСЗ" (рядок під "У відпустку:", зазвичай "ВП");
    - згадує "у відпустку за сімейними обставинами"
      (_is_family_leave_destination) - "ВПС";
    - заголовок дав "ВД", а колонка називає причиною навчання
      (_is_training_destination) - "Навч" (рядок під "У відрядження:" з "...до
      НЦ" чи "...на навчання до [назва академії]...").

    Ці перевірки колонки "Куди ..." працюють, навіть якщо заголовок не
    відповідає жодному з STATUS_HEADING_RULES. Реальний випадок: підзаголовок
    "Для проходження військової служби:" (під груповим "ПРИБУЛИ до пункту
    постійної дислокації ...:") сам жодного правила не має, але "Куди
    прибув" каже "Прибув до пункту постійної дислокації" - цього досить для
    "ППД".

    unresolved - рядок таблиці, лише якщо (1) заголовок не відповідає жодному
    правилу і колонка "Куди ..." теж нічого не підказала (справді незнайомий
    підзаголовок - рядок не губиться, а йде на ручну перевірку) або (2)
    колонку ПІБ визначити не вдалось."""
    if _is_ignored_table_heading(heading_text):
        return [], []

    status = _status_for_heading(heading_text)
    pib_col, date_col = _detect_pib_and_date_columns(table.rows[0])
    destination_col = _destination_column(table.rows[0])
    description_col = _destination_column(table.rows[0], _DESCRIPTION_COLUMN_WORDS)
    events, unresolved = [], []

    if pib_col is None:
        unresolved.append(_unresolved(f"Не вдалося визначити колонку ПІБ у таблиці під заголовком \"{heading_text.strip()}\".", None, report_date))
        return events, unresolved

    for row in table.rows[1:]:
        cells = row.cells
        pib_raw = cells[pib_col].text.strip()
        if not pib_raw:
            continue

        # Перевіряється раніше за "status is None" нижче, а не лише для вже
        # розпізнаного заголовка: "Для проходження військової служби:" (під
        # розпізнаним груповим "ПРИБУЛИ до пункту постійної дислокації ...:")
        # жодному правилу не відповідає, але "Куди прибув" каже "Прибув до
        # пункту постійної дислокації" - цього досить для "ППД" напряму з рядка.
        # Колонка "Куди ..." - точніший і надійніший сигнал, ніж підзаголовок.
        # Умова status == "ВД" для "Навч" лишається: для нерозпізнаного
        # заголовка status - None, тож ця гілка там і не спрацює - ніхто не
        # підтвердив, що "ВД" там доречне.
        destination_text = cells[destination_col].text.strip() if destination_col is not None else ""
        description_text = cells[description_col].text.strip() if description_col is not None else ""
        if status == TRANSFER_HEADING_MARKER:
            row_status = _transfer_status(destination_text)
            if row_status is None:
                unresolved.append(_unresolved(
                    "Переведення: з колонки \"Куди вибув\" не видно, чи це підрозділ бригади "
                    f"(\"{TRANSFER_TO_BRIGADE_STATUS}\"), чи інша військова частина "
                    f"(\"{TRANSFER_TO_OTHER_UNIT_STATUS}\") - перевірте вручну.",
                    pib_raw, report_date, raw_text=destination_text,
                ))
                continue
        elif _is_ppd_destination(description_text):
            row_status = "ППД"
        elif _is_health_leave_destination(description_text):
            row_status = "ВПСЗ"
        elif _is_family_leave_destination(description_text):
            row_status = "ВПС"
        elif status == "ВД" and _is_training_destination(description_text):
            row_status = "Навч"
        elif status is not None:
            row_status = status
        else:
            unresolved.append(_unresolved(f"Невідомий заголовок пункту рапорту: \"{heading_text.strip()}\".", pib_raw, report_date))
            continue

        date_text = cells[date_col].text.strip() if date_col is not None else ""
        event = {
            "pib_raw": pib_raw, "status": row_status, "effective_date": _parse_date(date_text) or report_date,
            "heading": heading_text.strip(), "report_date": report_date,
        }
        if _AFTER_HOURS_RE.search(date_text):
            event["after_hours"] = True
        events.append(event)

    return events, unresolved


def extract_status_events(file_path):
    """Читає один щоденний рапорт (.docx) і повертає (events, unresolved).

    events - [{"pib_raw", "status", "effective_date", "heading",
    "report_date"}], по одному на кожен рядок таблиці ПРИБУЛИ/ВИБУЛИ, чий
    підзаголовок розпізнано (STATUS_HEADING_RULES). effective_date - з колонки
    "Дата прибуття"/"Дата вибуття" цієї таблиці; вона може відрізнятись від
    дати рапорту (зміна "після 17:00" потрапляє в наступний рапорт, але зі
    справжньою датою в клітинці). Якщо клітинку не розпізнано - дата рапорту.

    unresolved - усе, що не вдалося однозначно перетворити на подію: рядки
    таблиць з нерозпізнаним підзаголовком чи без колонки ПІБ і абзаци розділу
    "Поза межами складу сил та засобів ...:" (вільний текст без таблиці,
    FREEFORM_GROUP_KEYWORD). Подіями стають лише абзаци з відомими зворотами:
    - сталий зворот з FREEFORM_STATUS_PATTERNS (_freeform_event, одна подія);
    - "виписаний ... ВЛК ... потребує відпустки ..." (_discharge_vlk_leave_events,
      до двох послідовних подій з одного речення);
    - "виписаний ... самовільно залишив військову частину (СЗЧ)"
      (_discharge_then_awol_event, "СЗЧ" на дату виписки);
    - "З ДАТА відпустка за станом здоров'я ..." (_confirmed_health_leave_event,
      "ВПСЗ" на явну дату з речення);
    - "завершив ВЛК"/"виписаний ..., з ДАТА направляється у відпустку ..."
      (_dated_treatment_leave_event, теж "ВПСЗ" на явну дату).
    Не розглядаються взагалі рядок підпису наприкінці рапорту
    (_is_signatory_paragraph), усе після нього та голий підзаголовок без ПІБ
    (_is_bare_status_heading, напр. "СЗЧ:" перед реченням про людину - він
    описує групу, а не подію). Решта вільного тексту не має сталої
    структури, тож навмисно не розбирається на статус і дату, а лишається
    для ручної перевірки.

    Окремо - заголовок-список "ПРИБУЛИ до пункту постійно(ї) дислокації
    військової частини ...:" (_is_arrival_list_heading,
    ARRIVAL_LIST_HEADING_KEYWORDS): кожен наступний абзац "звання ПІБ" (того
    самого стилю "List Paragraph", без таблиці) автоматично стає подією
    DEFAULT_STATUS ("РВЗ") на дату рапорту - це простий список з однозначним
    сенсом. Список закінчується на першому абзаці того самого стилю без ПІБ
    (_NAME_RE) - він уже звичайний заголовок (напр. наступного розділу).

    І ще окремо, лише всередині "Поза межами...", - заголовок-речення
    "Продовжено відрядження з ДЕНЬ МІСЯЦЬ РІК року терміном на N днів до
    [установа] ...:" (_extended_business_trip_date, _EXTENDED_BUSINESS_TRIP_RE)
    без ПІБ, стиль "Normal". Під ним - простий список "звання ПІБ"
    (_is_bare_name_paragraph - у рядку лише ПІБ), і кожен автоматично стає
    подією "ВД" на дату з заголовка. Список закінчується на першому рядку, що
    не є голим ПІБ, чи на іншому заголовку."""
    doc = Document(file_path)
    report_date = report_date_from_filename(file_path)

    events, unresolved = [], []
    current_heading = ""
    in_freeform_section = False
    in_arrival_section = False
    arrival_heading_text = ""
    pending_group_status = None
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
                pending_group_status = None
                if in_arrival_section:
                    arrival_heading_text = text
            elif in_freeform_section:
                name_match = _NAME_RE.search(text)

                if pending_group_status is not None and _is_bare_name_paragraph(name_match, text):
                    events.append({
                        "pib_raw": name_match.group(), "status": pending_group_status["status"],
                        "effective_date": pending_group_status["effective_date"],
                        "heading": pending_group_status["heading"], "report_date": report_date,
                    })
                    current_heading = text
                    continue
                pending_group_status = None

                if name_match is None:
                    extended_trip_date = _extended_business_trip_date(text)
                    if extended_trip_date is not None:
                        pending_group_status = {"status": "ВД", "effective_date": extended_trip_date, "heading": text}
                        current_heading = text
                        continue

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
