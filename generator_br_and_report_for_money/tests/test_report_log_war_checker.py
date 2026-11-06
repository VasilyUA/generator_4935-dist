import os
import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from docx import Document
from openpyxl import load_workbook

import checker_accounting.report_log_war_checker as checker
from checker_accounting.checker import GREEN_FILL, RED_FILL

_TABLE_HEADERS = ["№", "Посада", "Військове звання", "Прізвище, ім'я, по батькові", "Період участі", "Кількість днів", "Підстава для виплати"]


def _add_section(doc, heading_text, data_rows):
    """data_rows - список [посада, звання, піб, період, дні, підстава] - будує
    Heading 1 + таблицю (7 колонок, як create_table у formatting/docx_utils.py),
    той самий помічник, що й tests/test_report_document_reader.py."""
    doc.add_heading(heading_text, level=1)
    table = doc.add_table(rows=1 + len(data_rows), cols=7)
    for col_idx, header in enumerate(_TABLE_HEADERS):
        table.rows[0].cells[col_idx].text = header
    for row_idx, values in enumerate(data_rows, start=1):
        table.rows[row_idx].cells[0].text = str(row_idx)
        for col_idx, value in enumerate(values, start=1):
            table.rows[row_idx].cells[col_idx].text = value


def _write_report(path, sections):
    doc = Document()
    for heading_text, data_rows in sections:
        _add_section(doc, heading_text, data_rows)
    doc.save(str(path))
    return str(path)


def _write_log_war_doc(path, paragraphs):
    doc = Document()
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    cell.paragraphs[0].text = paragraphs[0] if paragraphs else ""
    for text in paragraphs[1:]:
        cell.add_paragraph(text)
    doc.save(str(path))
    return str(path)


# -------------------------
# _mentioned
# -------------------------
def test_mentioned_matches_case_and_whitespace_insensitively():
    assert checker._mentioned("ПЕРШИЙ Перший Перший", "  залучено   перший ПЕРШИЙ перший ; ")


def test_mentioned_false_for_empty_text():
    assert not checker._mentioned("ПЕРШИЙ Перший Перший", "")


def test_mentioned_false_when_absent():
    assert not checker._mentioned("ПЕРШИЙ Перший Перший", "залучено ДРУГИЙ Другий Другий;")


def test_mentioned_false_for_blank_pib():
    """normalize_name("") дає порожній рядок - _name_search_pattern не має
    жодного слова для побудови regex і повертає None (а не порожній/
    "збігається з усім" патерн)."""
    assert not checker._mentioned("", "залучено ПЕРШИЙ Перший Перший;")


def test_mentioned_tolerates_genitive_case_declension_of_name_and_patronymic():
    """РЕГРЕСІЯ (виявлено користувачем на реальному зразку): "Хід виконання"
    - вільний наратив, що відмінює ім'я/по батькові за відмінком речення
    (родовий: "...з батальйонного району оборони [кого?] Другого Третього"),
    а НЕ називний, як записано в рапорті ("Другий Третій") - точний збіг
    підрядка тут НЕ мав би спрацювати, толерантний до відмінка - мусить."""
    text = "Виведено з батальйонного району оборони ПЕРШИЙ Другого Третього;"
    assert checker._mentioned("ПЕРШИЙ Другий Третій", text)


def test_mentioned_different_surname_with_same_name_and_patronymic_does_not_match():
    """Прізвище ТЕЖ толерантне до відмінка (див. тест нижче), але це не
    означає "будь-яке прізвище" - інша людина з тим самим (відмінюваним)
    ім'ям/по батькові, але ІНШИМ прізвищем (інший КОРІНЬ - "ДРУГ" замість
    "ПЕРШ") не має збігатись."""
    text = "Виведено ДРУГИЙ Другого Третього;"
    assert not checker._mentioned("ПЕРШИЙ Другий Третій", text)


def test_mentioned_tolerates_adjective_type_surname_declension():
    """РЕГРЕСІЯ (виявлено користувачем, ДРУГИЙ реальний зразок): прізвища
    прикметникового типу (закінчення "-ИЙ"/"-СЬКИЙ" тощо, напр. реальне
    "ЧЕКЕРСЬКИЙ") відмінюються ЯК ПРИКМЕТНИК - "ЧЕКЕРСЬКИЙ" (називний) у
    рапорті, але "ЧЕКЕРСЬКОМУ" (давальний) у наративі ЖБД - раніше прізвище
    звірялось ТОЧНО (без толерантності до відмінка, на основі ІНШОГО зразка,
    де прізвище не відмінювалось), тож такий запис пропускався. Синтетичний
    приклад тут - "ПЕРШИЙ" як прізвище (сам порядковий числівник -
    прикметникового типу, відмінюється так само: ПЕРШИЙ -> ПЕРШОМУ)."""
    text = "Наказано ПЕРШОМУ Другому Третьому передислокуватись;"
    assert checker._mentioned("ПЕРШИЙ Другий Третій", text)


def test_mentioned_tolerates_replacive_ending_on_a_short_five_letter_word():
    """РЕГРЕСІЯ (виявлено користувачем, ТРЕТІЙ реальний зразок): коротке
    (РІВНО 5 літер, напр. реальне "ПАВЛО") слово з ЗАМІННИМ (не додатковим)
    закінченням відмінка - остання літера ЗАМІНЮЄТЬСЯ ("ПАВЛО" (називний) ->
    "ПАВЛУ" (давальний), а НЕ "ПАВЛО"+"У"="ПАВЛОУ", так не буває) - поріг
    "> 5" раніше НЕ обрізав такі короткі слова взагалі, тож точний 5-літерний
    підрядок НЕ збігався з формою, де 5-та літера інша (давальний/родовий).
    Поріг знижено до "> 3" - тепер обрізаються й короткі (4-5 літер) слова.
    Синтетичний приклад тут (не справжнє ім'я) - "ПЕРШО" (5 літер, закінчення
    "-О") -> "ПЕРШУ" (закінчення замінене на "-У", як у "ПАВЛО"/"ПАВЛУ")."""
    text = "Наказано ПЕРШИЙ ПЕРШУ Третій передислокуватись;"
    assert checker._mentioned("ПЕРШИЙ ПЕРШО Третій", text)


# -------------------------
# _covered_value - виняток пункту 30 (підтверджено користувачем): відсутній
# ПІБ у пункті 30 - ПОМИЛКА лише якщо в тому самому абзаці є фраза "з
# батальйонного району оборони"; інакше - "прощено" (None), не помилка.
# -------------------------
def test_covered_value_true_when_mentioned_regardless_of_waivable():
    assert checker._covered_value("ПЕРШИЙ Перший Перший", "залучено ПЕРШИЙ Перший Перший;", is_waivable=True) is True
    assert checker._covered_value("ПЕРШИЙ Перший Перший", "залучено ПЕРШИЙ Перший Перший;", is_waivable=False) is True


def test_covered_value_false_when_absent_and_not_waivable():
    assert checker._covered_value("ПЕРШИЙ Перший Перший", "залучено ДРУГИЙ Другий Другий;", is_waivable=False) is False


def test_covered_value_none_when_waivable_and_phrase_absent():
    assert checker._covered_value("ПЕРШИЙ Перший Перший", "залучено ДРУГИЙ Другий Другий;", is_waivable=True) is None


def test_covered_value_false_when_waivable_but_phrase_present():
    text = "залучено ДРУГИЙ Другий Другий; відведено з батальйонного району оборони."
    assert checker._covered_value("ПЕРШИЙ Перший Перший", text, is_waivable=True) is False


def test_covered_value_none_when_waivable_and_text_empty():
    assert checker._covered_value("ПЕРШИЙ Перший Перший", "", is_waivable=True) is None


# -------------------------
# _checks_for_entry_old / _checks_for_entry_new / _checks_for_entry_mixed /
# _missing_days_count - основна логіка покриття, "довгий" формат (список
# окремих перевірок дата+пункт, підтверджено користувачем), а не
# {дата: покрито}. Три методики (підтверджено користувачем): "стара" - лише
# вхід/вихід періоду; "нова" - кожен день окремо; "змішана" - стара до
# автоматично визначеної межі, потім нова.
# -------------------------
def _entry(pib, period, raw_value=100, heading_number=None):
    """heading_number - за замовчуванням узгоджений із raw_value (30->"1",
    інакше->"2"), як у реальному рапорті - явно передати інше значення можна
    для тестів, що НАВМИСНЕ перевіряють невідповідність (напр. 30к під іншим
    номером пункту)."""
    if heading_number is None:
        heading_number = "1" if raw_value == 30 else "2"
    return {
        "raw_value": raw_value, "ПОСАДА": "Стрілець", "ЗВАННЯ": "сержант", "ПІБ": pib, "ПЕРІОД": period,
        "ДНІ": "", "ПІДСТАВА": "", "ДАТА_ЗНИКНЕННЯ": None, "НОМЕР_ПУНКТУ": heading_number,
    }


def _log_war_files(*dates):
    return {date: f"/fake/ЖБД {date.strftime('%d.%m.%Y')}.docx" for date in dates}


# --- _checks_for_entry_new ("нова": ЛИШЕ присутність в пункті "6" САМОГО
# ОСТАННЬОГО дня під-діапазону - ВИПРАВЛЕНО ВДРУГЕ, підтверджено користувачем
# на реальних зразках за липень: ДЯКОВИЧ/МИРОНЕНКО (не по парі на кожен
# окремий день), ТІЩЕНКО (навіть на межі виходу - НЕ пара "5"+"6 наступного
# дня", а лише "6" САМОГО останнього дня - людина може вибути ПОЗА системою
# ЖБД узагалі, напр. у шпиталь, тож перевіряти наступний день безглуздо)) ---
def test_checks_new_only_last_day_of_subrange_gets_checked_not_every_day():
    """Період 01.05-02.05 (два дні) - ОДНА перевірка, (02.05,"6") - РЕГРЕСІЯ
    (виявлено користувачем, реальні зразки за липень): реальні документи НЕ
    підтверджують перший день безперервного перебування (тут - 01.05) окремим
    завданням, і НЕ вимагають перевірки наступного дня (03.05) - лише
    присутність у пункті "6" САМОГО останнього дня перебування."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-02.05.2026")
    sections_by_date = {
        datetime(2026, 5, 1): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 5, 2): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": "залучено ПЕРШИЙ Перший Перший;"},
        datetime(2026, 5, 3): {"section5": "", "section6_hid": ""},
    }
    dates = [datetime(2026, 5, 1), datetime(2026, 5, 2), datetime(2026, 5, 3)]
    checks = checker._checks_for_entry_new(entry, _log_war_files(*dates), sections_by_date)
    assert [(c["date"], c["point"]) for c in checks] == [(datetime(2026, 5, 2), "6")]
    assert all(c["covered"] for c in checks)
    assert checker._missing_days_count(checks) == 0


def test_checks_new_falls_back_to_next_day_when_last_day_not_mentioned():
    """РЕГРЕСІЯ (виявлено користувачем, реальний зразок - САСЮК): "Хід
    виконання" за ОСТАННІЙ день перебування (16.07) може фігурувати не в
    файлі ЦЬОГО ж дня, а в файлі НАСТУПНОГО (17.07 - звичайна конвенція:
    "Хід виконання" підтверджує ВЧОРАШНЄ виконання) - НАВІТЬ якщо 17.07
    людина вже у відрядженні. Перевірка тоді має "провалитись" на 16.07 і
    ВІДРАЗУ спробувати 17.07, перш ніж вважати це помилкою."""
    entry = _entry("ПЕРШИЙ Перший Перший", "15.05.2026-16.05.2026")
    sections_by_date = {
        datetime(2026, 5, 16): {"section5": "", "section6_hid": "Нічого про цю людину."},
        datetime(2026, 5, 17): {"section5": "", "section6_hid": "залучено ПЕРШИЙ Перший Перший;"},
    }
    dates = [datetime(2026, 5, 16), datetime(2026, 5, 17)]
    checks = checker._checks_for_entry_new(entry, _log_war_files(*dates), sections_by_date)
    assert [(c["date"], c["point"]) for c in checks] == [(datetime(2026, 5, 17), "6")]
    assert all(c["covered"] for c in checks)
    assert checker._missing_days_count(checks) == 0


def test_checks_new_error_recorded_against_next_day_when_neither_day_mentions_pib():
    """Ні останній день (16.05), ні наступний (17.05) не згадують ПІБ -
    ЛИШЕ ТОДІ це справжня помилка, і рядок помилки - за НАСТУПНИМ днем (там,
    де перевірку справді завершено)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "15.05.2026-16.05.2026")
    sections_by_date = {
        datetime(2026, 5, 16): {"section5": "", "section6_hid": ""},
        datetime(2026, 5, 17): {"section5": "", "section6_hid": ""},
    }
    dates = [datetime(2026, 5, 16), datetime(2026, 5, 17)]
    checks = checker._checks_for_entry_new(entry, _log_war_files(*dates), sections_by_date)
    assert [(c["date"], c["point"]) for c in checks] == [(datetime(2026, 5, 17), "6")]
    assert all(c["covered"] is False for c in checks)
    assert checker._missing_days_count(checks) == 1


def test_checks_new_missing_check_is_an_error():
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026")
    sections_by_date = {datetime(2026, 5, 1): {"section5": "", "section6_hid": ""}}
    checks = checker._checks_for_entry_new(entry, _log_war_files(datetime(2026, 5, 1)), sections_by_date)
    assert [c["point"] for c in checks] == ["6"]
    assert all(c["covered"] is False for c in checks)
    assert checker._missing_days_count(checks) == 1


def test_checks_new_missing_document_entirely_uses_placeholder_file_label():
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026")
    checks = checker._checks_for_entry_new(entry, {}, {})
    assert len(checks) == 1
    assert all(not c["covered"] for c in checks)
    assert all(checker._NO_FILE_LABEL in c["file_name"] for c in checks)
    assert checker._missing_days_count(checks) == 1


def test_checks_new_second_subrange_handled_independently():
    """"01.05-01.05; 03.05-03.05" - ДВА під-діапазони - кожен зі своєю
    перевіркою (свого останнього дня, тут - самого себе)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026; 03.05.2026-03.05.2026")
    sections_by_date = {
        datetime(2026, 5, 1): {"section5": "", "section6_hid": "залучено ПЕРШИЙ Перший Перший;"},
        datetime(2026, 5, 3): {"section5": "", "section6_hid": "залучено ПЕРШИЙ Перший Перший;"},
    }
    dates = [datetime(2026, 5, 1), datetime(2026, 5, 3)]
    checks = checker._checks_for_entry_new(entry, _log_war_files(*dates), sections_by_date)
    assert [(c["date"], c["point"]) for c in checks] == [
        (datetime(2026, 5, 1), "6"), (datetime(2026, 5, 3), "6"),
    ]
    assert all(c["covered"] for c in checks)


def test_checks_new_point_30_missing_pib_without_battalion_phrase_is_waived():
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026", raw_value=30)
    sections_by_date = {datetime(2026, 5, 1): {"section5": "", "section6_hid": ""}}
    checks = checker._checks_for_entry_new(entry, _log_war_files(datetime(2026, 5, 1)), sections_by_date)
    assert all(c["covered"] is None for c in checks)
    assert checker._missing_days_count(checks) == 0


def test_checks_new_ignores_person_dates_no_boundary_to_double_count():
    """person_dates не впливає на цю методику - перевірка ЛИШЕ ОДНОГО дня
    (останнього дня під-діапазону) сама по собі не залежить від того, чим
    зайнята людина наступного дня в ІНШОМУ записі (немає межі з двох
    календарних днів, яку можна було б випадково перевірити двічі)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "13.05.2026-14.05.2026")
    person_dates = {datetime(2026, 5, 15)}  # покрито ІНШИМ записом
    checks_with = checker._checks_for_entry_new(entry, {}, {}, person_dates=person_dates)
    checks_without = checker._checks_for_entry_new(entry, {}, {}, person_dates=frozenset())
    assert checks_with == checks_without
    # 14.05 (останній день) не знайдено -> fallback на 15.05 (наступний день).
    assert [(c["date"], c["point"]) for c in checks_with] == [(datetime(2026, 5, 15), "6")]


# --- _is_month_start / _is_month_end ---
def test_is_month_start_true_only_for_first_day():
    assert checker._is_month_start(datetime(2026, 5, 1)) is True
    assert checker._is_month_start(datetime(2026, 5, 2)) is False


def test_is_month_end_true_for_last_day_and_one_day_before():
    assert checker._is_month_end(datetime(2026, 5, 31)) is True  # травень - 31 день
    assert checker._is_month_end(datetime(2026, 5, 30)) is True  # допуск в 1 день
    assert checker._is_month_end(datetime(2026, 5, 29)) is False
    assert checker._is_month_end(datetime(2026, 4, 30)) is True  # квітень - 30 днів
    assert checker._is_month_end(datetime(2026, 4, 29)) is True


# --- _checks_for_entry_old ("стара": лише вхід/вихід періоду) ---
def test_checks_old_entry_and_exit_only_middle_days_not_checked():
    """Реальний приклад, підтверджений користувачем: період 02.04-15.04 ->
    перевірки ЛИШЕ за 01.04(п.5)+02.04(п.6) (вхід) і 14.04(п.5)+15.04(п.6)
    (вихід) - дні 03.04-13.04 узагалі НЕ перевіряються (0 рядків для них)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "02.04.2026-15.04.2026")
    sections_by_date = {
        datetime(2026, 4, 1): {"section5": "завдання ПЕРШИЙ Перший Перший в батальйонний район оборони", "section6_hid": ""},
        datetime(2026, 4, 2): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший увійшов;"},
        datetime(2026, 4, 14): {"section5": "завдання ПЕРШИЙ Перший Перший з батальйонного району оборони", "section6_hid": ""},
        datetime(2026, 4, 15): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший вийшов;"},
    }
    dates = [datetime(2026, 4, 1), datetime(2026, 4, 2), datetime(2026, 4, 14), datetime(2026, 4, 15)]
    checks = checker._checks_for_entry_old(entry, _log_war_files(*dates), sections_by_date)
    assert [(c["date"], c["point"]) for c in checks] == [
        (datetime(2026, 4, 1), "5"), (datetime(2026, 4, 2), "6"),
        (datetime(2026, 4, 14), "5"), (datetime(2026, 4, 15), "6"),
    ]
    assert all(c["covered"] for c in checks)
    assert checker._missing_days_count(checks) == 0


def test_checks_old_single_day_subrange_dedupes_entry_and_exit_into_one_pair():
    """Одноденний під-діапазон - вхід і вихід збігаються (та сама дата) -
    лише ОДНА пара перевірок, а не дві однакові."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026")
    sections_by_date = {
        datetime(2026, 4, 30): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 5, 1): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший;"},
    }
    checks = checker._checks_for_entry_old(entry, _log_war_files(datetime(2026, 4, 30), datetime(2026, 5, 1)), sections_by_date)
    assert [(c["date"], c["point"]) for c in checks] == [(datetime(2026, 4, 30), "5"), (datetime(2026, 5, 1), "6")]


def test_checks_old_entry_not_covered_is_an_error():
    """Вхід охоплює ДВІ РІЗНІ дати (день до + день входу), тож
    _missing_days_count (рахує ДАТИ) повертає 2. Період 10.05-10.05 (НЕ на
    межі місяця) - щоб не зачепити _waive_boundary_absence. Файли ЖБД за ОБИ
    дати ПРИСУТНІ (_log_war_files) - інакше нова перевірка "файл відсутній"
    сама по собі дала б None, і тест перестав би значуще перевіряти саме
    "файл є, але людину не згадано"."""
    entry = _entry("ПЕРШИЙ Перший Перший", "10.05.2026-10.05.2026")
    checks = checker._checks_for_entry_old(
        entry, _log_war_files(datetime(2026, 5, 9), datetime(2026, 5, 10)), {},
    )
    assert all(c["covered"] is False for c in checks)
    assert checker._missing_days_count(checks) == 2


def test_checks_old_point_30_waived_when_topic_absent():
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026", raw_value=30)
    checks = checker._checks_for_entry_old(
        entry, _log_war_files(datetime(2026, 4, 30), datetime(2026, 5, 1)), {},
    )
    assert all(c["covered"] is None for c in checks)
    assert checker._missing_days_count(checks) == 0


def test_checks_old_entry_skipped_when_previous_day_already_covered_by_same_person():
    """РЕГРЕСІЯ (виявлено користувачем, реальний зразок): людина БЕЗ РОЗРИВУ
    переходить з одного пункту рапорту в інший (напр. 100к до 04.05, 30к від
    05.05) - реальне переміщення в ЖБД задокументоване ОДИН раз (як вихід із
    100к), а НЕ ДВІЧІ - "вхід" у 30к (04.05(п.5)+05.05(п.6)) НЕ має шукати
    ДРУГЕ, неіснуюче підтвердження, якщо 04.05 вже покрито ІНШИМ записом ЦІЄЇ
    Ж людини (person_dates)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "05.05.2026-10.05.2026", raw_value=30)
    person_dates = {datetime(2026, 5, d) for d in range(1, 5)}  # 01.05-04.05 (інший пункт)
    checks = checker._checks_for_entry_old(entry, {}, {}, person_dates=person_dates)
    # Немає ЖОДНОЇ перевірки за 04.05(п.5)/05.05(п.6) - вхід "прощено" як
    # безшовне продовження.
    assert (datetime(2026, 5, 4), "5") not in {(c["date"], c["point"]) for c in checks}
    assert (datetime(2026, 5, 5), "6") not in {(c["date"], c["point"]) for c in checks}
    # Вихід (09.05/10.05) - НЕ покритий іншим записом - перевіряється як завжди.
    assert (datetime(2026, 5, 9), "5") in {(c["date"], c["point"]) for c in checks}
    assert (datetime(2026, 5, 10), "6") in {(c["date"], c["point"]) for c in checks}


def test_checks_old_exit_skipped_when_next_day_already_covered_by_same_person():
    """Симетрично - ВИХІД (09.05(п.5)+10.05(п.6)) НЕ перевіряється, якщо
    11.05 вже покрито ІНШИМ записом ЦІЄЇ Ж людини (продовження в іншому
    пункті, без розриву)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "05.05.2026-10.05.2026", raw_value=30)
    person_dates = {datetime(2026, 5, d) for d in range(11, 16)}  # 11.05-15.05 (інший пункт)
    checks = checker._checks_for_entry_old(entry, {}, {}, person_dates=person_dates)
    assert (datetime(2026, 5, 9), "5") not in {(c["date"], c["point"]) for c in checks}
    assert (datetime(2026, 5, 10), "6") not in {(c["date"], c["point"]) for c in checks}
    assert (datetime(2026, 5, 4), "5") in {(c["date"], c["point"]) for c in checks}
    assert (datetime(2026, 5, 5), "6") in {(c["date"], c["point"]) for c in checks}


def test_checks_old_single_day_subrange_only_checks_entry_adjacency_not_exit():
    """Одноденний під-діапазон - пара (D-1,"5")+(D,"6") ЗАВЖДИ представляє
    ВХІД (єдина перевірка для такого дня) - сусідній день ПІСЛЯ (D+1),
    покритий іншим записом, НЕ має жодного стосунку до цієї пари і не
    впливає на неї."""
    entry = _entry("ПЕРШИЙ Перший Перший", "05.05.2026-05.05.2026", raw_value=30)
    person_dates = {datetime(2026, 5, 6)}  # ЛИШЕ день ПІСЛЯ покритий іншим записом
    checks = checker._checks_for_entry_old(entry, {}, {}, person_dates=person_dates)
    # Вхід (04.05/05.05) УСЕ ОДНО перевіряється - 06.05 не має значення тут.
    assert (datetime(2026, 5, 4), "5") in {(c["date"], c["point"]) for c in checks}
    assert (datetime(2026, 5, 5), "6") in {(c["date"], c["point"]) for c in checks}


def test_checks_old_full_month_absence_is_waived_not_an_error():
    """Підтверджено користувачем: якщо ПІБ ЦІЛКОВИТО відсутній у ЖБД за ВЕСЬ
    календарний місяць (01.05.2026-31.05.2026) для пункту 1 чи 2 - "прощено"
    (covered None), а НЕ помилка - реальний вхід/вихід, найімовірніше, стався
    поза межами цього місяця. Файли ЖБД за ВСІ 4 дати ПРИСУТНІ - щоб
    перевіряти саме календарне пробачення, а не нову перевірку "файл
    відсутній"."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-31.05.2026")
    log_war_files = _log_war_files(
        datetime(2026, 4, 30), datetime(2026, 5, 1), datetime(2026, 5, 30), datetime(2026, 5, 31),
    )
    checks = checker._checks_for_entry_old(entry, log_war_files, {})
    assert len(checks) == 4  # вхід (30.04+01.05) і вихід (30.05+31.05)
    assert all(c["covered"] is None for c in checks)
    assert checker._missing_days_count(checks) == 0


def test_checks_old_full_month_partial_evidence_is_also_waived():
    """Підтверджено користувачем (уточнення): ЧАСТКОВА відсутність за весь
    місяць (напр. вхід знайдено, вихід - ні) ТЕЖ "прощується" - реальний
    вхід міг статись БУДЬ-ЯКОГО з перших днів місяця (скрипт не знає точної
    дати), тож НЕ вимагається, щоб АБСОЛЮТНО ВСЕ було відсутнє. Знайдені
    перевірки лишаються "Знайдено" (True), а НЕ знайдені - "прощені" (None)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-31.05.2026")
    sections_by_date = {
        datetime(2026, 4, 30): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 5, 1): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший;"},
    }
    log_war_files = _log_war_files(
        datetime(2026, 4, 30), datetime(2026, 5, 1), datetime(2026, 5, 30), datetime(2026, 5, 31),
    )
    checks = checker._checks_for_entry_old(entry, log_war_files, sections_by_date)
    entry_checks = [c for c in checks if c["date"] in (datetime(2026, 4, 30), datetime(2026, 5, 1))]
    exit_checks = [c for c in checks if c["date"] in (datetime(2026, 5, 30), datetime(2026, 5, 31))]
    assert all(c["covered"] is True for c in entry_checks)
    assert all(c["covered"] is None for c in exit_checks)  # прощено, НЕ помилка
    assert checker._missing_days_count(checks) == 0


def test_checks_old_full_month_fully_confirmed_is_not_touched():
    """Якщо ВЕСЬ місяць ПОВНІСТЮ підтверджений (і вхід, і вихід реально
    знайдені) - нічого прощати не потрібно, усе лишається "Знайдено"."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.04.2026-30.04.2026")
    sections_by_date = {
        datetime(2026, 3, 31): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 4, 1): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший;"},
        datetime(2026, 4, 29): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 4, 30): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший;"},
    }
    log_war_files = _log_war_files(*sections_by_date.keys())
    checks = checker._checks_for_entry_old(entry, log_war_files, sections_by_date)
    assert all(c["covered"] is True for c in checks)


def test_checks_old_partial_month_absence_is_still_an_error():
    """Період, що НЕ охоплює весь календарний місяць (напр. 05.05-20.05) -
    повна відсутність УСЕ ОДНО вважається помилкою (умова стосується ЛИШЕ
    ЦІЛОГО місяця, підтверджено користувачем). Файли ЖБД за ОБИ межі
    ПРИСУТНІ - щоб перевіряти саме "файл є, людину не згадано"."""
    entry = _entry("ПЕРШИЙ Перший Перший", "05.05.2026-20.05.2026")
    log_war_files = _log_war_files(
        datetime(2026, 5, 4), datetime(2026, 5, 5), datetime(2026, 5, 19), datetime(2026, 5, 20),
    )
    checks = checker._checks_for_entry_old(entry, log_war_files, {})
    assert all(c["covered"] is False for c in checks)


def test_checks_old_entry_waived_independently_when_start_is_month_start_but_end_is_not():
    """РЕГРЕСІЯ (уточнено користувачем): ВХІД прощається САМ ПО СОБІ, якщо
    start_date - 1-ше число місяця, НЕЗАЛЕЖНО від того, чи ВЕСЬ період -
    повний календарний місяць (напр. "01.05-10.05" - вихід (10.05) явно НЕ
    на межі місяця, тож лишається помилкою, якщо не знайдений; вхід (30.04+
    01.05) - прощається)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-10.05.2026")
    log_war_files = _log_war_files(
        datetime(2026, 4, 30), datetime(2026, 5, 1), datetime(2026, 5, 9), datetime(2026, 5, 10),
    )
    checks = checker._checks_for_entry_old(entry, log_war_files, {})
    entry_checks = [c for c in checks if c["date"] in (datetime(2026, 4, 30), datetime(2026, 5, 1))]
    exit_checks = [c for c in checks if c["date"] in (datetime(2026, 5, 9), datetime(2026, 5, 10))]
    assert all(c["covered"] is None for c in entry_checks)
    assert all(c["covered"] is False for c in exit_checks)


def test_checks_old_exit_waived_independently_when_end_is_month_end_but_start_is_not():
    """РЕГРЕСІЯ (уточнено користувачем): ВИХІД прощається САМ ПО СОБІ, якщо
    end_date - в останніх днях свого місяця, НЕЗАЛЕЖНО від того, чи ВХІД
    (start_date) теж на межі (напр. "08.05-31.05" - вхід (08.05) явно НЕ на
    межі місяця, тож лишається помилкою; вихід (30.05+31.05) - прощається)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "08.05.2026-31.05.2026")
    log_war_files = _log_war_files(
        datetime(2026, 5, 7), datetime(2026, 5, 8), datetime(2026, 5, 30), datetime(2026, 5, 31),
    )
    checks = checker._checks_for_entry_old(entry, log_war_files, {})
    entry_checks = [c for c in checks if c["date"] in (datetime(2026, 5, 7), datetime(2026, 5, 8))]
    exit_checks = [c for c in checks if c["date"] in (datetime(2026, 5, 30), datetime(2026, 5, 31))]
    assert all(c["covered"] is False for c in entry_checks)
    assert all(c["covered"] is None for c in exit_checks)


def test_checks_old_exit_waived_when_one_day_short_of_true_month_end():
    """РЕГРЕСІЯ (виявлено користувачем, реальний зразок): період типу
    "01.05.2026-30.05.2026" (травень має 31 день, а не 30) - вихід (30.05)
    ВСЕ ОДНО прощається (_is_month_end допускає 1 день допуску) - реальні
    рапорти регулярно закінчують період на день раніше буквального кінця
    місяця."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-30.05.2026")
    log_war_files = _log_war_files(
        datetime(2026, 4, 30), datetime(2026, 5, 1), datetime(2026, 5, 29), datetime(2026, 5, 30),
    )
    checks = checker._checks_for_entry_old(entry, log_war_files, {})
    assert all(c["covered"] is None for c in checks)


# --- Файл ЖБД за конкретну дату межі відсутній у log_war_files - "не
# помилка" (уточнено користувачем, реальний приклад: вхід 03.04.2026,
# вихід 16.05.2026, а завантажена папка ЖБД - лише за травень) ---
def test_checks_old_both_boundary_dates_missing_files_are_waived_not_error():
    """Обидві дати пари (день-до + день-входу) ВІДСУТНІ в log_war_files
    (симулює вхід 03.04, коли завантажена лише травнева папка) - ОБИДВІ
    covered=None, НЕЗАЛЕЖНО від календарної евристики (03.04 - НЕ 1-ше
    число жодного місяця, тож БЕЗ цього механізму це й досі була б False)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "03.04.2026-03.04.2026")
    checks = checker._checks_for_entry_old(entry, {}, {})
    assert len(checks) == 2
    assert all(c["covered"] is None for c in checks)
    assert checker._missing_days_count(checks) == 0


def test_checks_old_only_missing_half_of_boundary_is_waived_present_half_checked_for_real():
    """Нюанс, уточнений користувачем: якщо ЛИШЕ ОДНА дата пари відсутня в
    log_war_files, а ІНША присутня - відсутня прощається (None), а присутня
    перевіряється ЗВИЧАЙНО (реальний found/not-found), а НЕ прощається лише
    через сусідство відсутньої дати. Період 10.05-10.05 (НЕ на межі місяця) -
    щоб не зачепити календарне пробачення (_waive_boundary_absence) і
    ізольовано перевірити САМЕ файлову перевірку."""
    entry = _entry("ПЕРШИЙ Перший Перший", "10.05.2026-10.05.2026")
    log_war_files = _log_war_files(datetime(2026, 5, 10))  # 09.05 НЕМАЄ, 10.05 Є
    sections_by_date = {datetime(2026, 5, 10): {"section5": "", "section6_hid": "Нічого про цю людину."}}
    checks = checker._checks_for_entry_old(entry, log_war_files, sections_by_date)
    by_date = {c["date"]: c["covered"] for c in checks}
    assert by_date[datetime(2026, 5, 9)] is None
    assert by_date[datetime(2026, 5, 10)] is False


def test_checks_old_only_missing_half_of_boundary_present_half_found_is_true():
    """Той самий нюанс, але присутня дата ЗГАДУЄ людину -> True (а не
    прощена, і не хибно позначена через сусідство відсутньої дати)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "10.05.2026-10.05.2026")
    log_war_files = _log_war_files(datetime(2026, 5, 10))
    sections_by_date = {datetime(2026, 5, 10): {"section5": "", "section6_hid": "залучено ПЕРШИЙ Перший Перший;"}}
    checks = checker._checks_for_entry_old(entry, log_war_files, sections_by_date)
    by_date = {c["date"]: c["covered"] for c in checks}
    assert by_date[datetime(2026, 5, 9)] is None
    assert by_date[datetime(2026, 5, 10)] is True


# --- _checks_for_entry_mixed ("змішана": стара до old_cutoff, нова з new_cutoff) ---
def test_checks_mixed_entire_subrange_before_old_cutoff_behaves_like_old():
    entry = _entry("ПЕРШИЙ Перший Перший", "02.04.2026-15.04.2026")
    sections_by_date = {
        datetime(2026, 4, 1): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 4, 2): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший;"},
        datetime(2026, 4, 14): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 4, 15): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший;"},
    }
    dates = [datetime(2026, 4, 1), datetime(2026, 4, 2), datetime(2026, 4, 14), datetime(2026, 4, 15)]
    checks = checker._checks_for_entry_mixed(
        entry, _log_war_files(*dates), sections_by_date,
        old_cutoff=datetime(2026, 4, 30), new_cutoff=datetime(2026, 5, 1),
    )
    assert [(c["date"], c["point"]) for c in checks] == [
        (datetime(2026, 4, 1), "5"), (datetime(2026, 4, 2), "6"),
        (datetime(2026, 4, 14), "5"), (datetime(2026, 4, 15), "6"),
    ]


def test_checks_mixed_entire_subrange_after_new_cutoff_behaves_like_new():
    """Під-діапазон ЦІЛКОМ у "новій" зоні - ОДНА перевірка, (02.05,"6"), як і
    чиста "Нова" - НЕ пара, і НЕ перевірка наступного дня (ВИПРАВЛЕНО ВДРУГЕ,
    підтверджено користувачем на реальних зразках)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-02.05.2026")
    sections_by_date = {
        datetime(2026, 5, 1): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 5, 2): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший;"},
    }
    dates = [datetime(2026, 5, 1), datetime(2026, 5, 2)]
    checks = checker._checks_for_entry_mixed(
        entry, _log_war_files(*dates), sections_by_date,
        old_cutoff=datetime(2026, 4, 8), new_cutoff=datetime(2026, 5, 1),
    )
    assert [(c["date"], c["point"]) for c in checks] == [(datetime(2026, 5, 2), "6")]
    assert all(c["covered"] for c in checks)


def test_checks_mixed_subrange_spanning_the_gap_checks_entry_via_old_and_exit_via_new():
    """Під-діапазон ПЕРЕТИНАЄ розрив стилю (05.05-11.05, межі 06/09.05,
    end_date=11.05 - у "новій" зоні) - ВХІД перевіряється за "старою"
    (04.05(п.5)+05.05(п.6)), а ВИХІД - за "новою" методикою, ЛИШЕ присутність
    у пункті "6" САМОГО справжнього ОСТАННЬОГО дня під-діапазону (11.05,"6") -
    НЕ пара, НЕ перевірка наступного дня (ВИПРАВЛЕНО ВДРУГЕ) і НЕ штучний
    "вихід" на межі розриву (06-10.05 не мали б власних перевірок узагалі)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "05.05.2026-11.05.2026")
    sections_by_date = {
        datetime(2026, 5, 4): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 5, 5): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший увійшов;"},
        datetime(2026, 5, 11): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший;"},
    }
    dates = list(sections_by_date.keys())
    checks = checker._checks_for_entry_mixed(
        entry, _log_war_files(*dates), sections_by_date,
        old_cutoff=datetime(2026, 5, 6), new_cutoff=datetime(2026, 5, 9),
    )
    assert [(c["date"], c["point"]) for c in checks] == [
        (datetime(2026, 5, 4), "5"), (datetime(2026, 5, 5), "6"),
        (datetime(2026, 5, 11), "6"),
    ]
    assert all(c["covered"] for c in checks)
    # Дні строго між входом і виходом (06-10.05) - жодних перевірок.
    assert not any(
        c["date"] in (datetime(2026, 5, 6), datetime(2026, 5, 7), datetime(2026, 5, 8),
                      datetime(2026, 5, 9), datetime(2026, 5, 10))
        for c in checks
    )


def test_checks_mixed_dates_strictly_between_cutoffs_are_never_checked():
    """Період ЦІЛКОМ усередині розриву (07.05-08.05, межі 06/09.05) - жодної
    перевірки взагалі."""
    entry = _entry("ПЕРШИЙ Перший Перший", "07.05.2026-08.05.2026")
    checks = checker._checks_for_entry_mixed(
        entry, {}, {}, old_cutoff=datetime(2026, 5, 6), new_cutoff=datetime(2026, 5, 9),
    )
    assert checks == []


def test_checks_mixed_old_zone_entry_skipped_when_adjacent_day_covered_by_same_person():
    """Та сама РЕГРЕСІЯ, що й для "старої" методики (див.
    test_checks_old_entry_skipped_when_previous_day_already_covered_by_same_person),
    але у "старій" зоні "змішаної" методики - безшовний перехід між пунктами
    ЦІЄЇ Ж людини не потребує окремого підтвердження входу."""
    entry = _entry("ПЕРШИЙ Перший Перший", "05.05.2026-06.05.2026", raw_value=30)
    person_dates = {datetime(2026, 5, 4)}
    checks = checker._checks_for_entry_mixed(
        entry, {}, {}, old_cutoff=datetime(2026, 5, 20), new_cutoff=datetime(2026, 5, 21),
        person_dates=person_dates,
    )
    assert (datetime(2026, 5, 4), "5") not in {(c["date"], c["point"]) for c in checks}
    assert (datetime(2026, 5, 5), "6") not in {(c["date"], c["point"]) for c in checks}
    # Вихід (05.05/06.05) - не покритий іншим записом - перевіряється як завжди.
    assert (datetime(2026, 5, 5), "5") in {(c["date"], c["point"]) for c in checks}
    assert (datetime(2026, 5, 6), "6") in {(c["date"], c["point"]) for c in checks}


def test_checks_mixed_new_zone_exit_ignores_person_dates_no_boundary_to_double_count():
    """"Новостильний" вихід (ЛИШЕ присутність у пункті "6" самого останнього
    дня) НЕ залежить від person_dates - на відміну від "старозонного" входу/
    виходу (пари на межі двох днів), тут немає межі, яку можна було б
    випадково перевірити двічі."""
    entry = _entry("ПЕРШИЙ Перший Перший", "13.05.2026-14.05.2026")
    person_dates = {datetime(2026, 5, 15)}
    checks_with = checker._checks_for_entry_mixed(
        entry, {}, {}, old_cutoff=datetime(2026, 5, 7), new_cutoff=datetime(2026, 5, 9),
        person_dates=person_dates,
    )
    checks_without = checker._checks_for_entry_mixed(
        entry, {}, {}, old_cutoff=datetime(2026, 5, 7), new_cutoff=datetime(2026, 5, 9),
        person_dates=frozenset(),
    )
    assert checks_with == checks_without
    # 14.05 (останній день) не знайдено -> fallback на 15.05 (наступний день).
    assert [(c["date"], c["point"]) for c in checks_with] == [(datetime(2026, 5, 15), "6")]


def test_checks_mixed_full_month_absence_is_waived_not_an_error():
    """Та сама умова, що й для "старої" методики (див.
    test_checks_old_full_month_absence_is_waived_not_an_error), для
    "змішаної" - весь календарний місяць без жодної згадки ПІБ у ЖБД -
    "прощено", а не помилка. Межі поза [01.05-31.05] (весь під-діапазон у
    "старій" зоні) - "новостильний" вихід НЕ прощається цим механізмом
    узагалі (див. _checks_for_entry_new/_add_new_style_exit), тож тут
    навмисне перевіряється саме "старозонний" сценарій."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-31.05.2026")
    log_war_files = _log_war_files(
        datetime(2026, 4, 30), datetime(2026, 5, 1), datetime(2026, 5, 30), datetime(2026, 5, 31),
    )
    checks = checker._checks_for_entry_mixed(
        entry, log_war_files, {}, old_cutoff=datetime(2026, 5, 31), new_cutoff=datetime(2026, 6, 1),
    )
    assert len(checks) > 0
    assert all(c["covered"] is None for c in checks)
    assert checker._missing_days_count(checks) == 0


def test_checks_mixed_full_month_partial_evidence_is_also_waived():
    """Та сама умова, що й test_checks_old_full_month_partial_evidence_is_also_waived
    (уточнення користувача), для "змішаної" - ЧАСТКОВА відсутність за весь
    місяць теж "прощується" (не вимагається АБСОЛЮТНА відсутність)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-31.05.2026")
    sections_by_date = {
        datetime(2026, 4, 30): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 5, 1): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший;"},
    }
    log_war_files = _log_war_files(
        datetime(2026, 4, 30), datetime(2026, 5, 1), datetime(2026, 5, 30), datetime(2026, 5, 31),
    )
    checks = checker._checks_for_entry_mixed(
        entry, log_war_files, sections_by_date, old_cutoff=datetime(2026, 5, 31), new_cutoff=datetime(2026, 6, 1),
    )
    entry_checks = [c for c in checks if c["date"] in (datetime(2026, 4, 30), datetime(2026, 5, 1))]
    other_checks = [c for c in checks if c["date"] not in (datetime(2026, 4, 30), datetime(2026, 5, 1))]
    assert all(c["covered"] is True for c in entry_checks)
    assert all(c["covered"] is None for c in other_checks)
    assert checker._missing_days_count(checks) == 0


# --- _checks_for_entry_mixed_point1 (пункт "1"/30к під "змішаною" -
# "нова" зона й розрив ВЗАГАЛІ НЕ ПЕРЕВІРЯЮТЬСЯ, на відміну від пункту "2") ---
def test_checks_mixed_point1_entire_subrange_before_old_cutoff_behaves_like_old():
    entry = _entry("ПЕРШИЙ Перший Перший", "02.04.2026-15.04.2026", raw_value=30)
    sections_by_date = {
        datetime(2026, 4, 1): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 4, 2): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший;"},
        datetime(2026, 4, 14): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 4, 15): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший;"},
    }
    dates = [datetime(2026, 4, 1), datetime(2026, 4, 2), datetime(2026, 4, 14), datetime(2026, 4, 15)]
    checks = checker._checks_for_entry_mixed_point1(
        entry, _log_war_files(*dates), sections_by_date,
        old_cutoff=datetime(2026, 4, 30), new_cutoff=datetime(2026, 5, 1),
    )
    assert [(c["date"], c["point"]) for c in checks] == [
        (datetime(2026, 4, 1), "5"), (datetime(2026, 4, 2), "6"),
        (datetime(2026, 4, 14), "5"), (datetime(2026, 4, 15), "6"),
    ]


def test_checks_mixed_point1_entire_subrange_after_new_cutoff_is_not_checked_at_all():
    """На відміну від пункту "2" (де "нова" зона перевіряється щодня) - пункт
    "1" у "новій" зоні "змішаної" НЕ перевіряється взагалі (підтверджено
    користувачем: "нові" документи структурно не описують чергування в БРО)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-02.05.2026", raw_value=30)
    sections_by_date = {
        datetime(2026, 5, 1): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 5, 2): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": "ПЕРШИЙ Перший Перший;"},
    }
    checks = checker._checks_for_entry_mixed_point1(
        entry, {}, sections_by_date, old_cutoff=datetime(2026, 4, 8), new_cutoff=datetime(2026, 5, 1),
    )
    assert checks == []


def test_checks_mixed_point1_subrange_spanning_the_gap_checks_entry_only_no_exit():
    """Під-діапазон ПЕРЕТИНАЄ розрив/"нову" зону (05.05-11.05, межі 06/09.05) -
    ВХІД перевіряється за "старою" (04.05(п.5)+05.05(п.6)), а ВИХІД - НІ (не
    штучний "вихід" на межі розриву, і "нова" зона для пункту "1" узагалі не
    перевіряється - на відміну від пункту "2", де вихід ідуть через щоденні
    перевірки "нової" зони)."""
    entry = _entry("ПЕРШИЙ Перший Перший", "05.05.2026-11.05.2026", raw_value=30)
    sections_by_date = {
        datetime(2026, 5, 4): {"section5": "завдання ПЕРШИЙ Перший Перший", "section6_hid": ""},
        datetime(2026, 5, 5): {"section5": "", "section6_hid": "ПЕРШИЙ Перший Перший увійшов;"},
    }
    dates = list(sections_by_date.keys())
    checks = checker._checks_for_entry_mixed_point1(
        entry, _log_war_files(*dates), sections_by_date,
        old_cutoff=datetime(2026, 5, 6), new_cutoff=datetime(2026, 5, 9),
    )
    assert [(c["date"], c["point"]) for c in checks] == [
        (datetime(2026, 5, 4), "5"), (datetime(2026, 5, 5), "6"),
    ]
    assert all(c["covered"] for c in checks)


def test_checks_mixed_point1_dates_strictly_between_cutoffs_are_never_checked():
    entry = _entry("ПЕРШИЙ Перший Перший", "07.05.2026-08.05.2026", raw_value=30)
    checks = checker._checks_for_entry_mixed_point1(
        entry, {}, {}, old_cutoff=datetime(2026, 5, 6), new_cutoff=datetime(2026, 5, 9),
    )
    assert checks == []


def test_checks_mixed_point1_full_month_absence_is_waived_not_an_error():
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-31.05.2026", raw_value=30)
    log_war_files = _log_war_files(
        datetime(2026, 4, 30), datetime(2026, 5, 1), datetime(2026, 5, 30), datetime(2026, 5, 31),
    )
    checks = checker._checks_for_entry_mixed_point1(
        entry, log_war_files, {}, old_cutoff=datetime(2026, 5, 31), new_cutoff=datetime(2026, 6, 1),
    )
    assert len(checks) > 0
    assert all(c["covered"] is None for c in checks)
    assert checker._missing_days_count(checks) == 0


def test_checks_mixed_point1_old_zone_entry_skipped_when_adjacent_day_covered_by_same_person():
    entry = _entry("ПЕРШИЙ Перший Перший", "05.05.2026-06.05.2026", raw_value=30)
    person_dates = {datetime(2026, 5, 4)}
    checks = checker._checks_for_entry_mixed_point1(
        entry, {}, {}, old_cutoff=datetime(2026, 5, 20), new_cutoff=datetime(2026, 5, 21),
        person_dates=person_dates,
    )
    assert (datetime(2026, 5, 4), "5") not in {(c["date"], c["point"]) for c in checks}
    assert (datetime(2026, 5, 5), "6") not in {(c["date"], c["point"]) for c in checks}
    assert (datetime(2026, 5, 5), "5") in {(c["date"], c["point"]) for c in checks}
    assert (datetime(2026, 5, 6), "6") in {(c["date"], c["point"]) for c in checks}


# --- Диспетчер _checks_for_entry ---
def test_checks_dispatch_point1_under_new_logic_returns_no_checks_at_all():
    """Підтверджено користувачем (реальний зразок за липень): "нова" методика
    структурно не описує чергування в БРО - пункт "1" (30к) під "Нова" не
    перевіряється ВЗАГАЛІ (порожній список, а не "стара" методика й не
    "не знайдено")."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026", raw_value=30)
    assert checker._checks_for_entry(entry, {}, {}, checker.LOG_WAR_LOGIC_NEW) == []


def test_checks_dispatch_point1_under_mixed_without_cutoffs_returns_no_checks_at_all():
    """"Змішана" без визначеної межі тихо переходить на "нову" для всього
    рапорту - для пункту "1" це так само означає ПОРОЖНІЙ список."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026", raw_value=30)
    assert checker._checks_for_entry(entry, {}, {}, checker.LOG_WAR_LOGIC_MIXED) == []


def test_checks_dispatch_point1_under_mixed_with_cutoffs_uses_mixed_point1_variant():
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026", raw_value=30)
    result_mixed = checker._checks_for_entry(
        entry, {}, {}, checker.LOG_WAR_LOGIC_MIXED, old_cutoff=datetime(2026, 4, 30), new_cutoff=datetime(2026, 5, 1),
    )
    assert result_mixed == checker._checks_for_entry_mixed_point1(
        entry, {}, {}, old_cutoff=datetime(2026, 4, 30), new_cutoff=datetime(2026, 5, 1),
    )


def test_checks_dispatch_point1_under_old_logic_uses_old():
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026", raw_value=30)
    result_old = checker._checks_for_entry(entry, {}, {}, checker.LOG_WAR_LOGIC_OLD)
    assert result_old == checker._checks_for_entry_old(entry, {}, {})


def test_checks_dispatch_old():
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026")
    result_old = checker._checks_for_entry(entry, {}, {}, checker.LOG_WAR_LOGIC_OLD)
    assert result_old == checker._checks_for_entry_old(entry, {}, {})


def test_checks_dispatch_new():
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026")
    result_new = checker._checks_for_entry(entry, {}, {}, checker.LOG_WAR_LOGIC_NEW)
    assert result_new == checker._checks_for_entry_new(entry, {}, {})


def test_checks_dispatch_mixed_without_detected_cutoffs_falls_back_to_new():
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026")
    result_mixed = checker._checks_for_entry(entry, {}, {}, checker.LOG_WAR_LOGIC_MIXED)
    assert result_mixed == checker._checks_for_entry_new(entry, {}, {})


def test_checks_dispatch_mixed_with_cutoffs():
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026")
    result_mixed = checker._checks_for_entry(
        entry, {}, {}, checker.LOG_WAR_LOGIC_MIXED, old_cutoff=datetime(2026, 4, 30), new_cutoff=datetime(2026, 5, 1),
    )
    assert result_mixed == checker._checks_for_entry_mixed(
        entry, {}, {}, old_cutoff=datetime(2026, 4, 30), new_cutoff=datetime(2026, 5, 1),
    )


# -------------------------
# _detect_style_transition - автоматичне визначення межі переходу зі "старої"
# методики на "нову" (підтверджено користувачем: без потреби вводити чи
# пам'ятати дати) - за "порожніми" (< _GAP_MAX_COMBINED_LENGTH символів) днями.
# -------------------------
def _blank_sections():
    return {"section5": "x" * 50, "section6_hid": "x" * 50}


def _full_sections():
    return {"section5": "x" * 5000, "section6_hid": "x" * 5000}


def test_detect_style_transition_finds_gap_and_returns_boundaries():
    dates = [datetime(2026, 5, d) for d in range(3, 12)]
    sections_by_date = {d: _full_sections() for d in dates}
    sections_by_date[datetime(2026, 5, 7)] = _blank_sections()
    sections_by_date[datetime(2026, 5, 8)] = _blank_sections()
    log_war_files = _log_war_files(*dates)
    result = checker._detect_style_transition(log_war_files, sections_by_date)
    assert result == (datetime(2026, 5, 6), datetime(2026, 5, 9))


def test_detect_style_transition_returns_none_when_no_gap():
    dates = [datetime(2026, 5, d) for d in range(3, 12)]
    sections_by_date = {d: _full_sections() for d in dates}
    assert checker._detect_style_transition(_log_war_files(*dates), sections_by_date) is None


def test_detect_style_transition_treats_missing_sections_entry_as_not_blank():
    """sections_by_date.get(date) відсутній ВЗАГАЛІ (сам ключ дати відсутній
    у словнику, а не просто короткий текст) - _is_blank має трактувати це як
    "не порожній" (False), а не падати чи помилково рахувати розривом."""
    dates = [datetime(2026, 5, d) for d in range(3, 12)]
    sections_by_date = {d: _full_sections() for d in dates if d != datetime(2026, 5, 7)}
    assert checker._detect_style_transition(_log_war_files(*dates), sections_by_date) is None


def test_detect_style_transition_returns_none_when_gap_touches_folder_edge():
    """Розрив упирається в САМ ПОЧАТОК папки - немає дати "до" розриву, щоб
    визначити стару межу."""
    dates = [datetime(2026, 5, d) for d in range(1, 6)]
    sections_by_date = {d: _full_sections() for d in dates}
    sections_by_date[datetime(2026, 5, 1)] = _blank_sections()
    sections_by_date[datetime(2026, 5, 2)] = _blank_sections()
    assert checker._detect_style_transition(_log_war_files(*dates), sections_by_date) is None


def test_detect_style_transition_returns_none_when_too_few_dates():
    dates = [datetime(2026, 5, 1), datetime(2026, 5, 2)]
    sections_by_date = {d: _blank_sections() for d in dates}
    assert checker._detect_style_transition(_log_war_files(*dates), sections_by_date) is None


def test_detect_style_transition_picks_longest_gap_when_multiple():
    dates = [datetime(2026, 5, d) for d in range(1, 15)]
    sections_by_date = {d: _full_sections() for d in dates}
    sections_by_date[datetime(2026, 5, 4)] = _blank_sections()  # короткий розрив (1 день)
    sections_by_date[datetime(2026, 5, 9)] = _blank_sections()  # довший розрив (3 дні)
    sections_by_date[datetime(2026, 5, 10)] = _blank_sections()
    sections_by_date[datetime(2026, 5, 11)] = _blank_sections()
    result = checker._detect_style_transition(_log_war_files(*dates), sections_by_date)
    assert result == (datetime(2026, 5, 8), datetime(2026, 5, 12))


# -------------------------
# _print_style_detection_table - підтверджено користувачем: один рядок-
# підсумок ("стара по X, нова з Y") без доказів "виглядає бредово" - має бути
# видно ПОДЕННІ довжини тексту, з яких вирахувано межу.
# -------------------------
def test_print_style_detection_table_shows_per_day_lengths_and_labels(capsys):
    dates = [datetime(2026, 5, d) for d in range(3, 12)]
    sections_by_date = {d: _full_sections() for d in dates}
    sections_by_date[datetime(2026, 5, 7)] = _blank_sections()
    sections_by_date[datetime(2026, 5, 8)] = _blank_sections()
    log_war_files = _log_war_files(*dates)

    checker._print_style_detection_table(log_war_files, sections_by_date, datetime(2026, 5, 6), datetime(2026, 5, 9))

    out = capsys.readouterr().out
    assert "03.05.2026" in out and "стара методика" in out
    assert "07.05.2026" in out and "порожній день" in out
    assert "08.05.2026" in out and "порожній день" in out
    assert "09.05.2026" in out and "нова методика" in out
    assert "11.05.2026" in out and "нова методика" in out


def test_print_style_detection_table_no_transition_labels_nothing_as_old_or_new(capsys):
    dates = [datetime(2026, 5, d) for d in range(3, 12)]
    sections_by_date = {d: _full_sections() for d in dates}
    checker._print_style_detection_table(_log_war_files(*dates), sections_by_date, None, None)
    out = capsys.readouterr().out
    assert "стара методика" not in out
    assert "нова методика" not in out


# -------------------------
# _majority_year_month / _warn_if_folder_month_mismatch
# -------------------------
def test_majority_year_month_returns_most_common():
    dates = [datetime(2026, 5, 1), datetime(2026, 5, 2), datetime(2026, 6, 1)]
    assert checker._majority_year_month(dates) == (2026, 5)


def test_majority_year_month_empty_returns_none():
    assert checker._majority_year_month([]) is None


def test_result_file_name_includes_month_name():
    """Підтверджено користувачем: назва файлу результату містить НАЗВУ
    місяця перевірки (не лише номер) - напр. "..._Травень.xlsx"."""
    path = checker._result_file_name((2026, 5))
    assert os.path.basename(path) == f"{checker._RESULT_FILE_BASE_NAME}_Травень.xlsx"


def test_result_file_name_without_month_falls_back_to_base_name():
    path = checker._result_file_name(None)
    assert os.path.basename(path) == f"{checker._RESULT_FILE_BASE_NAME}.xlsx"


def test_warn_if_folder_month_mismatch_no_warning_when_matching(capsys):
    checker._warn_if_folder_month_mismatch((2026, 5), [datetime(2026, 5, 1), datetime(2026, 5, 2)])
    assert capsys.readouterr().out == ""


def test_warn_if_folder_month_mismatch_warns_when_different(capsys):
    checker._warn_if_folder_month_mismatch((2026, 5), [datetime(2026, 6, 1), datetime(2026, 6, 2)])
    assert "здебільшого стосується" in capsys.readouterr().out


def test_warn_if_folder_month_mismatch_no_dates_does_not_crash(capsys):
    checker._warn_if_folder_month_mismatch(None, [])
    assert capsys.readouterr().out == ""


# -------------------------
# _point_label - "1 (30к)"/"2 (100к)" (підтверджено користувачем).
# -------------------------
def test_point_label_formats_number_and_amount():
    assert checker._point_label(_entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026", raw_value=30)) == "1 (30к)"


def test_point_label_uses_entrys_own_heading_number():
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026", raw_value=100)
    entry["НОМЕР_ПУНКТУ"] = "2"
    assert checker._point_label(entry) == "2 (100к)"


# -------------------------
# _write_log_war_check_result - "довгий" формат: Посада/Звання/ПІБ/Пункт
# рапорту/Період об'єднані вертикально на всю групу рядків запису.
# -------------------------
def test_write_log_war_check_result_header_row():
    wb_path_headers = checker._IDENTITY_COLUMNS + checker._ROW_COLUMNS
    assert wb_path_headers == ["Посада", "Звання", "ПІБ", "Пункт рапорту", "Період (з рапорту)", "ДАТА ПЕРЕВІРКИ", "Пункт ЖБД", "Файл ЖБД", "Статус"]


def test_write_log_war_check_result_single_covered_row(tmp_path):
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026")
    checks = [{"date": datetime(2026, 5, 1), "point": "6", "file_name": "ЖБД 01.05.2026.docx", "covered": True}]
    output_path = str(tmp_path / "result.xlsx")

    checker._write_log_war_check_result([{"entry": entry, "checks": checks}], output_path)

    wb = load_workbook(output_path)
    ws = wb.active
    assert [c.value for c in ws[1]] == ["Посада", "Звання", "ПІБ", "Пункт рапорту", "Період (з рапорту)", "ДАТА ПЕРЕВІРКИ", "Пункт ЖБД", "Файл ЖБД", "Статус"]
    assert ws.max_row == 2
    assert ws.cell(row=2, column=1).value == "Стрілець"
    assert ws.cell(row=2, column=3).value == "ПЕРШИЙ Перший Перший"
    # "Пункт рапорту" - ЛІТЕРАЛЬНИЙ номер + сума в дужках (_point_label, напр.
    # "2 (100к)") - підтверджено користувачем: сам код raw_value без номера чи
    # "?" у цій колонці був незрозумілий. Номер пункту "2" тут - бо _entry()
    # узгоджує НОМЕР_ПУНКТУ з raw_value за замовчуванням (100 -> "2") - як і в
    # реальних рапортах (пункт "1" завжди 30к, "2" завжди 100к).
    assert ws.cell(row=2, column=4).value == "2 (100к)"
    assert ws.cell(row=2, column=6).value == "01.05.2026"
    assert ws.cell(row=2, column=7).value == "6"
    assert ws.cell(row=2, column=8).value == "ЖБД 01.05.2026.docx"
    assert ws.cell(row=2, column=9).value == "Знайдено"
    assert ws.cell(row=2, column=9).fill.fgColor.rgb == GREEN_FILL.fgColor.rgb


def test_write_log_war_check_result_group_colors_and_borders(tmp_path):
    """Підтверджено користувачем: колонки з рапорту (1-5), "ДАТА ПЕРЕВІРКИ"
    (6) окремим кольором, решта ЖБД-колонок (7-8) - ще одним, "Статус" (9) -
    БЕЗ групової заливки (у нього своя, змістовна) - і в шапці, і в даних.
    Тонка межа - на КОЖНІЙ клітинці табличці."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026")
    checks = [{"date": datetime(2026, 5, 1), "point": "6", "file_name": "ЖБД 01.05.2026.docx", "covered": True}]
    output_path = str(tmp_path / "result.xlsx")

    checker._write_log_war_check_result([{"entry": entry, "checks": checks}], output_path)

    wb = load_workbook(output_path)
    ws = wb.active
    for row in (1, 2):
        for col in range(1, 6):
            assert ws.cell(row=row, column=col).fill.fgColor.rgb == checker._REPORT_DATA_FILL.fgColor.rgb
        assert ws.cell(row=row, column=6).fill.fgColor.rgb == checker._CHECK_DATE_FILL.fgColor.rgb
        for col in (7, 8):
            assert ws.cell(row=row, column=col).fill.fgColor.rgb == checker._LOG_WAR_DATA_FILL.fgColor.rgb

    for row in (1, 2):
        for col in range(1, 10):
            cell = ws.cell(row=row, column=col)
            assert cell.border.left.style == "thin"
            assert cell.border.right.style == "thin"


def test_write_log_war_check_result_waived_row_uses_neutral_status_not_error(tmp_path):
    """covered is None (виняток пункту 30) -> нейтральний статус/заливка, НЕ
    "ВІДСУТНЄ"/RED_FILL і НЕ "Знайдено"/GREEN_FILL."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026", raw_value=30)
    checks = [{"date": datetime(2026, 5, 1), "point": "6", "file_name": "ЖБД 01.05.2026.docx", "covered": None}]
    output_path = str(tmp_path / "result.xlsx")

    checker._write_log_war_check_result([{"entry": entry, "checks": checks}], output_path)

    wb = load_workbook(output_path)
    ws = wb.active
    status_cell = ws.cell(row=2, column=9)
    assert status_cell.value not in ("Знайдено", "ВІДСУТНЄ")
    assert status_cell.fill.fgColor.rgb not in (GREEN_FILL.fgColor.rgb, RED_FILL.fgColor.rgb)


def test_write_log_war_check_result_missing_row_has_bright_red_font_and_fill(tmp_path):
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026")
    checks = [{"date": datetime(2026, 5, 1), "point": "6", "file_name": "ЖБД 01.05.2026.docx", "covered": False}]
    output_path = str(tmp_path / "result.xlsx")

    checker._write_log_war_check_result([{"entry": entry, "checks": checks}], output_path)

    wb = load_workbook(output_path)
    ws = wb.active
    status_cell = ws.cell(row=2, column=9)
    assert status_cell.value == "ВІДСУТНЄ"
    assert status_cell.fill.fgColor.rgb == RED_FILL.fgColor.rgb
    assert status_cell.font.bold is True


def test_write_log_war_check_result_merges_identity_columns_across_multiple_rows(tmp_path):
    """Посада/Звання/ПІБ/Пункт рапорту/Період - ОБ'ЄДНАНІ вертикально на всю
    групу рядків ОДНОГО запису (підтверджено користувачем), коли в нього
    кілька перевірок (напр. перехідний день - "5" і "6")."""
    entry = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026")
    checks = [
        {"date": datetime(2026, 5, 1), "point": "5", "file_name": "ЖБД 01.05.2026.docx", "covered": True},
        {"date": datetime(2026, 5, 1), "point": "6", "file_name": "ЖБД 02.05.2026.docx", "covered": True},
    ]
    output_path = str(tmp_path / "result.xlsx")

    checker._write_log_war_check_result([{"entry": entry, "checks": checks}], output_path)

    wb = load_workbook(output_path)
    ws = wb.active
    assert ws.max_row == 3  # header + 2 check rows
    merged_ranges = {str(r) for r in ws.merged_cells.ranges}
    # Кожна з 5 ідентифікаційних колонок (A-E) об'єднана з рядка 2 по рядок 3.
    for col_letter in ["A", "B", "C", "D", "E"]:
        assert f"{col_letter}2:{col_letter}3" in merged_ranges
    # Дата/Пункт/Файл/Статус (F-I) НЕ об'єднані - кожен рядок свій.
    for col_letter in ["F", "G", "H", "I"]:
        assert f"{col_letter}2:{col_letter}3" not in merged_ranges


def test_write_log_war_check_result_multiple_entries_each_get_own_merge_block(tmp_path):
    entry1 = _entry("ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026")
    entry2 = _entry("ДРУГИЙ Другий Другий", "01.05.2026-01.05.2026")
    checks1 = [
        {"date": datetime(2026, 5, 1), "point": "5", "file_name": "ЖБД 01.05.2026.docx", "covered": True},
        {"date": datetime(2026, 5, 1), "point": "6", "file_name": "ЖБД 02.05.2026.docx", "covered": True},
    ]
    checks2 = [{"date": datetime(2026, 5, 1), "point": "6", "file_name": "ЖБД 01.05.2026.docx", "covered": False}]
    output_path = str(tmp_path / "result.xlsx")

    checker._write_log_war_check_result(
        [{"entry": entry1, "checks": checks1}, {"entry": entry2, "checks": checks2}], output_path,
    )

    wb = load_workbook(output_path)
    ws = wb.active
    assert ws.max_row == 4  # header + 2 rows (entry1) + 1 row (entry2)
    assert ws.cell(row=2, column=3).value == "ПЕРШИЙ Перший Перший"
    assert ws.cell(row=4, column=3).value == "ДРУГИЙ Другий Другий"


# -------------------------
# check_report_against_log_war - наскрізь
# -------------------------
def test_check_report_against_log_war_only_covers_headings_numbered_1_and_2(tmp_path):
    """Підтверджено користувачем: перевірка стосується ЛИШЕ пунктів,
    ЛІТЕРАЛЬНО пронумерованих "1" і "2" у самому рапорті - пункт 170к (тут -
    "3") НЕ фігурує у файлі результату взагалі, навіть якщо людина явно
    відсутня в ЖБД."""
    report_path = _write_report(tmp_path / "report.docx", [
        (
            "1. Виплатити додаткову винагороду у розмірі 30 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "15.05.2026-15.05.2026", "1", ""]],
        ),
        (
            "2. Виплатити додаткову винагороду у розмірі 100 000 грн. 00 коп. військовослужбовцям",
            [["Оператор", "старший матрос", "ДРУГИЙ Другий Другий", "15.05.2026-15.05.2026", "1", ""]],
        ),
        (
            "3. Виплатити додаткову винагороду у розмірі 170 000 грн. 00 коп. військовослужбовцям",
            [["Кулеметник", "молодший сержант", "ТРЕТІЙ Третій Третій", "15.05.2026-15.05.2026", "1", ""]],
        ),
    ])

    log_war_dir = tmp_path / "log_war"
    log_war_dir.mkdir()
    # 15.05 (СЕРЕДИНА місяця, НЕ на межі) - щоб не зачепити
    # _waive_boundary_absence (тут перевіряємо саме ОХОПЛЕННЯ пунктів).
    _write_log_war_doc(log_war_dir / "ЖБД 15.05.2026.docx", [
        "5. Рішення командира.",
        "Хтось інший.",
        "6. Відомості про виконання.",
        "Хід виконання спланованих завдань:",
        # Пункт "1" (30к) під "новою" методикою ВЗАГАЛІ НЕ ПЕРЕВІРЯЄТЬСЯ
        # (_POINT1_HEADING_NUMBER) - тож тут явно обираємо "Стара" (_checks_for_entry_old),
        # де для одноденного періоду 15.05-15.05 єдина перевірка "в межах" - це
        # пункт "6" САМЕ 15.05.2026 (дня "до" входу, 14.05.2026, у папці взагалі
        # немає файлу - ця частина природно "прощається", бо скрипт не знає
        # точної дати входу). Фраза "з батальйонного району оборони" тут - щоб
        # відсутність ПЕРШОГО НЕ була "прощена" винятком пункту 30
        # (_covered_value) - тут перевіряємо саме ОХОПЛЕННЯ пунктів, а не цей
        # окремий виняток.
        "Хтось інший з батальйонного району оборони.",
    ])

    output_path = str(tmp_path / "result.xlsx")
    checker.check_report_against_log_war(
        report_path, str(log_war_dir), output_path=output_path, logic_mode=checker.LOG_WAR_LOGIC_OLD,
    )

    wb = load_workbook(output_path)
    ws = wb.active
    pib_values = {v for v in (ws.cell(row=r, column=3).value for r in range(2, ws.max_row + 1)) if v is not None}
    assert pib_values == {"ПЕРШИЙ Перший Перший", "ДРУГИЙ Другий Другий"}
    assert "ТРЕТІЙ Третій Третій" not in pib_values


def test_check_report_against_log_war_ignores_30k_100k_under_a_different_heading_number(tmp_path):
    """РЕГРЕСІЯ (виявлено користувачем): пункт "5", що ТЕЖ платить 30к
    (напр. ретроактивна поправка за минулий місяць) - НЕ пункт "1"/"2" -
    ігнорується перевіркою, навіть якщо сума ті самі 30к і людина явно
    відсутня в ЖБД. Раніше фільтр був за СУМОЮ (raw_value), а не за
    ЛІТЕРАЛЬНИМ номером пункту, тож такий запис помилково потрапляв у файл."""
    report_path = _write_report(tmp_path / "report.docx", [
        (
            "5. Виплатити додаткову винагороду у розмірі 30 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ЧЕТВЕРТИЙ Четвертий Четвертий", "01.05.2026-01.05.2026", "1", ""]],
        ),
    ])

    log_war_dir = tmp_path / "log_war"
    log_war_dir.mkdir()
    _write_log_war_doc(log_war_dir / "ЖБД 01.05.2026.docx", [
        "5. Рішення командира.",
        "Нічого особливого.",
        "6. Відомості про виконання.",
        "Хід виконання спланованих завдань:",
        "Нічого не згадано.",
    ])

    output_path = str(tmp_path / "result.xlsx")
    checker.check_report_against_log_war(report_path, str(log_war_dir), output_path=output_path)

    wb = load_workbook(output_path)
    ws = wb.active
    assert ws.max_row == 1  # заголовок лише - пункт "5" не перевіряється взагалі


def test_check_report_against_log_war_end_to_end_fully_covered_produces_header_only(tmp_path):
    """Підтверджено користувачем: "Знайдено"/"ОК" рядки НЕ потрапляють у файл
    узагалі - людина, повністю покрита ЖБД, у файлі геть не з'являється
    (лишається лише заголовок). "Стара" методика (logic_mode) - найпростіший
    сценарій для перевірки САМЕ фільтрації файлу (вхід: 30.04(п.5)+01.05(п.6))."""
    report_path = _write_report(tmp_path / "report.docx", [
        (
            "1. Виплатити додаткову винагороду у розмірі 100 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026", "1", ""]],
        ),
    ])

    log_war_dir = tmp_path / "log_war"
    log_war_dir.mkdir()
    _write_log_war_doc(log_war_dir / "ЖБД 30.04.2026.docx", [
        "5. Рішення командира військової частини та бойові завдання.",
        "Завдання: ПЕРШИЙ Перший Перший - зайняти позицію.",
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 01.05.2026.docx", [
        "6. Відомості про виконання.",
        "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])

    output_path = str(tmp_path / "result.xlsx")
    result_path = checker.check_report_against_log_war(
        report_path, str(log_war_dir), output_path=output_path, logic_mode=checker.LOG_WAR_LOGIC_OLD,
    )

    assert result_path == output_path
    wb = load_workbook(output_path)
    ws = wb.active
    assert ws.max_row == 1  # лише заголовок - нічого виправляти не треба


def test_check_report_against_log_war_finds_missing_coverage(tmp_path):
    report_path = _write_report(tmp_path / "report.docx", [
        (
            "2. Виплатити додаткову винагороду у розмірі 100 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026", "1", ""]],
        ),
    ])

    log_war_dir = tmp_path / "log_war"
    log_war_dir.mkdir()
    _write_log_war_doc(log_war_dir / "ЖБД 01.05.2026.docx", [
        "5. Рішення командира.",
        "Нічого про цю людину.",
        "6. Відомості про виконання.",
        "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ДРУГИЙ Другий Другий;",
    ])

    output_path = str(tmp_path / "result.xlsx")
    checker.check_report_against_log_war(report_path, str(log_war_dir), output_path=output_path)

    wb = load_workbook(output_path)
    ws = wb.active
    # Перевіряється ЛИШЕ присутність у пункті "6" самого 01.05 - ПЕРШИЙ там не
    # згаданий (згаданий лише ДРУГИЙ) - ОДИН рядок помилки.
    statuses = [ws.cell(row=r, column=9).value for r in range(2, ws.max_row + 1)]
    assert statuses == ["ВІДСУТНЄ"]


def test_check_report_against_log_war_new_logic_multiday_stay_checks_only_the_exit(tmp_path):
    """Наскрізь: РЕГРЕСІЯ (виявлено користувачем, реальні зразки за липень -
    ДЯКОВИЧ, МИРОНЕНКО, ТІЩЕНКО): період з ДВОДЕННИМ безперервним
    перебуванням ("01.05.2026-02.05.2026") - "Нова" методика раніше хибно
    перевіряла і ПЕРШИЙ день (01.05,"5")+(02.05,"6"), і ще й пару на межі
    виходу ((02.05,"5")+(03.05,"6")), хоча реальні документи підтверджують
    ЛИШЕ присутність людини в пункті "6" ОСТАННЬОГО дня перебування (02.05,
    знайдено одразу - fallback на 03.05 тут не потрібен, див. окремий тест
    test_check_report_against_log_war_new_logic_falls_back_to_next_day_when_exit_day_not_mentioned)."""
    report_path = _write_report(tmp_path / "report.docx", [
        (
            "2. Виплатити додаткову винагороду у розмірі 100 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "01.05.2026-02.05.2026", "1", ""]],
        ),
    ])

    log_war_dir = tmp_path / "log_war"
    log_war_dir.mkdir()
    # 01.05 - НАВМИСНЕ без згадки ПЕРШИЙ у жодній секції (перший день
    # безперервного перебування - реально НЕ підтверджується окремо).
    _write_log_war_doc(log_war_dir / "ЖБД 01.05.2026.docx", [
        "5. Рішення командира.",
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
    ])
    # 02.05 (останній день перебування) - ПЕРШИЙ згаданий САМЕ в пункті "6"
    # ЦЬОГО Ж дня - єдине, що перевіряється.
    _write_log_war_doc(log_war_dir / "ЖБД 02.05.2026.docx", [
        "5. Рішення командира.",
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])

    output_path = str(tmp_path / "result.xlsx")
    checker.check_report_against_log_war(report_path, str(log_war_dir), output_path=output_path)

    wb = load_workbook(output_path)
    ws = wb.active
    assert ws.max_row == 1  # усе покрито - "01.05" узагалі не мав би перевірятись


def test_check_report_against_log_war_new_logic_falls_back_to_next_day_when_exit_day_not_mentioned(tmp_path):
    """Наскрізь: РЕГРЕСІЯ (виявлено користувачем, реальний зразок - САСЮК):
    "Хід виконання" за ОСТАННІЙ день перебування (16.07) підтверджено в
    файлі НАСТУПНОГО дня (17.07 - звичайна конвенція), а НЕ ЦЬОГО ж дня -
    перевірка має спробувати ОБИДВА дні, перш ніж вважати це помилкою."""
    report_path = _write_report(tmp_path / "report.docx", [
        (
            "2. Виплатити додаткову винагороду у розмірі 100 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "15.07.2026-16.07.2026", "1", ""]],
        ),
    ])

    log_war_dir = tmp_path / "log_war"
    log_war_dir.mkdir()
    # 16.07 (останній день) - НАВМИСНЕ без згадки ПЕРШИЙ у пункті "6".
    _write_log_war_doc(log_war_dir / "ЖБД 16.07.2026.docx", [
        "5. Рішення командира.",
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
    ])
    # 17.07 - ПЕРШИЙ уже у відрядженні, але "Хід виконання" тут підтверджує
    # ЙОГО ЖЕ виконання завдання ВЧОРА (16.07).
    _write_log_war_doc(log_war_dir / "ЖБД 17.07.2026.docx", [
        "5. Рішення командира.",
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])

    output_path = str(tmp_path / "result.xlsx")
    checker.check_report_against_log_war(report_path, str(log_war_dir), output_path=output_path)

    wb = load_workbook(output_path)
    ws = wb.active
    assert ws.max_row == 1  # усе покрито - знайдено на 17.07 через fallback


def test_check_report_against_log_war_fully_covered_entry_is_omitted_when_mixed_with_a_broken_one(tmp_path):
    """Кілька записів рапорту: ОДИН повністю покритий (не з'являється у файлі
    взагалі), ДРУГИЙ - ні (з'являється, лише з реальними помилками). "Стара"
    методика - найпростіший сценарій для перевірки САМЕ фільтрації файлу."""
    report_path = _write_report(tmp_path / "report.docx", [
        (
            "1. Виплатити додаткову винагороду у розмірі 100 000 грн. 00 коп. військовослужбовцям",
            [
                ["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "15.05.2026-15.05.2026", "1", ""],
                ["Оператор", "старший матрос", "ДРУГИЙ Другий Другий", "15.05.2026-15.05.2026", "1", ""],
            ],
        ),
    ])

    log_war_dir = tmp_path / "log_war"
    log_war_dir.mkdir()
    # 15.05 (СЕРЕДИНА місяця, НЕ на межі) - щоб не зачепити
    # _waive_boundary_absence (тут перевіряємо саме фільтрацію файлу).
    _write_log_war_doc(log_war_dir / "ЖБД 14.05.2026.docx", [
        "5. Рішення командира.",
        "Завдання: ПЕРШИЙ Перший Перший - зайняти позицію.",
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 15.05.2026.docx", [
        "6. Відомості про виконання.",
        "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])

    output_path = str(tmp_path / "result.xlsx")
    checker.check_report_against_log_war(
        report_path, str(log_war_dir), output_path=output_path, logic_mode=checker.LOG_WAR_LOGIC_OLD,
    )

    wb = load_workbook(output_path)
    ws = wb.active
    # ПЕРШИЙ - повністю покритий, не з'являється взагалі. ДРУГИЙ - жодних
    # згадок узагалі - вхід (14.05(п.5)+15.05(п.6)) не підтверджений - 2
    # рядки (PIB - лише в ПЕРШОМУ з них, решта об'єднаної колонки - None, як
    # у решті тестів merge).
    pib_values = [v for v in (ws.cell(row=r, column=3).value for r in range(2, ws.max_row + 1)) if v is not None]
    assert pib_values == ["ДРУГИЙ Другий Другий"]
    assert ws.max_row == 3


def test_check_report_against_log_war_mixed_end_to_end_auto_detects_gap_and_applies_new_logic(tmp_path):
    """Наскрізь: "змішана" методика САМА визначає межу переходу (03-06.05 -
    "повні" дні, 07-08.05 - "порожній" розрив, 09.05+ - знову "повні") і
    застосовує "нову" методику для дати, що потрапляє в "нову" зону -
    підтверджено користувачем: без потреби вводити чи пам'ятати дати."""
    # Пункт "2" (100к) - у зоні "нова" методика дійсно застосовується (пункт
    # "1"/30к ЗАВЖДИ "стара", незалежно від logic_mode -
    # _ALWAYS_OLD_LOGIC_HEADING_NUMBER - цей тест саме про "нову" зону).
    report_path = _write_report(tmp_path / "report.docx", [
        (
            "2. Виплатити додаткову винагороду у розмірі 100 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "10.05.2026-10.05.2026", "1", ""]],
        ),
    ])

    log_war_dir = tmp_path / "log_war"
    log_war_dir.mkdir()
    for day in (3, 4, 5, 6, 9):
        _write_log_war_doc(log_war_dir / f"ЖБД {day:02d}.05.2026.docx", [
            "5. Рішення командира.",
            "x" * 500,
            "6. Відомості про виконання.",
            "Хід виконання спланованих завдань:",
            "x" * 500,
        ])
    for day in (7, 8):
        _write_log_war_doc(log_war_dir / f"ЖБД {day:02d}.05.2026.docx", [
            "5. Рішення командира.",
            "6. Відомості про виконання.",
            "Хід виконання спланованих завдань:",
        ])
    # 10.05 (єдиний день періоду, у "новій" зоні) - ПЕРШИЙ згаданий САМЕ в
    # пункті "6" ЦЬОГО Ж дня - єдине, що перевіряється (ВИПРАВЛЕНО ВДРУГЕ -
    # перевірка НАСТУПНОГО дня, 11.05, більше не виконується).
    _write_log_war_doc(log_war_dir / "ЖБД 10.05.2026.docx", [
        "5. Рішення командира.",
        "x" * 500,
        "6. Відомості про виконання.",
        "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])

    output_path = str(tmp_path / "result.xlsx")
    checker.check_report_against_log_war(
        report_path, str(log_war_dir), output_path=output_path, logic_mode=checker.LOG_WAR_LOGIC_MIXED,
    )

    wb = load_workbook(output_path)
    ws = wb.active
    assert ws.max_row == 1  # повністю покритий за автоматично визначеною "новою" зоною


def test_check_report_against_log_war_mixed_end_to_end_falls_back_to_new_with_warning_when_no_gap(tmp_path, capsys):
    """Немає жодного "порожнього" дня в папці - межу переходу визначити
    неможливо - друкується попередження (червоним), і застосовується "нова"
    методика для всього рапорту (підтверджено користувачем - безпечний типовий
    варіант замість зупинки)."""
    report_path = _write_report(tmp_path / "report.docx", [
        (
            "1. Виплатити додаткову винагороду у розмірі 100 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "01.05.2026-01.05.2026", "1", ""]],
        ),
    ])

    log_war_dir = tmp_path / "log_war"
    log_war_dir.mkdir()
    for day in (1, 2, 3):
        _write_log_war_doc(log_war_dir / f"ЖБД {day:02d}.05.2026.docx", [
            "5. Рішення командира.",
            "x" * 500,
            "6. Відомості про виконання.",
            "Хід виконання спланованих завдань:",
            "x" * 500,
        ])

    output_path = str(tmp_path / "result.xlsx")
    checker.check_report_against_log_war(
        report_path, str(log_war_dir), output_path=output_path, logic_mode=checker.LOG_WAR_LOGIC_MIXED,
    )

    assert "не вдалось автоматично визначити межу переходу" in capsys.readouterr().out


def test_check_report_against_log_war_old_logic_entry_before_loaded_folder_is_not_an_error(tmp_path):
    """Наскрізь: приклад користувача - в/сл зайшов у БРО 03.04.2026 (задовго
    ДО завантаженої папки ЖБД - лише за травень), вийшов 16.05.2026 (У
    травні, папка ЖБД це покриває). ВХІД (02.04"5"+03.04"6") НЕ з'являється
    як помилка взагалі (файлів за квітень немає в обраній папці) - ВИХІД
    (15.05"5"+16.05"6") перевіряється ЗВИЧАЙНО, і тут - справжня помилка
    (ПІБ ніде не згаданий), тож ЦЕЙ рядок з'являється."""
    report_path = _write_report(tmp_path / "report.docx", [
        (
            "2. Виплатити додаткову винагороду у розмірі 100 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "03.04.2026-16.05.2026", "1", ""]],
        ),
    ])

    log_war_dir = tmp_path / "log_war"
    log_war_dir.mkdir()
    # Лише травневі файли - жодного квітневого (навіть "перехлюпу" 30.04).
    for day in (15, 16):
        _write_log_war_doc(log_war_dir / f"ЖБД {day:02d}.05.2026.docx", [
            "5. Рішення командира.",
            "6. Відомості про виконання.",
            "Хід виконання спланованих завдань:",
            "Хтось інший.",
        ])

    output_path = str(tmp_path / "result.xlsx")
    checker.check_report_against_log_war(
        report_path, str(log_war_dir), output_path=output_path, logic_mode=checker.LOG_WAR_LOGIC_OLD,
    )

    wb = load_workbook(output_path)
    ws = wb.active
    # Лише ВИХІД (справжня помилка) - 2 рядки (15.05"5" + 16.05"6"), вхід
    # (02.04/03.04) узагалі не з'являється - файлів за квітень немає.
    dates_shown = {ws.cell(row=r, column=6).value for r in range(2, ws.max_row + 1)}
    assert dates_shown == {"15.05.2026", "16.05.2026"}
    assert all(ws.cell(row=r, column=9).value == "ВІДСУТНЄ" for r in range(2, ws.max_row + 1))


def test_check_report_against_log_war_old_logic_no_false_error_on_seamless_category_transition(tmp_path):
    """Наскрізь: РЕГРЕСІЯ, виявлена користувачем на реальних даних - людина
    БЕЗ РОЗРИВУ переходить з пункту 100к (до 04.05) у пункт 30к (від 05.05).
    Реальне переміщення задокументоване РІВНО ОДИН раз (30.04(п.5)+01.05(п.6) -
    вхід у 100к; 07.05(п.5)+08.05(п.6) - вихід із 30к, наприкінці) - дні
    02.05-06.05 не потребують ЖОДНОЇ окремої згадки, бо ні "вихід із 100к",
    ні "вхід у 30к" на межі 04.05/05.05 НЕ перевіряються окремо (безшовне
    продовження) - весь запис має вийти ПОВНІСТЮ покритим (лише заголовок)."""
    report_path = _write_report(tmp_path / "report.docx", [
        (
            "2. Виплатити додаткову винагороду у розмірі 100 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "01.05.2026-04.05.2026", "4", ""]],
        ),
        (
            "1. Виплатити додаткову винагороду у розмірі 30 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "05.05.2026-08.05.2026", "4", ""]],
        ),
    ])

    log_war_dir = tmp_path / "log_war"
    log_war_dir.mkdir()
    _write_log_war_doc(log_war_dir / "ЖБД 30.04.2026.docx", [
        "5. Рішення командира.",
        "Завдання: ПЕРШИЙ Перший Перший - зайняти позицію.",
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 01.05.2026.docx", [
        "6. Відомості про виконання.",
        "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 07.05.2026.docx", [
        "5. Рішення командира.",
        "Завдання: ПЕРШИЙ Перший Перший - залишити позицію.",
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 08.05.2026.docx", [
        "6. Відомості про виконання.",
        "Хід виконання спланованих завдань:",
        "Вибув особовий склад: ПЕРШИЙ Перший Перший;",
    ])

    output_path = str(tmp_path / "result.xlsx")
    checker.check_report_against_log_war(
        report_path, str(log_war_dir), output_path=output_path, logic_mode=checker.LOG_WAR_LOGIC_OLD,
    )

    wb = load_workbook(output_path)
    ws = wb.active
    assert ws.max_row == 1  # лише заголовок - жодної хибної помилки на межі 04.05/05.05


def test_check_report_against_log_war_mixed_no_false_error_on_seamless_transition_in_new_zone(tmp_path):
    """Наскрізь: РЕГРЕСІЯ, виявлена користувачем на реальних даних (МАКОГІН) -
    людина БЕЗ РОЗРИВУ переходить з пункту 30к (09.05-12.05) у пункт 100к
    (13.05-16.05), обидва під-діапазони починаються в "новій" зоні (межа
    07/09.05, визначена автоматично). Пункт 30к у "новій" зоні взагалі НЕ
    перевіряється (_checks_for_entry_mixed_point1 - підтверджено окремо
    користувачем за липень), а пункт 100к перевіряє ЛИШЕ свій ВЛАСНИЙ вихід
    (16.05(п.5)+17.05(п.6), під-діапазон ЦІЛКОМ у "новій" зоні) - жодна з
    двох перевірок не зачіпає межу 12.05/13.05 узагалі, тож "друге" завдання
    там просто нізвідки взятись."""
    report_path = _write_report(tmp_path / "report.docx", [
        (
            "1. Виплатити додаткову винагороду у розмірі 30 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "09.05.2026-12.05.2026", "4", ""]],
        ),
        (
            "2. Виплатити додаткову винагороду у розмірі 100 000 грн. 00 коп. військовослужбовцям",
            [["Стрілець", "сержант", "ПЕРШИЙ Перший Перший", "13.05.2026-16.05.2026", "4", ""]],
        ),
    ])

    log_war_dir = tmp_path / "log_war"
    log_war_dir.mkdir()
    for day in (3, 4, 5, 6):
        _write_log_war_doc(log_war_dir / f"ЖБД {day:02d}.05.2026.docx", [
            "5. Рішення командира.", "x" * 500,
            "6. Відомості про виконання.", "Хід виконання спланованих завдань:", "x" * 500,
        ])
    for day in (7, 8):
        _write_log_war_doc(log_war_dir / f"ЖБД {day:02d}.05.2026.docx", [
            "5. Рішення командира.",
            "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
        ])
    # 09-12.05 (пункт 30к, "новозонний" щоденний цикл: свіже завдання щодня,
    # підтверджене наступного дня) - 13.05 навмисне БЕЗ згадки в п.5.
    _write_log_war_doc(log_war_dir / "ЖБД 09.05.2026.docx", [
        "5. Рішення командира.", "Завдання: ПЕРШИЙ Перший Перший - лишатись на позиції.",
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:", "x" * 500,
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 10.05.2026.docx", [
        "5. Рішення командира.", "Завдання: ПЕРШИЙ Перший Перший - лишатись на позиції.",
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 11.05.2026.docx", [
        "5. Рішення командира.", "Завдання: ПЕРШИЙ Перший Перший - лишатись на позиції.",
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 12.05.2026.docx", [
        "5. Рішення командира.", "Завдання: ПЕРШИЙ Перший Перший - залишити позицію.",
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 13.05.2026.docx", [
        "5. Рішення командира.",  # НАВМИСНЕ без згадки ПЕРШИЙ - продовження, не нове завдання.
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
        "Вибув особовий склад: ПЕРШИЙ Перший Перший;",
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 14.05.2026.docx", [
        "5. Рішення командира.", "Завдання: ПЕРШИЙ Перший Перший - нове завдання.",
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 15.05.2026.docx", [
        "5. Рішення командира.", "Завдання: ПЕРШИЙ Перший Перший - нове завдання.",
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 16.05.2026.docx", [
        "5. Рішення командира.", "Завдання: ПЕРШИЙ Перший Перший - нове завдання.",
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])
    _write_log_war_doc(log_war_dir / "ЖБД 17.05.2026.docx", [
        "5. Рішення командира.",
        "6. Відомості про виконання.", "Хід виконання спланованих завдань:",
        "Залучено особовий склад: ПЕРШИЙ Перший Перший;",
    ])

    output_path = str(tmp_path / "result.xlsx")
    checker.check_report_against_log_war(
        report_path, str(log_war_dir), output_path=output_path, logic_mode=checker.LOG_WAR_LOGIC_MIXED,
    )

    wb = load_workbook(output_path)
    ws = wb.active
    assert ws.max_row == 1  # лише заголовок - жодної хибної помилки на межі 12.05/13.05
