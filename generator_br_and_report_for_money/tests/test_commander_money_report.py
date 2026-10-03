import os
import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import constants
from helpers_for_TEST_WORK import read_docx_text


def _row(pib, posada, days=None, zvannya="сержант"):
    row = {"ПІБ": pib, "ПОСАДА": posada, "ЗВАННЯ": zvannya}
    row.update(days or {})
    return row


def _tvo_row(pib, posada, start, end, tvo=True):
    return {"ПІБ": pib, "ПОСАДА": posada, "Start": start, "End": end, "ТВО": tvo}


# -------------------------
# _find_regular_commander_row / _collect_commander_rows
# -------------------------
def test_find_regular_commander_row_matches_exact_position():
    import generators.generate_report_for_commander_money as cm

    rows = [_row("ПЕРШИЙ Перший", "Стрілець"), _row("ДРУГИЙ Другий", constants.HIGHER_COMMANDER_TITLE)]
    result = cm._find_regular_commander_row(rows)
    assert result["ПІБ"] == "ДРУГИЙ Другий"


def test_find_regular_commander_row_returns_none_when_absent():
    import generators.generate_report_for_commander_money as cm

    rows = [_row("ПЕРШИЙ Перший", "Стрілець")]
    assert cm._find_regular_commander_row(rows) is None


def test_collect_commander_rows_puts_regular_first_then_tvo_each_scoped_to_own_period():
    """Штатний командир рахується лише за дні, НЕ покриті ТВО (тут - 01.07), а
    ТВО - лише за дні своєї заявленої заміни (тут - 15.07 і 31.07, Start=15.07) -
    ЛИШЕ коли підписує ТВО (narrator_row з позначкою ТВО)."""
    import generators.generate_report_for_commander_money as cm

    d1, d_mid, d2 = datetime(2026, 7, 1), datetime(2026, 7, 15), datetime(2026, 7, 31)
    regular = _row("ДРУГИЙ Другий", constants.HIGHER_COMMANDER_TITLE, {d1: 100, d_mid: 100, d2: 100})
    deputy_tvo = _row("ТРЕТІЙ Третій", "Заступник командира батальйону", {d_mid: 100, d2: 100})
    rows_with_data = [deputy_tvo, regular]
    rows_with_tvo_data = [_tvo_row("ТРЕТІЙ Третій", constants.HIGHER_COMMANDER_TITLE, "15.07.2026", "31.07.2026")]
    narrator_row = {"ПІБ": "ТРЕТІЙ Третій", "ТВО": True}

    result = cm._collect_commander_rows(rows_with_data, rows_with_tvo_data, [d1, d_mid, d2], d1, d2, narrator_row)

    assert [r["ПІБ"] for r in result] == ["ДРУГИЙ Другий", "ТРЕТІЙ Третій"]
    regular_result, tvo_result = result
    assert set(regular_result) & {d1, d_mid, d2} == {d1}
    assert set(tvo_result) & {d1, d_mid, d2} == {d_mid, d2}


def test_collect_commander_rows_scopes_each_tvo_to_its_own_window_when_multiple():
    """Двоє різних ТВО протягом місяця (кожен - свій окремий період без перетину) -
    кожен рядок обмежений ЛИШЕ днями своєї заміни, "по періодам" - підписує другий ТВО."""
    import generators.generate_report_for_commander_money as cm

    d1, d_mid, d2 = datetime(2026, 7, 1), datetime(2026, 7, 15), datetime(2026, 7, 31)
    tvo_first = _row("ПЕРШИЙ ТВО", "Заступник командира батальйону", {d1: 100, d_mid: 100})
    tvo_second = _row("ДРУГИЙ ТВО", "Начальник штабу батальйону", {d_mid: 100, d2: 100})
    rows_with_data = [tvo_first, tvo_second]
    rows_with_tvo_data = [
        _tvo_row("ПЕРШИЙ ТВО", constants.HIGHER_COMMANDER_TITLE, "01.07.2026", "14.07.2026"),
        _tvo_row("ДРУГИЙ ТВО", constants.HIGHER_COMMANDER_TITLE, "15.07.2026", "31.07.2026"),
    ]
    narrator_row = {"ПІБ": "ДРУГИЙ ТВО", "ТВО": True}

    result = cm._collect_commander_rows(rows_with_data, rows_with_tvo_data, [d1, d_mid, d2], d1, d2, narrator_row)

    by_pib = {r["ПІБ"]: r for r in result}
    assert set(by_pib["ПЕРШИЙ ТВО"]) & {d1, d_mid, d2} == {d1}
    assert set(by_pib["ДРУГИЙ ТВО"]) & {d1, d_mid, d2} == {d_mid, d2}


def test_collect_commander_rows_merges_dates_if_regular_was_also_tvo():
    """Якщо штатний командир одночасно (рідкісний збіг) значиться ТВО на цю ж
    посаду - обидва діапазони об'єднуються в ОДИН його рядок, без дублювання."""
    import generators.generate_report_for_commander_money as cm

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 31)
    regular = _row("ДРУГИЙ Другий", constants.HIGHER_COMMANDER_TITLE, {d1: 100, d2: 100})
    rows_with_tvo_data = [_tvo_row("ДРУГИЙ Другий", constants.HIGHER_COMMANDER_TITLE, "01.07.2026", "31.07.2026")]
    narrator_row = {"ПІБ": "ДРУГИЙ Другий", "ТВО": True}

    result = cm._collect_commander_rows([regular], rows_with_tvo_data, [d1, d2], d1, d2, narrator_row)

    assert [r["ПІБ"] for r in result] == ["ДРУГИЙ Другий"]
    assert set(result[0]) & {d1, d2} == {d1, d2}


def test_collect_commander_rows_empty_when_no_regular_and_no_tvo():
    import generators.generate_report_for_commander_money as cm

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 31)
    rows = [_row("ПЕРШИЙ Перший", "Стрілець", {d1: 100})]
    result = cm._collect_commander_rows(rows, [], [d1, d2], d1, d2, {})
    assert result == []


def test_collect_commander_rows_skips_tvo_person_missing_from_personnel_data():
    """ТВО-запис (аркуш "ТВО") існує, але відповідного ПІБ немає серед rows_with_data
    (напр. людину вже звільнено/переведено) - такий ТВО просто пропускається, без падіння."""
    import generators.generate_report_for_commander_money as cm

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 31)
    rows_with_tvo_data = [_tvo_row("ЗНИКЛИЙ ТВО", constants.HIGHER_COMMANDER_TITLE, "01.07.2026", "31.07.2026")]
    narrator_row = {"ПІБ": "ЗНИКЛИЙ ТВО", "ТВО": True}

    result = cm._collect_commander_rows([], rows_with_tvo_data, [d1, d2], d1, d2, narrator_row)

    assert result == []


# -------------------------
# НОВЕ ПРАВИЛО (підтверджено користувачем): хто підписує (narrator_row) визначає,
# чи додавати ТВО окремим рядком узагалі - а не сам факт, що ТВО був частину місяця.
# -------------------------
def test_collect_commander_rows_regular_signer_excludes_tvo_entirely():
    """Штатний командир підписує (narrator_row БЕЗ позначки ТВО - наразі немає
    активного на "сьогодні" ТВО) - ТВО НЕ додається взагалі, навіть якщо він
    реально був ТВО частину місяця; штатний отримує ОДИН рядок за ВЕСЬ період."""
    import generators.generate_report_for_commander_money as cm

    d1, d_mid, d2 = datetime(2026, 8, 1), datetime(2026, 8, 15), datetime(2026, 8, 31)
    regular = _row("ЧЕТВЕРТИЙ Четвертий", constants.HIGHER_COMMANDER_TITLE, {d1: 100, d_mid: 100, d2: 100})
    former_tvo = _row("ШОСТИЙ Шостий", "Заступник командира батальйону", {d1: 100, d_mid: 100})
    rows_with_data = [former_tvo, regular]
    rows_with_tvo_data = [_tvo_row("ШОСТИЙ Шостий", constants.HIGHER_COMMANDER_TITLE, "01.08.2026", "14.08.2026")]
    narrator_row = {"ПІБ": "ЧЕТВЕРТИЙ Четвертий", "ПОСАДА": constants.HIGHER_COMMANDER_TITLE}

    result = cm._collect_commander_rows(rows_with_data, rows_with_tvo_data, [d1, d_mid, d2], d1, d2, narrator_row)

    assert [r["ПІБ"] for r in result] == ["ЧЕТВЕРТИЙ Четвертий"]
    assert set(result[0]) & {d1, d_mid, d2} == {d1, d_mid, d2}


def test_collect_commander_rows_tvo_signer_includes_both_tvo_and_regular():
    """ТВО підписує (narrator_row З позначкою ТВО) - і ТВО, і штатний командир
    потрапляють РАЗОМ, кожен за СВІЙ період (як і раніше, коли ТВО активний)."""
    import generators.generate_report_for_commander_money as cm

    d1, d_mid, d2 = datetime(2026, 8, 1), datetime(2026, 8, 15), datetime(2026, 8, 31)
    regular = _row("ЧЕТВЕРТИЙ Четвертий", constants.HIGHER_COMMANDER_TITLE, {d1: 100, d_mid: 100, d2: 100})
    active_tvo = _row("ШОСТИЙ Шостий", "Заступник командира батальйону", {d1: 100, d_mid: 100})
    rows_with_data = [active_tvo, regular]
    rows_with_tvo_data = [_tvo_row("ШОСТИЙ Шостий", constants.HIGHER_COMMANDER_TITLE, "01.08.2026", "14.08.2026")]
    narrator_row = {"ПІБ": "ШОСТИЙ Шостий", "ТВО": True}

    result = cm._collect_commander_rows(rows_with_data, rows_with_tvo_data, [d1, d_mid, d2], d1, d2, narrator_row)

    by_pib = {r["ПІБ"]: r for r in result}
    assert set(by_pib.keys()) == {"ЧЕТВЕРТИЙ Четвертий", "ШОСТИЙ Шостий"}
    assert set(by_pib["ШОСТИЙ Шостий"]) & {d1, d_mid, d2} == {d1}
    assert set(by_pib["ЧЕТВЕРТИЙ Четвертий"]) & {d1, d_mid, d2} == {d_mid, d2}


def test_collect_commander_rows_regular_signer_with_no_regular_row_is_empty():
    """Штатний командир підписує (за find_higher_commander), але його рядка
    немає серед rows_with_data (напр. посада вакантна на папері) - порожній
    результат, а не падіння."""
    import generators.generate_report_for_commander_money as cm

    d1, d2 = datetime(2026, 8, 1), datetime(2026, 8, 31)
    rows_with_tvo_data = [_tvo_row("КОЛИШНІЙ ТВО", constants.HIGHER_COMMANDER_TITLE, "01.08.2026", "14.08.2026")]

    result = cm._collect_commander_rows([], rows_with_tvo_data, [d1, d2], d1, d2, {})

    assert result == []


# -------------------------
# _add_commander_intro
# -------------------------
def test_add_commander_intro_substitutes_nominative_rank_pib_posada():
    """rank/pib/posada беруться напряму з narrator_row (ОБЛІК.xlsx, називний відмінок) -
    без окремого файлу з відмінками ПІБ/посади."""
    import generators.generate_report_for_commander_money as cm
    from docx import Document

    doc = Document()
    narrator_row = {"ЗВАННЯ": "сержант", "ПІБ": "ДЕСЯТИЙ Десятий Десятий", "ПОСАДА": "Командир батальйону"}

    cm._add_commander_intro(doc, 7, narrator_row)

    text = "\n".join(p.text for p in doc.paragraphs)
    assert "сержант" in text
    assert "ДЕСЯТИЙ Десятий Десятий" in text
    assert "Командир батальйону" in text
    assert "липня" in text


# -------------------------
# process_generate_report_for_commander_money - наскрізна генерація (реальні resources/*)
# -------------------------
def test_process_generate_report_for_commander_money_returns_none_without_dates():
    from generators.generate_report_for_commander_money import process_generate_report_for_commander_money

    result = process_generate_report_for_commander_money([], [], ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ"])
    assert result is None


def test_process_generate_report_for_commander_money_returns_none_without_commander_or_tvo(monkeypatch):
    import generators.generate_report_for_commander_money as cm

    d1 = datetime(2026, 7, 1)
    rows = [_row("ПЕРШИЙ Перший", "Стрілець", {d1: 100})]
    result = cm.process_generate_report_for_commander_money(rows, [], [d1])
    assert result is None


def test_process_generate_report_for_commander_money_returns_none_without_matching_categories():
    """Штатний командир є, але жоден його день місяця не потрапляє в жодну
    категорію (усі комірки дат порожні) - рапорт КБ/ТВО не формується."""
    import generators.generate_report_for_commander_money as cm

    d1 = datetime(2026, 7, 1)
    rows = [_row("ВОСЬМИЙ Восьмий", constants.HIGHER_COMMANDER_TITLE, {d1: None})]
    result = cm.process_generate_report_for_commander_money(rows, [], [d1])
    assert result is None


def test_process_generate_report_for_commander_money_uses_commander_specific_categories(tmp_path, monkeypatch):
    """Рапорт на командира/ТВО будує підстави з COMMANDER_MONEY_REPORT_CATEGORIES,
    а НЕ з головного MONEY_REPORT_CATEGORIES (тут навіть не монкіпатчиться - лишається
    реальним) - підміна однієї не впливає на іншу."""
    import generators.generate_report_for_commander_money as cm
    from helpers_for_TEST_WORK import read_docx_text

    monkeypatch.setattr(cm, "OUTPUT_DIR", str(tmp_path))
    # _build_categories завжди звертається до всіх 4 пунктів (30/70/100/170) - тож
    # навіть коли тест цікавить лише 100, довелось додати мінімальну структуру й для
    # решти (лише "general" + catch-all "БД(СЗ)", без власних підстав).
    _minimal_point = {"general": [], "БД(СЗ)": {"grounds": [], "use_brs": False, "exclude_general": [], "default": True}}
    monkeypatch.setattr(cm, "COMMANDER_MONEY_REPORT_CATEGORIES", {
        30: {"general": [], "ЖИТТЄДІЯЛЬНІСТЬ": {"grounds": [], "use_brs": False, "exclude_general": [], "default": True}, "ЗВРез": {"grounds": [], "use_brs": False, "exclude_general": []}},
        70: _minimal_point,
        100: {
            "general": [],
            "РТГр": {"grounds": [], "use_brs": False, "exclude_general": []},
            "БД(СЗ)": {"grounds": [{"start": "01.07.2026", "end": "31.07.2026", "lines": ["ОКРЕМА ПІДСТАВА КОМАНДИРА"]}], "use_brs": False, "exclude_general": [], "default": True},
            "МЕДИК": {"grounds": [], "use_brs": False, "exclude_general": []},
        },
        170: _minimal_point,
    })

    d1 = datetime(2026, 7, 1)
    rows = [_row("СЬОМИЙ Сьомий Сьомий", constants.HIGHER_COMMANDER_TITLE, {d1: 100}, zvannya="сержант")]

    result_path = cm.process_generate_report_for_commander_money(rows, [], [d1])

    assert result_path is not None
    text = read_docx_text(result_path)
    assert "ОКРЕМА ПІДСТАВА КОМАНДИРА" in text
    assert "СЬОМИЙ Сьомий Сьомий" in text


def test_process_generate_report_for_commander_money_includes_not_paid_section_for_szch(tmp_path, monkeypatch):
    """_NOT_PAID_POINT - тепер псевдо-пункт прямо в COMMANDER_MONEY_REPORT_CATEGORIES
    (реальний словник, не монкіпатчиться) - рапорт КБ/ТВО теж рендерить пункт
    "Не виплачувати..." в кінці, якщо сам штатний командир мав "СЗЧ" хоч один
    день місяця, з підставою з колонки "ПІДСТАВИ" його рядка ОБЛІК.xlsx."""
    import generators.generate_report_for_commander_money as cm
    from helpers_for_TEST_WORK import read_docx_text

    monkeypatch.setattr(cm, "OUTPUT_DIR", str(tmp_path))

    d1, d2 = datetime(2026, 7, 1), datetime(2026, 7, 2)
    rows = [_row("СЬОМИЙ Сьомий Сьомий", constants.HIGHER_COMMANDER_TITLE, {d1: 100, d2: "СЗЧ", "ПІДСТАВИ": "Довідка №1 від 03.07.2026"})]

    result_path = cm.process_generate_report_for_commander_money(rows, [], [d1, d2])

    assert result_path is not None
    text = read_docx_text(result_path)
    assert cm.STATIK["COMMANDER_SECTIONS"]["NOT_PAID"] in text
    assert "Довідка №1 від 03.07.2026" in text


def test_process_generate_report_for_commander_money_uses_nominative_data_without_extra_file(tmp_path, monkeypatch):
    """Раніше давальний відмінок ПІБ/посади бралися з окремого опційного файлу
    (ЗАДАЧІ.xlsm) - якщо для того, хто підписує, там не було запису, рапорт КБ/ТВО
    взагалі не формувався. Тепер rank/pib/posada беруться напряму з ОБЛІК.xlsx
    (називний відмінок, той самий rows_with_data) - жодного окремого файлу не треба."""
    import generators.generate_report_for_commander_money as cm
    from helpers_for_TEST_WORK import read_docx_text

    monkeypatch.setattr(cm, "OUTPUT_DIR", str(tmp_path))

    d1 = datetime(2026, 7, 1)
    rows = [_row("СЬОМИЙ Сьомий Сьомий", constants.HIGHER_COMMANDER_TITLE, {d1: 100})]

    result_path = cm.process_generate_report_for_commander_money(rows, [], [d1])

    assert result_path is not None
    text = read_docx_text(result_path)
    assert "виплатити мені" in text
    assert "СЬОМИЙ Сьомий Сьомий" in text


def test_process_generate_report_for_commander_money_end_to_end(tmp_path, monkeypatch):
    import generators.generate_report_for_commander_money as cm
    from process_generate_report_for_get_additional_money import process_generate_report_for_commander_money
    from utils.excel_reader import read_datafile, read_optional_datafile, PERSONEL_LIST_FALLBACK_SHEET_NAMES

    monkeypatch.setattr(cm, "OUTPUT_DIR", str(tmp_path))

    rows_data = read_datafile(
        constants.PERSONEL_LIST_FILE_NAME, constants.PERSONEL_LIST_SHEET_NAME, constants.PERSONEL_LIST_COLUMNS_LETTERS,
        extra_fallback_sheet_names=PERSONEL_LIST_FALLBACK_SHEET_NAMES,
    )
    tvo_rows = read_optional_datafile(constants.PERSONEL_LIST_FILE_NAME, constants.TVO_LIST_SHEET_NAME, constants.TVO_LIST_COLUMNS_LETTERS, sheet_can_be_missing=True)

    result_path = process_generate_report_for_commander_money(
        rows_data.get("rows", []), tvo_rows, rows_data.get("columns", []),
    )

    assert result_path is not None
    assert os.path.isfile(result_path)
    assert os.path.dirname(os.path.abspath(result_path)) == str(tmp_path)

    text = read_docx_text(result_path)
    assert "РАПОРТ" in text
    assert "прошу виплатити мені" in text
