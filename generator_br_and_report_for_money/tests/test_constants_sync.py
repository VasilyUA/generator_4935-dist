import os
import sys
import zipfile
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from docx import Document

import sync.constants_sync as constants_sync


def _write_docx(path, paragraphs):
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(str(path))
    return str(path)


def _zip_up(zip_path, *file_paths):
    with zipfile.ZipFile(zip_path, "w") as archive:
        for file_path in file_paths:
            archive.write(file_path, os.path.basename(file_path))
    return str(zip_path)


# -------------------------
# _derive_general_br_bat
# -------------------------
def test_derive_general_br_bat_uses_latest_week_entry_not_later_than_day():
    week_dict = {"01.09.2026": {"бат": "166"}, "07.09.2026": {"бат": "186"}}
    result = constants_sync._derive_general_br_bat(["01.09.2026", "05.09.2026", "07.09.2026"], week_dict)
    assert result["01.09.2026"] == "166 від 01.09.2026"
    assert result["05.09.2026"] == "166 від 01.09.2026"  # між двома тижневими - лишається останнє видане
    assert result["07.09.2026"] == "186 від 07.09.2026"


def test_derive_general_br_bat_skips_invalid_day_key_instead_of_crashing(capsys):
    """РЕГРЕСІЯ (реальний випадок користувача, вересень 2026): constants.py
    (людина редагує вручну щомісяця, копіюючи структуру попереднього місяця)
    лишив "зайвий" ключ "31.09.2026" зі скопійованого 31-денного місяця -
    такого календарного дня в вересні не існує. Раніше ЦЕ обривало весь
    sync_document_numbers_from_folder (ValueError) і губило вже успішно
    відскановану папку з документами - тепер лише ЦЕЙ ключ пропускається."""
    week_dict = {"01.09.2026": {"бат": "166"}}
    result = constants_sync._derive_general_br_bat(["01.09.2026", "31.09.2026"], week_dict)

    assert result["01.09.2026"] == "166 від 01.09.2026"
    assert "31.09.2026" not in result
    assert "Некоректна дата-ключ" in capsys.readouterr().out


def test_derive_general_br_bat_skips_invalid_week_key_instead_of_crashing(capsys):
    """Той самий випадок, але сама "зайва" дата - у NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK
    (не в EVERY_DAY, як у test_derive_general_br_bat_skips_invalid_day_key_instead_of_crashing) -
    решта тижневих записів і далі мають рахуватись, а не обривати все."""
    week_dict = {"01.09.2026": {"бат": "166"}, "31.09.2026": {"бат": "241"}}
    result = constants_sync._derive_general_br_bat(["01.09.2026", "05.09.2026"], week_dict)

    assert result["01.09.2026"] == "166 від 01.09.2026"
    assert result["05.09.2026"] == "166 від 01.09.2026"
    assert "Некоректна дата-ключ" in capsys.readouterr().out


# -------------------------
# _own_document_date / _own_document_number
# -------------------------
def test_own_document_date_finds_first_date_in_header():
    text = "Розпорядження. Дата видання 07.08.2026. Далі текст з іншою датою 31.12.2099."
    date = constants_sync._own_document_date(text)
    assert (date.day, date.month, date.year) == (7, 8, 2026)


def test_own_document_date_handles_two_digit_year():
    text = "Дата 07.08.26 щось."
    date = constants_sync._own_document_date(text)
    assert (date.day, date.month, date.year) == (7, 8, 2026)


def test_own_document_date_none_when_absent():
    assert constants_sync._own_document_date("Текст без жодної дати такого формату.") is None


def test_own_document_date_ignores_dates_beyond_header_window():
    padding = "x" * constants_sync._HEADER_SEARCH_WINDOW
    text = padding + " 07.08.2026"
    assert constants_sync._own_document_date(text) is None


def test_own_document_number_skips_short_section_label():
    text = "Розпорядження. Пункт № 1 Затвердити. Розпорядження № 2726 від 07.08.2026."
    assert constants_sync._own_document_number(text) == "2726"


def test_own_document_number_none_when_absent():
    assert constants_sync._own_document_number("Текст без жодного номера.") is None


# -------------------------
# _carry_forward_field
# -------------------------
def test_carry_forward_field_uses_found_values_directly():
    number_dict = {"01.09.2026": {"брг": ""}, "02.09.2026": {"брг": ""}}
    constants_sync._carry_forward_field(
        number_dict, ["01.09.2026", "02.09.2026"], "брг",
        {"01.09.2026": "100", "02.09.2026": "101"},
    )
    assert number_dict["01.09.2026"]["брг"] == "100"
    assert number_dict["02.09.2026"]["брг"] == "101"


def test_carry_forward_field_repeats_last_found_value_for_gap():
    number_dict = {"01.09.2026": {"брг": ""}, "02.09.2026": {"брг": ""}, "03.09.2026": {"брг": ""}}
    gaps = constants_sync._carry_forward_field(
        number_dict, ["01.09.2026", "02.09.2026", "03.09.2026"], "брг",
        {"01.09.2026": "100", "03.09.2026": "102"},
    )
    assert number_dict["01.09.2026"]["брг"] == "100"
    assert number_dict["02.09.2026"]["брг"] == "100"  # прогалина - повторює 01.09
    assert number_dict["03.09.2026"]["брг"] == "102"
    # Повертає (дата, дата-джерело) для КОЖНОЇ дати, заповненої повторенням -
    # підтверджено користувачем: реальний випадок (29.08.2026, документа
    # немає) без цього виглядав як "правильно опрацьовано", хоча насправді
    # значення - лише повторення сусіднього дня, не підтверджене документом.
    assert gaps == [("02.09.2026", "01.09.2026")]


def test_carry_forward_field_stays_empty_when_nothing_found_and_nothing_preexisting():
    number_dict = {"01.09.2026": {"брг": ""}, "02.09.2026": {"брг": ""}}
    constants_sync._carry_forward_field(
        number_dict, ["01.09.2026", "02.09.2026"], "брг", {},
    )
    assert number_dict["01.09.2026"]["брг"] == ""  # незмінне - нічого не знайдено й нема звідки повторити
    assert number_dict["02.09.2026"]["брг"] == ""  # теж незмінне - нема з чого повторити


def test_carry_forward_field_existing_values_never_leak_into_unrelated_keys():
    """РЕГРЕСІЯ (виявлено користувачем, реальні дані серпня): раніше ІСНУЮЧЕ
    (до синхронізації) значення 1-го ключа, для якого нічого не знайдено,
    хибно "просочувалось" як last_known і перезаписувало ВЖЕ ПРАВИЛЬНЕ,
    АЛЕ ТЕЖ НЕЗНАЙДЕНЕ, значення 2-го ключа - обидва мають лишитись
    НЕЗМІННИМИ (кожен - своїм власним значенням), коли скан НІЧОГО не знайшов
    для жодного з них."""
    number_dict = {"01.09.2026": {"брг": "999"}, "02.09.2026": {"брг": "888"}}
    constants_sync._carry_forward_field(
        number_dict, ["01.09.2026", "02.09.2026"], "брг", {},
    )
    assert number_dict["01.09.2026"]["брг"] == "999"  # незмінне
    assert number_dict["02.09.2026"]["брг"] == "888"  # НЕЗМІННЕ - не перезаписане чужим значенням


def test_carry_forward_field_leading_gap_before_first_found_value_stays_unchanged():
    """РЕГРЕСІЯ: 1-ше число місяця, для якого нічого не знайдено (реальне
    джерело - документ, датований ОСТАННІМ днем ПОПЕРЕДНЬОГО місяця, якого в
    обраній папці найчастіше просто немає) - лишається НЕЗМІННИМ, а не
    отримує номер, "запозичений" звідкись іще ПЕРЕД ним у списку ключів."""
    number_dict = {"31.08.2026": {"брг": "2665"}, "01.09.2026": {"брг": "2682"}, "02.09.2026": {"брг": ""}}
    constants_sync._carry_forward_field(
        number_dict, ["31.08.2026", "01.09.2026", "02.09.2026"], "брг",
        {"02.09.2026": "2726"},
    )
    assert number_dict["31.08.2026"]["брг"] == "2665"  # незмінне
    assert number_dict["01.09.2026"]["брг"] == "2682"  # незмінне - НЕ перезаписане значенням 31.08
    assert number_dict["02.09.2026"]["брг"] == "2726"  # знайдено напряму


def test_carry_forward_field_trailing_gap_after_last_found_value_stays_unchanged():
    """РЕГРЕСІЯ (виявлено користувачем, реальні дані серпня): ОСТАННІЙ день
    місяця, для якого документа зазвичай іще просто немає, - лишається
    НЕЗМІННИМ (тут - порожнім), а НЕ отримує номер попереднього дня
    (екстраполяція В НЕВІДОМЕ - на відміну від прогалини МІЖ двома
    знахідками, яку й далі треба заповнювати повторенням)."""
    number_dict = {"29.08.2026": {"брг": ""}, "30.08.2026": {"брг": ""}, "31.08.2026": {"брг": ""}}
    constants_sync._carry_forward_field(
        number_dict, ["29.08.2026", "30.08.2026", "31.08.2026"], "брг",
        {"30.08.2026": "3267"},
    )
    assert number_dict["29.08.2026"]["брг"] == ""  # немає ЗНАЙДЕНОГО значення ПІЗНІШЕ 29-го теж не рахується (29 - НЕ прогалина, а сам не знайдений)
    assert number_dict["30.08.2026"]["брг"] == "3267"  # знайдено напряму
    assert number_dict["31.08.2026"]["брг"] == ""  # незмінне - НЕМАЄ знахідки ПІСЛЯ 31-го, екстраполяція не відбувається


# -------------------------
# sync_brigade_document_numbers_from_folder (наскрізь, persist=False)
# -------------------------
def test_sync_brigade_document_numbers_assigns_to_date_plus_one_day(tmp_path):
    """Підтверджено користувачем, звірено з реальними даними: документ,
    датований 07.08.2026, стосується ДНЯ 08.08.2026 - фізична папка, куди файл
    поклали, для розрахунку дати не використовується взагалі."""
    doc_path = _write_docx(tmp_path / "doc.docx", [
        "Розпорядження № 2726 від 07.08.2026.",
        "Наказую продовжити обороняти батальйонний район оборони.",
    ])
    zip_path = _zip_up(tmp_path / "any_folder_name.zip", doc_path)
    folder = tmp_path / "src"
    folder.mkdir()
    os.replace(zip_path, folder / "any_folder_name.zip")

    day_dict = {"07.08.2026": {"брг": ""}, "08.08.2026": {"брг": ""}}
    day_keys = ["07.08.2026", "08.08.2026"]
    save_dict, save_keys = {}, []

    constants_sync.sync_brigade_document_numbers_from_folder(
        str(folder), "unused.py", day_dict, day_keys, save_dict, save_keys, persist=False,
    )

    assert day_dict["08.08.2026"]["брг"] == "2726"
    assert day_dict["07.08.2026"]["брг"] == ""  # нічого не знайдено САМЕ за цей день


def test_sync_brigade_document_numbers_recognizes_alternate_defense_wording(tmp_path):
    """Реальні дані серпня (документ №3071 від 19.08.2026) довели: та сама
    команда "продовжити оборону" в деяких документах сформульована ІНАКШЕ, ніж
    _PHRASE_CONTINUE_DEFENSE - без розпізнавання цього варіанту документ мовчки
    ігнорувався, хоча за всіма іншими ознаками (номер, дата) був справжнім
    бойовим розпорядженням, що продовжує дію попереднього."""
    doc_path = _write_docx(tmp_path / "doc.docx", [
        "Розпорядження № 3071 від 19.08.2026.",
        "Наказую продовжити ведення оборонного бою у складі І ешелону оборони в батальйонному районі оборони.",
    ])
    zip_path = _zip_up(tmp_path / "any_folder_name.zip", doc_path)
    folder = tmp_path / "src"
    folder.mkdir()
    os.replace(zip_path, folder / "any_folder_name.zip")

    day_dict = {"19.08.2026": {"брг": ""}, "20.08.2026": {"брг": ""}}
    day_keys = ["19.08.2026", "20.08.2026"]

    constants_sync.sync_brigade_document_numbers_from_folder(
        str(folder), "unused.py", day_dict, day_keys, {}, [], persist=False,
    )

    assert day_dict["20.08.2026"]["брг"] == "3071"


def test_sync_brigade_document_numbers_overwrites_stale_value_with_freshly_found_one(tmp_path):
    """Реальний випадок (серпень, документ №3193 від 25.08.2026 у
    resources/серпень/25/7825дск.zip): 26.08.2026 у constants.py вже МАВ якесь
    (застаріле/помилкове) значення "3174" - користувач очікує, що знайдений
    документ його ПЕРЕЗАПИШЕ (26.08 = 25.08+1 день), а НЕ лишить старе."""
    doc_path = _write_docx(tmp_path / "doc.docx", [
        "Розпорядження № 3193 від 25.08.2026.",
        "Наказую продовжити обороняти батальйонний район оборони.",
    ])
    inner_zip = _zip_up(tmp_path / "01_7825дск.zip", doc_path)
    folder = tmp_path / "src"
    folder.mkdir()
    os.replace(_zip_up(tmp_path / "outer.zip", inner_zip), folder / "7825дск.zip")

    day_dict = {"25.08.2026": {"брг": "3174"}, "26.08.2026": {"брг": "3174"}, "27.08.2026": {"брг": ""}}
    day_keys = ["25.08.2026", "26.08.2026", "27.08.2026"]

    constants_sync.sync_brigade_document_numbers_from_folder(
        str(folder), "unused.py", day_dict, day_keys, {}, [], persist=False,
    )

    assert day_dict["26.08.2026"]["брг"] == "3193"
    assert day_dict["25.08.2026"]["брг"] == "3174"  # незмінне - нічого не знайдено САМЕ за 25.08


def test_sync_brigade_document_numbers_warns_about_gap_filled_by_repetition(tmp_path, capsys):
    """Реальний випадок (29.08.2026 - жодного документа немає у наданій
    папці) - людина має ОДРАЗУ побачити в консолі, що конкретна дата НЕ
    підтверджена документом (значення - лише повторення сусіднього дня), а
    не з'ясовувати це вручну діфом constants.py, як сталось раніше."""
    doc1 = _write_docx(tmp_path / "doc1.docx", [
        "Розпорядження № 500 від 01.09.2026.",
        "Наказую продовжити обороняти батальйонний район оборони.",
    ])
    doc2 = _write_docx(tmp_path / "doc2.docx", [
        "Розпорядження № 510 від 03.09.2026.",
        "Наказую продовжити обороняти батальйонний район оборони.",
    ])
    folder = tmp_path / "src"
    folder.mkdir()
    os.replace(_zip_up(tmp_path / "a.zip", doc1), folder / "a.zip")
    os.replace(_zip_up(tmp_path / "b.zip", doc2), folder / "b.zip")

    day_dict = {"02.09.2026": {"брг": ""}, "03.09.2026": {"брг": ""}, "04.09.2026": {"брг": ""}}
    day_keys = ["02.09.2026", "03.09.2026", "04.09.2026"]

    constants_sync.sync_brigade_document_numbers_from_folder(
        str(folder), "unused.py", day_dict, day_keys, {}, [], persist=False,
    )

    assert day_dict["03.09.2026"]["брг"] == "500"  # прогалина - повторює 02.09
    out = capsys.readouterr().out
    assert "03.09.2026" in out and "документ НЕ знайдено" in out
    assert "з 02.09.2026" in out


def test_sync_brigade_document_numbers_handles_safety_phrase_and_nested_zip(tmp_path):
    """Реальні зразки: документ надходить у zip, ЩЕ РАЗ запакованому у zip -
    сканування має розпаковувати ВКЛАДЕНІ архіви теж."""
    doc_path = _write_docx(tmp_path / "doc.docx", [
        "Розпорядження № 1036 від 01.09.2026.",
        "Про невиконання заходів безпеки застосування військ або їх зрив доповідати командиру.",
    ])
    inner_zip = _zip_up(tmp_path / "inner.zip", doc_path)
    outer_zip = _zip_up(tmp_path / "outer.zip", inner_zip)
    folder = tmp_path / "src"
    folder.mkdir()
    os.replace(outer_zip, folder / "outer.zip")

    day_dict, day_keys = {}, []
    save_dict = {"02.09.2026": {"бз_брг": ""}}
    save_keys = ["02.09.2026"]

    constants_sync.sync_brigade_document_numbers_from_folder(
        str(folder), "unused.py", day_dict, day_keys, save_dict, save_keys, persist=False,
    )

    assert save_dict["02.09.2026"]["бз_брг"] == "1036"


def test_sync_brigade_document_numbers_carries_forward_across_gap(tmp_path):
    doc1 = _write_docx(tmp_path / "doc1.docx", [
        "Розпорядження № 500 від 01.09.2026.",
        "Наказую продовжити обороняти батальйонний район оборони.",
    ])
    doc2 = _write_docx(tmp_path / "doc2.docx", [
        "Розпорядження № 510 від 03.09.2026.",
        "Наказую продовжити обороняти батальйонний район оборони.",
    ])
    folder = tmp_path / "src"
    folder.mkdir()
    os.replace(_zip_up(tmp_path / "a.zip", doc1), folder / "a.zip")
    os.replace(_zip_up(tmp_path / "b.zip", doc2), folder / "b.zip")

    day_dict = {"02.09.2026": {"брг": ""}, "03.09.2026": {"брг": ""}, "04.09.2026": {"брг": ""}}
    day_keys = ["02.09.2026", "03.09.2026", "04.09.2026"]

    constants_sync.sync_brigade_document_numbers_from_folder(
        str(folder), "unused.py", day_dict, day_keys, {}, [], persist=False,
    )

    assert day_dict["02.09.2026"]["брг"] == "500"
    assert day_dict["03.09.2026"]["брг"] == "500"  # прогалина - повторює 02.09
    assert day_dict["04.09.2026"]["брг"] == "510"


def test_sync_brigade_document_numbers_warns_on_conflicting_duplicate_dates(tmp_path, capsys):
    doc1 = _write_docx(tmp_path / "doc1.docx", [
        "Розпорядження № 500 від 01.09.2026.",
        "Наказую продовжити обороняти батальйонний район оборони.",
    ])
    doc2 = _write_docx(tmp_path / "doc2.docx", [
        "Розпорядження № 600 від 01.09.2026.",
        "Наказую продовжити обороняти батальйонний район оборони.",
    ] + ["Додатковий абзац - щоб файл цього документа гарантовано був більшим."] * 40)
    folder = tmp_path / "src"
    folder.mkdir()
    os.replace(_zip_up(tmp_path / "a.zip", doc1), folder / "a.zip")
    os.replace(_zip_up(tmp_path / "b.zip", doc2), folder / "b.zip")

    day_dict = {"02.09.2026": {"брг": ""}}
    day_keys = ["02.09.2026"]

    constants_sync.sync_brigade_document_numbers_from_folder(
        str(folder), "unused.py", day_dict, day_keys, {}, [], persist=False,
    )

    assert day_dict["02.09.2026"]["брг"] == "600"  # лишається НАЙБІЛЬШИЙ за розміром файлу
    out = capsys.readouterr().out
    assert "КІЛЬКА документів з РІЗНИМИ номерами" in out
    assert "залишено 600" in out and "відхилено: 500" in out


def test_sync_brigade_document_numbers_largest_file_wins_regardless_of_zip_order(tmp_path):
    """Результат не залежить від того, який архів трапився першим у папці (os.walk
    порядок на різних машинах різний) - обидва варіанти порядку мають дати той самий номер."""
    short_doc = _write_docx(tmp_path / "short.docx", [
        "Розпорядження № 500 від 01.09.2026.",
        "Наказую продовжити обороняти батальйонний район оборони.",
    ])
    long_doc = _write_docx(tmp_path / "long.docx", [
        "Розпорядження № 600 від 01.09.2026.",
        "Наказую продовжити обороняти батальйонний район оборони.",
    ] + ["Додатковий абзац - щоб файл цього документа гарантовано був більшим."] * 40)

    for case, (long_zip, short_zip) in enumerate((("a.zip", "b.zip"), ("b.zip", "a.zip"))):
        folder = tmp_path / f"src_{case}"
        folder.mkdir()
        os.replace(_zip_up(tmp_path / f"long_{case}.zip", long_doc), folder / long_zip)
        os.replace(_zip_up(tmp_path / f"short_{case}.zip", short_doc), folder / short_zip)

        day_dict = {"02.09.2026": {"брг": ""}}
        constants_sync.sync_brigade_document_numbers_from_folder(
            str(folder), "unused.py", day_dict, ["02.09.2026"], {}, [], persist=False,
        )

        assert day_dict["02.09.2026"]["брг"] == "600"


def test_pick_largest_candidates_keeps_largest_file_and_ranks_rejected_ones():
    candidates = {"06.09.2026": {"1061": 2000, "1065": 9000, "1060": 2000}}

    found, duplicates = constants_sync._pick_largest_candidates(candidates)

    assert found == {"06.09.2026": "1065"}
    # однаковий розмір (1060, 1061) - менший номер першим, щоб результат не залежав від порядку файлів
    assert duplicates["06.09.2026"] == [("1065", 9000), ("1060", 2000), ("1061", 2000)]


def test_pick_largest_candidates_single_candidate_is_not_a_duplicate():
    found, duplicates = constants_sync._pick_largest_candidates({"06.09.2026": {"1061": 2000}})

    assert found == {"06.09.2026": "1061"}
    assert duplicates == {}


def test_remember_candidate_keeps_largest_size_per_number():
    candidates = {}
    constants_sync._remember_candidate(candidates, "06.09.2026", "1061", 2000)
    constants_sync._remember_candidate(candidates, "06.09.2026", "1061", 500)  # та сама копія меншого розміру

    assert candidates == {"06.09.2026": {"1061": 2000}}


def test_sync_brigade_document_numbers_no_matches_prints_warning(tmp_path, capsys):
    doc_path = _write_docx(tmp_path / "doc.docx", ["Звичайний документ без потрібних фраз."])
    folder = tmp_path / "src"
    folder.mkdir()
    os.replace(_zip_up(tmp_path / "a.zip", doc_path), folder / "a.zip")

    constants_sync.sync_brigade_document_numbers_from_folder(
        str(folder), "unused.py", {}, [], {}, [], persist=False,
    )

    assert "не знайдено жодного документа" in capsys.readouterr().out


def test_sync_brigade_document_numbers_prints_progress_percent(tmp_path, capsys):
    doc_path = _write_docx(tmp_path / "doc.docx", [
        "Розпорядження № 500 від 01.09.2026.",
        "Наказую продовжити обороняти батальйонний район оборони.",
    ])
    folder = tmp_path / "src"
    folder.mkdir()
    os.replace(_zip_up(tmp_path / "a.zip", doc_path), folder / "a.zip")

    constants_sync.sync_brigade_document_numbers_from_folder(
        str(folder), "unused.py", {}, [], {}, [], persist=False,
    )

    out = capsys.readouterr().out
    assert "Сканування документів БРГ" in out
    assert "100%" in out


# -------------------------
# _iter_document_texts_from_zip - самопосилання / запобіжник глибини
# -------------------------
def test_iter_document_texts_from_zip_self_reference_stops_immediately_not_hang(tmp_path, capsys):
    """Реальні дані (архів "5_736_461дск.zip", серпень) - .zip, що містить
    ВКЛАДЕНИЙ .zip із ТИМ САМИМ іменем, що й у нього самого (найімовірніше,
    помилка пакування) - раніше це марно розпаковувалось знову й знову аж до
    _MAX_ZIP_NESTING_DEPTH; тепер має зупинитись одразу, ОДНИМ попередженням
    про самопосилання (і, найголовніше, взагалі не зависнути/не впасти)."""
    doc_path = _write_docx(tmp_path / "doc.docx", ["Звичайний документ."])
    inner_build_path = _zip_up(tmp_path / "inner_build.zip", doc_path)

    outer_path = tmp_path / "same_name.zip"
    with zipfile.ZipFile(outer_path, "w") as archive:
        # Вкладений запис МАЄ те саме ім'я, що й сам зовнішній архів.
        archive.write(inner_build_path, "same_name.zip")

    texts = list(constants_sync._iter_document_texts_from_zip(str(outer_path), str(tmp_path / "tmp_root")))

    assert texts == []  # самопосилання пропущено - жодного тексту НЕ витягнуто з нього
    assert "самопосилання" in capsys.readouterr().out


def test_iter_document_texts_from_zip_depth_guard_still_works_for_distinct_names(tmp_path, capsys):
    """Запасний запобіжник (_MAX_ZIP_NESTING_DEPTH) і далі має спрацьовувати,
    навіть коли самого повторення ІМЕНІ немає (напр. гіпотетичний ланцюжок із
    РІЗНИМИ іменами на кожному рівні) - перевіряємо виклик НАПРЯМУ з великим
    depth, а не будуючи справді 10+ рівнів вкладеності."""
    doc_path = _write_docx(tmp_path / "doc.docx", ["Звичайний документ."])
    zip_path = _zip_up(tmp_path / "any.zip", doc_path)

    texts = list(constants_sync._iter_document_texts_from_zip(
        zip_path, str(tmp_path / "tmp_root"), depth=constants_sync._MAX_ZIP_NESTING_DEPTH + 1,
    ))

    assert texts == []
    assert "Забагато вкладених .zip" in capsys.readouterr().out


# -------------------------
# sync_brigade_document_numbers_from_folder - persist=True записує напряму
# (як і БАТ/БЗ-синхронізація - без окремого підтвердження)
# -------------------------
def _constants_file_with_brigade_block(tmp_path, day_lines, save_lines=("",)):
    constants_path = tmp_path / "constants.py"
    body = "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY = {\n" + "".join(day_lines) + "}\n"
    body += "NUMBER_OF_DOCUMENTS_BRS_SAVE = {\n" + "".join(save_lines if save_lines != ("",) else []) + "}\n"
    constants_path.write_text(body, encoding="utf-8")
    return constants_path


def test_sync_brigade_document_numbers_persist_writes_directly(tmp_path):
    doc_path = _write_docx(tmp_path / "doc.docx", [
        "Розпорядження № 500 від 01.09.2026.",
        "Наказую продовжити обороняти батальйонний район оборони.",
    ])
    folder = tmp_path / "src"
    folder.mkdir()
    os.replace(_zip_up(tmp_path / "a.zip", doc_path), folder / "a.zip")

    day_dict = {"02.09.2026": {"брг": ""}}
    day_keys = ["02.09.2026"]
    constants_path = _constants_file_with_brigade_block(
        tmp_path, ["    '02.09.2026': {'брг': ''},\n"],
    )

    constants_sync.sync_brigade_document_numbers_from_folder(
        str(folder), str(constants_path), day_dict, day_keys, {}, [], persist=True,
    )

    assert "'брг': '500'" in constants_path.read_text(encoding="utf-8")
    assert (tmp_path / "constants.py.bak").exists()


# -------------------------
# persist_number_dicts_to_constants_file (узагальнена сигнатура - fields-список)
# -------------------------
def test_persist_number_dicts_writes_only_given_fields(tmp_path):
    constants_path = tmp_path / "constants.py"
    constants_path.write_text(
        "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY = {\n"
        "    '01.09.2026': {'бат': '167', 'брг': ''},\n"
        "}\n",
        encoding="utf-8",
    )
    day_dict = {"01.09.2026": {"бат": "167", "брг": "2682"}}

    constants_sync.persist_number_dicts_to_constants_file(str(constants_path), [
        ("NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", "брг", day_dict, ["01.09.2026"]),
    ])

    updated = constants_path.read_text(encoding="utf-8")
    assert "'брг': '2682'" in updated
    assert "'бат': '167'" in updated  # поле, яке НЕ передали в fields, лишилось як було
    assert os.path.isfile(str(constants_path) + ".bak")


# -------------------------
# sync_document_numbers_from_folder - очищення номерів ВИДАЛЕНИХ з папки документів
# -------------------------
def _folder_with_files(tmp_path, file_names):
    folder = tmp_path / "src"
    folder.mkdir(exist_ok=True)
    for name in file_names:
        (folder / name).write_bytes(b"")
    return str(folder)


def test_sync_document_numbers_clears_safety_number_of_deleted_document(tmp_path, capsys):
    """Реальний випадок: два застосування безпеки видалили з папки, але їхні номери в
    NUMBER_OF_DOCUMENTS_BRS_SAVE лишились - br_save генерувався з неіснуючим розпорядженням.
    Дата без файлу очищується, дата з файлом отримує свій номер, бз_брг не чіпається, а
    у повідомленні користувачу - лише реально очищені дати."""
    folder = _folder_with_files(tmp_path, ["№500 застосування безпеки 02.09.2026.docx"])
    save_dict = {
        "01.09.2026": {"бз_бат": "480", "бз_брг": "1031"},
        "02.09.2026": {"бз_бат": "499", "бз_брг": "1036"},
        "03.09.2026": {"бз_бат": "", "бз_брг": "1042"},
    }

    constants_sync.sync_document_numbers_from_folder(
        folder, "unused.py", {}, [], {}, [], save_dict, list(save_dict), persist=False,
    )

    assert save_dict["02.09.2026"]["бз_бат"] == "500"
    assert save_dict["01.09.2026"]["бз_бат"] == ""
    assert save_dict["01.09.2026"]["бз_брг"] == "1031"
    out = capsys.readouterr().out
    assert "01.09.2026" in out
    assert "03.09.2026" not in out  # вже було порожнім - нічого не очищено, нема що повідомляти


def test_sync_document_numbers_keeps_numbers_of_document_kinds_absent_from_folder(tmp_path):
    """Папка лише з ЗАВДАННЯМИ не має стерти ні ЩОДЕННА, ні БЗ-номери: очищається лише той
    вид документа, якого в папці знайдено хоча б один (тут - тижневий "бат")."""
    folder = _folder_with_files(tmp_path, ["№7 ЗАВДАННЯ 01.09.2026.docx"])
    day_dict = {"01.09.2026": {"бат": "166", "брг": "", "general_br_bat": ""}}
    week_dict = {
        "01.09.2026": {"бат": "", "посилання_брг": "", "посилання_бат": ""},
        "08.09.2026": {"бат": "9", "посилання_брг": "", "посилання_бат": ""},
    }
    save_dict = {"01.09.2026": {"бз_бат": "480", "бз_брг": "1031"}}

    constants_sync.sync_document_numbers_from_folder(
        folder, "unused.py",
        day_dict, list(day_dict), week_dict, list(week_dict), save_dict, list(save_dict),
        persist=False,
    )

    assert day_dict["01.09.2026"]["бат"] == "166"
    assert save_dict["01.09.2026"]["бз_бат"] == "480"
    assert week_dict["01.09.2026"]["бат"] == "7"
    assert week_dict["08.09.2026"]["бат"] == ""  # тижневий документ за цю дату більше не знайдено


def test_sync_document_numbers_empty_folder_clears_nothing(tmp_path, capsys):
    """Жодного відповідного файлу - ранній вихід ДО очищення: номери в пам'яті не чіпаються."""
    folder = tmp_path / "empty"
    folder.mkdir()
    save_dict = {"01.09.2026": {"бз_бат": "480", "бз_брг": "1031"}}

    constants_sync.sync_document_numbers_from_folder(
        str(folder), "unused.py", {}, [], {}, [], save_dict, list(save_dict), persist=False,
    )

    assert save_dict["01.09.2026"]["бз_бат"] == "480"
    assert "не знайдено файлів" in capsys.readouterr().out


def test_sync_document_numbers_persists_cleared_safety_number_to_constants(tmp_path):
    """Очищене бз_бат має потрапити у constants.py НА ДИСКУ (persist=True за замовчуванням),
    а не лишатися лише в пам'яті поточного запуску."""
    folder = _folder_with_files(tmp_path, ["№500 застосування безпеки 02.09.2026.docx"])
    constants_path = tmp_path / "constants.py"
    constants_path.write_text(
        "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY = {\n"
        "}\n"
        "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK = {\n"
        "}\n"
        "NUMBER_OF_DOCUMENTS_BRS_SAVE = {\n"
        "    '01.09.2026': {'бз_бат': '480', 'бз_брг': '1031'},\n"
        "    '02.09.2026': {'бз_бат': '499', 'бз_брг': '1036'},\n"
        "}\n",
        encoding="utf-8",
    )
    save_dict = {
        "01.09.2026": {"бз_бат": "480", "бз_брг": "1031"},
        "02.09.2026": {"бз_бат": "499", "бз_брг": "1036"},
    }

    constants_sync.sync_document_numbers_from_folder(
        folder, str(constants_path), {}, [], {}, [], save_dict, ["01.09.2026", "02.09.2026"],
    )

    updated = constants_path.read_text(encoding="utf-8")
    assert "'01.09.2026': {'бз_бат': '', 'бз_брг': '1031'}," in updated
    assert "'02.09.2026': {'бз_бат': '500', 'бз_брг': '1036'}," in updated


# -------------------------
# refresh_stage_backup_numbers - бекап етапу після синхронізації БРГ
# -------------------------
_STAGE_BACKUP_TEXT = (
    "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY = {\n"
    "    '01.09.2026': {'бат': '167', 'брг': '', 'general_br_bat': '166 від 01.09.2026'},\n"
    "}\n"
    "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK = {\n"
    "    '01.09.2026': {'бат': '166', 'посилання_брг': '', 'посилання_бат': ''},\n"
    "}\n"
    "NUMBER_OF_DOCUMENTS_BRS_SAVE = {\n"
    "    '01.09.2026': {'бз_бат': '256', 'бз_брг': ''},\n"
    "}\n"
)


def test_refresh_stage_backup_numbers_writes_brg_numbers_from_memory(tmp_path):
    """РЕГРЕСІЯ: після синхронізації БРГ у режимі етапу бекап мав отримати нові БРГ -
    інакше наступний запуск з цим етапом зчитає старі значення з бекапу."""
    backup_path = tmp_path / "constants.py.bak"
    backup_path.write_text(_STAGE_BACKUP_TEXT, encoding="utf-8")
    day_dict = {"01.09.2026": {"бат": "167", "брг": "2682"}}  # general_br_bat немає в пам'яті - не чіпаємо
    week_dict = {"01.09.2026": {"бат": "166", "посилання_брг": "", "посилання_бат": ""}}
    save_dict = {"01.09.2026": {"бз_бат": "256", "бз_брг": "1031"}}

    constants_sync.refresh_stage_backup_numbers(
        str(backup_path), day_dict, week_dict, save_dict, ["01.09.2026"], ["01.09.2026"], ["01.09.2026"],
    )

    updated = backup_path.read_text(encoding="utf-8")
    assert "'брг': '2682'" in updated
    assert "'бз_брг': '1031'" in updated
    assert "'general_br_bat': '166 від 01.09.2026'" in updated
    assert "'посилання_брг': ''" in updated


def test_refresh_stage_backup_numbers_does_nothing_without_backup(tmp_path):
    backup_path = tmp_path / "constants.py.bak"

    constants_sync.refresh_stage_backup_numbers(
        str(backup_path), {}, {}, {}, [], [], [],
    )

    assert not backup_path.exists()


def test_refresh_stage_backup_numbers_leaves_old_format_backup_untouched(tmp_path, capsys):
    backup_path = tmp_path / "constants.py.bak"
    old_format = _STAGE_BACKUP_TEXT.split("NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK")[0]  # без словника тижневих
    backup_path.write_text(old_format, encoding="utf-8")

    constants_sync.refresh_stage_backup_numbers(
        str(backup_path), {"01.09.2026": {"брг": "2682"}}, {}, {"01.09.2026": {"бз_брг": "1031"}},
        ["01.09.2026"], [], ["01.09.2026"],
    )

    assert backup_path.read_text(encoding="utf-8") == old_format
    assert "застарілий формат" in capsys.readouterr().out


def test_refresh_stage_backup_numbers_rejects_value_that_would_break_the_file(tmp_path, capsys):
    """Значення з апострофом зіпсувало б Python-файл бекапу - він має лишитися без змін,
    а не стати нечитабельним (наступний запуск інакше падав би на імпорті)."""
    backup_path = tmp_path / "constants.py.bak"
    backup_path.write_text(_STAGE_BACKUP_TEXT, encoding="utf-8")

    constants_sync.refresh_stage_backup_numbers(
        str(backup_path), {}, {}, {"01.09.2026": {"бз_брг": "2'682"}},
        [], [], ["01.09.2026"],
    )

    assert backup_path.read_text(encoding="utf-8") == _STAGE_BACKUP_TEXT
    assert "бекап лишено без змін" in capsys.readouterr().out
