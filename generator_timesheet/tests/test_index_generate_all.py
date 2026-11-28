from datetime import date
from types import SimpleNamespace

import index
from constants import ADDON_BOOK_KINDS, REPORT_DIR
from content.signal_reports import SyncResult
from generators.check_addon_mismatches import AddonCheckResult, AddonCrossCheckResult

_FETCH_DATE = date(2026, 11, 20)


def _run_main(monkeypatch, choice, yes_to=()):
    """index.main() з підставленими відповідями: "Так" - на питання, що
    починаються з будь-якого з yes_to, решта - "Ні"; choice - пункт меню "Що
    зробити?". Повертає події в тому порядку, як їх побачить користувач."""
    events = []

    def ask_date(message, default):
        events.append("дата")
        return _FETCH_DATE

    def ask_yes_no(question, default):
        events.append(question)
        return question.startswith(tuple(yes_to)) if yes_to else False

    def menu(**_kwargs):
        events.append("меню")
        return choice

    monkeypatch.setattr(index, "ask_date", ask_date)
    monkeypatch.setattr(index, "ask_yes_no", ask_yes_no)
    monkeypatch.setattr(index, "ListPrompt", lambda **kwargs: SimpleNamespace(execute=lambda: menu(**kwargs)))
    monkeypatch.setattr(index, "_sync_reports_from_signal", lambda day: events.append(("рапорти до", day)))
    monkeypatch.setattr(index, "_fetch_bchs_files", lambda day: events.append(("БЧС за", day)))
    monkeypatch.setattr(index, "_cross_check_addons", lambda: events.append("звірка додатків"))
    monkeypatch.setattr(index, "clear_output_directory", lambda path: events.append("очистити output"))
    monkeypatch.setattr(index, "_generate_all", lambda: events.append("згенерувати все") or [])
    monkeypatch.setattr(index, "open_folder", lambda path: events.append("відкрити output"))
    monkeypatch.setattr(index, "_backup_resources_and_output", lambda: events.append("зберегти"))
    try:
        index.main()
    except SystemExit:
        events.append("кінець")
    return events


def test_main_asks_one_date_first_and_uses_it_for_both_downloads(monkeypatch):
    """Рішення користувача: найперше - дата; рапорти й відомості - до неї
    включно, файли БЧС - за неї (окремого питання про дату БЧС немає)."""
    events = _run_main(monkeypatch, index.STOP_SCRIPT, yes_to=("Отримати файли",))

    assert events == [
        "дата",
        "Отримати файли resources/report та resources/schedule?", ("рапорти до", _FETCH_DATE),
        "Отримати файли БЧС?", ("БЧС за", _FETCH_DATE),
        "меню", "кінець",
    ]


def test_cross_check_of_addons_is_a_menu_item_and_ends_the_script(monkeypatch):
    """Рішення користувача: "Звірити додатки (resources/addons) між собою" -
    пункт меню "Що зробити?", а не питання на старті; лише він, без очищення
    output/ і без решти питань."""
    events = _run_main(monkeypatch, index.CROSS_CHECK_ADDONS)

    assert events == [
        "дата", "Отримати файли resources/report та resources/schedule?", "Отримати файли БЧС?",
        "меню", "звірка додатків", "кінець",
    ]


def test_opening_output_ends_the_script(monkeypatch):
    """Рішення користувача: "Відкрити output? Так" - кінець, без питання про
    збереження resources та output."""
    events = _run_main(monkeypatch, index.GENERATE_ALL, yes_to=("Відкрити output",))

    assert events[events.index("меню"):] == [
        "меню", "очистити output", "згенерувати все",
        "Здійснити звірку з addons?", "Відкрити output?", "відкрити output", "кінець",
    ]


def _fake_generators(monkeypatch, payments_mismatch=None, schedule_mismatch=None):
    """Генератори "Згенерувати все" без файлів; повертає список подій у тому
    порядку, як їх побачить користувач."""
    events = []
    monkeypatch.setattr(index, "generate_timesheet", lambda output_file_name: ("стройова.xlsx", []))
    monkeypatch.setattr(index, "generate_payments_timesheet", lambda: ("виплати.xlsx", [], payments_mismatch))
    monkeypatch.setattr(index, "generate_rop_vop_statement",
                        lambda output_file_name: (None, None, schedule_mismatch) if schedule_mismatch
                        else (output_file_name, 3, None))
    monkeypatch.setattr(index, "_check_category_100_in", lambda path: events.append(("категорія 100", path)))
    monkeypatch.setattr(index, "_open_if_confirmed", lambda path: events.append(("відкрити?", path)))
    return events


def test_generate_all_asks_to_open_the_rop_vop_mismatch_report(monkeypatch):
    """Рішення користувача: після "N розбіжність(ей) у Відомості РОП/ВОП" -
    питання "Відкрити ...?", як в окремому пункті меню."""
    events = _fake_generators(monkeypatch, schedule_mismatch="error_mis_statuses_Відомість.xlsx")

    obliks = index._generate_all()

    assert events == [("категорія 100", "виплати.xlsx"), ("відкрити?", "error_mis_statuses_Відомість.xlsx")]
    assert [path for _name, path in obliks] == ["виплати.xlsx", "стройова.xlsx"]


def test_generate_all_asks_about_each_report_right_after_its_generator(monkeypatch):
    events = _fake_generators(monkeypatch, payments_mismatch="error_mis_statuses.xlsx",
                              schedule_mismatch="error_mis_statuses_Відомість.xlsx")

    index._generate_all()

    assert events == [
        ("відкрити?", "error_mis_statuses.xlsx"),
        ("категорія 100", "виплати.xlsx"),
        ("відкрити?", "error_mis_statuses_Відомість.xlsx"),
    ]


def test_generate_all_does_not_offer_the_statement_itself(monkeypatch):
    events = _fake_generators(monkeypatch)

    index._generate_all()

    assert events == [("категорія 100", "виплати.xlsx")]


def test_addon_check_gets_the_reports_to_see_statuses_behind_numbers(monkeypatch):
    """Без рапортів звірка готового ОБЛІК для виплат не знає, що за числом 10:
    ВД, ППД чи БЗВП."""
    calls = []

    def fake_check(path, **kwargs):
        calls.append((path, kwargs))
        return AddonCheckResult([], [], None, ["book61.xlsm"], [], [])

    monkeypatch.setattr(index, "addon_files", lambda: ["book61.xlsm"])
    monkeypatch.setattr(index, "check_addon_mismatches_file", fake_check)

    index._check_addon_mismatches_in("виплати.xlsx")

    assert calls == [("виплати.xlsx", {"report_dir": REPORT_DIR})]


def _capture_prints(monkeypatch):
    printed = {"червоним": [], "жовтим": []}
    monkeypatch.setattr(index, "print_red", printed["червоним"].append)
    monkeypatch.setattr(index, "print_yellow", printed["жовтим"].append)
    monkeypatch.setattr(index, "print_green", lambda text: None)
    return printed


def test_missing_book_is_only_a_yellow_note_in_the_addon_check(monkeypatch):
    """Рішення користувача: файлу немає - звіряємо без нього. Це не помилка, але
    видно, що саме не звірялось."""
    printed = _capture_prints(monkeypatch)
    monkeypatch.setattr(index, "addon_files", lambda: ["book61.xlsm"])
    monkeypatch.setattr(index, "check_addon_mismatches_file", lambda path, **kwargs: AddonCheckResult(
        [], [], None, ["book61.xlsm"], [], [], ("ЦВС", "журнал лікування"),
    ))

    index._check_addon_mismatches_in("ОБЛІК.xlsx")

    assert printed["червоним"] == []
    assert [("ЦВС, журнал лікування" in text and "звірено без них" in text) for text in printed["жовтим"]] == [True]


def test_name_only_difference_is_not_counted_as_a_period_mismatch(monkeypatch):
    printed = _capture_prints(monkeypatch)
    monkeypatch.setattr(index, "addon_files", lambda: ["book61.xlsm"])
    monkeypatch.setattr(index, "_open_if_confirmed", lambda path: None)
    name_only = {"source_file": "book61", "dates": {}, "misspelled": [{"pib_raw": "ПЕРШИЙ Перший Перши"}],
                 "roster_pib_raw": "ПЕРШИЙ Перший Перший", "addon_pib_raw": "ПЕРШИЙ Перший Перши"}
    monkeypatch.setattr(index, "check_addon_mismatches_file", lambda path, **kwargs: AddonCheckResult(
        [name_only], [], "mis_match.xlsx", ["book61.xlsm"], [], [],
    ))

    index._check_addon_mismatches_in("ОБЛІК.xlsx")

    (message,) = printed["червоним"]
    assert "1 людей з іншим написанням ПІБ" in message and "періодів" not in message


def test_all_books_in_place_print_no_note(monkeypatch):
    printed = _capture_prints(monkeypatch)
    monkeypatch.setattr(index, "addon_files", lambda: ["book61.xlsm"])
    monkeypatch.setattr(index, "check_addon_mismatches_file",
                        lambda path, **kwargs: AddonCheckResult([], [], None, ["book61.xlsm"], [], []))

    index._check_addon_mismatches_in("ОБЛІК.xlsx")

    assert printed == {"червоним": [], "жовтим": []}


def test_missing_book_is_a_yellow_note_in_the_cross_check_too(monkeypatch):
    printed = _capture_prints(monkeypatch)
    monkeypatch.setattr(index, "cross_check_addons", lambda: AddonCrossCheckResult(
        [], None, ["book61.xlsm", "cvs.xlsx"], [], [], ("журнал обліку БП",),
    ))

    index._cross_check_addons()

    assert printed["червоним"] == []
    assert ["журнал обліку БП" in text for text in printed["жовтим"]] == [True]


def test_unrecognized_addon_message_names_every_kind_of_book(monkeypatch):
    printed = _capture_prints(monkeypatch)

    index._print_unrecognized_addons(["resources/addons/щось.xlsx"])

    (message,) = printed["червоним"]
    assert "щось.xlsx" in message and all(name in message for name in ADDON_BOOK_KINDS.values())


def test_missing_signal_files_are_one_red_message_with_the_group_and_a_line_per_day(monkeypatch):
    """Рішення користувача: замість рядка на кожен відсутній файл - заголовок з
    назвою групи Signal (з великої літери) і по рядку на файл; відомість за
    день має в назві наступну дату."""
    monkeypatch.setattr(index, "REPORTS_SIGNAL_GROUP_NAME", "група підрозділу")
    result = SyncResult("відомості ВОП-РОП", "відомості ВОП-РОП", missing=[
        (date(2026, 11, 24), "ПідрозділВОП-РОП(ЛИСТОПАД) 25.11.2026.xlsx"),
        (date(2026, 11, 25), "ПідрозділВОП-РОП(ЛИСТОПАД) 26.11.2026.xlsx"),
    ])
    printed = _capture_prints(monkeypatch)
    monkeypatch.setattr(index, "sync_reports_and_schedules", lambda today: [result])

    index._sync_reports_from_signal(date(2026, 11, 25))

    assert printed["червоним"] == [
        "В Signal у групі Група підрозділу, відомості ВОП-РОП немає:\n"
        " - файлу за 24.11.2026 - ПідрозділВОП-РОП(ЛИСТОПАД) 25.11.2026.xlsx (ні в Signal, ні в теці)\n"
        " - файлу за 25.11.2026 - ПідрозділВОП-РОП(ЛИСТОПАД) 26.11.2026.xlsx (ні в Signal, ні в теці)."
    ]


def _check_category_100(monkeypatch, violations, check_path):
    """Що побачить користувач після перевірки категорії 100: [(колір, текст)]
    і файли, які запропоновано відкрити."""
    printed, asked = [], []
    monkeypatch.setattr(index, "check_payment_category_100_file", lambda path: (violations, check_path))
    monkeypatch.setattr(index, "print_red", lambda text: printed.append(("червоний", text)))
    monkeypatch.setattr(index, "print_green", lambda text: printed.append(("зелений", text)))
    monkeypatch.setattr(index, "print_yellow", lambda text: printed.append(("жовтий", text)))
    monkeypatch.setattr(index, "_open_if_confirmed", asked.append)
    index._check_category_100_in("ОБЛІК.xlsx")
    return printed, asked


def test_category_100_violations_are_yellow_and_offer_to_open_the_table(monkeypatch):
    """Рішення користувача: порушення - попередження, а не помилка."""
    printed, asked = _check_category_100(monkeypatch, [{}, {}], "таблиця.xlsx")

    assert printed == [("жовтий", "2 порушень - див. таблиця.xlsx")]
    assert asked == ["таблиця.xlsx"]


def test_category_100_table_with_only_marked_people_is_not_a_violation(monkeypatch):
    """Таблиця є лише через людей з кольором: виправляти нічого, тож "Так" на
    питанні "Відкрити" не має зупиняти "Згенерувати все"."""
    printed, asked = _check_category_100(monkeypatch, [], "таблиця.xlsx")

    assert printed == [("зелений", "Порушень правила категорії 100 поспіль не знайдено, люди з кольором - див. таблиця.xlsx")]
    assert asked == []


def test_category_100_without_a_table_says_there_are_no_violations(monkeypatch):
    printed, asked = _check_category_100(monkeypatch, [], None)

    assert printed == [("зелений", "Порушень правила категорії 100 поспіль не знайдено.")]
    assert asked == []
