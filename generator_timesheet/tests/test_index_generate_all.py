import index
from constants import ADDON_BOOK_KINDS, REPORT_DIR
from generators.check_addon_mismatches import AddonCheckResult, AddonCrossCheckResult


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
