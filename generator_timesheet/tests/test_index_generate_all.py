import index


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
