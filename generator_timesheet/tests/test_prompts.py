from datetime import date

import pytest
from InquirerPy.prompts.input import InputPrompt
from InquirerPy.prompts.list import ListPrompt

import utils.prompts as prompts
from utils.prompts import ask_date, ask_yes_no


def _stub_init(self, message, choices, default):
    pass


def test_ask_yes_no_returns_execute_result(monkeypatch):
    """ListPrompt.__init__ саме по собі торкається реальної консолі
    (NoConsoleScreenBufferError поза справжнім Windows-консольним вікном,
    напр. під Git Bash) - тож __init__ теж заглушено, не лише execute."""
    monkeypatch.setattr(ListPrompt, "__init__", _stub_init)
    monkeypatch.setattr(ListPrompt, "execute", lambda self: True)

    assert ask_yes_no("Відкрити файл?", default=True) is True


def test_ask_yes_no_returns_false_when_declined(monkeypatch):
    monkeypatch.setattr(ListPrompt, "__init__", _stub_init)
    monkeypatch.setattr(ListPrompt, "execute", lambda self: False)

    assert ask_yes_no("Відкрити файл?", default=True) is False


def test_ask_yes_no_passes_message_and_default_to_the_prompt(monkeypatch):
    captured = {}

    def _capture_init(self, message, choices, default):
        captured["message"] = message
        captured["default"] = default
        captured["choice_values"] = [choice["value"] for choice in choices]

    monkeypatch.setattr(ListPrompt, "__init__", _capture_init)
    monkeypatch.setattr(ListPrompt, "execute", lambda self: True)

    ask_yes_no("Відкрити output/ОБЛІК.xlsx?", default=False)

    assert captured["message"] == "Відкрити output/ОБЛІК.xlsx?"
    assert captured["default"] is False
    assert captured["choice_values"] == [False, True]


def _stub_input_init(self, message, default, validate, invalid_message):
    pass


def test_ask_date_returns_the_parsed_date(monkeypatch):
    monkeypatch.setattr(InputPrompt, "__init__", _stub_input_init)
    monkeypatch.setattr(InputPrompt, "execute", lambda self: "01.10.2026")

    assert ask_date("Оберіть дату", default=date(2026, 10, 6)) == date(2026, 10, 1)


def test_ask_date_passes_the_message_and_the_default_as_dd_mm_yyyy_text(monkeypatch):
    captured = {}

    def _capture_init(self, message, default, validate, invalid_message):
        captured["message"] = message
        captured["default"] = default

    monkeypatch.setattr(InputPrompt, "__init__", _capture_init)
    monkeypatch.setattr(InputPrompt, "execute", lambda self: "06.10.2026")

    ask_date("Оберіть дату, за яку вивантажити файли", default=date(2026, 10, 6))

    assert captured["message"] == "Оберіть дату, за яку вивантажити файли"
    assert captured["default"] == "06.10.2026"


@pytest.mark.parametrize("text, expected", [
    ("01.10.2026", True),
    ("2026.10.01", False),  # не той порядок
    ("31.02.2026", False),  # невалідна дата (лютий не має 31 дня)
    ("", False),
    (None, False),
])
def test_is_valid_date_text(text, expected):
    assert prompts._is_valid_date_text(text) is expected
