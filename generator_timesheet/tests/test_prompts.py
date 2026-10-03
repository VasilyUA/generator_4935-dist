from InquirerPy.prompts.list import ListPrompt

from utils.prompts import ask_yes_no


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
