from utils import prompts
from utils.logging_utils import print_green, print_purple, print_red


class _Prompt:
    answers = []
    created = []

    def __init__(self, **kwargs):
        _Prompt.created.append(kwargs)

    def execute(self):
        return _Prompt.answers.pop(0)


def test_ask_yes_no_and_ask_text(monkeypatch):
    monkeypatch.setattr(prompts, "ListPrompt", _Prompt)
    monkeypatch.setattr(prompts, "InputPrompt", _Prompt)
    _Prompt.answers = [True, " 2027 ", "x"]
    _Prompt.created = []
    assert prompts.ask_yes_no("Так?", default=False) is True
    assert prompts.ask_text("Рік:", default=2027, validate=str.isdigit) == "2027"
    assert prompts.ask_text("Текст:") == "x"
    assert _Prompt.created[1]["default"] == "2027" and "validate" in _Prompt.created[1]
    assert "validate" not in _Prompt.created[2]


def test_print_helpers(capsys):
    print_red("червоний")
    print_green("зелений")
    print_purple("фіолетовий")
    out = capsys.readouterr().out
    assert "червоний" in out and "зелений" in out and "фіолетовий" in out
