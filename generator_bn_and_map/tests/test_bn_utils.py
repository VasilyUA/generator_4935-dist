from datetime import date

import pytest
from docx import Document

import utils.prompts as prompts
from utils.docx_utils import save_docx_safely


class _FakePrompt:
    answers = []
    calls = []

    def __init__(self, **kwargs):
        self.kwargs = kwargs
        _FakePrompt.calls.append(kwargs)

    def execute(self):
        return _FakePrompt.answers.pop(0) if _FakePrompt.answers else self.kwargs.get("default")


@pytest.fixture
def fake_prompts(monkeypatch):
    # Підміна імен, прив'язаних саме в utils.prompts - глобальні класи InquirerPy
    # не чіпаються (їх патчать conftest-и інших проєктів).
    monkeypatch.setattr(prompts, "InputPrompt", _FakePrompt)
    monkeypatch.setattr(prompts, "ListPrompt", _FakePrompt)
    _FakePrompt.answers, _FakePrompt.calls = [], []
    return _FakePrompt


def test_ask_yes_no(fake_prompts):
    fake_prompts.answers = [True]
    assert prompts.ask_yes_no("Питання?", default=False) is True
    assert [c["value"] for c in fake_prompts.calls[0]["choices"]] == [False, True]


def test_ask_date_default_and_validation(fake_prompts):
    assert prompts.ask_date("Дата", default=date(2026, 8, 24)) == date(2026, 8, 24)
    validate = fake_prompts.calls[0]["validate"]
    assert validate("24.08.2026") and not validate("2026-08-24") and not validate(None)


def test_ask_text(fake_prompts):
    fake_prompts.answers = ["  7 "]
    assert prompts.ask_text("Номер", default=5, validate=str.isdigit) == "7"
    assert fake_prompts.calls[0]["default"] == "5" and "validate" in fake_prompts.calls[0]
    assert prompts.ask_text("Без перевірки") == ""
    assert "validate" not in fake_prompts.calls[1]


def test_save_docx_safely_retries_then_raises(tmp_path, monkeypatch):
    doc = Document()
    attempts = []

    def locked(path):
        attempts.append(path)
        raise PermissionError("locked")

    monkeypatch.setattr(doc, "save", locked)
    with pytest.raises(PermissionError, match="заблоковано"):
        save_docx_safely(doc, str(tmp_path / "x.docx"), attempts=3, delay=0)
    assert len(attempts) == 3


def test_save_docx_safely_saves(tmp_path):
    path = tmp_path / "x.docx"
    save_docx_safely(Document(), str(path))
    assert path.is_file()


def test_clear_output_directory(tmp_path):
    from utils.output_folder import clear_output_directory
    target = tmp_path / "output"
    (target / "map").mkdir(parents=True)
    (target / "map" / "старий.csv").write_text("x")
    (target / "старий.docx").write_text("x")
    clear_output_directory(str(target))
    assert target.is_dir() and list(target.iterdir()) == []
    clear_output_directory(str(tmp_path / "нова"))
    assert (tmp_path / "нова").is_dir()
