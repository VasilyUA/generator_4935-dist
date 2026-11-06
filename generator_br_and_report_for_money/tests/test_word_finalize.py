import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import pytest
from docx import Document

from generators import word_finalize
from generators.word_finalize import (
    FinalizeError,
    REPORT_SHEETS_BLANK,
    REPORT_SHEETS_MARKER,
    finalize_report_sheet_count,
    finalize_report_with_word,
    finalize_report_without_word,
)


class _FakeFind:
    def __init__(self):
        self.executed = []

    def ClearFormatting(self):
        pass

    @property
    def Replacement(self):
        return self

    def Execute(self, *args):
        self.executed.append(args)


class _FakeDoc:
    def __init__(self, pages, fail=False):
        self.pages = pages
        self.fail = fail
        self.find = _FakeFind()
        self.saved = self.closed = False

    def Repaginate(self):
        if self.fail:
            raise RuntimeError("COM error")

    def ComputeStatistics(self, kind):
        assert kind == word_finalize.WD_STATISTIC_PAGES
        return self.pages

    @property
    def Content(self):
        return self

    @property
    def Find(self):
        return self.find

    def Save(self):
        self.saved = True

    def Close(self, save):
        self.closed = True


class _FakeWord:
    def __init__(self, doc):
        self.doc = doc
        self.opened = []
        self.quit = False

    @property
    def Documents(self):
        return self

    def Open(self, path, AddToRecentFiles):
        self.opened.append(path)
        return self.doc

    def Quit(self):
        self.quit = True


def _opener(word, released):
    return lambda: (word, lambda: released.append(True))


def test_finalize_report_with_word_computes_pages_and_replaces_marker(tmp_path):
    doc = _FakeDoc(133)
    word = _FakeWord(doc)
    released = []
    assert finalize_report_with_word(tmp_path / "b.docx", open_word=_opener(word, released)) == 133
    replace = doc.find.executed[0]
    assert replace[0] == REPORT_SHEETS_MARKER
    assert replace[9] == "133 (ста тридцяти трьох)"
    assert replace[10] == word_finalize.WD_REPLACE_ALL
    assert doc.saved and doc.closed and word.quit and released


def test_finalize_report_with_word_wraps_com_errors(tmp_path):
    word = _FakeWord(_FakeDoc(1, fail=True))
    released = []
    with pytest.raises(FinalizeError, match="COM error"):
        finalize_report_with_word(tmp_path / "b.docx", open_word=_opener(word, released))
    assert word.doc.closed and released


def test_finalize_report_with_word_reports_missing_word(tmp_path):
    def _no_word():
        raise ImportError("no pywin32")
    with pytest.raises(FinalizeError, match="недоступний"):
        finalize_report_with_word(tmp_path / "b.docx", open_word=_no_word)


def test_finalize_report_with_word_ignores_errors_while_closing(tmp_path):
    doc = _FakeDoc(3)
    word = _FakeWord(doc)

    def _quit():
        raise RuntimeError("already closed")
    word.Quit = _quit
    assert finalize_report_with_word(tmp_path / "b.docx", open_word=_opener(word, [])) == 3


def test_finalize_report_without_word_replaces_marker_with_blank(tmp_path):
    path = tmp_path / "b.docx"
    doc = Document()
    doc.add_paragraph(f"Рапорт подається на {REPORT_SHEETS_MARKER} аркушах.")
    doc.add_paragraph("інший текст")
    doc.save(path)

    finalize_report_without_word(path)

    result = Document(path)
    assert result.paragraphs[0].text == f"Рапорт подається на {REPORT_SHEETS_BLANK} аркушах."
    assert result.paragraphs[1].text == "інший текст"


def test_finalize_report_sheet_count_uses_word_result_when_available(tmp_path, monkeypatch):
    calls = []
    monkeypatch.setattr(word_finalize, "finalize_report_with_word", lambda path: calls.append(path) or 42)
    monkeypatch.setattr(word_finalize, "finalize_report_without_word", lambda path: calls.append(("fallback", path)))

    finalize_report_sheet_count("report.docx")

    assert calls == ["report.docx"]


def test_finalize_report_sheet_count_falls_back_without_crashing_when_word_fails(tmp_path, monkeypatch, capsys):
    """РЕГРЕСІЯ: відсутність/збій Word НЕ має валити всю генерацію рапорту -
    лишається прочерк, лише жовте попередження в термінал."""
    monkeypatch.setattr(word_finalize, "finalize_report_with_word", lambda path: (_ for _ in ()).throw(FinalizeError("Word недоступний")))
    fallback_calls = []
    monkeypatch.setattr(word_finalize, "finalize_report_without_word", lambda path: fallback_calls.append(path))

    finalize_report_sheet_count("report.docx")

    assert fallback_calls == ["report.docx"]
    assert "Word недоступний" in capsys.readouterr().out
