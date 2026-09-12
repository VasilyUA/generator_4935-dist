import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import pytest

import utils.excel_writer as excel_writer_module


class _FakeExcelWorkbook:
    def __init__(self, full_name):
        self.FullName = full_name
        self.closed_with = None

    def Close(self, SaveChanges=False):
        self.closed_with = SaveChanges


class _BrokenExcelWorkbook:
    @property
    def FullName(self):
        raise RuntimeError("помилка COM при читанні шляху")


# -------------------------
# _close_excel_workbook_if_open
# -------------------------
def test_close_excel_workbook_if_open_excel_com_unavailable(tmp_path, monkeypatch):
    monkeypatch.setitem(sys.modules, "win32com.client", None)
    excel_writer_module._close_excel_workbook_if_open(str(tmp_path / "result.xlsx"))


def test_close_excel_workbook_if_open_excel_not_running(monkeypatch):
    import pythoncom
    import win32com.client

    monkeypatch.setattr(pythoncom, "CoInitialize", lambda: None)
    monkeypatch.setattr(pythoncom, "CoUninitialize", lambda: None)

    def _raise(name):
        raise Exception("Excel.Application не запущено")

    monkeypatch.setattr(win32com.client, "GetActiveObject", _raise)

    excel_writer_module._close_excel_workbook_if_open("result.xlsx")


def test_close_excel_workbook_if_open_closes_only_exact_match(tmp_path, monkeypatch, capsys):
    import pythoncom
    import win32com.client

    target_file = tmp_path / "result.xlsx"
    target_file.write_text("x")

    matching_wb = _FakeExcelWorkbook(str(target_file))
    unrelated_wb = _FakeExcelWorkbook(str(tmp_path / "other.xlsx"))
    broken_wb = _BrokenExcelWorkbook()

    class FakeExcelApp:
        Workbooks = [matching_wb, broken_wb, unrelated_wb]

    monkeypatch.setattr(pythoncom, "CoInitialize", lambda: None)
    monkeypatch.setattr(pythoncom, "CoUninitialize", lambda: None)
    monkeypatch.setattr(win32com.client, "GetActiveObject", lambda name: FakeExcelApp())

    excel_writer_module._close_excel_workbook_if_open(str(target_file))

    assert matching_wb.closed_with is False
    assert unrelated_wb.closed_with is None
    assert "Закриваю відкритий у Excel файл" in capsys.readouterr().out


# -------------------------
# save_workbook_safely
# -------------------------
def test_save_workbook_safely_succeeds_first_try(tmp_path, monkeypatch):
    monkeypatch.setattr(excel_writer_module, "_close_excel_workbook_if_open", lambda path: None)

    class FakeWorkbook:
        def __init__(self):
            self.saved_to = None

        def save(self, path):
            self.saved_to = path

    workbook = FakeWorkbook()
    target = str(tmp_path / "result.xlsx")
    excel_writer_module.save_workbook_safely(workbook, target)

    assert workbook.saved_to == target


def test_save_workbook_safely_retries_then_succeeds(monkeypatch):
    monkeypatch.setattr(excel_writer_module, "_close_excel_workbook_if_open", lambda path: None)
    monkeypatch.setattr(excel_writer_module.time, "sleep", lambda s: None)

    call_count = {"n": 0}

    class FlakyWorkbook:
        def save(self, path):
            call_count["n"] += 1
            if call_count["n"] == 1:
                raise PermissionError("заблоковано")

    excel_writer_module.save_workbook_safely(FlakyWorkbook(), "result.xlsx")

    assert call_count["n"] >= 2


def test_save_workbook_safely_raises_after_all_retries_fail(monkeypatch):
    monkeypatch.setattr(excel_writer_module, "_close_excel_workbook_if_open", lambda path: None)
    monkeypatch.setattr(excel_writer_module.time, "sleep", lambda s: None)

    class AlwaysFailsWorkbook:
        def save(self, path):
            raise PermissionError("завжди заблоковано")

    with pytest.raises(PermissionError, match="Не вдалось зберегти"):
        excel_writer_module.save_workbook_safely(AlwaysFailsWorkbook(), "result.xlsx")
