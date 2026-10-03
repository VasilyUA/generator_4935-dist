import os

import pytest

import utils.output_folder as output_folder
from utils.output_folder import clear_output_directory, open_file


def test_clear_output_directory_removes_existing_files(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    (output_dir / "стара ОБЛІК.xlsx").write_text("старе")

    clear_output_directory(str(output_dir))

    assert output_dir.is_dir()
    assert list(output_dir.iterdir()) == []


def test_clear_output_directory_creates_missing_directory(tmp_path):
    output_dir = tmp_path / "output"
    assert not output_dir.exists()

    clear_output_directory(str(output_dir))

    assert output_dir.is_dir()


def test_clear_output_directory_removes_nested_files_too(tmp_path):
    output_dir = tmp_path / "output"
    nested = output_dir / "sub"
    nested.mkdir(parents=True)
    (nested / "файл.txt").write_text("щось")

    clear_output_directory(str(output_dir))

    assert output_dir.is_dir()
    assert not nested.exists()
    assert os.listdir(str(output_dir)) == []


def test_clear_output_directory_skips_a_locked_file_without_crashing(tmp_path, monkeypatch):
    """Реальний випадок: ОБЛІК.xlsx (чи його тимчасовий "~$..." lock-файл)
    відкритий у Excel - shutil.rmtree впадав з PermissionError і зривав весь
    запуск (traceback). Тепер такий файл просто лишається як є, решта теки
    видаляється нормально, а сам виклик не падає."""
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    locked_name = "~$ОБЛІК.xlsx"
    (output_dir / locked_name).write_text("lock")
    (output_dir / "звичайний.txt").write_text("щось")

    original_unlink = os.unlink

    def _unlink(path, *args, **kwargs):
        if os.path.basename(path) == locked_name:
            raise PermissionError("[WinError 32] заблоковано іншою програмою")
        return original_unlink(path, *args, **kwargs)

    monkeypatch.setattr(output_folder.os, "unlink", _unlink)

    clear_output_directory(str(output_dir))  # не мало підняти виняток

    assert output_dir.is_dir()
    assert (output_dir / locked_name).exists()  # заблокований файл лишився
    assert not (output_dir / "звичайний.txt").exists()  # решта видалена


def test_ignore_os_error_swallows_os_errors():
    output_folder._ignore_os_error(None, None, OSError("заблоковано"))  # не мало підняти виняток


def test_ignore_os_error_reraises_non_os_errors():
    """shutil.rmtree сам ЗАВЖДИ викликає onexc лише з OSError (перехоплює
    рівно except OSError у власному коді) - ця гілка недосяжна через реальний
    виклик clear_output_directory, тож перевіряємо контракт _ignore_os_error
    напряму, як предохоронник про всяк випадок."""
    with pytest.raises(ValueError, match="щось геть неочікуване"):
        output_folder._ignore_os_error(None, None, ValueError("щось геть неочікуване"))


def test_open_file_calls_os_startfile_for_an_existing_file(tmp_path, monkeypatch):
    file_path = tmp_path / "review.txt"
    file_path.write_text("щось")
    calls = []
    monkeypatch.setattr(output_folder.os, "startfile", calls.append, raising=False)

    open_file(str(file_path))

    assert calls == [str(file_path)]


def test_open_file_does_nothing_for_a_missing_file(monkeypatch):
    calls = []
    monkeypatch.setattr(output_folder.os, "startfile", calls.append, raising=False)

    open_file("не/існує.txt")

    assert calls == []


def test_open_file_does_nothing_for_an_empty_path(monkeypatch):
    calls = []
    monkeypatch.setattr(output_folder.os, "startfile", calls.append, raising=False)

    open_file("")

    assert calls == []


def test_open_file_swallows_os_error_when_startfile_fails(tmp_path, monkeypatch):
    file_path = tmp_path / "review.txt"
    file_path.write_text("щось")

    def _raise(path):
        raise OSError("немає типової програми для відкриття")

    monkeypatch.setattr(output_folder.os, "startfile", _raise, raising=False)

    open_file(str(file_path))  # не мало підняти виняток


def test_open_file_swallows_attribute_error_when_startfile_unavailable(tmp_path, monkeypatch):
    file_path = tmp_path / "review.txt"
    file_path.write_text("щось")
    monkeypatch.delattr(output_folder.os, "startfile", raising=False)

    open_file(str(file_path))  # не мало підняти виняток (не-Windows платформа)
