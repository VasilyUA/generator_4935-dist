import os

import pytest

import utils.output_folder as output_folder
from utils.output_folder import clear_information_unit_directory, clear_output_directory, open_file


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


def test_clear_information_unit_directory_keeps_the_named_file_and_removes_the_rest(tmp_path):
    directory = tmp_path / "information_unit"
    directory.mkdir()
    (directory / "Управління_Штаб.xlsx").write_text("лишити")
    (directory / "БЧС ВЗ 01.10.2026.xlsx").write_text("прибрати")

    clear_information_unit_directory(str(directory), "Управління_Штаб.xlsx")

    assert os.listdir(str(directory)) == ["Управління_Штаб.xlsx"]


def test_clear_information_unit_directory_creates_the_directory_if_missing(tmp_path):
    directory = tmp_path / "information_unit"
    assert not directory.exists()

    clear_information_unit_directory(str(directory), "Управління_Штаб.xlsx")

    assert directory.is_dir()
    assert os.listdir(str(directory)) == []


def test_clear_information_unit_directory_skips_a_locked_file_without_crashing(tmp_path, monkeypatch):
    """Той самий реальний випадок, що й clear_output_directory вище - файл,
    відкритий в Excel, лишається, решта видаляється, виклик не падає."""
    directory = tmp_path / "information_unit"
    directory.mkdir()
    locked_name = "БЧС ВЗ 01.10.2026.xlsx"
    (directory / locked_name).write_text("відкрито в Excel")
    (directory / "БЧС 1 РМП 01.10.2026.xlsx").write_text("звичайний")

    original_remove = os.remove

    def _remove(path, *args, **kwargs):
        if os.path.basename(path) == locked_name:
            raise PermissionError("[WinError 32] заблоковано іншою програмою")
        return original_remove(path, *args, **kwargs)

    monkeypatch.setattr(output_folder.os, "remove", _remove)

    clear_information_unit_directory(str(directory), "Управління_Штаб.xlsx")  # не мало підняти виняток

    assert (directory / locked_name).exists()
    assert not (directory / "БЧС 1 РМП 01.10.2026.xlsx").exists()


def test_clear_information_unit_directory_does_not_remove_subdirectories(tmp_path):
    directory = tmp_path / "information_unit"
    sub = directory / "архів"
    sub.mkdir(parents=True)

    clear_information_unit_directory(str(directory), "Управління_Штаб.xlsx")

    assert sub.is_dir()


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
