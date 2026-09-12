import os
from datetime import date

from utils.backup import backup_resources_and_output, latest_report_date


def _touch(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("щось")


def test_latest_report_date_picks_the_maximum_among_report_files(tmp_path):
    report_dir = tmp_path / "report"
    _touch(report_dir / "01.08.2026 - щоденний рапорт.docx")
    _touch(report_dir / "23.08.2026 - щоденний рапорт.docx")
    _touch(report_dir / "15.08.2026 - щоденний рапорт.docx")

    assert latest_report_date(str(report_dir)) == date(2026, 8, 23)


def test_latest_report_date_ignores_office_lock_files(tmp_path):
    """Реальний випадок: рапорт ВІДКРИТИЙ у Word - тимчасовий lock-файл
    "~$..." має ТЕ САМЕ розширення (.docx), тож інакше хибно "переміг" би
    справжні дати (сам lock-файл - не рапорт, лише маленька позначка)."""
    report_dir = tmp_path / "report"
    _touch(report_dir / "01.08.2026 - щоденний рапорт.docx")
    _touch(report_dir / "~$31.12.2026 - щоденний рапорт.docx")

    assert latest_report_date(str(report_dir)) == date(2026, 8, 1)


def test_latest_report_date_returns_none_when_no_file_has_a_recognizable_date(tmp_path):
    report_dir = tmp_path / "report"
    _touch(report_dir / "щоденний рапорт.docx")

    assert latest_report_date(str(report_dir)) is None


def test_latest_report_date_returns_none_for_missing_directory(tmp_path):
    assert latest_report_date(str(tmp_path / "не_існує")) is None


def test_backup_resources_and_output_copies_both_directories_into_a_dated_subfolder(tmp_path):
    report_dir = tmp_path / "resources" / "report"
    _touch(report_dir / "23.08.2026 - щоденний рапорт.docx")
    resources_dir = tmp_path / "resources"
    _touch(resources_dir / "ОБЛІК.xlsx")
    output_dir = tmp_path / "output"
    _touch(output_dir / "ОБЛІК.xlsx")
    destination_root = tmp_path / "backup"

    target_dir = backup_resources_and_output(str(destination_root), str(report_dir), str(resources_dir), str(output_dir))

    assert target_dir == os.path.join(str(destination_root), "23.08.2026")
    assert (destination_root / "23.08.2026" / "resources" / "ОБЛІК.xlsx").is_file()
    assert (destination_root / "23.08.2026" / "resources" / "report" / "23.08.2026 - щоденний рапорт.docx").is_file()
    assert (destination_root / "23.08.2026" / "output" / "ОБЛІК.xlsx").is_file()


def test_backup_resources_and_output_returns_none_without_a_dated_report(tmp_path):
    """Немає жодного рапорту з розпізнаваною датою - немає, як назвати нову
    підтеку, тож нічого не копіюється взагалі (не половинчастий результат)."""
    report_dir = tmp_path / "resources" / "report"
    report_dir.mkdir(parents=True)
    resources_dir = tmp_path / "resources"
    _touch(resources_dir / "ОБЛІК.xlsx")
    output_dir = tmp_path / "output"
    _touch(output_dir / "ОБЛІК.xlsx")
    destination_root = tmp_path / "backup"

    target_dir = backup_resources_and_output(str(destination_root), str(report_dir), str(resources_dir), str(output_dir))

    assert target_dir is None
    assert not destination_root.exists()


def test_backup_resources_and_output_overwrites_on_a_repeated_run_for_the_same_date(tmp_path):
    """Повторний прогін на ТУ САМУ дату (напр. ще раз того самого дня) - НЕ
    падає з FileExistsError (dirs_exist_ok=True), а домішує/перезаписує
    вміст підтеки оновленими файлами."""
    report_dir = tmp_path / "resources" / "report"
    _touch(report_dir / "23.08.2026 - щоденний рапорт.docx")
    resources_dir = tmp_path / "resources"
    _touch(resources_dir / "ОБЛІК.xlsx")
    output_dir = tmp_path / "output"
    _touch(output_dir / "ОБЛІК.xlsx")
    destination_root = tmp_path / "backup"

    backup_resources_and_output(str(destination_root), str(report_dir), str(resources_dir), str(output_dir))
    (resources_dir / "ОБЛІК.xlsx").write_text("оновлено")
    target_dir = backup_resources_and_output(str(destination_root), str(report_dir), str(resources_dir), str(output_dir))

    assert target_dir == os.path.join(str(destination_root), "23.08.2026")
    assert (destination_root / "23.08.2026" / "resources" / "ОБЛІК.xlsx").read_text() == "оновлено"
