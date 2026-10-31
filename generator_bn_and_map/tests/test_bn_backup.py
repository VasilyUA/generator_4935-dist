"""Збереження resources/ і output/ в іншу теку в кінці запуску (як у generator_timesheet)."""
from datetime import date

from utils.backup import backup_resources_and_output, snapshot_folder_name


def test_snapshot_folder_name_prefers_order_date_then_pbd_date():
    assert snapshot_folder_name(date(2026, 8, 24), "2026-10-09T20:00") == "24.08.2026"
    assert snapshot_folder_name(None, "2026-10-09T20:00") == "09.10.2026"          # лише карта - дата ПБД
    assert snapshot_folder_name(None, "", today=date(2026, 10, 12)) == "12.10.2026"


def test_backup_copies_both_folders_and_keeps_originals(tmp_path):
    resources, output = tmp_path / "resources", tmp_path / "output"
    (resources / "ПБД").mkdir(parents=True)
    (resources / "ПБД" / "ПБД.docx").write_bytes(b"pbd")
    (resources / "data.json").write_text("{}", encoding="utf-8")
    (output / "map").mkdir(parents=True)
    (output / "Бойовий наказ.docx").write_bytes(b"bn")
    (output / "~$йовий наказ.docx").write_bytes(b"lock")      # замок Word, поки БН відкритий
    (output / "map" / "карта.kmz").write_bytes(b"kmz")

    target = backup_resources_and_output(str(tmp_path / "архів"), "24.08.2026", str(resources), str(output))

    assert target == str(tmp_path / "архів" / "24.08.2026")
    saved = tmp_path / "архів" / "24.08.2026"
    assert (saved / "resources" / "ПБД" / "ПБД.docx").read_bytes() == b"pbd"
    assert (saved / "resources" / "data.json").is_file()
    assert (saved / "output" / "Бойовий наказ.docx").read_bytes() == b"bn"
    assert (saved / "output" / "map" / "карта.kmz").read_bytes() == b"kmz"
    assert not (saved / "output" / "~$йовий наказ.docx").exists()
    assert (resources / "data.json").is_file() and (output / "Бойовий наказ.docx").is_file()   # копія, не переміщення


def test_backup_again_the_same_day_merges(tmp_path):
    resources, output = tmp_path / "resources", tmp_path / "output"
    resources.mkdir()
    output.mkdir()
    (output / "перший.txt").write_text("1", encoding="utf-8")
    backup_resources_and_output(str(tmp_path / "архів"), "24.08.2026", str(resources), str(output))
    (output / "перший.txt").write_text("оновлено", encoding="utf-8")
    (output / "другий.txt").write_text("2", encoding="utf-8")
    backup_resources_and_output(str(tmp_path / "архів"), "24.08.2026", str(resources), str(output))
    saved = tmp_path / "архів" / "24.08.2026" / "output"
    assert (saved / "перший.txt").read_text(encoding="utf-8") == "оновлено" and (saved / "другий.txt").is_file()
