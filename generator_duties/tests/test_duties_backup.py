from utils.backup import save_output


def test_save_output_copies_files_skips_office_locks_and_overwrites(tmp_path):
    output = tmp_path / "output"
    output.mkdir()
    (output / "книга.docx").write_text("1", encoding="utf-8")
    (output / "~$книга.docx").write_text("lock", encoding="utf-8")
    target = save_output(str(output), str(tmp_path / "архів"), "Обов'язки 2027")
    assert sorted(p.name for p in (tmp_path / "архів" / "Обов'язки 2027").iterdir()) == ["книга.docx"]
    (output / "книга.docx").write_text("2", encoding="utf-8")
    assert save_output(str(output), str(tmp_path / "архів"), "Обов'язки 2027") == target
    assert (tmp_path / "архів" / "Обов'язки 2027" / "книга.docx").read_text(encoding="utf-8") == "2"
