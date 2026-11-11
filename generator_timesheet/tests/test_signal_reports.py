"""Звіряння resources/report і resources/schedule з групою Signal
(content/signal_reports.py) - справжня база Signal не потрібна: вкладення -
словники, "розшифрування" - функція, що пише відомі байти."""

import os
import sqlite3
from datetime import date, datetime

import pytest

import content.signal_reports as sr
from content.schedule_reader import latest_schedule_file
from content.signal_bchs import SignalFetchError

_SCHEDULE = "ПідрозділВОП-РОП(СЕРПЕНЬ) {}.xlsx"
_TODAY = date(2026, 10, 25)


def _ms(dt):
    return int(dt.timestamp() * 1000)


def _row(file_name, sent, content, edit=-1):
    return {"fileName": file_name, "sentAt": _ms(sent), "content": content, "editHistoryIndex": edit, "path": "p"}


def _decrypt(row, path):
    if row["content"] is None:
        return False
    with open(path, "wb") as f:
        f.write(row["content"])
    return True


def _kinds(tmp_path):
    report_kind, schedule_kind = sr._kinds(_TODAY, str(tmp_path / "report"), str(tmp_path / "schedule"))
    return report_kind, schedule_kind


@pytest.mark.parametrize("file_name, expected", [
    ("17.10.2026 - щоденний рапорт.docx", date(2026, 10, 17)),
    ("17.10.2026_-_щоденний_рапорт.docx", date(2026, 10, 17)),
    (_SCHEDULE.format("25.10.2026"), date(2026, 10, 25)),
    ("БЧС 07.10.26.xlsx", date(2026, 10, 7)),
    ("Без дати.docx", None),
    ("31.02.2026 - щоденний рапорт.docx", None),
])
def test_name_date(file_name, expected):
    assert sr.name_date(file_name) == expected


def test_versions_by_date_puts_the_latest_submission_of_each_date_last(tmp_path):
    """Рапорт подають кілька разів - найновіший за часом останній; дати до
    останнього дня минулого місяця, чужі файли й інші розширення - не
    беруться; рядки історії редагування з тим самим часом - поточна версія
    найновіша."""
    report_kind, _schedule_kind = _kinds(tmp_path)
    rows = [
        _row("17.10.2026 - щоденний рапорт.docx", datetime(2026, 10, 17, 13, 30), b"v1"),
        _row("17.10.2026 - щоденний рапорт.docx", datetime(2026, 10, 17, 16, 51), b"v3"),
        _row("17.10.2026 - щоденний рапорт.docx", datetime(2026, 10, 17, 15, 53), b"v2"),
        _row("18.10.2026 - щоденний рапорт.docx", datetime(2026, 10, 18, 15, 0), b"old", edit=0),
        _row("18.10.2026 - щоденний рапорт.docx", datetime(2026, 10, 18, 15, 0), b"current", edit=-1),
        _row("30.09.2026 - щоденний рапорт.docx", datetime(2026, 9, 30, 16, 0), b"baseline"),
        _row("29.09.2026 - щоденний рапорт.docx", datetime(2026, 9, 29, 16, 0), b"too early"),
        _row("Рапорт поранення 17.10.2026.docx", datetime(2026, 10, 17, 18, 0), b"other"),
        _row("17.10.2026 - щоденний рапорт.pdf", datetime(2026, 10, 17, 19, 0), b"pdf"),
    ]

    versions = sr.versions_by_date(rows, report_kind)

    assert {day: [row["content"] for row in day_versions] for day, day_versions in versions.items()} == {
        date(2026, 10, 17): [b"v1", b"v2", b"v3"], date(2026, 10, 18): [b"old", b"current"],
        date(2026, 9, 30): [b"baseline"],
    }


def test_schedules_start_from_the_first_day_of_the_month(tmp_path):
    _report_kind, schedule_kind = _kinds(tmp_path)
    rows = [
        _row(_SCHEDULE.format("30.09.2026"), datetime(2026, 9, 30, 9), b"september"),
        _row(_SCHEDULE.format("01.10.2026"), datetime(2026, 10, 1, 9), b"first"),
        _row(_SCHEDULE.format("26.10.2026"), datetime(2026, 10, 25, 11), b"sent a day early"),
    ]

    assert sorted(sr.versions_by_date(rows, schedule_kind)) == [date(2026, 10, 1), date(2026, 10, 26)]


def _set_mtime(path, when):
    os.utime(path, (when.timestamp(), when.timestamp()))


def test_sync_kind_downloads_new_keeps_actual_and_replaces_outdated_files(tmp_path):
    """Немає в теці - завантажується під звичною назвою; той самий вміст -
    лишається; збережений РАНІШЕ за найновішу версію в Signal і з іншим
    вмістом - замінюється (на місці, з тією самою назвою), а попередня
    версія - у ".попередні версії"; файли лише в теці - не чіпаються; минулі
    дати без рапорту ніде - у списку відсутніх."""
    report_kind, _schedule_kind = _kinds(tmp_path)
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    (report_dir / "16.10.2026 - щоденний рапорт.docx").write_bytes(b"same")
    (report_dir / "17.10.2026_-_щоденний_рапорт.docx").write_bytes(b"old")  # інша назва, та сама дата
    (report_dir / "19.10.2026 - щоденний рапорт.docx").write_bytes(b"local only")
    for name in ("16.10.2026 - щоденний рапорт.docx", "17.10.2026_-_щоденний_рапорт.docx"):
        _set_mtime(report_dir / name, datetime(2026, 10, 1, 9, 0))  # збережено до надсилання версій нижче
    rows = [
        _row("16.10.2026 - щоденний рапорт.docx", datetime(2026, 10, 16, 16), b"same"),
        _row("17.10.2026 - щоденний рапорт.docx", datetime(2026, 10, 17, 16), b"new version"),
        _row("18.10.2026_-_щоденний_рапорт.docx", datetime(2026, 10, 18, 16), b"fresh"),
        _row("20.10.2026 - щоденний рапорт.docx", datetime(2026, 10, 20, 16), None),  # не розшифрувалось
    ]

    result = sr.sync_kind(report_kind, rows, _decrypt, _TODAY, now=datetime(2026, 10, 25, 12, 0, 0))

    assert result.actual == [date(2026, 10, 16)]
    assert result.updated == [date(2026, 10, 17)]
    assert result.new == [date(2026, 10, 18)]
    assert result.failed == ["20.10.2026 - щоденний рапорт.docx"]
    assert (report_dir / "17.10.2026_-_щоденний_рапорт.docx").read_bytes() == b"new version"
    assert (report_dir / "18.10.2026 - щоденний рапорт.docx").read_bytes() == b"fresh"
    assert (report_dir / "19.10.2026 - щоденний рапорт.docx").read_bytes() == b"local only"
    backups = os.listdir(report_dir / ".попередні версії")
    assert backups == ["17.10.2026_-_щоденний_рапорт (замінено 25.10.2026 12-00-00).docx"]
    missing = dict(result.missing)
    assert missing[date(2026, 9, 30)] == "30.09.2026 - щоденний рапорт.docx"
    assert date(2026, 10, 21) in missing and _TODAY in missing  # "з попередніх до поточної" - сьогодні теж
    assert date(2026, 10, 19) not in missing and date(2026, 10, 26) not in missing


def test_sync_kind_keeps_a_local_file_saved_after_the_newest_signal_version(tmp_path):
    """Реальний випадок: рапорти, виправлені вручну ПІСЛЯ завантаження з
    Signal, не збігаються з жодною версією там - але збережені пізніше за
    найновішу, тож актуальні й НЕ перезаписуються (навіть не завантажуються)."""
    report_kind, _schedule_kind = _kinds(tmp_path)
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    local = report_dir / "17.10.2026 - щоденний рапорт.docx"
    local.write_bytes(b"corrected by hand")
    _set_mtime(local, datetime(2026, 10, 18, 19, 43))
    rows = [_row("17.10.2026 - щоденний рапорт.docx", datetime(2026, 10, 17, 16, 51), None)]  # None - не завантажувати

    result = sr.sync_kind(report_kind, rows, _decrypt, _TODAY)

    assert (result.actual, result.updated, result.failed) == ([date(2026, 10, 17)], [], [])
    assert local.read_bytes() == b"corrected by hand"
    assert not (report_dir / ".попередні версії").exists()


def test_sync_kind_replaces_an_exact_copy_of_an_older_version_saved_later(tmp_path):
    """Збережено не ту версію: точна копія попередньої версії з Signal - не
    актуальна, хоч і збережена пізніше за найновішу; замінюється найновішою."""
    report_kind, _schedule_kind = _kinds(tmp_path)
    report_dir = tmp_path / "report"
    report_dir.mkdir()
    local = report_dir / "17.10.2026 - щоденний рапорт.docx"
    local.write_bytes(b"v1")
    _set_mtime(local, datetime(2026, 10, 18, 9, 0))
    rows = [
        _row("17.10.2026 - щоденний рапорт.docx", datetime(2026, 10, 17, 13, 30), b"v1"),
        _row("17.10.2026 - щоденний рапорт.docx", datetime(2026, 10, 17, 16, 51), b"v2"),
    ]

    result = sr.sync_kind(report_kind, rows, _decrypt, _TODAY, now=datetime(2026, 10, 25, 12, 0, 0))

    assert (result.actual, result.updated) == ([], [date(2026, 10, 17)])
    assert local.read_bytes() == b"v2"
    assert os.listdir(report_dir / ".попередні версії") == [
        "17.10.2026 - щоденний рапорт (замінено 25.10.2026 12-00-00).docx",
    ]


def test_missing_schedules_are_listed_with_the_expected_file_name(tmp_path):
    """Рішення користувача: відомості - теж дата й назва файлу
    за кожен відсутній день від 1-го числа до сьогодні; назва - за зразком
    найпізнішого відомого файлу (у Signal чи в теці) з потрібною датою."""
    today = date(2026, 10, 4)
    _report_kind, schedule_kind = sr._kinds(today, str(tmp_path / "report"), str(tmp_path / "schedule"))
    schedule_dir = tmp_path / "schedule"
    schedule_dir.mkdir()
    (schedule_dir / _SCHEDULE.format("01.10.2026")).write_bytes(b"x")
    rows = [_row(_SCHEDULE.format("03.10.2026"), datetime(2026, 10, 2, 18), b"y")]

    result = sr.sync_kind(schedule_kind, rows, _decrypt, today)

    assert result.missing == [
        (date(2026, 10, 2), _SCHEDULE.format("02.10.2026")),
        (date(2026, 10, 4), _SCHEDULE.format("04.10.2026")),
    ]


def test_with_date_keeps_the_year_style():
    assert sr._with_date("Звіт 07.10.26.xlsx", date(2026, 11, 2)) == "Звіт 02.11.26.xlsx"
    assert sr._with_date("Звіт 07.10.2026 (07.10.2026).xlsx", date(2026, 11, 2)) == "Звіт 02.11.2026 (07.10.2026).xlsx"


def test_generators_do_not_see_replaced_versions(tmp_path):
    """Попередні версії лежать у підтеці з крапкою - glob("*") генераторів
    (тут - вибір найновішої відомості) її не бачить."""
    schedule_dir = tmp_path / "schedule"
    hidden = schedule_dir / ".попередні версії"
    hidden.mkdir(parents=True)
    (schedule_dir / _SCHEDULE.format("20.10.2026")).write_bytes(b"x")
    (hidden / _SCHEDULE.format("25.10.2026")).write_bytes(b"x")

    latest = latest_schedule_file(str(schedule_dir), 2026, 10)
    assert latest is not None and os.path.basename(latest) == _SCHEDULE.format("20.10.2026")


def test_group_attachments_reads_only_downloaded_attachments_of_the_group():
    conn = sqlite3.connect(":memory:")
    conn.execute(
        "CREATE TABLE message_attachments (conversationId TEXT, fileName TEXT, sentAt INTEGER, path TEXT, "
        "localKey TEXT, size INTEGER, version INTEGER, editHistoryIndex INTEGER, error TEXT)"
    )
    conn.executemany("INSERT INTO message_attachments VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", [
        ("cid", "a.docx", 1, "p", "k", 1, 2, -1, None),
        ("cid", "b.docx", 1, None, "k", 1, 2, -1, None),  # ще не завантажено
        ("cid", "c.docx", 1, "p", "k", 1, 2, -1, "помилка"),
        ("other", "d.docx", 1, "p", "k", 1, 2, -1, None),
    ])

    rows = sr._group_attachments(conn.cursor(), "cid")

    assert [row["fileName"] for row in rows] == ["a.docx"]
    assert rows[0]["editHistoryIndex"] == -1


class _FakeDb:
    def __init__(self):
        self.closed = False

    def close(self):
        self.closed = True


def test_sync_reports_and_schedules_needs_a_configured_and_existing_group(monkeypatch):
    with pytest.raises(SignalFetchError, match="REPORTS_SIGNAL_GROUP_NAME"):
        sr.sync_reports_and_schedules(group_name=None)

    db = _FakeDb()
    monkeypatch.setattr(sr.signal_bchs, "_signal_source_dir", lambda: "src")
    monkeypatch.setattr(sr.signal_bchs, "_open_db", lambda source_dir: (db, object()))
    monkeypatch.setattr(sr.signal_bchs, "_find_conversation_id", lambda cursor, name: None)
    with pytest.raises(SignalFetchError, match="не знайдено"):
        sr.sync_reports_and_schedules(group_name="група")
    assert db.closed


def test_sync_reports_and_schedules_fills_both_folders(tmp_path, monkeypatch):
    rows = [
        _row("24.10.2026 - щоденний рапорт.docx", datetime(2026, 10, 24, 13), b"report"),
        _row(_SCHEDULE.format("25.10.2026"), datetime(2026, 10, 25, 11), b"schedule"),
    ]
    monkeypatch.setattr(sr.signal_bchs, "_signal_source_dir", lambda: "src")
    monkeypatch.setattr(sr.signal_bchs, "_open_db", lambda source_dir: (_FakeDb(), object()))
    monkeypatch.setattr(sr.signal_bchs, "_find_conversation_id", lambda cursor, name: "cid")
    monkeypatch.setattr(sr, "_group_attachments", lambda cursor, cid: rows)
    monkeypatch.setattr(sr.signal_bchs, "_decrypt_to", lambda source_dir, row, path: _decrypt(row, path))

    reports, schedules = sr.sync_reports_and_schedules(
        group_name="група", today=_TODAY, report_dir=str(tmp_path / "report"), schedule_dir=str(tmp_path / "schedule"),
    )

    assert (reports.label, reports.new, schedules.label, schedules.new) == (
        "рапорти", [date(2026, 10, 24)], "відомості ВОП-РОП", [date(2026, 10, 25)],
    )
    assert (tmp_path / "report" / "24.10.2026 - щоденний рапорт.docx").read_bytes() == b"report"
    assert (tmp_path / "schedule" / _SCHEDULE.format("25.10.2026")).read_bytes() == b"schedule"
