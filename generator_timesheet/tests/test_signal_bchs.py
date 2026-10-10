import sqlite3
from datetime import date, datetime

import pytest

import content.signal_bchs as signal_bchs
from content.signal_bchs import SignalFetchError, _candidate_rows, _contains_marker, _dedupe_keep_latest, _has_supported_extension, _save_attachment, _sent_on_date, fetch_bchs_files


def _epoch_ms(dt):
    return int(dt.timestamp() * 1000)


# --- _contains_marker --------------------------------------------------

@pytest.mark.parametrize("file_name, expected", [
    ("БЧС ВЗ 01.10.2026.xlsx", True),
    ("бчс_1РМП 01.10.26.xlsx", True),  # нижній регістр - теж збіг
    ("МП 01.10.26. новий xlsx.xlsx", False),  # реальний випадок: немає "БЧС" узагалі
    (None, False),
])
def test_contains_marker(file_name, expected):
    assert _contains_marker(file_name, "БЧС") is expected


# --- _has_supported_extension ---------------------------------------------

@pytest.mark.parametrize("file_name, expected", [
    ("БЧС ВЗ 01.10.2026.xlsx", True),
    ("БЧС ВЗ 01.10.2026.xlsm", True),
    ("БЧС ВЗ 01.10.2026.xls", True),
    ("БЧС ВЗ 01.10.2026.docx", False),
    (None, False),
])
def test_has_supported_extension_matches_information_unit_reader(file_name, expected):
    """Ті самі розширення, що й information_unit_reader._SUPPORTED_EXTENSIONS
    (.xlsx/.xlsm/.xls) - не окремий список, який довелось би синхронізувати."""
    assert _has_supported_extension(file_name) is expected


# --- _sent_on_date -------------------------------------------------------

def test_sent_on_date_matches_the_local_calendar_day():
    sent_at = _epoch_ms(datetime(2026, 10, 1, 23, 59))

    assert _sent_on_date(sent_at, date(2026, 10, 1)) is True
    assert _sent_on_date(sent_at, date(2026, 10, 2)) is False


def test_sent_on_date_handles_a_missing_timestamp():
    assert _sent_on_date(None, date(2026, 10, 1)) is False


# --- _dedupe_keep_latest --------------------------------------------------

def test_dedupe_keep_latest_keeps_the_later_duplicate_by_exact_name():
    """Реальний випадок: повідомлення, відредаговане в Signal, лишає кілька
    рядків message_attachments з ТОЧНО тією самою назвою файлу - береться та
    з найбільшим sentAt, незалежно від порядку у вхідному списку."""
    older = {"fileName": "БЧС ВЗ 01.10.2026.xlsx", "sentAt": 100}
    newer = {"fileName": "БЧС ВЗ 01.10.2026.xlsx", "sentAt": 200}

    assert _dedupe_keep_latest([older, newer]) == [newer]
    assert _dedupe_keep_latest([newer, older]) == [newer]


def test_dedupe_keep_latest_keeps_distinct_names_separately():
    """Різні підрозділи - різні назви файлів за ТУ САМУ дату - не дублікати
    одне одного, лишаються ВСІ."""
    one = {"fileName": "БЧС ВЗ 01.10.2026.xlsx", "sentAt": 100}
    two = {"fileName": "БЧС 1 РМП 01.10.2026.xlsx", "sentAt": 100}

    kept = _dedupe_keep_latest([one, two])

    assert {row["fileName"] for row in kept} == {one["fileName"], two["fileName"]}


def test_dedupe_keep_latest_with_no_rows():
    assert _dedupe_keep_latest([]) == []


# --- _candidate_rows (справжній SQL, звичайний sqlite3 - та сама підмножина колонок) ---

def _make_db(rows, conversation_id="cid"):
    """rows - [(fileName, sentAt, path, localKey, size, version, error), ...]."""
    conn = sqlite3.connect(":memory:")
    conn.execute(
        "CREATE TABLE message_attachments (conversationId TEXT, fileName TEXT, sentAt INTEGER, "
        "path TEXT, localKey TEXT, size INTEGER, version INTEGER, error TEXT)"
    )
    conn.executemany(
        "INSERT INTO message_attachments VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        [(conversation_id, *row) for row in rows],
    )
    conn.commit()
    return conn


def test_candidate_rows_filters_by_marker_date_and_extension():
    target = date(2026, 10, 1)
    same_day = _epoch_ms(datetime(2026, 10, 1, 10, 0))
    other_day = _epoch_ms(datetime(2026, 10, 2, 10, 0))
    conn = _make_db([
        ("БЧС ВЗ 01.10.2026.xlsx", same_day, "a/b", "key1", 10, 2, None),
        ("БЧС ВЗ 02.10.2026.xlsx", other_day, "a/c", "key2", 10, 2, None),  # не той день
        ("Якийсь протокол 01.10.2026.docx", same_day, "a/d", "key3", 10, 2, None),  # не .xlsx
        ("Список особового складу 01.10.2026.xlsx", same_day, "a/e", "key4", 10, 2, None),  # немає "БЧС"
    ])

    rows = _candidate_rows(conn.cursor(), "cid", target)

    assert [row["fileName"] for row in rows] == ["БЧС ВЗ 01.10.2026.xlsx"]


def test_candidate_rows_excludes_errored_or_not_yet_downloaded_attachments():
    target = date(2026, 10, 1)
    same_day = _epoch_ms(datetime(2026, 10, 1, 10, 0))
    conn = _make_db([
        ("БЧС ВЗ 01.10.2026.xlsx", same_day, None, "key1", 10, 2, None),  # path NULL - ще не завантажено
        ("БЧС 1 РМП 01.10.2026.xlsx", same_day, "a/b", "key2", 10, 2, "щось пішло не так"),  # error
        ("БЧС ІСВ 01.10.2026.xlsx", same_day, "a/c", "key3", 10, 2, None),  # ОК
    ])

    rows = _candidate_rows(conn.cursor(), "cid", target)

    assert [row["fileName"] for row in rows] == ["БЧС ІСВ 01.10.2026.xlsx"]


def test_candidate_rows_only_looks_at_the_given_conversation():
    target = date(2026, 10, 1)
    same_day = _epoch_ms(datetime(2026, 10, 1, 10, 0))
    conn = _make_db([("БЧС ВЗ 01.10.2026.xlsx", same_day, "a/b", "key1", 10, 2, None)], conversation_id="other-cid")

    rows = _candidate_rows(conn.cursor(), "cid", target)

    assert rows == []


# --- _save_attachment -----------------------------------------------------

def test_save_attachment_plain_copies_a_pre_v2_attachment(tmp_path):
    source_dir = tmp_path / "signal"
    attachments_dir = source_dir / "attachments.noindex" / "sub"
    attachments_dir.mkdir(parents=True)
    (attachments_dir / "raw.bin").write_bytes(b"xlsx bytes")
    dest_dir = tmp_path / "dest"
    dest_dir.mkdir()
    row = {"fileName": "БЧС ВЗ 01.10.2026.xlsx", "path": "sub/raw.bin", "version": 1, "size": 10, "localKey": ""}

    assert _save_attachment(str(source_dir), str(dest_dir), row) is True
    assert (dest_dir / row["fileName"]).read_bytes() == b"xlsx bytes"


def test_save_attachment_decrypts_a_v2_attachment(tmp_path, monkeypatch):
    dest_dir = tmp_path / "dest"
    dest_dir.mkdir()
    row = {"fileName": "БЧС ВЗ 01.10.2026.xlsx", "path": "raw.bin", "version": 2, "size": 10, "localKey": "k"}

    def _fake_decrypt(att, src_path, dst_path):
        with open(dst_path, "wb") as fh:
            fh.write(b"decrypted")

    monkeypatch.setattr(signal_bchs, "decrypt_attachment", _fake_decrypt)

    assert _save_attachment(str(tmp_path / "signal"), str(dest_dir), row) is True
    assert (dest_dir / row["fileName"]).read_bytes() == b"decrypted"


def test_save_attachment_returns_false_without_raising_when_decrypt_fails(tmp_path, monkeypatch):
    row = {"fileName": "БЧС ВЗ 01.10.2026.xlsx", "path": "raw.bin", "version": 2, "size": 10, "localKey": "k"}
    monkeypatch.setattr(signal_bchs, "decrypt_attachment", lambda att, src_path, dst_path: (_ for _ in ()).throw(ValueError("MAC mismatch")))

    assert _save_attachment(str(tmp_path / "signal"), str(tmp_path / "dest"), row) is False


def test_save_attachment_returns_false_when_the_source_file_is_missing(tmp_path):
    row = {"fileName": "БЧС ВЗ 01.10.2026.xlsx", "path": "raw.bin", "version": 0, "size": 10, "localKey": ""}

    assert _save_attachment(str(tmp_path / "signal"), str(tmp_path / "dest"), row) is False


# --- _signal_source_dir ---------------------------------------------------

def test_signal_source_dir_returns_the_signal_folder_when_it_looks_installed(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    signal_dir = tmp_path / "Signal"
    signal_dir.mkdir()
    (signal_dir / "config.json").write_text("{}")

    assert signal_bchs._signal_source_dir() == str(signal_dir)


# --- fetch_bchs_files (помилки верхнього рівня) ---------------------------

def test_fetch_bchs_files_raises_when_appdata_is_not_set(monkeypatch):
    monkeypatch.delenv("APPDATA", raising=False)

    with pytest.raises(SignalFetchError, match="APPDATA"):
        fetch_bchs_files(date(2026, 10, 1), "dest")


def test_fetch_bchs_files_raises_when_signal_is_not_installed(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))  # немає Signal/config.json

    with pytest.raises(SignalFetchError, match="Signal Desktop"):
        fetch_bchs_files(date(2026, 10, 1), "dest")


def test_fetch_bchs_files_raises_when_the_group_is_not_found(monkeypatch):
    class _FakeDb:
        def close(self):
            pass

    monkeypatch.setattr(signal_bchs, "_signal_source_dir", lambda: "src")
    monkeypatch.setattr(signal_bchs, "_open_db", lambda source_dir: (_FakeDb(), object()))
    monkeypatch.setattr(signal_bchs, "_find_conversation_id", lambda cursor, group_name: None)

    with pytest.raises(SignalFetchError, match="не знайдено"):
        fetch_bchs_files(date(2026, 10, 1), "dest")


def test_fetch_bchs_files_closes_the_db_even_when_the_group_is_not_found(monkeypatch):
    closed = []

    class _FakeDb:
        def close(self):
            closed.append(True)

    monkeypatch.setattr(signal_bchs, "_signal_source_dir", lambda: "src")
    monkeypatch.setattr(signal_bchs, "_open_db", lambda source_dir: (_FakeDb(), object()))
    monkeypatch.setattr(signal_bchs, "_find_conversation_id", lambda cursor, group_name: None)

    with pytest.raises(SignalFetchError):
        fetch_bchs_files(date(2026, 10, 1), "dest")

    assert closed == [True]


def test_fetch_bchs_files_returns_only_the_rows_that_were_saved_successfully(monkeypatch):
    class _FakeDb:
        def close(self):
            pass

    rows = [
        {"fileName": "БЧС ВЗ 01.10.2026.xlsx", "sentAt": 1},
        {"fileName": "БЧС 1 РМП 01.10.2026.xlsx", "sentAt": 1},
    ]
    saved_calls = []

    def _fake_save(source_dir, dest_dir, row):
        saved_calls.append(row["fileName"])
        return row["fileName"] == "БЧС ВЗ 01.10.2026.xlsx"

    monkeypatch.setattr(signal_bchs, "_signal_source_dir", lambda: "src")
    monkeypatch.setattr(signal_bchs, "_open_db", lambda source_dir: (_FakeDb(), object()))
    monkeypatch.setattr(signal_bchs, "_find_conversation_id", lambda cursor, group_name: "cid")
    monkeypatch.setattr(signal_bchs, "_candidate_rows", lambda cursor, cid, target_date: rows)
    monkeypatch.setattr(signal_bchs, "_save_attachment", _fake_save)

    result = fetch_bchs_files(date(2026, 10, 1), "dest")

    assert result == ["БЧС ВЗ 01.10.2026.xlsx"]
    assert saved_calls == ["БЧС ВЗ 01.10.2026.xlsx", "БЧС 1 РМП 01.10.2026.xlsx"]
