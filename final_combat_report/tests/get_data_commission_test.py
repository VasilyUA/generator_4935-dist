import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from get_normalized_data import get_data_commission as m


# -------------------------
# _strip_service_prefixes
# -------------------------
def test_strip_service_prefixes_removes_propyshit():
    assert m._strip_service_prefixes("Пропишіть: текст") == "текст"


def test_strip_service_prefixes_case_insensitive_no_colon():
    assert m._strip_service_prefixes("пропишіть текст") == "текст"


def test_strip_service_prefixes_leaves_other_text_untouched():
    assert m._strip_service_prefixes("звичайний текст") == "звичайний текст"


def test_strip_service_prefixes_non_string_passthrough():
    assert m._strip_service_prefixes(None) is None


# -------------------------
# _parse_message_date
# -------------------------
def test_parse_message_date_datetime_passthrough():
    dt = datetime(2026, 7, 25, 9, 49)
    assert m._parse_message_date(dt) is dt


def test_parse_message_date_iso_string():
    assert m._parse_message_date("2026-07-25T09:49:00") == datetime(2026, 7, 25, 9, 49)


def test_parse_message_date_invalid_string_returns_none():
    assert m._parse_message_date("не дата") is None


def test_parse_message_date_other_type_returns_none():
    assert m._parse_message_date(12345) is None


# -------------------------
# get_normalized_data
# -------------------------
def test_get_normalized_data_prepends_time_when_missing():
    messages = [{"body": "перевірка КСП", "date": "2026-07-25T09:49:00"}]
    assert m.get_normalized_data(messages) == ["09:49 25.07.2026 перевірка КСП"]


def test_get_normalized_data_does_not_duplicate_existing_time_prefix():
    messages = [{"body": "09:49 перевірка КСП", "date": "2026-07-25T09:49:00"}]
    assert m.get_normalized_data(messages) == ["09:49 перевірка КСП"]


def test_get_normalized_data_strips_service_prefix_before_output():
    messages = [{"body": "Пропишіть: перевірка КСП", "date": "2026-07-25T09:49:00"}]
    assert m.get_normalized_data(messages) == ["09:49 25.07.2026 перевірка КСП"]


def test_get_normalized_data_drops_messages_without_letters():
    messages = [{"body": "12345", "date": "2026-07-25T09:49:00"}]
    assert m.get_normalized_data(messages) == []


def test_get_normalized_data_sorts_by_date():
    messages = [
        {"body": "другий", "date": "2026-07-25T10:00:00"},
        {"body": "перший", "date": "2026-07-25T09:00:00"},
    ]
    result = m.get_normalized_data(messages)
    assert "перший" in result[0]
    assert "другий" in result[1]


# -------------------------
# get_data_commission
# -------------------------
def test_get_data_commission_reads_filters_and_normalizes(monkeypatch):
    raw_messages = [{"body": "перевірка КСП", "date": "2026-07-25T09:49:00"}]
    monkeypatch.setattr(m, "get_signal_data", lambda path: raw_messages)
    monkeypatch.setattr(m, "get_filtered_messages_for_date", lambda msgs, hour, date: msgs)

    result = m.get_data_commission(18, "25.07.2026")

    assert result == ["09:49 25.07.2026 перевірка КСП"]
