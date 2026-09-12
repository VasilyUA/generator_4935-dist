import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import pytest

import filters as m
from constants import UNIT_BATTALION, UNIT_BRIGADE_MIXED_CASE


# -------------------------
# get_filtered_messages_for_date
# -------------------------
def test_get_filtered_messages_for_date_keeps_only_window():
    messages = [
        {"date": "2026-07-24T17:59:00", "text": "до вікна"},
        {"date": "2026-07-24T18:00:00", "text": "початок вікна"},
        {"date": "2026-07-25T09:00:00", "text": "всередині"},
        {"date": "2026-07-25T17:59:59", "text": "кінець вікна"},
        {"date": "2026-07-25T18:00:00", "text": "після вікна"},
    ]
    result = m.get_filtered_messages_for_date(messages, 18, "25.07.2026")
    texts = [x["text"] for x in result]
    assert texts == ["початок вікна", "всередині", "кінець вікна"]


def test_get_filtered_messages_for_date_sorted_by_date():
    messages = [
        {"date": "2026-07-25T10:00:00", "text": "другий"},
        {"date": "2026-07-25T09:00:00", "text": "перший"},
    ]
    result = m.get_filtered_messages_for_date(messages, 18, "25.07.2026")
    assert [x["text"] for x in result] == ["перший", "другий"]


def test_get_filtered_messages_for_date_hour_24_is_calendar_day():
    messages = [
        {"date": "2026-07-24T23:59:59", "text": "вчора"},
        {"date": "2026-07-25T00:00:00", "text": "сьогодні початок"},
        {"date": "2026-07-25T23:59:59", "text": "сьогодні кінець"},
    ]
    result = m.get_filtered_messages_for_date(messages, 24, "25.07.2026")
    texts = [x["text"] for x in result]
    assert texts == ["сьогодні початок", "сьогодні кінець"]


def test_get_filtered_messages_for_date_defaults_to_today_when_date_is_none():
    # report_date=None -> сьогодні (datetime.now()); hour=24 означає календарну
    # добу [сьогодні 00:00, завтра 00:00) - поточний момент завжди в цьому вікні
    messages = [{"date": datetime.now().isoformat(), "text": "сьогодні"}]
    result = m.get_filtered_messages_for_date(messages, 24, None)
    assert [x["text"] for x in result] == ["сьогодні"]


def test_get_filtered_messages_for_date_accepts_datetime_objects():
    messages = [{"date": datetime(2026, 7, 25, 9, 0), "text": "тест"}]
    result = m.get_filtered_messages_for_date(messages, 18, "25.07.2026")
    assert len(result) == 1


def test_get_filtered_messages_for_date_skips_messages_without_date():
    messages = [{"text": "без дати"}]
    assert m.get_filtered_messages_for_date(messages, 18, "25.07.2026") == []


def test_get_filtered_messages_for_date_invalid_format_raises():
    with pytest.raises(ValueError):
        m.get_filtered_messages_for_date([], 18, "not-a-date")


# -------------------------
# get_filtered_body_from_cost_messages
# -------------------------
def test_get_filtered_body_from_cost_messages_requires_spend_keyword_by_default():
    messages = [
        {"body": "Результат: уражено"},
        {"body": "Витрата: щось"},
        {"body": "Знищено: щось"},
        {"body": "просто текст"},
    ]
    result = m.get_filtered_body_from_cost_messages(messages)
    assert result == ["Результат: уражено", "Витрата: щось", "Знищено: щось"]


def test_get_filtered_body_from_cost_messages_need_spend_false_returns_all_nonempty():
    messages = [{"body": "просто текст"}, {"body": ""}]
    result = m.get_filtered_body_from_cost_messages(messages, need_spend=False)
    assert result == ["просто текст"]


# -------------------------
# get_filtered_body_from_logistics_uav_messages / get_filtered_body_from_loss_uav_messages
# -------------------------
def test_get_filtered_body_from_logistics_uav_messages_requires_both_keywords():
    messages = [
        {"body": "Звіт по логістика доставлено"},
        {"body": "Звіт без потрібного слова"},
        {"body": "логістика без цього слова"},  # "звіт" немає навіть як підрядок
    ]
    result = m.get_filtered_body_from_logistics_uav_messages(messages)
    assert result == ["Звіт по логістика доставлено"]


def test_get_filtered_body_from_loss_uav_messages_requires_both_keywords():
    messages = [
        {"body": "Виліт 5 логістичне забезпечення"},
        {"body": "Виліт без потрібного слова"},
    ]
    result = m.get_filtered_body_from_loss_uav_messages(messages)
    assert result == ["Виліт 5 логістичне забезпечення"]


# -------------------------
# get_filtered_body_from_kasper_logistics_messages
# -------------------------
def test_get_filtered_body_from_kasper_logistics_messages_requires_both_keywords():
    # регресія: екіпаж "Каспер" звітує форматом "...бомбер-вампір... працював
    # на логістику..." - не містить ні "логістика"+"звіт", ні "логістичне",
    # тому без цього фільтра вильоти повністю випадали з підрахунку
    messages = [
        {"body": f"Звіт {UNIT_BATTALION} {UNIT_BRIGADE_MIXED_CASE}: бомбер-вампір(Старлінк) працював на логістику."},
        {"body": "бомбер-вампір без другого слова"},
        {"body": "працював на логістику без бомбера"},
    ]
    result = m.get_filtered_body_from_kasper_logistics_messages(messages)
    assert result == [f"Звіт {UNIT_BATTALION} {UNIT_BRIGADE_MIXED_CASE}: бомбер-вампір(Старлінк) працював на логістику."]
