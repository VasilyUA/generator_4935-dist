import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from get_normalized_data import get_active_data as m
from constants import (
    UNIT_BRIGADE, UNIT_BATTALION, UNIT_COMPANY_ONE, UNIT_COMPANY_TWO,
    UNIT_COMPANY_DSHR, UNIT_COMPANY_ARTILLERY, UNIT_COMPANY_RECONNAISSANCE,
)


# -------------------------
# get_normalized_data
# -------------------------
def test_get_normalized_data_parses_date_and_time():
    data = [f"15:49 21.04.2026 {UNIT_COMPANY_ONE} {UNIT_BRIGADE} знищено ворога"]
    result = m.get_normalized_data(data)
    assert len(result) == 1
    assert result[0]["date"] == datetime(2026, 4, 21, 15, 49)
    assert UNIT_COMPANY_ONE in result[0]["text"]


def test_get_normalized_data_accepts_dot_time_separator():
    data = ["15.49 21.04.2026 текст"]
    result = m.get_normalized_data(data)
    assert result[0]["date"] == datetime(2026, 4, 21, 15, 49)


def test_get_normalized_data_skips_lines_without_date():
    data = ["без дати взагалі", ""]
    assert m.get_normalized_data(data) == []


def test_get_normalized_data_sorts_by_date():
    data = ["10:00 21.04.2026 другий", "09:00 21.04.2026 перший"]
    result = m.get_normalized_data(data)
    assert [r["text"].split()[-1] for r in result] == ["перший", "другий"]


def test_get_normalized_data_skips_line_with_invalid_calendar_date():
    # регекс приймає "32.13.2026" (правильна форма dd.mm.yyyy), але
    # datetime.strptime падає - такого дня/місяця не існує
    data = ["15:49 32.13.2026 текст"]
    assert m.get_normalized_data(data) == []


# -------------------------
# _inject_missing_date (мін. батр. пише лише час, без дати в тексті)
# -------------------------
def test_inject_missing_date_adds_date_from_signal_timestamp():
    body = f"13:45 мін. батр. {UNIT_BATTALION} {UNIT_BRIGADE} Витрати: 5,56х45мм-150 шт."
    result = m._inject_missing_date(body, "2026-07-25T13:57:08.534000")
    assert result == f"13:45 25.07.2026 мін. батр. {UNIT_BATTALION} {UNIT_BRIGADE} Витрати: 5,56х45мм-150 шт."


def test_inject_missing_date_leaves_body_unchanged_when_date_already_present():
    body = f"09:23 25.07.2026 {UNIT_COMPANY_ONE} {UNIT_BATTALION} {UNIT_BRIGADE}"
    assert m._inject_missing_date(body, "2026-07-25T09:30:00") == body


def test_inject_missing_date_leaves_body_unchanged_without_leading_time():
    body = "без часу на початку рядка"
    assert m._inject_missing_date(body, "2026-07-25T09:30:00") == body


def test_inject_missing_date_leaves_body_unchanged_without_msg_date():
    body = "13:45 мін. батр. текст"
    assert m._inject_missing_date(body, None) == body


def test_inject_missing_date_leaves_body_unchanged_when_msg_date_unparseable():
    body = "13:45 мін. батр. текст"
    assert m._inject_missing_date(body, "не дата взагалі") == body


def test_get_active_data_skips_messages_with_empty_body(monkeypatch):
    messages = [{"date": "2026-07-25T09:00:00", "body": "   "}]
    monkeypatch.setattr(m, "get_signal_data", lambda path: messages)
    monkeypatch.setattr(m, "get_filtered_messages_for_date", lambda msgs, hour, date: msgs)

    assert m.get_active_data(18, "25.07.2026") == []


def test_get_active_data_recovers_dateless_minbatr_message(monkeypatch):
    # регресія: "мін. батр." у реальних повідомленнях ніколи не пише дату в
    # тексті - без підстановки дати з Signal-мітки get_normalized_data просто
    # пропускав би такі повідомлення, і мінбатр ніколи б не рахувався
    messages = [{
        "date": "2026-07-25T13:57:08.534000",
        "body": f'13:45 мін. батр. {UNIT_BATTALION} {UNIT_BRIGADE} Витрати: 5,56х45 мм-150 шт.',
    }]
    monkeypatch.setattr(m, "get_signal_data", lambda path: messages)
    monkeypatch.setattr(m, "get_filtered_messages_for_date", lambda msgs, hour, date: msgs)

    result = m.get_active_data(20, "25.07.2026")
    assert len(result) == 1
    assert result[0]["date"] == datetime(2026, 7, 25, 13, 45)
    assert m.calculate(result, UNIT_COMPANY_ARTILLERY) == "5.56х45мм – 150 шт."


# -------------------------
# _norm (формат калібру - кирилична х, крапка, без пробілу перед "мм")
# -------------------------
def test_norm_formats_caliber_cyrillic_period_no_space():
    assert m._norm("5,56x45мм") == "5.56х45мм"


def test_norm_only_handles_latin_x_input():
    # _norm отримує вхід виключно з регекспів calculate() (завжди латинська "x")
    # - кирилична "х" на вході не є реальним сценарієм і не перетворюється.
    assert m._norm("5.56х45мм") == "5.56х45"


def test_norm_passthrough_when_no_x_present():
    assert m._norm("12") == "12"


# -------------------------
# _unit_in_text
# -------------------------
def test_unit_in_text_direct_match():
    assert m._unit_in_text(UNIT_COMPANY_DSHR, f"{UNIT_COMPANY_DSHR} {UNIT_BRIGADE} знищено") is True


def test_unit_in_text_compact_fallback_matches_no_space_variant():
    assert m._unit_in_text(UNIT_COMPANY_ONE, f"{UNIT_COMPANY_ONE.replace(' ', '')} {UNIT_BRIGADE} знищено") is True


def test_unit_in_text_no_match():
    assert m._unit_in_text("тос", f"{UNIT_COMPANY_ONE} {UNIT_BRIGADE} знищено") is False


def test_unit_in_text_matches_dotted_abbreviation():
    # реальні повідомлення пишуть UNIT_COMPANY_ARTILLERY як "мін. батр." (крапка + пробіл)
    assert m._unit_in_text(UNIT_COMPANY_ARTILLERY, f"13:45 мін. батр. {UNIT_BATTALION} {UNIT_BRIGADE}") is True


# -------------------------
# calculate - основний сценарій підрахунку витрати БК по калібру
# -------------------------
def test_calculate_extracts_qty_and_caliber_with_dash_pattern():
    data = [{"text": f'5,56x45 мм - 300 набоїв. {UNIT_COMPANY_ONE} {UNIT_BRIGADE} по ворогу.'}]
    result = m.calculate(data, UNIT_COMPANY_ONE)
    assert result == "5.56х45мм – 300 шт."


def test_calculate_extracts_qty_and_caliber_qty_first_pattern():
    data = [{"text": f'{UNIT_COMPANY_ONE} {UNIT_BRIGADE} використано 360 шт (5,56x45мм) по ворогу.'}]
    result = m.calculate(data, UNIT_COMPANY_ONE)
    assert result == "5.56х45мм – 360 шт."


def test_calculate_sums_across_multiple_messages():
    data = [
        {"text": f'{UNIT_COMPANY_ONE} {UNIT_BRIGADE} 100 шт (5,56x45мм).'},
        {"text": f'{UNIT_COMPANY_ONE} {UNIT_BRIGADE} 50 шт (5,56x45мм).'},
    ]
    result = m.calculate(data, UNIT_COMPANY_ONE)
    assert result == "5.56х45мм – 150 шт."


def test_calculate_sorts_multiple_calibers_by_priority():
    data = [{"text": f'{UNIT_COMPANY_ONE} {UNIT_BRIGADE} 10 шт (7,62x39мм), 20 шт (5,56x45мм).'}]
    result = m.calculate(data, UNIT_COMPANY_ONE)
    assert result == "5.56х45мм – 20 шт., 7.62х39мм – 10 шт."


def test_calculate_no_matching_unit_returns_empty_string():
    data = [{"text": f'{UNIT_COMPANY_ONE} {UNIT_BRIGADE} 100 шт (5,56x45мм).'}]
    assert m.calculate(data, "тос") == ""


def test_calculate_empty_data_returns_empty_string():
    assert m.calculate([], UNIT_COMPANY_ONE) == ""


def test_calculate_dshr_excludes_rvp_lines():
    # рядки РВП містять "{UNIT_COMPANY_DSHR}/" + UNIT_BATTALION без пробілу - їх не можна плутати з реальними даними UNIT_COMPANY_DSHR
    data = [{"text": f'{UNIT_COMPANY_DSHR}/{UNIT_BATTALION.replace(' ', '')} рвп {UNIT_BRIGADE} 100 шт (5,56x45мм).'}]
    assert m.calculate(data, UNIT_COMPANY_DSHR) == ""


def test_calculate_reconnaissance_excludes_rvp_substring_collision():
    # UNIT_COMPANY_RECONNAISSANCE ("рв") - підрядок "рвп", без винятку задвоїло б лічильник рвп в рв
    data = [{"text": f'рвп {UNIT_BRIGADE} 100 шт (5,56x45мм).'}]
    assert m.calculate(data, UNIT_COMPANY_RECONNAISSANCE) == ""


def test_calculate_1rmp_excludes_2rmp_lines():
    data = [{"text": f'{UNIT_COMPANY_TWO.replace(' ', '')} {UNIT_BRIGADE} 100 шт (5,56x45мм).'}]
    assert m.calculate(data, UNIT_COMPANY_ONE) == ""


def test_calculate_1rmp_excludes_lines_that_also_mention_2rmp():
    # рядок, де UNIT_COMPANY_ONE ЗНАЙДЕНО в тексті, але поруч є й UNIT_COMPANY_TWO (без пробілу) - все одно
    # виключається, щоб не задвоїти з витратою UNIT_COMPANY_TWO
    data = [{"text": f'{UNIT_COMPANY_ONE.replace(' ', '')} {UNIT_COMPANY_TWO.replace(' ', '')} {UNIT_BRIGADE} 100 шт (5,56x45мм).'}]
    assert m.calculate(data, UNIT_COMPANY_ONE) == ""


def test_calculate_2rmp_excluded_when_only_dotted_abbreviation_present():
    # _unit_in_text знаходить підрозділ через compact-фолбек (крапка теж
    # прибирається), але власна перевірка calculate() прибирає лише пробіли -
    # тому "2.рмп" (з крапкою) не визнається як UNIT_COMPANY_TWO.replace(' ', '') для цієї перевірки
    data = [{"text": f'2.рмп {UNIT_BRIGADE} 100 шт (5,56x45мм).'}]
    assert m.calculate(data, UNIT_COMPANY_TWO) == ""


def test_calculate_matches_unit_without_space_via_compact_fallback():
    data = [{"text": f'{UNIT_COMPANY_ONE.replace(' ', '')} {UNIT_BRIGADE} 100 шт (5,56x45мм).'}]
    assert m.calculate(data, UNIT_COMPANY_ONE) == "5.56х45мм – 100 шт."


def test_calculate_12x70_special_case():
    data = [{"text": f'{UNIT_COMPANY_ONE} {UNIT_BRIGADE} використано 12x70 дробовий, 15 набоїв.'}]
    assert m.calculate(data, UNIT_COMPANY_ONE) == "12х70 – 15 шт."


def test_calculate_12_caliber_special_case():
    data = [{"text": f'{UNIT_COMPANY_ONE} {UNIT_BRIGADE} 12 калібру - використано 20 шт.'}]
    assert m.calculate(data, UNIT_COMPANY_ONE) == "12 – 20 шт."
