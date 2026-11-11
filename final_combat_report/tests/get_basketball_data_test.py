import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from get_normalized_data import get_basketball_data as m
from constants import UNIT_BATTALION, UNIT_BRIGADE, UNIT_BRIGADE_MIXED_CASE


# -------------------------
# get_basketball_data
# -------------------------
def test_get_basketball_data_reads_filters_parses_and_normalizes(monkeypatch):
    raw_messages = [{"body": "Дата: 24.07.26\nЧас: 09:00\nВитрата: ОФ-843Б, 2шт. (120мм)\nРезультат: уражено"}]
    monkeypatch.setattr(m, "get_signal_data", lambda path: raw_messages)
    monkeypatch.setattr(m, "get_filtered_messages_for_date", lambda msgs, hour, date: msgs)

    result = m.get_basketball_data(18, "24.07.2026")

    assert len(result) == 1
    assert result[0]["spend"] == {"ОФ-843Б 120мм": 2}
    assert result[0]["date"] == "24.07.26"


# реальний формат щоденного донесення ВГрМ "Пежо" (перевірено на десятках
# повідомлень 07.07.2026-02.08.2026 - структура завжди однакова: подвійна
# пряма лапка без закриваючої навколо позивного позиції, "N БМП N ОБрМП"
# дублюється в кінці 1 рядка, координати МІЖ "Характер цілі:" і "Ціль:",
# опис району в дужках ОДРАЗУ за кодом цілі)
REAL_VGRM_TEXT = (
    f'ВГрМ 1.1""Пежо {UNIT_BRIGADE_MIXED_CASE}\n'
    f' {UNIT_BATTALION.upper()} {UNIT_BRIGADE_MIXED_CASE}\n'
    ' Дата: 02.08.26\n'
    ' Час: 12:10-12:13\n'
    ' Характер цілі: УКРИТТЯ\n'
    ' 37U CP 31696 29708\n'
    ' Ціль: 020840101\n'
    ' (в р-н н.п. Філія)\n'
    ' Витрата: ОФ-843Б, 2шт. зар.6, М12, 2шт. (120мм)\n'
    ' Результат: Заборона дій\n'
    ' Стрім: Без стріму  '
)


def test_get_parsed_extracts_all_fields_from_real_vgrm_format():
    record = m.get_parsed(REAL_VGRM_TEXT)
    assert record["ПОЗИЦІЯ"] == 'ВГрМ 1.1 "Пежо"'
    assert record["Підрозділ"] == f"{UNIT_BATTALION} {UNIT_BRIGADE}"
    assert record["Дата"] == "02.08.26"
    assert record["Час"] == "12:10-12:13"
    assert record["Характер цілі"] == "УКРИТТЯ 37U CP 31696 29708"
    assert record["Ціль"] == "020840101 (в р-н н.п. Філія)"
    assert record["Витрата"] == "ОФ-843Б 120мм. -2шт."
    assert record["spend"] == {"ОФ-843Б 120мм": 2}
    assert record["Результат"] == "Заборона дій"
    assert record["Стрим"] == "Без стріму"


def test_get_messages_matches_real_vgrm_report_line():
    record = m.get_parsed(REAL_VGRM_TEXT)
    text = m.get_messages(record)
    assert text == (
        f'12:10-12:13 02.08.26 ВГрМ 1.1 "Пежо" {UNIT_BATTALION} {UNIT_BRIGADE}. '
        'Характер цілі: УКРИТТЯ 37U CP 31696 29708. '
        'Ціль: 020840101 (в р-н н.п. Філія). '
        'Витрата: ОФ-843Б 120мм. -2шт. '
        'Результат: Заборона дій. '
        'Стрім: Без стріму.'
    )


def test_get_parsed_spend_ignores_charge_and_extracts_only_primary_munition():
    # реальний формат: перше "назва, Nшт." - міна, "зар.N" і друге "назва, Nшт."
    # - заряд (метальний), що йде РАЗОМ з тією ж міною й не рахується окремо
    record = m.get_parsed("Витрата: ОФ-843Б, 3шт. зар.6, М12, 3шт. (120мм)")
    assert record["spend"] == {"ОФ-843Б 120мм": 3}
    assert record["Витрата"] == "ОФ-843Б 120мм. -3шт."


def test_get_parsed_spend_unparseable_format_kept_for_display_only():
    # нерозпізнаний формат - лишається в тексті (для 3.3.6), але не в spend
    # (щоб не рахувати "1" для того, що насправді не вдалося розпарсити)
    record = m.get_parsed("Витрата: невідомий формат без числа")
    assert record["Витрата"] == "невідомий формат без числа"
    assert "spend" not in record


def test_get_parsed_stream_kept_as_is_when_not_empty():
    record = m.get_parsed("Стрім: https://example.com/live")
    assert record["Стрим"] == "https://example.com/live"


def test_get_parsed_stream_bez_strimu_kept_literally_not_rewritten():
    # "Без стріму" вже й так означає "немає трансляції" - на відміну від
    # УБпАК-формату тут не переписуємо на "Відсутній"
    record = m.get_parsed("Стрім: Без стріму")
    assert record["Стрим"] == "Без стріму"


def test_get_parsed_missing_fields_are_absent():
    record = m.get_parsed("Просто текст без міток")
    assert "Дата" not in record
    assert "Характер цілі" not in record
    assert "Ціль" not in record


def test_get_messages_formats_available_fields():
    text = m.get_messages({"Час": "09:00", "Дата": "24.07.26", "Ціль": "5", "Результат": "уражено"})
    assert "09:00" in text
    assert "24.07.26" in text
    assert "Ціль: 5." in text
    assert "Результат: уражено." in text


def test_get_messages_omits_missing_fields():
    text = m.get_messages({"Час": "09:00", "Дата": "24.07.26"})
    assert "Характер цілі" not in text
    assert "Ціль" not in text
    assert "Витрата" not in text
    assert "Результат" not in text
    assert "Стрім" not in text


def test_normalize_data_wraps_parsed_records():
    result = m.normalize_data([{"Дата": "24.07.26", "Час": "09:00", "spend": {"x": 1}}])
    assert result[0]["date"] == "24.07.26"
    assert result[0]["time"] == "24.07.26 09:00"
    assert result[0]["spend"] == {"x": 1}
    assert "text" in result[0]


def test_calculate_basketball_data_sums_qty_and_groups_by_ammo_name():
    data = [
        {"spend": {"ОФ-843Б 120мм": 2}},
        {"spend": {"ОФ 152мм": 1}},
    ]
    result = m.calculate_basketball_data(data)
    assert result == "ОФ-843Б 120мм. – 2 шт., ОФ 152мм. – 1 шт."


def test_calculate_basketball_data_sums_same_ammo_across_messages():
    data = [{"spend": {"ОФ-843Б 120мм": 2}}, {"spend": {"ОФ-843Б 120мм": 3}}]
    assert m.calculate_basketball_data(data) == "ОФ-843Б 120мм. – 5 шт."


def test_calculate_basketball_data_ignores_messages_without_spend():
    data = [{"text": "щось без витрати"}]
    assert m.calculate_basketball_data(data) == ""


def test_calculate_basketball_data_empty_input():
    assert m.calculate_basketball_data([]) == ""
