import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from get_normalized_data import get_paintball_data as m
from constants import UNIT_BATTALION, UNIT_BRIGADE


def make_msg(date_iso, body):
    return {"date": date_iso, "body": body}


# -------------------------
# get_paintball_data
# -------------------------
def test_get_paintball_data_reads_filters_and_sorts(monkeypatch):
    raw_messages = [make_msg("2026-07-24T09:30:00", "09:30 артобстріл по позиціях")]
    monkeypatch.setattr(m, "get_signal_data", lambda path: raw_messages)
    monkeypatch.setattr(m, "get_filtered_messages_for_date", lambda msgs, hour, date: msgs)

    result = m.get_paintball_data(18, "24.07.2026")

    assert result == ['09:30 24.07.2026 артобстріл по позиціях. Втрати уточнюються.']


# -------------------------
# _normalize_time
# -------------------------
def test_normalize_time_pads_single_digit_hour():
    assert m._normalize_time("9:05") == "09:05"


def test_normalize_time_converts_dot_to_colon():
    assert m._normalize_time("14.30") == "14:30"


def test_normalize_time_leaves_already_valid_time():
    assert m._normalize_time("09:05") == "09:05"


# -------------------------
# get_sorted_messages
# -------------------------
def test_get_sorted_messages_single_time():
    messages = [make_msg("2026-07-24T09:30:00", "09:30 артобстріл по позиціях")]
    result = m.get_sorted_messages(messages)
    assert result == ['09:30 24.07.2026 артобстріл по позиціях. Втрати уточнюються.']


def test_get_sorted_messages_time_range_kept_whole():
    messages = [make_msg("2026-07-25T09:30:00", '09:30-09:35 артобстріл (2 АС 152мм) по ТЗ "Газда"')]
    result = m.get_sorted_messages(messages)
    assert result == ['09:30-09:35 25.07.2026 артобстріл (2 АС 152мм) по ТЗ "Газда". Втрати уточнюються.']


def test_get_sorted_messages_missing_time_defaults_to_0000():
    messages = [make_msg("2026-07-24T12:00:00", "обстріл без вказаного часу")]
    result = m.get_sorted_messages(messages)
    assert result[0].startswith("00:00 24.07.2026")


def test_get_sorted_messages_skips_utochnennia():
    messages = [make_msg("2026-07-24T09:30:00", "09:30 Уточнення по попередньому обстрілу")]
    assert m.get_sorted_messages(messages) == []


def test_get_sorted_messages_skips_empty_body_or_missing_date():
    messages = [make_msg("2026-07-24T09:30:00", ""), make_msg(None, "09:30 текст")]
    assert m.get_sorted_messages(messages) == []


def test_get_sorted_messages_skips_message_with_unparseable_date():
    messages = [make_msg("не дата взагалі", "09:30 текст")]
    assert m.get_sorted_messages(messages) == []


def test_get_sorted_messages_out_of_range_hour_falls_back_to_midnight_sort_key():
    # "25:10" - валідні 1-2 цифри для години за regex, але datetime.replace(hour=25)
    # падає (година має бути 0-23) - падаємо на північ для сортування, а не валимось
    messages = [make_msg("2026-07-24T09:30:00", "25:10 текст з некоректною годиною")]
    result = m.get_sorted_messages(messages)
    assert result == ['25:10 24.07.2026 текст з некоректною годиною. Втрати уточнюються.']


def test_get_sorted_messages_falls_back_to_midnight_when_time_unparseable(monkeypatch):
    # захисна гілка на випадок, якщо _normalize_time колись поверне щось без
    # двокрапки - парсинг год/хв не повинен валити всю обробку повідомлення
    monkeypatch.setattr(m, "_normalize_time", lambda t: "invalid")
    messages = [make_msg("2026-07-24T09:30:00", "09:30 текст")]
    result = m.get_sorted_messages(messages)
    assert result == ['invalid 24.07.2026 текст. Втрати уточнюються.']


def test_get_sorted_messages_night_message_evening_event_rolls_back_a_day():
    # повідомлення надійшло вночі (02:00), але описує вечірню подію (19:30) -
    # подія тоді трапилась учора, а не сьогодні
    messages = [make_msg("2026-04-22T02:00:00", "19:30 удар по позиціях")]
    result = m.get_sorted_messages(messages)
    assert result == ['19:30 21.04.2026 удар по позиціях. Втрати уточнюються.']


def test_get_sorted_messages_sorted_by_time():
    messages = [
        make_msg("2026-07-24T10:00:00", "10:00 другий обстріл"),
        make_msg("2026-07-24T09:00:00", "09:00 перший обстріл"),
    ]
    result = m.get_sorted_messages(messages)
    assert "перший" in result[0]
    assert "другий" in result[1]


# -------------------------
# classify_attacks
# -------------------------
def test_classify_attacks_counts_matching_category():
    stats = m.classify_attacks(["09:00 артобстріл позицій"])
    assert stats["артилерійських_обстрілів"]["count"] == 1
    assert stats["всього"]["count"] == 1


def test_classify_attacks_skips_utochnennia_messages():
    stats = m.classify_attacks(["09:00 Уточнення по обстрілу"])
    assert stats["всього"]["count"] == 0


def test_classify_attacks_message_can_match_multiple_categories():
    # ключові слова перевіряються незалежно - одне повідомлення може одразу
    # містити ознаки декількох категорій (наприклад авіа + арт в описі)
    stats = m.classify_attacks(["авіаудар (каб) та артобстріл одночасно"])
    assert stats["авіаційних_ударів"]["count"] == 1
    assert stats["артилерійських_обстрілів"]["count"] == 1
    assert stats["всього"]["count"] == 2


def test_classify_attacks_zero_categories_have_no_messages():
    stats = m.classify_attacks(["09:00 артобстріл позицій"])
    assert stats["танковий"]["count"] == 0
    assert stats["танковий"]["messages"] == []


def test_classify_attacks_fpv_and_mavic_are_separate_categories():
    stats = m.classify_attacks(["удар fpv-дроном", "скид з mavic"])
    assert stats["удари_FPV_дронів"]["count"] == 1
    assert stats["скиди_бпла"]["count"] == 1


def test_classify_attacks_bmp_category():
    stats = m.classify_attacks(["обстріл з БМП по позиціях"])
    assert stats["БМП"]["count"] == 1


def test_classify_attacks_bmp_ignores_own_unit_designation():
    # регресія: "{UNIT_BATTALION} {UNIT_BRIGADE}" - штатне позначення СВОГО підрозділу в майже
    # кожному повідомленні каналу, не обстріл ворожою БМП - не повинно рахуватись
    stats = m.classify_attacks([f"09:00 25.07.2026 артобстріл по ТЗ \"Газда\" {UNIT_BATTALION} {UNIT_BRIGADE}"])
    assert stats["БМП"]["count"] == 0


def test_classify_attacks_striletska_zbroya_excluded_when_no_own_losses():
    # "стрілецький бій" здебільшого - наша пошукова група сама знайшла й
    # знищила ворога БЕЗ власних втрат; п.3.1.2 рахує УДАРИ ПРОТИВНИКА по
    # нас, тож такі випадки не повинні йти в "стріл. зброя".
    stats = m.classify_attacks([
        "07:20 - 07:35 внаслідок пошукових дій відбувся стрілецький бій. "
        "Втрати наших військ: без втрат. Втрати противника: 1 в/сл. - 200 (безповоротні)."
    ])
    assert stats["стріл_зброя"]["count"] == 0


def test_classify_attacks_striletska_zbroya_counted_when_own_losses_reported():
    stats = m.classify_attacks([
        "16:10 стрілецький бій. Втрати наших військ: 1 в/сл. поранений. "
        "Втрати противника: без втрат."
    ])
    assert stats["стріл_зброя"]["count"] == 1


def test_classify_attacks_striletska_zbroya_counted_when_no_own_losses_label_at_all():
    # історичний реальний приклад: втрати описані наративно (загибель
    # конкретного військовослужбовця), без мітки "Втрати наших військ:"
    # узагалі - типовий "без втрат" детектор тут нічого не знайде, тож
    # повідомлення МАЄ порахуватись (типова помилка - трактувати "не знайшов
    # мітку" як "немає втрат", хоча насправді втрати є, просто інакше описані)
    stats = m.classify_attacks([
        "в результаті стрілецького бою, противника було знищено1-200, одночасно "
        "отримав поранення не сумісні з життям матрос ШОСТИЙ Шостий Шостий"
    ])
    assert stats["стріл_зброя"]["count"] == 1
