import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import process_generate_br as module

# Ізольовано тестує ЛИШЕ оркестрацію process_generate_br.py (які генератори
# викликаються, для яких дат, за яких умов) - самі генератори (їхній реальний
# вміст) уже перевірені окремо (test_generate_documents_br_weekly_task.py,
# test_combat_log_extract_generators.py тощо), тож тут вони підмінені фейками,
# які лише запам'ятовують виклики. Раніше ця гілка (щотижневий БР - "ЗАВДАННЯ" +
# "Витяг з ЖБД") покривалась ЛИШЕ через реальні дані ОБЛІК.xlsx
# (test_integration.py) - тепер синтетично, незалежно від того, який місяць
# зараз веде користувач у реальному файлі.


def _install_fakes(monkeypatch, calls):
    monkeypatch.setattr(module, "create_or_clear_output_directory", lambda *a, **kw: None)
    monkeypatch.setattr(module, "find_higher_commander", lambda *a, **kw: {})
    monkeypatch.setattr(module, "generate_content_br_general", lambda *a, **kw: {})
    monkeypatch.setattr(module, "generate_documents_br_general", lambda *a, **kw: calls.append("daily_br"))
    monkeypatch.setattr(module, "generate_extract_log_war_general_br_every_day", lambda *a, **kw: calls.append("daily_extract"))
    monkeypatch.setattr(module, "generate_documents_br_save", lambda *a, **kw: calls.append("br_save"))
    monkeypatch.setattr(module, "generate_documents_br", lambda *a, **kw: calls.append("weekly_task"))
    monkeypatch.setattr(module, "generate_documents_combat_log_extract_war_every_week", lambda *a, **kw: calls.append("weekly_extract"))


def test_process_generate_br_runs_weekly_generators_when_bat_configured(monkeypatch):
    """Дата з ФАКТИЧНИМ тижневим БР бат (NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK,
    поле 'бат' непусте) і в межах обраного MONTH/YEAR - генерує "ЗАВДАННЯ" (x2 -
    звичайний і витяг) та щотижневий "Витяг з ЖБД"."""
    calls = []
    _install_fakes(monkeypatch, calls)
    monkeypatch.setattr(module, "MONTH", "07")
    monkeypatch.setattr(module, "YEAR", "2026")
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", {})
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {"01.07.2026": {"бат": "87"}})

    col = datetime(2026, 7, 1)
    result = module.process_generate_br([col], [], [], [], "КИЇВ", "36T TT 12345 67890")

    assert result is True
    assert calls.count("weekly_task") == 2
    assert calls.count("weekly_extract") == 1
    assert calls.count("br_save") == 1
    # Немає запису в NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY за цю дату - щоденні БР/
    # витяг НЕ генеруються.
    assert "daily_br" not in calls
    assert "daily_extract" not in calls


def test_process_generate_br_skips_weekly_generators_without_bat_number(monkeypatch):
    calls = []
    _install_fakes(monkeypatch, calls)
    monkeypatch.setattr(module, "MONTH", "07")
    monkeypatch.setattr(module, "YEAR", "2026")
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", {})
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {})

    module.process_generate_br([datetime(2026, 7, 1)], [], [], [], "КИЇВ", "36T TT 12345 67890")

    assert "weekly_task" not in calls
    assert "weekly_extract" not in calls
    assert calls.count("br_save") == 1


def test_process_generate_br_skips_weekly_generators_for_other_months(monkeypatch):
    """Дата має 'бат', але НЕ належить обраному MONTH/YEAR (напр. перехідний
    хвіст попереднього місяця в ОБЛІК.xlsx) - щотижневі генератори не викликаються."""
    calls = []
    _install_fakes(monkeypatch, calls)
    monkeypatch.setattr(module, "MONTH", "07")
    monkeypatch.setattr(module, "YEAR", "2026")
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", {})
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {"30.06.2026": {"бат": "80"}})

    module.process_generate_br([datetime(2026, 6, 30)], [], [], [], "КИЇВ", "36T TT 12345 67890")

    assert "weekly_task" not in calls
    assert "weekly_extract" not in calls


def test_process_generate_br_runs_daily_generators_when_configured(monkeypatch):
    calls = []
    _install_fakes(monkeypatch, calls)
    monkeypatch.setattr(module, "MONTH", "07")
    monkeypatch.setattr(module, "YEAR", "2026")
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", {"01.07.2026": {"бат": "88", "брг": "1"}})
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {})

    module.process_generate_br([datetime(2026, 7, 1)], [], [], [], "КИЇВ", "36T TT 12345 67890")

    assert calls.count("daily_br") == 2
    assert calls.count("daily_extract") == 1


def test_process_generate_br_merges_pridani_sheet_rows_into_rows_with_data(monkeypatch):
    """Аркуш "ПРИДАНІ" У САМОМУ ОБЛІК.xlsx (rows_with_pridani_sheet_data) -
    рядки просто ДОДАЮТЬСЯ до rows_with_data ще до generate_content_br_general,
    щоб приданий особовий склад отримав БР/ЖБД за свої дні 100/70/170 через
    ТУ САМУ логіку категоризації, що й звичайний особовий склад (жодних змін у
    content/br_general_catalog.py не потрібно)."""
    calls = []
    _install_fakes(monkeypatch, calls)
    monkeypatch.setattr(module, "MONTH", "07")
    monkeypatch.setattr(module, "YEAR", "2026")
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", {"01.07.2026": {"бат": "88", "брг": "1"}})
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {})

    seen_rows_with_data = []
    monkeypatch.setattr(
        module, "generate_content_br_general",
        lambda rows_with_data, *a, **kw: seen_rows_with_data.append(rows_with_data) or {},
    )

    col = datetime(2026, 7, 1)
    regular_row = {"ПІБ": "ПЕРШИЙ Перший", col: 100}
    pridani_row = {"ПІБ": "ДРУГИЙ Другий", col: 100}

    module.process_generate_br([col], [regular_row], [], [], "КИЇВ", "36T TT 12345 67890", [pridani_row])

    pibs = {row["ПІБ"] for rows in seen_rows_with_data for row in rows}
    assert pibs == {"ПЕРШИЙ Перший", "ДРУГИЙ Другий"}


def test_process_generate_br_works_without_pridani_sheet_data(monkeypatch):
    """rows_with_pridani_sheet_data - опційний параметр (за замовчуванням
    порожній) - виклики без нього (як досі) поводяться так само, як і раніше."""
    calls = []
    _install_fakes(monkeypatch, calls)
    monkeypatch.setattr(module, "MONTH", "07")
    monkeypatch.setattr(module, "YEAR", "2026")
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", {})
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {})

    result = module.process_generate_br([datetime(2026, 7, 1)], [], [], [], "КИЇВ", "36T TT 12345 67890")

    assert result is True


def test_process_generate_br_ignores_non_date_columns(monkeypatch):
    """Колонки ПІДРОЗДІЛ/ПОСАДА/ЗВАННЯ/ПІБ (не datetime) - просто пропускаються,
    без жодного виклику генераторів."""
    calls = []
    _install_fakes(monkeypatch, calls)
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY", {})
    monkeypatch.setattr(module, "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK", {})

    module.process_generate_br(["ПІДРОЗДІЛ", "ПОСАДА"], [], [], [], "КИЇВ", "36T TT 12345 67890")

    assert calls == []
