import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import constants as m


# -------------------------
# _parse_pipe_keyed_overrides
# -------------------------
def test_parse_pipe_keyed_overrides_empty_dict_stays_empty():
    assert m._parse_pipe_keyed_overrides({}) == {}


def test_parse_pipe_keyed_overrides_parses_pipe_delimited_keys():
    raw = {"екіпаж1|Сокіл": "варіант1", "_comment": "ігнорується"}
    result = m._parse_pipe_keyed_overrides(raw)
    assert result == {("екіпаж1", "Сокіл"): "варіант1"}


# -------------------------
# resources/data.json - ambiguous_callsign_overrides / bk_synonym_crew_overrides
# розділи об'єднані в один файл разом зі "unit" і щоденним вмістом звіту
# -------------------------
def test_ambiguous_callsign_overrides_loaded_from_merged_data_json():
    assert isinstance(m.AMBIGUOUS_CALLSIGN_OVERRIDES, dict)
    assert all(isinstance(k, tuple) and len(k) == 2 for k in m.AMBIGUOUS_CALLSIGN_OVERRIDES)


def test_bk_synonym_crew_overrides_loaded_from_merged_data_json():
    assert isinstance(m.BK_SYNONYM_CREW_OVERRIDES, dict)
    assert all(isinstance(k, tuple) and len(k) == 2 for k in m.BK_SYNONYM_CREW_OVERRIDES)
