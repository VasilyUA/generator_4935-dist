import math

import pytest

from utils.mgrs import _K0, _meridian_arc, find_mgrs, mgrs_to_latlon, utm_to_latlon


def _distance_m(a, b):
    lat1, lon1, lat2, lon2 = map(math.radians, (*a, *b))
    h = math.sin((lat2 - lat1) / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin((lon2 - lon1) / 2) ** 2
    return 2 * 6371000 * math.asin(math.sqrt(h))


@pytest.mark.parametrize("mgrs, expected", [
    # Публічні орієнтири (координати округлені до 4 знаків - звідси допуск).
    ("31U DQ 48251 11932", (48.8583, 2.2945)),   # Ейфелева вежа
    ("18S UJ 23483 06479", (38.8895, -77.0352)),  # Монумент Вашингтона
])
def test_landmarks(mgrs, expected):
    assert _distance_m(mgrs_to_latlon(mgrs), expected) < 15


def test_central_meridian_is_exact():
    # На осьовому меридіані northing = k0 * довжина дуги меридіана до широти.
    lat, lon = utm_to_latlon(31, 500000, _K0 * _meridian_arc(48.0))
    assert lon == pytest.approx(3.0, abs=1e-9)
    assert lat == pytest.approx(48.0, abs=1e-7)


def test_formats_are_equivalent():
    spaced = mgrs_to_latlon("31U DQ 48251 11932")
    assert mgrs_to_latlon("31UDQ4825111932") == spaced
    assert mgrs_to_latlon("31u dq 48251-11932") == spaced


def test_cyrillic_lookalike_letters():
    # Х і Р кирилицею - як часто набирають у документах.
    assert mgrs_to_latlon("33U ХР 12345 67890") == mgrs_to_latlon("33U XP 12345 67890")


def test_lower_precision_is_south_west_corner():
    assert mgrs_to_latlon("31U DQ 482 119") == mgrs_to_latlon("31U DQ 48200 11900")
    assert mgrs_to_latlon("31U DQ 482-119") == mgrs_to_latlon("31U DQ 48200 11900")


def test_southern_hemisphere_band():
    lat, lon = mgrs_to_latlon("33H VG 00000 00000")
    assert -32 <= lat < -24
    assert 12 < lon < 18


@pytest.mark.parametrize("text", ["", "без координат", "31U DQ 48251 1193", "31U JQ 00000 00000"])
def test_invalid_returns_none(text):
    assert mgrs_to_latlon(text) is None


def test_find_mgrs_positions_refer_to_original_text():
    text = "точка (33U ХР 123-678), інша (31U DQ 48251 1193) кінець"
    found = find_mgrs(text)
    assert [m.valid for m in found] == [True, False]
    assert text[found[0].start:found[0].end] == found[0].raw == "33U ХР 123-678"
    assert found[0].normalized == "33U XP 123 678"


def test_find_mgrs_does_not_match_inside_numbers():
    assert find_mgrs("12345 67890") == []
    assert find_mgrs("№ 1234") == []
