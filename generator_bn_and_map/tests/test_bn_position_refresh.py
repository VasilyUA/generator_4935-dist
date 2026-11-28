from content.position_refresh import PositionRefresher

POINTS = [
    {"type": "ВП", "name": "АЛЬФА", "mgrs": "31U DQ 40100 10100"},
    {"type": "КСП", "name": "", "mgrs": "31U DQ 40500 10500"},
    {"type": "ВП", "name": "НОВА", "mgrs": "31U DQ 40300 10300"},
    {"type": "ТЗ", "name": "БРАВО", "mgrs": "31U DQ 40700 10700"},
    {"type": "ВП", "name": "БЕЗ КООРДИНАТ", "mgrs": ""},
]


def _refresher():
    return PositionRefresher(POINTS, {"ГОРА 1": ("КСП", "")})


def test_coordinates_refreshed_by_name():
    result = _refresher().refresh("ВП «АЛЬФА» за координатами (31U DQ 49999 19999);")
    assert result.text == "ВП «АЛЬФА» за координатами (31U DQ 40100 10100);"
    assert result.refreshed == 1 and result.auto_mgrs == {"31U DQ 40100 10100"}


def test_alias_for_unnamed_command_post():
    text = "КСП «ГОРА 1» мати околиці н.п. «ТЕСТОВЕ» за координатами (31U DQ 41111 11111)."
    assert _refresher().refresh(text).text.endswith("(31U DQ 40500 10500).")


def test_renamed_position_found_by_coordinates():
    result = _refresher().refresh("ПВ «СТАРА» (31U DQ 40301 10301)")
    assert result.text == "ПВ «НОВА» (31U DQ 40300 10300)"
    assert result.renamed == [("ПВ «СТАРА»", "ВП «НОВА»")]


def test_unknown_position_kept_and_reported():
    text = "ВП «ЗНИКЛА» (31U DQ 45000 15000)"
    result = _refresher().refresh(text)
    assert result.text == text and result.unknown == ["ВП «ЗНИКЛА»"] and result.refreshed == 0


def test_rename_requires_compatible_type():
    text = "ХАБ «ІНШИЙ» (31U DQ 40701 10701)"
    assert _refresher().refresh(text).unknown == ["ХАБ «ІНШИЙ»"]


def test_areas_are_not_touched():
    text = "ВОП «БРАВО»: ліс (31U DQ 40100 10101), ліс (31U DQ 40102 10100)"
    assert _refresher().refresh(text).text == text


def test_coordinates_before_the_name():
    result = _refresher().refresh("збір за координатами 31U DQ 49999 19999 — ВП «АЛЬФА».")
    assert result.text == "збір за координатами 31U DQ 40100 10100 — ВП «АЛЬФА»."


def test_each_coordinate_used_once_and_mention_without_coordinates():
    text = "ВП «АЛЬФА» і ВП «НОВА» (31U DQ 49999 19999)"
    result = _refresher().refresh(text)
    assert result.text == "ВП «АЛЬФА» і ВП «НОВА» (31U DQ 40300 10300)"


def test_no_points_or_text():
    assert PositionRefresher([]).refresh("ВП «АЛЬФА» (31U DQ 1 1)").text == "ВП «АЛЬФА» (31U DQ 1 1)"
    assert _refresher().refresh("").text == ""
