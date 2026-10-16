from bn_fixtures import PBD_PARAGRAPHS, write_docx
from content.docx_blocks import item_text
from content.pbd_reader import read_pbd


def _read(tmp_path, paragraphs=PBD_PARAGRAPHS):
    return read_pbd(str(write_docx(tmp_path / "ПБД.docx", paragraphs)))


def _texts(items):
    return [item_text(item) for item in items]


def test_header(tmp_path):
    report = _read(tmp_path)
    assert (report.ksp_place, report.report_time, report.report_date, report.map_edition) == ("ТЕСТОВЕ", "20.00", "09.10.2026", "2024")
    assert report.ksp_mgrs == "31U DQ 40500 10500"
    assert report.warnings == []


def test_battle_order_block_stops_at_ksp_line(tmp_path):
    report = _read(tmp_path)
    assert _texts(report.blocks["battle_order"]) == ["Бойовий порядок в один ешелон.", "І ешелон:", "друга рота;"]


def test_positions_block_keeps_group_headers(tmp_path):
    texts = _texts(_read(tmp_path).blocks["positions"])
    assert texts[0].startswith("друга рота здійснює стійку оборону РОП «АЛЬФА»")
    assert "Окремі позиції в межах БРО:" in texts
    assert "MAVIK:" in texts and "НРК:" in texts
    assert texts[-1].startswith("ХАБ «ГОТЕЛЬ»")


def test_rop_blocks(tmp_path):
    blocks = _read(tmp_path).blocks
    head = _texts(blocks["rop:АЛЬФА:head"])
    assert len(head) == 2 and head[1].startswith("з переднім краєм")
    assert blocks["rop:АЛЬФА:head"][0][1] == ["РОП «АЛЬФА»"]
    assert [t.split(" ")[0] for t in _texts(blocks["rop:АЛЬФА:positions"])] == ["КСП", "ВОП", "ВП", "СП"]


def test_uav_points_block_excludes_ground_robotics(tmp_path):
    assert _texts(_read(tmp_path).blocks["uav_points"]) == ["MAVIK:", "ТЗ «ФОКСТРОТ» (31U DQ 40400 10400);"]


def test_reserves_and_observation_blocks(tmp_path):
    blocks = _read(tmp_path).blocks
    assert len(blocks["reserves"]) == 1
    assert len(blocks["observation"]) == 1


def test_points_types_groups_and_duplicates(tmp_path):
    points = {(p.type, p.name): p for p in _read(tmp_path).points}
    assert set(points) == {
        ("КСП", ""), ("КСП", "АЛЬФА"), ("ВП", "ЧАРЛІ"), ("СП", "ДЕЛЬТА"), ("ВП", "ЕХО"),
        ("ТЗ", "ФОКСТРОТ"), ("ПУ", "ГОЛЬФ"), ("ХАБ", "ГОТЕЛЬ"), ("БнГ", ""),
    }
    assert points[("ВП", "ЧАРЛІ")].rop == "АЛЬФА" and points[("ВП", "ЧАРЛІ")].vop == "БРАВО"
    assert points[("ТЗ", "ФОКСТРОТ")].group == "MAVIK"
    assert points[("ПУ", "ГОЛЬФ")].group == "НРК"
    assert points[("ХАБ", "ГОТЕЛЬ")].group == "Логістичні хаби"
    assert points[("БнГ", "")].group == "Резерви"


def test_lines(tmp_path):
    lines = {(l.kind, l.ref): len(l.mgrs) for l in _read(tmp_path).lines}
    assert lines == {
        ("bro", ""): 3, ("flot", ""): 2,
        ("boundary_right", "праворуч з A0002"): 2, ("boundary_left", "ліворуч з A0003"): 2,
        ("rop", "РОП «АЛЬФА»"): 3, ("flot_rop", "РОП «АЛЬФА»"): 2,
        ("vop", "ВОП «БРАВО»"): 3, ("flot_vop", "ВОП «БРАВО»"): 2,
    }


def test_duplicate_with_other_coordinates_is_reported(tmp_path):
    paragraphs = list(PBD_PARAGRAPHS)
    paragraphs[paragraphs.index("Спостережні позиції:") + 1] = "СП «ДЕЛЬТА» за координатами (31U DQ 49000 19000);"
    report = _read(tmp_path, paragraphs)
    assert any("ДЕЛЬТА" in w and "двічі" in w for w in report.warnings)


def test_position_without_coordinates_is_reported(tmp_path):
    paragraphs = list(PBD_PARAGRAPHS)
    paragraphs[paragraphs.index("ВП «ЧАРЛІ» в районі з центром (31U DQ 40115 10115);")] = "ВП «ЧАРЛІ» в районі з центром (31U DQ 40115 1011);"
    report = _read(tmp_path, paragraphs)
    assert not any(p.name == "ЧАРЛІ" for p in report.points)
    assert any("ЧАРЛІ" in w for w in report.warnings)


def test_missing_section(tmp_path):
    report = _read(tmp_path, ["Порожній документ"])
    assert report.points == [] and len(report.warnings) == 2


def test_signature_and_reserve_areas(tmp_path):
    report = _read(tmp_path)
    assert _texts(report.blocks["signature"]) == ["Командир A0001", "звання\tПІДПИС"]
    assert report.reserve_areas == {"БнГ": "лісосмуга (31U DQ 40700 10700)"}


def test_signature_without_following_line(tmp_path):
    report = _read(tmp_path, PBD_PARAGRAPHS[:-2])
    assert _texts(report.blocks["signature"]) == ["Командир A0001"]


def test_no_signature(tmp_path):
    report = _read(tmp_path, PBD_PARAGRAPHS[:-4])
    assert "signature" not in report.blocks
