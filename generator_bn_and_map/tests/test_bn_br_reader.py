from bn_fixtures import BR_PARAGRAPHS, write_docx
from content.br_reader import read_br, read_br_header
from content.docx_blocks import item_text


def _texts(items):
    return [item_text(item) for item in items]


def test_header(tmp_path):
    path = write_docx(tmp_path / "БР.docx", BR_PARAGRAPHS)
    order = read_br(str(path))
    assert (order.number, order.date, order.time) == ("1234", "19.08.2026", "16.00")
    header = read_br_header(str(path))
    assert (header.number, header.date) == ("1234", "19.08.2026")
    assert header.blocks == {}


def test_enemy_block_is_not_cut_by_text_starting_with_commander(tmp_path):
    order = read_br(str(write_docx(tmp_path / "БР.docx", BR_PARAGRAPHS)))
    assert _texts(order.blocks["br_enemy"]) == ["Противник веде оборону.", "Командир противника діє пасивно."]


def test_fire_tasks_keep_bold_phrases(tmp_path):
    order = read_br(str(write_docx(tmp_path / "БР.docx", BR_PARAGRAPHS)))
    assert order.blocks["br_fire_tasks"] == [["101   ЖС   31U DQ 48251 11932;", ["101"]]]


def test_mission_stops_before_offensive_stages(tmp_path):
    order = read_br(str(write_docx(tmp_path / "БР.docx", BR_PARAGRAPHS)))
    assert _texts(order.blocks["br_mission"]) == [
        "A0001 продовжити ведення оборонного бою.",
        "Бойовий порядок мати в один ешелон.",
        "Основні зусилля зосередити на утриманні рубежу.",
    ]


def test_neighbours_are_prefixed_and_continuation_merged(tmp_path):
    order = read_br(str(write_docx(tmp_path / "БР.docx", BR_PARAGRAPHS)))
    assert _texts(order.blocks["br_neighbours"]) == [
        "Сусід праворуч: A0002 обороняє район.",
        "Сусід ліворуч: A0003 обороняє район: лісосмуга (31U DQ 48000 11000).",
    ]


def test_readiness_stops_at_signature(tmp_path):
    order = read_br(str(write_docx(tmp_path / "БР.docx", BR_PARAGRAPHS)))
    assert _texts(order.blocks["br_readiness"]) == ["Час готовності до дій - з отриманням розпорядження."]
    assert order.warnings == []


def test_missing_sections_and_header_are_reported(tmp_path):
    path = write_docx(tmp_path / "БР.docx", ["Текст без розділів."])
    order = read_br(str(path))
    assert order.blocks == {}
    assert len(order.warnings) == 2 + 5


def test_mission_without_stop_phrase_is_kept_whole(tmp_path):
    paragraphs = [p for p in BR_PARAGRAPHS if not (isinstance(p, str) and p.startswith("З початком"))]
    order = read_br(str(write_docx(tmp_path / "БР.docx", paragraphs)))
    assert "Етап перший." in _texts(order.blocks["br_mission"])
    assert any("повністю" in w for w in order.warnings)


def test_stage_blocks_and_ksp(tmp_path):
    blocks = read_br(str(write_docx(tmp_path / "БР.docx", BR_PARAGRAPHS))).blocks
    assert _texts(blocks["br_ksp"]) == ["КСП A0001 - ТЕСТОВЕ.", "Запасне КСП - ТЕСТОВЕ."]
    assert _texts(blocks["br_stage:1:right"]) == ["праворуч: точка (31U DQ 48000 11000), відповідальний сусід."]
    assert _texts(blocks["br_stage:1:left"]) == ["ліворуч: точка (31U DQ 49000 11000), відповідальний сусід."]
    assert _texts(blocks["br_stage:1:withdrawal"]) == ["з отриманням сигналу ‘‘ГРІМ - 101’’ відійти на рубіж."]
    assert _texts(blocks["br_stage:1:routes"]) == ["№1 точка (31U DQ 48500 11500).", "Запасні маршрути відходу визначити своїм рішенням."]
    assert _texts(blocks["br_stage:1:actions"])[0].startswith("З початком наступальних дій")
    assert _texts(blocks["br_stage:2:all"]) == ["У разі вклинення противника в межі відсічного рубежу діяти своїм рішенням."]
    assert "br_stage:2:right" not in blocks


def test_stage_text_after_boundaries_returns_to_actions(tmp_path):
    paragraphs = list(BR_PARAGRAPHS)
    index = paragraphs.index("Порядок відходу мати:")
    paragraphs.insert(index, "Додаткова вказівка після меж.")
    blocks = read_br(str(write_docx(tmp_path / "БР.docx", paragraphs))).blocks
    assert "Додаткова вказівка після меж." in _texts(blocks["br_stage:1:actions"])
