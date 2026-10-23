import pytest

from constants import SECTION_DUTIES, SECTION_GENERAL, SECTION_RIGHTS
from content.statute_reader import find_section, heading_key, parse_statute_html, read_statute
from duties_fixtures import STATUTE_HTML, write_statute


def _sections():
    return parse_statute_html(STATUTE_HTML)


def test_heading_key_drops_parentheses_and_footnote_marks():
    assert heading_key("Командир батальйону (корабля 3 рангу)*") == "командир батальйону"
    assert heading_key("Сержант із матеріального забезпечення роти (батареї), старшина роти (батареї)") == \
        "сержант із матеріального забезпечення роти, старшина роти"


def test_parse_splits_general_duties_and_rights_and_collects_articles():
    commander = find_section(_sections(), "Командир батальйону")
    assert commander.articles == ["101", "102"]
    assert commander.source == "Статут ВС, ст. 101, 102"
    assert commander.title == "Командир батальйону"
    assert commander.items(SECTION_GENERAL) == [
        "Командир батальйону (корабля 3 рангу) відповідає за бойову готовність.",
        "Командир батальйону підпорядковується командирові бригади.",
    ]
    assert commander.items(SECTION_DUTIES) == ["підтримувати бойову готовність;", "керувати бойовою підготовкою."]
    staff = find_section(_sections(), "Начальник штабу батальйону")
    assert staff.items(SECTION_RIGHTS) == ["віддавати розпорядження."]


def test_parse_skips_amendment_notes_footnotes_and_additional_lists():
    sections = _sections()
    responsibility = find_section(sections, "Відповідальність військовослужбовців")
    assert responsibility.items(SECTION_GENERAL) == [
        "Військовослужбовці несуть дисциплінарну відповідальність.",
        "Стягнення не звільняє від матеріальної відповідальності.",
    ]
    company = find_section(sections, "Командир роти")
    assert company.items(SECTION_DUTIES) == ["керувати ротою;"]
    assert company.items(SECTION_GENERAL)[-1] == "Окрема стаття після переліку."
    assert company.articles[-1] == "112-1"


def test_find_section_prefers_exact_then_shortest_prefix():
    sections = _sections()
    assert find_section(sections, "Загальні обов’язки військовослужбовців").heading.startswith("Загальні обов’язки")
    assert find_section(sections, "Командир полку") is None
    assert find_section(sections, "Розділ").key.startswith("розділ 3")


def test_section_without_articles_has_plain_source():
    sections = parse_statute_html('<p class="rvps7">Заголовок</p><p class="rvps2">Текст без номера.</p>')
    assert sections[0].source == "Статут ВС"


def test_text_before_first_heading_is_ignored():
    sections = parse_statute_html('<p class="rvps2">Преамбула</p><p></p><p class="rvps7">Заголовок</p>')
    assert len(sections) == 1 and sections[0].sections == {}


def test_read_statute_handles_utf8_and_cp1251(tmp_path):
    assert len(read_statute(write_statute(tmp_path / "u.htm"))) == len(_sections())
    assert len(read_statute(write_statute(tmp_path / "w.htm", encoding="cp1251"))) == len(_sections())


def test_read_statute_reports_unknown_encoding(tmp_path):
    path = tmp_path / "x.htm"
    path.write_bytes(b"\x98")
    with pytest.raises(ValueError, match="кодування"):
        read_statute(path)
