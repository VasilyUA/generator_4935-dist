import pytest

from constants import (
    ALL_POSITIONS, SECTION_DUTIES, SECTION_GENERAL, SECTION_RESPONSIBILITY, SECTION_RIGHTS, SHEET_FUNCTIONAL, SHEET_JOB,
    STATUTE_RANK_AND_FILE, TO_FILL,
)
from content.prefill import common_set, label_names, match_statute_rule, prefill_set
from content.sample_reader import read_sample
from content.sources import Sources
from content.statute_reader import parse_statute_html
from duties_fixtures import STANDARD_RESPONSIBILITY, STATUTE_HTML, make_positions, write_sample


@pytest.fixture
def sources(tmp_path):
    return Sources(statute=parse_statute_html(STATUTE_HTML), sample=read_sample(write_sample(tmp_path / "s.docx")))


@pytest.mark.parametrize("title, heading", [
    ("Командир батальйону", "Командир батальйону"),
    ("Командир господарчого відділення взводу роти", "Командир відділення"),
    ("Командир взводу роти", "Командир взводу"),
    ("Головний сержант – командир автомобільного відділення взводу", "Командир відділення"),
    ("Головний сержант-командир міномета мінометного взводу мінометної батареї", "Командир відділення"),
    ("Головний сержант-командир відділення управління командира батальйону взводу зв’язку", "Командир відділення"),
    ("Начальник штабу-заступник командира батальйону", "Начальник штабу батальйону"),
    ("Заступник командира роти з психологічної підтримки персоналу", "Заступник командира роти з морально-психологічного"),
    ("Заступник командира роти з озброєння роти батальйону", "Заступник командира роти з озброєння"),
    ("Стрілець відділення взводу", STATUTE_RANK_AND_FILE),
    ("Старший водій-електрик відділення", STATUTE_RANK_AND_FILE),
    ("Офіцер штабу", None),
    ("Старший офіцер відділення", None),
    ("Головний сержант – офіцер штабу", None),
    ("Начальник групи персоналу штабу", None),
    ("Командир взводу-старший офіцер на батареї мінометного взводу мінометної батареї", "Командир взводу"),
    ("Командир міномета мінометного взводу мінометної батареї", "Командир відділення"),
    ("Заступник командира бойової машини-навідник-оператор відділення взводу роти", STATUTE_RANK_AND_FILE),
    ("Заступник командира мінометної батареї з психологічної підтримки персоналу",
     "Заступник командира роти з морально-психологічного"),
    ("Заступник командира роти вогневої підтримки з озброєння", "Заступник командира роти з озброєння"),
    ("Стрілець-помічник гранатометника відділення взводу", STATUTE_RANK_AND_FILE),
    ("Помічник гранатометника відділення", STATUTE_RANK_AND_FILE),
    ("Помічник командира батальйону", None),
    ("Штаб-сержант 1 категорії групи розвідки штабу", None),
    ("Сержант-менеджер 2 категорії групи персоналу штабу", None),
    ("Молодший сержант відділення", None),
])
def test_match_statute_rule(title, heading):
    assert match_statute_rule(title) == heading


def test_match_statute_rule_empty_title():
    assert match_statute_rule("12") is None


def _texts(duty_set, section):
    return [item.text for item in duty_set.sections.get(section, [])]


def test_job_set_prefers_statute_and_is_named_after_it(sources):
    commander = make_positions()[0]
    duty_set = prefill_set(commander, SHEET_JOB, sources)
    assert duty_set.name == "Командир батальйону (Статут)"
    assert duty_set.sections[SECTION_GENERAL][0].source == "Статут ВС, ст. 101, 102"
    assert _texts(duty_set, SECTION_DUTIES) == ["підтримувати бойову готовність;", "керувати бойовою підготовкою."]
    assert SECTION_RESPONSIBILITY not in duty_set.sections


def test_functional_set_prefers_sample(sources):
    commander = make_positions()[0]
    duty_set = prefill_set(commander, SHEET_FUNCTIONAL, sources)
    assert _texts(duty_set, SECTION_DUTIES) == ["керувати батальйоном;"]
    assert duty_set.name == "Командир батальйону (зразок)"


def test_sample_block_matches_position_by_title_prefix(sources):
    driver = make_positions([(1, "Водій автомобільного відділення взводу", "1", "солдат", "")])[0]
    duty_set = prefill_set(driver, SHEET_FUNCTIONAL, sources)
    assert _texts(duty_set, SECTION_RIGHTS) == ["отримувати пальне;"]
    assert duty_set.name == "Водій автомобільного відділення (зразок)"


def test_mixed_sources_are_named_with_both(sources):
    chief = make_positions()[1]
    duty_set = prefill_set(chief, SHEET_JOB, sources)
    assert _texts(duty_set, SECTION_RIGHTS) == ["віддавати розпорядження."]
    assert duty_set.name == "Начальник штабу батальйону (Статут)"
    driver = make_positions()[4]
    assert prefill_set(driver, SHEET_FUNCTIONAL, sources).name == \
        "Загальні обов’язки військовослужбовців рядового складу (Статут)"


def test_position_without_sources_gets_to_fill_set_named_after_it():
    driver = make_positions()[4]
    duty_set = prefill_set(driver, SHEET_JOB, Sources())
    assert duty_set.name == driver.title
    assert duty_set.is_to_fill
    assert [(s, [i.text for i in items]) for s, items in duty_set.sections.items()] == [
        (SECTION_GENERAL, [TO_FILL]), (SECTION_DUTIES, [TO_FILL]),
    ]


def test_common_sets_differ_between_sheets(sources):
    functional = common_set(SHEET_FUNCTIONAL, sources)
    assert functional.name == ALL_POSITIONS
    assert _texts(functional, SECTION_RESPONSIBILITY) == list(STANDARD_RESPONSIBILITY)
    job = common_set(SHEET_JOB, sources)
    assert _texts(job, SECTION_RESPONSIBILITY) == [
        "Військовослужбовці несуть дисциплінарну відповідальність.",
        "Стягнення не звільняє від матеріальної відповідальності.",
    ]
    assert job.sections[SECTION_RESPONSIBILITY][0].source == "Статут ВС, ст. 26, 27"


def test_common_sets_without_sources_or_section():
    assert common_set(SHEET_FUNCTIONAL, Sources()).sections == {}
    assert common_set(SHEET_JOB, Sources(statute=parse_statute_html("<p class='rvps7'>Інше</p>"))).sections == {}


def test_label_names(sources):
    names = label_names(sources)
    assert names["Статут ВС, ст. 125, 126"] == "Командир відділення (Статут)"
    assert label_names(Sources()) == {}
