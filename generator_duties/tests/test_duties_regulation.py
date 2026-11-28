from docx import Document

from content.docx_text import read_paragraph_specs
from content.positions import auto_suffixes, group_positions
from content.regulation_draft import (
    DEFAULT_PURPOSE, battalion_names, ensure_regulation_draft, leadership, purpose, subunit_groups,
)
from content.roster_reader import Slot, read_roster
from duties_fixtures import BATTALION, FULL_HEADER, FULL_ROWS, UNIT, write_roster


def _staff(tmp_path):
    write_roster(tmp_path / "ШТАТ.xlsx", rows=FULL_ROWS, header=FULL_HEADER, title_rows=0)
    slots = read_roster(str(tmp_path)).slots
    suffixes = auto_suffixes(slots)
    return slots, group_positions(slots, suffixes), suffixes


def test_battalion_names_from_staff_chain_or_data():
    assert battalion_names({}, [UNIT, BATTALION]) == ("батальйон A0001", BATTALION, UNIT)
    data = {"unit": {"BATTALION_GENITIVE": "2 батальйону A0002", "MILITARY_UNIT_GENITIVE": "військової частини A0003"}}
    assert battalion_names(data, [UNIT, BATTALION]) == ("2 батальйон A0002", "2 батальйону A0002", "військової частини A0003")
    assert battalion_names({}, []) == ("батальйон", "батальйону", "")


def test_subunit_groups_purposes_and_leadership(tmp_path):
    slots, positions, _ = _staff(tmp_path)
    assert subunit_groups(slots) == ["управління", "рота"]
    assert subunit_groups([Slot(1, "x", group="Рота"), Slot(2, "y", group="рота"), Slot(3, "z")]) == ["Рота"]
    assert purpose("Мінометна батарея").startswith("вогнева підтримка")
    assert purpose("Взвод зв'язку").startswith("організація та підтримання")
    assert purpose("Група об'єднаної вогневої підтримки").startswith("планування, координація")
    assert purpose("Рота вогневої підтримки").startswith("вогнева підтримка")
    assert purpose("Щось інше") == DEFAULT_PURPOSE
    assert leadership(positions) == ["Заступник командира батальйону з озброєння"]
    nested = group_positions([Slot(1, "Головний сержант відділення управління командира батальйону взводу"),
                              Slot(2, "Головний сержант батальйону")])
    assert leadership(nested) == ["Головний сержант батальйону"]


def test_ensure_regulation_draft_creates_document_once(tmp_path):
    slots, positions, suffixes = _staff(tmp_path)
    path = ensure_regulation_draft(str(tmp_path), positions, slots, {}, suffixes)
    assert path.endswith("Положення про батальйон.docx")
    specs = read_paragraph_specs(Document(path))
    texts = [spec.text for spec in specs]
    assert texts[:3] == ["ПОЛОЖЕННЯ", "про батальйон A0001", UNIT]
    assert all(spec.center and spec.bold for spec in specs[:3])
    assert texts[3] == "1. Загальні положення" and specs[3].bold and not specs[3].center
    assert f"1.1. Батальйон A0001 {UNIT} (далі – батальйон) є тактичним підрозділом і призначений для ведення бойових " \
           f"дій самостійно або у складі військової частини." in texts
    assert "– управління;" in texts and "– рота;" in texts
    assert any(text.startswith("– рота – виконання бойових завдань") for text in texts)
    assert f"4.1. Батальйон очолює командир батальйону, який підпорядковується командиру {UNIT} і є прямим начальником " \
           f"усього особового складу батальйону." in texts
    assert "– Заступник командира батальйону з озброєння;" in texts
    assert ensure_regulation_draft(str(tmp_path), positions, slots, {}, suffixes) is None


def test_regulation_draft_without_groups_or_unit(tmp_path):
    slots = [Slot(1, "Стрілець", vos="1")]
    path = ensure_regulation_draft(str(tmp_path), group_positions(slots), slots, {}, [])
    texts = [spec.text for spec in read_paragraph_specs(Document(path))]
    assert texts[:2] == ["ПОЛОЖЕННЯ", "про батальйон"]
    assert "– ЗАПОВНИТИ: перелік підрозділів;" in texts
    assert any("підпорядковується старшому командиру" in text for text in texts)
