from content.positions import (
    auto_suffixes, common_suffix, compare_key, find_commander, group_positions, normalize_dashes, normalize_title,
    strip_chain,
)
from content.roster_reader import Slot, read_roster
from duties_fixtures import BATTALION, FULL_HEADER, FULL_ROWS, SUFFIX, UNIT, make_positions, write_roster


def test_normalize_title_strips_numbers_and_higher_unit_chain():
    assert normalize_title(f"Стрілець 2 відділення 3 взводу {SUFFIX}", [SUFFIX]) == "Стрілець відділення взводу"


def test_normalize_title_keeps_first_word_of_chain_when_nothing_else_left():
    assert normalize_title(f"Командир {SUFFIX}", [SUFFIX]) == "Командир батальйону"


def test_normalize_title_keeps_chain_word_after_dangling_genitive():
    title = normalize_title(f"Начальник штабу - заступник командира {SUFFIX}", [SUFFIX])
    assert title == "Начальник штабу – заступник командира батальйону"


def test_composite_positions_written_with_hyphen_get_a_dash():
    assert normalize_dashes("Головний сержант-командир відділення") == "Головний сержант – командир відділення"
    assert normalize_dashes("Начальник штабу-Заступник командира") == "Начальник штабу – Заступник командира"
    assert normalize_dashes("Водій-електрик відділення") == "Водій-електрик відділення"
    assert normalize_dashes("Начальник штабу — заступник") == "Начальник штабу – заступник"
    assert compare_key("Начальник штабу-заступник командира") == compare_key("начальник штабу – заступник командира")


def test_normalize_title_unifies_apostrophes_spaces_and_case():
    assert normalize_title("  командир   відділення зв'язку  №1 ") == "Командир відділення зв’язку"


def test_normalize_title_ignores_empty_suffixes_and_suffix_at_start():
    assert normalize_title(SUFFIX, ["", SUFFIX]) == "Батальйону A0000"


def test_normalize_title_keeps_specialization_after_chain_in_the_middle():
    title = normalize_title(f"Заступник командира {SUFFIX} з озброєння", [SUFFIX])
    assert title == "Заступник командира батальйону з озброєння"


def test_normalize_title_keeps_category_rank_and_class_numbers():
    assert normalize_title("Штаб-сержант 1 категорії групи 2 роти") == "Штаб-сержант 1 категорії групи роти"
    assert normalize_title("Водій 3 класу 4 відділення") == "Водій 3 класу відділення"


def test_normalize_title_with_numbers_only_suffix_keeps_head():
    assert normalize_title("Командир 7 8", ["7 8"]) == "Командир"


def test_compare_key_ignores_case_dashes_and_edge_punctuation():
    assert compare_key("Начальник штабу — заступник командира.") == compare_key("начальник штабу - заступник командира")


def test_group_positions_keeps_staff_order_and_collects_slots():
    positions = make_positions()
    assert [p.title for p in positions] == [
        "Командир батальйону",
        "Начальник штабу – заступник командира батальйону",
        "Командир відділення взводу",
        "Стрілець відділення взводу",
        "Водій господарчого відділення",
    ]
    rifleman = positions[3]
    assert len(rifleman.slots) == 2
    assert rifleman.occupied == 1
    assert rifleman.vos_codes == ["1000004"]


def test_group_positions_skips_slots_with_empty_title():
    assert group_positions([Slot(1, "12")]) == []


def test_find_commander_prefers_occupied_slot():
    vacant = Slot(1, "Командир батальйону", vos="1")
    occupied = Slot(2, "Командир батальйону", "майор", "ПЕРШИЙ Перший Перший", "1")
    positions = group_positions([vacant, occupied])
    assert find_commander(positions) is occupied


def test_strip_chain_keeps_numbers_and_uses_last_occurrence():
    raw = f"Стрілець 1 відділення роти {BATTALION} роти {BATTALION} {UNIT}"
    assert strip_chain(raw, [BATTALION, UNIT]) == f"Стрілець 1 відділення роти {BATTALION} роти"


def test_common_suffix_whole_words():
    assert common_suffix(["а б в г", "д в г", "", "в г"]) == "в г"
    assert common_suffix(["а б", "в г"]) == ""
    assert common_suffix([]) == ""


def test_auto_suffixes_from_full_staff_list(tmp_path):
    write_roster(tmp_path / "ШТАТ.xlsx", rows=FULL_ROWS, header=FULL_HEADER, title_rows=0)
    slots = read_roster(str(tmp_path)).slots
    suffixes = auto_suffixes(slots)
    assert suffixes == [UNIT, BATTALION]
    positions = group_positions(slots, suffixes)
    assert [p.title for p in positions] == [
        "Командир батальйону", "Заступник командира батальйону з озброєння", "Стрілець відділення взводу роти",
    ]
    assert [slot.display for slot in positions[2].slots] == ["Стрілець 1 відділення 2 взводу роти",
                                                             "Стрілець 2 відділення 2 взводу роти"]


def test_auto_suffixes_rejects_non_chain_and_whole_titles():
    assert auto_suffixes([Slot(1, "Стрілець запасу"), Slot(2, "Кухар запасу")]) == []
    assert auto_suffixes([Slot(1, f"Командир {UNIT}")]) == []
    assert auto_suffixes([Slot(1, "Стрілець роти", subunit="роти"), Slot(2, "Кухар", subunit="")]) == []
    slots = [Slot(1, f"Стрілець {UNIT}", subunit=f"роти {BATTALION}"), Slot(2, f"Кухар {UNIT}", subunit=BATTALION)]
    assert auto_suffixes(slots) == [UNIT, BATTALION]


def test_find_commander_falls_back_to_vacancy_or_none():
    vacant = Slot(1, "Командир батальйону", vos="1")
    assert find_commander(group_positions([vacant])) is vacant
    assert find_commander(group_positions([Slot(1, "Стрілець", vos="1")])) is None
