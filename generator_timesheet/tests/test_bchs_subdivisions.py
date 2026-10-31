"""Підрозділ за назвою файлу БЧС (content/bchs_subdivisions.py) - назви
підрозділів синтетичні, але ті самі форми написання, що й у реальних файлах."""

from content.bchs_subdivisions import file_subdivision, missing_subdivisions

_SUBDIVISIONS = ("1 абв", "2 абв", "гдежзи", "жзк", "кл", "клм", "нпр")
_FILES = [
    "БЧС 1 АБВ 27.10.xlsx",
    "БЧС.2АБВ 27.10.2026.xlsx",
    "БЧС ГДЕ.Ж. 27.10.2026.xlsx",
    "БЧС. ЖЗКаб A0000 27.10.2026.xlsx",
    "БЧС_КЛМ__27.10.2026.xlsx",
]


def test_missing_subdivisions_reports_subdivisions_without_a_bchs_file():
    """Рішення користувача: після завантаження червоним, чийого
    файлу БЧС немає. Підрозділ пишуть з пробілом і без ("1 АБВ" / "2АБВ"), з
    крапками ("ГДЕ.Ж." - "гдежзи"), із закінченням ("ЖЗКаб" - "жзк");
    дволітерне - лише точно ("КЛМ" - не "кл")."""
    assert missing_subdivisions(_FILES, _SUBDIVISIONS) == ["кл", "нпр"]
    assert missing_subdivisions(_FILES + ["БЧС КЛ_A0000_27.10.2026.xlsx"], _SUBDIVISIONS) == ["нпр"]
    assert missing_subdivisions([], ("кл",)) == ["кл"]


def test_file_subdivision_names_the_owner_of_the_file():
    assert [file_subdivision(name, _SUBDIVISIONS) for name in _FILES] == ["1 абв", "2 абв", "гдежзи", "жзк", "клм"]
    assert file_subdivision("БЧС КЛ_A0000_27.10.2026.xlsx", _SUBDIVISIONS) == "кл"
    assert file_subdivision("Інший файл.xlsx", _SUBDIVISIONS) is None
