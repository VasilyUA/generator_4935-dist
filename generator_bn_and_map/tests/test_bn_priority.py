"""Пріоритет джерел для даних, що можуть бути і в БР, і в ПБД (координати
КСП/ЗКП, н.п. КСП, видання карти): що першим - БР чи ПБД - обирається
запитанням при запуску, попередній бойовий наказ (data_auto.json → previous)
завжди третій. Синтетичні дані."""
from datetime import date
from types import SimpleNamespace

import docx
import pytest

import generators.pipeline as pipeline
import index as bn_index
import utils.prompts as prompts
from bn_fixtures import BR_PARAGRAPHS, PBD_PARAGRAPHS, make_data, write_docx, write_rop_vop
from content.br_reader import read_br
from content.data_store import split_data
from generators.generate_bn import (
    PRIORITY_BR_FIRST, PRIORITY_PBD_FIRST, br_priority_values, generate_bn, priority_order, priority_overview,
    resolve_priority_values, review_lines, settlement_name,
)

BR_KSP = "КСП A0001 - ВИГАДАНЕ (31U DQ 40900 10900)."
BR_ZKP = "Запасне КСП - ВІДДАЛЕНЕ (31U DQ 40950 10950)."
BR_VALUES = {"ksp_mgrs": "31U DQ 40900 10900", "zkp_mgrs": "31U DQ 40950 10950", "ksp_place": "ВИГАДАНЕ", "map_edition": "2025"}
PBD_VALUES = {"ksp_mgrs": "31U DQ 40500 10500", "zkp_mgrs": "31U DQ 40800 10800", "ksp_place": "ТЕСТОВЕ", "map_edition": "2024"}
PREVIOUS = {"KSP_MGRS": "31U DQ 40100 10100", "ZKP_MGRS": "31U DQ 40200 10200", "KSP_PLACE": "ПОПЕРЕДНЄ", "MAP_EDITION": "2023"}
PREVIOUS_VALUES = {"ksp_mgrs": "31U DQ 40100 10100", "zkp_mgrs": "31U DQ 40200 10200", "ksp_place": "ПОПЕРЕДНЄ",
                   "map_edition": "2023"}


def _br(ksp=BR_KSP, zkp=BR_ZKP, edition="2025"):
    lines = [line for line in (ksp, zkp) if line]
    return SimpleNamespace(blocks={"br_ksp": lines} if lines else {}, map_edition=edition)


def _expect(values, source):
    return {key: (value, source) for key, value in values.items()}


# --- порядок джерел -------------------------------------------------------------------
def test_priority_order():
    assert priority_order("br") == PRIORITY_BR_FIRST == ("br", "pbd", "previous")
    assert priority_order("pbd") == PRIORITY_PBD_FIRST == ("pbd", "br", "previous")
    assert priority_order(None) == PRIORITY_BR_FIRST


@pytest.mark.parametrize("order, br, pbd, expected", [
    # 1) БР - головне: є в усіх трьох джерелах -> з БР.
    (PRIORITY_BR_FIRST, _br(), PBD_VALUES, _expect(BR_VALUES, "br")),
    # 1) ПБД - головне: є в усіх трьох -> з ПБД, хоч у БР теж є.
    (PRIORITY_PBD_FIRST, _br(), PBD_VALUES, _expect(PBD_VALUES, "pbd")),
    # 2) Першого джерела немає -> друге.
    (PRIORITY_BR_FIRST, None, PBD_VALUES, _expect(PBD_VALUES, "pbd")),
    (PRIORITY_PBD_FIRST, _br(), {}, _expect(BR_VALUES, "br")),
    # 3) Ні БР, ні ПБД -> попередній бойовий наказ (за будь-якого вибору).
    (PRIORITY_BR_FIRST, None, {}, _expect(PREVIOUS_VALUES, "previous")),
    (PRIORITY_PBD_FIRST, _br(None, None, ""), {}, _expect(PREVIOUS_VALUES, "previous")),
])
def test_three_sources_in_order(order, br, pbd, expected):
    assert resolve_priority_values(br, pbd, PREVIOUS, order) == expected


def test_each_value_resolved_independently():
    br = _br(zkp=None, edition="")                    # у БР лише КСП
    pbd = {"zkp_mgrs": PBD_VALUES["zkp_mgrs"]}        # у ПБД лише ЗКП
    resolved = resolve_priority_values(br, pbd, PREVIOUS, PRIORITY_PBD_FIRST)
    assert resolved == {
        "ksp_mgrs": (BR_VALUES["ksp_mgrs"], "br"),
        "zkp_mgrs": (PBD_VALUES["zkp_mgrs"], "pbd"),
        "ksp_place": (BR_VALUES["ksp_place"], "br"),
        "map_edition": (PREVIOUS_VALUES["map_edition"], "previous"),
    }


def test_nothing_anywhere_and_incomplete_order():
    empty = {key: ("", "") for key in BR_VALUES}
    assert resolve_priority_values(None, {}, {}, PRIORITY_BR_FIRST) == empty
    assert resolve_priority_values(None, None, None) == empty
    # Порядок без попереднього наказу чи з зайвим - попередній однаково останній.
    assert resolve_priority_values(None, {}, PREVIOUS, ("pbd", "br")) == _expect(PREVIOUS_VALUES, "previous")
    assert resolve_priority_values(_br(), {}, PREVIOUS, ("невідоме",)) == _expect(BR_VALUES, "br")


# --- значення з БР ---------------------------------------------------------------------
def test_settlement_name():
    assert settlement_name("КСП A0001 - ВИГАДАНЕ (31U DQ 40900 10900).") == "ВИГАДАНЕ"
    assert settlement_name("КСП A0001 A0000 - півд.- зах. ВИГАДАНЕ.") == "ВИГАДАНЕ"
    assert settlement_name("Запасне КСП - сх. ДВА СЛОВА.") == "ДВА СЛОВА"
    assert settlement_name("КСП A0001 - без назви.") == ""
    assert settlement_name("") == ""


def test_br_values_read_from_document(tmp_path):
    paragraphs = [BR_KSP if p == "КСП A0001 - ТЕСТОВЕ." else BR_ZKP if p == "Запасне КСП - ТЕСТОВЕ." else p
                  for p in BR_PARAGRAPHS]
    br = read_br(write_docx(tmp_path / "БР 1234.docx", paragraphs))
    assert br.map_edition == "2025"
    assert br_priority_values(br) == BR_VALUES
    assert br_priority_values(None) == {key: "" for key in BR_VALUES}


# --- повна генерація ---------------------------------------------------------------------
@pytest.fixture
def sources(tmp_path):
    res = tmp_path / "resources"
    folder = res / "ПБД"
    folder.mkdir(parents=True)
    write_docx(folder / "ПБД 09.10.2026.docx", PBD_PARAGRAPHS)
    write_rop_vop(folder / "РОП_ВОП 09.10.2026.xlsx")
    paragraphs = [BR_KSP if p == "КСП A0001 - ТЕСТОВЕ." else BR_ZKP if p == "Запасне КСП - ТЕСТОВЕ." else p
                  for p in BR_PARAGRAPHS]
    br = read_br(write_docx(res / "БР 1234.docx", paragraphs))
    cache, _, _ = pipeline.refresh_pbd_cache({}, str(res))
    return br, cache


def _texts(result):
    return [p.text for p in docx.Document(result.path).paragraphs]


@pytest.mark.parametrize("order, place, edition, ksp, zkp", [
    (PRIORITY_BR_FIRST, "Вигадане", "2025", BR_VALUES["ksp_mgrs"], BR_VALUES["zkp_mgrs"]),
    (PRIORITY_PBD_FIRST, "Тестове", "2024", PBD_VALUES["ksp_mgrs"], PBD_VALUES["zkp_mgrs"]),
])
def test_generated_order_follows_chosen_priority(sources, tmp_path, order, place, edition, ksp, zkp):
    br, cache = sources
    assert {key: cache[key] for key in PBD_VALUES} == PBD_VALUES
    data = make_data()
    data["bn"].update(PREVIOUS)
    result = generate_bn(data, br, cache, date(2026, 8, 24), "06.00", "5", str(tmp_path / "out"), priority=order)
    texts = _texts(result)
    assert f"{place} – 2026 рік" in texts
    assert f"КСП – {place} 06.00 24.08.2026. Карта 25 000, видання {edition}р." in texts
    assert f"КСП ({ksp}), ЗКП ({zkp}), {{невідомий}}" in texts
    assert result.priority == order
    assert {source for _value, source in result.priority_values.values()} == {order[0]}


def test_generated_order_falls_back_to_previous(sources, tmp_path):
    _br_order, cache = sources
    for key in PBD_VALUES:
        cache[key] = ""
    data = make_data()
    data["bn"].update(PREVIOUS)
    result = generate_bn(data, None, cache, date(2026, 8, 24), "06.00", "5", str(tmp_path / "out"))
    texts = _texts(result)
    assert "КСП – Попереднє 06.00 24.08.2026. Карта 25 000, видання 2023р." in texts
    assert "КСП (31U DQ 40100 10100), ЗКП (31U DQ 40200 10200), {невідомий}" in texts
    assert result.priority_values == _expect(PREVIOUS_VALUES, "previous")
    assert not any("немає ні в БР, ні в ПБД" in w for w in result.warnings)


def test_review_shows_chosen_priority(sources, tmp_path):
    br, cache = sources
    result = generate_bn(make_data(), br, cache, date(2026, 8, 24), "06.00", "5", str(tmp_path), priority=PRIORITY_PBD_FIRST)
    review = "\n".join(review_lines(result))
    assert "пріоритет, обраний при запуску: ПБД -> БР -> попередній наказ" in review
    assert "  - Координати КСП: ПБД / РОП_ВОП" in review and "  - Видання карти: ПБД / РОП_ВОП" in review


# --- запитання в терміналі і запам'ятовування ----------------------------------------------
class _FakeListPrompt:
    calls = []
    answer = None

    def __init__(self, **kwargs):
        self.kwargs = kwargs
        _FakeListPrompt.calls.append(kwargs)

    def execute(self):
        return _FakeListPrompt.answer if _FakeListPrompt.answer is not None else self.kwargs["default"]


def test_terminal_question(monkeypatch):
    monkeypatch.setattr(prompts, "ListPrompt", _FakeListPrompt)
    _FakeListPrompt.calls, _FakeListPrompt.answer = [], None
    assert prompts.ask_source_priority() == "br"                       # Enter - БР першим
    call = _FakeListPrompt.calls[0]
    assert [choice["value"] for choice in call["choices"]] == ["br", "pbd"]
    assert call["choices"][0]["name"].startswith("1) Бойове розпорядження → 2) ПБД → 3) попередній")
    _FakeListPrompt.answer = "pbd"
    assert prompts.ask_source_priority() == "pbd"


def test_index_asks_and_shows_sources(monkeypatch):
    printed = []
    monkeypatch.setattr(bn_index, "print_purple", printed.append)
    monkeypatch.setattr(bn_index, "ask_source_priority", lambda default: "pbd")
    order = bn_index._ask_priority({"bn": PREVIOUS}, _br(), PBD_VALUES)
    assert order == PRIORITY_PBD_FIRST
    assert ("  - Координати КСП: БР - 31U DQ 40900 10900; ПБД - 31U DQ 40500 10500; "
            "попередній наказ - 31U DQ 40100 10100") in printed
    assert priority_overview(None, {}, {})[0] == "Координати КСП: БР - немає; ПБД - немає; попередній наказ - немає"


def test_index_remembers_values_for_next_order(monkeypatch):
    saved = []
    monkeypatch.setattr(bn_index, "save_auto_data", lambda path, data: saved.append(data))
    data = {"bn": {"NUMBER": "5", "KSP_MGRS": "старе", "MAP_EDITION": "2023"}}
    result = SimpleNamespace(priority_values={
        "ksp_mgrs": ("31U DQ 40900 10900", "br"),
        "zkp_mgrs": ("31U DQ 40800 10800", "pbd"),
        "ksp_place": ("ПОПЕРЕДНЄ", "previous"),
        "map_edition": ("", ""),
    })
    bn_index._remember_run(data, result, SimpleNamespace(delta=None))
    assert data["bn"] == {"NUMBER": "5", "KSP_MGRS": "31U DQ 40900 10900", "ZKP_MGRS": "31U DQ 40800 10800",
                          "MAP_EDITION": "2023"}
    assert saved == [data]


def test_previous_values_live_in_auto_file():
    data = make_data()
    data["bn"].update(PREVIOUS)
    manual, auto = split_data(data)
    assert manual["bn"] == {"NUMBER": "5", "TIME": "06.00"}
    assert auto["previous"] == PREVIOUS
