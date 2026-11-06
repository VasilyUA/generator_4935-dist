import importlib.util
import json
import os
from pathlib import Path

from content.roster_reader import read_roster
from content.sources import load_sources


def _load_module():
    path = Path(__file__).resolve().parent.parent / "resources_example" / "generate_sample.py"
    spec = importlib.util.spec_from_file_location("duties_generate_sample", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_generate_sample_creates_readable_resources_once(tmp_path, monkeypatch, capsys):
    module = _load_module()
    target = tmp_path / "resources"
    monkeypatch.setattr(module, "RESOURCES_DIR", str(target))
    module.main()
    assert sorted(os.listdir(target)) == ["data.json", "Зразки приклад.docx", "Статут приклад.htm", "Штат приклад.xlsx"]
    data = json.loads((target / "data.json").read_text(encoding="utf-8"))
    assert data["unit"]["MILITARY_UNIT"] == "Військова частина A0000"
    assert len(read_roster(str(target), data["roster"]).slots) == 6
    sources = load_sources(str(target))
    assert sources.statute and sources.sample.blocks
    module.main()
    assert "вже існує" in capsys.readouterr().out
