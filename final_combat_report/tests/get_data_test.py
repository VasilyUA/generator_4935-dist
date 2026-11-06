import json
import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import get_data as m


# -------------------------
# get_static_data_json
# -------------------------
def test_get_static_data_json_missing_file_returns_empty_list(tmp_path):
    assert m.get_static_data_json(str(tmp_path / "no_such_file.json")) == []


def test_get_static_data_json_reads_valid_json(tmp_path):
    path = tmp_path / "data.json"
    path.write_text(json.dumps({"paragraph_one": "текст"}), encoding="utf-8")
    assert m.get_static_data_json(str(path)) == {"paragraph_one": "текст"}


def test_get_static_data_json_invalid_json_returns_empty_dict(tmp_path):
    path = tmp_path / "broken.json"
    path.write_text("{не json", encoding="utf-8")
    assert m.get_static_data_json(str(path)) == {}


# -------------------------
# get_date / get_hour_of_report / get_actual_data_json (через FakePrompt з conftest)
# -------------------------
def test_get_date_uses_default_on_enter(inquirer_inputs):
    from datetime import datetime
    result = m.get_date(18)
    assert result == datetime.now().strftime("%d.%m.%Y")


def test_get_date_returns_scripted_input(inquirer_inputs):
    inquirer_inputs.append("24.07.2026")
    assert m.get_date(18) == "24.07.2026"


def test_get_hour_of_report_default_is_18(inquirer_inputs):
    assert m.get_hour_of_report() == 18


def test_get_hour_of_report_returns_scripted_input(inquirer_inputs):
    inquirer_inputs.append("20")
    assert m.get_hour_of_report() == 20


def test_get_actual_data_json_default_is_false(inquirer_inputs):
    assert m.get_actual_data_json() is False


def test_get_actual_data_json_returns_scripted_input(inquirer_inputs):
    inquirer_inputs.append(True)
    assert m.get_actual_data_json() is True


# -------------------------
# run_signal_export
# -------------------------
def test_run_signal_export_raises_without_appdata(monkeypatch):
    monkeypatch.delenv("APPDATA", raising=False)
    import pytest
    with pytest.raises(EnvironmentError):
        m.run_signal_export("Basketball")


def test_run_signal_export_invokes_subprocess_with_expected_args(monkeypatch, tmp_path, capsys):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    captured = {}

    class FakeResult:
        stdout = "ok"
        stderr = ""

    def fake_run(cmd, capture_output, text):
        captured["cmd"] = cmd
        return FakeResult()

    monkeypatch.setattr(m.subprocess, "run", fake_run)

    m.run_signal_export("Кречет 2.0", "Пейнтбол")

    assert "--chats=Кречет 2.0,Пейнтбол" in captured["cmd"]
    assert "--overwrite" in captured["cmd"]
    assert "signal_json" in captured["cmd"]
    assert "ok" in capsys.readouterr().out
