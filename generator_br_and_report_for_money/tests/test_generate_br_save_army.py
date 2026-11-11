import os
import sys
from datetime import datetime
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from docx import Document

import generators.generate_br_save_army as save_army_module


def test_generate_documents_br_save_skips_when_no_safety_number_configured(tmp_path, monkeypatch, capsys):
    """Немає номера бойового розпорядження застосування безпеки на цю дату
    (NUMBER_OF_DOCUMENTS_BRS_SAVE без запису чи з порожніми полями) - документ
    НЕ генерується (жодного файлу), лише повідомлення в консоль."""
    monkeypatch.setattr(save_army_module, "NUMBER_OF_DOCUMENTS_BRS_SAVE", {})

    save_army_module.generate_documents_br_save(
        Document(), datetime(2026, 7, 1), {"ЗВАННЯ": "підполковник", "ПІБ": "ПЕРШИЙ Перший Перший", "ТВО": False},
        "КИЇВ", "36T TT 12345 67890", str(tmp_path),
    )

    assert os.listdir(tmp_path) == []
    assert "Відсутній номер" in capsys.readouterr().out


def test_generate_documents_br_save_writes_full_document(tmp_path, monkeypatch):
    monkeypatch.setattr(save_army_module, "NUMBER_OF_DOCUMENTS_BRS_SAVE", {
        "01.07.2026": {"бз_бат": "91", "бз_брг": "500"},
    })

    save_army_module.generate_documents_br_save(
        Document(), datetime(2026, 7, 1), {"ЗВАННЯ": "підполковник", "ПІБ": "ПЕРШИЙ Перший Перший", "ТВО": True},
        "КИЇВ", "36T TT 12345 67890", str(tmp_path),
    )

    files = os.listdir(tmp_path)
    assert any(f.endswith("застосування безпеки 01.07.2026.docx") for f in files)
