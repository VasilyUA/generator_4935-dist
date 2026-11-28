import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from docx import Document

import constants
import template.sanitation
# template/__init__.py робить "from .sanitation import sanitation", тож
# ОДНОЙМЕННИЙ атрибут пакета template.sanitation переписується на функцію -
# дістаємо справжній модуль напряму із sys.modules, а не через атрибут пакета.
m = sys.modules["template.sanitation"]


def _soldier():
    return {
        "position_full_dative": "Стрільцю",
        "rank_fact_dative": "сержанту",
        "name_dative": "ДВАНАДЦЯТИЙ Дванадцятий Дванадцятий",
        "name_nominative": "ДВАНАДЦЯТИЙ Дванадцятий Дванадцятий",
    }


def _higher_commander():
    return {
        "position_full_nominative": "Командир батальйону",
        "rank_fact_nominative": "підполковник",
        "name_nominative": "ТРИНАДЦЯТИЙ Тринадцятий Тринадцятий",
    }


def test_sanitation_generates_document(tmp_path, monkeypatch):
    monkeypatch.setattr(m, "OUTPUT_DIR", str(tmp_path))

    doc_name = m.sanitation(_soldier(), _higher_commander())

    assert doc_name == str(tmp_path / "ДВАНАДЦЯТИЙ Дванадцятий Дванадцятий на оздоровчі.docx")
    text = "\n".join(p.text for p in Document(doc_name).paragraphs)
    assert "грошову допомогу на оздоровлення" in text
    assert constants.FULL_MILITARY_UNIT in text
    assert "Тринадцятий ТРИНАДЦЯТИЙ" in text
