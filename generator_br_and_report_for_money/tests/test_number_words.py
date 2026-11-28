import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import pytest

from utils.number_words import number_to_words_genitive, sheets_locative_phrase


@pytest.mark.parametrize("number, words", [
    (1, "одного"), (2, "двох"), (10, "десяти"), (11, "одинадцяти"), (19, "дев’ятнадцяти"),
    (20, "двадцяти"), (40, "сорока"), (90, "дев’яноста"), (100, "ста"), (118, "ста вісімнадцяти"),
    (133, "ста тридцяти трьох"), (200, "двохсот"), (999, "дев’ятисот дев’яноста дев’яти"),
    (1000, "однієї тисячі"), (2001, "двох тисяч одного"), (5000, "п’яти тисяч"),
])
def test_number_to_words_genitive(number, words):
    assert number_to_words_genitive(number) == words


def test_number_to_words_genitive_rejects_out_of_range():
    with pytest.raises(ValueError):
        number_to_words_genitive(0)


@pytest.mark.parametrize("number, phrase", [
    (118, "118 (ста вісімнадцяти)"),
    (133, "133 (ста тридцяти трьох)"),
])
def test_sheets_locative_phrase(number, phrase):
    """Реальні приклади - "133 (ста тридцяти трьох) аркушах" (користувач),
    "118 (ста вісімнадцяти) аркушах" (tests/test_report_document_reader.py) -
    саме слово "аркушах" тут не входить (лишається незмінним у шаблоні,
    generators/word_finalize.py)."""
    assert sheets_locative_phrase(number) == phrase
