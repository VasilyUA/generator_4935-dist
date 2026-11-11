"""Підрозділ - за назвою файлу БЧС. Реальні назви пишуть підрозділ
по-різному: з пробілом і без ("БЧС 1 X" / "БЧС.2X"), з крапками всередині
скорочення ("X.Y."), із закінченням ("Xаб"). Перелік підрозділів -
BCHS_EXPECTED_SUBDIVISIONS (resources/data.json)."""

import os
import re

from constants import BCHS_EXPECTED_SUBDIVISIONS

_NAME_WORD_RE = re.compile(r"[0-9A-Za-zА-Яа-яІіЇїЄєҐґ']+")


def _name_words(file_name):
    """Слова назви файлу (без розширення) великими літерами плюс склеєні
    сусідні пари ("1 X" -> "1X", "X.Y." -> "XY")."""
    words = _NAME_WORD_RE.findall(os.path.splitext(file_name)[0].upper())
    return words + [first + second for first, second in zip(words, words[1:])]


def _key(subdivision):
    return "".join(subdivision.upper().split())


def _names_subdivision(word, subdivision_key):
    """Чи позначає слово з назви файлу підрозділ (ключ - великими, без
    пробілів): вони однакові, або одне починає інше, якщо коротше має хоча б
    3 знаки (скорочення з закінченням: "МІН.Б." - це "мінбатр"). Дволітерні -
    лише точний збіг, інакше "XY" вважалось би й підрозділом "XYZ"."""
    if word == subdivision_key:
        return True
    shorter = min(word, subdivision_key, key=len)
    return len(shorter) >= 3 and (word.startswith(subdivision_key) or subdivision_key.startswith(word))


def file_subdivision(file_name, subdivisions=BCHS_EXPECTED_SUBDIVISIONS):
    """Підрозділ (з subdivisions), чий це файл БЧС, або None. Точний збіг
    слова має перевагу."""
    words = _name_words(file_name)
    exact = [s for s in subdivisions if _key(s) in words]
    if exact:
        return exact[0]
    return next((s for s in subdivisions if any(_names_subdivision(word, _key(s)) for word in words)), None)


def missing_subdivisions(file_names, subdivisions=BCHS_EXPECTED_SUBDIVISIONS):
    """Підрозділи subdivisions (у їхньому порядку), від яких серед file_names
    немає файлу БЧС. Після завантаження з Signal вони друкуються червоним."""
    words = [word for file_name in file_names for word in _name_words(file_name)]
    return [
        subdivision for subdivision in subdivisions
        if not any(_names_subdivision(word, _key(subdivision)) for word in words)
    ]
