"""Пошук вхідних файлів у resources/: найновіше БР і найновіша тека з
додатками донесення ("ДД.ММ.РРРР <будь-що>") з ПБД і РОП_ВОП усередині."""
import glob
import os
import re
from datetime import datetime

from constants import BR_FILE_GLOB, SOURCE_FOLDER_DATE_RE
from content.br_reader import read_br_header


def _parse_date(text):
    try:
        return datetime.strptime(text, "%d.%m.%Y").date()
    except (TypeError, ValueError):
        return None


def find_latest_source_folder(resources_dir):
    """-> (шлях, дата) теки з найпізнішою датою в назві, або (None, None)."""
    best = (None, None)
    if not os.path.isdir(resources_dir):
        return best
    for name in os.listdir(resources_dir):
        path = os.path.join(resources_dir, name)
        m = re.match(SOURCE_FOLDER_DATE_RE, name)
        if not (m and os.path.isdir(path)):
            continue
        folder_date = _parse_date(m.group(1))
        if folder_date and (best[1] is None or folder_date > best[1]):
            best = (path, folder_date)
    return best


def find_file(folder, markers, extension):
    """Перший файл з розширенням extension, у назві якого є ВСІ markers (без
    урахування регістру); тимчасові "~$…" файли Word/Excel пропускаються."""
    if not folder or not os.path.isdir(folder):
        return None
    for name in sorted(os.listdir(folder)):
        lower = name.lower()
        if name.startswith("~$") or not lower.endswith(extension.lower()):
            continue
        if all(marker.lower() in lower for marker in markers):
            return os.path.join(folder, name)
    return None


def find_latest_br(resources_dir):
    """Найновіше БР за датою з його шапки (за її відсутності - за часом зміни
    файлу)."""
    candidates = [
        path for path in glob.glob(os.path.join(resources_dir, BR_FILE_GLOB))
        if not os.path.basename(path).startswith("~$")
    ]
    if not candidates:
        return None

    def sort_key(path):
        try:
            header_date = _parse_date(read_br_header(path).date)
        except Exception:
            header_date = None
        return (header_date is not None, header_date or datetime.min.date(), os.path.getmtime(path))

    return max(candidates, key=sort_key)
