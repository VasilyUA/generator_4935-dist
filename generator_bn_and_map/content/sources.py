"""Пошук вхідних файлів у resources/: найновіше БР, найновіший ПБД з РОП_ВОП
(тека ПБД/ або тека "ДД.ММ.РРРР …"), розпорядження з розвідувального
забезпечення, позатермінові донесення з перейменуваннями, таблиця сигналів і
останній CSV-експорт карти з Дельти."""
import glob
import os
import re
from datetime import datetime

from constants import (
    BR_FILE_GLOB, DELTA_FILE_MARKER, DELTA_FOLDER_NAME, PBD_FILE_MARKERS, PBD_FOLDER_NAMES, SOURCE_FOLDER_DATE_RE,
)
from content.br_reader import read_br_header

_DATE_IN_NAME_RE = re.compile(r"(?<!\d)(\d{2}\.\d{2}\.\d{4})(?!\d)")


def _parse_date(text):
    try:
        return datetime.strptime(text, "%d.%m.%Y").date()
    except (TypeError, ValueError):
        return None


def date_in_name(name):
    """Остання дата ДД.ММ.РРРР у назві файлу чи теки, або None."""
    dates = [_parse_date(m.group(1)) for m in _DATE_IN_NAME_RE.finditer(name or "")]
    dates = [d for d in dates if d]
    return dates[-1] if dates else None


def _is_temporary(name):
    return name.startswith("~$")


def find_latest_source_folder(resources_dir):
    """-> (шлях, дата) теки "ДД.ММ.РРРР …" з найпізнішою датою, або (None, None)."""
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
        if _is_temporary(name) or not lower.endswith(extension.lower()):
            continue
        if all(marker.lower() in lower for marker in markers):
            return os.path.join(folder, name)
    return None


def _pbd_folders(resources_dir):
    """Теки, де можуть лежати додатки донесення: ПБД/ та "ДД.ММ.РРРР …"."""
    if not os.path.isdir(resources_dir):
        return []
    folders = []
    for name in sorted(os.listdir(resources_dir)):
        path = os.path.join(resources_dir, name)
        if not os.path.isdir(path):
            continue
        if name.upper() in {n.upper() for n in PBD_FOLDER_NAMES} or re.match(SOURCE_FOLDER_DATE_RE, name):
            folders.append(path)
    return folders


def find_latest_pbd(resources_dir):
    """-> (шлях до ПБД .docx, тека) найновішого ПБД або (None, None).

    Порядок: дата в назві файлу ПБД, інакше дата в назві теки, далі - час
    зміни файлу. РОП_ВОП береться з тієї ж теки (find_file)."""
    candidates = []
    for folder in _pbd_folders(resources_dir):
        for name in os.listdir(folder):
            lower = name.lower()
            if _is_temporary(name) or not lower.endswith(".docx"):
                continue
            if all(marker.lower() in lower for marker in PBD_FILE_MARKERS):
                path = os.path.join(folder, name)
                when = date_in_name(name) or date_in_name(os.path.basename(folder))
                candidates.append(((when is not None, when or datetime.min.date(), os.path.getmtime(path)), path, folder))
    if not candidates:
        return None, None
    _key, path, folder = max(candidates)
    return path, folder


def find_rop_vop(folder, markers, pbd_path=None):
    """РОП_ВОП з теки ПБД; якщо таких кілька - з тією ж датою в назві, що й ПБД."""
    if not folder or not os.path.isdir(folder):
        return None
    matches = [
        os.path.join(folder, name) for name in sorted(os.listdir(folder))
        if not _is_temporary(name) and name.lower().endswith(".xlsx")
        and all(marker.lower() in name.lower() for marker in markers)
    ]
    if len(matches) > 1 and pbd_path:
        wanted = date_in_name(os.path.basename(pbd_path))
        same_date = [path for path in matches if date_in_name(os.path.basename(path)) == wanted]
        matches = same_date or matches
    return matches[-1] if matches else None


def find_latest_by_glob(resources_dir, pattern):
    """Найновіший (за датою в назві, далі - за часом зміни) файл resources/<pattern>."""
    candidates = [
        path for path in glob.glob(os.path.join(resources_dir, pattern))
        if os.path.isfile(path) and not _is_temporary(os.path.basename(path))
    ]
    if not candidates:
        return None

    def sort_key(path):
        when = date_in_name(os.path.basename(path))
        return (when is not None, when or datetime.min.date(), os.path.getmtime(path))

    return max(candidates, key=sort_key)


def find_all_by_glob(resources_dir, pattern):
    """Усі файли resources/<pattern> (без тимчасових), від старіших до новіших."""
    paths = [
        path for path in glob.glob(os.path.join(resources_dir, pattern))
        if os.path.isfile(path) and not _is_temporary(os.path.basename(path))
    ]
    return sorted(paths, key=lambda path: (os.path.getmtime(path), path))


def find_delta_export(resources_dir):
    """Найновіший CSV-експорт Дельти: resources/Дельта/*.csv або resources/*дельта*.csv."""
    candidates = []
    folder = os.path.join(resources_dir, DELTA_FOLDER_NAME)
    if os.path.isdir(folder):
        candidates += [os.path.join(folder, n) for n in os.listdir(folder) if n.lower().endswith(".csv")]
    if os.path.isdir(resources_dir):
        candidates += [
            os.path.join(resources_dir, n) for n in os.listdir(resources_dir)
            if n.lower().endswith(".csv") and DELTA_FILE_MARKER in n.lower()
        ]
    candidates = [p for p in candidates if os.path.isfile(p) and not _is_temporary(os.path.basename(p))]
    return max(candidates, key=os.path.getmtime) if candidates else None


def find_delta_kmz(resources_dir):
    """KMZ-експорти карти в resources/Дельта/ (звідти беруться значки SIDC)."""
    folder = os.path.join(resources_dir, DELTA_FOLDER_NAME)
    if not os.path.isdir(folder):
        return []
    return sorted(os.path.join(folder, name) for name in os.listdir(folder)
                  if name.lower().endswith(".kmz") and not _is_temporary(name))


def find_latest_br(resources_dir):
    """Найновіше БР за датою з його шапки (за її відсутності - за часом зміни
    файлу)."""
    candidates = [
        path for path in glob.glob(os.path.join(resources_dir, BR_FILE_GLOB))
        if not _is_temporary(os.path.basename(path))
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
