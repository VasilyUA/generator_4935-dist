"""Кроки між запитаннями index.py - без жодного введення, щоб їх можна було
перевіряти тестами."""
import glob
import os

from constants import RENAME_FILE_GLOB, ROP_VOP_FILE_MARKERS, RR_FILE_GLOB, SIGNAL_TABLE_GLOB
from content.br_reader import read_br
from content.data_store import build_pbd_cache
from content.delta_reader import read_delta_export, read_kmz_icons
from content.pbd_reader import read_pbd
from content.rename_reader import read_renames
from content.rop_vop_reader import read_rop_vop
from content.rr_reader import read_rr
from content.signal_table import read_signal_table
from content.sources import (
    find_all_by_glob, find_delta_export, find_delta_kmz, find_latest_br, find_latest_by_glob, find_latest_pbd,
    find_rop_vop,
)
from generators.generate_bn import BnExtras


class SourceError(Exception):
    """Немає обов'язкового вхідного файлу - зрозуміле повідомлення замість traceback."""


def refresh_pbd_cache(data, resources_dir):
    """Найновіший ПБД (resources/ПБД/ або тека "ДД.ММ.РРРР …") + РОП_ВОП з тієї ж
    теки -> новий data["pbd"] (сам data не змінюється). -> (cache, warnings,
    використані файли)."""
    pbd_path, folder = find_latest_pbd(resources_dir)
    if pbd_path is None:
        raise SourceError(
            "У resources/ немає ПБД: покладіть .docx з «ПБД» у назві в теку resources/ПБД/ "
            "(або в теку, назва якої починається з дати ДД.ММ.РРРР)."
        )
    rop_vop_path = find_rop_vop(folder, ROP_VOP_FILE_MARKERS, pbd_path)

    table = read_rop_vop(rop_vop_path) if rop_vop_path else None
    cache, warnings = build_pbd_cache(read_pbd(pbd_path), table, data.get("pbd"), os.path.basename(folder))
    if table is None:
        warnings.append("Файл РОП/ВОП (.xlsx з «РОП» і «ВОП» у назві) не знайдено - кількість о/с точок не заповнено.")
    return cache, warnings, [path for path in (pbd_path, rop_vop_path) if path]


def load_latest_br(resources_dir):
    path = find_latest_br(resources_dir)
    return read_br(path) if path else None


def _read(reader, path, problems):
    try:
        return reader(path)
    except Exception as e:  # пошкоджений чи відкритий в іншій програмі файл не зупиняє генерацію
        problems.append(f"{os.path.basename(path)}: не вдалося прочитати ({e.__class__.__name__}) - пропущено.")
        return None


def load_extras(resources_dir):
    """Необов'язкові джерела: РР, позатермінові донесення, таблиця сигналів,
    експорт Дельти. -> (BnExtras, [використані файли], [проблеми])."""
    problems, used = [], []
    extras = BnExtras()
    rr_path = find_latest_by_glob(resources_dir, RR_FILE_GLOB)
    if rr_path:
        extras.rr = _read(read_rr, rr_path, problems)
    for path in find_all_by_glob(resources_dir, RENAME_FILE_GLOB):
        report = _read(read_renames, path, problems)
        if report is not None:
            extras.renames.append(report)
            used.append(path)
    signals_path = find_latest_by_glob(resources_dir, SIGNAL_TABLE_GLOB)
    if signals_path:
        extras.signals = _read(read_signal_table, signals_path, problems)
    delta_path = find_delta_export(resources_dir)
    if delta_path:
        extras.delta = _read(read_delta_export, delta_path, problems)
    used += [path for path, source in ((rr_path, extras.rr), (signals_path, extras.signals), (delta_path, extras.delta))
             if path and source is not None]
    for kmz_path in find_delta_kmz(resources_dir):
        icons = read_kmz_icons(kmz_path)
        if icons:
            used.append(kmz_path)
            for sidc, png in icons.items():
                extras.icons.setdefault(sidc, png)
    return extras, used, problems


def save_missing_icons(icons_dir, icons):
    """Значки з KMZ Дельти, яких ще немає в resources/map_icons/, -
    зберігаються туди (лишаються, навіть якщо KMZ потім прибрати). -> к-ть нових."""
    if not icons:
        return 0
    os.makedirs(icons_dir, exist_ok=True)
    saved = 0
    for sidc, png in icons.items():
        path = os.path.join(icons_dir, f"{sidc}.png")
        if not os.path.exists(path):
            with open(path, "wb") as f:
                f.write(png)
            saved += 1
    return saved


def remove_backup_files(resources_dir):
    """Видаляє резервні копії (*.bak, *.bak-*) з resources/ - у них ті самі
    дані, що й у data.json. -> [видалені файли]."""
    removed = []
    for pattern in ("*.bak", "*.bak-*"):
        for path in glob.glob(os.path.join(resources_dir, pattern)):
            if os.path.isfile(path):
                try:
                    os.remove(path)
                    removed.append(path)
                except OSError:
                    pass
    return removed
