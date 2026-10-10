"""Кроки між запитаннями index.py - без жодного введення, щоб їх можна було
перевіряти тестами."""
import os

from constants import PBD_FILE_MARKERS, ROP_VOP_FILE_MARKERS
from content.br_reader import read_br
from content.data_store import build_pbd_cache
from content.pbd_reader import read_pbd
from content.rop_vop_reader import read_rop_vop
from content.sources import find_file, find_latest_br, find_latest_source_folder


class SourceError(Exception):
    """Немає обов'язкового вхідного файлу - зрозуміле повідомлення замість traceback."""


def refresh_pbd_cache(data, resources_dir):
    """Найновіша тека "ДД.ММ.РРРР …" -> ПБД + РОП_ВОП -> новий data["pbd"]
    (сам data не змінюється). -> (cache, warnings, використані файли)."""
    folder, _folder_date = find_latest_source_folder(resources_dir)
    if folder is None:
        raise SourceError("У resources/ немає теки з назвою, що починається з дати ДД.ММ.РРРР (додатки донесення).")
    pbd_path = find_file(folder, PBD_FILE_MARKERS, ".docx")
    if pbd_path is None:
        raise SourceError(f"У теці «{os.path.basename(folder)}» немає .docx з «ПБД» у назві.")
    rop_vop_path = find_file(folder, ROP_VOP_FILE_MARKERS, ".xlsx")

    table = read_rop_vop(rop_vop_path) if rop_vop_path else None
    cache, warnings = build_pbd_cache(read_pbd(pbd_path), table, data.get("pbd"), os.path.basename(folder))
    if table is None:
        warnings.append("Файл РОП/ВОП (.xlsx з «РОП» і «ВОП» у назві) не знайдено - кількість о/с точок не заповнено.")
    return cache, warnings, [path for path in (pbd_path, rop_vop_path) if path]


def load_latest_br(resources_dir):
    path = find_latest_br(resources_dir)
    return read_br(path) if path else None
