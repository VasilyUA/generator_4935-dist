"""Збереження результату (output/) в іншу теку - копіювання, а не
переміщення (як generator_bn_and_map/utils/backup.py)."""
import os
import shutil

# Файли-замки Word/Excel ("~$назва.docx") поки документ відкритий - копіювати їх
# немає сенсу, а відкритий замок зірвав би копіювання.
_SKIP = shutil.ignore_patterns("~$*")


def save_output(output_dir, destination_root, folder_name):
    """Копіює вміст output_dir у destination_root/folder_name/ (повторне
    збереження з тією самою назвою перезаписує файли). → шлях підтеки."""
    target_dir = os.path.join(destination_root, folder_name)
    shutil.copytree(output_dir, target_dir, dirs_exist_ok=True, ignore=_SKIP)
    return target_dir
