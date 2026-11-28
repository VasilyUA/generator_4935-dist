"""Ручний "знімок" поточних resources/ і output/ в іншу теку - за зразком
generator_timesheet/utils/backup.py: копіювання, а не переміщення
(resources/ і output/ лишаються на місці, готові до наступного запуску)."""
import os
import re
import shutil
from datetime import date

# Файли-замки Word/Excel ("~$назва.docx") поки документ відкритий - копіювати їх
# немає сенсу, а відкритий замок зірвав би копіювання.
_SKIP = shutil.ignore_patterns("~$*")


def snapshot_folder_name(bn_date=None, report_datetime="", today=None):
    """Назва підтеки (дд.мм.рррр): дата бойового наказу; якщо БН цього разу не
    генерувався (лише карта) - дата ПБД; інакше - сьогоднішня."""
    if bn_date is not None:
        return bn_date.strftime("%d.%m.%Y")
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", report_datetime or "")
    if m:
        return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
    return (today or date.today()).strftime("%d.%m.%Y")


def backup_resources_and_output(destination_root, folder_name, resources_dir, output_dir):
    """Копіює resources_dir і output_dir у destination_root/folder_name/
    (resources, output). dirs_exist_ok=True - повторне збереження з тією самою
    назвою (напр. ще раз того самого дня) домішує й перезаписує, а не падає.
    -> шлях нової підтеки."""
    target_dir = os.path.join(destination_root, folder_name)
    shutil.copytree(resources_dir, os.path.join(target_dir, "resources"), dirs_exist_ok=True, ignore=_SKIP)
    shutil.copytree(output_dir, os.path.join(target_dir, "output"), dirs_exist_ok=True, ignore=_SKIP)
    return target_dir
