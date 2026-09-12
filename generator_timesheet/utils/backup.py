import glob
import os
import shutil

from content.daily_report_reader import report_date_from_filename


def latest_report_date(report_dir):
    """Дата НАЙПІЗНІШОГО щоденного рапорту (report_date_from_filename) серед
    файлів report_dir - той самий розбір імені файлу, що й скрізь у проєкті
    (лишень файли з розпізнаваною датою в імені враховуються; lock-файли
    Excel/Word "~$..." пропускаються - той самий принцип, що й
    content/information_unit_reader._is_office_lock_file). None, якщо
    жодного такого файлу немає (порожня тека чи лише файли без дати в
    імені) - викликач тоді не має, як назвати нову підтеку."""
    dates = (
        report_date_from_filename(path)
        for path in glob.glob(os.path.join(report_dir, "*"))
        if not os.path.basename(path).startswith("~$")
    )
    dates = [d for d in dates if d is not None]
    return max(dates) if dates else None


def backup_resources_and_output(destination_root, report_dir, resources_dir, output_dir):
    """За прямою вказівкою користувача - копіює ПОТОЧНІ resources_dir і
    output_dir у НОВУ підтеку destination_root, названу датою
    НАЙПІЗНІШОГО рапорту в report_dir (latest_report_date, формат
    DD.MM.YYYY - той самий, що й у самих іменах файлів рапортів) - ручний
    "знімок" поточного стану (напр. перед початком нового місяця), а не
    заміна чи переміщення - resources/output лишаються на місці, готові до
    наступного прогону.

    dirs_exist_ok=True (copytree) - повторний запуск для ТІЄЇ САМОЇ дати
    (напр. ще раз того самого дня) домішує/перезаписує, а не падає з
    FileExistsError.

    Повертає шлях НОВОЇ підтеки, або None, якщо в report_dir немає жодного
    файлу з розпізнаваною датою (немає за чим назвати підтеку - нічого не
    копіюється)."""
    date_value = latest_report_date(report_dir)
    if date_value is None:
        return None
    target_dir = os.path.join(destination_root, date_value.strftime("%d.%m.%Y"))
    shutil.copytree(resources_dir, os.path.join(target_dir, "resources"), dirs_exist_ok=True)
    shutil.copytree(output_dir, os.path.join(target_dir, "output"), dirs_exist_ok=True)
    return target_dir
