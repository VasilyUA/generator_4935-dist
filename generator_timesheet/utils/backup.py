import os
import shutil

from content.daily_report_reader import latest_report_date


def backup_resources_and_output(destination_root, report_dir, resources_dir, output_dir):
    """Копіює поточні resources_dir і output_dir у нову підтеку destination_root,
    названу датою найпізнішого рапорту в report_dir (дд.мм.рррр, як у назвах
    рапортів). Це ручний знімок стану (напр. перед новим місяцем): самі теки
    лишаються на місці для наступного прогону. Повторний запуск за ту саму
    дату дописує й перезаписує файли, а не падає.

    Повертає шлях нової підтеки або None, якщо в report_dir немає жодного
    рапорту з датою в назві (нема як назвати підтеку - нічого не копіюється)."""
    date_value = latest_report_date(report_dir)
    if date_value is None:
        return None
    target_dir = os.path.join(destination_root, date_value.strftime("%d.%m.%Y"))
    shutil.copytree(resources_dir, os.path.join(target_dir, "resources"), dirs_exist_ok=True)
    shutil.copytree(output_dir, os.path.join(target_dir, "output"), dirs_exist_ok=True)
    return target_dir
