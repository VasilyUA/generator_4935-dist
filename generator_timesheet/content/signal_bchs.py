"""Завантаження файлів БЧС за обрану дату з групи Signal BCHS_SIGNAL_GROUP_NAME.
index.py питає про це окремо ("Отримати файли БЧС?" і дата), очищає
INFORMATION_UNIT_DIR (clear_information_unit_directory) і кладе туди щойно
завантажені файли.

Чому не через сам пакет sigexport (як final_combat_report/get_data.py):
sigexport==3.6.0 шукає вкладення в старій схемі Signal Desktop (ключ
"attachments" у JSON повідомлення), а встановлений Signal уже тримає їх в
окремій таблиці message_attachments. Повний експорт групи через sigexport
дав 0 вкладень на 6539 повідомлень, хоча в 4835 із них
message_attachments.hasAttachments=1. Тому таблиця читається напряму, але
ключ бази й розшифрування вкладень - з sigexport (crypto і
files.decrypt_attachment), а не власна копія AES/HMAC."""

import os
import re
import shutil
from datetime import datetime
from pathlib import Path

from sigexport import crypto
from sigexport.files import decrypt_attachment
from sqlcipher3 import dbapi2

from constants import BCHS_FILENAME_MARKER, BCHS_SIGNAL_GROUP_NAME
from content.bchs_subdivisions import file_subdivision
from content.information_unit_reader import _SUPPORTED_EXTENSIONS
from utils.logging_utils import print_red, print_yellow


class SignalFetchError(Exception):
    """Очікувана помилка, яку треба показати червоним і йти далі: Signal
    Desktop не встановлено чи не розблоковано, групу не знайдено тощо.
    Непередбачені баги натомість мають лишатися видимим traceback'ом."""


def _signal_source_dir():
    appdata = os.getenv("APPDATA")
    if not appdata:
        raise SignalFetchError("Змінна середовища APPDATA не встановлена - Signal Desktop шукати нема де.")
    source_dir = os.path.join(appdata, "Signal")
    if not os.path.isfile(os.path.join(source_dir, "config.json")):
        raise SignalFetchError(f"Не знайдено встановлений Signal Desktop у {source_dir}.")
    return source_dir


def _open_db(source_dir):
    try:
        # crypto.get_key робить appdir / "config.json", тож потрібен Path, а
        # не рядок (з рядком - TypeError).
        key = crypto.get_key(Path(source_dir), None)
    except Exception as e:
        raise SignalFetchError(f"Не вдалося розшифрувати ключ бази Signal Desktop: {e}") from e
    db = dbapi2.connect(os.path.join(source_dir, "sql", "db.sqlite"))
    cursor = db.cursor()
    # Ті самі PRAGMA, що й у sigexport.data.fetch_data; їх треба виконувати
    # одразу після кожного підключення.
    cursor.execute(f"PRAGMA KEY = \"x'{key}'\"")
    cursor.execute("PRAGMA cipher_page_size = 4096")
    cursor.execute("PRAGMA kdf_iter = 64000")
    cursor.execute("PRAGMA cipher_hmac_algorithm = HMAC_SHA512")
    cursor.execute("PRAGMA cipher_kdf_algorithm = PBKDF2_HMAC_SHA512")
    return db, cursor


def _find_conversation_id(cursor, group_name):
    """conversations.name - назва групи. profileName - поле, за яким
    sigexport.data.fetch_data шукає --chats; у групи воно зазвичай порожнє,
    але перевіряється про всяк випадок."""
    cursor.execute("SELECT id FROM conversations WHERE name = ? OR profileName = ?", (group_name, group_name))
    row = cursor.fetchone()
    return row[0] if row else None


def _contains_marker(file_name, marker):
    """Без урахування регістру: назви файлів люди пишуть вручну
    ("БЧС_Підрозділ1 07.10.26.xlsx", "БЧС.Підрозділ2 07.10.2026.xlsx", "БЧС  Підрозділ_3 ...")."""
    return marker.upper() in (file_name or "").upper()


def _sent_on_date(sent_at_ms, target_date):
    """sentAt - мілісекунди Unix-часу. Порівнюється локальна дата
    надсилання, а не дата в назві файлу: її пишуть по-різному
    ("07.10.2026"/"07.10.26"/"07.10", із зайвими пробілами чи без неї взагалі),
    а sentAt завжди однозначний."""
    if sent_at_ms is None:
        return False
    return datetime.fromtimestamp(sent_at_ms / 1000).date() == target_date


def _has_supported_extension(file_name):
    """Чи вміє information_unit_reader прочитати такий файл (.xlsx/.xlsm/.xls).
    Перелік береться звідти ж (_SUPPORTED_EXTENSIONS), щоб не тримати копію."""
    return os.path.splitext(file_name or "")[1].lower() in _SUPPORTED_EXTENSIONS


def _candidate_rows(cursor, conversation_id, target_date, marker=BCHS_FILENAME_MARKER):
    """[{fileName, sentAt, path, localKey, size, version}, ...] - вкладення
    цієї розмови, які information_unit_reader уміє прочитати, з marker у назві
    й надіслані target_date. Вкладення, які Signal позначив як непридатні
    (error) чи ще не завантажив на цей комп'ютер (path IS NULL), пропускаються -
    їх нема що розшифровувати."""
    cursor.execute(
        """
        SELECT fileName, sentAt, path, localKey, size, version
        FROM message_attachments
        WHERE conversationId = ?
          AND error IS NULL
          AND path IS NOT NULL
        """,
        (conversation_id,),
    )
    rows = []
    for file_name, sent_at, path, local_key, size, version in cursor.fetchall():
        if not _has_supported_extension(file_name):
            continue
        if not _contains_marker(file_name, marker):
            continue
        if not _sent_on_date(sent_at, target_date):
            continue
        rows.append({
            "fileName": file_name, "sentAt": sent_at, "path": path,
            "localKey": local_key, "size": size, "version": version,
        })
    return rows


# Усе, що не буква й не цифра (пробіли, підкреслення, крапки тощо) - для
# _dedupe_key. Явний перелік літер, а не "\W+": той залежить від locale і
# ненадійний для кирилиці.
_NON_ALNUM_RE = re.compile(r"[^0-9A-Za-zА-Яа-яІіЇїЄєҐґ]+")


def _dedupe_key(file_name):
    """Назва файлу без розділювачів і регістру - для файлів, чий підрозділ не
    розпізнано (_dedupe_keep_latest). Назви пишуть вручну, тож "БЧС  Підрозділ_3
    16.10.2026.xlsx" (подвійний пробіл) і "БЧС Підрозділ_3_16.10.2026.xlsx"
    (підкреслення перед датою) - той самий файл. Інші літери чи цифри дають
    інші ключі."""
    return _NON_ALNUM_RE.sub("", file_name or "").upper()


def _dedupe_keep_latest(rows, target_date):
    """Один файл від підрозділу за target_date (рішення користувача): підрозділ
    - за назвою файлу (file_subdivision). Назва, де підрозділ не розпізнано,
    порівнюється лише з такою самою (_dedupe_key), тож чужий файл її не витіснить.

    Лишається файл із цією датою в назві, а з кількох таких - найсвіжіший
    (найбільший sentAt). Просто найсвіжіший не підходить, бо за день надходять і
    виправлення за вчора. Реальні випадки: "БЧС X 31.10.2026.xlsx" і пізніше
    "БЧС X 31.10.2026 Виправленний.xlsx" - лишається другий; "БЧС X
    01.10.2026.xlsx" о 16:08 і "БЧС X 30.09.26 ВИПРАВЛЕННИЙ.xlsx" о 18:06 -
    лишається перший. Файла з цією датою немає - найсвіжіший. Той самий файл
    кількома рядками - це відредаговане в Signal повідомлення
    (editHistoryIndex -1/0/1 на одне вкладення).

    Файл, який не взято, друкується жовтим: це не помилка, але мовчки його
    губити не можна. Той самий файл, надісланий повторно з іншими пробілами чи
    підкресленнями (_dedupe_key), - без повідомлення: таке буває щотижня."""
    day_mark = target_date.strftime("%d.%m")
    groups = {}
    for row in rows:
        groups.setdefault(file_subdivision(row["fileName"]) or _dedupe_key(row["fileName"]), []).append(row)

    kept = []
    for group in groups.values():
        best = max(group, key=lambda row: (day_mark in row["fileName"], row["sentAt"]))
        skipped = sorted({row["fileName"] for row in group if _dedupe_key(row["fileName"]) != _dedupe_key(best["fileName"])})
        if skipped:
            print_yellow(f"Кілька файлів БЧС від одного підрозділу: взято {best['fileName']}, пропущено {', '.join(skipped)}")
        kept.append(best)
    return kept


def _save_attachment(source_dir, dest_dir, row):
    """Копіює чи розшифровує одне вкладення в dest_dir під його власним
    fileName. На відміну від sigexport.files.copy_attachments, не
    перейменовує на "<дата>_<індекс>_<ім'я>": information_unit_reader читає
    файли за вмістом. Повертає False (з причиною червоним), а не піднімає
    виняток, - одне непридатне вкладення не має зупиняти решту."""
    return _decrypt_to(source_dir, row, os.path.join(dest_dir, row["fileName"]))


def _decrypt_to(source_dir, row, dst_path):
    """Розшифровує чи копіює одне вкладення в dst_path (будь-який шлях, напр.
    тимчасовий файл для порівняння з наявним у content/signal_reports.py).
    False (з причиною червоним), якщо не вдалося."""
    normalized_path = row["path"].replace("\\", "/").replace("/", os.sep)
    src_path = os.path.join(source_dir, "attachments.noindex", normalized_path)
    if int(row.get("version") or 0) >= 2:
        try:
            decrypt_attachment(row, Path(src_path), Path(dst_path))
        except ValueError as e:
            print_red(f"Не вдалося розшифрувати {row['fileName']}: {e}")
            return False
    else:
        try:
            shutil.copy2(src_path, dst_path)
        except (FileNotFoundError, OSError) as e:
            print_red(f"Не вдалося скопіювати {row['fileName']}: {e}")
            return False
    return True


def fetch_bchs_files(target_date, dest_dir, group_name=BCHS_SIGNAL_GROUP_NAME):
    """Завантажує в dest_dir вкладення групи Signal group_name, які
    information_unit_reader уміє прочитати, мають "БЧС" у назві й надіслані
    target_date (datetime.date), - по одному від підрозділу
    (_dedupe_keep_latest). Повертає імена успішно збережених файлів.
    Якщо групу не знайдено, Signal не встановлено чи не розблоковано тощо -
    SignalFetchError: викликач друкує її і йде далі."""
    source_dir = _signal_source_dir()
    db, cursor = _open_db(source_dir)
    try:
        conversation_id = _find_conversation_id(cursor, group_name)
        if conversation_id is None:
            raise SignalFetchError(f"У Signal не знайдено чат/групу \"{group_name}\".")
        rows = _dedupe_keep_latest(_candidate_rows(cursor, conversation_id, target_date), target_date)
    finally:
        db.close()

    saved = []
    for row in rows:
        if _save_attachment(source_dir, dest_dir, row):
            saved.append(row["fileName"])
    return saved
