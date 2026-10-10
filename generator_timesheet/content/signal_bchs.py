"""Завантаження файлів БЧС з групи Signal BCHS_SIGNAL_GROUP_NAME за
обрану дату - за прямою вказівкою користувача: окреме запитання перед
звичним запуском index.py ("отримати файли БЧС?" + дата), після чого
INFORMATION_UNIT_DIR очищається (utils/output_folder.py,
clear_information_unit_directory) і заповнюється щойно завантаженими
файлами.

Чому НЕ через сам пакет sigexport (як final_combat_report/get_data.py,
run_signal_export) - на цій машині емпірично підтверджено: sigexport==3.6.0
читає вкладення зі СТАРОЇ схеми Signal Desktop (JSON-блоб повідомлення ->
ключ "attachments"), а встановлений тут Signal Desktop ВЖЕ переніс
вкладення в окрему таблицю message_attachments - повний експорт групи
ВІДПОЧИНОК/ЕКСПЕДИЦІЇ самим sigexport дав 0 вкладень при 6539 повідомленнях,
хоча message_attachments.hasAttachments=1 для 4835 з них. Тож ця таблиця
читається напряму (той самий ключ розшифрування БД - sigexport.crypto -
і те саме, вже протестоване розшифрування самого вкладення -
sigexport.files.decrypt_attachment, а не власна копія AES/HMAC)."""

import os
import shutil
from datetime import datetime
from pathlib import Path

from sigexport import crypto
from sigexport.files import decrypt_attachment
from sqlcipher3 import dbapi2

from constants import BCHS_FILENAME_MARKER, BCHS_SIGNAL_GROUP_NAME
from content.information_unit_reader import _SUPPORTED_EXTENSIONS
from utils.logging_utils import print_red


class SignalFetchError(Exception):
    """Очікувана помилка ("показати червоним і продовжити роботу скрипта"):
    Signal Desktop не встановлено/не розлочено на цій машині, групу не
    знайдено тощо - на відміну від непередбаченого бага, який мусить лишитись
    видимим traceback'ом."""


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
        # crypto.get_key робить appdir / "config.json" (pathlib) - рядок
        # (os.path.join) тут НЕ підійде (TypeError), на відміну від решти
        # цього модуля, де source_dir лишається звичайним рядком.
        key = crypto.get_key(Path(source_dir), None)
    except Exception as e:
        raise SignalFetchError(f"Не вдалося розшифрувати ключ бази Signal Desktop: {e}") from e
    db = dbapi2.connect(os.path.join(source_dir, "sql", "db.sqlite"))
    cursor = db.cursor()
    # Той самий набір PRAGMA, що й sigexport.data.fetch_data - потрібен
    # щоразу одразу після підключення, це НЕ налаштування "раз і назавжди".
    cursor.execute(f"PRAGMA KEY = \"x'{key}'\"")
    cursor.execute("PRAGMA cipher_page_size = 4096")
    cursor.execute("PRAGMA kdf_iter = 64000")
    cursor.execute("PRAGMA cipher_hmac_algorithm = HMAC_SHA512")
    cursor.execute("PRAGMA cipher_kdf_algorithm = PBKDF2_HMAC_SHA512")
    return db, cursor


def _find_conversation_id(cursor, group_name):
    """conversations.name - заголовок групи; profileName - те саме поле,
    яким sigexport.data.fetch_data узгоджує --chats (для групи зазвичай
    порожнє, звіряється про всяк випадок - той самий принцип)."""
    cursor.execute("SELECT id FROM conversations WHERE name = ? OR profileName = ?", (group_name, group_name))
    row = cursor.fetchone()
    return row[0] if row else None


def _contains_marker(file_name, marker):
    """Без урахування регістру - реальні назви пишуть люди вручну
    ("БЧС_ВБАК 07.10.26.xlsx", "БЧС.2РМП 07.10.2026.xlsx", "БЧС  РВ_1БМП ...")."""
    return marker.upper() in (file_name or "").upper()


def _sent_on_date(sent_at_ms, target_date):
    """sentAt - мілісекунди Unix-часу; порівнюємо за ЛОКАЛЬНОЮ календарною
    датою відправлення, а НЕ за датою, вписаною в текст назви файлу - люди
    пишуть її по-різному ("07.10.2026"/"07.10.26"/"07.10"/із зайвими
    пробілами чи взагалі без неї), а sentAt завжди однозначний."""
    if sent_at_ms is None:
        return False
    return datetime.fromtimestamp(sent_at_ms / 1000).date() == target_date


def _has_supported_extension(file_name):
    """_SUPPORTED_EXTENSIONS (content/information_unit_reader.py) - ТЕ САМЕ
    ОДНЕ джерело правди, що визначає, які файли information_unit_reader
    узагалі здатен прочитати (.xlsx/.xlsm/.xls) - не власний, окремий
    список тут, який довелось би синхронізувати вручну."""
    return os.path.splitext(file_name or "")[1].lower() in _SUPPORTED_EXTENSIONS


def _candidate_rows(cursor, conversation_id, target_date, marker=BCHS_FILENAME_MARKER):
    """[{fileName, sentAt, path, localKey, size, version}, ...] - вкладення
    цієї розмови з розширенням, яке вміє прочитати information_unit_reader
    (_has_supported_extension), чия назва містить marker і дата
    відправлення дорівнює target_date. error IS NULL/path IS NOT NULL -
    вкладення, які сам Signal позначив як непридатні чи ще не завантажені
    на цей комп'ютер, пропускаються (інакше потім нема що
    розшифровувати/копіювати)."""
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


def _dedupe_keep_latest(rows):
    """За прямою вказівкою користувача - якщо файли за цю саму дату
    дублюються, береться актуальніший. "Дублюються" - ТОЧНА (verbatim)
    назва файлу повторюється серед rows (реальний випадок, підтверджений
    емпірично: повідомлення, відредаговане в Signal, лишає в
    message_attachments кілька рядків - editHistoryIndex -1/0/1 - для
    ОДНОГО й того самого вкладення) - серед однакових назв лишається ЛИШЕ
    та, що з найбільшим sentAt. Файли з РІЗНИМИ назвами (різні підрозділи)
    - не дублікати один одного, лишаються всі."""
    latest_by_name = {}
    for row in rows:
        current = latest_by_name.get(row["fileName"])
        if current is None or row["sentAt"] > current["sentAt"]:
            latest_by_name[row["fileName"]] = row
    return list(latest_by_name.values())


def _save_attachment(source_dir, dest_dir, row):
    """Копіює/розшифровує ОДНЕ вкладення під його власним (сирим) fileName
    у dest_dir - на відміну від sigexport.files.copy_attachments, НЕ
    перейменовує на "<дата>_<індекс>_<ім'я>" (information_unit_reader.py
    читає файли information_unit_dir за змістом, довільна назва з
    розширенням .xlsx цілком підходить). Повертає False (і друкує причину
    червоним), а не підіймає виняток, - ОДНЕ непридатне вкладення не має
    зупиняти завантаження решти файлів за цю дату."""
    normalized_path = row["path"].replace("\\", "/").replace("/", os.sep)
    src_path = os.path.join(source_dir, "attachments.noindex", normalized_path)
    dst_path = os.path.join(dest_dir, row["fileName"])
    if int(row.get("version") or 0) >= 2:
        try:
            decrypt_attachment(row, src_path, dst_path)
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
    """Повертає список імен файлів, успішно завантажених у dest_dir - лише
    вкладення групи Signal group_name з розширенням, яке вміє прочитати
    information_unit_reader (.xlsx/.xlsm/.xls), чия назва містить "БЧС" і
    чия дата відправлення дорівнює target_date (datetime.date). Групу не
    знайдено, Signal Desktop не встановлено/не розлочено тощо -
    SignalFetchError (викликач друкує її як звичайне повідомлення й веде
    скрипт далі, а не падає traceback'ом)."""
    source_dir = _signal_source_dir()
    db, cursor = _open_db(source_dir)
    try:
        conversation_id = _find_conversation_id(cursor, group_name)
        if conversation_id is None:
            raise SignalFetchError(f"У Signal не знайдено чат/групу \"{group_name}\".")
        rows = _dedupe_keep_latest(_candidate_rows(cursor, conversation_id, target_date))
    finally:
        db.close()

    saved = []
    for row in rows:
        if _save_attachment(source_dir, dest_dir, row):
            saved.append(row["fileName"])
    return saved
