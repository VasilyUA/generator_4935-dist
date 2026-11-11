"""Звіряння resources/report (щоденні рапорти) і resources/schedule (відомості
ВОП-РОП) з групою Signal REPORTS_SIGNAL_GROUP_NAME (resources/data.json) -
за питанням "Отримати файли resources/report та resources/schedule?" на
старті index.py. База й вкладення читаються так само, як у
content/signal_bchs.py.

- Рапорти: "ДД.ММ.РРРР - щоденний рапорт.docx" (буває "_-_" замість " - "),
  дата з назви, починаючи з останнього дня минулого місяця. Зберігаються під
  звичною назвою "ДД.ММ.РРРР - щоденний рапорт.docx".
- Відомості: ".xlsx" з "ВОП-РОП" і датою подання в назві, з першого числа
  місяця. Зберігаються під назвою, з якою їх надіслали.
- Той самий файл буває надіслано кілька разів - береться найновіший за часом
  надсилання.
- Файл за цю дату вже є в теці: якщо його збережено після надсилання
  найновішої версії (завантажено раніше чи виправлено вручну - такі рапорти
  справді трапляються) і це не точна копія однієї з попередніх версій, він
  актуальний і лишається. Так само лишається файл з тим самим вмістом, що й
  у найновішої версії. Інакше (у Signal з'явилася новіша версія чи збережено
  не ту) файл замінюється, а старий переїжджає в
  SIGNAL_REPLACED_VERSIONS_DIR_NAME, якої генератори не бачать. Якщо файлу
  немає - він завантажується. Файлів теки, яких немає в Signal, ніхто не
  чіпає."""

import filecmp
import glob
import os
import re
import shutil
import tempfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

from constants import REPORT_DIR, REPORTS_SIGNAL_GROUP_NAME, SCHEDULE_DIR, SIGNAL_REPLACED_VERSIONS_DIR_NAME
from content import signal_bchs
from content.information_unit_reader import _is_office_lock_file
from content.signal_bchs import SignalFetchError

_DATE_IN_NAME_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{4}|\d{2})(?!\d)")
_REPORT_MARKER_RE = re.compile(r"щоденн\w*[\s_]+рапорт", re.IGNORECASE)
_SCHEDULE_MARKER = "ВОП-РОП"


@dataclass
class _Kind:
    label: str
    dest_dir: str
    extensions: tuple
    matches: object  # (назва файлу) -> bool
    first_date: date
    canonical_name: object  # (дата, назва в Signal чи None) -> назва файлу в теці


@dataclass
class SyncResult:
    label: str
    new: list = field(default_factory=list)
    updated: list = field(default_factory=list)
    actual: list = field(default_factory=list)
    failed: list = field(default_factory=list)
    missing: list = field(default_factory=list)  # [(дата, очікувана назва файлу)]


def name_date(file_name):
    """Перша дата "дд.мм.рррр" (чи "дд.мм.рр") у назві файлу або None."""
    match = _DATE_IN_NAME_RE.search(file_name or "")
    if match is None:
        return None
    day, month, year = match.groups()
    try:
        return date(int(year) + (2000 if len(year) == 2 else 0), int(month), int(day))
    except ValueError:
        return None


def _is_report(file_name):
    return bool(_REPORT_MARKER_RE.search(file_name or ""))


def _is_schedule(file_name):
    return _SCHEDULE_MARKER in (file_name or "").upper()


def _kinds(today, report_dir=REPORT_DIR, schedule_dir=SCHEDULE_DIR):
    month_start = date(today.year, today.month, 1)
    return [
        _Kind("рапорти", report_dir, (".docx",), _is_report, month_start - timedelta(days=1),
              lambda day, _signal_name: f"{day:%d.%m.%Y} - щоденний рапорт.docx"),
        _Kind("відомості ВОП-РОП", schedule_dir, (".xlsx", ".xlsm", ".xls"), _is_schedule, month_start,
              lambda day, signal_name: (os.path.basename(signal_name) if signal_name
                                        else f"{_SCHEDULE_MARKER} {day:%d.%m.%Y}.xlsx")),
    ]


def _with_date(file_name, day):
    """file_name, у якому перша дата замінена на day (рік - стільки ж цифр,
    скільки було: 4 чи 2)."""
    def replace(match):
        return f"{day:%d.%m.}" + (f"{day:%Y}" if len(match.group(3)) == 4 else f"{day:%y}")
    return _DATE_IN_NAME_RE.sub(replace, file_name, count=1)


def versions_by_date(rows, kind):
    """{дата з назви: вкладення цього виду від найстарішого до найновішого за
    sentAt}, лише від kind.first_date. При однаковому sentAt (рядки історії
    редагування повідомлення) найновішою вважається поточна версія
    (editHistoryIndex -1)."""
    versions = {}
    for row in rows:
        file_name = row.get("fileName") or ""
        if os.path.splitext(file_name)[1].lower() not in kind.extensions or not kind.matches(file_name):
            continue
        day = name_date(file_name)
        if day is None or day < kind.first_date:
            continue
        versions.setdefault(day, []).append(row)
    for day_versions in versions.values():
        day_versions.sort(key=lambda row: (row.get("sentAt") or 0, row.get("editHistoryIndex", -1) == -1))
    return versions


def _local_files(kind):
    """Наявні в теці файли цього виду з датою в назві, за абеткою."""
    for path in sorted(glob.glob(os.path.join(kind.dest_dir, "*"))):
        file_name = os.path.basename(path)
        if (os.path.isfile(path) and not _is_office_lock_file(path)
                and os.path.splitext(file_name)[1].lower() in kind.extensions
                and kind.matches(file_name) and name_date(file_name) is not None):
            yield path


def _local_file(kind, day):
    """Наявний у теці файл цього виду за дату day (перший за абеткою) або None."""
    return next((path for path in _local_files(kind) if name_date(os.path.basename(path)) == day), None)


def _missing(kind, by_date, today):
    """[(дата, очікувана назва файлу)] - дати від kind.first_date до today
    включно, за які файлу немає ні в Signal, ні в теці. Очікувана назва
    будується за зразком найпізнішого відомого файлу цього виду (у Signal чи в
    теці) з підставленою датою (kind.canonical_name)."""
    known = [(day, versions[-1]["fileName"]) for day, versions in by_date.items()]
    known += [(name_date(os.path.basename(path)), os.path.basename(path)) for path in _local_files(kind)]
    template = max(known)[1] if known else None
    missing, day = [], kind.first_date
    while day <= today:
        if day not in by_date and _local_file(kind, day) is None:
            missing.append((day, kind.canonical_name(day, _with_date(template, day) if template else None)))
        day += timedelta(days=1)
    return missing


def _saved_after_sending(path, row):
    """Файл у теці збережено після надсилання цієї (найновішої) версії в
    Signal: він її вже містить або виправлений вручну пізніше."""
    return os.path.getmtime(path) * 1000 >= (row.get("sentAt") or 0)


def _is_older_version(path, older_rows, decrypt_to, tmp_dir):
    """Файл у теці - точна копія однієї з попередніх версій у Signal (збережено
    не ту), тож він не актуальний, хоч і збережений пізніше за найновішу."""
    for index, row in enumerate(older_rows):
        copy = os.path.join(tmp_dir, f"older_{index}{os.path.splitext(row['fileName'])[1]}")
        if decrypt_to(row, copy) and filecmp.cmp(copy, path, shallow=False):
            return True
    return False


def _move_to_previous_versions(path, now):
    folder = os.path.join(os.path.dirname(path), SIGNAL_REPLACED_VERSIONS_DIR_NAME)
    os.makedirs(folder, exist_ok=True)
    stem, extension = os.path.splitext(os.path.basename(path))
    shutil.move(path, os.path.join(folder, f"{stem} (замінено {now:%d.%m.%Y %H-%M-%S}){extension}"))


def sync_kind(kind, rows, decrypt_to, today, now=None):
    """Звіряє теку kind.dest_dir з вкладеннями (versions_by_date).
    decrypt_to(row, шлях) -> bool розшифровує вкладення в шлях."""
    now = now or datetime.now()
    result = SyncResult(kind.label)
    os.makedirs(kind.dest_dir, exist_ok=True)
    by_date = versions_by_date(rows, kind)
    with tempfile.TemporaryDirectory() as tmp_dir:
        for day, versions in sorted(by_date.items()):
            row = versions[-1]
            local = _local_file(kind, day)
            if (local is not None and _saved_after_sending(local, row)
                    and not _is_older_version(local, versions[:-1], decrypt_to, tmp_dir)):
                result.actual.append(day)
                continue
            downloaded = os.path.join(tmp_dir, f"{day:%Y%m%d}{os.path.splitext(row['fileName'])[1]}")
            if not decrypt_to(row, downloaded):
                result.failed.append(row["fileName"])
                continue
            if local is not None and filecmp.cmp(downloaded, local, shallow=False):
                result.actual.append(day)
                continue
            if local is not None:
                _move_to_previous_versions(local, now)
                result.updated.append(day)
                target = local
            else:
                result.new.append(day)
                target = os.path.join(kind.dest_dir, kind.canonical_name(day, row["fileName"]))
            shutil.copyfile(downloaded, target)
    result.missing = _missing(kind, by_date, today)
    return result


def _group_attachments(cursor, conversation_id):
    """Усі вкладення розмови, які Signal уже завантажив і не позначив помилкою."""
    cursor.execute(
        """
        SELECT fileName, sentAt, path, localKey, size, version, editHistoryIndex
        FROM message_attachments
        WHERE conversationId = ? AND error IS NULL AND path IS NOT NULL
        """,
        (conversation_id,),
    )
    columns = ("fileName", "sentAt", "path", "localKey", "size", "version", "editHistoryIndex")
    return [dict(zip(columns, values)) for values in cursor.fetchall()]


def sync_reports_and_schedules(group_name=REPORTS_SIGNAL_GROUP_NAME, today=None,
                               report_dir=REPORT_DIR, schedule_dir=SCHEDULE_DIR):
    """[SyncResult рапортів, SyncResult відомостей]. Якщо групу не
    налаштовано чи не знайдено або Signal Desktop недоступний -
    SignalFetchError (викликач друкує її червоним і йде далі)."""
    if not group_name:
        raise SignalFetchError(
            "Звіряння рапортів з Signal не налаштовано: у resources/data.json немає groups.REPORTS_SIGNAL_GROUP_NAME."
        )
    today = today or date.today()
    source_dir = signal_bchs._signal_source_dir()
    db, cursor = signal_bchs._open_db(source_dir)
    try:
        conversation_id = signal_bchs._find_conversation_id(cursor, group_name)
        if conversation_id is None:
            raise SignalFetchError(f"У Signal не знайдено чат/групу \"{group_name}\".")
        rows = _group_attachments(cursor, conversation_id)
    finally:
        db.close()

    def decrypt_to(row, path):
        return signal_bchs._decrypt_to(source_dir, row, path)

    return [sync_kind(kind, rows, decrypt_to, today) for kind in _kinds(today, report_dir, schedule_dir)]
