"""Виправлена копія таблиці сигналів: сигнали з тексту БН, яких у таблиці
немає, дописуються окремим розділом у кінці таблиці, а сигнали виходу резервів
на вогневі рубежі ("ВР №N" у завданнях БнГ / ЗВРез) замінюються в рядках
"Вийти на вогневий рубіж №N" відповідної групи, якщо в таблиці вони інші.

Оригінал у resources/ не змінюється - копія пишеться в output/ під тією самою
назвою; після перевірки штаб кладе її в resources/ замість старої."""
import copy
import os
import re
from dataclasses import dataclass, field

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from content.signal_table import cell_texts, parse_rows, split_signal
from utils.docx_utils import save_docx_safely
from utils.mgrs import find_mgrs

_VR_RE = re.compile(r"^\s*ВР\s*№\s*(?P<number>\d+)")
_ROW_NUMBER_RE = re.compile(r"№\s*(?P<number>\d+)")
_RESERVE_GROUPS = ("БНГ", "ЗВРЕЗ")
_MEANING_LIMIT = 180
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
# Дія, яку запускає сигнал: "здійснити відхід на …", "відійти на …" тощо - до
# першої коми, крапки з комою чи уточнення координат.
_ACTION_RE = re.compile(
    r"(?P<action>(?:здійсн\w*|розпоча\w*)\s+(?:поступов\w+\s+)?(?:відхід|переміщення|висування|заміну|перерозподіл)\b.+?"
    r"|відійти\s+на\s+.+?|переміститис\w*\s+на\s+.+?|висунут\w*\s+.+?|зайняти\s+.+?)"
    r"(?=\s*(?:[,;]|\s+за\s+координатами|\s+та\s+(?:зайня|займ)|\s+з\s+переднім|\s+під\s+прикриттям|$))",
    re.IGNORECASE,
)
# "до «С» + 00.45" на початку змісту - час, а не дія.
_TIME_PREFIX_RE = re.compile(r"^(?:до\s*)?[«‘'\"“]{0,2}\s*С\s*[»’'\"”]{0,2}\s*\+\s*\d{1,2}[.:]\d{2}[,\s]*", re.IGNORECASE)


@dataclass
class SignalsResult:
    path: str = ""
    added: list = field(default_factory=list)      # [(назва, зміст)]
    updated: list = field(default_factory=list)    # [(група, зміст рядка, було, стало)]
    duplicates: list = field(default_factory=list)  # [(назва, [змісти])] - один сигнал на різні дії


def _set_cell_text(tc, text):
    """Текст комірки -> text у першому run першого абзацу (форматування
    комірки й run лишається, решта абзаців і runs видаляються)."""
    paragraphs = list(tc.iter(qn("w:p")))
    for paragraph in paragraphs[1:]:
        paragraph.getparent().remove(paragraph)
    runs = list(paragraphs[0].iter(qn("w:r")))
    if not runs:
        runs = [OxmlElement("w:r")]
        paragraphs[0].append(runs[0])
    for run in runs[1:]:
        run.getparent().remove(run)
    first = runs[0]
    for child in list(first):
        if child.tag in (qn("w:t"), qn("w:tab"), qn("w:br")):
            first.remove(child)
    t = OxmlElement("w:t")
    t.set(_XML_SPACE, "preserve")
    t.text = text
    first.append(t)


def _clause(text, position):
    """Речення (до "." чи ";"), у якому стоїть позиція position."""
    start = 0
    for m in re.finditer(r"[.;](?=\s+[А-ЯІЇЄҐA-Z«])", text):
        if m.end() > position:
            return text[start:m.end()]
        start = m.end()
    return text[start:]


def _unit_label(unit):
    """"Командиру БнГ командира …" -> "командиру БнГ командира …" (для дужок)."""
    unit = (unit or "").strip().rstrip(":")
    return unit[:1].lower() + unit[1:] if unit else ""


def _tidy(text):
    for match in reversed(find_mgrs(text)):
        text = text[:match.start] + text[match.end:]
    text = re.sub(r"\(\s*[,;\s]*\)", "", text)
    text = re.sub(r"\s{2,}", " ", text)
    return re.sub(r"\s+([,.;:])", r"\1", text).strip(" ,;:.–-")


def signal_meaning(text, name, unit=""):
    """Зміст сигналу для таблиці з речення БН: для "ВР №N" - "Вийти на вогневий
    рубіж №N", інакше - дія після сигналу ("здійснити відхід на …"), а якщо її
    не знайдено - речення без координат і самого сигналу (до ~180 символів).
    unit - заголовок завдання ("Командиру …") - дописується в дужках."""
    vr = _VR_RE.match(text)
    if vr:
        meaning = f"Вийти на вогневий рубіж №{vr.group('number')}"
    else:
        quoted = re.compile(
            r"(?:(?:з\s+отриманням\s+сигналу|за\s+сигналом)\s*)?[«‘'\"“]{1,2}\s*" + re.escape(name) + r"\s*[»’'\"”]{1,2}",
            re.IGNORECASE,
        )
        m = quoted.search(text)
        clause = _clause(text, m.start() if m else 0)
        clause_end = text.find(clause) + len(clause)
        after = _tidy(text[m.end():clause_end] if m else clause)
        actions = [a.group("action") for a in _ACTION_RE.finditer(after)]
        # "відхід на «С» + 01.20" - це час, а не рубіж: беремо першу дію з місцем.
        action = next((a for a in actions if not re.search(r"\bна\s*[«‘'\"“]{0,2}\s*С\s*[»’'\"”]{0,2}\s*\+", a)),
                      actions[0] if actions else "")
        meaning = _tidy(action) if action else _TIME_PREFIX_RE.sub("", _tidy(quoted.sub("", clause)))
        if len(meaning) > _MEANING_LIMIT:
            meaning = meaning[:_MEANING_LIMIT].rsplit(" ", 1)[0] + "…"
        meaning = meaning[:1].upper() + meaning[1:]
    label = _unit_label(unit)
    return f"{meaning} ({label})" if label else meaning


def _reserve_group(unit):
    upper = (unit or "").upper()
    return next((group for group in _RESERVE_GROUPS if group in upper), "")


def _vr_corrections(rows, mentions):
    """[(рядок таблиці, назва сигналу з БН)] для ВР №N резервів, де сигнали різні."""
    corrections = []
    for mention in mentions:
        vr = _VR_RE.match(mention.text)
        group = _reserve_group(mention.unit)
        if not vr or not group:
            continue
        wanted = re.compile(rf"вогнев\w*\s+рубіж\w*\s*№\s*{vr.group('number')}(?!\d)", re.IGNORECASE)
        row = next((r for r in rows if group in r.group.upper().replace(" ", "") and wanted.search(r.meaning)), None)
        if row is not None and row.key != mention.key and all(r is not row for r, _ in corrections):
            corrections.append((row, mention.name))
    return corrections


def _next_section_number(trs):
    numbers = []
    for tr in trs:
        filled = [text for text in cell_texts(tr) if text]
        m = re.match(r"^(\d+)\.\s*\D", filled[0]) if len(filled) == 1 else None
        if m:
            numbers.append(int(m.group(1)))
    return max(numbers, default=0) + 1


def _missing_mentions(rows, mentions):
    known = {row.key for row in rows}
    missing = []
    for mention in mentions:
        if mention.key not in known and all(mention.key != m.key for m in missing):
            missing.append(mention)
    return missing


def find_duplicates(mentions):
    """Сигнали БН, які стоять за різними діями (напр. той самий сигнал для
    ВР №3 і ВР №21) - імовірна помилка. -> [(назва, [змісти])]."""
    meanings = {}
    names = {}
    for mention in mentions:
        meaning = signal_meaning(mention.text, mention.name, mention.unit)
        meanings.setdefault(mention.key, [])
        if meaning not in meanings[mention.key]:
            meanings[mention.key].append(meaning)
        names.setdefault(mention.key, mention.name)
    return [(names[key], values) for key, values in meanings.items() if len(values) > 1]


def _reserve_anchor(rows, trs, mention):
    """Рядок таблиці, після якого вставити сигнал "ВР №N" резерву: останній
    рядок тієї ж групи з номером того самого десятка (рубіж), інакше - останній
    рядок групи. None - групи в таблиці немає."""
    vr = _VR_RE.match(mention.text)
    group = _reserve_group(mention.unit)
    if not vr or not group:
        return None
    in_group = [r for r in rows if group in r.group.upper().replace(" ", "")]
    if not in_group:
        return None
    tens = int(vr.group("number")) // 10
    same_line = [r for r in in_group if (m := _ROW_NUMBER_RE.search(r.meaning)) and int(m.group("number")) // 10 == tens]
    return trs[(same_line or in_group)[-1].index]


def write_corrected_table(table, mentions, output_dir, title):
    """table - SignalTable (resources/), mentions - SignalMention з БН, title -
    "бойового наказу №N від ДД.ММ.РРРР". -> SignalsResult; path порожній,
    якщо виправляти нічого."""
    result = SignalsResult()
    if table is None or not table.rows:
        return result
    doc = Document(table.path)
    docx_table = doc.tables[table.table_index]
    rows = parse_rows(docx_table)
    trs = list(docx_table._tbl.tr_lst)

    for row, name in _vr_corrections(rows, mentions):
        word, number = split_signal(name)
        tcs = list(trs[row.index].iter(qn("w:tc")))
        old = row.name
        _set_cell_text(tcs[1], word)
        _set_cell_text(tcs[2], number)
        row.word, row.number = word, number
        result.updated.append((row.group, row.meaning, old, row.name))

    # Зразок рядка - останній сигнал із заповненим змістом (у порожній комірці
    # немає run, з якого можна взяти шрифт).
    template = trs[next((r.index for r in reversed(rows) if r.meaning), rows[-1].index)]

    def new_row(mention):
        word, digits = split_signal(mention.name)
        meaning = signal_meaning(mention.text, mention.name, mention.unit)
        tr = copy.deepcopy(template)
        tcs = list(tr.iter(qn("w:tc")))
        _set_cell_text(tcs[1], word)
        _set_cell_text(tcs[2], digits)
        _set_cell_text(tcs[3], meaning)
        result.added.append((f"{word}-{digits}", meaning))
        return tr

    rest, inserted = [], []
    for mention in _missing_mentions(rows, mentions):
        anchor = _reserve_anchor(rows, trs, mention)
        if anchor is None:
            rest.append(mention)
            continue
        # Кілька вставок після того самого рядка - кожна наступна стає нижче попередньої.
        following = anchor.getnext()
        while following is not None and any(following is row for row in inserted):
            anchor, following = following, following.getnext()
        row_tr = new_row(mention)
        anchor.addnext(row_tr)
        inserted.append(row_tr)

    if rest:
        anchor = trs[-1]
        section_tr = next((tr for tr in reversed(trs) if len([t for t in cell_texts(tr) if t]) == 1), None)
        if section_tr is not None:
            header = copy.deepcopy(section_tr)
            label_tc = next(tc for tc in header.iter(qn("w:tc")) if "".join(t.text or "" for t in tc.iter(qn("w:t"))).strip())
            _set_cell_text(label_tc, f"{_next_section_number(trs)}. СИГНАЛИ {title.upper()} (ДОДАНО ПРОГРАМОЮ - ПЕРЕВІРТЕ)")
            anchor.addnext(header)
            anchor = header
        for mention in rest:
            row_tr = new_row(mention)
            anchor.addnext(row_tr)
            anchor = row_tr

    result.duplicates = find_duplicates(mentions)
    if result.added or result.updated:
        os.makedirs(output_dir, exist_ok=True)
        result.path = os.path.join(output_dir, os.path.basename(table.path))
        save_docx_safely(doc, result.path)
    return result
