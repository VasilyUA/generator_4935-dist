"""Початкові набори обов'язків для нових посад - зі Статуту і зі зразка;
чого немає ні там, ні там - "ЗАПОВНИТИ". Набір отримує назву за джерелом
("Командир відділення (Статут)"), тож посади з однаковими обов'язками
потрапляють в один набір (DutiesTable.add_set)."""
import re

from constants import (
    ALL_POSITIONS, COMPOSITE_PRIORITY, LEADER_ASSISTANT_OF, LEADER_WORDS, RANK_ADJECTIVES, SECTION_RESPONSIBILITY,
    SECTIONS, SECTIONS_REQUIRED, SHEET_FUNCTIONAL, SHEET_JOB, STATUTE_RANK_AND_FILE, STATUTE_RESPONSIBILITY,
    STATUTE_RULES, SUBUNIT_WORDS, TO_FILL,
)
from content.duties_table import DutySet, name_from_sources
from content.positions import compare_key, normalize_dashes
from content.sample_reader import find_block
from content.statute_reader import find_section

SOURCE_SAMPLE = "sample"
SOURCE_STATUTE = "statute"
# На якому аркуші яке джерело головніше (решта розділів - з наступного).
SOURCE_PRIORITY = {
    SHEET_FUNCTIONAL: (SOURCE_SAMPLE, SOURCE_STATUTE),
    SHEET_JOB: (SOURCE_STATUTE, SOURCE_SAMPLE),
}
# Слова з дефісом ("взводу-старший", "морально-психологічного") - окремими
# словами, і в назві посади, і в правилі.
_WORD_RE = re.compile(r"[^\W\d_]+(?:’[^\W\d_]+)*")


def _words(text):
    return _WORD_RE.findall(compare_key(text))


def _rule_match(rule_words, title_words):
    """→ індекс останнього збігу, якщо слова правила йдуть у назві в тому самому
    порядку (перше - першим словом назви), інакше None. Між ними можуть бути
    лише уточнення ("командир господарчого відділення"), але не інший
    підрозділ ("начальник групи … штабу" - це не начальник штабу)."""
    if not title_words or title_words[0] != rule_words[0]:
        return None
    position = 0
    for word in rule_words[1:]:
        try:
            found = title_words.index(word, position + 1)
        except ValueError:
            return None
        if any(between in SUBUNIT_WORDS for between in title_words[position + 1:found]):
            return None
        position = found
    return position


def _best_rule(title_words):
    best = None
    for rule, heading in STATUTE_RULES:
        end = _rule_match(_words(rule), title_words)
        if end is None:
            continue
        score = (len(_words(rule)), -end)
        if best is None or score > best[0]:
            best = (score, heading)
    return best[1] if best else None


def _composite_rank(part):
    words = _words(part)
    return COMPOSITE_PRIORITY.index(words[0]) if words and words[0] in COMPOSITE_PRIORITY else len(COMPOSITE_PRIORITY)


def match_statute_rule(title):
    """Назва посади → початок заголовка підрозділу Статуту (див.
    constants.STATUTE_RULES) або загальні обов'язки рядового складу.
    Складена посада "X – Y" перевіряється спершу за головною частиною
    (constants.COMPOSITE_PRIORITY), потім цілком."""
    parts = [part for part in normalize_dashes(title).split(" – ") if part.strip()]
    candidates = sorted(parts, key=_composite_rank) + [title] if len(parts) > 1 else [title]
    for candidate in candidates:
        heading = _best_rule(_words(candidate))
        if heading:
            return heading
    if _words(title) and not _is_leader(title):
        return STATUTE_RANK_AND_FILE
    return None


def _is_leader(title):
    """Керівна посада (див. constants.LEADER_WORDS): за першим словом (зокрема
    частинами слова з дефісом - "штаб-сержант"), "старший/головний + офіцер,
    технік…" чи "помічник командира/начальника"."""
    words = _words(title)
    first_token = _words(normalize_dashes(title).split()[0])
    if any(word in LEADER_WORDS for word in first_token):
        return True
    if len(words) > 1 and words[0] in RANK_ADJECTIVES and words[1] in LEADER_WORDS:
        return True
    return len(words) > 1 and words[0] == "помічник" and words[1] in LEADER_ASSISTANT_OF


def _statute_items(position, sources):
    heading = match_statute_rule(position.title)
    section = find_section(sources.statute, heading) if heading and sources.statute else None
    if section is None:
        return {}, ""
    return section.sections, section.source


def _sample_items(position, sources):
    block = find_block(sources.sample.blocks, position.key) if sources.sample else None
    if block is None:
        return {}, ""
    return block.sections, f"Зразок: {block.title}"


_READERS = {SOURCE_SAMPLE: _sample_items, SOURCE_STATUTE: _statute_items}


def label_names(sources):
    """Джерело пункту → назва набору: "Статут ВС, ст. 125, 126" → "Командир
    відділення (Статут)". Зразок іменується за назвою блоку (name_from_sources)."""
    return {section.source: f"{section.title} (Статут)" for section in sources.statute}


def prefill_set(position, sheet, sources):
    """Набір обов'язків однієї нової посади для аркуша sheet."""
    found = [_READERS[kind](position, sources) for kind in SOURCE_PRIORITY[sheet]]
    duty_set = DutySet(position.title)
    for section in SECTIONS:
        for items, label in found:
            if items.get(section):
                for text in items[section]:
                    duty_set.add(section, text, label)
                break
        else:
            if section in SECTIONS_REQUIRED:
                duty_set.add(section, TO_FILL, TO_FILL)
    duty_set.name = name_from_sources(duty_set, position.title, label_names(sources))
    return duty_set


def common_set(sheet, sources):
    """Набір "Для всіх": на "Функціональних" - спільні пункти зразка, на
    "Посадових" - відповідальність військовослужбовців зі Статуту."""
    duty_set = DutySet(ALL_POSITIONS)
    if sheet == SHEET_FUNCTIONAL and sources.sample:
        for section in SECTIONS:
            for text in sources.sample.common.get(section, []):
                duty_set.add(section, text, "Зразок (спільне)")
    if sheet == SHEET_JOB and sources.statute:
        statute = find_section(sources.statute, STATUTE_RESPONSIBILITY)
        if statute is not None:
            for items in statute.sections.values():
                for text in items:
                    duty_set.add(SECTION_RESPONSIBILITY, text, statute.source)
    return duty_set
