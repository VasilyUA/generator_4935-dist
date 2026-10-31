"""Унікальні посади зі штату: назва без номерів підрозділів і без ланцюга
вищого підрозділу ("…батальйону…", "військової частини …" - data.json →
roster.position_suffixes або, якщо там порожньо, auto_suffixes), у порядку
першої появи в штаті; кожна посада знає всі свої слоти (людей і вакансії)."""
import re
from dataclasses import dataclass, field

from constants import CHAIN_WORDS, COMPOSITE_HEADS, DANGLING_WORDS, SUBUNIT_WORDS
from content.docx_text import clean_text

# Номер підрозділу ("2 відділення", "№3"); номер категорії/розряду/класу/рангу
# - частина назви посади ("Штаб-сержант 1 категорії"), лишається.
_NUMBER_TOKEN_RE = re.compile(r"(?<!\S)(?:№\s*)?\d+(?:-?(?:й|го|ої|а|ї|я))?(?!\S)"
                              r"(?!\s+(?:категорії|розряду|класу|рангу)\b)")
_DASH_RE = re.compile(r"\s+[-–—]\s+")
_COMPOSITE_RE = re.compile(rf"(?<=\w)\s*[-–—]\s*(?=(?:{'|'.join(COMPOSITE_HEADS)})\w*\b)", re.IGNORECASE)
_EDGE_PUNCTUATION = " ,;:.–-"


def normalize_dashes(text):
    """Тире між частинами назви - завжди " – " (з пробілами), зокрема у
    складених посадах, записаних через дефіс: "сержант-командир" →
    "сержант – командир"."""
    return _DASH_RE.sub(" – ", _COMPOSITE_RE.sub(" – ", clean_text(text)))


@dataclass
class UniquePosition:
    title: str
    key: str
    slots: list = field(default_factory=list)

    @property
    def vos_codes(self):
        codes = []
        for slot in self.slots:
            if slot.vos and slot.vos not in codes:
                codes.append(slot.vos)
        return codes

    @property
    def occupied(self):
        return sum(1 for slot in self.slots if not slot.vacant)


def compare_key(text):
    """Ключ для порівняння назв: без регістру, з однаковими апострофами,
    тире й пробілами, без розділових знаків на краях."""
    return normalize_dashes(text).strip(_EDGE_PUNCTUATION).lower()


def _needs_subunit(head):
    words = head.lower().replace(",", " ").split()
    return not words or not any(word in SUBUNIT_WORDS for word in words) or words[-1] in DANGLING_WORDS


def _first_word(text):
    for word in clean_text(text).split():
        if not _NUMBER_TOKEN_RE.fullmatch(word):
            return word
    return ""


def _capitalize(text):
    text = clean_text(text).strip(_EDGE_PUNCTUATION)
    return text[:1].upper() + text[1:]


def strip_chain(raw, suffixes=()):
    """Прибирає ланцюг вищого підрозділу там, де він стоїть (останній збіг;
    те, що після нього - "… з озброєння", - лишається). Якщо перед ланцюгом
    лишилось саме "Командир" (чи "…заступник командира") - залишається перше
    слово ланцюга: "Командир батальйону". Номери підрозділів не чіпає."""
    text = normalize_dashes(raw)
    for suffix in sorted((clean_text(s) for s in suffixes if s), key=len, reverse=True):
        index = text.lower().rfind(suffix.lower())
        if index > 0:
            head = text[:index].strip(_EDGE_PUNCTUATION)
            tail = text[index + len(suffix):].strip(_EDGE_PUNCTUATION)
            if _needs_subunit(head):
                head = f"{head} {_first_word(suffix)}".strip()
            text = f"{head} {tail}".strip()
    return _capitalize(text)


def normalize_title(raw, suffixes=()):
    """"Водій 2 господарчого відділення взводу … 1 батальйону X" → "Водій
    господарчого відділення взводу …" - назва унікальної посади."""
    return _capitalize(_NUMBER_TOKEN_RE.sub(" ", strip_chain(raw, suffixes)))


def common_suffix(texts):
    """Спільне закінчення (цілими словами) усіх непорожніх текстів."""
    word_lists = [clean_text(text).split() for text in texts if clean_text(text)]
    if not word_lists:
        return ""
    common = []
    for words in zip(*(reversed(words) for words in word_lists)):
        if any(word != words[0] for word in words):
            break
        common.append(words[0])
    return " ".join(reversed(common))


def _is_chain(text):
    words = text.lower().split()
    return len(words) >= 2 and any(word in CHAIN_WORDS for word in words)


def auto_suffixes(slots):
    """Ланцюг вищого підрозділу без data.json: спільне закінчення всіх назв
    посад ("військової частини …") і всіх значень колонки підрозділу ("…
    батальйону …") - якщо в ньому є слово-підрозділ, а від кожної назви посади
    щось лишається."""
    positions = [slot.position for slot in slots if slot.position]
    chains = []
    position_chain = common_suffix(positions)
    if _is_chain(position_chain) and all(len(clean_text(p).split()) > len(position_chain.split()) for p in positions):
        chains.append(position_chain)
    subunit_chain = common_suffix([slot.subunit for slot in slots])
    if _is_chain(subunit_chain) and subunit_chain not in chains:
        chains.append(subunit_chain)
    return chains


def group_positions(slots, suffixes=()):
    """Слоти штату → [UniquePosition] у порядку першої появи. Кожен слот
    отримує display - посаду без ланцюга, але з номерами (лист ознайомлення)."""
    by_key = {}
    for slot in slots:
        slot.display = strip_chain(slot.position, suffixes)
        title = normalize_title(slot.position, suffixes)
        if not title:
            continue
        key = compare_key(title)
        if key not in by_key:
            by_key[key] = UniquePosition(title, key)
        by_key[key].slots.append(slot)
    return list(by_key.values())


def find_commander(positions, key="командир батальйону"):
    """Слот комбата (для підписів): перший зайнятий слот посади з цим ключем."""
    for position in positions:
        if position.key == key:
            occupied = [slot for slot in position.slots if not slot.vacant]
            return (occupied or position.slots)[0]
    return None
