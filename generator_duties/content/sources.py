"""Джерела в resources/, з яких програма заповнює таблицю обов'язків і
будує книги: Статут (.htm), отриманий зразок (.docx), Положення (.docx)."""
import glob
import os
from dataclasses import dataclass, field

from docx import Document

from constants import REGULATION_FILE_GLOB, SAMPLE_FILE_GLOB, STATUTE_FILE_GLOB
from content.docx_text import read_paragraph_specs
from content.sample_reader import read_sample
from content.statute_reader import read_statute

STATUTE_URL = "https://zakon.rada.gov.ua/laws/show/548-14"


@dataclass
class Sources:
    statute: list = field(default_factory=list)
    sample: object = None
    regulation: list = field(default_factory=list)
    used_files: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


def find_resource(resources_dir, pattern):
    """Найновіший файл за шаблоном (тимчасові файли Office "~$…" пропускаються)."""
    paths = [p for p in glob.glob(os.path.join(resources_dir, pattern)) if not os.path.basename(p).startswith("~$")]
    return max(paths, key=os.path.getmtime) if paths else None


def load_sources(resources_dir):
    sources = Sources()
    statute_path = find_resource(resources_dir, STATUTE_FILE_GLOB)
    if statute_path is None:
        sources.warnings.append(
            f"Немає Статуту (resources/{STATUTE_FILE_GLOB}) - збережіть сторінку {STATUTE_URL} "
            f"як «Статут внутрішньої служби.htm», інакше обов'язки зі Статуту не підставляться."
        )
    else:
        try:
            sources.statute = read_statute(statute_path)
            sources.used_files.append(statute_path)
        except (OSError, ValueError) as e:
            sources.warnings.append(f"Не вдалось прочитати Статут {os.path.basename(statute_path)}: {e}")
    sample_path = find_resource(resources_dir, SAMPLE_FILE_GLOB)
    if sample_path is None:
        sources.warnings.append(f"Немає зразка (resources/{SAMPLE_FILE_GLOB}) - без витягу зі Статуту і текстів зі зразка.")
    else:
        try:
            sources.sample = read_sample(sample_path)
            sources.used_files.append(sample_path)
        except Exception as e:
            sources.warnings.append(f"Не вдалось прочитати зразок {os.path.basename(sample_path)}: {e}")
    regulation_path = find_resource(resources_dir, REGULATION_FILE_GLOB)
    if regulation_path is None:
        sources.warnings.append(
            f"Немає Положення про батальйон (resources/{REGULATION_FILE_GLOB}) - розділ «Положення» пропущено."
        )
    else:
        try:
            sources.regulation = read_paragraph_specs(Document(regulation_path))
            sources.used_files.append(regulation_path)
        except Exception as e:
            sources.warnings.append(f"Не вдалось прочитати {os.path.basename(regulation_path)}: {e}")
    return sources
