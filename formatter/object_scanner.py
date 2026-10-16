from __future__ import annotations

import re
from dataclasses import dataclass

from .utils import normalize_space, paragraph_has_page_break


SECTION_HEADING_RE = re.compile(
    r"^(?:[IVXLCDM]+\.|[0-9]+\.|[A-Z]\.)\s+[A-Z][A-Z0-9 ,:/()\-]+$"
)


@dataclass
class ScanItem:
    kind: str
    number: int | None
    paragraph_index: int | None
    text: str
    issue: str | None = None


def detect_main_report_start(document) -> int:
    """Find the first paragraph that looks like the real report body."""

    page_break_candidate: int | None = None
    abstract_indices: list[int] = []
    saw_contents = False
    for index, paragraph in enumerate(document.paragraphs):
        text = normalize_space(paragraph.text)
        lower = text.lower()
        if lower in {"table of contents", "contents", "目录"}:
            saw_contents = True
        if lower in {"abstract", "abstract.", "摘要"} or lower.startswith("abstract"):
            abstract_indices.append(index)
            if not saw_contents:
                return index
            if len(abstract_indices) >= 2:
                return index
        if re.match(r"^(?:i\.|1\.)\s+introduction\b", lower):
            return index
        if SECTION_HEADING_RE.match(text):
            return index
        if page_break_candidate is None and paragraph_has_page_break(paragraph):
            page_break_candidate = min(index + 1, len(document.paragraphs) - 1)
    if abstract_indices:
        return abstract_indices[-1]
    return page_break_candidate or 0


def looks_like_heading(text: str) -> bool:
    normalized = normalize_space(text)
    return bool(SECTION_HEADING_RE.match(normalized) or normalized.lower() == "abstract")
