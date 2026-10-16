from __future__ import annotations

import re

from .caption_detector import CaptionMatch
from .equation_utils import (
    extract_trailing_equation_numbers,
    get_visible_text_from_xml,
    has_omml_math,
    is_unnumbered_equation_block,
    match_equation_label_paragraph,
    match_equation_label_text,
)
from .utils import iter_block_items, normalize_space, should_skip_for_body_inventory


def match_equation_number(text: str):
    return match_equation_label_text(text)


def is_equation_like(text: str) -> bool:
    return match_equation_number(text) is not None


def detect_equation_numbers(document, start_index: int = 0) -> list[CaptionMatch]:
    matches: list[CaptionMatch] = []
    for index, paragraph, block_index in _iter_body_paragraphs(document, start_index):
        raw = get_visible_text_from_xml(paragraph) or paragraph.text
        match = match_equation_label_paragraph(paragraph)
        if not match:
            continue
        number_token = match.group("number")
        number = int(re.search(r"\d+", number_token).group(0))
        matches.append(
            CaptionMatch(
                kind="equation",
                paragraph_index=index,
                number=number,
                text=normalize_space(raw),
                raw_text=normalize_space(raw),
                pattern=match.re.pattern,
                number_token=number_token,
                source="equation_block",
                block_index=block_index,
                paragraph=paragraph,
            )
        )
    table_matches, _consumed = _detect_table_row_equation_numbers(document, start_index)
    return sorted([*matches, *table_matches], key=_equation_match_sort_key)


def detect_unnumbered_equation_blocks(document, start_index: int = 0) -> list[CaptionMatch]:
    matches: list[CaptionMatch] = []
    for index, paragraph, block_index in _iter_body_paragraphs(document, start_index):
        if not is_unnumbered_equation_block(paragraph):
            continue
        raw = get_visible_text_from_xml(paragraph) or paragraph.text
        matches.append(
            CaptionMatch(
                kind="equation",
                paragraph_index=index,
                number=0,
                text=normalize_space(raw),
                raw_text=normalize_space(raw),
                pattern="heuristic_math_line",
                number_token=None,
                source="unnumbered_equation_block",
                block_index=block_index,
                paragraph=paragraph,
            )
        )
    matches.extend(_detect_unnumbered_table_row_equations(document, start_index))
    return sorted(matches, key=_equation_match_sort_key)


def detect_equation_blocks(document, start_index: int = 0) -> list[CaptionMatch]:
    return sorted(
        [*detect_equation_numbers(document, start_index), *detect_unnumbered_equation_blocks(document, start_index)],
        key=_equation_match_sort_key,
    )


def _iter_body_paragraphs(document, start_index: int):
    for block_index, block in enumerate(iter_block_items(document)):
        if block.kind != "paragraph":
            continue
        paragraph = block.element
        index = block.paragraph_index
        if should_skip_for_body_inventory(paragraph, index, start_index):
            continue
        yield index, paragraph, block_index * 1000


def _detect_table_row_equation_numbers(document, start_index: int) -> tuple[list[CaptionMatch], set[int]]:
    matches: list[CaptionMatch] = []
    consumed_paragraphs: set[int] = set()
    for block_index, block in enumerate(iter_block_items(document)):
        if block.kind != "table" or (block.paragraph_index is not None and block.paragraph_index < start_index):
            continue
        table = block.element
        for row_offset, row in enumerate(table.rows):
            row_paragraphs = _unique_row_paragraphs(row)
            formula_paragraphs = [
                paragraph for paragraph in row_paragraphs if _paragraph_has_structural_equation_body(paragraph)
            ]
            if not formula_paragraphs:
                continue
            label_candidates = []
            for paragraph in row_paragraphs:
                text = get_visible_text_from_xml(paragraph) or paragraph.text
                numbers = extract_trailing_equation_numbers(text)
                if numbers:
                    label_candidates.append((paragraph, text, numbers))
            if not label_candidates:
                continue
            label_paragraph, label_text, numbers = label_candidates[-1]
            row_text = _row_text(row_paragraphs)
            matches.append(
                CaptionMatch(
                    kind="equation",
                    paragraph_index=None,
                    number=numbers[-1],
                    text=normalize_space(row_text),
                    raw_text=normalize_space(row_text),
                    pattern="table_row_equation_number",
                    number_token=str(numbers[-1]),
                    source="equation_table_row",
                    block_index=block_index * 1000 + row_offset,
                    paragraph=label_paragraph,
                )
            )
            consumed_paragraphs.update(id(paragraph._p) for paragraph in row_paragraphs)
    return matches, consumed_paragraphs


def _detect_unnumbered_table_row_equations(document, start_index: int) -> list[CaptionMatch]:
    matches: list[CaptionMatch] = []
    _numbered_matches, consumed = _detect_table_row_equation_numbers(document, start_index)
    for block_index, block in enumerate(iter_block_items(document)):
        if block.kind != "table" or (block.paragraph_index is not None and block.paragraph_index < start_index):
            continue
        table = block.element
        for row_offset, row in enumerate(table.rows):
            row_paragraphs = _unique_row_paragraphs(row)
            if any(id(paragraph._p) in consumed for paragraph in row_paragraphs):
                continue
            formula_paragraphs = [
                paragraph for paragraph in row_paragraphs if _paragraph_has_structural_equation_body(paragraph)
            ]
            if not formula_paragraphs:
                continue
            paragraph = formula_paragraphs[-1]
            raw = get_visible_text_from_xml(paragraph) or paragraph.text
            matches.append(
                CaptionMatch(
                    kind="equation",
                    paragraph_index=None,
                    number=0,
                    text=normalize_space(_row_text(row_paragraphs) or raw),
                    raw_text=normalize_space(_row_text(row_paragraphs) or raw),
                    pattern="table_row_heuristic_math_line",
                    number_token=None,
                    source="unnumbered_equation_table_row",
                    block_index=block_index * 1000 + row_offset,
                    paragraph=paragraph,
                )
            )
    return matches


def _paragraph_has_equation_body(paragraph) -> bool:
    text = get_visible_text_from_xml(paragraph) or paragraph.text
    return (
        has_omml_math(paragraph)
        or match_equation_label_paragraph(paragraph) is not None
        or bool(text and is_unnumbered_equation_block(paragraph))
    )


def _paragraph_has_structural_equation_body(paragraph) -> bool:
    return has_omml_math(paragraph) or match_equation_label_paragraph(paragraph) is not None


def _unique_row_paragraphs(row) -> list:
    paragraphs = []
    seen: set[int] = set()
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            identity = id(paragraph._p)
            if identity in seen:
                continue
            seen.add(identity)
            if should_skip_for_body_inventory(paragraph):
                continue
            paragraphs.append(paragraph)
    return paragraphs


def _row_text(paragraphs: list) -> str:
    return "\t".join(filter(None, [(get_visible_text_from_xml(paragraph) or paragraph.text) for paragraph in paragraphs]))


def _equation_match_sort_key(item: CaptionMatch):
    if item.block_index is not None:
        return (item.block_index, item.paragraph_index if item.paragraph_index is not None else 10**9)
    return (10**9, item.paragraph_index if item.paragraph_index is not None else 10**9)
