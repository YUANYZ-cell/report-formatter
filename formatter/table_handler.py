from __future__ import annotations

from dataclasses import dataclass

from .caption_detector import match_table_number
from .format_profiles import format_table_caption
from .style_profile import StyleProfile
from .utils import (
    apply_paragraph_style_from_profile,
    insert_paragraph_before_table,
    iter_block_items,
    normalize_space,
    should_skip_for_body_inventory,
)


@dataclass
class MissingTableCaption:
    paragraph_index_before_table: int
    message: str
    table: object


def find_uncaptioned_tables(document, start_index: int = 0) -> list[MissingTableCaption]:
    missing: list[MissingTableCaption] = []
    for block in iter_block_items(document):
        if block.kind != "table":
            continue
        paragraph_index = block.paragraph_index or 0
        if paragraph_index < start_index:
            continue
        if paragraph_index < len(document.paragraphs) and should_skip_for_body_inventory(
            document.paragraphs[paragraph_index], paragraph_index, start_index
        ):
            continue
        previous_lines = _previous_nonempty_paragraphs(document.paragraphs, paragraph_index - 1, limit=3)
        if any(match_table_number(text) for text in previous_lines):
            continue
        missing.append(
            MissingTableCaption(
                paragraph_index_before_table=paragraph_index,
                message=f"Possible table before paragraph {paragraph_index + 1} has no TABLE N title line.",
                table=block.element,
            )
        )
    return missing


def insert_missing_table_placeholders(
    document,
    style_profile: StyleProfile,
    format_profile: dict,
    report_language: str,
    ui_language: str = "en",
    start_index: int = 0,
) -> list[str]:
    inserted: list[str] = []
    for offset, item in enumerate(find_uncaptioned_tables(document, start_index), start=1):
        rendered = format_table_caption(offset, "", format_profile, report_language, placeholder=True)
        if isinstance(rendered, list):
            number_text, caption_text = rendered
            number_line = insert_paragraph_before_table(item.table, number_text)
            caption_line = insert_paragraph_before_table(item.table, caption_text)
            apply_paragraph_style_from_profile(number_line, style_profile.table_number)
            apply_paragraph_style_from_profile(caption_line, style_profile.table_caption)
        else:
            number_line = insert_paragraph_before_table(item.table, rendered)
            apply_paragraph_style_from_profile(number_line, style_profile.table_number)
        if ui_language == "zh":
            inserted.append(f"已在第 {item.paragraph_index_before_table + 1} 段之前插入表题占位符。")
        else:
            inserted.append(f"Inserted placeholder table caption before paragraph {item.paragraph_index_before_table + 1}.")
    return inserted


def _previous_nonempty_paragraphs(paragraphs, start: int, limit: int = 3) -> list[str]:
    lines: list[str] = []
    for index in range(start, -1, -1):
        text = normalize_space(paragraphs[index].text)
        if text:
            lines.append(text)
            if len(lines) >= limit:
                break
    return lines
