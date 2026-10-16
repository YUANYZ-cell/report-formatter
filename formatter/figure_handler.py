from __future__ import annotations

from dataclasses import dataclass

from .caption_detector import match_figure_caption
from .format_profiles import format_figure_caption
from .style_profile import StyleProfile
from .utils import (
    apply_paragraph_style_from_profile,
    insert_paragraph_after,
    normalize_space,
    paragraph_has_drawing,
    should_skip_for_body_inventory,
)


@dataclass
class MissingFigureCaption:
    paragraph_index: int
    message: str


def find_uncaptioned_figures(document, start_index: int = 0) -> list[MissingFigureCaption]:
    missing: list[MissingFigureCaption] = []
    paragraphs = document.paragraphs
    for index, paragraph in enumerate(paragraphs[start_index:], start=start_index):
        if should_skip_for_body_inventory(paragraph, index, start_index):
            continue
        if not paragraph_has_drawing(paragraph):
            continue
        next_text = _next_nonempty_text(paragraphs, index + 1)
        if next_text and match_figure_caption(next_text):
            continue
        missing.append(
            MissingFigureCaption(
                paragraph_index=index,
                message=f"Possible figure at paragraph {index + 1} has no following figure caption.",
            )
        )
    return missing


def insert_missing_figure_placeholders(
    document,
    style_profile: StyleProfile,
    format_profile: dict,
    report_language: str,
    ui_language: str = "en",
    start_index: int = 0,
) -> list[str]:
    inserted: list[str] = []
    missing = find_uncaptioned_figures(document, start_index)
    for offset, item in enumerate(missing, start=1):
        paragraph = document.paragraphs[item.paragraph_index]
        text = format_figure_caption(offset, "", format_profile, report_language, placeholder=True)
        placeholder = insert_paragraph_after(paragraph, text)
        apply_paragraph_style_from_profile(placeholder, style_profile.figure_caption)
        if ui_language == "zh":
            inserted.append(f"已在第 {item.paragraph_index + 1} 段之后插入图题占位符。")
        else:
            inserted.append(f"Inserted placeholder figure caption after paragraph {item.paragraph_index + 1}.")
    return inserted


def _next_nonempty_text(paragraphs, start: int) -> str | None:
    for index in range(start, min(start + 4, len(paragraphs))):
        text = normalize_space(paragraphs[index].text)
        if text:
            return text
    return None
