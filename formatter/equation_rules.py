from __future__ import annotations

from .caption_rules import FormattingAction, _renumber_message, _should_apply_existing_caption_formatting
from .equation_detector import detect_equation_blocks
from .equation_utils import (
    EQUATION_REFERENCE_RE,
    append_equation_number,
    collapse_trailing_equation_label_cluster,
    extract_trailing_equation_numbers,
    get_visible_text_from_xml,
    iter_equation_text_paragraphs,
    is_equation_label_paragraph,
    match_equation_label_text,
    replace_final_equation_number,
    replace_trailing_equation_numbers_in_text,
    replace_text_span_in_runs,
)
from .format_profiles import format_equation_number
from .style_profile import StyleProfile
from .utils import apply_paragraph_style_from_profile, normalize_space, should_skip_for_body_inventory


def renumber_equations(
    document,
    style_profile: StyleProfile,
    format_profile: dict,
    ui_language: str = "en",
    start_index: int = 0,
    apply_formatting: bool = True,
) -> tuple[list[FormattingAction], dict[int, int]]:
    actions: list[FormattingAction] = []
    number_map: dict[int, int] = {}
    matches = _dedupe_equation_matches(detect_equation_blocks(document, start_index))
    label_indices = {match.paragraph_index for match in matches}

    for new_number, match in enumerate(matches, start=1):
        if match.paragraph is not None:
            paragraph = match.paragraph
        elif match.paragraph_index is not None:
            paragraph = document.paragraphs[match.paragraph_index]
        else:
            continue
        before = get_visible_text_from_xml(paragraph) or paragraph.text
        formatted_number = format_equation_number(new_number, format_profile)
        existing_numbers = extract_trailing_equation_numbers(before)
        changed = False
        if not existing_numbers and match.number:
            existing_numbers = [match.number]
        if existing_numbers:
            after = _replace_equation_number(before, formatted_number)
            for old_number in existing_numbers:
                number_map[old_number] = new_number
            if before != after:
                allow_standalone_label = getattr(match, "source", "") == "equation_table_row"
                if replace_final_equation_number(paragraph, formatted_number, allow_standalone_label):
                    changed = True
                    equation_id = f"equation-{new_number}"
                    actions.append(
                        FormattingAction(
                            "equation",
                            match.paragraph_index,
                            normalize_space(before),
                            normalize_space(after),
                            _equation_number_update_message(existing_numbers, new_number, ui_language),
                            object_id=equation_id,
                            equation_id=equation_id,
                            old_number=existing_numbers[-1],
                            new_number=new_number,
                        )
                    )
        elif append_equation_number(paragraph, formatted_number):
            after = get_visible_text_from_xml(paragraph) or paragraph.text
            changed = True
            equation_id = f"equation-{new_number}"
            actions.append(
                FormattingAction(
                    "equation",
                    match.paragraph_index,
                    normalize_space(before),
                    normalize_space(after),
                    _insert_equation_number_message(new_number, ui_language),
                    object_id=equation_id,
                    equation_id=equation_id,
                    new_number=new_number,
                )
            )
        if _should_apply_existing_caption_formatting(apply_formatting, format_profile, changed):
            apply_paragraph_style_from_profile(paragraph, style_profile.equation)

    actions.extend(_update_equation_references(document, number_map, label_indices, ui_language, start_index))
    actions.extend(cleanup_duplicate_equation_numbers(document, ui_language, start_index))
    return actions, number_map


def cleanup_duplicate_equation_numbers(
    document,
    ui_language: str = "en",
    start_index: int = 0,
) -> list[FormattingAction]:
    actions: list[FormattingAction] = []
    for paragraph_index, paragraph in iter_equation_text_paragraphs(document, start_index):
        changed, before, after = collapse_trailing_equation_label_cluster(paragraph)
        if not changed:
            continue
        actions.append(
            FormattingAction(
                "equation",
                paragraph_index,
                normalize_space(before),
                normalize_space(after),
                _duplicate_equation_number_cleanup_message(ui_language),
            )
        )
    return actions


def _dedupe_equation_matches(matches: list) -> list:
    deduped = []
    seen: set[tuple[str, int]] = set()
    for match in matches:
        paragraph = getattr(match, "paragraph", None)
        if paragraph is not None:
            key = ("paragraph", id(paragraph._p))
        elif match.paragraph_index is not None:
            key = ("paragraph_index", match.paragraph_index)
        else:
            key = ("block", match.block_index if match.block_index is not None else id(match))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(match)
    return deduped


def _replace_equation_number(text: str, formatted_number: str) -> str:
    match = match_equation_label_text(text or "")
    if not match:
        return replace_trailing_equation_numbers_in_text(text, formatted_number)
    return replace_trailing_equation_numbers_in_text(text, formatted_number)


def _equation_number_update_message(existing_numbers: list[int], new_number: int, ui_language: str) -> str:
    unique_existing = list(dict.fromkeys(existing_numbers))
    if len(existing_numbers) > 1:
        if ui_language == "zh":
            return f"已清理重复公式编号并保留 ({new_number})。"
        return f"Cleaned duplicate equation numbers and kept ({new_number})."
    if len(unique_existing) == 1:
        return _renumber_message("equation", unique_existing[0], new_number, ui_language)
    return _insert_equation_number_message(new_number, ui_language)


def _update_equation_references(
    document,
    number_map: dict[int, int],
    label_indices: set[int],
    ui_language: str,
    start_index: int,
) -> list[FormattingAction]:
    actions: list[FormattingAction] = []
    if not number_map:
        return actions

    for index, paragraph in enumerate(document.paragraphs[start_index:], start=start_index):
        if should_skip_for_body_inventory(paragraph, index, start_index):
            continue
        if index in label_indices or is_equation_label_paragraph(paragraph):
            continue
        before = get_visible_text_from_xml(paragraph) or paragraph.text
        replacements: list[tuple[int, int, str]] = []
        for match in EQUATION_REFERENCE_RE.finditer(before):
            old_number = int(match.group("number"))
            new_number = number_map.get(old_number)
            if new_number is None or new_number == old_number:
                continue
            replacements.append((match.start(), match.end(), f"{match.group('prefix')} ({new_number})"))

        if not replacements:
            continue

        expected = before
        replaced_all = True
        for start, end, replacement in reversed(replacements):
            replaced_all = replace_text_span_in_runs(paragraph, start, end, replacement) and replaced_all
            expected = expected[:start] + replacement + expected[end:]
        after = get_visible_text_from_xml(paragraph) or paragraph.text
        if not replaced_all or after != expected or after == before:
            continue
        actions.append(
            FormattingAction(
                "equation_reference",
                index,
                normalize_space(before),
                normalize_space(after),
                _equation_reference_update_message(ui_language),
            )
        )
    return actions


def _equation_reference_update_message(ui_language: str) -> str:
    if ui_language == "zh":
        return "已根据公式重编号同步更新正文公式引用。"
    return "Updated in-text equation reference after equation renumbering."


def _duplicate_equation_number_cleanup_message(ui_language: str) -> str:
    if ui_language == "zh":
        return "已清理重复公式编号，保留最终编号。"
    return "Cleaned duplicate trailing equation numbers and kept the final label."


def _insert_equation_number_message(new_number: int, ui_language: str) -> str:
    if ui_language == "zh":
        return f"已为未编号公式插入编号 ({new_number})。"
    return f"Inserted equation number ({new_number}) for an unnumbered display equation."
