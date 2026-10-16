from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable

from .caption_detector import (
    CaptionMatch,
    detect_figure_captions,
    detect_figure_references,
    detect_table_captions,
    detect_table_inventory,
    match_figure_caption,
    match_table_number,
)
from .equation_detector import detect_equation_numbers
from .equation_utils import EQUATION_REFERENCE_RE
from .format_profiles import preferred_reference
from .localization import CATEGORY_LABELS, label
from .utils import normalize_space, roman_to_int, should_skip_for_body_inventory


TABLE_REF_RE = re.compile(
    r"\b(TABLE|Table)\s+([IVXLCDM]+|\d+)\b|\u8868(?:\u683c)?\s*([IVXLCDM]+|\d+)",
    re.IGNORECASE,
)
TABLE_REVIEW_SOURCES = {"image_table_candidate", "image_table_confirmed", "caption_only_table", "caption_only"}
EQUATION_REF_RE = EQUATION_REFERENCE_RE


@dataclass
class ReferenceIssue:
    severity: str
    category: str
    message: str
    paragraph_index: int | None = None


def check_references(
    document,
    figures: Iterable[CaptionMatch] | None = None,
    tables: Iterable[CaptionMatch] | None = None,
    equations: Iterable[CaptionMatch] | None = None,
    start_index: int = 0,
    format_profile: dict | None = None,
    report_language: str = "en",
    ui_language: str = "en",
) -> list[ReferenceIssue]:
    figures = list(figures if figures is not None else detect_figure_captions(document, start_index))
    tables = list(tables if tables is not None else detect_table_inventory(document, start_index))
    equations = list(equations if equations is not None else detect_equation_numbers(document, start_index))
    format_profile = format_profile or _fallback_reference_profile()

    issues: list[ReferenceIssue] = []
    figure_numbers = [item.number for item in figures]
    confirmed_tables = [item for item in tables if item.source == "native_table"]
    uncertain_tables = [item for item in tables if item.source in TABLE_REVIEW_SOURCES]
    table_numbers = [item.number for item in confirmed_tables]
    equation_numbers = [item.number for item in equations]
    figure_set = set(figure_numbers)
    table_set = set(table_numbers)
    uncertain_table_set = {item.number for item in uncertain_tables}
    equation_set = set(equation_numbers)

    _check_sequence("figure", figure_numbers, issues, ui_language)
    _check_sequence("table", table_numbers, issues, ui_language)
    _check_sequence("equation", equation_numbers, issues, ui_language)

    refs = _collect_references(document, start_index)

    for ref in refs["figures"]:
        preferred = preferred_reference("figure", ref["number"], format_profile)
        if ref["number"] not in figure_set:
            issues.append(
                ReferenceIssue(
                    "error",
                    "figure",
                    _missing_target_message(ref["text"], "figure", ui_language),
                    ref["paragraph_index"],
                )
            )
        elif _normalize_reference(ref["text"]) != _normalize_reference(preferred):
            issues.append(
                ReferenceIssue(
                    "warning",
                    "figure",
                    _inconsistent_style_message(ref["text"], preferred, ui_language),
                    ref["paragraph_index"],
                )
            )

    for ref in refs["tables"]:
        preferred = preferred_reference("table", ref["number"], format_profile)
        if ref["number"] in table_set:
            if _normalize_reference(ref["text"]) != _normalize_reference(preferred):
                issues.append(
                    ReferenceIssue(
                        "warning",
                        "table",
                        _inconsistent_style_message(ref["text"], preferred, ui_language),
                        ref["paragraph_index"],
                    )
                )
        elif ref["number"] in uncertain_table_set:
            issues.append(
                ReferenceIssue(
                    "warning",
                    "table",
                    _possible_image_table_message(ref["text"], ui_language),
                    ref["paragraph_index"],
                )
            )
        else:
            issues.append(
                ReferenceIssue(
                    "error",
                    "table",
                    _missing_target_message(ref["text"], "table", ui_language),
                    ref["paragraph_index"],
                )
            )

    for ref in refs["equations"]:
        preferred = preferred_reference("equation", ref["number"], format_profile)
        if ref["number"] not in equation_set:
            issues.append(
                ReferenceIssue(
                    "error",
                    "equation",
                    _missing_target_message(ref["text"], "equation", ui_language),
                    ref["paragraph_index"],
                )
            )
        elif not _equation_reference_style_ok(ref["text"], preferred):
            issues.append(
                ReferenceIssue(
                    "warning",
                    "equation",
                    _inconsistent_style_message(ref["text"], preferred, ui_language),
                    ref["paragraph_index"],
                )
            )

    _check_unused("figure", figure_set, {ref["number"] for ref in refs["figures"]}, issues, ui_language)
    referenced_tables = {ref["number"] for ref in refs["tables"]}
    _check_unused("table", table_set, referenced_tables, issues, ui_language)
    _check_unused("table", uncertain_table_set, referenced_tables, issues, ui_language)
    _check_unused("equation", equation_set, {ref["number"] for ref in refs["equations"]}, issues, ui_language)

    _check_first_citation_order("figure", refs["figures"], issues, ui_language)
    _check_first_citation_order("table", refs["tables"], issues, ui_language)
    _check_first_citation_order("equation", refs["equations"], issues, ui_language)

    return issues


def _collect_references(document, start_index: int) -> dict[str, list[dict]]:
    caption_indices = _caption_paragraph_indices(document, start_index)
    references = {"figures": [], "tables": [], "equations": []}

    for ref in detect_figure_references(document, start_index):
        references["figures"].append(
            {
                "number": ref.number,
                "text": ref.text,
                "label": "Figure" if ref.text.lower().startswith("figure") else "Fig.",
                "paragraph_index": ref.paragraph_index,
            }
        )

    for index, paragraph in enumerate(document.paragraphs[start_index:], start=start_index):
        if should_skip_for_body_inventory(paragraph, index, start_index):
            continue
        if index in caption_indices:
            continue
        text = normalize_space(paragraph.text)
        if not text or match_figure_caption(text) or match_table_number(text):
            continue

        for match in TABLE_REF_RE.finditer(text):
            label_text = match.group(1) or "Table"
            raw_number = match.group(2) or match.group(3)
            number = int(raw_number) if raw_number.isdigit() else roman_to_int(raw_number) or 0
            references["tables"].append(
                {
                    "number": number,
                    "text": match.group(0),
                    "label": label_text,
                    "raw_number": raw_number,
                    "paragraph_index": index,
                }
            )

        for match in EQUATION_REF_RE.finditer(text):
            raw_number = match.group("number")
            if raw_number:
                references["equations"].append(
                    {
                        "number": int(raw_number),
                        "text": match.group(0).strip(),
                        "label": match.group("prefix"),
                        "paragraph_index": index,
                    }
                )

    return references


def _caption_paragraph_indices(document, start_index: int) -> set[int]:
    indices = {
        item.paragraph_index
        for item in detect_figure_captions(document, start_index)
        if item.paragraph_index is not None and item.source == "paragraph"
    }
    for item in detect_table_captions(document, start_index):
        if item.paragraph_index is not None:
            indices.add(item.paragraph_index)
        if item.caption_paragraph_index is not None:
            indices.add(item.caption_paragraph_index)
    indices.update(item.paragraph_index for item in detect_equation_numbers(document, start_index))
    return indices


def _check_sequence(kind: str, numbers: list[int], issues: list[ReferenceIssue], ui_language: str) -> None:
    if not numbers:
        return
    seen = set()
    for number in numbers:
        if number in seen:
            issues.append(ReferenceIssue("error", kind, _duplicate_number_message(kind, number, ui_language)))
        seen.add(number)
    expected = set(range(1, max(numbers) + 1))
    missing = sorted(expected - set(numbers))
    for number in missing:
        if kind == "figure":
            issues.append(ReferenceIssue("warning", kind, _possible_numbering_gap_message(kind, number, ui_language)))
        else:
            issues.append(ReferenceIssue("error", kind, _missing_number_message(kind, number, ui_language)))


def _check_unused(
    kind: str, detected: set[int], referenced: set[int], issues: list[ReferenceIssue], ui_language: str
) -> None:
    for number in sorted(detected - referenced):
        issues.append(ReferenceIssue("warning", kind, _unused_message(kind, number, ui_language)))


def _check_first_citation_order(
    kind: str, refs: list[dict], issues: list[ReferenceIssue], ui_language: str
) -> None:
    first_seen: list[int] = []
    for ref in refs:
        number = ref["number"]
        if number not in first_seen:
            first_seen.append(number)
    if len(first_seen) < 2:
        return
    if first_seen != sorted(first_seen):
        issues.append(ReferenceIssue("warning", kind, _out_of_order_message(kind, first_seen, ui_language)))


def _normalize_reference(text: str) -> str:
    return re.sub(r"\s+", "", text).lower().replace(".", "")


def _equation_reference_style_ok(ref_text: str, preferred: str) -> bool:
    if _normalize_reference(ref_text) == _normalize_reference(preferred):
        return True
    return EQUATION_REF_RE.fullmatch(ref_text.strip()) is not None


def _missing_target_message(ref_text: str, kind: str, ui_language: str) -> str:
    kind_label = label(CATEGORY_LABELS, ui_language, kind).lower()
    if kind == "equation":
        if ui_language == "zh":
            return f"缺少公式引用：{ref_text}。"
        return f"Missing equation reference: {ref_text}."
    if ui_language == "zh":
        return f"引用 {ref_text}，但未检测到对应的{kind_label}。"
    return f"{ref_text} points to a nonexistent {kind_label}."


def _possible_image_table_message(ref_text: str, ui_language: str) -> str:
    if ui_language == "zh":
        return f"引用 {ref_text}，未检测到原生 Word 表格，但发现疑似图片表格/表格标题，请人工确认。"
    return f"{ref_text} has no native Word table, but a nearby possible image table or table title was found; verify manually."


def _inconsistent_style_message(ref_text: str, preferred: str, ui_language: str) -> str:
    if ui_language == "zh":
        return f"{ref_text} 与当前格式规则不一致；建议写作 {preferred}。"
    return f"{ref_text} is inconsistent with the selected profile; prefer {preferred}."


def _duplicate_number_message(kind: str, number: int, ui_language: str) -> str:
    kind_label = label(CATEGORY_LABELS, ui_language, kind).lower()
    if ui_language == "zh":
        return f"{kind_label}编号 {number} 重复。"
    return f"Duplicate {kind_label} number {number}."


def _missing_number_message(kind: str, number: int, ui_language: str) -> str:
    kind_label = label(CATEGORY_LABELS, ui_language, kind).lower()
    if ui_language == "zh":
        return f"缺少{kind_label}编号 {number}。"
    return f"Missing {kind_label} number {number}."


def _possible_numbering_gap_message(kind: str, number: int, ui_language: str) -> str:
    kind_label = label(CATEGORY_LABELS, ui_language, kind).lower()
    if ui_language == "zh":
        return f"{kind_label}编号 {number} 可能存在跳号，请人工复核。"
    return f"Possible {kind_label} numbering gap at {number}; verify manually."


def _unused_message(kind: str, number: int, ui_language: str) -> str:
    kind_label = label(CATEGORY_LABELS, ui_language, kind)
    if ui_language == "zh":
        return f"{kind_label} {number} 已存在，但正文中未引用。"
    return f"{kind_label} {number} exists but is never referenced."


def _out_of_order_message(kind: str, first_seen: list[int], ui_language: str) -> str:
    kind_label = label(CATEGORY_LABELS, ui_language, kind).lower()
    if ui_language == "zh":
        return f"{kind_label}的首次引用顺序可能异常：{first_seen}。"
    return f"Possible out-of-order first citation for {kind_label}s: {first_seen}."


def _fallback_reference_profile() -> dict:
    return {
        "figure": {"prefix": "Fig.", "number_style": "arabic"},
        "table": {"prefix": "Table", "number_style": "roman_upper"},
        "equation": {"number_style": "arabic", "wrapper": "parentheses"},
        "references": {"figure": "Fig. {n}", "table": "Table {roman}", "equation": "{equation_n}"},
    }
