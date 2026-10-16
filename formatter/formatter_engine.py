from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from .caption_detector import dedupe_table_inventory, detect_figure_captions, detect_table_inventory
from .caption_rules import (
    FormattingAction,
    renumber_figure_captions,
    renumber_table_captions,
)
from .equation_detector import detect_equation_blocks, detect_equation_numbers, detect_unnumbered_equation_blocks
from .equation_utils import detect_ieee_citation_tokens
from .equation_rules import cleanup_duplicate_equation_numbers, renumber_equations
from .figure_handler import find_uncaptioned_figures, insert_missing_figure_placeholders
from .heading_rules import renumber_ieee_headings
from .format_profiles import (
    build_custom_profile,
    infer_profile_from_template,
    load_format_profile,
    normalize_profile_name,
)
from .image_table_detector import ImageScanResult, detect_document_images, image_table_matches_from_records
from .language_detector import detect_report_language, resolve_report_language
from .object_scanner import detect_main_report_start
from .reference_checker import ReferenceIssue, check_references
from .report_writer import generate_check_report
from .style_extractor import extract_style_profile, load_fallback_profile
from .style_profile import StyleProfile
from .table_handler import find_uncaptioned_tables, insert_missing_table_placeholders
from .utils import count_skipped_toc_entries


FORMATTER_VERSION = "2026-06-25-p0-safe-equation-pixel-diff-v2"
EQUATION_AUTO_EDIT_ENABLED = False


class ProcessingMode(str, Enum):
    ANALYZE_ONLY = "Analyze only"
    FIX_NUMBERING_AND_FORMATTING = "Fix numbering and formatting"
    FIX_NUMBERING_ONLY = "Fix numbering only"
    CHECK_REFERENCES_ONLY = "Check references only"


@dataclass
class FormattingOptions:
    preserve_cover_page: bool = True
    preserve_two_column_layout: bool = True
    insert_placeholder_captions: bool = False
    renumber_figures: bool = True
    renumber_tables: bool = True
    renumber_equations: bool = True
    renumber_headings: bool = True
    check_in_text_references: bool = True
    generate_check_report: bool = True
    ui_language: str = "en"
    report_language_override: str = "auto"
    format_profile_name: str = "template_matching"
    custom_profile_options: dict = field(default_factory=dict)


@dataclass
class ProcessResult:
    style_profile: StyleProfile
    start_index: int
    language_detection: dict
    report_language: str
    ui_language: str
    format_profile: dict
    style_source: str
    figures: list = field(default_factory=list)
    tables: list = field(default_factory=list)
    native_tables: list = field(default_factory=list)
    image_tables: list = field(default_factory=list)
    images: list[ImageScanResult] = field(default_factory=list)
    equation_blocks: list = field(default_factory=list)
    equations: list = field(default_factory=list)
    unnumbered_equations: list = field(default_factory=list)
    reference_citations: list = field(default_factory=list)
    skipped_toc_entries_count: int = 0
    issues: list[ReferenceIssue] = field(default_factory=list)
    actions: list[FormattingAction] = field(default_factory=list)
    manual_review: list[str] = field(default_factory=list)
    headings: list = field(default_factory=list)
    report_text: str = ""


def process_document(
    document,
    template_document=None,
    mode: str | ProcessingMode = ProcessingMode.ANALYZE_ONLY,
    style_source: str = "Auto-detect from uploaded report",
    options: FormattingOptions | None = None,
) -> ProcessResult:
    options = options or FormattingOptions()
    mode = ProcessingMode(mode)
    start_index = detect_main_report_start(document) if options.preserve_cover_page else 0
    language_detection = detect_report_language(document)
    report_language = resolve_report_language(language_detection, options.report_language_override)
    style_profile = _select_style_profile(document, template_document, style_source)
    format_profile = _select_format_profile(document, template_document, options)

    actions: list[FormattingAction] = []
    manual_review: list[str] = []
    images = detect_document_images(document, start_index)

    image_paragraph_indices = {image.paragraph_index for image in images if image.paragraph_index is not None}
    missing_figures = [
        item for item in find_uncaptioned_figures(document, start_index) if item.paragraph_index not in image_paragraph_indices
    ]
    missing_tables = find_uncaptioned_tables(document, start_index)
    manual_review.extend(_missing_figure_message(item.paragraph_index, options.ui_language) for item in missing_figures)
    manual_review.extend(
        _missing_table_message(item.paragraph_index_before_table, options.ui_language) for item in missing_tables
    )
    manual_review.extend(_image_review_messages(images, options.ui_language))

    should_fix = mode in {
        ProcessingMode.FIX_NUMBERING_AND_FORMATTING,
        ProcessingMode.FIX_NUMBERING_ONLY,
    }
    apply_formatting = mode == ProcessingMode.FIX_NUMBERING_AND_FORMATTING

    if should_fix and options.insert_placeholder_captions:
        manual_review.extend(
            insert_missing_figure_placeholders(
                document, style_profile, format_profile, report_language, options.ui_language, start_index
            )
        )
        manual_review.extend(
            insert_missing_table_placeholders(
                document, style_profile, format_profile, report_language, options.ui_language, start_index
            )
        )

    if should_fix:
        if options.renumber_figures:
            figure_actions, _ = renumber_figure_captions(
                document,
                style_profile,
                format_profile,
                report_language,
                options.ui_language,
                start_index,
                apply_formatting,
            )
            actions.extend(figure_actions)
        if options.renumber_tables:
            table_actions, _ = renumber_table_captions(
                document,
                style_profile,
                format_profile,
                report_language,
                options.ui_language,
                start_index,
                apply_formatting,
            )
            actions.extend(table_actions)
        if options.renumber_equations:
            if EQUATION_AUTO_EDIT_ENABLED:
                equation_actions, _ = renumber_equations(
                    document,
                    style_profile,
                    format_profile,
                    options.ui_language,
                    start_index,
                    apply_formatting,
                )
                actions.extend(equation_actions)
                actions.extend(cleanup_duplicate_equation_numbers(document, options.ui_language, start_index))
            else:
                manual_review.append(_equation_auto_edit_suspended_message(options.ui_language))
        if options.renumber_headings and format_profile.get("profile_key") == "ieee":
            actions.extend(
                renumber_ieee_headings(
                    document,
                    start_index=start_index,
                    ui_language=options.ui_language,
                    apply_formatting=apply_formatting,
                )
            )

    figures = detect_figure_captions(document, start_index)
    tables = dedupe_table_inventory(
        [*detect_table_inventory(document, start_index), *image_table_matches_from_records(images)]
    )
    native_tables = [table for table in tables if table.source == "native_table"]
    image_tables = [table for table in tables if table.source in {"image_table_confirmed", "image_table_candidate"}]
    equation_blocks = detect_equation_blocks(document, start_index)
    equations = detect_equation_numbers(document, start_index)
    unnumbered_equations = detect_unnumbered_equation_blocks(document, start_index)
    reference_citations = detect_ieee_citation_tokens(document, start_index)
    skipped_toc_entries_count = count_skipped_toc_entries(document, start_index)

    issues: list[ReferenceIssue] = []
    if options.check_in_text_references or mode == ProcessingMode.CHECK_REFERENCES_ONLY:
        issues = check_references(
            document,
            figures,
            tables,
            equations,
            start_index,
            format_profile,
            report_language,
            options.ui_language,
        )

    report_text = ""
    if options.generate_check_report:
        report_text = generate_check_report(
            style_profile=style_profile,
            figures=figures,
            tables=tables,
            native_tables=native_tables,
            image_tables=image_tables,
            images=images,
            equation_blocks=equation_blocks,
            equations=equations,
            unnumbered_equations=unnumbered_equations,
            reference_citations=reference_citations,
            skipped_toc_entries_count=skipped_toc_entries_count,
            issues=issues,
            actions=actions,
            manual_review=manual_review,
            ui_language=options.ui_language,
            language_detection=language_detection,
            report_language=report_language,
            format_profile=format_profile,
            style_source=style_source,
        )

    return ProcessResult(
        style_profile=style_profile,
        start_index=start_index,
        language_detection=language_detection,
        report_language=report_language,
        ui_language=options.ui_language,
        format_profile=format_profile,
        style_source=style_source,
        figures=figures,
        tables=tables,
        native_tables=native_tables,
        image_tables=image_tables,
        images=images,
        equation_blocks=equation_blocks,
        equations=equations,
        unnumbered_equations=unnumbered_equations,
        reference_citations=reference_citations,
        skipped_toc_entries_count=skipped_toc_entries_count,
        issues=issues,
        actions=actions,
        manual_review=manual_review,
        headings=[action for action in actions if action.kind == "heading"],
        report_text=report_text,
    )


def _select_style_profile(document, template_document, style_source: str) -> StyleProfile:
    if style_source in {"Use fallback defaults", "Use fallback IEEE defaults"}:
        return load_fallback_profile()
    if style_source == "Use uploaded template as style source" and template_document is not None:
        return extract_style_profile(document, template_document)
    return extract_style_profile(document)


def _equation_auto_edit_suspended_message(ui_language: str) -> str:
    if ui_language == "zh":
        return (
            "\u5b89\u5168\u4fdd\u62a4\uff1a\u516c\u5f0f\u81ea\u52a8\u91cd\u7f16\u53f7\u5df2\u6682\u505c\u3002"
            "\u672c\u6b21\u4ec5\u62a5\u544a\u68c0\u6d4b\u7ed3\u679c\uff0c"
            "\u4e0d\u63d2\u5165\u3001\u4e0d\u987a\u5ef6\u3001\u4e0d\u6e05\u7406\u516c\u5f0f\u7f16\u53f7\u3002"
        )
    return (
        "Safety guard: automatic equation renumbering is suspended. "
        "The run reports detected equations only and does not insert, shift, or clean equation numbers."
    )


def _select_format_profile(document, template_document, options: FormattingOptions) -> dict:
    profile_name = normalize_profile_name(options.format_profile_name)
    if profile_name == "custom_default":
        return build_custom_profile(options.custom_profile_options)
    if profile_name == "ieee":
        return load_format_profile("ieee")
    source_document = template_document or document
    return infer_profile_from_template(source_document, load_format_profile("template_matching"))


def _missing_figure_message(paragraph_index: int, ui_language: str) -> str:
    if ui_language == "zh":
        return f"第 {paragraph_index + 1} 段附近可能存在未添加图题的图片。"
    return f"Possible figure near paragraph {paragraph_index + 1} has no following figure caption."


def _missing_table_message(paragraph_index: int, ui_language: str) -> str:
    if ui_language == "zh":
        return f"第 {paragraph_index + 1} 段之前可能存在未添加表题的表格。"
    return f"Possible table before paragraph {paragraph_index + 1} has no table caption."


def _image_review_messages(images: list[ImageScanResult], ui_language: str) -> list[str]:
    messages: list[str] = []
    for image in images:
        if image.is_suspected_table:
            messages.append(_image_table_review_message(image, ui_language))
            continue
        if not image.ocr_available:
            messages.append(_unparsed_image_message(image, ui_language))
            continue
        if image.is_suspected_figure and image.caption_type == "figure_caption_embedded_in_image":
            messages.append(_embedded_figure_caption_message(image, ui_language))
    return messages


def _image_table_review_message(image: ImageScanResult, ui_language: str) -> str:
    title = image.ocr_title_text or "not detected"
    location = _image_location(image, ui_language)
    size = _image_size_text(image)
    if ui_language == "zh":
        message = f"疑似图片表格：{location}，文件：{image.file_name}，尺寸：{size}，OCR 标题：{title}，置信度：{image.confidence:.2f}。"
        if image.caption_type == "table_caption_embedded_in_image":
            message += " 该表题嵌入图片，无法安全自动改写，请手动把表题移到图片外作为 Word 文本 caption。"
        elif not image.ocr_available:
            message += " OCR 不可用，已列入未解析图片，需要人工检查。"
        return message
    message = (
        f"Suspected image table: {location}, file: {image.file_name}, size: {size}, "
        f"OCR title: {title}, confidence: {image.confidence:.2f}."
    )
    if image.caption_type == "table_caption_embedded_in_image":
        message += " The table caption is embedded in the image and cannot be safely rewritten; move it outside the image as a Word caption."
    elif not image.ocr_available:
        message += " OCR is unavailable, so this image requires manual review."
    return message


def _unparsed_image_message(image: ImageScanResult, ui_language: str) -> str:
    location = _image_location(image, ui_language)
    size = _image_size_text(image)
    error = image.ocr_error or "OCR unavailable"
    if ui_language == "zh":
        return f"未解析图片，需要人工检查：{location}，文件：{image.file_name}，尺寸：{size}，原因：{error}。"
    return f"Unparsed image requires manual review: {location}, file: {image.file_name}, size: {size}, reason: {error}."


def _embedded_figure_caption_message(image: ImageScanResult, ui_language: str) -> str:
    title = image.ocr_title_text or "not detected"
    location = _image_location(image, ui_language)
    if ui_language == "zh":
        return f"检测到嵌入图片的图题：{location}，文件：{image.file_name}，OCR 标题：{title}。请手动移到图片外作为 Word 文本 caption。"
    return (
        f"Detected figure caption embedded in image: {location}, file: {image.file_name}, OCR title: {title}. "
        "Move it outside the image as a Word caption."
    )


def _image_location(image: ImageScanResult, ui_language: str) -> str:
    if image.paragraph_index is None:
        return f"第 {image.image_index} 张图片，附近段落未知" if ui_language == "zh" else f"image {image.image_index}, paragraph unknown"
    paragraph = image.paragraph_index + 1
    return f"第 {image.image_index} 张图片，附近第 {paragraph} 段" if ui_language == "zh" else f"image {image.image_index}, near paragraph {paragraph}"


def _image_size_text(image: ImageScanResult) -> str:
    if image.width_px and image.height_px:
        return f"{image.width_px}x{image.height_px}px"
    if image.width_inches and image.height_inches:
        return f"{image.width_inches}x{image.height_inches}in"
    return "unknown"
