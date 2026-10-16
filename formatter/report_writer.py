from __future__ import annotations

from collections import Counter

from .caption_detector import CaptionMatch
from .caption_rules import FormattingAction
from .image_table_detector import ImageScanResult
from .localization import (
    ALIGNMENT_LABELS,
    CATEGORY_LABELS,
    LANGUAGE_NAMES,
    PROFILE_LABELS,
    ROLE_LABELS,
    SEVERITY_LABELS,
    label,
)
from .reference_checker import ReferenceIssue
from .style_profile import StyleItem, StyleProfile


def generate_check_report(
    style_profile: StyleProfile,
    figures: list[CaptionMatch],
    tables: list[CaptionMatch],
    equations: list[CaptionMatch],
    issues: list[ReferenceIssue],
    actions: list[FormattingAction],
    manual_review: list[str],
    ui_language: str = "en",
    language_detection: dict | None = None,
    report_language: str = "en",
    format_profile: dict | None = None,
    style_source: str = "uploaded report",
    reference_citations: list[dict] | None = None,
    native_tables: list[CaptionMatch] | None = None,
    image_tables: list[CaptionMatch] | None = None,
    images: list[ImageScanResult] | None = None,
    equation_blocks: list[CaptionMatch] | None = None,
    unnumbered_equations: list[CaptionMatch] | None = None,
    skipped_toc_entries_count: int = 0,
) -> str:
    return _generate_zh(
        style_profile,
        figures,
        tables,
        equations,
        reference_citations or [],
        native_tables or [],
        image_tables or [],
        images or [],
        equation_blocks or equations,
        unnumbered_equations or [],
        skipped_toc_entries_count,
        issues,
        actions,
        manual_review,
        language_detection or {},
        report_language,
        format_profile or {},
        style_source,
    ) if ui_language == "zh" else _generate_en(
        style_profile,
        figures,
        tables,
        equations,
        reference_citations or [],
        native_tables or [],
        image_tables or [],
        images or [],
        equation_blocks or equations,
        unnumbered_equations or [],
        skipped_toc_entries_count,
        issues,
        actions,
        manual_review,
        language_detection or {},
        report_language,
        format_profile or {},
        style_source,
    )


def _generate_en(
    style_profile: StyleProfile,
    figures: list[CaptionMatch],
    tables: list[CaptionMatch],
    equations: list[CaptionMatch],
    reference_citations: list[dict],
    native_tables: list[CaptionMatch],
    image_tables: list[CaptionMatch],
    images: list[ImageScanResult],
    equation_blocks: list[CaptionMatch],
    unnumbered_equations: list[CaptionMatch],
    skipped_toc_entries_count: int,
    issues: list[ReferenceIssue],
    actions: list[FormattingAction],
    manual_review: list[str],
    language_detection: dict,
    report_language: str,
    format_profile: dict,
    style_source: str,
) -> str:
    counts = Counter(action.kind for action in actions)
    lines = [
        "Template-aware Lab Report Formatter Check Report",
        "",
        "Summary:",
        f"- UI language selected: English",
        f"- Detected report language: {_language_name('en', language_detection.get('language', 'unknown'))}",
        f"- Final report language used: {_language_name('en', report_language)}",
        f"- Formatting profile selected: {_profile_name('en', format_profile)}",
        f"- Style source used: {style_source}",
        f"- Figures detected: {len(figures)}",
        f"- Native tables detected: {len(native_tables)}",
        f"- Image table candidates detected: {len(image_tables)}",
        f"- Images scanned: {len(images)}",
        f"- Tables detected: {len(tables)}",
        f"- Equation blocks detected: {len(equation_blocks)}",
        f"- Equation numbers detected: {len(equations)}",
        f"- Unnumbered equation blocks detected: {len(unnumbered_equations)}",
        f"- IEEE reference citations detected: {len(reference_citations)}",
        f"- Skipped TOC/front-matter entries: {skipped_toc_entries_count}",
        f"- Figure captions fixed: {counts.get('figure', 0)}",
        f"- Table captions fixed: {counts.get('table', 0)}",
        f"- Equation numbers fixed: {counts.get('equation', 0)}",
        f"- Reference issues found: {len(issues)}",
        "",
        "Language detection:",
        f"- Confidence: {language_detection.get('confidence', 0):.2f}",
        f"- Chinese characters: {language_detection.get('chinese_char_count', 0)}",
        f"- English words: {language_detection.get('english_word_count', 0)}",
        "",
        "Detected style profile:",
    ]
    lines.extend(f"- {row}" for row in _style_summary("en", style_profile))
    lines.append("")

    _append_image_scan_section(lines, images, "en")
    _append_issue_section(lines, "Figure issues:", [issue for issue in issues if issue.category == "figure"], "en")
    _append_issue_section(lines, "Table issues:", [issue for issue in issues if issue.category == "table"], "en")
    _append_issue_section(lines, "Equation issues:", [issue for issue in issues if issue.category == "equation"], "en")
    _append_actions(lines, actions, "en")
    _append_manual_review(lines, manual_review, "en")
    lines.append("Notes:")
    lines.append("- This tool is rule-based and does not use AI APIs, external APIs, or local ML models.")
    lines.append("- Complex floating images, MathType objects, screenshot equations, and caption repositioning may require manual review.")
    return "\n".join(lines)


def _generate_zh(
    style_profile: StyleProfile,
    figures: list[CaptionMatch],
    tables: list[CaptionMatch],
    equations: list[CaptionMatch],
    reference_citations: list[dict],
    native_tables: list[CaptionMatch],
    image_tables: list[CaptionMatch],
    images: list[ImageScanResult],
    equation_blocks: list[CaptionMatch],
    unnumbered_equations: list[CaptionMatch],
    skipped_toc_entries_count: int,
    issues: list[ReferenceIssue],
    actions: list[FormattingAction],
    manual_review: list[str],
    language_detection: dict,
    report_language: str,
    format_profile: dict,
    style_source: str,
) -> str:
    counts = Counter(action.kind for action in actions)
    lines = [
        "模板感知实验报告格式化工具检查报告",
        "",
        "摘要：",
        "- 已选择界面语言：中文",
        f"- 自动识别的报告语言：{_language_name('zh', language_detection.get('language', 'unknown'))}",
        f"- 最终使用的报告语言：{_language_name('zh', report_language)}",
        f"- 已选择格式规则：{_profile_name('zh', format_profile)}",
        f"- 样式来源：{_style_source_zh(style_source)}",
        f"- 检测到图题：{len(figures)}",
        f"- 检测到表题：{len(tables)}",
        f"- 检测到公式：{len(equations)}",
        f"- 已修正图题：{counts.get('figure', 0)}",
        f"- 已修正表题：{counts.get('table', 0)}",
        f"- 已修正公式编号：{counts.get('equation', 0)}",
        f"- 发现引用问题：{len(issues)}",
        "",
        "语言检测：",
        f"- 置信度：{language_detection.get('confidence', 0):.2f}",
        f"- 中文字符数：{language_detection.get('chinese_char_count', 0)}",
        f"- 英文单词数：{language_detection.get('english_word_count', 0)}",
        "",
        "检测到的样式画像：",
    ]
    lines.extend(f"- {row}" for row in _style_summary("zh", style_profile))
    lines.append("")
    _append_image_scan_section(lines, images, "zh")

    _append_issue_section(lines, "图题问题：", [issue for issue in issues if issue.category == "figure"], "zh")
    _append_issue_section(lines, "表题问题：", [issue for issue in issues if issue.category == "table"], "zh")
    _append_issue_section(lines, "公式问题：", [issue for issue in issues if issue.category == "equation"], "zh")
    _append_actions(lines, actions, "zh")
    _append_manual_review(lines, manual_review, "zh")
    lines.append("说明：")
    lines.append("- 本工具完全基于规则，不使用 AI API、外部 API 或本地机器学习模型。")
    lines.append("- 复杂浮动图片、MathType 公式、截图公式和题注位置移动可能需要人工复核。")
    return "\n".join(lines)


def _append_issue_section(lines: list[str], title: str, issues: list[ReferenceIssue], language: str) -> None:
    lines.append(title)
    if issues:
        for issue in issues:
            severity = label(SEVERITY_LABELS, language, issue.severity)
            location = _paragraph_location(language, issue.paragraph_index)
            lines.append(f"- {severity}: {issue.message}{location}" if language == "en" else f"- {severity}：{issue.message}{location}")
    else:
        lines.append("- None." if language == "en" else "- 无。")
    lines.append("")


def _append_actions(lines: list[str], actions: list[FormattingAction], language: str) -> None:
    lines.append("Formatting actions performed:" if language == "en" else "已执行的格式化动作：")
    if actions:
        for action in actions:
            lines.append(f"- {action.message}{_paragraph_location(language, action.paragraph_index)}")
    else:
        lines.append("- No formatting actions were performed." if language == "en" else "- 未执行格式化动作。")
    lines.append("")


def _append_manual_review(lines: list[str], manual_review: list[str], language: str) -> None:
    lines.append("Items requiring manual review:" if language == "en" else "需要人工复核的项目：")
    if manual_review:
        lines.extend(f"- {item}" for item in manual_review)
    else:
        lines.append("- None detected by rule-based checks." if language == "en" else "- 规则检查未发现需要人工复核的项目。")
    lines.append("")


def _append_image_scan_section(lines: list[str], images: list[ImageScanResult], language: str) -> None:
    lines.append("Image scan:" if language == "en" else "图片扫描：")
    if not images:
        lines.append("- No images detected." if language == "en" else "- 未检测到图片。")
        lines.append("")
        return

    for image in images:
        if image.is_suspected_table:
            kind = "Suspected image table" if language == "en" else "疑似图片表格"
        elif image.is_suspected_figure:
            kind = "Suspected image figure" if language == "en" else "疑似图片图"
        else:
            kind = "Unparsed image" if language == "en" else "未解析图片"
        title = image.ocr_title_text or ("not detected" if language == "en" else "未识别")
        location = _image_location(language, image)
        caption_type = image.caption_type or ("not detected" if language == "en" else "未识别")
        line = (
            f"- {kind}: {location}, file: {image.file_name}, size: {_image_size_text(image)}, "
            f"OCR title: {title}, caption type: {caption_type}, confidence: {image.confidence:.2f}."
            if language == "en"
            else f"- {kind}：{location}，文件：{image.file_name}，尺寸：{_image_size_text(image)}，"
            f"OCR 标题：{title}，caption 类型：{caption_type}，置信度：{image.confidence:.2f}。"
        )
        if image.caption_type == "table_caption_embedded_in_image":
            line += (
                " Caption is embedded in the image; move it outside as a Word text caption."
                if language == "en"
                else " 该表题嵌入图片，无法安全自动改写，请手动把表题移到图片外作为 Word 文本 caption。"
            )
        elif not image.ocr_available:
            reason = image.ocr_error or ("OCR unavailable" if language == "en" else "OCR 不可用")
            line += (" OCR unavailable: " + reason if language == "en" else f" OCR 不可用：{reason}")
        lines.append(line)
    lines.append("")


def _image_location(language: str, image: ImageScanResult) -> str:
    if image.paragraph_index is None:
        return f"image {image.image_index}, paragraph unknown" if language == "en" else f"第 {image.image_index} 张图片，附近段落未知"
    paragraph = image.paragraph_index + 1
    return f"image {image.image_index}, near paragraph {paragraph}" if language == "en" else f"第 {image.image_index} 张图片，附近第 {paragraph} 段"


def _image_size_text(image: ImageScanResult) -> str:
    if image.width_px and image.height_px:
        return f"{image.width_px}x{image.height_px}px"
    if image.width_inches and image.height_inches:
        return f"{image.width_inches}x{image.height_inches}in"
    return "unknown"


def _style_summary(language: str, profile: StyleProfile) -> list[str]:
    rows = []
    for role, item in [
        ("body_text", profile.body_text),
        ("figure_caption", profile.figure_caption),
        ("table_number", profile.table_number),
        ("table_caption", profile.table_caption),
        ("equation", profile.equation),
    ]:
        rows.append(f"{label(ROLE_LABELS, language, role)}: {_style_item_summary(language, item)}")
    return rows


def _style_item_summary(language: str, item: StyleItem) -> str:
    parts = []
    if item.font_name:
        parts.append(item.font_name)
    if item.font_size:
        parts.append(f"{item.font_size:g} pt")
    if item.alignment:
        parts.append(label(ALIGNMENT_LABELS, language, item.alignment))
    if not parts:
        return "not detected" if language == "en" else "未检测到"
    return ", ".join(parts) if language == "en" else "，".join(parts)


def _paragraph_location(language: str, paragraph_index: int | None) -> str:
    if paragraph_index is None:
        return ""
    return f" (paragraph {paragraph_index + 1})" if language == "en" else f"（第 {paragraph_index + 1} 段）"


def _language_name(language: str, value: str) -> str:
    return LANGUAGE_NAMES.get(language, LANGUAGE_NAMES["en"]).get(value, value)


def _profile_name(language: str, format_profile: dict) -> str:
    key = format_profile.get("profile_key") or format_profile.get("profile_name", "")
    if key == "custom_default":
        key = "custom"
    return label(PROFILE_LABELS, language, key)


def _style_source_zh(style_source: str) -> str:
    return {
        "Auto-detect from uploaded report": "上传报告自动识别",
        "Use uploaded template as style source": "上传模板",
        "Use fallback defaults": "默认样式",
        "Use fallback IEEE defaults": "默认样式",
    }.get(style_source, style_source)
