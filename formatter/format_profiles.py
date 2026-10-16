from __future__ import annotations

import copy
import json
import re
from pathlib import Path
from typing import Any

from .caption_detector import detect_figure_captions, detect_table_captions
from .utils import int_to_roman


PROFILE_DIR = Path(__file__).resolve().parents[1] / "templates" / "format_profiles"


def load_format_profile(profile_name: str) -> dict[str, Any]:
    normalized = normalize_profile_name(profile_name)
    path = PROFILE_DIR / f"{normalized}.json"
    if not path.exists():
        path = PROFILE_DIR / "template_matching.json"
    profile = json.loads(path.read_text(encoding="utf-8"))
    profile["profile_key"] = normalized
    return validate_format_profile(profile)


def normalize_profile_name(profile_name: str | None) -> str:
    value = (profile_name or "template_matching").strip().lower().replace("-", "_").replace(" ", "_")
    aliases = {
        "template": "template_matching",
        "template_matching": "template_matching",
        "ieee": "ieee",
        "custom": "custom_default",
        "custom_default": "custom_default",
    }
    return aliases.get(value, "template_matching")


def validate_format_profile(profile: dict[str, Any]) -> dict[str, Any]:
    profile = copy.deepcopy(profile)
    profile.setdefault("profile_name", "Template-matching")
    profile.setdefault("profile_key", normalize_profile_name(profile.get("profile_name")))
    profile.setdefault("figure", {})
    profile.setdefault("table", {})
    profile.setdefault("equation", {})
    profile.setdefault("references", {})

    profile["figure"].setdefault("prefix", "Fig.")
    profile["figure"].setdefault("number_style", "arabic")
    profile["figure"].setdefault("separator_after_number", ".")
    profile["figure"].setdefault("caption_position", "below")

    profile["table"].setdefault("prefix", "Table")
    profile["table"].setdefault("number_style", "roman_upper")
    profile["table"].setdefault("caption_position", "above")
    profile["table"].setdefault("number_line_only", False)
    profile["table"].setdefault("caption_text_uppercase", False)
    profile["table"].setdefault("separator_after_number", ".")

    profile["equation"].setdefault("number_style", "arabic")
    profile["equation"].setdefault("wrapper", "parentheses")

    profile["references"].setdefault("figure", f"{profile['figure']['prefix']} {{n}}")
    profile["references"].setdefault("table", f"{profile['table']['prefix']} {{table_n}}")
    profile["references"].setdefault("equation", "{equation_n}")
    return profile


def build_custom_profile(user_options: dict[str, Any]) -> dict[str, Any]:
    profile = load_format_profile("custom_default")
    profile["profile_key"] = "custom"
    profile["profile_name"] = "Custom"
    figure = profile["figure"]
    table = profile["table"]
    equation = profile["equation"]

    figure["prefix"] = user_options.get("figure_prefix", figure["prefix"])
    figure["number_style"] = user_options.get("figure_number_style", figure["number_style"])
    figure["separator_after_number"] = user_options.get(
        "figure_separator", figure["separator_after_number"]
    )

    table["prefix"] = user_options.get("table_prefix", table["prefix"])
    table["number_style"] = user_options.get("table_number_style", table["number_style"])
    table_layout = user_options.get("table_layout", "single_line")
    table["number_line_only"] = table_layout == "number_line_plus_caption"
    table["separator_after_number"] = user_options.get("table_separator", table["separator_after_number"])

    equation["wrapper"] = user_options.get("equation_wrapper", equation["wrapper"])
    profile["references"] = _reference_templates(profile)
    return validate_format_profile(profile)


def infer_profile_from_template(document, base_profile: dict[str, Any] | None = None) -> dict[str, Any]:
    profile = copy.deepcopy(base_profile or load_format_profile("template_matching"))
    profile["profile_key"] = "template_matching"
    profile["profile_name"] = "Template-matching"

    figures = detect_figure_captions(document)
    if figures:
        raw = figures[0].raw_text
        if re.match(r"^Figure\s+\d+", raw, re.IGNORECASE):
            profile["figure"]["prefix"] = "Figure"
            profile["figure"]["separator_after_number"] = "."
        elif re.match(r"^图\s*\d+", raw):
            profile["figure"]["prefix"] = "图"
            profile["figure"]["separator_after_number"] = "."
        else:
            profile["figure"]["prefix"] = "Fig."
            profile["figure"]["separator_after_number"] = "."

    tables = detect_table_captions(document)
    if tables:
        raw = tables[0].raw_text
        separator = _detected_caption_separator(raw, r"(?:Table|TABLE)")
        if re.match(r"^TABLE\s+", raw):
            profile["table"]["prefix"] = "TABLE"
            profile["table"]["number_style"] = "roman_upper"
            profile["table"]["number_line_only"] = True
            profile["table"]["caption_text_uppercase"] = True
        elif re.match(r"^表\s+", raw):
            profile["table"]["prefix"] = "表"
            profile["table"]["number_line_only"] = not tables[0].inline_caption
        elif re.search(r"\d+", raw):
            profile["table"]["prefix"] = "Table"
            profile["table"]["number_style"] = "arabic"
            profile["table"]["number_line_only"] = not tables[0].inline_caption
            if separator:
                profile["table"]["separator_after_number"] = separator
        else:
            profile["table"]["prefix"] = "Table"

    profile["references"] = _reference_templates(profile)
    return validate_format_profile(profile)


def _detected_caption_separator(raw_text: str, prefix_pattern: str) -> str | None:
    match = re.match(
        rf"^\s*{prefix_pattern}\s+(?:[IVXLCDM]+|\d+)\s*(?P<separator>[.:\uff1a])",
        raw_text,
        re.IGNORECASE,
    )
    if not match:
        return None
    if match.group("separator") in {":", "\uff1a"}:
        return ":"
    return "."


def format_number(number: int, style: str) -> str:
    if style == "roman_upper":
        return int_to_roman(number)
    if style == "roman_lower":
        return int_to_roman(number).lower()
    return str(number)


def format_figure_caption(
    number: int,
    caption_text: str,
    profile: dict[str, Any],
    report_language: str,
    placeholder: bool = False,
) -> str:
    figure = profile["figure"]
    prefix = figure.get("prefix") or ("图" if report_language == "zh" else "Fig.")
    number_text = format_number(number, figure.get("number_style", "arabic"))
    separator = figure.get("separator_after_number", ".")
    text = caption_text.strip()
    if placeholder or not text:
        text = "请输入图片标题" if report_language == "zh" else "Enter figure caption here"
    return f"{prefix} {number_text}{separator} {text}".strip()


def format_table_caption(
    number: int,
    caption_text: str,
    profile: dict[str, Any],
    report_language: str,
    placeholder: bool = False,
) -> list[str] | str:
    table = profile["table"]
    prefix = table.get("prefix") or ("表" if report_language == "zh" else "Table")
    number_text = format_number(number, table.get("number_style", "roman_upper"))
    text = caption_text.strip()
    if placeholder or not text:
        text = "请输入表格标题" if report_language == "zh" else "ENTER TABLE CAPTION HERE"
    if table.get("caption_text_uppercase") and report_language != "zh":
        text = text.upper()
    if table.get("number_line_only"):
        return [f"{prefix} {number_text}".strip(), text]
    separator = table.get("separator_after_number", ".")
    return f"{prefix} {number_text}{separator} {text}".strip()


def format_equation_number(number: int, profile: dict[str, Any]) -> str:
    equation = profile["equation"]
    number_text = format_number(number, equation.get("number_style", "arabic"))
    wrapper = equation.get("wrapper", "parentheses")
    if wrapper == "brackets":
        return f"[{number_text}]"
    if wrapper == "eq_parentheses":
        return f"Eq. ({number_text})"
    if wrapper == "equation_parentheses":
        return f"Equation ({number_text})"
    return f"({number_text})"


def preferred_reference(kind: str, number: int, profile: dict[str, Any]) -> str:
    table_number = format_number(number, profile["table"].get("number_style", "roman_upper"))
    equation_number = format_equation_number(number, profile)
    template = profile.get("references", {}).get(kind, "")
    if kind == "figure":
        return template.format(n=number)
    if kind == "table":
        return template.format(n=number, roman=table_number, table_n=table_number)
    if kind == "equation":
        return template.format(n=number, equation_n=equation_number)
    return str(number)


def _reference_templates(profile: dict[str, Any]) -> dict[str, str]:
    figure_prefix = profile["figure"].get("prefix", "Fig.")
    table_prefix = profile["table"].get("prefix", "Table")
    table_n = "{table_n}"
    return {
        "figure": f"{figure_prefix} {{n}}",
        "table": f"{table_prefix} {table_n}",
        "equation": "{equation_n}",
    }
