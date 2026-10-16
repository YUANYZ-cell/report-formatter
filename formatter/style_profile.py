from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass
class StyleItem:
    """Serializable paragraph/run formatting captured from a source paragraph."""

    role: str
    paragraph_style: str | None = None
    font_name: str | None = None
    font_size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    all_caps: bool | None = None
    alignment: str | None = None
    left_indent: float | None = None
    right_indent: float | None = None
    first_line_indent: float | None = None
    space_before: float | None = None
    space_after: float | None = None
    line_spacing: float | None = None
    source: str | None = None
    pattern: str | None = None
    run_level: dict[str, Any] = field(default_factory=dict)
    paragraph_level: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class StyleProfile:
    """Formatting rules inferred from the report/template or fallback defaults."""

    body_text: StyleItem
    section_heading: StyleItem
    figure_caption: StyleItem
    table_number: StyleItem
    table_caption: StyleItem
    equation: StyleItem
    equation_number: StyleItem
    code_block: StyleItem | None = None
    references: StyleItem | None = None
    source: str = "fallback"
    warnings: list[str] = field(default_factory=list)

    def get(self, role: str) -> StyleItem:
        normalized = role.replace("-", "_")
        item = getattr(self, normalized, None)
        if item is None:
            return self.body_text
        return item

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        return data

    def short_summary(self) -> list[str]:
        rows: list[str] = []
        for label, item in [
            ("Body text", self.body_text),
            ("Figure captions", self.figure_caption),
            ("Table number", self.table_number),
            ("Table captions", self.table_caption),
            ("Equations", self.equation),
        ]:
            parts = []
            if item.font_name:
                parts.append(item.font_name)
            if item.font_size:
                parts.append(f"{item.font_size:g} pt")
            if item.alignment:
                parts.append(item.alignment)
            rows.append(f"{label}: {', '.join(parts) if parts else 'not detected'}")
        return rows
