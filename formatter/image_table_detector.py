from __future__ import annotations

import re
from dataclasses import dataclass, field
from functools import lru_cache
from io import BytesIO
from pathlib import PurePosixPath
from typing import Any

from docx.oxml.ns import qn
from lxml import etree

from .caption_detector import CaptionMatch, get_document_blocks_in_order, match_figure_caption, match_table_number
from .utils import normalize_space, roman_to_int, should_skip_for_body_inventory


EMU_PER_INCH = 914400

TABLE_CAPTION_OCR_RE = re.compile(
    r"(?im)^\s*(?P<prefix>Table|TABLE|\u8868(?:\u683c)?)\s+"
    r"(?P<number>[IVXLCDM]+|\d+)\b[.:\uff1a]?\s*(?P<title>.*)$"
)
FIGURE_CAPTION_OCR_RE = re.compile(
    r"(?im)^\s*(?P<prefix>Fig\.?|Figure|\u56fe)\s+"
    r"(?P<number>\d+[A-Za-z]?)\b[.:\uff1a]?\s*(?P<title>.*)$"
)
TABLE_MENTION_OCR_RE = re.compile(
    r"\b(?:Table|TABLE)\s+(?P<number>[IVXLCDM]+|\d+)\b|\u8868(?:\u683c)?\s*(?P<zh_number>\d+)",
    re.IGNORECASE,
)
TABLE_COLUMN_HINT_RE = re.compile(
    r"\b(load|current|voltage|power|speed|torque|time|frequency|resistance|"
    r"load\s*condition|secondary|primary|measured|calculated|unit|value)\b",
    re.IGNORECASE,
)


@dataclass
class OcrResult:
    available: bool
    engine: str | None
    text: str = ""
    error: str | None = None


@dataclass
class TableStructureResult:
    is_table: bool
    confidence: float
    horizontal_lines: int = 0
    vertical_lines: int = 0
    method: str = "unavailable"
    error: str | None = None


@dataclass
class NearbyParagraph:
    paragraph_index: int | None
    text: str


@dataclass
class ImageScanResult:
    image_index: int
    paragraph_index: int | None
    block_index: int | None
    relationship_id: str | None
    file_name: str
    content_type: str | None
    shape_type: str
    width_px: int | None = None
    height_px: int | None = None
    width_inches: float | None = None
    height_inches: float | None = None
    nearby_paragraphs: list[NearbyParagraph] = field(default_factory=list)
    ocr_text: str = ""
    ocr_engine: str | None = None
    ocr_available: bool = False
    ocr_error: str | None = None
    ocr_title_text: str = ""
    caption_type: str | None = None
    caption_number_token: str | None = None
    caption_number: int | None = None
    caption_text: str = ""
    table_structure: TableStructureResult = field(default_factory=lambda: TableStructureResult(False, 0.0))
    column_hints: list[str] = field(default_factory=list)
    is_suspected_table: bool = False
    is_suspected_figure: bool = False
    confidence: float = 0.0


@dataclass
class _ImageOccurrence:
    relationship_id: str | None
    shape_type: str
    width_emu: int | None = None
    height_emu: int | None = None


def detect_document_images(document, start_index: int = 0) -> list[ImageScanResult]:
    """Scan DOCX drawing/pict image occurrences and classify likely image tables."""

    blocks = get_document_blocks_in_order(document)
    records: list[ImageScanResult] = []
    for block in blocks:
        if block.paragraph is None:
            continue
        if should_skip_for_body_inventory(block.paragraph):
            continue
        for occurrence in _iter_paragraph_image_occurrences(block.paragraph):
            record = _build_image_record(
                document=document,
                paragraph=block.paragraph,
                occurrence=occurrence,
                image_index=len(records) + 1,
                paragraph_index=block.paragraph_index,
                block_index=block.block_index,
                blocks=blocks,
            )
            records.append(record)
    return records


def image_table_matches_from_records(records: list[ImageScanResult]) -> list[CaptionMatch]:
    matches: list[CaptionMatch] = []
    for record in records:
        if not record.is_suspected_table or not record.caption_number:
            continue
        source = (
            "image_table_confirmed"
            if record.caption_type == "table_caption_embedded_in_image" or record.table_structure.is_table
            else "image_table_candidate"
        )
        matches.append(
            CaptionMatch(
                kind="table",
                paragraph_index=record.paragraph_index,
                number=record.caption_number,
                text=record.caption_text,
                raw_text=record.ocr_title_text or f"Image {record.image_index}: {record.file_name}",
                pattern=record.caption_type or "image_table_detection",
                caption_paragraph_index=None,
                prefix="Table",
                number_token=record.caption_number_token,
                inline_caption=True,
                source=source,
                confidence=record.confidence,
                block_index=record.block_index,
                paragraph=None,
            )
        )
    return matches


def _build_image_record(
    *,
    document,
    paragraph,
    occurrence: _ImageOccurrence,
    image_index: int,
    paragraph_index: int | None,
    block_index: int | None,
    blocks,
) -> ImageScanResult:
    part = _related_image_part(paragraph, occurrence.relationship_id)
    blob = getattr(part, "blob", None) if part is not None else None
    file_name = _image_file_name(part, occurrence.relationship_id, image_index)
    width_px, height_px = _image_pixel_size(blob)
    width_inches = _emu_to_inches(occurrence.width_emu)
    height_inches = _emu_to_inches(occurrence.height_emu)
    nearby = _nearby_paragraphs(document, blocks, block_index, paragraph_index)
    ocr = _ocr_image_blob(blob) if blob else OcrResult(False, None, error="Image data is not embedded in the DOCX.")
    structure = _detect_table_structure(blob) if blob else TableStructureResult(False, 0.0, error="No embedded image data.")
    column_hints = _column_hints(ocr.text)
    caption = _caption_from_ocr(ocr.text)
    nearby_table = _nearby_table_caption_or_mention(nearby)
    nearby_figure = _nearby_figure_caption(nearby)

    is_table = _is_suspected_table(caption, structure, column_hints, nearby_table)
    is_figure = not is_table and (caption.get("type") == "figure_caption_embedded_in_image" or nearby_figure)
    confidence = _classification_confidence(caption, structure, column_hints, nearby_table, ocr)

    if not caption and nearby_table:
        caption = nearby_table
    if not caption and nearby_figure:
        caption = nearby_figure

    return ImageScanResult(
        image_index=image_index,
        paragraph_index=paragraph_index,
        block_index=block_index,
        relationship_id=occurrence.relationship_id,
        file_name=file_name,
        content_type=getattr(part, "content_type", None) if part is not None else None,
        shape_type=occurrence.shape_type,
        width_px=width_px,
        height_px=height_px,
        width_inches=width_inches,
        height_inches=height_inches,
        nearby_paragraphs=nearby,
        ocr_text=normalize_space(ocr.text),
        ocr_engine=ocr.engine,
        ocr_available=ocr.available,
        ocr_error=ocr.error,
        ocr_title_text=caption.get("title_text", "") if caption else _first_ocr_line(ocr.text),
        caption_type=caption.get("type") if caption else None,
        caption_number_token=caption.get("number_token") if caption else None,
        caption_number=caption.get("number") if caption else None,
        caption_text=caption.get("caption_text", "") if caption else "",
        table_structure=structure,
        column_hints=column_hints,
        is_suspected_table=is_table,
        is_suspected_figure=is_figure,
        confidence=confidence,
    )


def _iter_paragraph_image_occurrences(paragraph) -> list[_ImageOccurrence]:
    occurrences: list[_ImageOccurrence] = []
    try:
        blips = paragraph._p.xpath(".//a:blip")
    except Exception:
        blips = []
    for blip in blips:
        relationship_id = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
        width_emu, height_emu = _extent_from_ancestors(blip)
        occurrences.append(
            _ImageOccurrence(
                relationship_id=relationship_id,
                shape_type=_drawing_shape_type(blip),
                width_emu=width_emu,
                height_emu=height_emu,
            )
        )

    try:
        legacy_images = paragraph._p.xpath(".//v:imagedata")
    except Exception:
        legacy_images = []
    for image in legacy_images:
        relationship_id = image.get(qn("r:id")) or image.get(_safe_qn("o:relid"))
        occurrences.append(_ImageOccurrence(relationship_id=relationship_id, shape_type="pict"))
    return occurrences


def _extent_from_ancestors(node) -> tuple[int | None, int | None]:
    for ancestor in node.iterancestors():
        for child in ancestor:
            if _local_name(child) != "extent":
                continue
            cx = _safe_int(child.get("cx"))
            cy = _safe_int(child.get("cy"))
            if cx and cy:
                return cx, cy
    return None, None


def _drawing_shape_type(node) -> str:
    for ancestor in node.iterancestors():
        local = _local_name(ancestor)
        if local == "anchor":
            return "floating"
        if local == "inline":
            return "inline"
    return "drawing"


def _related_image_part(paragraph, relationship_id: str | None):
    if not relationship_id:
        return None
    try:
        return paragraph.part.related_parts.get(relationship_id)
    except Exception:
        return None


def _image_file_name(part, relationship_id: str | None, image_index: int) -> str:
    partname = str(getattr(part, "partname", "") or "")
    if partname:
        return str(PurePosixPath(partname).name)
    if relationship_id:
        return f"{relationship_id}"
    return f"image-{image_index}"


def _image_pixel_size(blob: bytes | None) -> tuple[int | None, int | None]:
    if not blob:
        return None, None
    try:
        from PIL import Image

        image = Image.open(BytesIO(blob))
        return image.size
    except Exception:
        return None, None


def _nearby_paragraphs(document, blocks, block_index: int | None, paragraph_index: int | None) -> list[NearbyParagraph]:
    nearby: list[NearbyParagraph] = []
    if paragraph_index is not None:
        start = max(0, paragraph_index - 3)
        end = min(len(document.paragraphs), paragraph_index + 4)
        for index in range(start, end):
            text = normalize_space(document.paragraphs[index].text)
            if text:
                nearby.append(NearbyParagraph(index, text))
        return nearby

    if block_index is None:
        return nearby
    for block in blocks:
        if abs(block.block_index - block_index) > 3:
            continue
        text = normalize_space(block.text)
        if text:
            nearby.append(NearbyParagraph(block.paragraph_index, text))
    return nearby


def _ocr_image_blob(blob: bytes | None) -> OcrResult:
    if not blob:
        return OcrResult(False, None, error="No image data.")
    for reader in (_ocr_with_paddle, _ocr_with_easyocr, _ocr_with_tesseract):
        result = reader(blob)
        if result.available:
            return result
    return OcrResult(
        False,
        None,
        error="OCR is unavailable. Install PaddleOCR, EasyOCR, or pytesseract with a local Tesseract binary.",
    )


def _ocr_with_paddle(blob: bytes) -> OcrResult:
    try:
        import numpy as np
        from PIL import Image

        reader = _paddle_reader()
        if reader is None:
            return OcrResult(False, "paddleocr", error="PaddleOCR is not installed.")
        image = np.asarray(Image.open(BytesIO(blob)).convert("RGB"))
        output = reader.ocr(image, cls=True)
        lines: list[str] = []
        for page in output or []:
            for item in page or []:
                if len(item) >= 2 and item[1]:
                    lines.append(str(item[1][0]))
        return OcrResult(True, "paddleocr", "\n".join(lines))
    except Exception as exc:
        return OcrResult(False, "paddleocr", error=str(exc))


@lru_cache(maxsize=1)
def _paddle_reader():
    try:
        from paddleocr import PaddleOCR

        return PaddleOCR(use_angle_cls=True, lang="en", show_log=False)
    except Exception:
        return None


def _ocr_with_easyocr(blob: bytes) -> OcrResult:
    try:
        import numpy as np
        from PIL import Image

        reader = _easyocr_reader()
        if reader is None:
            return OcrResult(False, "easyocr", error="EasyOCR is not installed.")
        image = np.asarray(Image.open(BytesIO(blob)).convert("RGB"))
        lines = reader.readtext(image, detail=0, paragraph=True)
        return OcrResult(True, "easyocr", "\n".join(str(line) for line in lines))
    except Exception as exc:
        return OcrResult(False, "easyocr", error=str(exc))


@lru_cache(maxsize=1)
def _easyocr_reader():
    try:
        import easyocr

        return easyocr.Reader(["en"], gpu=False)
    except Exception:
        return None


def _ocr_with_tesseract(blob: bytes) -> OcrResult:
    try:
        import pytesseract
        from PIL import Image

        image = Image.open(BytesIO(blob)).convert("RGB")
        return OcrResult(True, "pytesseract", pytesseract.image_to_string(image))
    except Exception as exc:
        return OcrResult(False, "pytesseract", error=str(exc))


def _detect_table_structure(blob: bytes | None) -> TableStructureResult:
    if not blob:
        return TableStructureResult(False, 0.0, error="No image data.")
    cv2_result = _detect_table_structure_with_cv2(blob)
    if cv2_result.method != "unavailable":
        return cv2_result
    return _detect_table_structure_with_pillow(blob)


def _detect_table_structure_with_cv2(blob: bytes) -> TableStructureResult:
    try:
        import cv2
        import numpy as np

        data = np.frombuffer(blob, dtype=np.uint8)
        gray = cv2.imdecode(data, cv2.IMREAD_GRAYSCALE)
        if gray is None:
            return TableStructureResult(False, 0.0, method="opencv", error="OpenCV could not decode image.")
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        height, width = binary.shape[:2]
        horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(10, width // 25), 1))
        vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(10, height // 25)))
        horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel)
        vertical = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel)
        horizontal_lines = _projection_line_count(horizontal, axis=1)
        vertical_lines = _projection_line_count(vertical, axis=0)
        is_table = horizontal_lines >= 2 and vertical_lines >= 2
        confidence = min(0.95, 0.25 + 0.12 * horizontal_lines + 0.12 * vertical_lines) if is_table else 0.0
        return TableStructureResult(is_table, confidence, horizontal_lines, vertical_lines, "opencv")
    except Exception as exc:
        return TableStructureResult(False, 0.0, method="unavailable", error=str(exc))


def _detect_table_structure_with_pillow(blob: bytes) -> TableStructureResult:
    try:
        import numpy as np
        from PIL import Image

        image = Image.open(BytesIO(blob)).convert("L")
        width, height = image.size
        if width < 30 or height < 30:
            return TableStructureResult(False, 0.0, method="pillow")
        arr = np.asarray(image)
        dark = arr < 150
        horizontal_lines = _grouped_projection_count(dark.mean(axis=1) > 0.42)
        vertical_lines = _grouped_projection_count(dark.mean(axis=0) > 0.42)
        is_table = horizontal_lines >= 2 and vertical_lines >= 2
        confidence = min(0.8, 0.18 + 0.1 * horizontal_lines + 0.1 * vertical_lines) if is_table else 0.0
        return TableStructureResult(is_table, confidence, horizontal_lines, vertical_lines, "pillow")
    except Exception as exc:
        return TableStructureResult(False, 0.0, method="unavailable", error=str(exc))


def _projection_line_count(image, axis: int) -> int:
    try:
        import numpy as np

        projection = image.mean(axis=axis) > 20
        return _grouped_projection_count(np.asarray(projection))
    except Exception:
        return 0


def _grouped_projection_count(mask) -> int:
    count = 0
    in_group = False
    for value in mask:
        if value and not in_group:
            count += 1
            in_group = True
        elif not value:
            in_group = False
    return count


def _caption_from_ocr(text: str) -> dict[str, Any]:
    if not text:
        return {}
    table = TABLE_CAPTION_OCR_RE.search(text)
    if table:
        token = table.group("number")
        return {
            "type": "table_caption_embedded_in_image",
            "title_text": normalize_space(table.group(0)),
            "number_token": token,
            "number": _table_number_to_int(token),
            "caption_text": normalize_space(table.group("title")),
        }
    figure = FIGURE_CAPTION_OCR_RE.search(text)
    if figure:
        token = figure.group("number")
        return {
            "type": "figure_caption_embedded_in_image",
            "title_text": normalize_space(figure.group(0)),
            "number_token": token,
            "number": _safe_int(re.sub(r"\D", "", token)),
            "caption_text": normalize_space(figure.group("title")),
        }
    return {}


def _nearby_table_caption_or_mention(nearby: list[NearbyParagraph]) -> dict[str, Any]:
    for item in nearby:
        match = match_table_number(item.text)
        if match:
            token = match.group(2)
            return {
                "type": "table_caption_near_image",
                "title_text": normalize_space(item.text),
                "number_token": token,
                "number": _table_number_to_int(token),
                "caption_text": normalize_space(match.group(3)) if len(match.groups()) >= 3 else "",
            }
    for item in nearby:
        match = TABLE_MENTION_OCR_RE.search(item.text)
        if match:
            token = match.group("number") or match.group("zh_number")
            return {
                "type": "table_mention_near_image",
                "title_text": normalize_space(item.text),
                "number_token": token,
                "number": _table_number_to_int(token),
                "caption_text": "",
            }
    return {}


def _nearby_figure_caption(nearby: list[NearbyParagraph]) -> dict[str, Any]:
    for item in nearby:
        match = match_figure_caption(item.text)
        if match:
            return {
                "type": "figure_caption_near_image",
                "title_text": normalize_space(item.text),
                "number_token": match.group(2),
                "number": _safe_int(re.sub(r"\D", "", match.group(2))),
                "caption_text": normalize_space(match.group(3)),
            }
    return {}


def _column_hints(text: str) -> list[str]:
    seen: set[str] = set()
    hints: list[str] = []
    for match in TABLE_COLUMN_HINT_RE.finditer(text or ""):
        value = normalize_space(match.group(0)).lower()
        if value in seen:
            continue
        seen.add(value)
        hints.append(value)
    return hints


def _is_suspected_table(
    caption: dict[str, Any],
    structure: TableStructureResult,
    column_hints: list[str],
    nearby_table: dict[str, Any],
) -> bool:
    if caption.get("type") == "table_caption_embedded_in_image":
        return True
    if structure.is_table:
        return True
    if len(column_hints) >= 3:
        return True
    return bool(nearby_table)


def _classification_confidence(
    caption: dict[str, Any],
    structure: TableStructureResult,
    column_hints: list[str],
    nearby_table: dict[str, Any],
    ocr: OcrResult,
) -> float:
    if caption.get("type") == "table_caption_embedded_in_image":
        return 0.95
    if structure.is_table and len(column_hints) >= 2:
        return max(0.88, structure.confidence)
    if structure.is_table:
        return max(0.75, structure.confidence)
    if len(column_hints) >= 3:
        return 0.72
    if nearby_table:
        return 0.65 if ocr.available else 0.45
    if not ocr.available:
        return 0.2
    return 0.35


def _first_ocr_line(text: str) -> str:
    for line in (text or "").splitlines():
        line = normalize_space(line)
        if line:
            return line
    return ""


def _table_number_to_int(token: str | None) -> int | None:
    if not token:
        return None
    if token.isdigit():
        return int(token)
    return roman_to_int(token)


def _emu_to_inches(value: int | None) -> float | None:
    if not value:
        return None
    return round(value / EMU_PER_INCH, 2)


def _safe_int(value: str | None) -> int | None:
    try:
        return int(value) if value not in (None, "") else None
    except Exception:
        return None


def _safe_qn(tag: str) -> str:
    try:
        return qn(tag)
    except Exception:
        return tag


def _local_name(node) -> str:
    try:
        return etree.QName(node).localname
    except Exception:
        return ""
