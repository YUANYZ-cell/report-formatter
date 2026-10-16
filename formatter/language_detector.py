from __future__ import annotations

import re

from .utils import normalize_space


CHINESE_CHAR_RE = re.compile(r"[\u4e00-\u9fff]")
ENGLISH_WORD_RE = re.compile(r"\b[A-Za-z]+\b")


def detect_report_language(document) -> dict:
    text = _document_text(document)
    chinese_char_count = len(CHINESE_CHAR_RE.findall(text))
    english_word_count = len(ENGLISH_WORD_RE.findall(text))
    total = chinese_char_count + english_word_count

    if total < 20:
        language = "unknown"
        confidence = 0.0
    else:
        chinese_ratio = chinese_char_count / total
        english_ratio = english_word_count / total
        if chinese_ratio >= 0.62:
            language = "zh"
            confidence = round(chinese_ratio, 2)
        elif english_ratio >= 0.62:
            language = "en"
            confidence = round(english_ratio, 2)
        else:
            language = "unknown"
            confidence = round(max(chinese_ratio, english_ratio), 2)

    return {
        "language": language,
        "confidence": confidence,
        "chinese_char_count": chinese_char_count,
        "english_word_count": english_word_count,
    }


def resolve_report_language(detection: dict, override: str) -> str:
    if override in {"en", "zh"}:
        return override
    detected = detection.get("language", "unknown")
    return detected if detected in {"en", "zh"} else "en"


def _document_text(document) -> str:
    chunks: list[str] = []
    for paragraph in document.paragraphs:
        text = normalize_space(paragraph.text)
        if text and not _looks_like_noise(text):
            chunks.append(text)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                text = normalize_space(cell.text)
                if text and not _looks_like_noise(text):
                    chunks.append(text)
    return " ".join(chunks)


def _looks_like_noise(text: str) -> bool:
    if re.fullmatch(r"[\d\s,.;:()\[\]{}+\-*/=]+", text):
        return True
    if re.fullmatch(r"\[[0-9]+\].*", text):
        return True
    return False
