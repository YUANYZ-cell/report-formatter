from __future__ import annotations

from pathlib import Path
from typing import BinaryIO

from docx import Document


def load_document(path_or_stream: str | Path | BinaryIO):
    return Document(path_or_stream)


def save_document(document, output_path: str | Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path
