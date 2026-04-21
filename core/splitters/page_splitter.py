from __future__ import annotations

from core.parsers.pdf_parser import parse_pdf_by_page
from core.parsers.word_parser import parse_word_by_page
from core.parsers.excel_parser import parse_excel_by_sheet


def split_by_page(file_bytes: bytes, file_ext: str) -> list[dict]:
    """
    Split a document by page/sheet.
    Returns list of {"content": str, "metadata": dict}.
    """
    ext = file_ext.lower().lstrip(".")

    if ext == "pdf":
        return parse_pdf_by_page(file_bytes)
    elif ext == "docx":
        return parse_word_by_page(file_bytes)
    elif ext in ("xlsx", "xls"):
        return parse_excel_by_sheet(file_bytes)
    else:
        raise ValueError(f"Unsupported file type: .{ext}")
