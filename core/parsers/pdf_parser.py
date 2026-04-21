from __future__ import annotations

import fitz


def parse_pdf_by_page(file_bytes: bytes) -> list[dict]:
    """Extract text from each page of a PDF, returning one chunk per page."""
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    chunks = []
    for i, page in enumerate(doc):
        text = page.get_text("text").strip()
        if text:
            chunks.append({
                "content": text,
                "metadata": {"page": i + 1},
            })
    doc.close()
    return chunks


def parse_pdf_full_text(file_bytes: bytes) -> tuple[str, list[dict]]:
    """Extract full text from PDF with page boundary markers for semantic splitting."""
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    paragraphs: list[dict] = []
    for i, page in enumerate(doc):
        text = page.get_text("text").strip()
        if not text:
            continue
        for para in text.split("\n\n"):
            para = para.strip()
            if para:
                paragraphs.append({
                    "text": para,
                    "page": i + 1,
                })
    doc.close()
    full_text = "\n\n".join(p["text"] for p in paragraphs)
    return full_text, paragraphs
