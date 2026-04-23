from __future__ import annotations

import base64
import io
from typing import Iterator

from docx import Document
from docx.oxml.ns import qn

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
C_NS = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
MC_NS = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"


def _is_page_break(paragraph) -> bool:
    """Check if a paragraph contains a manual page break."""
    for run in paragraph.runs:
        for child in run._element:
            if child.tag == qn("w:br"):
                if child.get(qn("w:type")) == "page":
                    return True
    return False


def _extract_images_from_element(element, doc_part, image_counter: list[int]) -> list[dict]:
    """Walk an XML element and extract embedded images (pictures + chart fallbacks)."""
    images: list[dict] = []

    for blip in element.iter(f"{A_NS}blip"):
        r_embed = blip.get(f"{R_NS}embed") or blip.get(f"{R_NS}link")
        if not r_embed:
            continue
        try:
            image_part = doc_part.related_parts[r_embed]
        except KeyError:
            continue

        image_bytes = getattr(image_part, "blob", None) or getattr(image_part, "_blob", None)
        if not image_bytes:
            continue

        mime_type = getattr(image_part, "content_type", "image/png")
        kind = "image"
        parent = blip.getparent()
        while parent is not None:
            if parent.tag == f"{MC_NS}Fallback":
                kind = "chart"
                break
            parent = parent.getparent()

        image_counter[0] += 1
        images.append({
            "id": f"image_{image_counter[0]}",
            "mime_type": mime_type,
            "base64": base64.b64encode(image_bytes).decode("ascii"),
            "kind": kind,
        })

    for chart_ref in element.iter(f"{C_NS}chart"):
        if chart_ref.get(f"{R_NS}id"):
            image_counter[0] += 1
            images.append({
                "id": f"image_{image_counter[0]}",
                "mime_type": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                "base64": "",
                "kind": "chart_xml",
                "note": "Chart without fallback image; only chart XML reference preserved.",
            })

    return images


def _paragraph_to_text_with_images(paragraph, doc_part, image_counter: list[int]) -> tuple[str, list[dict]]:
    """Convert a paragraph to text, replacing image anchors with [IMAGE:id] placeholders."""
    text_parts: list[str] = []
    images: list[dict] = []

    for child in paragraph._element.iter():
        if child.tag == f"{W_NS}t":
            if child.text:
                text_parts.append(child.text)
        elif child.tag == f"{W_NS}tab":
            text_parts.append("\t")
        elif child.tag == f"{W_NS}br":
            text_parts.append("\n")
        elif child.tag == f"{W_NS}drawing":
            drawing_imgs = _extract_images_from_element(child, doc_part, image_counter)
            seen_ids: set[str] = set()
            for img in drawing_imgs:
                if img["id"] in seen_ids:
                    continue
                seen_ids.add(img["id"])
                images.append(img)
                text_parts.append(f"[IMAGE:{img['id']}]")

    return "".join(text_parts), images


def _table_to_markdown(table) -> str:
    rows = []
    for row in table.rows:
        cells = [cell.text.replace("\n", " ").replace("|", "\\|").strip() for cell in row.cells]
        rows.append(cells)

    if not rows:
        return ""

    max_cols = max(len(r) for r in rows)
    rows = [r + [""] * (max_cols - len(r)) for r in rows]

    lines = ["| " + " | ".join(rows[0]) + " |"]
    lines.append("| " + " | ".join("---" for _ in range(max_cols)) + " |")
    for r in rows[1:]:
        lines.append("| " + " | ".join(r) + " |")
    return "\n".join(lines)


def _iter_block_items(doc) -> Iterator:
    """Iterate paragraphs and tables in document order."""
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def parse_word_by_page(file_bytes: bytes, page_indices: set[int] | None = None) -> list[dict]:
    """
    Split Word document by manual page breaks. Extracts text, tables (as markdown),
    and images (as base64 in metadata).
    page_indices: optional set of 0-based page indices to keep. Expensive work is
    skipped for pages outside the range to speed up large docs.
    """
    doc = Document(io.BytesIO(file_bytes))
    doc_part = doc.part
    image_counter = [0]

    pages: list[dict] = [{"lines": [], "images": []}]
    current_page = 0
    max_needed = max(page_indices) if page_indices else None

    def in_range() -> bool:
        return page_indices is None or current_page in page_indices

    for block in _iter_block_items(doc):
        if block.__class__.__name__ == "Paragraph":
            if _is_page_break(block):
                pages.append({"lines": [], "images": []})
                current_page += 1
                if max_needed is not None and current_page > max_needed:
                    break
            if in_range():
                text, imgs = _paragraph_to_text_with_images(block, doc_part, image_counter)
                text = text.strip()
                if text:
                    pages[-1]["lines"].append(text)
                pages[-1]["images"].extend(imgs)
        else:
            if in_range():
                md = _table_to_markdown(block)
                if md:
                    pages[-1]["lines"].append(md)
                tbl_imgs = _extract_images_from_element(block._element, doc_part, image_counter)
                seen: set[str] = set()
                for img in tbl_imgs:
                    if img["id"] in seen:
                        continue
                    seen.add(img["id"])
                    pages[-1]["images"].append(img)

    chunks = []
    for i, page in enumerate(pages):
        content = "\n\n".join(page["lines"]).strip()
        if content or page["images"]:
            chunks.append({
                "content": content,
                "metadata": {
                    "page": i + 1,
                    "images": page["images"],
                },
            })

    return chunks
