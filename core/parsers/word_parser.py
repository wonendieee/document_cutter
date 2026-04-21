from __future__ import annotations

import base64
import io
import re
import unicodedata
from typing import Iterator

from docx import Document
from docx.oxml.ns import qn

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PIC_NS = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
C_NS = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
MC_NS = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"


def _is_page_break(paragraph) -> bool:
    """Check if a paragraph contains a page break."""
    for run in paragraph.runs:
        for child in run._element:
            if child.tag == qn("w:br"):
                if child.get(qn("w:type")) == "page":
                    return True
    return False


def _get_heading_level(paragraph) -> int | None:
    """Return heading level (1-9) or None if not a heading."""
    style_name = paragraph.style.name or ""
    if style_name.startswith("Heading"):
        try:
            return int(style_name.replace("Heading", "").strip())
        except ValueError:
            return None
    return None


def _extract_images_from_element(element, doc_part, image_counter: list[int]) -> list[dict]:
    """
    Walk an XML element and extract all embedded images (pictures, chart fallbacks, etc.).
    Returns list of {"id", "mime_type", "base64", "kind"}.
    """
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

        in_chart_fallback = False
        parent = blip.getparent()
        while parent is not None:
            if parent.tag == f"{MC_NS}Fallback":
                in_chart_fallback = True
                break
            parent = parent.getparent()
        if in_chart_fallback:
            kind = "chart"

        image_counter[0] += 1
        images.append({
            "id": f"image_{image_counter[0]}",
            "mime_type": mime_type,
            "base64": base64.b64encode(image_bytes).decode("ascii"),
            "kind": kind,
        })

    for chart_ref in element.iter(f"{C_NS}chart"):
        r_id = chart_ref.get(f"{R_NS}id")
        if not r_id:
            continue
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
    """
    Convert a paragraph to text, replacing image anchors with [IMAGE:id] placeholders.
    Returns (text, images).
    """
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
    """Convert a docx table to a Markdown table."""
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
    """
    Iterate paragraphs and tables in document order (python-docx's doc.paragraphs
    skips tables, so we need to walk body directly).
    """
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


def parse_word_by_page(file_bytes: bytes) -> list[dict]:
    """
    Split Word document by page breaks. Extracts text, tables (as markdown),
    and images (as base64 in metadata).
    """
    doc = Document(io.BytesIO(file_bytes))
    doc_part = doc.part
    image_counter = [0]

    pages: list[dict] = [{"lines": [], "images": []}]

    for block in _iter_block_items(doc):
        if block.__class__.__name__ == "Paragraph":
            if _is_page_break(block):
                pages.append({"lines": [], "images": []})
            text, imgs = _paragraph_to_text_with_images(block, doc_part, image_counter)
            text = text.strip()
            if text:
                pages[-1]["lines"].append(text)
            pages[-1]["images"].extend(imgs)
        else:
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

    if len(chunks) == 1:
        chunks[0]["metadata"]["page"] = 1

    return chunks


_PUNCT_RE = re.compile(r"[\s\u3000\.\,\:\;\!\?\-\_\(\)\[\]\{\}\'\"`~\*\#\/\\。，：；！？（）【】《》、]+")


def _normalize(text: str) -> str:
    """Normalize text for anchor matching: NFKC, lowercase, strip punctuation/whitespace."""
    if not text:
        return ""
    t = unicodedata.normalize("NFKC", text).lower()
    t = _PUNCT_RE.sub("", t)
    return t


def parse_word_by_anchors(file_bytes: bytes, anchors: list[dict]) -> list[dict]:
    """
    Split Word document by upstream-provided heading anchors.

    anchors: list of dicts with keys:
      - section_standard (str): semantic label for the chunk
      - anchor_primary (str): primary heading text to match
      - anchor_fallbacks (list[str]): alternative heading texts
      - anchor_pos (int, optional): paragraph index hint (1-based, counts non-empty paragraphs)
      - found (bool): only process anchors where found=True

    Matching: normalized (NFKC + lowercased + stripped punctuation/spaces).
    Falls back to anchor_pos (N-th non-empty paragraph) when text matching fails.
    Content before first anchor becomes a preamble chunk with section_standard="_preamble".
    """
    doc = Document(io.BytesIO(file_bytes))
    doc_part = doc.part
    image_counter = [0]

    active_anchors = [a for a in anchors if a.get("found")]

    blocks: list[dict] = []
    nonempty_para_idx = 0

    for block in _iter_block_items(doc):
        if block.__class__.__name__ == "Paragraph":
            text, imgs = _paragraph_to_text_with_images(block, doc_part, image_counter)
            text = text.strip()
            entry = {
                "type": "paragraph",
                "text": text,
                "images": imgs,
                "para_idx": None,
            }
            if text:
                nonempty_para_idx += 1
                entry["para_idx"] = nonempty_para_idx
            blocks.append(entry)
        else:
            md = _table_to_markdown(block)
            tbl_imgs = _extract_images_from_element(block._element, doc_part, image_counter)
            seen: set[str] = set()
            dedup_imgs = []
            for img in tbl_imgs:
                if img["id"] in seen:
                    continue
                seen.add(img["id"])
                dedup_imgs.append(img)
            blocks.append({
                "type": "table",
                "text": md,
                "images": dedup_imgs,
                "para_idx": None,
            })

    cut_points: list[tuple[int, dict]] = []
    used_block_indices: set[int] = set()
    last_cut_block = -1

    for anchor in active_anchors:
        candidates = [anchor.get("anchor_primary", "")] + list(anchor.get("anchor_fallbacks") or [])
        candidates = [c for c in candidates if c]
        normalized_targets = [(_normalize(c), c) for c in candidates]

        matched_block_idx = None
        matched_anchor_text = ""
        match_type = ""

        for tgt_norm, tgt_raw in normalized_targets:
            if not tgt_norm:
                continue
            for b_idx in range(last_cut_block + 1, len(blocks)):
                if b_idx in used_block_indices:
                    continue
                block = blocks[b_idx]
                if block["type"] != "paragraph" or not block["text"]:
                    continue
                if _normalize(block["text"]) == tgt_norm:
                    matched_block_idx = b_idx
                    matched_anchor_text = tgt_raw
                    match_type = "primary" if tgt_raw == anchor.get("anchor_primary") else "fallback"
                    break
            if matched_block_idx is not None:
                break

        if matched_block_idx is None:
            pos = anchor.get("anchor_pos", -1)
            if isinstance(pos, int) and pos > 0:
                for b_idx in range(last_cut_block + 1, len(blocks)):
                    if b_idx in used_block_indices:
                        continue
                    block = blocks[b_idx]
                    if block["type"] == "paragraph" and block.get("para_idx") == pos:
                        matched_block_idx = b_idx
                        matched_anchor_text = blocks[b_idx]["text"]
                        match_type = "position_fallback"
                        break

        if matched_block_idx is not None:
            used_block_indices.add(matched_block_idx)
            last_cut_block = matched_block_idx
            cut_points.append((matched_block_idx, {
                "section_standard": anchor.get("section_standard", ""),
                "matched_anchor": matched_anchor_text,
                "match_type": match_type,
                "confidence": anchor.get("confidence", 0),
            }))

    chunks: list[dict] = []

    def _build_chunk(start: int, end: int, meta: dict) -> dict | None:
        lines: list[str] = []
        imgs: list[dict] = []
        seen_ids: set[str] = set()
        for b in blocks[start:end]:
            if b["text"]:
                lines.append(b["text"])
            for img in b["images"]:
                if img["id"] in seen_ids:
                    continue
                seen_ids.add(img["id"])
                imgs.append(img)
        content = "\n\n".join(lines).strip()
        if not content and not imgs:
            return None
        return {
            "content": content,
            "metadata": {**meta, "images": imgs},
        }

    if cut_points:
        first_cut = cut_points[0][0]
        preamble = _build_chunk(0, first_cut, {
            "section_standard": "_preamble",
            "matched_anchor": "",
            "match_type": "preamble",
        })
        if preamble:
            chunks.append(preamble)

        for i, (start_idx, meta) in enumerate(cut_points):
            end_idx = cut_points[i + 1][0] if i + 1 < len(cut_points) else len(blocks)
            chunk = _build_chunk(start_idx, end_idx, meta)
            if chunk:
                chunks.append(chunk)
    else:
        whole = _build_chunk(0, len(blocks), {
            "section_standard": "_full",
            "matched_anchor": "",
            "match_type": "no_anchor_matched",
        })
        if whole:
            chunks.append(whole)

    return chunks


def parse_word_by_heading(file_bytes: bytes) -> list[dict]:
    """
    Split Word document by heading structure. Each section contains text,
    tables, and images (as base64 in metadata).
    """
    doc = Document(io.BytesIO(file_bytes))
    doc_part = doc.part
    image_counter = [0]

    sections: list[dict] = []
    current = {"heading": "", "lines": [], "images": []}

    def _flush():
        text = "\n\n".join(current["lines"]).strip()
        if text or current["images"]:
            sections.append({
                "text": text,
                "heading": current["heading"],
                "images": list(current["images"]),
            })

    for block in _iter_block_items(doc):
        if block.__class__.__name__ == "Paragraph":
            level = _get_heading_level(block)
            text, imgs = _paragraph_to_text_with_images(block, doc_part, image_counter)
            text = text.strip()

            if level is not None and text:
                _flush()
                current = {"heading": text, "lines": [text], "images": list(imgs)}
            else:
                if text:
                    current["lines"].append(text)
                current["images"].extend(imgs)
        else:
            md = _table_to_markdown(block)
            if md:
                current["lines"].append(md)
            tbl_imgs = _extract_images_from_element(block._element, doc_part, image_counter)
            seen: set[str] = set()
            for img in tbl_imgs:
                if img["id"] in seen:
                    continue
                seen.add(img["id"])
                current["images"].append(img)

    _flush()
    return sections
