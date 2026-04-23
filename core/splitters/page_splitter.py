from __future__ import annotations

import fitz

from core.parsers.pdf_parser import parse_pdf_by_page
from core.parsers.word_parser import parse_word_by_page
from core.parsers.excel_parser import parse_excel_by_sheet


def _safe_int(s: str) -> int | None:
    s = s.strip()
    if not s:
        return None
    try:
        return int(s)
    except ValueError:
        return None


def _parse_page_range_bounds(page_range: str) -> tuple[int | None, int | None]:
    """
    Parse a 1-based inclusive page range into (start_1based, end_1based).
    None on either side means open-ended. Empty string = (None, None).
    """
    if not page_range or not str(page_range).strip():
        return None, None
    s = str(page_range).strip()
    if "-" not in s:
        n = _safe_int(s)
        return (n, n) if n is not None else (None, None)
    left, _, right = s.partition("-")
    start = _safe_int(left)
    end = _safe_int(right)
    if start is None and end is None:
        raise ValueError(
            f"Invalid page_range: {page_range!r}. Expected '1-10' / '5-' / '-3' / '7' or empty."
        )
    return start, end


def _bounds_to_indices(total: int, start_1b: int | None, end_1b: int | None) -> list[int]:
    start = (start_1b - 1) if start_1b else 0
    end = end_1b if end_1b else total
    start = max(0, start)
    end = min(total, max(start, end))
    return list(range(start, end))


def _group_pages(selected: list[dict], pages_per_chunk: int) -> list[dict]:
    step = max(1, int(pages_per_chunk or 1))
    if step == 1:
        return selected

    merged = []
    for i in range(0, len(selected), step):
        group = selected[i:i + step]
        if not group:
            continue
        parts = []
        all_images = []
        image_ids = []
        first_meta = dict(group[0].get("metadata", {}))
        last_meta = group[-1].get("metadata", {})
        for p in group:
            parts.append(p.get("content", ""))
            meta = p.get("metadata", {})
            if "images" in meta:
                all_images.extend(meta["images"])
            if "image_ids" in meta:
                image_ids.extend(meta["image_ids"])
        merged_meta = first_meta
        if "page" in first_meta and "page" in last_meta:
            merged_meta["page_start"] = first_meta["page"]
            merged_meta["page_end"] = last_meta["page"]
            merged_meta.pop("page", None)
        if all_images:
            merged_meta["images"] = all_images
        if image_ids:
            merged_meta["image_ids"] = image_ids
        merged.append({
            "content": "\n\n".join(parts),
            "metadata": merged_meta,
        })
    return merged


def split_by_page(
    file_bytes: bytes,
    file_ext: str,
    page_range: str = "",
    pages_per_chunk: int = 1,
) -> list[dict]:
    """
    Split a document by page/sheet.
    Returns list of {"content": str, "metadata": dict}.

    page_range: e.g. "1-10", "5-", "-3", "7" (1-based inclusive). Empty = all.
    pages_per_chunk: group every N pages into one chunk. Default 1 (per page).
    """
    ext = file_ext.lower().lstrip(".")
    start_1b, end_1b = _parse_page_range_bounds(page_range)

    if ext == "pdf":
        with fitz.open(stream=file_bytes, filetype="pdf") as doc:
            total = doc.page_count
        indices = _bounds_to_indices(total, start_1b, end_1b)
        pages = parse_pdf_by_page(file_bytes, page_indices=set(indices))
    elif ext == "docx":
        if end_1b is not None or start_1b is not None:
            upper = end_1b if end_1b is not None else 10_000
            indices = set(range(max(0, (start_1b or 1) - 1), upper))
        else:
            indices = None
        pages = parse_word_by_page(file_bytes, page_indices=indices)
    elif ext in ("xlsx", "xls"):
        all_sheets = parse_excel_by_sheet(file_bytes)
        indices = _bounds_to_indices(len(all_sheets), start_1b, end_1b)
        pages = [all_sheets[i] for i in indices]
    else:
        raise ValueError(f"Unsupported file type: .{ext}")

    return _group_pages(pages, pages_per_chunk)
