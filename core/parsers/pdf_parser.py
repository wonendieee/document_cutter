from __future__ import annotations

import base64

import fitz

EXT_MIME_MAP = {
    "png": "image/png",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tiff": "image/tiff",
    "webp": "image/webp",
}


def _extract_page_images(doc, page, image_counter: list[int]) -> list[dict]:
    """Extract all embedded images from a page as base64 entries."""
    images: list[dict] = []
    seen_xrefs: set[int] = set()

    for img_info in page.get_images(full=True):
        xref = img_info[0]
        if xref in seen_xrefs:
            continue
        seen_xrefs.add(xref)

        try:
            extracted = doc.extract_image(xref)
        except Exception:
            continue

        image_bytes = extracted.get("image")
        ext = (extracted.get("ext") or "png").lower()
        if not image_bytes:
            continue

        image_counter[0] += 1
        images.append({
            "id": f"image_{image_counter[0]}",
            "mime_type": EXT_MIME_MAP.get(ext, f"image/{ext}"),
            "base64": base64.b64encode(image_bytes).decode("ascii"),
            "kind": "image",
        })

    return images


def parse_pdf_by_page(file_bytes: bytes, page_indices: set[int] | None = None) -> list[dict]:
    """
    Extract text and images from each page of a PDF, returning one chunk per page.
    page_indices: optional set of 0-based page indices to keep. None = all pages.
    """
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    image_counter = [0]
    chunks = []

    for i in range(doc.page_count):
        if page_indices is not None and i not in page_indices:
            continue
        page = doc.load_page(i)
        text = page.get_text("text").strip()
        images = _extract_page_images(doc, page, image_counter)

        parts = []
        if text:
            parts.append(text)
        for img in images:
            parts.append(f"[IMAGE:{img['id']}]")

        content = "\n\n".join(parts)
        if content or images:
            chunks.append({
                "content": content,
                "metadata": {
                    "page": i + 1,
                    "images": images,
                },
            })

    doc.close()
    return chunks


