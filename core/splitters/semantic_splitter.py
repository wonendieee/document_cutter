from __future__ import annotations

from core.parsers.pdf_parser import parse_pdf_full_text
from core.parsers.word_parser import parse_word_by_heading
from core.parsers.excel_parser import parse_excel_by_sheet_with_row_split


def _merge_paragraphs_into_chunks(
    paragraphs: list[dict],
    max_chunk_size: int,
    overlap_size: int,
    heading_key: str | None = None,
    meta_keys: list[str] | None = None,
) -> list[dict]:
    """
    Merge a list of paragraphs into chunks respecting max_chunk_size.
    Each paragraph dict must have a 'text' key.
    Overlap is applied by repeating trailing text from the previous chunk.
    """
    if not paragraphs:
        return []

    chunks: list[dict] = []
    current_texts: list[str] = []
    current_len = 0
    current_meta: dict = {}

    def _build_meta(para: dict) -> dict:
        meta = {}
        if heading_key and heading_key in para:
            meta["heading"] = para[heading_key]
        if meta_keys:
            for k in meta_keys:
                if k in para:
                    meta[k] = para[k]
        return meta

    def _flush():
        if not current_texts:
            return
        content = "\n".join(current_texts)
        chunks.append({"content": content, "metadata": {**current_meta}})

    for para in paragraphs:
        text = para["text"]
        text_len = len(text)

        if current_len + text_len > max_chunk_size and current_texts:
            _flush()
            overlap_text = ""
            if overlap_size > 0 and chunks:
                last_content = chunks[-1]["content"]
                overlap_text = last_content[-overlap_size:]
            current_texts = [overlap_text] if overlap_text else []
            current_len = len(overlap_text)
            current_meta = _build_meta(para)

        current_texts.append(text)
        current_len += text_len
        if not current_meta:
            current_meta = _build_meta(para)

    _flush()
    return chunks


def split_semantic(
    file_bytes: bytes,
    file_ext: str,
    max_chunk_size: int = 2000,
    overlap_size: int = 200,
) -> list[dict]:
    """
    Split a document by semantic structure.
    Returns list of {"content": str, "metadata": dict}.
    """
    ext = file_ext.lower().lstrip(".")

    if ext == "pdf":
        _full_text, paragraphs = parse_pdf_full_text(file_bytes)
        return _merge_paragraphs_into_chunks(
            paragraphs, max_chunk_size, overlap_size, meta_keys=["page"]
        )

    elif ext == "docx":
        sections = parse_word_by_heading(file_bytes)
        chunks = []
        for sec in sections:
            text = sec["text"]
            heading = sec.get("heading", "")
            images = sec.get("images", [])
            if len(text) <= max_chunk_size:
                chunks.append({
                    "content": text,
                    "metadata": {
                        "heading": heading,
                        "char_count": len(text),
                        "images": images,
                    },
                })
            else:
                sub_paras = [{"text": p.strip()} for p in text.split("\n") if p.strip()]
                sub_chunks = _merge_paragraphs_into_chunks(
                    sub_paras, max_chunk_size, overlap_size
                )
                for idx, sc in enumerate(sub_chunks):
                    sc["metadata"]["heading"] = heading
                    sc["metadata"]["images"] = images if idx == 0 else []
                chunks.extend(sub_chunks)
        return chunks

    elif ext in ("xlsx", "xls"):
        max_rows = max(max_chunk_size // 80, 10)
        return parse_excel_by_sheet_with_row_split(file_bytes, max_rows=max_rows)

    else:
        raise ValueError(f"Unsupported file type: .{ext}")
