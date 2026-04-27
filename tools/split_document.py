from __future__ import annotations

import base64
import os
import re
from collections.abc import Generator
from typing import Any

import requests
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from core.extractors.file_extractor import extract_file
from core.splitters.page_splitter import split_by_page

SUPPORTED_EXTENSIONS = {"pdf", "docx", "xlsx", "xls"}

INTERNAL_FILES_URL_CANDIDATES = [
    os.environ.get("INTERNAL_FILES_URL", "").rstrip("/"),
    os.environ.get("FILES_URL", "").rstrip("/"),
    "http://api:5001",
    "http://host.docker.internal:5001",
]

EXT_TO_MIME = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xls": "application/vnd.ms-excel",
}


def _fetch_by_url_fallback(url: str) -> bytes:
    """If url is relative (no scheme), try prefixing internal Dify API hosts."""
    last_err: Exception | None = None
    candidates = [u for u in INTERNAL_FILES_URL_CANDIDATES if u] or ["http://api:5001"]
    seen = set()
    for base in candidates:
        if base in seen:
            continue
        seen.add(base)
        full = base + url if url.startswith("/") else f"{base}/{url}"
        try:
            resp = requests.get(full, timeout=30)
            resp.raise_for_status()
            return resp.content
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(
        f"Could not fetch file from any internal URL candidate for path '{url}'. Last error: {last_err}"
    )


def _extract_file_bytes(file_info) -> tuple[str, bytes]:
    """Resolve filename and byte content from a Dify File-like object or dict."""
    if isinstance(file_info, (list, tuple)):
        if not file_info:
            raise ValueError("File list is empty.")
        file_info = file_info[0]

    # Probe Dify File pydantic model by class definition, NOT by hasattr on the instance.
    # hasattr(instance, 'blob') would trigger the .blob property which fires httpx.get and can
    # raise ValueError (not AttributeError). hasattr propagates that out, bypassing fallback.
    is_file_like = (
        not isinstance(file_info, dict)
        and hasattr(file_info.__class__, "blob")
        and hasattr(file_info, "url")
    )
    if is_file_like:
        file_name = (
            getattr(file_info, "filename", None)
            or getattr(file_info, "name", None)
            or "unknown"
        )
        # Prefer URL-based fetch when URL is relative — skip the broken SDK .blob entirely.
        url_pre = ""
        try:
            url_pre = getattr(file_info, "url", "") or ""
        except Exception:
            pass
        if url_pre and not url_pre.startswith(("http://", "https://")):
            return file_name, _fetch_by_url_fallback(url_pre)
        try:
            return file_name, file_info.blob
        except Exception as blob_err:
            # 1) try .url attribute (may be None when FILES_URL is unset)
            url = ""
            try:
                url = getattr(file_info, "url", "") or ""
            except Exception:
                pass
            # 2) if .url is empty/absolute, try to extract relative path from
            #    the SDK error message, e.g. "Invalid file URL '/files/UUID/...'"
            if not url or url.startswith(("http://", "https://")):
                import re as _re
                m = _re.search(r"'(/files/[^']+)'", str(blob_err))
                if m:
                    url = m.group(1)
            # 3) last resort: use related_id to build minimal path
            if not url or url.startswith(("http://", "https://")):
                related_id = (
                    getattr(file_info, "related_id", None)
                    or getattr(file_info, "upload_file_id", None)
                )
                if related_id:
                    url = f"/files/{related_id}/file-preview"
            if url and not url.startswith(("http://", "https://")):
                return file_name, _fetch_by_url_fallback(url)
            raise

    if isinstance(file_info, dict):
        file_name = file_info.get("filename") or file_info.get("name") or "unknown"
        file_bytes = file_info.get("blob") or file_info.get("content") or b""
        if isinstance(file_bytes, str):
            file_bytes = base64.b64decode(file_bytes)
        if not file_bytes:
            url = file_info.get("url") or ""
            if url and not url.startswith(("http://", "https://")):
                file_bytes = _fetch_by_url_fallback(url)
            elif url:
                resp = requests.get(url, timeout=30)
                resp.raise_for_status()
                file_bytes = resp.content
        return file_name, file_bytes

    if hasattr(file_info, "read"):
        return (
            getattr(file_info, "filename", None) or getattr(file_info, "name", "unknown"),
            file_info.read(),
        )

    raise TypeError(
        f"Unsupported file object type: {type(file_info).__name__}. "
        "Expected Dify File object, dict, or list of File."
    )


def _safe_int(s: str) -> int | None:
    s = s.strip()
    if not s:
        return None
    try:
        return int(s)
    except ValueError:
        return None


def _parse_bounds(page_range: str) -> tuple[int | None, int | None]:
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


def _sanitize_basename(name: str, default: str = "document") -> str:
    name = str(name or "").strip()
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", name)
    name = name.strip(" ._") or default
    return name[:120]


def _build_output_filename(
    original: str, start: int | None, end: int | None, custom: str | None = None
) -> str:
    base, ext = os.path.splitext(original)
    if custom:
        base = _sanitize_basename(custom, default=base or "document")
        tag = ""
    else:
        base = _sanitize_basename(base, default="document")
        if start and end and start == end:
            tag = f"_p{start}"
        elif start and end:
            tag = f"_p{start}-{end}"
        elif start:
            tag = f"_p{start}-end"
        elif end:
            tag = f"_p1-{end}"
        else:
            tag = "_all"
    return f"{base}{tag}{ext}"


class SplitDocumentTool(Tool):
    def _invoke(
        self, tool_parameters: dict[str, Any]
    ) -> Generator[ToolInvokeMessage]:
        file_info = tool_parameters.get("file") or tool_parameters.get("files")
        if not file_info:
            yield self.create_text_message("Error: no file provided.")
            return

        try:
            file_name, file_bytes = _extract_file_bytes(file_info)
        except Exception as e:
            yield self.create_text_message(f"Error reading file: {e}")
            return

        if not file_bytes:
            yield self.create_text_message("Error: file content is empty.")
            return

        ext = os.path.splitext(file_name)[1].lower().lstrip(".")
        if ext not in SUPPORTED_EXTENSIONS:
            yield self.create_text_message(
                f"Error: unsupported file type '.{ext}'. Supported: {', '.join(SUPPORTED_EXTENSIONS)}"
            )
            return

        split_mode = tool_parameters.get("split_mode") or "page_file"
        page_range = str(tool_parameters.get("page_range") or "")
        pages_per_chunk = int(tool_parameters.get("pages_per_chunk") or 1)
        custom_name = str(tool_parameters.get("output_filename") or "").strip()

        try:
            if split_mode == "page_file":
                start_1b, end_1b = _parse_bounds(page_range)
                out_bytes, mime = extract_file(file_bytes, ext, start_1b, end_1b)
                out_name = _build_output_filename(file_name, start_1b, end_1b, custom_name or None)

                result = {
                    "file_name": out_name,
                    "mime_type": mime,
                    "size_bytes": len(out_bytes),
                    "source_file": file_name,
                    "page_range": page_range or "all",
                }

                yield self.create_json_message(result)
                for key, value in result.items():
                    yield self.create_variable_message(key, value)
                yield self.create_blob_message(
                    blob=out_bytes,
                    meta={"file_name": out_name, "mime_type": mime},
                )
                return

            raw_chunks = split_by_page(
                file_bytes, ext,
                page_range=page_range,
                pages_per_chunk=pages_per_chunk,
            )
        except Exception as e:
            yield self.create_text_message(f"Error processing document: {e}")
            return

        chunks = []
        for i, chunk in enumerate(raw_chunks):
            meta = chunk.get("metadata", {})
            meta["char_count"] = len(chunk["content"])
            chunks.append({
                "index": i,
                "content": chunk["content"],
                "metadata": meta,
            })

        yield self.create_json_message({
            "total_chunks": len(chunks),
            "file_name": file_name,
            "file_type": ext,
            "split_mode": split_mode,
            "chunks": chunks,
        })
