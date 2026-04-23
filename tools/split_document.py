from __future__ import annotations

import base64
import io
import json
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

EXT_TO_MIME = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xls": "application/vnd.ms-excel",
}


def _extract_file_bytes(file_info) -> tuple[str, bytes]:
    """Resolve filename and byte content from a Dify File-like object or dict."""
    if isinstance(file_info, (list, tuple)):
        if not file_info:
            raise ValueError("File list is empty.")
        file_info = file_info[0]

    if hasattr(file_info, "blob"):
        return (
            getattr(file_info, "filename", None)
            or getattr(file_info, "name", None)
            or "unknown",
            file_info.blob,
        )

    if isinstance(file_info, dict):
        file_name = file_info.get("filename") or file_info.get("name") or "unknown"
        file_bytes = file_info.get("blob") or file_info.get("content") or b""
        if isinstance(file_bytes, str):
            file_bytes = base64.b64decode(file_bytes)
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


def _load_json_object(raw: str, default: dict) -> dict:
    raw = (raw or "").strip()
    if not raw:
        return dict(default)
    try:
        obj = json.loads(raw)
    except Exception as e:
        raise ValueError(f"Invalid JSON: {e}") from e
    if not isinstance(obj, dict):
        raise ValueError("JSON must decode to an object.")
    return obj


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

        split_mode = tool_parameters.get("split_mode") or "page_text"
        page_range = str(tool_parameters.get("page_range") or "")
        pages_per_chunk = int(tool_parameters.get("pages_per_chunk") or 1)
        delivery_mode = str(tool_parameters.get("delivery_mode") or "blob")
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
                    "delivery_mode": delivery_mode,
                }

                if delivery_mode == "upload_link":
                    yield from self._deliver_upload(out_bytes, out_name, mime, result)
                    return

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

    def _deliver_upload(
        self,
        out_bytes: bytes,
        out_name: str,
        mime: str,
        result: dict,
    ) -> Generator[ToolInvokeMessage]:
        creds = getattr(self.runtime, "credentials", {}) or {}

        upload_url = str(creds.get("upload_url") or "").strip()
        if not upload_url:
            yield self.create_text_message(
                "Error: Provider credential 'upload_url' is required when delivery_mode=upload_link."
            )
            return

        upload_token = str(creds.get("upload_token") or "").strip()
        file_field = (str(creds.get("upload_file_field_name") or "file").strip() or "file")
        resp_url_field = (
            str(creds.get("response_download_url_field") or "download_url").strip()
            or "download_url"
        )
        resp_name_field = (
            str(creds.get("response_file_name_field") or "file_name").strip()
            or "file_name"
        )

        try:
            headers = _load_json_object(str(creds.get("upload_headers_json") or ""), {})
            form_data = _load_json_object(str(creds.get("upload_form_data_json") or ""), {})
        except ValueError as e:
            yield self.create_text_message(f"Error: invalid provider credential JSON. {e}")
            return

        form_data.setdefault("desired_name", out_name)
        if upload_token:
            headers.setdefault("Authorization", f"Bearer {upload_token}")

        files = {file_field: (out_name, io.BytesIO(out_bytes), mime)}

        try:
            resp = requests.post(
                upload_url,
                headers=headers,
                data=form_data,
                files=files,
                timeout=60,
            )
            resp.raise_for_status()
            resp_json = resp.json()
        except Exception as e:
            yield self.create_text_message(f"Error uploading file to {upload_url}: {e}")
            return

        if not isinstance(resp_json, dict):
            yield self.create_text_message("Error: upload service must return a JSON object.")
            return

        download_url = resp_json.get(resp_url_field)
        uploaded_name = resp_json.get(resp_name_field) or out_name

        if not download_url:
            yield self.create_text_message(
                f"Error: upload succeeded but response field '{resp_url_field}' not found. "
                f"Response keys: {list(resp_json.keys())}"
            )
            return

        result["download_url"] = download_url
        result["returned_file_name"] = uploaded_name

        yield self.create_json_message(result)
        for key, value in result.items():
            yield self.create_variable_message(key, value)
        yield self.create_text_message(f"Download URL: {download_url}")
