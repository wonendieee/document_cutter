from __future__ import annotations

import os
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from core.splitters.anchor_splitter import split_by_anchors
from core.splitters.page_splitter import split_by_page
from core.splitters.semantic_splitter import split_semantic

SUPPORTED_EXTENSIONS = {"pdf", "docx", "xlsx", "xls"}


class SplitDocumentTool(Tool):
    def _invoke(
        self, tool_parameters: dict[str, Any]
    ) -> Generator[ToolInvokeMessage]:
        file_info = tool_parameters.get("file")
        if not file_info:
            yield self.create_text_message("Error: no file provided.")
            return

        if isinstance(file_info, dict):
            file_name = file_info.get("filename", file_info.get("name", "unknown"))
            file_bytes = file_info.get("content", b"")
            if isinstance(file_bytes, str):
                import base64
                file_bytes = base64.b64decode(file_bytes)
        else:
            file_name = getattr(file_info, "filename", None) or getattr(file_info, "name", "unknown")
            file_bytes = file_info.read() if hasattr(file_info, "read") else bytes(file_info)

        ext = os.path.splitext(file_name)[1].lower().lstrip(".")
        if ext not in SUPPORTED_EXTENSIONS:
            yield self.create_text_message(
                f"Error: unsupported file type '.{ext}'. Supported: {', '.join(SUPPORTED_EXTENSIONS)}"
            )
            return

        split_mode = tool_parameters.get("split_mode", "page")
        max_chunk_size = int(tool_parameters.get("max_chunk_size", 2000))
        overlap_size = int(tool_parameters.get("overlap_size", 200))

        try:
            if split_mode == "page":
                raw_chunks = split_by_page(file_bytes, ext)
            elif split_mode == "anchor":
                anchors_input = tool_parameters.get("anchors")
                if not anchors_input:
                    yield self.create_text_message(
                        "Error: 'anchors' parameter is required when split_mode='anchor'."
                    )
                    return
                raw_chunks = split_by_anchors(file_bytes, ext, anchors_input)
            else:
                raw_chunks = split_semantic(
                    file_bytes, ext,
                    max_chunk_size=max_chunk_size,
                    overlap_size=overlap_size,
                )
        except Exception as e:
            yield self.create_text_message(f"Error splitting document: {e}")
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
