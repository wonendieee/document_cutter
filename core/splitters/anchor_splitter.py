from __future__ import annotations

import json

from core.parsers.word_parser import parse_word_by_anchors


def _normalize_anchors_input(anchors_raw) -> list[dict]:
    """
    Accept anchors input in multiple forms:
      - list of dicts (already parsed)
      - JSON string of list
      - JSON string of object with key 'anchors_json'
      - dict with key 'anchors_json'
    """
    if anchors_raw is None:
        return []

    if isinstance(anchors_raw, str):
        try:
            anchors_raw = json.loads(anchors_raw)
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid anchors JSON string: {e}")

    if isinstance(anchors_raw, dict):
        if "anchors_json" in anchors_raw:
            anchors_raw = anchors_raw["anchors_json"]
        elif "anchors" in anchors_raw:
            anchors_raw = anchors_raw["anchors"]
        else:
            raise ValueError("anchors dict must contain 'anchors_json' or 'anchors' key")

    if not isinstance(anchors_raw, list):
        raise ValueError(f"anchors must be a list, got {type(anchors_raw).__name__}")

    return anchors_raw


def split_by_anchors(file_bytes: bytes, file_ext: str, anchors_raw) -> list[dict]:
    """
    Split a document by upstream-provided heading anchors.
    Currently supports only .docx (anchor matching requires structural paragraphs).
    """
    ext = file_ext.lower().lstrip(".")
    anchors = _normalize_anchors_input(anchors_raw)

    if ext != "docx":
        raise ValueError(
            f"Anchor split mode only supports .docx files, got .{ext}. "
            "Use 'page' or 'semantic' mode for other formats."
        )

    return parse_word_by_anchors(file_bytes, anchors)
