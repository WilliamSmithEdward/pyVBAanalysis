"""The `'''` documentation-comment grammar (ported from src/analyzer/docs)."""

from __future__ import annotations

from .doc_comment import (
    DocBlockLine,
    DocTagOccurrence,
    attached_comments_start,
    leading_doc_lines,
    scan_doc_tags,
)

__all__ = [
    "DocBlockLine",
    "DocTagOccurrence",
    "attached_comments_start",
    "leading_doc_lines",
    "scan_doc_tags",
]
