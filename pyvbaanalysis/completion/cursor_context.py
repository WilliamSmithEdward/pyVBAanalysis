"""Cursor-context detection for the completion stack (reduced port).

Ported from the prefix-significant-token slice of
xlide_vscode/src/analyzer/completion/cursorContext.ts. The diagnostics seam only
needs ``significant_tokens`` (the prefix tokenized, comments dropped, newlines
kept as statement boundaries); the partial-identifier peel, in-comment/in-string
classification, space-trigger gate, and the per-request cache are completion-UX
and are intentionally dropped.
"""

from __future__ import annotations

from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize


def completion_significant_tokens(source: str, offset: int) -> list[VbaToken]:
    """Prefix tokens before ``offset`` with comments removed, newlines kept.

    Mirrors ``completionCursorContext(source, offset).significantTokens``: tokenize
    only the text up to the clamped cursor and drop comment tokens. Newlines stay
    so a dangling member-access dot on an earlier line is not merged into the chain
    being resolved.
    """
    safe_offset = max(0, min(offset, len(source)))
    return [t for t in tokenize(source[:safe_offset]) if t.kind is not TokenKind.COMMENT]
