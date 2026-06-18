"""Small token utilities shared across analyzer surfaces.

Ported from xlide_vscode/src/analyzer/lexer/tokenHelpers.ts. Keeps statement-level
token handling (comment/newline filtering, identifier extraction, leading line
numbers, paren matching) from drifting between surfaces. VBA is case-insensitive
(MS-VBAL 3.3.5), so name matching folds case.
"""

from __future__ import annotations

import re
from collections.abc import Sequence

from .token_kinds import TokenKind, VbaToken
from .tokenize import tokenize

IDENT_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_DECIMAL_RE = re.compile(r"^\d+$")


def is_ident_like(token: VbaToken) -> bool:
    """True when the token reads as a bare identifier (identifier or keyword)."""
    return (
        token.kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD)
        and IDENT_RE.match(token.raw_text) is not None
    )


def statement_tokens(source: str, start: int, end: int) -> list[VbaToken]:
    """Significant tokens of a source span, excluding comments and newlines."""
    return [
        t
        for t in tokenize(source[start:end])
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]


def token_name(token: VbaToken | None) -> str | None:
    """Identifier-like name of a token (unwraps bracketed identifiers)."""
    if token is None:
        return None
    if token.kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD):
        return token.raw_text
    if token.kind is TokenKind.BRACKETED_IDENTIFIER:
        return token.raw_text[1:-1]
    return None


def token_word(token: VbaToken | None) -> str:
    """Canonical (case-folded) text of a token, used for keyword matching.

    Keyword tokens carry canonical_text from the lexer; otherwise the raw text is
    lowered because VBA is case-insensitive (MS-VBAL 3.3.5).
    """
    if token is None:
        return ""
    text = token.canonical_text if token.canonical_text is not None else token.raw_text
    return text.lower()


def is_decimal_line_number(token: VbaToken | None) -> bool:
    """True when the token is a decimal line-number literal."""
    return (
        token is not None
        and token.kind is TokenKind.INTEGER_LITERAL
        and _DECIMAL_RE.match(token.raw_text) is not None
    )


def tokens_without_leading_line_number(tokens: Sequence[VbaToken]) -> list[VbaToken]:
    """Drops the leading line-number token when one prefixes the statement."""
    if len(tokens) > 1 and is_decimal_line_number(tokens[0]):
        return list(tokens[1:])
    return list(tokens)


def match_paren_from(tokens: Sequence[VbaToken], open_index: int) -> int:
    """Index of the ')' matching the '(' at open_index, or -1 when unmatched."""
    depth = 0
    for i in range(open_index, len(tokens)):
        raw = tokens[i].raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
            if depth == 0:
                return i
    return -1


def split_top_level_token_groups(
    tokens: Sequence[VbaToken],
    start: int,
    separator: str,
    end: int | None = None,
) -> list[list[VbaToken]]:
    """Split tokens[start:end) into separator-delimited groups at paren depth 0."""
    if end is None:
        end = len(tokens)
    groups: list[list[VbaToken]] = []
    current: list[VbaToken] = []
    depth = 0
    for i in range(start, end):
        raw = tokens[i].raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth = max(0, depth - 1)
        if depth == 0 and raw == separator:
            groups.append(current)
            current = []
            continue
        current.append(tokens[i])
    groups.append(current)
    return groups
