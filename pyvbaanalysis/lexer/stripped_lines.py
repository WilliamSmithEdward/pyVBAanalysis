"""Lexer-derived stripped-line substrate.

Ported from xlide_vscode/src/analyzer/lexer/strippedLines.ts. Produces every
physical line with string-literal and comment token spans blanked to spaces,
length and column alignment preserved, so a single lexer defines string/comment
semantics for the consumers built on this substrate.
"""

from __future__ import annotations

import re

from .token_kinds import TokenKind
from .tokenize import tokenize

_LINE_SPLIT_RE = re.compile(r"\r\n|\r|\n")


def lexer_stripped_lines(source: str) -> list[str]:
    """Each physical line with comment and string-literal spans blanked to spaces.

    Comments and string literals never span physical lines (MS-VBAL 3.3.1 / 3.3.4),
    so blanking on the token's start line covers the whole token.
    """
    chars = [list(line) for line in _LINE_SPLIT_RE.split(source)]
    for token in tokenize(source):
        if token.kind is not TokenKind.COMMENT and token.kind is not TokenKind.STRING_LITERAL:
            continue
        if token.line >= len(chars):
            continue
        line_chars = chars[token.line]
        end = min(token.character + len(token.raw_text), len(line_chars))
        for col in range(token.character, end):
            line_chars[col] = " "
    return ["".join(line_chars) for line_chars in chars]


def lexer_stripped_line(line: str) -> str:
    """Single-line variant for call sites stripping one physical line in isolation."""
    result = lexer_stripped_lines(line)
    return result[0] if result else line
