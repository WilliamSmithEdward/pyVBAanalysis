"""VBA lexical token kinds.

Ported from xlide_vscode/src/analyzer/lexer/tokenKinds.ts. Verified against
MS-VBAL v20250520, section 3.3 (Lexical Tokens).

The lexer is loss-aware: every character of the source appears in exactly one
token's raw_text or in a token's leading_trivia (and any end-of-file trailing
trivia on the last token), so concatenating the stream reproduces the source
exactly. That round-trip property is the lexer acceptance gate.

Offsets are code-point offsets into the Python string. XLIDE uses UTF-16 offsets;
for the BMP (all normal VBA) they coincide. Astral characters (rare, only inside
strings or comments) would differ in offset, but round-trip still holds because
slicing is consistent.
"""

from __future__ import annotations

import enum
from dataclasses import dataclass


class TokenKind(str, enum.Enum):
    """The category of a lexical token (MS-VBAL 3.3)."""

    NEWLINE = "newline"
    COMMENT = "comment"
    KEYWORD = "keyword"
    IDENTIFIER = "identifier"
    BRACKETED_IDENTIFIER = "bracketedIdentifier"
    INTEGER_LITERAL = "integerLiteral"
    FLOAT_LITERAL = "floatLiteral"
    DATE_LITERAL = "dateLiteral"
    STRING_LITERAL = "stringLiteral"
    OPERATOR = "operator"
    PUNCTUATION = "punctuation"
    COLON = "colon"
    DIRECTIVE = "directive"
    UNKNOWN = "unknown"


class TriviaKind(str, enum.Enum):
    """Insignificant text attached to the following token (MS-VBAL 3.2.2)."""

    WHITESPACE = "whitespace"
    LINE_CONTINUATION = "lineContinuation"


@dataclass(frozen=True, slots=True)
class Trivia:
    """Whitespace or a line continuation attached to a token."""

    kind: TriviaKind
    text: str
    start: int
    end: int


@dataclass(slots=True)
class VbaToken:
    """A VBA lexical token.

    Mutable so the tokenizer can attach end-of-file trailing trivia to the final
    token after the fact; everything else is set at construction.
    """

    kind: TokenKind
    raw_text: str
    start: int
    end: int
    line: int
    character: int
    # Canonical capitalization for keyword tokens (MS-VBAL 3.3.5.2); None otherwise.
    canonical_text: str | None = None
    # Whitespace / line continuations immediately preceding this token.
    leading_trivia: tuple[Trivia, ...] = ()
    # Trailing trivia, only populated on the final token at end of file so the
    # stream stays perfectly round-trippable.
    trailing_trivia: tuple[Trivia, ...] = ()


# WSC: whitespace characters (MS-VBAL 3.2.2), excluding line terminators. This
# reproduces tokenKinds.ts isWsc exactly: the explicit XLIDE set (tab, the
# eom-character U+0019, space, ideographic space U+3000) unioned with the
# JavaScript \s class minus the line terminators it drops ({\n \r \v \f}). U+2028
# and U+2029 are in \s and are not line terminators here, so they count as WSC.
# Code points are spelled out so the source stays plain ASCII.
_WSC: frozenset[str] = frozenset(
    {
        chr(0x09),  # tab
        chr(0x19),  # eom-character (XLIDE explicit)
        chr(0x20),  # space
        chr(0xA0),  # no-break space
        chr(0x1680),  # ogham space mark
        chr(0x2028),  # line separator
        chr(0x2029),  # paragraph separator
        chr(0x202F),  # narrow no-break space
        chr(0x205F),  # medium mathematical space
        chr(0x3000),  # ideographic (DBCS) space
        chr(0xFEFF),  # zero-width no-break space
    }
    | {chr(c) for c in range(0x2000, 0x200B)}  # U+2000 - U+200A spaces
)


def is_wsc(ch: str) -> bool:
    """True when ch is a whitespace character (MS-VBAL 3.2.2 WSC)."""
    return ch in _WSC


def is_line_terminator(ch: str) -> bool:
    """True when ch is a line terminator (CR or LF)."""
    return ch == "\r" or ch == "\n"
