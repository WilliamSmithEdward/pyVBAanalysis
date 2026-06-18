"""Fixed-length string type detection (String * length).

Ported from xlide_vscode/src/analyzer/parser/fixedLengthString.ts. MS-VBAL
5.2.3.3 / 5.3.1.4: `String * length` is a fixed-length string whose length is an
integer literal or a constant name.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

from ..lexer.token_helpers import token_word
from ..lexer.token_kinds import TokenKind, VbaToken


@dataclass(frozen=True, slots=True)
class FixedLengthStringType:
    """Token indices of a recognized `String * length` triple."""

    type_start: int
    star_index: int
    length_index: int
    end_index: int


def _at(tokens: Sequence[VbaToken], i: int) -> VbaToken | None:
    return tokens[i] if 0 <= i < len(tokens) else None


def _is_fixed_length_string_length_token(token: VbaToken | None) -> bool:
    if token is None:
        return False
    return token.kind in (
        TokenKind.INTEGER_LITERAL,
        TokenKind.IDENTIFIER,
        TokenKind.BRACKETED_IDENTIFIER,
    )


def parse_fixed_length_string_type(
    tokens: Sequence[VbaToken], type_start: int
) -> FixedLengthStringType | None:
    """Match `String * length` at type_start, or None when it is not present."""
    first = _at(tokens, type_start)
    if first is None or token_word(first) != "string":
        return None
    star = _at(tokens, type_start + 1)
    if star is None or star.raw_text != "*":
        return None
    if not _is_fixed_length_string_length_token(_at(tokens, type_start + 2)):
        return None
    return FixedLengthStringType(
        type_start=type_start,
        star_index=type_start + 1,
        length_index=type_start + 2,
        end_index=type_start + 3,
    )
