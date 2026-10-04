"""The file numbers a module's Open statements name (XLIDE issue #419).

Ported from xlide_vscode/src/analyzer/diagnostics/openedFileNumbers.ts. A file
statement on a number no Open in the project names raises 52, "Bad file name or
number", wherever it runs; one Open whose number is a variable or FreeFile may
open any number, and then none is judged.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass

from ..lexer.token_helpers import token_word
from ..lexer.token_kinds import TokenKind
from ..lexer.tokenize import tokenize_cached

_DECIMAL_RE = re.compile(r"^\d+$")


@dataclass(frozen=True, slots=True)
class OpenedFileNumbers:
    """What a module's, or a project's, Open statements name."""

    # An Open whose number is not a literal: it may open any.
    any: bool
    # The literal numbers Opens name.
    numbers: frozenset[int]


def opened_file_numbers_in(source: str) -> OpenedFileNumbers:
    """The numbers `Open ... As #n` names in a module's source: a statement of its
    own, after a colon or a line number, or a one-line If's arm."""
    toks = [tok for tok in tokenize_cached(source) if tok.kind is not TokenKind.COMMENT]
    numbers: set[int] = set()
    any_ = False
    for i, tok in enumerate(toks):
        if token_word(tok) != "open":
            continue
        before = toks[i - 1] if i > 0 else None
        line_number = (
            before is not None
            and before.kind is TokenKind.INTEGER_LITERAL
            and (i < 2 or toks[i - 2].kind is TokenKind.NEWLINE)
        )
        starts = (
            before is None
            or before.kind is TokenKind.NEWLINE
            or before.kind is TokenKind.COLON
            or line_number
            or token_word(before) == "then"
            or token_word(before) == "else"
        )
        if not starts:
            continue
        j = i + 1
        while j < len(toks) and toks[j].kind is not TokenKind.NEWLINE and toks[j].kind is not TokenKind.COLON:
            if token_word(toks[j]) != "as":
                j += 1
                continue
            after = toks[j + 1] if j + 1 < len(toks) else None
            number = (toks[j + 2] if j + 2 < len(toks) else None) if after is not None and after.raw_text == "#" else after
            if (
                number is not None
                and number.kind is TokenKind.INTEGER_LITERAL
                and _DECIMAL_RE.match(number.raw_text)
            ):
                numbers.add(int(number.raw_text))
            else:
                any_ = True
            break
    return OpenedFileNumbers(any=any_, numbers=frozenset(numbers))


def merge_opened_file_numbers(parts: Sequence[OpenedFileNumbers]) -> OpenedFileNumbers:
    """Both modules' Opens, or a project's and one module's current text."""
    return OpenedFileNumbers(
        any=any(part.any for part in parts),
        numbers=frozenset(n for part in parts for n in part.numbers),
    )
