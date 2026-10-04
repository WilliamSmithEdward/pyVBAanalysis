"""Rule: places where the VBE refuses a line continuation (XLIDE issue #126).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/lineContinuations.ts.
Measured in Excel 16.0 (build 20326, 2026-09-25):

- 25 continuations in one logical line: AddFromString refuses the module, "Too
  many line continuations". 24 compile.
- A continuation followed by a blank line: `x = 1 + _` then an empty line. Syntax
  error.
- Any continuation inside an Enum body, on a member line or before `End Enum`:
  "Invalid inside Enum". The same splits inside a Type compile, and so does one
  after `Private` on the Enum header line.

A Declare continued between its Lib (or Alias) string and its parameter list
compiles. It was reported until a recheck on 2026-10-01: the "Syntax error" came
from VBComponents.CodeModule.AddFromString, which stores a stray `()` line after
such a Declare. The same module imported from a .bas file, saved and reopened,
compiles, and so does VBA-JSON, which has this shape in its Windows branch.

The analyzer follows continuations everywhere else the VBE does; those places are
covered by XLIDE's tests/diagnostics/lineContinuations.test.ts.
"""

from __future__ import annotations

import re

from ...lexer.token_helpers import first_token_at_or_after
from ...lexer.token_kinds import TokenKind, Trivia, TriviaKind
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import EnumNode, ModuleNode, Span
from ..context import PushFn

_MAX_CONTINUATIONS = 24

_LINE_TERMINATOR_RE = re.compile(r"[\r\n]")
_SPACES_AND_TABS_RE = re.compile(r"[ \t]*")


def check_line_continuation_limits(source: str, mod: ModuleNode, push: PushFn) -> None:
    tokens = tokenize_cached(source)
    continuations: list[Trivia] = []
    in_logical_line = 0
    for tok in tokens:
        for trivia in tok.leading_trivia:
            if trivia.kind is not TriviaKind.LINE_CONTINUATION:
                continue
            continuations.append(trivia)
            in_logical_line += 1
            if in_logical_line == _MAX_CONTINUATIONS + 1:
                push(
                    "invalidLineContinuation",
                    f"This is the {_MAX_CONTINUATIONS + 1}th line continuation in one logical "
                    f'line; the VBE allows {_MAX_CONTINUATIONS} ("Too many line continuations").',
                    Span(trivia.start, trivia.end),
                )
            # The continuation joined this line to the next; a newline token
            # right after it means the next physical line is empty.
            if tok.kind is TokenKind.NEWLINE and _only_whitespace_between(
                source, trivia.end, tok.start
            ):
                push(
                    "invalidLineContinuation",
                    "A line continuation must be followed by more of the statement; the next "
                    "line is empty. This is a VBE compile error: Syntax error.",
                    Span(trivia.start, trivia.end),
                )
        if tok.kind is TokenKind.NEWLINE:
            in_logical_line = 0
    if len(continuations) == 0:
        return
    for member in mod.members:
        if isinstance(member, EnumNode):
            header_end = _line_end_at_or_after(
                source, member.name_span.end if member.name_span is not None else member.span.start
            )
            for i in range(first_token_at_or_after(continuations, header_end + 1), len(continuations)):
                trivia = continuations[i]
                if trivia.start >= member.span.end:
                    break
                push(
                    "invalidLineContinuation",
                    f"A line continuation is not allowed inside Enum '{member.name}': neither "
                    'on a member line nor before End Enum ("Invalid inside Enum").',
                    Span(trivia.start, trivia.end),
                )


def _only_whitespace_between(source: str, start: int, end: int) -> bool:
    # fullmatch, not match with `$`: Python's `$` also matches before a final
    # newline, where the JavaScript `/^[ \t]*$/` does not.
    return _SPACES_AND_TABS_RE.fullmatch(source, start, max(start, end)) is not None


def _line_end_at_or_after(source: str, start: int) -> int:
    """Offset of the first line terminator at or after `start`.

    Private copy of upstream's vbaSourceScan.ts lineEndAtOrAfter (outside the
    analyzer tree the port mirrors).
    """
    found = _LINE_TERMINATOR_RE.search(source, start)
    return found.start() if found is not None else len(source)
