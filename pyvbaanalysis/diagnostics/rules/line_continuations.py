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
- In a Declare, a continuation between the Lib (or Alias) string and the parameter
  list. Syntax error. After PtrSafe, after the name, after Private and inside the
  parameter list all compile.

The analyzer follows continuations everywhere else the VBE does; those places are
covered by XLIDE's tests/diagnostics/lineContinuations.test.ts.
"""

from __future__ import annotations

import re
from bisect import bisect_left
from collections.abc import Sequence

from ...lexer.token_kinds import TokenKind, Trivia, TriviaKind, VbaToken
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import DeclareNode, EnumNode, ModuleNode, Span
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
            header_end = _line_end_after(
                source, member.name_span.end if member.name_span is not None else member.span.start
            )
            for trivia in continuations:
                if trivia.start > header_end and trivia.start < member.span.end:
                    push(
                        "invalidLineContinuation",
                        f"A line continuation is not allowed inside Enum '{member.name}': neither "
                        'on a member line nor before End Enum ("Invalid inside Enum").',
                        Span(trivia.start, trivia.end),
                    )
        elif isinstance(member, DeclareNode):
            gap = _declare_lib_to_params_gap(tokens, member.span)
            if gap is None:
                continue
            for trivia in continuations:
                if trivia.start >= gap.start and trivia.end <= gap.end:
                    push(
                        "invalidLineContinuation",
                        f"Declare '{member.name}' cannot break the line between its Lib or Alias "
                        "string and the parameter list. This is a VBE compile error: Syntax error.",
                        Span(trivia.start, trivia.end),
                    )


def _only_whitespace_between(source: str, start: int, end: int) -> bool:
    # fullmatch, not match with `$`: Python's `$` also matches before a final
    # newline, where the JavaScript `/^[ \t]*$/` does not.
    return _SPACES_AND_TABS_RE.fullmatch(source, start, max(start, end)) is not None


def _line_end_after(source: str, start: int) -> int:
    """Offset of the first line terminator at or after `start`."""
    found = _LINE_TERMINATOR_RE.search(source, start)
    return found.start() if found is not None else len(source)


def _declare_lib_to_params_gap(tokens: Sequence[VbaToken], span: Span) -> Span | None:
    """The source between a Declare's last Lib/Alias string literal and its `(`."""
    last_string: VbaToken | None = None
    # Upstream walks the module's tokens from the first one, skipping those that
    # start before the Declare; the tokens are in source order, so a binary search
    # finds the same first token without the walk.
    for i in range(bisect_left(tokens, span.start, key=_token_start), len(tokens)):
        tok = tokens[i]
        if tok.start >= span.end:
            break
        if tok.kind is TokenKind.STRING_LITERAL:
            last_string = tok
        elif tok.raw_text == "(" and last_string is not None:
            return Span(last_string.end, tok.start)
    return None


def _token_start(tok: VbaToken) -> int:
    return tok.start
