"""Rule family: lexical source-shape rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/lexical.ts: unterminated
string literals and invalid line continuations. Self-contained (source + tokens).
"""

from __future__ import annotations

import re

from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import Span
from ..context import PushFn

# VBA whitespace characters relevant to line continuations: tab, the eom character
# U+0019, space, and the ideographic (DBCS) space U+3000.
_VBA_WSC: frozenset[str] = frozenset({"\t", chr(0x19), " ", chr(0x3000)})
_IDENT_PART: frozenset[str] = frozenset(
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_"
)


def _char(source: str, i: int) -> str:
    """Char at i, or '' when out of range (matches the TS undefined behavior)."""
    return source[i] if 0 <= i < len(source) else ""


def is_vba_wsc(ch: str) -> bool:
    return ch in _VBA_WSC


def is_identifier_part_char(ch: str) -> bool:
    return ch in _IDENT_PART


def _count_quotes(text: str) -> int:
    return text.count('"')


# A blank, an underscore and nothing but blanks after it: what the VBE reads as a
# line continuation. (JavaScript's `$` without the m flag is the end of the text,
# which Python spells `\Z`.)
_CONTINUATION_AT_END = re.compile(r"[ \t]_[ \t]*\Z")


def check_unterminated_strings(source: str, push: PushFn) -> None:
    """A string left open at the end of its line, when that line ends like a line
    continuation.

    The VBE closes any other open string at the end of its line, so `Main = "abc`
    compiles and gives "abc" (issue #681, measured in Excel 16.0); one that
    swallowed a needed `)` is unbalanced-parens. But `Main = "abc _` is a Syntax
    error, and so is `"abc _ `, while `"abc_`, `"abc __` and `"a _b` compile.
    """
    toks = tokenize_cached(source)
    previous: VbaToken | None = None
    for tok in toks:
        if tok.kind is TokenKind.STRING_LITERAL and _count_quotes(tok.raw_text) % 2 == 1:
            # One that opens a statement stands alone once closed, `"` or `"abc`
            # on a line of its own, after `:` or a label, or after Then: no
            # statement (issue #740, measured in Excel 16.0).
            opens_statement = (
                previous is None
                or previous.kind is TokenKind.NEWLINE
                or previous.raw_text == ":"
                or (previous.kind is TokenKind.KEYWORD and previous.raw_text.lower() in ("then", "else"))
            )
            if _CONTINUATION_AT_END.search(tok.raw_text):
                push(
                    "unterminatedString",
                    "Unterminated string literal: its line ends with ' _', which the VBE reads as a "
                    "line continuation. This is a VBE compile error: Syntax error.",
                    Span(tok.start, tok.end),
                )
            elif opens_statement:
                push(
                    "unterminatedString",
                    "Unterminated string literal: closed at the end of its line, it stands alone "
                    "where a statement goes. This is a VBE compile error: Syntax error.",
                    Span(tok.start, tok.end),
                )
        if tok.kind is not TokenKind.COMMENT:
            previous = tok


def check_invalid_line_continuations(source: str, push: PushFn) -> None:
    """VBA line-continuation trivia is strictly `1*WSC "_" line-terminator`.

    A continuation underscore with trailing text/comment, or without the required
    whitespace before it, is a settled compile-time syntax error.
    """
    n = len(source)
    line_start = 0
    while line_start < n:
        line_end = line_start
        while line_end < n and source[line_end] != "\r" and source[line_end] != "\n":
            line_end += 1
        _check_line(source, line_start, line_end, push)
        if line_end >= n:
            break
        line_start = (
            line_end + 2
            if (source[line_end] == "\r" and _char(source, line_end + 1) == "\n")
            else line_end + 1
        )


def _check_line(source: str, line_start: int, line_end: int, push: PushFn) -> None:
    comment_start = _physical_line_comment_start(source, line_start, line_end)
    if comment_start is None:
        comment_start = line_end
    code_last = _last_non_wsc_offset(source, line_start, comment_start)
    if code_last is None:
        return
    visible_line_end = _last_non_wsc_offset(source, line_start, line_end)
    span_end = line_end if visible_line_end is None else visible_line_end + 1

    for underscore in _underscores_outside_strings(source, line_start, comment_start):
        prev = _char(source, underscore - 1)
        nxt = _char(source, underscore + 1)
        prev_is_wsc = underscore > line_start and is_vba_wsc(prev)
        next_starts_identifier = is_identifier_part_char(nxt)
        has_trailing_text = _first_non_wsc_offset(source, underscore + 1, line_end) is not None

        if prev_is_wsc and has_trailing_text and not next_starts_identifier:
            push(
                "invalidLineContinuation",
                "Line continuation '_' must be the final non-whitespace character on the physical line.",
                Span(underscore, max(underscore + 1, span_end)),
            )
            return

        if (
            underscore == code_last
            and line_end < len(source)
            and not prev_is_wsc
            and not is_identifier_part_char(prev)
        ):
            push(
                "invalidLineContinuation",
                "Line continuation '_' must be preceded by whitespace.",
                Span(underscore, underscore + 1),
            )
            return


def _physical_line_comment_start(source: str, line_start: int, line_end: int) -> int | None:
    in_string = False
    statement_start = True
    i = line_start
    while i < line_end:
        ch = source[i]
        if in_string:
            if ch == '"':
                if _char(source, i + 1) == '"':
                    i += 1
                else:
                    in_string = False
            i += 1
            continue
        if ch == '"':
            in_string = True
            statement_start = False
            i += 1
            continue
        if ch == "'":
            return i
        if is_vba_wsc(ch):
            i += 1
            continue
        if ch == ":":
            statement_start = True
            i += 1
            continue
        if statement_start and _starts_rem_comment(source, i, line_end):
            return i
        statement_start = False
        i += 1
    return None


def _underscores_outside_strings(source: str, start: int, end: int) -> list[int]:
    offsets: list[int] = []
    in_string = False
    i = start
    while i < end:
        ch = source[i]
        if in_string:
            if ch == '"':
                if _char(source, i + 1) == '"':
                    i += 1
                else:
                    in_string = False
            i += 1
            continue
        if ch == '"':
            in_string = True
            i += 1
            continue
        if ch == "_":
            offsets.append(i)
        i += 1
    return offsets


def _starts_rem_comment(source: str, offset: int, end: int) -> bool:
    return (
        offset + 3 <= end
        and source[offset : offset + 3].lower() == "rem"
        and not is_identifier_part_char(_char(source, offset + 3))
    )


def _first_non_wsc_offset(source: str, start: int, end: int) -> int | None:
    for i in range(start, end):
        if not is_vba_wsc(source[i]):
            return i
    return None


def _last_non_wsc_offset(source: str, start: int, end: int) -> int | None:
    for i in range(end - 1, start - 1, -1):
        if not is_vba_wsc(source[i]):
            return i
    return None
