"""Leading-trivia scanner for the VBA lexer.

Ported from xlide_vscode/src/analyzer/lexer/trivia.ts. Verified against MS-VBAL
v20250520, section 3.2.2 (WSC, line-continuation).
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from .token_kinds import _WSC, Trivia, TriviaKind, is_line_terminator, is_wsc

# The WSC class as a run matcher, built from the same frozenset the predicate
# uses so the two can never disagree. Always matches (possibly empty).
_WSC_RUN_RE = re.compile("[" + re.escape("".join(sorted(_WSC))) + "]*")


@dataclass(slots=True)
class TriviaScan:
    """A run of scanned trivia plus the advanced cursor."""

    trivia: list[Trivia]
    pos: int
    line: int
    character: int


def scan_leading_trivia(src: str, pos: int, line: int, character: int) -> TriviaScan:
    """Consume whitespace and line-continuation trivia starting at pos.

    Stops at the first character that begins a real token (including a line
    terminator, which is a significant newline token, not trivia). A
    line-continuation (1*WSC underscore line-terminator, MS-VBAL 3.2.2) is merged
    into a single lineContinuation trivia so the logical line is preserved while
    the raw text round-trips. The VBE also continues a line whose underscore has
    whitespace after it, and drops that whitespace when it stores the line, so the
    trivia takes it in too (XLIDE issue #83).
    """
    length = len(src)
    # Fast path: most tokens have no leading trivia at all, so answer that
    # without allocating a list. The regex below always matches (possibly
    # empty), so end == pos means "no whitespace here".
    run_end = _WSC_RUN_RE.match(src, pos).end()  # type: ignore[union-attr]
    if run_end == pos:
        return TriviaScan(trivia=[], pos=pos, line=line, character=character)

    trivia: list[Trivia] = []
    while pos < length:
        if not is_wsc(src[pos]):
            break
        start = pos
        # WSC never contains a line terminator, so a whole run advances the
        # column by its length; matching runs beats one predicate call per space.
        pos = _WSC_RUN_RE.match(src, pos).end()  # type: ignore[union-attr]
        character += pos - start
        # A line-continuation is whitespace + '_' + line terminator.
        terminator = continuation_terminator(src, pos)
        if terminator >= 0:
            pos = terminator  # consume '_' and any whitespace after it
            # consume the line terminator (CRLF, CR, or LF)
            if src[pos] == "\r" and pos + 1 < length and src[pos + 1] == "\n":
                pos += 2
            else:
                pos += 1
            line += 1
            character = 0
            trivia.append(
                Trivia(kind=TriviaKind.LINE_CONTINUATION, text=src[start:pos], start=start, end=pos)
            )
        else:
            trivia.append(
                Trivia(kind=TriviaKind.WHITESPACE, text=src[start:pos], start=start, end=pos)
            )
    return TriviaScan(trivia=trivia, pos=pos, line=line, character=character)


def continuation_terminator(src: str, pos: int) -> int:
    """Where the line terminator starts when ``pos`` holds the underscore of a line
    continuation, whitespace after the underscore allowed; otherwise -1."""
    if pos >= len(src) or src[pos] != "_":
        return -1
    at = _WSC_RUN_RE.match(src, pos + 1).end()  # type: ignore[union-attr]
    return at if at < len(src) and is_line_terminator(src[at]) else -1
