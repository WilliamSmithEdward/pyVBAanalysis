"""Leading-trivia scanner for the VBA lexer.

Ported from xlide_vscode/src/analyzer/lexer/trivia.ts. Verified against MS-VBAL
v20250520, section 3.2.2 (WSC, line-continuation).
"""

from __future__ import annotations

from dataclasses import dataclass

from .token_kinds import Trivia, TriviaKind, is_line_terminator, is_wsc


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
    the raw text round-trips.
    """
    trivia: list[Trivia] = []
    length = len(src)
    while pos < length:
        if not is_wsc(src[pos]):
            break
        start = pos
        while pos < length and is_wsc(src[pos]):
            pos += 1
            character += 1
        # A line-continuation is whitespace + '_' + line terminator.
        if (
            pos < length
            and src[pos] == "_"
            and pos + 1 < length
            and is_line_terminator(src[pos + 1])
        ):
            pos += 1  # consume '_'
            character += 1
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
