"""Parser state: logical statements and a forward cursor.

Ported from xlide_vscode/src/analyzer/parser/parserState.ts. MS-VBAL 3.3.1: a
logical line ends at a line terminator or a ':' statement separator, so the token
stream splits on newline and colon tokens. Line continuations are already merged
into trivia by the lexer, so a continued physical line is one logical statement.
Statement boundaries are the parser's natural recovery points (error tolerance).
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

from ..lexer.token_helpers import token_word
from ..lexer.token_kinds import TokenKind, VbaToken

__all__ = [
    "LogicalStatement",
    "split_logical_statements",
    "StatementCursor",
    "code_tokens",
    "token_word",
]


@dataclass(slots=True)
class LogicalStatement:
    """The significant tokens between two separators (no newline/colon)."""

    # Significant tokens (no newline/colon separators; trailing comment kept).
    tokens: list[VbaToken]
    # Absolute offset of the first token.
    start: int
    # Absolute offset just past the last token.
    end: int
    # Zero-based line of the first token.
    line: int


def split_logical_statements(tokens: Sequence[VbaToken]) -> list[LogicalStatement]:
    """Split a token stream into logical statements (MS-VBAL 3.3.1 EOS).

    Newline and colon tokens act as separators and are not included in any
    statement. Empty statements (blank lines, doubled separators) are dropped.
    """
    statements: list[LogicalStatement] = []
    current: list[VbaToken] = []

    def flush() -> None:
        nonlocal current
        if not current:
            return
        first = current[0]
        last = current[-1]
        statements.append(
            LogicalStatement(tokens=current, start=first.start, end=last.end, line=first.line)
        )
        current = []

    for token in tokens:
        if token.kind is TokenKind.NEWLINE or token.kind is TokenKind.COLON:
            flush()
            continue
        current.append(token)
    flush()
    return statements


class StatementCursor:
    """A forward cursor over a list of logical statements.

    The parser consumes statements one at a time; block parsers peek ahead to find
    their closers.
    """

    __slots__ = ("_statements", "_index")

    def __init__(self, statements: Sequence[LogicalStatement]) -> None:
        self._statements = statements
        self._index = 0

    def at_end(self) -> bool:
        """True when no statements remain."""
        return self._index >= len(self._statements)

    def peek(self) -> LogicalStatement | None:
        """The current statement without consuming it, or None at end."""
        if self._index < len(self._statements):
            return self._statements[self._index]
        return None

    def next(self) -> LogicalStatement | None:
        """Consume and return the current statement, or None at end.

        Mirrors the TS `statements[index++]`: the index always advances, even at
        end (so a guarded `while not at_end()` loop terminates cleanly).
        """
        stmt = self._statements[self._index] if self._index < len(self._statements) else None
        self._index += 1
        return stmt

    def position(self) -> int:
        """Current cursor position (for span bookkeeping)."""
        return self._index


def code_tokens(statement: LogicalStatement) -> list[VbaToken]:
    """Significant tokens excluding any trailing comment (MS-VBAL 3.3.1)."""
    return [t for t in statement.tokens if t.kind is not TokenKind.COMMENT]
