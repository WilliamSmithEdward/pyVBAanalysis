"""Contextual keywords are keywords only inside the statement that makes them so.

Ported from xlide_vscode/src/analyzer/lexer/contextualKeywords.ts (XLIDE issue
#86). Everywhere else they are names, and the VBE spells a variable called
`text`, `binary` or `output` the way it was declared: it capitalizes Text in
`Option Compare Text` and nowhere else.

    Explicit, Base, Compare, Binary, Text   Option (Binary also in Open ... For Binary)
    Lib, Alias, PtrSafe                     Declare
    Step                                    For
    Error                                   On Error, and the Error statement
    Output, Append, Random, Read            Open (Read after Access or Lock)

The lexer calls this once it has the whole token stream, so every consumer sees
the same answer: a keyword token where the word is grammar, an identifier token
where it is a name.
"""

from __future__ import annotations

import re
from collections.abc import Sequence

from .token_kinds import TokenKind, VbaToken

# The words, in the canonical spelling the lexer gives their keyword tokens.
_STATEMENT_BOUND = frozenset(
    {
        "Explicit", "Base", "Compare", "Binary", "Text",
        "Lib", "Alias", "PtrSafe",
        "Step",
        "Error",
        "Output", "Append", "Random", "Read",
    }
)

_LINE_NUMBER_RE = re.compile(r"^\d+$")

_OPERAND_KEYWORDS = frozenset({"true", "false", "nothing", "empty", "null", "me"})

_OPERAND_KINDS = frozenset(
    {
        TokenKind.IDENTIFIER,
        TokenKind.BRACKETED_IDENTIFIER,
        TokenKind.INTEGER_LITERAL,
        TokenKind.FLOAT_LITERAL,
        TokenKind.STRING_LITERAL,
        TokenKind.DATE_LITERAL,
    }
)


def settle_contextual_keywords(tokens: Sequence[VbaToken]) -> None:
    """Turn each statement-bound keyword standing outside its statement into an
    identifier."""
    statement_start = 0
    for i, token in enumerate(tokens):
        if token.kind is TokenKind.NEWLINE or token.kind is TokenKind.COLON:
            statement_start = i + 1
        elif token.kind is TokenKind.KEYWORD and (token.canonical_text or "") in _STATEMENT_BOUND:
            # Only what comes before the word decides it, and the words are rare,
            # so the statement is gathered for them alone.
            statement = [t for t in tokens[statement_start : i + 1] if t.kind is not TokenKind.COMMENT]
            if not _is_keyword_at(statement, len(statement) - 1):
                token.kind = TokenKind.IDENTIFIER
                token.canonical_text = None


def _at(statement: Sequence[VbaToken], i: int) -> VbaToken | None:
    return statement[i] if 0 <= i < len(statement) else None


def _word(token: VbaToken | None) -> str:
    """The word a token spells, lower-cased: its canonical text when it has one."""
    if token is None:
        return ""
    return (token.canonical_text if token.canonical_text is not None else token.raw_text).lower()


def _is_member_access(token: VbaToken | None) -> bool:
    return token is not None and (
        (token.kind is TokenKind.PUNCTUATION and token.raw_text == ".")
        or (token.kind is TokenKind.OPERATOR and token.raw_text == "!")
    )


def _is_keyword_at(statement: Sequence[VbaToken], i: int) -> bool:
    prev = _at(statement, i - 1)
    if _is_member_access(prev):
        return False
    before = _word(prev)
    word = _word(statement[i])
    if word in ("explicit", "base", "compare"):
        return before == "option"
    if word == "text":
        return before == "compare" and _word(_at(statement, i - 2)) == "option"
    if word == "binary":
        return (before == "compare" and _word(_at(statement, i - 2)) == "option") or (
            before == "for" and _in_open_statement(statement, i)
        )
    if word in ("output", "append", "random"):
        return before == "for" and _in_open_statement(statement, i)
    if word == "read":
        return before in ("access", "lock") and _in_open_statement(statement, i)
    if word in ("ptrsafe", "lib", "alias"):
        return _in_declare_header(statement, i)
    if word == "step":
        return _in_for_header(statement, i) and _ends_operand(prev)
    if word == "error":
        return (
            before == "on"
            or (before == "local" and _word(_at(statement, i - 2)) == "on")
            or _starts_statement(statement, i)
        )
    return True


def _starts_statement(statement: Sequence[VbaToken], i: int) -> bool:
    """At the start of a statement: first, after a line number, or after a
    single-line If's Then or Else."""
    if i == 0:
        return True
    prev = statement[i - 1]
    if i == 1 and prev.kind is TokenKind.INTEGER_LITERAL and _LINE_NUMBER_RE.match(prev.raw_text):
        return True
    return prev.kind is TokenKind.KEYWORD and _word(prev) in ("then", "else")


def _in_open_statement(statement: Sequence[VbaToken], i: int) -> bool:
    """An `Open` statement runs before ``i``: the reserved word, not a member
    called Open."""
    for j in range(i - 1, -1, -1):
        token = statement[j]
        if _word(token) == "open" and token.kind is TokenKind.KEYWORD and not _is_member_access(_at(statement, j - 1)):
            return True
    return False


def _in_declare_header(statement: Sequence[VbaToken], i: int) -> bool:
    """Inside a Declare header, before its parameter list."""
    declare = False
    for token in statement[:i]:
        if token.kind is TokenKind.PUNCTUATION and token.raw_text == "(":
            return False
        if _word(token) == "declare" and token.kind is TokenKind.KEYWORD:
            declare = True
    return declare


def _in_for_header(statement: Sequence[VbaToken], i: int) -> bool:
    """After the For and To of a counted For header, outside any parentheses."""
    depth = 0
    saw_to = False
    for j in range(i - 1, -1, -1):
        token = statement[j]
        if token.kind is TokenKind.PUNCTUATION and token.raw_text == ")":
            depth += 1
        elif token.kind is TokenKind.PUNCTUATION and token.raw_text == "(":
            if depth == 0:
                return False
            depth -= 1
        elif depth == 0 and token.kind is TokenKind.KEYWORD and _word(token) == "to":
            saw_to = True
        elif depth == 0 and token.kind is TokenKind.KEYWORD and _word(token) == "for":
            return saw_to and _word(_at(statement, j - 1)) != "exit"
    return False


def _ends_operand(token: VbaToken | None) -> bool:
    """The token can end an operand, so what follows it is not the next part of one."""
    if token is None:
        return False
    if token.kind in _OPERAND_KINDS:
        return True
    if token.kind is TokenKind.PUNCTUATION:
        return token.raw_text == ")"
    if token.kind is TokenKind.KEYWORD:
        return _word(token) in _OPERAND_KEYWORDS
    return False
