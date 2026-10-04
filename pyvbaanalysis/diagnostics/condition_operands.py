"""Ported from xlide_vscode/src/analyzer/diagnostics/conditionOperands.ts.

Where a statement reads a bare name as a truth value or a Boolean operand
(XLIDE issue #424): the whole condition of If, ElseIf, Do While/Until, Loop
While/Until and While, a Select Case subject, IIf's first argument, and an
operand of Not, And, Or, Xor, Eqv or Imp. An object, an array or a Variant
holding an array has no value there, and what each raises was measured in
Excel 16.0; the rules that know each kind of name judge it.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from typing import Literal

from ..lexer.token_helpers import token_name
from ..lexer.token_helpers import token_word as token_text
from ..lexer.token_kinds import TokenKind, VbaToken

ConditionForm = Literal["condition", "select", "iif", "not", "logical"]


@dataclass(frozen=True, slots=True)
class ConditionOperand:
    index: int
    form: ConditionForm


_LOGICAL: frozenset[str] = frozenset({"and", "or", "xor", "eqv", "imp"})
_CONDITION_HEADS: tuple[tuple[str, ...], ...] = (
    ("do", "while"),
    ("do", "until"),
    ("loop", "while"),
    ("loop", "until"),
    ("while",),
    ("select", "case"),
)
_STARTING_WORDS: frozenset[str] = frozenset({"then", "if", "elseif", "while", "until", "case", "not", "else"})


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw(tok: VbaToken | None) -> str | None:
    return tok.raw_text if tok is not None else None


def condition_operands(toks: Sequence[VbaToken]) -> list[ConditionOperand]:
    """The bare names a statement's tokens read as a condition or a Boolean operand."""
    out: list[ConditionOperand] = []

    def bare(i: int) -> bool:
        tok = _at(toks, i)
        return (
            token_name(tok) is not None
            and tok is not None
            and tok.kind is not TokenKind.KEYWORD
            and (_raw(_at(toks, i - 1)) or "") not in (".", "!")
            and (_raw(_at(toks, i + 1)) or "") not in ("(", ".", "!", "$")
        )

    head = token_text(_at(toks, 0))
    # `If x Then` and `ElseIf x Then`, a block's or a one-line If's.
    if (head == "if" or head == "elseif") and token_text(_at(toks, 2)) == "then" and bare(1):
        out.append(ConditionOperand(1, "condition"))
    for words in _CONDITION_HEADS:
        if (
            len(toks) == len(words) + 1
            and all(token_text(toks[k]) == word for k, word in enumerate(words))
            and bare(len(words))
        ):
            out.append(ConditionOperand(len(words), "select" if words[0] == "select" else "condition"))

    # An operand ends where the expression does or another Boolean operator starts.
    def ends(tok: VbaToken | None) -> bool:
        return (
            tok is None
            or tok.raw_text in (")", ",")
            or token_text(tok) == "then"
            or token_text(tok) in _LOGICAL
        )

    def starts(tok: VbaToken | None) -> bool:
        return (
            tok is None
            or tok.raw_text in ("(", ",", "=")
            or token_text(tok) in _STARTING_WORDS
            or token_text(tok) in _LOGICAL
        )

    for i in range(1, len(toks)):
        if not bare(i) or any(hit.index == i for hit in out):
            continue
        before = _at(toks, i - 1)
        after = _at(toks, i + 1)
        if token_text(before) == "not" and ends(after):
            out.append(ConditionOperand(i, "not"))
        elif (token_text(before) in _LOGICAL and ends(after)) or (
            token_text(after) in _LOGICAL and starts(before)
        ):
            out.append(ConditionOperand(i, "logical"))
        elif (
            _raw(before) == "("
            and token_text(_at(toks, i - 2)) == "iif"
            and _raw(after) == ","
            and _raw(_at(toks, i - 3)) != "."
        ):
            out.append(ConditionOperand(i, "iif"))
    return out
