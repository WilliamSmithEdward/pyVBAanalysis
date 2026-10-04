"""Ported from xlide_vscode/src/analyzer/diagnostics/nullOperators.ts.

Whether an expression gives Null through its operators (XLIDE issues #324 and
#556, each measured in Excel 16.0): arithmetic, `+`, unary minus, Not, a
comparison and Abs give Null when an operand is Null, and so do Xor and Eqv;
And, Or and Imp give a value when the other side decides it; `&` never does.
"""

from __future__ import annotations

import math
from collections.abc import Callable, Sequence
from typing import Literal

from ..js_compat import js_number
from ..lexer.token_helpers import match_paren_from
from ..lexer.token_helpers import token_word as token_text
from ..lexer.token_kinds import TokenKind, VbaToken
from .call_extraction import unwrap_outer_parens

# The binary operators whose result is Null when an operand is. `&` is here to be refused.
_NULL_PROPAGATING: frozenset[str] = frozenset(
    {"+", "-", "*", "/", "\\", "^", "mod", "=", "<>", "<", ">", "<=", ">=", "and", "or", "xor", "eqv", "imp", "&"}
)

_LITERAL_SUFFIXES = "%&^!#@"


def _literal_number(operand: Sequence[VbaToken]) -> float | None:
    """The number a literal operand is, True as -1 and False as 0: `1`, `-2.5`, `True`."""
    toks = unwrap_outer_parens(list(operand))
    sign = -1 if len(toks) == 2 and toks[0].raw_text == "-" else 1
    tok = (
        toks[0]
        if len(toks) == 1
        else toks[1]
        if len(toks) == 2 and (toks[0].raw_text == "-" or toks[0].raw_text == "+")
        else None
    )
    word = token_text(tok)
    if word == "true" or word == "false":
        return float(sign * (-1 if word == "true" else 0))
    if tok is None or (tok.kind is not TokenKind.INTEGER_LITERAL and tok.kind is not TokenKind.FLOAT_LITERAL):
        return None
    raw = tok.raw_text
    value = js_number(raw[:-1] if raw and raw[-1] in _LITERAL_SUFFIXES else raw)
    return sign * value if math.isfinite(value) else None


def operator_yields_null(toks: Sequence[VbaToken], holds_null: Callable[[VbaToken], bool]) -> bool:
    """Whether the tokens give Null. `holds_null` says whether one token does: the
    literal Null, or a local known to hold it. A single token is asked whole."""
    try:
        return _operator_yields_null(toks, holds_null)
    except RecursionError:
        # Port-only: upstream recurses once per prefix operator and parenthesis
        # on JavaScript's deeper stack. An expression nested past Python's
        # limit is not known to give Null.
        return False


def _operator_yields_null(toks: Sequence[VbaToken], holds_null: Callable[[VbaToken], bool]) -> bool:
    part = unwrap_outer_parens(list(toks))
    if len(part) == 1:
        return holds_null(part[0])
    head = token_text(part[0]) if part else ""
    if head == "-" or head == "not":
        return _operator_yields_null(part[1:], holds_null)
    if (
        head == "abs"
        and len(part) > 1
        and part[1].raw_text == "("
        and match_paren_from(part, 1) == len(part) - 1
    ):
        return _operator_yields_null(part[2:-1], holds_null)
    operands: list[list[VbaToken]] = [[]]
    operators: list[str] = []
    depth = 0
    for tok in part:
        depth += 1 if tok.raw_text == "(" else -1 if tok.raw_text == ")" else 0
        word = tok.raw_text if tok.kind is TokenKind.OPERATOR else token_text(tok)
        current = operands[-1]
        if (
            depth == 0
            and len(current) > 0
            and word in _NULL_PROPAGATING
            and tok.kind is not TokenKind.STRING_LITERAL
        ):
            operators.append(word)
            operands.append([])
        else:
            current.append(tok)
    if len(operators) == 0 or "&" in operators or any(len(operand) == 0 for operand in operands):
        return False
    # And, Or and Imp give a value when the other side decides it (issue
    # #556): Null And 0 is 0, but Null And 1 is Null; 40000 Or Null is 40000,
    # but 0 Or Null is Null; Null Imp 12 is 12 and False Imp Null is True, but
    # Null Imp False is Null. Judged with one operator only.
    logical = next((op for op in operators if op == "and" or op == "or" or op == "imp"), None)
    if logical is not None:
        if len(operators) != 1:
            return all(_operator_yields_null(operand, holds_null) for operand in operands)
        sides: list[float | Literal["null"] | None] = [
            "null" if _operator_yields_null(operand, holds_null) else _literal_number(operand)
            for operand in operands
        ]
        left, right = sides

        def decided(other: float | Literal["null"] | None, other_on_left: bool) -> bool:
            if other == "null":
                return True
            if other is None:
                return False
            if logical == "and":
                return other != 0
            if logical == "or":
                return other == 0
            return other != 0 if other_on_left else other == 0

        return (left == "null" and decided(right, False)) or (right == "null" and decided(left, True))
    return any(_operator_yields_null(operand, holds_null) for operand in operands)
