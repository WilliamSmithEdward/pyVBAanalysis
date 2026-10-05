"""Ported from xlide_vscode/src/analyzer/diagnostics/nullOperators.ts.

Whether an expression gives Null through its operators (XLIDE issues #324 and
#556, each measured in Excel 16.0): arithmetic, `+`, unary minus, Not, a
comparison and Abs give Null when an operand is Null, and so do Xor and Eqv;
And, Or and Imp give a value when the other side decides it; `&` never does.
"""

from __future__ import annotations

import math
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from typing import Literal

from ..js_compat import js_number
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
    return _operator_yields_null(toks, holds_null)


@dataclass
class _NullFrame:
    ranges: list[tuple[int, int]]
    mode: str
    next: int = 1
    left: float | Literal["null"] | None = None


def _operator_yields_null(toks: Sequence[VbaToken], holds_null: Callable[[VbaToken], bool]) -> bool:
    if len(toks) == 1:
        return holds_null(toks[0])
    if len(toks) == 2 and token_text(toks[0]) in ("not", "-"):
        return holds_null(toks[1])
    # Shared token windows and explicit continuations keep deep wrappers safe.
    parens: dict[int, int] = {}
    pending: list[int] = []
    for i, tok in enumerate(toks):
        if tok.raw_text == "(":
            pending.append(i)
        elif tok.raw_text == ")" and pending:
            parens[pending.pop()] = i
    frames: list[_NullFrame] = []
    current = (0, len(toks))
    result: bool | None = None
    while True:
        if result is None:
            start, end = current
            while True:
                if end - start >= 2 and toks[start].raw_text == "(" and parens.get(start) == end - 1:
                    start += 1
                    end -= 1
                if end - start == 1:
                    result = holds_null(toks[start])
                    break
                if end <= start:
                    result = False
                    break
                head = token_text(toks[start])
                if head in ("-", "not"):
                    start += 1
                    continue
                if head == "abs" and toks[start + 1].raw_text == "(" and parens.get(start + 1) == end - 1:
                    start += 2
                    end -= 1
                    continue
                break
            if result is None:
                ranges: list[tuple[int, int]] = []
                operators: list[str] = []
                segment = start
                depth = 0
                i = start
                while i < end:
                    tok = toks[i]
                    close = parens.get(i)
                    if depth == 0 and tok.raw_text == "(" and close is not None and close < end:
                        i = close + 1
                        continue
                    depth += 1 if tok.raw_text == "(" else -1 if tok.raw_text == ")" else 0
                    word = tok.raw_text if tok.kind is TokenKind.OPERATOR else token_text(tok)
                    if depth == 0 and i > segment and word in _NULL_PROPAGATING and tok.kind is not TokenKind.STRING_LITERAL:
                        ranges.append((segment, i))
                        operators.append(word)
                        segment = i + 1
                    i += 1
                ranges.append((segment, end))
                if not operators or "&" in operators or any(a == b for a, b in ranges):
                    result = False
                else:
                    logical = next((op for op in operators if op in ("and", "or", "imp")), None)
                    mode = (logical if len(operators) == 1 else "every") if logical else "some"
                    frames.append(_NullFrame(ranges, mode))
                    current = ranges[0]
                    continue
        while True:
            if not frames:
                return result
            frame = frames[-1]
            if frame.mode in ("some", "every"):
                if (result if frame.mode == "some" else not result) or frame.next == len(frame.ranges):
                    frames.pop()
                    continue
                current = frame.ranges[frame.next]
                frame.next += 1
                result = None
                break
            a, b = frame.ranges[frame.next - 1]
            value: float | Literal["null"] | None = "null" if result else _literal_number(toks[a:b])
            if frame.next == 1:
                frame.left = value
                frame.next = 2
                current = frame.ranges[1]
                result = None
                break

            def decided(other: float | Literal["null"] | None, on_left: bool) -> bool:
                if other == "null":
                    return True
                if other is None:
                    return False
                if frame.mode == "and":
                    return other != 0
                if frame.mode == "or":
                    return other == 0
                return other != 0 if on_left else other == 0

            result = (frame.left == "null" and decided(value, False)) or (value == "null" and decided(frame.left, True))
            frames.pop()
