"""What a worksheet function refuses when its arguments are literals (XLIDE issue
#442).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/worksheetFunctionArguments.ts.

Through WorksheetFunction a worksheet error value is raised as Run-time error
1004; through Application the same call returns the error value and runs, so
only WorksheetFunction is judged. Each case measured in Excel 16.0 (build
20430, 2026-10-02):

 - Ln(0), Ln(-1) and Log10(0): no logarithm of a number at or below 0.
 - Power(-1, 0.5): a negative base to a fractional power; Power(0, -1).
 - Sum("abc"), Max("abc"), Average("x", 1): a string that is no number.
   Sum("5") is 5.
 - Dec2Bin(1000) and Dec2Bin(-513): outside -512 to 511.
 - Large(Array(1, 2), 3) and Small(Array(1, 2), 0): k outside 1 to Count.
 - Index(Array(1, 2), 3): past the last element.
 - Match("zzz", Array("a", "b"), 0): an exact match finds nothing. Text
   compares without case: Match("B", ...) is 2.
"""

from __future__ import annotations

import json
import math
import re
from collections.abc import Sequence

from ...constants.integer_constant_expression import parse_vba_integer_literal
from ...js_compat import JS_WHITESPACE, js_number, js_number_to_string
from ...lexer.token_helpers import split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ..call_extraction import string_literal_value
from ..walker import match_paren_from, token_text

# A number or a string literal's text.
Literal = int | float | str

# JavaScript's `\s`, `\d` and `$`: whitespace of its own set, ASCII digits, and
# no match before a trailing newline.
_NUMERIC_TEXT = re.compile(
    "^[" + JS_WHITESPACE + r"]*[-+]?([0-9]+\.?[0-9]*|\.[0-9]+)([eE][-+]?[0-9]+)?["
    + JS_WHITESPACE
    + r"]*\Z"
)
_FLOAT_SUFFIX = re.compile(r"[!#@]\Z")


def _is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _is_integer(value: int | float) -> bool:
    return isinstance(value, int) or (math.isfinite(value) and value.is_integer())


def _literal(arg: Sequence[VbaToken] | None) -> Literal | None:
    """A number or string literal, a number optionally signed."""
    if arg is None:
        return None
    toks = [tok for tok in arg if tok.kind is not TokenKind.COMMENT]
    if len(toks) == 1 and toks[0].kind is TokenKind.STRING_LITERAL:
        return string_literal_value(toks[0].raw_text)
    sign: int | None = (
        (-1 if toks[0].raw_text == "-" else 1)
        if len(toks) == 2 and (toks[0].raw_text == "-" or toks[0].raw_text == "+")
        else None
    )
    number = (toks[0] if len(toks) > 0 else None) if sign is None else toks[1]
    if len(toks) != (1 if sign is None else 2) or number is None:
        return None
    value: int | float | None
    if number.kind is TokenKind.INTEGER_LITERAL:
        value = parse_vba_integer_literal(number.raw_text)
    elif number.kind is TokenKind.FLOAT_LITERAL:
        value = js_number(_FLOAT_SUFFIX.sub("", number.raw_text))
    else:
        value = None
    if value is None or not math.isfinite(value):
        return None
    return (sign if sign is not None else 1) * value


def _array_literal(arg: Sequence[VbaToken] | None) -> list[Literal] | None:
    """The literals of `Array(...)`, each a number or string, or None."""
    if arg is None:
        return None
    toks = [tok for tok in arg if tok.kind is not TokenKind.COMMENT]
    if (
        len(toks) == 0
        or token_text(toks[0]) != "array"
        or len(toks) < 2
        or toks[1].raw_text != "("
        or match_paren_from(toks, 1) != len(toks) - 1
    ):
        return None
    items = (
        []
        if len(toks) == 3
        else [_literal(group) for group in split_top_level_token_groups(toks, 2, ",", len(toks) - 1)]
    )
    out: list[Literal] = []
    for item in items:
        if item is None:
            return None
        out.append(item)
    return out


def _is_numeric_text(text: str) -> bool:
    return _NUMERIC_TEXT.search(text) is not None


def _json(text: str) -> str:
    """JSON.stringify of a string."""
    return json.dumps(text, ensure_ascii=False)


def _values_word(count: int) -> str:
    return "value" if count == 1 else "values"


def worksheet_function_refusal(name: str, args: Sequence[Sequence[VbaToken]]) -> str | None:
    """Why the worksheet function refuses these literal arguments, or None.
    `name` is the function's name, lowercased."""
    values = [_literal(arg) for arg in args]
    a = values[0] if len(values) > 0 else None
    b = values[1] if len(values) > 1 else None
    if name == "ln" or name == "log10":
        if isinstance(a, (int, float)) and a <= 0:
            fn = "Ln" if name == "ln" else "Log10"
            return f"{fn} has no value at {js_number_to_string(a)}: a logarithm takes a number above 0"
        return None
    if name == "power":
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            if a < 0 and not _is_integer(b):
                return (
                    f"Power({js_number_to_string(a)}, {js_number_to_string(b)}) raises a negative "
                    "number to a fractional power"
                )
            if a == 0 and b < 0:
                return f"Power(0, {js_number_to_string(b)}) divides by zero"
        return None
    if name in ("sum", "max", "min", "average", "product"):
        text = next(
            (value for value in values if isinstance(value, str) and not _is_numeric_text(value)),
            None,
        )
        return f"{_json(text)} is no number for the worksheet function to take" if text is not None else None
    if name == "dec2bin":
        if isinstance(a, (int, float)) and _is_integer(a) and (a > 511 or a < -512):
            return f"Dec2Bin takes -512 to 511, and {js_number_to_string(a)} is outside that"
        return None
    if name == "large" or name == "small":
        items = _array_literal(args[0] if len(args) > 0 else None)
        if items is not None and isinstance(b, (int, float)) and _is_integer(b) and (b < 1 or b > len(items)):
            return (
                f"the array holds {len(items)} {_values_word(len(items))}, so k = "
                f"{js_number_to_string(b)} names none"
            )
        return None
    if name == "index":
        items = _array_literal(args[0] if len(args) > 0 else None)
        if (
            items is not None
            and len(args) == 2
            and isinstance(b, (int, float))
            and _is_integer(b)
            and b > len(items)
        ):
            return (
                f"the array holds {len(items)} {_values_word(len(items))}, so index "
                f"{js_number_to_string(b)} is past the end"
            )
        return None
    if name == "match":
        items = _array_literal(args[1] if len(args) > 1 else None)
        third = values[2] if len(values) > 2 else None
        exact = _is_number(third) and third == 0
        if items is None or not exact or a is None or any(_is_number(item) != _is_number(a) for item in items):
            return None

        def same(item: Literal) -> bool:
            if isinstance(a, str) and isinstance(item, str):
                return a.lower() == item.lower()
            return a == item

        if any(same(item) for item in items):
            return None
        shown = _json(a) if isinstance(a, str) else js_number_to_string(a)
        return f"an exact Match finds {shown} nowhere in the array"
    return None
