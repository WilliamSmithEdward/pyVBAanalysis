"""Shared VBA integer-constant expression evaluation.

Ported from xlide_vscode/src/analyzer/constants/integerConstantExpression.ts. One
evaluator for declared-constant and enum-member integer expressions, used by both
the project symbol graph (exported constant surfaces) and the diagnostics engine
(fixed-length strings, runtime argument bounds, division by zero). A single copy
guarantees the project-visible constant values and the diagnostics rules can never
disagree on the same expression.

The grammar is deliberately conservative: +, -, * (binary and unary +/-),
parentheses, integer literals (decimal, &H hex, &O octal, with an optional %/&/^
type suffix), bare constant names, and Module.Constant qualified names. Anything
else evaluates to None so callers never guess.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Generator, Mapping
from typing import Any, Protocol, cast

from ..js_compat import js_number, js_number_to_string, js_trim
from ..lexer.token_helpers import token_name
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize
from ..parser.expression_stack import run_expression

# JavaScript Number.MAX_SAFE_INTEGER; mirrors the Number.isSafeInteger gating that
# the TypeScript source uses to reject magnitudes that lose precision.
_MAX_SAFE_INTEGER = 2**53 - 1

_DECIMAL_RE = re.compile(r"^\d+$")
_HEX_RE = re.compile(r"^&[hH]([0-9A-Fa-f]+)$")
_OCTAL_RE = re.compile(r"^&[oO]([0-7]+)$")


class IntegerConstantLookup(Protocol):
    """Lookup of numeric constant values by lowercased (possibly qualified) name.

    Most supported forms are integers; Val can also produce a fraction.
    """

    # Positional-only so a plain dict/Mapping of resolved constants satisfies the
    # protocol (mirrors the ReadonlyMap the TypeScript rules pass in).
    def get(self, name: str, /) -> float | None: ...


def _is_safe_integer(value: int) -> bool:
    return -_MAX_SAFE_INTEGER <= value <= _MAX_SAFE_INTEGER


def parse_decimal_integer_literal(raw: str) -> int | None:
    """Parses an unsigned decimal integer literal, rejecting unsafe magnitudes."""
    if _DECIMAL_RE.match(raw) is None:
        return None
    value = int(raw)
    return value if _is_safe_integer(value) else None


def parse_vba_integer_literal(raw: str) -> int | None:
    """Parses a VBA integer literal (decimal, &H, &O; optional %/&/^ suffix)."""
    trimmed = js_trim(raw)
    suffix = trimmed[-1] if trimmed and trimmed[-1] in "%&^" else None
    text = trimmed[:-1] if suffix else trimmed
    hex_match = _HEX_RE.match(text)
    octal_match = None if hex_match else _OCTAL_RE.match(text)
    matched = hex_match or octal_match
    if matched is not None:
        digits = matched.group(1)
        value = int(digits, 16 if hex_match else 8)
        if not _is_safe_integer(value):
            return None
        # A hex or octal literal is the signed value of its bits (MS-VBAL 3.3.2,
        # measured in Excel 16.0, XLIDE issue #141): up to four hex digits it is
        # an Integer, so &H8000 is -32768 and &HFFFF is -1; up to eight it is a
        # Long, so &H80000000 is -2147483648. An `&` suffix makes a short literal
        # a Long (&HFFFF& is 65535) but still wraps at 32 bits; a `^` suffix
        # (LongLong) does not wrap here.
        fits16 = len(digits) <= 4 if hex_match else value <= 0xFFFF
        if suffix != "&" and suffix != "^" and fits16 and value > 0x7FFF:
            return value - 0x10000
        if suffix != "^" and 0x7FFFFFFF < value <= 0xFFFFFFFF:
            return value - 0x100000000
        return value
    return parse_decimal_integer_literal(text)


def bankers_round(value: float) -> float:
    """VBA's rounding to a whole number: banker's rounding at .5."""
    if not math.isfinite(value):
        return value  # JavaScript's Math.floor passes Infinity and NaN through
    floor = math.floor(value)
    fraction = value - floor
    if fraction > 0.5:
        return floor + 1
    if fraction < 0.5:
        return floor
    return floor if floor % 2 == 0 else floor + 1


def safe_integer(value: float) -> int | None:
    """Clamps an arithmetic result to a safe integer; None when out of range
    (JavaScript's Number.isSafeInteger, so a fraction is None too)."""
    if isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():
            return None
        value = int(value)
    return value if _is_safe_integer(value) else None


def enum_member_raw_expression(explicit_raw: str | None, previous_name: str | None) -> str:
    """Raw value expression of an enum member.

    The explicit initializer when present, otherwise the implicit MS-VBAL rule of
    previous member + 1 (the first member defaults to 0).
    """
    if explicit_raw is not None:
        return explicit_raw
    return f"{previous_name} + 1" if previous_name else "0"


def evaluate_integer_constant_expression(raw: str, constants: IntegerConstantLookup) -> float | None:
    """Evaluate a supported numeric constant expression (Val may be fractional)."""
    return _IntegerConstantExpressionParser(raw, constants).parse()


# Match upstream's descent budget, even though readers use an explicit stack.
# Past this depth return None so callers never guess.
_MAX_RECURSION_DEPTH = 300


class _NotACall:
    """Returned where the current token starts no rounding call."""

    __slots__ = ()


_NOT_A_CALL = _NotACall()

# The calls that make a number whole, with the range each result must fit.
_ROUNDING_CALLS: dict[str, tuple[int, int]] = {
    "cint": (-32768, 32767),
    "clng": (-2147483648, 2147483647),
    "cbyte": (0, 255),
    "int": (-_MAX_SAFE_INTEGER, _MAX_SAFE_INTEGER),
    "fix": (-_MAX_SAFE_INTEGER, _MAX_SAFE_INTEGER),
    "round": (-_MAX_SAFE_INTEGER, _MAX_SAFE_INTEGER),
}

# The logical operators, loosest first.
_LOGICAL_LEVELS: tuple[str, ...] = ("xor", "or", "and")

_FLOAT_SUFFIX_RE = re.compile(r"[!#@]$")


def _string_literal_text(raw: str) -> str:
    """The text of a string literal token: its quotes off, a doubled quote one."""
    body = raw[1:-1] if raw.endswith('"') and len(raw) > 1 else raw[1:]
    return body.replace('""', '"')


def _is_long(value: float) -> bool:
    if isinstance(value, float) and not value.is_integer():
        return False
    return -2147483648 <= value <= 2147483647


def _whole(value: float) -> float:
    """An integral float as an int, the way a JavaScript number prints and keys."""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


class _IntegerConstantExpressionParser:
    # Values are JavaScript numbers: whole numbers as int, and a fraction (from
    # Val or a Double literal inside a rounding call) as float.

    def __init__(self, raw: str, constants: IntegerConstantLookup) -> None:
        self._constants = constants
        self._tokens: list[VbaToken] = [
            t
            for t in tokenize(raw)
            if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
        ]
        self._index = 0
        self._depth = 0
        # Inside a call's arguments, where True and False pass as -1 and 0.
        self._in_arguments = 0

    def parse(self) -> float | None:
        try:
            return run_expression(self._parse())
        except RecursionError:
            # Caller-provided constant lookups may still recurse.
            return None

    def _parse(self) -> Generator[Any, Any, float | None]:
        if not self._tokens:
            return None
        value = cast("float | None", (yield self._expression()))
        if value is None or self._current() is not None:
            return None
        # Val can return a fraction. Preserve it for a later rounding call or
        # another constant, as upstream's JavaScript number map does.
        return _whole(value)

    def _lookup(self, name: str) -> Generator[Any, Any, float | None]:
        if isinstance(self._constants, _GeneratorLookup):
            return cast("float | None", (yield self._constants.resolve(name)))
        return self._constants.get(name)

    def _expression(self) -> Generator[Any, Any, float | None]:
        # Depth guard: untrusted Const text can nest arbitrarily deep; bail to
        # None rather than overflowing the stack.
        self._depth += 1
        try:
            if self._depth > _MAX_RECURSION_DEPTH:
                return None
            return (cast("float | None", (yield self._logical(0))))
        finally:
            self._depth -= 1

    def _logical(self, level: int) -> Generator[Any, Any, float | None]:
        """Xor, Or and And, loosest first, then Not, over whole numbers within
        the Long range: ``Const K0 = 15 And 255`` is 15 (XLIDE issue #496,
        measured in Excel 16.0)."""
        if level == len(_LOGICAL_LEVELS):
            count = 0
            while self._accept_word("not"):
                count += 1
            operand = cast("float | None", (yield self._expression_inner()))
            for _ in range(count):
                operand = None if operand is None or not _is_long(operand) else ~int(operand)
            return operand
        word = _LOGICAL_LEVELS[level]
        value = (cast("float | None", (yield self._logical(level + 1))))
        while value is not None and self._accept_word(word):
            right = (cast("float | None", (yield self._logical(level + 1))))
            if right is None or not _is_long(value) or not _is_long(right):
                return None
            left_int = int(value)
            right_int = int(right)
            if word == "and":
                value = left_int & right_int
            elif word == "or":
                value = left_int | right_int
            else:
                value = left_int ^ right_int
        return value

    def _expression_inner(self) -> Generator[Any, Any, float | None]:
        value = (cast("float | None", (yield self._modulo())))
        while value is not None:
            if self._accept("+"):
                right = (cast("float | None", (yield self._modulo())))
                value = None if right is None else safe_integer(value + right)
                continue
            if self._accept("-"):
                right = (cast("float | None", (yield self._modulo())))
                value = None if right is None else safe_integer(value - right)
                continue
            break
        return value

    def _modulo(self) -> Generator[Any, Any, float | None]:
        """Mod binds below ``\\``, and ``\\`` below ``*``: ``50 Mod 7 + 10`` is 11."""
        value = (cast("float | None", (yield self._integer_division())))
        while value is not None and self._accept_word("mod"):
            right = (cast("float | None", (yield self._integer_division())))
            # JavaScript's % keeps the dividend's sign, as math.fmod does.
            value = None if right is None or right == 0 else safe_integer(math.fmod(value, right))
        return value

    def _integer_division(self) -> Generator[Any, Any, float | None]:
        value = (cast("float | None", (yield self._term())))
        while value is not None and self._accept("\\"):
            right = (cast("float | None", (yield self._term())))
            value = None if right is None or right == 0 else safe_integer(math.trunc(value / right))
        return value

    def _term(self) -> Generator[Any, Any, float | None]:
        value = (cast("float | None", (yield self._factor())))
        while value is not None:
            if not self._accept("*"):
                break
            right = (cast("float | None", (yield self._factor())))
            value = None if right is None else safe_integer(value * right)
        return value

    def _factor(self) -> Generator[Any, Any, float | None]:
        # Depth guard: unary +/- chains and nested parens recurse through factor;
        # bail to None once the ceiling is hit (see _expression()).
        self._depth += 1
        try:
            if self._depth > _MAX_RECURSION_DEPTH:
                return None
            return (cast("float | None", (yield self._factor_inner())))
        finally:
            self._depth -= 1

    def _factor_inner(self) -> Generator[Any, Any, float | None]:
        if self._accept("+"):
            return (cast("float | None", (yield self._factor())))
        if self._accept("-"):
            value = (cast("float | None", (yield self._factor())))
            return None if value is None else safe_integer(-value)
        if self._accept("("):
            value = (cast("float | None", (yield self._expression())))
            return value if value is not None and self._accept(")") else None
        token = self._current()
        if token is None:
            return None
        if token.kind is TokenKind.INTEGER_LITERAL:
            self._index += 1
            return parse_vba_integer_literal(token.raw_text)
        # True is -1 as a number, False 0: `F(False)` (XLIDE issue #562). Only
        # there: a Byte Const of True holds 255.
        word = token.raw_text.lower()
        previous = self._peek(-1)
        if (
            self._in_arguments > 0
            and word in ("true", "false")
            and (previous is None or previous.raw_text != ".")
        ):
            self._index += 1
            return -1 if word == "true" else 0
        qualified = self._qualified_name()
        if qualified:
            return cast("float | None", (yield self._lookup(qualified.lower())))
        rounded = (cast("float | None | _NotACall", (yield self._rounding_call())))
        if not isinstance(rounded, _NotACall):
            return rounded
        name = token_name(token)
        if name:
            self._index += 1
            following = self._peek(0)
            # `F()`: a lookup may know a Function's result as `f()` (XLIDE issue #448).
            if following is not None and following.raw_text == "(":
                after = self._peek(1)
                if after is not None and after.raw_text == ")":
                    self._index += 2
                    return cast("float | None", (yield self._lookup(f"{name.lower()}()")))
                # `F(-1)`: a call with whole-number arguments is `f(-1)` to a
                # lookup (XLIDE issue #562).
                before = self._peek(-2)
                if before is None or before.raw_text != ".":
                    self._index += 1
                    self._in_arguments += 1
                    args: list[float | None] = [(cast("float | None", (yield self._expression())))]
                    while self._accept(","):
                        args.append((cast("float | None", (yield self._expression()))))
                    self._in_arguments -= 1
                    if not self._accept(")") or any(arg is None for arg in args):
                        return None
                    key = ",".join(js_number_to_string(arg) for arg in args if arg is not None)
                    return cast("float | None", (yield self._lookup(f"{name.lower()}({key})")))
            return cast("float | None", (yield self._lookup(name.lower())))
        return None

    def _rounding_call(self) -> Generator[Any, Any, float | None | _NotACall]:
        """``CInt(3.5)``, ``Int(-0.1)``, ``Fix(-0.9)``, ``Round(0.5)``: a
        conversion or rounding of a number, which comes out whole (XLIDE issue
        #286, measured in Excel 16.0). CInt, CLng and Round round half to even,
        Int down and Fix toward zero. _NOT_A_CALL where the current token starts
        no such call; None where it does and the argument is not known or the
        result does not fit."""
        current_name = token_name(self._current())
        word = current_name.lower() if current_name is not None else None
        previous = self._peek(-1)
        after_dot = previous is not None and previous.raw_text == "."
        # `Val("0,5")` is 0 in every locale (XLIDE issue #703).
        open_paren = self._peek(1)
        literal = self._peek(2)
        close_paren = self._peek(3)
        if (
            word == "val"
            and open_paren is not None
            and open_paren.raw_text == "("
            and literal is not None
            and literal.kind is TokenKind.STRING_LITERAL
            and close_paren is not None
            and close_paren.raw_text == ")"
            and not after_dot
        ):
            # Imported here: string_conversion imports this module.
            from ..diagnostics.string_conversion import val_prefix_value

            val = val_prefix_value(_string_literal_text(literal.raw_text))
            if isinstance(val, (int, float)):
                self._index += 4
                return val
            return _NOT_A_CALL
        range_ = _ROUNDING_CALLS.get(word) if word else None
        if range_ is None or open_paren is None or open_paren.raw_text != "(" or after_dot:
            return _NOT_A_CALL
        start = self._index
        self._index += 2
        argument = self._index
        negative = self._accept("-")
        tok = self._current()
        value: float | None
        following = self._peek(1)
        closes_next = following is not None and following.raw_text == ")"
        if tok is not None and tok.kind is TokenKind.STRING_LITERAL and not negative and closes_next:
            # `CInt("(5)")`, `CLng("&H0")`: a string every locale reads alike
            # (XLIDE issue #703, measured in Excel 16.0).
            from ..diagnostics.string_conversion import numeric_string_verdict

            self._index += 1
            verdict = numeric_string_verdict(_string_literal_text(tok.raw_text))
            verdict_value = getattr(verdict, "value", None)
            value = (
                verdict_value
                if getattr(verdict, "kind", None) == "number" and isinstance(verdict_value, (int, float))
                else None
            )
        elif tok is not None and tok.kind is TokenKind.FLOAT_LITERAL and closes_next:
            self._index += 1
            read = js_number(_FLOAT_SUFFIX_RE.sub("", tok.raw_text, count=1))
            value = (-read if negative else read) if math.isfinite(read) else None
        else:
            self._index = argument
            value = (cast("float | None", (yield self._expression())))
        if value is None or not self._accept(")"):
            self._index = start
            return _NOT_A_CALL
        if not math.isfinite(value):
            return None
        whole: float
        if word == "int":
            whole = math.floor(value)
        elif word == "fix":
            whole = math.trunc(value)
        else:
            whole = bankers_round(value)
        return int(whole) if range_[0] <= whole <= range_[1] else None

    def _qualified_name(self) -> str | None:
        qualifier = token_name(self._current())
        dot = self._peek(1)
        member = token_name(self._peek(2))
        if not qualifier or dot is None or dot.raw_text != "." or not member:
            return None
        self._index += 3
        return f"{qualifier}.{member}"

    def _current(self) -> VbaToken | None:
        return self._peek(0)

    def _peek(self, offset: int) -> VbaToken | None:
        i = self._index + offset
        return self._tokens[i] if 0 <= i < len(self._tokens) else None

    def _accept(self, raw: str) -> bool:
        current = self._current()
        if current is None or current.raw_text != raw:
            return False
        self._index += 1
        return True

    def _accept_word(self, word: str) -> bool:
        token = self._current()
        if token is None or token.kind is not TokenKind.KEYWORD or token.raw_text.lower() != word:
            return False
        self._index += 1
        return True


class _GeneratorLookup:
    """Resume dependency lookups on the same explicit expression stack."""

    __slots__ = ("_fn",)

    def __init__(self, fn: Callable[[str], Generator[Any, Any, float | None]]) -> None:
        self._fn = fn

    def get(self, name: str) -> float | None:
        return run_expression(self.resolve(name))

    def resolve(self, name: str) -> Generator[Any, Any, float | None]:
        return self._fn(name)


def resolve_raw_integer_constants(
    raw_constants: Mapping[str, str | None],
    base: Mapping[str, float | None] | None = None,
) -> dict[str, float | None]:
    """Resolves raw constant expressions to integer values, memoized, cycle-safe.

    raw_constants maps a lowercased (possibly qualified) name to its raw expression
    text (None for an ambiguous duplicate). Names absent from raw_constants fall
    back to the optional base map of already-resolved values; the returned map only
    contains raw_constants keys. A reference cycle resolves to None.
    """
    base_map: Mapping[str, float | None] = {} if base is None else base
    resolved: dict[str, float | None] = {}
    resolving: set[str] = set()

    def resolve(name: str) -> Generator[Any, Any, float | None]:
        key = name.lower()
        if key in resolved:
            return resolved[key]
        if key not in raw_constants:
            return base_map.get(key)
        if key in resolving:
            resolved[key] = None
            return None
        raw = raw_constants[key]
        if raw is None:
            resolved[key] = None
            return None
        resolving.add(key)
        parser = _IntegerConstantExpressionParser(raw, _GeneratorLookup(resolve))
        value = cast("float | None", (yield parser._parse()))
        resolving.discard(key)
        resolved[key] = value
        return value

    for key in raw_constants:
        run_expression(resolve(key))
    return resolved
