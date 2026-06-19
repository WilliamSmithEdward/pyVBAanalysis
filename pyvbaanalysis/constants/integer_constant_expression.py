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

import re
from collections.abc import Callable, Mapping
from typing import Protocol

from ..lexer.token_helpers import token_name
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize

# JavaScript Number.MAX_SAFE_INTEGER; mirrors the Number.isSafeInteger gating that
# the TypeScript source uses to reject magnitudes that lose precision.
_MAX_SAFE_INTEGER = 2**53 - 1

_DECIMAL_RE = re.compile(r"^\d+$")
_SUFFIX_RE = re.compile(r"[%&^]$")
_HEX_RE = re.compile(r"^&[hH]([0-9A-Fa-f]+)$")
_OCTAL_RE = re.compile(r"^&[oO]([0-7]+)$")


class IntegerConstantLookup(Protocol):
    """Lookup of integer constant values by lowercased (possibly qualified) name."""

    # Positional-only so a plain dict/Mapping of resolved constants satisfies the
    # protocol (mirrors the ReadonlyMap the TypeScript rules pass in).
    def get(self, name: str, /) -> int | None: ...


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
    text = _SUFFIX_RE.sub("", raw.strip())
    hex_match = _HEX_RE.match(text)
    if hex_match:
        value = int(hex_match.group(1), 16)
        return value if _is_safe_integer(value) else None
    octal_match = _OCTAL_RE.match(text)
    if octal_match:
        value = int(octal_match.group(1), 8)
        return value if _is_safe_integer(value) else None
    return parse_decimal_integer_literal(text)


def safe_integer(value: int) -> int | None:
    """Clamps an arithmetic result to a safe integer; None when out of range."""
    return value if _is_safe_integer(value) else None


def enum_member_raw_expression(explicit_raw: str | None, previous_name: str | None) -> str:
    """Raw value expression of an enum member.

    The explicit initializer when present, otherwise the implicit MS-VBAL rule of
    previous member + 1 (the first member defaults to 0).
    """
    if explicit_raw is not None:
        return explicit_raw
    return f"{previous_name} + 1" if previous_name else "0"


def evaluate_integer_constant_expression(raw: str, constants: IntegerConstantLookup) -> int | None:
    """Evaluates one raw constant expression against already-known constants."""
    return _IntegerConstantExpressionParser(raw, constants).parse()


class _IntegerConstantExpressionParser:
    def __init__(self, raw: str, constants: IntegerConstantLookup) -> None:
        self._constants = constants
        self._tokens: list[VbaToken] = [
            t
            for t in tokenize(raw)
            if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
        ]
        self._index = 0

    def parse(self) -> int | None:
        if not self._tokens:
            return None
        value = self._expression()
        return value if value is not None and self._current() is None else None

    def _expression(self) -> int | None:
        value = self._term()
        while value is not None:
            if self._accept("+"):
                right = self._term()
                value = None if right is None else safe_integer(value + right)
                continue
            if self._accept("-"):
                right = self._term()
                value = None if right is None else safe_integer(value - right)
                continue
            break
        return value

    def _term(self) -> int | None:
        value = self._factor()
        while value is not None:
            if not self._accept("*"):
                break
            right = self._factor()
            value = None if right is None else safe_integer(value * right)
        return value

    def _factor(self) -> int | None:
        if self._accept("+"):
            return self._factor()
        if self._accept("-"):
            value = self._factor()
            return None if value is None else safe_integer(-value)
        if self._accept("("):
            value = self._expression()
            return value if value is not None and self._accept(")") else None
        token = self._current()
        if token is None:
            return None
        if token.kind is TokenKind.INTEGER_LITERAL:
            self._index += 1
            return parse_vba_integer_literal(token.raw_text)
        qualified = self._qualified_name()
        if qualified:
            return self._constants.get(qualified.lower())
        name = token_name(token)
        if name:
            self._index += 1
            return self._constants.get(name.lower())
        return None

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


class _CallableLookup:
    """Adapts a resolve callback to the IntegerConstantLookup get-by-name protocol."""

    __slots__ = ("_fn",)

    def __init__(self, fn: Callable[[str], int | None]) -> None:
        self._fn = fn

    def get(self, name: str) -> int | None:
        return self._fn(name)


def resolve_raw_integer_constants(
    raw_constants: Mapping[str, str | None],
    base: Mapping[str, int | None] | None = None,
) -> dict[str, int | None]:
    """Resolves raw constant expressions to integer values, memoized, cycle-safe.

    raw_constants maps a lowercased (possibly qualified) name to its raw expression
    text (None for an ambiguous duplicate). Names absent from raw_constants fall
    back to the optional base map of already-resolved values; the returned map only
    contains raw_constants keys. A reference cycle resolves to None.
    """
    base_map: Mapping[str, int | None] = {} if base is None else base
    resolved: dict[str, int | None] = {}
    resolving: set[str] = set()

    def resolve(name: str) -> int | None:
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
        value = evaluate_integer_constant_expression(raw, _CallableLookup(resolve))
        resolving.discard(key)
        resolved[key] = value
        return value

    for key in raw_constants:
        resolve(key)
    return resolved
