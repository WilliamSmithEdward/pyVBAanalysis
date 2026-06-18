"""M1: integer-constant expression evaluation (integerConstantExpression.ts parity)."""

from __future__ import annotations

import pytest

from pyvbaanalysis.constants import (
    enum_member_raw_expression,
    evaluate_integer_constant_expression,
    parse_decimal_integer_literal,
    parse_vba_integer_literal,
    resolve_raw_integer_constants,
    safe_integer,
)

_MAX_SAFE = 2**53 - 1


def _eval(raw: str, constants: dict[str, int | None] | None = None) -> int | None:
    table = constants or {}
    return evaluate_integer_constant_expression(raw, _DictLookup(table))


class _DictLookup:
    def __init__(self, table: dict[str, int | None]) -> None:
        self._table = table

    def get(self, name: str) -> int | None:
        return self._table.get(name)


def test_parse_decimal_integer_literal() -> None:
    assert parse_decimal_integer_literal("42") == 42
    assert parse_decimal_integer_literal("007") == 7
    assert parse_decimal_integer_literal("-1") is None  # unsigned only
    assert parse_decimal_integer_literal("1.5") is None
    assert parse_decimal_integer_literal(str(_MAX_SAFE + 1)) is None


def test_parse_vba_integer_literal_radixes_and_suffix() -> None:
    assert parse_vba_integer_literal("255") == 255
    assert parse_vba_integer_literal("&HFF") == 255
    assert parse_vba_integer_literal("&hff") == 255
    assert parse_vba_integer_literal("&O17") == 15
    assert parse_vba_integer_literal("&o17") == 15
    assert parse_vba_integer_literal("100&") == 100  # type suffix stripped
    assert parse_vba_integer_literal("16%") == 16
    assert parse_vba_integer_literal("&HZZ") is None


def test_safe_integer_bounds() -> None:
    assert safe_integer(0) == 0
    assert safe_integer(_MAX_SAFE) == _MAX_SAFE
    assert safe_integer(_MAX_SAFE + 1) is None
    assert safe_integer(-_MAX_SAFE - 1) is None


@pytest.mark.parametrize(
    ("expr", "expected"),
    [
        ("5", 5),
        ("&HFF", 255),
        ("&O17", 15),
        ("2 + 3", 5),
        ("10 - 4", 6),
        ("2 + 3 * 4", 14),  # precedence: * binds tighter than +
        ("(2 + 3) * 4", 20),
        ("-5", -5),
        ("+7", 7),
        ("- -5", 5),
        ("3 * -2", -6),
        ("", None),
        ("2 +", None),  # trailing operator
        ("2 3", None),  # trailing token
        ("2 / 3", None),  # division is outside the grammar
        ("(1 + 2", None),  # unbalanced
    ],
)
def test_evaluate_grammar(expr: str, expected: int | None) -> None:
    assert _eval(expr) == expected


def test_evaluate_named_and_qualified_constants() -> None:
    table = {"foo": 10, "mod.bar": 3}
    assert _eval("FOO", table) == 10  # case-insensitive
    assert _eval("Foo + 5", table) == 15
    assert _eval("Mod.Bar * 2", table) == 6
    assert _eval("Unknown", table) is None


def test_evaluate_overflow_is_none() -> None:
    assert _eval(f"{_MAX_SAFE} + 1") is None


def test_enum_member_raw_expression() -> None:
    assert enum_member_raw_expression("5", None) == "5"
    assert enum_member_raw_expression(None, "Prev") == "Prev + 1"
    assert enum_member_raw_expression(None, None) == "0"


def test_resolve_raw_integer_constants_chain() -> None:
    raw = {"a": "1", "b": "a + 1", "c": "b * 2"}
    resolved = resolve_raw_integer_constants(raw)
    assert resolved == {"a": 1, "b": 2, "c": 4}


def test_resolve_raw_integer_constants_cycle_is_none() -> None:
    raw = {"a": "b", "b": "a"}
    resolved = resolve_raw_integer_constants(raw)
    assert resolved == {"a": None, "b": None}


def test_resolve_raw_integer_constants_uses_base() -> None:
    raw = {"b": "a + 1"}  # a is not in raw; comes from base
    resolved = resolve_raw_integer_constants(raw, base={"a": 41})
    assert resolved == {"b": 42}


def test_resolve_raw_integer_constants_ambiguous_is_none() -> None:
    raw: dict[str, str | None] = {"a": None, "b": "a + 1"}  # a ambiguous duplicate
    resolved = resolve_raw_integer_constants(raw)
    assert resolved == {"a": None, "b": None}
