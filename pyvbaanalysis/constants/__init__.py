"""Shared constant-expression evaluation (lexer-level, parser-independent)."""

from .integer_constant_expression import (
    IntegerConstantLookup,
    enum_member_raw_expression,
    evaluate_integer_constant_expression,
    parse_decimal_integer_literal,
    parse_vba_integer_literal,
    resolve_raw_integer_constants,
    safe_integer,
)

__all__ = [
    "IntegerConstantLookup",
    "enum_member_raw_expression",
    "evaluate_integer_constant_expression",
    "parse_decimal_integer_literal",
    "parse_vba_integer_literal",
    "resolve_raw_integer_constants",
    "safe_integer",
]
