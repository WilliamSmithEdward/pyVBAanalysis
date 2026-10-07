"""Pure, host-free VBA type-name helpers.

Ported from the host-free core of
xlide_vscode/src/analyzer/diagnostics/typeInference.ts: type-name normalization
and classification, plus numeric-literal bounds. They have no host or completion
dependency, so every layer can use them; expression typing lives in
diagnostics/argument_inference.py.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

_TRAILING_PARENS = re.compile(r"\s*\(\s*\)\s*$")

_NUMERIC_TYPES: frozenset[str] = frozenset(
    {"byte", "integer", "long", "longlong", "longptr", "single", "double", "currency", "decimal"}
)


def normalize_type(type_name: str | None) -> str | None:
    """Normalize a declared type name: strip a trailing (), then
    trim and lowercase. Returns None for an empty/missing name."""
    if not type_name:
        return None
    stripped = _TRAILING_PARENS.sub("", type_name)
    return stripped.strip().lower()


def is_numeric_type(type_name: str) -> bool:
    """True when the normalized type name is one of VBA's numeric types (Byte/Integer/Long/LongLong/LongPtr/Single/Double/Currency/Decimal)."""
    return type_name in _NUMERIC_TYPES


def is_known_scalar_type(type_name: str) -> bool:
    """True when the normalized type name is a known scalar: String, Boolean, Date, or any numeric type."""
    return (
        type_name == "string"
        or type_name == "boolean"
        or type_name == "date"
        or is_numeric_type(type_name)
    )


def is_string_concatenation_operand_type(type_name: str) -> bool:
    return (
        type_name == "string"
        or type_name == "boolean"
        or type_name == "date"
        or is_numeric_type(type_name)
    )


def is_provably_non_numeric_string(value: str) -> bool:
    """Whether no locale converts the string to a number (see string_conversion)."""
    # Function-local: diagnostics/ imports this module while it initializes.
    from ..diagnostics.string_conversion import is_invalid_numeric_string

    return is_invalid_numeric_string(value)


@dataclass(frozen=True, slots=True)
class NumericBounds:
    min: int
    max: int
    label: str


def numeric_literal_bounds(expected: str) -> NumericBounds | None:
    """Inclusive overflow bounds for a numeric type, or None when not range-checked.

    Only Byte/Integer/Long/Currency are bounded; LongLong/LongPtr are omitted
    because every reachable safe-integer literal already fits them.
    """
    if expected == "byte":
        return NumericBounds(0, 255, "Byte")
    if expected == "integer":
        return NumericBounds(-32768, 32767, "Integer")
    if expected == "long":
        return NumericBounds(-2147483648, 2147483647, "Long")
    if expected == "currency":
        return NumericBounds(-922337203685477, 922337203685477, "Currency")
    return None
