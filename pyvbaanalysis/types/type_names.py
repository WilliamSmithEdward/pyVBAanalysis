"""Pure, host-free VBA type-name helpers.

Ported from the host-free core of
xlide_vscode/src/analyzer/diagnostics/typeInference.ts: type-name normalization
and classification, plus numeric-literal bounds. The host/completion-coupled
inference engine lands in type_inference.py (M8); these helpers have no host or
completion dependency, so the rule families can use them now.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

_TRAILING_PARENS = re.compile(r"\s*\(\s*\)\s*$")
_LEADING_VB = re.compile(r"^vb", re.IGNORECASE)

_NUMERIC_TYPES: frozenset[str] = frozenset(
    {"byte", "integer", "long", "longlong", "longptr", "single", "double", "currency", "decimal"}
)


def normalize_type(type_name: str | None) -> str | None:
    """Normalize a declared type name: strip a trailing (), a leading 'vb', then
    trim and lowercase. Returns None for an empty/missing name."""
    if not type_name:
        return None
    stripped = _LEADING_VB.sub("", _TRAILING_PARENS.sub("", type_name))
    return stripped.strip().lower()


def is_numeric_type(type_name: str) -> bool:
    return type_name in _NUMERIC_TYPES


def is_known_scalar_type(type_name: str) -> bool:
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


_HAS_DIGIT = re.compile(r"[0-9]")
_BOOLEAN_STRING = re.compile(r"^(true|false|0|-?1)$", re.IGNORECASE)


def is_provably_non_numeric_string(value: str) -> bool:
    trimmed = value.strip()
    return bool(trimmed) and _HAS_DIGIT.search(trimmed) is None


def is_boolean_string(value: str) -> bool:
    return _BOOLEAN_STRING.match(value.strip()) is not None


def is_known_object_assignment_type(type_name: str | None) -> bool:
    """True when a declared type names an object that Set-binds and supports members.

    Host-free slice (M7): the generic ``Object`` type qualifies; Variant and the
    scalar types do not. Host aliases (Excel/Office) and project class types are
    resolved by the full host-coupled inference in M8; until then they
    conservatively return False, so the analyzer never invents a false positive
    (it only misses some true positives).
    """
    normalized = normalize_type(type_name)
    if normalized is None or normalized == "variant":
        return False
    return normalized == "object"


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
