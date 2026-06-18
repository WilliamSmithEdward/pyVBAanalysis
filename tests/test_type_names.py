"""Foundation: pure type-name helpers (typeInference.ts host-free core parity)."""

from __future__ import annotations

from pyvbaanalysis.types import (
    NumericBounds,
    is_known_scalar_type,
    is_numeric_type,
    normalize_type,
    numeric_literal_bounds,
)


def test_normalize_type() -> None:
    assert normalize_type("Long") == "long"
    assert normalize_type("vbLong") == "long"  # leading vb stripped
    assert normalize_type("VBString") == "string"  # case-insensitive vb strip
    assert normalize_type("String()") == "string"  # trailing () stripped
    assert normalize_type("  Variant ( ) ") == "variant"
    assert normalize_type(None) is None
    assert normalize_type("") is None


def test_is_numeric_and_scalar() -> None:
    for t in ("byte", "integer", "long", "longlong", "longptr", "single", "double", "currency", "decimal"):
        assert is_numeric_type(t)
        assert is_known_scalar_type(t)
    assert not is_numeric_type("string")
    assert is_known_scalar_type("string")
    assert is_known_scalar_type("boolean")
    assert is_known_scalar_type("date")
    assert not is_known_scalar_type("object")
    assert not is_known_scalar_type("variant")


def test_numeric_literal_bounds() -> None:
    assert numeric_literal_bounds("byte") == NumericBounds(0, 255, "Byte")
    assert numeric_literal_bounds("integer") == NumericBounds(-32768, 32767, "Integer")
    assert numeric_literal_bounds("long") == NumericBounds(-2147483648, 2147483647, "Long")
    assert numeric_literal_bounds("currency") == NumericBounds(-922337203685477, 922337203685477, "Currency")
    assert numeric_literal_bounds("double") is None  # not range-checked
    assert numeric_literal_bounds("longlong") is None
