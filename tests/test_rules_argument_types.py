"""M8: call-argument type checking (argumentTypes.ts parity)."""

from __future__ import annotations

from oracle_support import (
    accepted_cases,
    assert_oracle_behavior,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import analyze_module

_CODES = (
    "argument-type-mismatch",
    "argument-object-type-mismatch",
    "byref-argument-type-mismatch",
    "string-arithmetic-coercion",
)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_nonnumeric_string_literal_for_numeric_param() -> None:
    src = 'Sub Foo(n As Long)\nEnd Sub\nSub S()\n    Foo "blah"\nEnd Sub'
    assert "argument-type-mismatch" in _codes(src)


def test_numeric_literal_overflow() -> None:
    src = "Sub Foo(n As Byte)\nEnd Sub\nSub S()\n    Foo 300\nEnd Sub"
    assert "argument-type-mismatch" in _codes(src)


def test_byref_exact_type_mismatch() -> None:
    src = "Sub Foo(ByRef n As Long)\nEnd Sub\nSub S()\n    Dim s As String\n    Foo s\nEnd Sub"
    assert "byref-argument-type-mismatch" in _codes(src)


def test_string_arithmetic_into_numeric() -> None:
    src = 'Sub Foo(n As Long)\nEnd Sub\nSub S()\n    Foo 1 + "abc"\nEnd Sub'
    assert "string-arithmetic-coercion" in _codes(src)


def test_compatible_arguments_silent() -> None:
    src = "Sub Foo(n As Long)\nEnd Sub\nSub S()\n    Dim x As Long\n    Foo x\nEnd Sub"
    assert not (_codes(src) & set(_CODES))
    # Variant / unknown arguments are accepted.
    src2 = "Sub Foo(n As Long)\nEnd Sub\nSub S()\n    Dim v As Variant\n    Foo v\nEnd Sub"
    assert not (_codes(src2) & set(_CODES))


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, _CODES)
        assert not spurious, f"{case.id}: argument-type false positive {spurious}"
