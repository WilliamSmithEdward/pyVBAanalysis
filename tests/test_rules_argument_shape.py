"""M8: argument-shape-mismatch (argumentShape.ts parity)."""

from __future__ import annotations

from oracle_support import (
    accepted_cases,
    assert_oracle_behavior,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import analyze_module

_CODES = ("argument-shape-mismatch",)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_array_to_scalar() -> None:
    src = "Sub Foo(n As Long)\nEnd Sub\nSub S()\n    Dim a(3) As Long\n    Foo a\nEnd Sub"
    assert "argument-shape-mismatch" in _codes(src)


def test_scalar_to_array() -> None:
    src = "Sub Foo(a() As Long)\nEnd Sub\nSub S()\n    Dim n As Long\n    Foo n\nEnd Sub"
    assert "argument-shape-mismatch" in _codes(src)


def test_udt_to_scalar() -> None:
    src = (
        "Type T\n    x As Long\nEnd Type\nSub Foo(n As Long)\nEnd Sub\n"
        "Sub S()\n    Dim t As T\n    Foo t\nEnd Sub"
    )
    assert "argument-shape-mismatch" in _codes(src)


def test_matching_shapes_silent() -> None:
    # array -> array, and array -> Variant (boxes), both accepted.
    assert "argument-shape-mismatch" not in _codes(
        "Sub Foo(a() As Long)\nEnd Sub\nSub S()\n    Dim a(3) As Long\n    Foo a\nEnd Sub"
    )
    assert "argument-shape-mismatch" not in _codes(
        "Sub Foo(v As Variant)\nEnd Sub\nSub S()\n    Dim a(3) As Long\n    Foo a\nEnd Sub"
    )


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, _CODES)
        assert not spurious, f"{case.id}: argument-shape false positive {spurious}"
