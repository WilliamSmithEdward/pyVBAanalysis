"""M8: division by a provably-zero divisor (expressions.ts parity)."""

from __future__ import annotations

from oracle_support import (
    accepted_cases,
    assert_oracle_behavior,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import analyze_module

_CODES = ("division-by-zero",)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_literal_zero_divisor() -> None:
    code = "division-by-zero"
    assert code in _codes("Sub S()\n    Dim x As Long\n    x = 1 / 0\nEnd Sub")
    assert code in _codes("Sub S()\n    Dim x As Long\n    x = 5 Mod 0\nEnd Sub")
    assert code in _codes("Sub S()\n    Dim x As Long\n    x = 5 \\ 0\nEnd Sub")
    # A non-zero divisor does not fire.
    assert code not in _codes("Sub S()\n    Dim x As Long\n    x = 1 / 2\nEnd Sub")


def test_constant_zero_divisor() -> None:
    code = "division-by-zero"
    source = "Const Z As Long = 0\nSub S()\n    Dim x As Long\n    x = 10 / Z\nEnd Sub"
    assert code in _codes(source)
    nonzero = "Const Z As Long = 3\nSub S()\n    Dim x As Long\n    x = 10 / Z\nEnd Sub"
    assert code not in _codes(nonzero)


def test_parenthesized_and_signed_zero_divisor() -> None:
    code = "division-by-zero"
    assert code in _codes("Sub S()\n    Dim x As Long\n    x = 1 / (0)\nEnd Sub")
    assert code in _codes("Sub S()\n    Dim x As Long\n    x = 1 / -0\nEnd Sub")
    # A non-constant divisor cannot be proven zero, so no diagnostic.
    assert code not in _codes("Sub S(n As Long)\n    Dim x As Long\n    x = 1 / n\nEnd Sub")


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, _CODES)
        assert not spurious, f"{case.id}: division-by-zero false positive {spurious}"
