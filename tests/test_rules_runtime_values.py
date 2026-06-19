"""M8: deterministic runtime argument / conversion values (runtimeValues.ts parity)."""

from __future__ import annotations

from oracle_support import (
    accepted_cases,
    assert_oracle_behavior,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import analyze_module

_CODES = ("runtime-argument-value", "runtime-conversion-value")


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_runtime_argument_value_negative_length() -> None:
    code = "runtime-argument-value"
    # Left(s, -1) compiles but raises Run-time error 5.
    assert code in _codes('Sub S()\n    Dim t As String\n    t = Left("ab", -1)\nEnd Sub')
    # A valid length does not fire.
    assert code not in _codes('Sub S()\n    Dim t As String\n    t = Left("ab", 1)\nEnd Sub')


def test_runtime_argument_value_chr_out_of_range() -> None:
    code = "runtime-argument-value"
    assert code in _codes("Sub S()\n    Dim t As String\n    t = Chr(256)\nEnd Sub")
    assert code not in _codes("Sub S()\n    Dim t As String\n    t = Chr(65)\nEnd Sub")


def test_runtime_argument_value_user_shadow_suppressed() -> None:
    # A user-defined Left procedure shadows the intrinsic; do not fire.
    code = "runtime-argument-value"
    source = (
        "Function Left(ByVal s As String, ByVal n As Integer) As String\n"
        "End Function\n"
        "Sub S()\n    Dim t As String\n    t = Left(\"ab\", -1)\nEnd Sub"
    )
    assert code not in _codes(source)


def test_runtime_conversion_value_invalid_date_string() -> None:
    code = "runtime-conversion-value"
    assert code in _codes('Sub S()\n    Dim d As Date\n    d = CDate("hello")\nEnd Sub')
    # A date-like string is not flagged (conservative).
    assert code not in _codes('Sub S()\n    Dim d As Date\n    d = CDate("January 1")\nEnd Sub')


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, _CODES)
        assert not spurious, f"{case.id}: runtime-value false positive {spurious}"
