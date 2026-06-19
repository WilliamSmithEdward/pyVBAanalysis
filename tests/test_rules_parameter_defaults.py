"""M8: optional-parameter default value type checking (declarations.ts parity)."""

from __future__ import annotations

from oracle_support import (
    accepted_cases,
    assert_oracle_behavior,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import analyze_module

_CODES = ("parameter-default-type-mismatch", "parameter-default-not-constant")


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_non_constant_default_is_flagged() -> None:
    src = "Function F() As Long\nEnd Function\nSub S(Optional n As Long = F())\nEnd Sub"
    assert "parameter-default-not-constant" in _codes(src)


def test_constant_default_not_flagged() -> None:
    assert "parameter-default-not-constant" not in _codes("Sub S(Optional n As Long = 5)\nEnd Sub")
    # A bare/qualified identifier may be a constant, so it is left alone.
    assert "parameter-default-not-constant" not in _codes(
        "Sub S(Optional n As Long = SomeConst)\nEnd Sub"
    )


def test_string_default_for_numeric_param() -> None:
    assert "parameter-default-type-mismatch" in _codes('Sub S(Optional n As Long = "blah")\nEnd Sub')


def test_non_nothing_default_for_object_param() -> None:
    assert "parameter-default-type-mismatch" in _codes("Sub S(Optional o As Object = 5)\nEnd Sub")


def test_valid_defaults_silent() -> None:
    assert "parameter-default-type-mismatch" not in _codes("Sub S(Optional n As Long = 5)\nEnd Sub")
    assert "parameter-default-type-mismatch" not in _codes(
        "Sub S(Optional o As Object = Nothing)\nEnd Sub"
    )


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, _CODES)
        assert not spurious, f"{case.id}: parameter-default false positive {spurious}"
