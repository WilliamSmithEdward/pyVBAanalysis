"""M6: suffixed numeric-literal overflow (numericLiterals.ts parity)."""

from __future__ import annotations

from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

_CODE = "suffixed-literal-overflow"


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_integer_suffix_overflow() -> None:
    assert _CODE in _codes("x = 40000%")
    assert _CODE in _codes("x = 32768%")  # one past Integer max
    assert _CODE not in _codes("x = 32767%")  # the max is accepted
    # Only the % suffix; & is ambiguous with string concatenation.
    assert _CODE not in _codes('x = 3000000000&"y"')
    # Hex / octal / no-suffix literals never match.
    assert _CODE not in _codes("x = 40000")
    assert _CODE not in _codes("x = &HFFFF")


def test_oracle_asserted_cases() -> None:
    if asserted_cases(_CODE):
        assert assert_oracle_behavior(_CODE) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        assert _CODE not in case_codes(case), f"{case.id}: {_CODE} false positive"
