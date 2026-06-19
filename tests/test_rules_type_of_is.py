"""M8: TypeOf ... Is rules (typeOfIs.ts parity)."""

from __future__ import annotations

from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

_CODES = ("typeof-missing-operand",)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_typeof_missing_operand() -> None:
    code = "typeof-missing-operand"
    # `TypeOf Is Y` with no operand between is a syntax error.
    assert code in _codes("Sub S(obj As Object)\n    If TypeOf Is Worksheet Then\n    End If\nEnd Sub")
    # A valid `TypeOf x Is Y` has an operand.
    assert code not in _codes(
        "Sub S(obj As Object)\n    If TypeOf obj Is Worksheet Then\n    End If\nEnd Sub"
    )


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    codes = set(_CODES)
    for case in accepted_cases():
        spurious = case_codes(case) & codes
        assert not spurious, f"{case.id}: TypeOf-Is false positive {spurious}"
