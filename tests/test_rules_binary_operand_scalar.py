"""M8: non-scalar binary operand (binaryOperandScalar.ts parity)."""

from __future__ import annotations

from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

_CODE = "non-scalar-binary-operand"


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_non_scalar_binary_operand() -> None:
    arr = "Dim a(3) As Long\n    Dim s As String\n"
    # An array operand of a scalar operator (concat / arithmetic / comparison).
    assert _CODE in _codes(f'Sub S\n    {arr}    s = a & "x"\nEnd Sub')
    assert _CODE in _codes(f"Sub S\n    {arr}    Dim n As Long\n    n = a + 1\nEnd Sub")
    assert _CODE in _codes(f"Sub S\n    {arr}    Dim b As Boolean\n    b = a < 1\nEnd Sub")
    # A same-module user-defined Type operand.
    udt = "Private Type TPoint\n    x As Long\nEnd Type\n"
    assert _CODE in _codes(f'{udt}Sub S\n    Dim p As TPoint\n    Dim s As String\n    s = p & "x"\nEnd Sub')
    # An indexed element a(0) is a scalar element -> quiet.
    assert _CODE not in _codes(f'Sub S\n    {arr}    s = a(0) & "x"\nEnd Sub')
    # Scalar / Variant operands -> quiet.
    assert _CODE not in _codes('Sub S\n    Dim n As Long\n    Dim s As String\n    s = n & "x"\nEnd Sub')
    assert _CODE not in _codes('Sub S\n    Dim v As Variant\n    Dim s As String\n    s = v & "x"\nEnd Sub')
    # Is (object operands) is out of scope for this rule.
    assert _CODE not in _codes("Sub S\n    Dim o As Object\n    If o Is Nothing Then\n    End If\nEnd Sub")


def test_oracle_asserted_cases() -> None:
    if asserted_cases(_CODE):
        assert assert_oracle_behavior(_CODE) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        assert _CODE not in case_codes(case), f"{case.id}: {_CODE} false positive"
