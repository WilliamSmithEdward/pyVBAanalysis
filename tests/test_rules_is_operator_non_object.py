"""M9: is-operator-non-object (typeOfIs.ts checkIsOperatorOperands parity).

The `Is` operator requires object operands; a provably-scalar operand is a type
error. The rule is a BinaryExpr `Is` expression visitor: it fires on expression-
reachable forms (`If x Is Nothing`, `b = x Is Nothing`) and is silent on
`Debug.Print x Is Nothing` (a reserved-name-receiver statement parses as a raw
StatementNode, so the inner `Is` never reaches the expression walk). This matches
XLIDE exactly - verified by running XLIDE's own analyzer, which is likewise silent
on the Debug.Print form and fires on the If form. The entire oracle corpus uses the
dormant Debug.Print form, so those cases are asserted as XLIDE-dormant here rather
than as positive firings.
"""

from __future__ import annotations

from oracle_support import AUDIT, CASES, accepted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

_CODE = "is-operator-non-object"


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_fires_on_if_condition() -> None:
    src = "Sub S()\n    Dim n As Long\n    If n Is Nothing Then\n    End If\nEnd Sub"
    assert _CODE in _codes(src)


def test_fires_on_assignment_rhs() -> None:
    src = "Sub S()\n    Dim n As Long\n    Dim b As Boolean\n    b = n Is Nothing\nEnd Sub"
    assert _CODE in _codes(src)


def test_fires_on_literal_operand() -> None:
    assert _CODE in _codes("Sub S()\n    Dim b As Boolean\n    b = 1 Is Nothing\nEnd Sub")


def test_object_operand_is_silent() -> None:
    src = "Sub S()\n    Dim o As Object\n    If o Is Nothing Then\n    End If\nEnd Sub"
    assert _CODE not in _codes(src)


def test_unknown_operand_is_silent() -> None:
    # An undeclared / unknown-typed operand is not provably scalar.
    assert _CODE not in _codes("Sub S()\n    If foo Is Nothing Then\n    End If\nEnd Sub")


def test_debug_print_form_is_dormant() -> None:
    # The Debug.Print receiver statement parses raw, so the Is never reaches the
    # walk -- silent, exactly as XLIDE is on this form.
    src = "Sub S()\n    Dim n As Long\n    Debug.Print n Is Nothing\nEnd Sub"
    assert _CODE not in _codes(src)


def test_oracle_cases_match_xlide_dormancy() -> None:
    # Every asserted is-operator-non-object oracle case uses the Debug.Print form,
    # on which XLIDE is dormant; the faithful port reproduces that silence.
    ids = AUDIT[_CODE].asserted_oracle_cases
    cases = [CASES[i] for i in ids if i in CASES and CASES[i].expected == "rejected"]
    assert cases, "expected asserted is-operator-non-object oracle cases"
    for case in cases:
        assert _CODE not in case_codes(case), f"{case.id}: expected XLIDE-dormant silence"


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        assert _CODE not in case_codes(case), f"{case.id}: is-operator-non-object false positive"
