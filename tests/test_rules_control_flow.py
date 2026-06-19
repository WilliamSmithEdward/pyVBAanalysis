"""M6: control-flow rule family (controlFlow.ts parity, self-contained slice)."""

from __future__ import annotations

import pytest
from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

_CF_CODES = (
    "exit-wrong-proc",
    "if-missing-then",
    "case-outside-select",
    "member-access-outside-with",
    "exit-outside-block",
    "next-variable-mismatch",
    "duplicate-case-else",
    "else-without-if",
    "else-branch-order",
    "duplicate-label",
    "undefined-label",
    "invalid-assignment-target",
    "open-missing-for",
)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_exit_statement_kind() -> None:
    assert "exit-wrong-proc" in _codes("Function F() As Long\n    Exit Sub\nEnd Function")
    assert "exit-wrong-proc" not in _codes("Sub S\n    Exit Sub\nEnd Sub")
    # Exit For / Exit Do are loop exits, not procedure exits.
    assert "exit-wrong-proc" not in _codes("Sub S\n    For i = 1 To 2\n        Exit For\n    Next i\nEnd Sub")


def test_if_missing_then() -> None:
    assert "if-missing-then" in _codes("Sub S\n    If x\nEnd Sub")
    assert "if-missing-then" not in _codes("Sub S\n    If x Then y = 1\nEnd Sub")


def test_case_outside_select() -> None:
    assert "case-outside-select" in _codes("Sub S\n    Case 1\nEnd Sub")
    assert "case-outside-select" not in _codes(
        "Sub S\n    Select Case x\n    Case 1\n    End Select\nEnd Sub"
    )


def test_leading_dot_outside_with() -> None:
    assert "member-access-outside-with" in _codes("Sub S\n    .Value = 1\nEnd Sub")
    assert "member-access-outside-with" not in _codes(
        "Sub S\n    With obj\n        .Value = 1\n    End With\nEnd Sub"
    )


def test_loop_exit_context() -> None:
    assert "exit-outside-block" in _codes("Sub S\n    Exit For\nEnd Sub")
    assert "exit-outside-block" in _codes("Sub S\n    Exit Do\nEnd Sub")
    assert "exit-outside-block" not in _codes("Sub S\n    For i = 1 To 2\n        Exit For\n    Next i\nEnd Sub")


def test_next_variable_mismatch() -> None:
    assert "next-variable-mismatch" in _codes("Sub S\n    For i = 1 To 10\n    Next j\nEnd Sub")
    assert "next-variable-mismatch" not in _codes("Sub S\n    For i = 1 To 10\n    Next i\nEnd Sub")
    # A bare `Next` (no variable) is fine.
    assert "next-variable-mismatch" not in _codes("Sub S\n    For i = 1 To 10\n    Next\nEnd Sub")


def test_duplicate_case_else() -> None:
    src = "Sub S\n    Select Case x\n    Case Else\n    Case Else\n    End Select\nEnd Sub"
    assert "duplicate-case-else" in _codes(src)
    ok = "Sub S\n    Select Case x\n    Case 1\n    Case Else\n    End Select\nEnd Sub"
    assert "duplicate-case-else" not in _codes(ok)


def test_else_without_if() -> None:
    assert "else-without-if" in _codes("Sub S\n    Else\nEnd Sub")
    assert "else-without-if" not in _codes(
        "Sub S\n    If x Then\n        a = 1\n    Else\n        a = 2\n    End If\nEnd Sub"
    )


def test_malformed_statements() -> None:
    assert "invalid-assignment-target" in _codes("Sub S\n    1 = x\nEnd Sub")
    assert "invalid-assignment-target" not in _codes("Sub S\n    x = 1\nEnd Sub")
    assert "open-missing-for" in _codes('Sub S\n    Open "f.txt" As #1\nEnd Sub')
    assert "open-missing-for" not in _codes('Sub S\n    Open "f.txt" For Input As #1\nEnd Sub')


def test_undefined_labels() -> None:
    code = "undefined-label"
    # GoTo / GoSub / On Error GoTo / On n GoTo targets must exist in the procedure.
    assert code in _codes("Sub S\n    GoTo Missing\nEnd Sub")
    assert code not in _codes("Sub S\n    GoTo Done\nDone:\nEnd Sub")
    assert code in _codes("Sub S\n    On Error GoTo Handler\nEnd Sub")
    assert code not in _codes("Sub S\n    On Error GoTo Handler\nHandler:\nEnd Sub")
    # Decimal line labels are matched after leading-zero normalization.
    assert code not in _codes("Sub S\n    GoTo 10\n10:\nEnd Sub")
    # On Error Resume Next / GoTo 0 / GoTo -1 are non-label forms, never a target.
    assert code not in _codes("Sub S\n    On Error Resume Next\nEnd Sub")
    assert code not in _codes("Sub S\n    On Error GoTo 0\nEnd Sub")
    assert code not in _codes("Sub S\n    On Error GoTo -1\nEnd Sub")
    # Labels do not leak across procedures.
    cross = "Sub A\n    GoTo Shared\nEnd Sub\nSub B\nShared:\nEnd Sub"
    assert code in _codes(cross)


def test_duplicate_labels() -> None:
    code = "duplicate-label"
    assert code in _codes("Sub S\nDup:\nDup:\nEnd Sub")
    assert code not in _codes("Sub S\nA:\nB:\nEnd Sub")
    # Same label name in two different procedures is fine.
    assert code not in _codes("Sub A\nDone:\nEnd Sub\nSub B\nDone:\nEnd Sub")


def test_else_branch_order() -> None:
    code = "else-branch-order"
    # Runtime If blocks: ElseIf/Else after the final Else is rejected; the
    # normal ElseIf-before-Else ordering is accepted.
    assert code in _codes("Sub S\n    If x Then\n    Else\n    ElseIf y Then\n    End If\nEnd Sub")
    assert code in _codes("Sub S\n    If x Then\n    Else\n    Else\n    End If\nEnd Sub")
    assert code not in _codes("Sub S\n    If x Then\n    ElseIf y Then\n    Else\n    End If\nEnd Sub")
    # Conditional-compilation directives follow the same branch-order rule.
    assert code in _codes("Sub S\n#If False Then\n#Else\n#ElseIf True Then\n#End If\nEnd Sub")
    assert code in _codes("Sub S\n#If False Then\n#Else\n#Else\n#End If\nEnd Sub")
    assert code not in _codes("Sub S\n#If False Then\n#ElseIf True Then\n#Else\n#End If\nEnd Sub")


@pytest.mark.parametrize("code", _CF_CODES)
def test_oracle_asserted_cases(code: str) -> None:
    if not asserted_cases(code):
        pytest.skip(f"{code} has no asserted corpus cases")
    assert assert_oracle_behavior(code) > 0


def test_no_control_flow_false_positives_on_accepted_cases() -> None:
    cf = set(_CF_CODES) | {"invalid-expression-syntax"}
    for case in accepted_cases():
        spurious = case_codes(case) & cf
        assert not spurious, f"{case.id}: control-flow false positive {spurious}"
