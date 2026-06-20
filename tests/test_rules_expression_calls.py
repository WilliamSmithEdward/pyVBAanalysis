"""M8/M10: parenthesized/parenless call-shape + expression-syntax rules (expressions.ts).

M10 slice 3 un-defers the standalone member-call parentheses form (`obj.Method()` ->
call-statement-forbids-parens). A leading-dot member call (`.Method()` inside With)
only fires when the member resolves against the receiver surface (no-FP gate); the
bare leading-dot incomplete-member case already fires through the parser.
"""

from __future__ import annotations

from oracle_support import (
    accepted_cases,
    assert_oracle_behavior,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import analyze_module

_CODES = (
    "call-requires-parens",
    "call-statement-forbids-parens",
    "invalid-explicit-call-target",
    "expression-call-requires-parens",
    "invalid-expression-syntax",
)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_standalone_member_call_parens_fires() -> None:
    # obj.Method() as a statement with empty parens is forbidden.
    assert "call-statement-forbids-parens" in _codes(
        "Public Sub S()\n    ThisWorkbook.CanCheckIn()\nEnd Sub\n"
    )


def test_call_requires_parens() -> None:
    assert "call-requires-parens" in _codes(
        "Sub Foo(a As Long)\nEnd Sub\nSub S()\n    Call Foo 1\nEnd Sub"
    )


def test_call_statement_forbids_parens() -> None:
    assert "call-statement-forbids-parens" in _codes("Sub Foo()\nEnd Sub\nSub S()\n    Foo()\nEnd Sub")


def test_invalid_explicit_call_target() -> None:
    # DoEvents forbids explicit Call.
    assert "invalid-explicit-call-target" in _codes("Sub S()\n    Call DoEvents\nEnd Sub")


def test_expression_call_requires_parens() -> None:
    src = "Function F(a As Long) As Long\nEnd Function\nSub S()\n    Dim x As Long\n    x = F 1\nEnd Sub"
    assert "expression-call-requires-parens" in _codes(src)


def test_invalid_operator_sequence() -> None:
    assert "invalid-expression-syntax" in _codes("Sub S()\n    Dim x As Long\n    x = 1 * / 2\nEnd Sub")


def test_valid_calls_silent() -> None:
    src = "Sub Foo(a As Long)\nEnd Sub\nSub S()\n    Call Foo(1)\nEnd Sub"
    assert not (_codes(src) & set(_CODES))


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, _CODES)
        assert not spurious, f"{case.id}: expression-call false positive {spurious}"
