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

from pyvbaanalysis import analyze_project
from pyvbaanalysis.diagnostics import analyze_module
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind

_CODES = (
    "call-requires-parens",
    "call-statement-forbids-parens",
    "call-statement-multi-arg-parens",
    "invalid-explicit-call-target",
    "expression-call-requires-parens",
    "invalid-expression-syntax",
)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def _diags(source: str, code: str) -> list:
    return [d for d in analyze_module(source) if d.code == code]


def _project_codes(caller_source: str, helper_source: str) -> set[str]:
    results = analyze_project(
        [
            ModuleInput("Caller", ModuleSymbolKind.STANDARD, caller_source),
            ModuleInput("Helpers", ModuleSymbolKind.STANDARD, helper_source),
        ]
    )
    return {d.code for d in results["Caller"]}


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


def test_multi_arg_parens_bare_sub_call() -> None:
    src = 'Sub T()\n    F("a", "b")\nEnd Sub\nSub F(a As String, b As String)\nEnd Sub\n'
    hits = _diags(src, "call-statement-multi-arg-parens")
    assert len(hits) == 1
    assert src[hits[0].span.start : hits[0].span.end] == 'F("a", "b")'


def test_multi_arg_parens_function_call_statement() -> None:
    src = (
        'Sub T()\n    G("a", "b")\nEnd Sub\n'
        "Function G(a As String, b As String) As String\nEnd Function\n"
    )
    assert "call-statement-multi-arg-parens" in _codes(src)


def test_multi_arg_parens_cross_module_bare() -> None:
    codes = _project_codes(
        'Sub mySub()\n    Helper("a", "b")\nEnd Sub\n',
        "Public Sub Helper(a As String, b As String)\nEnd Sub\n",
    )
    assert "call-statement-multi-arg-parens" in codes


def test_multi_arg_parens_module_qualified() -> None:
    codes = _project_codes(
        'Sub mySub()\n    Helpers.Helper("a", "b")\nEnd Sub\n',
        "Public Sub Helper(a As String, b As String)\nEnd Sub\n",
    )
    assert "call-statement-multi-arg-parens" in codes


def test_multi_arg_parens_unknown_module_silent() -> None:
    codes = _project_codes(
        'Sub mySub()\n    Unknownz.Helper("a", "b")\nEnd Sub\n',
        "Public Sub Helper(a As String, b As String)\nEnd Sub\n",
    )
    assert "call-statement-multi-arg-parens" not in codes


def test_multi_arg_parens_excluded_forms_silent() -> None:
    # Single argument (legal ByVal grouping).
    assert "call-statement-multi-arg-parens" not in _codes(
        'Sub T()\n    F("a")\nEnd Sub\nSub F(a As String)\nEnd Sub\n'
    )
    # Zero-argument empty parens is owned by call-statement-forbids-parens.
    zero = _codes("Sub T()\n    F()\nEnd Sub\nSub F()\nEnd Sub\n")
    assert "call-statement-multi-arg-parens" not in zero
    assert "call-statement-forbids-parens" in zero
    # Parenless call statement form.
    assert "call-statement-multi-arg-parens" not in _codes(
        'Sub T()\n    F "a", "b"\nEnd Sub\nSub F(a As String, b As String)\nEnd Sub\n'
    )
    # Explicit Call with parentheses.
    assert "call-statement-multi-arg-parens" not in _codes(
        'Sub T()\n    Call F("a", "b")\nEnd Sub\nSub F(a As String, b As String)\nEnd Sub\n'
    )
    # A parenthesized call used as an expression value.
    assert "call-statement-multi-arg-parens" not in _codes(
        'Sub T()\n    Dim s As String\n    s = G("a", "b")\nEnd Sub\n'
        "Function G(a As String, b As String) As String\nEnd Function\n"
    )
    # An unknown bare name (array-index / external-reference safety).
    assert "call-statement-multi-arg-parens" not in _codes('Sub T()\n    Maybe("a", "b")\nEnd Sub\n')
    # An object member/property call is deferred to oracle evidence.
    assert "call-statement-multi-arg-parens" not in _codes("Sub T()\n    obj.Method(1, 2)\nEnd Sub\n")


def test_invalid_operator_sequence() -> None:
    assert "invalid-expression-syntax" in _codes("Sub S()\n    Dim x As Long\n    x = 1 * / 2\nEnd Sub")


def _flags_juxtaposition(body: str) -> bool:
    src = f"Sub T()\n    Dim n As Long\n    Dim arr As Variant\n    {body}\nEnd Sub\n"
    return any("expected end of statement" in d.message.lower() for d in analyze_module(src))


def test_juxtaposed_rhs_values_flagged() -> None:
    for body in (
        "n = 1 n 1",
        'n = 1 MsgBox("hello") 1',
        "n = 1 1",
        "n = Foo() bar",
        'n = "a" 1',
    ):
        assert _flags_juxtaposition(body), body


def test_juxtaposed_rhs_values_valid_forms_silent() -> None:
    for body in (
        "n = 1 + 1",
        "n = Foo(1)",
        "n = MsgBox (1)",  # a call written with a space before the paren
        "n = arr(1)(2)",  # jagged-array / call-chain access
        "n = Count&",  # Long type-declaration suffix
        "n = a.b",  # member access
        'n = "x" & "y"',
        "n = IIf(1, 2, 3)",
        "n = (1 = 2)",  # parenthesized comparison
        "n = New Collection",  # New <Type>
        'MsgBox "hello"',  # implicit call statement (no '=')
        "Debug.Print n",  # call statement
        "Foo n, 1",  # call statement with args
    ):
        assert not _flags_juxtaposition(body), body


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
