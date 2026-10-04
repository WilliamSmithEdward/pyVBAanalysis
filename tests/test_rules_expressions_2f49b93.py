"""Upstream 2f49b93's changes to expressions.ts, each measured in Excel 16.0
upstream: a lone `Not` (issue #234), `x^=5` (issue #369), a multi-argument
member call in parentheses (issue #224), `1 \\ False` (issue #458), a
conversion of 0 as a divisor (issue #219), and Not binding below the
comparisons (issue #361)."""

from __future__ import annotations

from pyvbaanalysis.diagnostics import analyze_module
from pyvbaanalysis.diagnostics.model import VbaDiagnostic


def _module(lines: list[str], extra: str = "") -> str:
    body = "\n".join(f"    {line}" for line in lines)
    return f"Option Explicit\n{extra}Function Main() As Variant\n{body}\nEnd Function\n"


def _found(source: str, code: str) -> list[VbaDiagnostic]:
    return [diag for diag in analyze_module(source) if diag.code == code]


def _text(source: str, diag: VbaDiagnostic) -> str:
    return source[diag.span.start : diag.span.end]


def test_not_with_nothing_to_negate() -> None:
    src = _module(["Dim x As Variant", "x = Not", "Main = x"])
    found = _found(src, "invalid-expression-syntax")
    assert [diag.message for diag in found] == [
        "'Not' has nothing after it to negate. This is a VBE compile error: Syntax error."
    ]
    assert _text(src, found[0]) == "Not"


def test_glued_longlong_suffix_assignment_is_no_operator_run() -> None:
    src = _module(["Dim x^", "x^=5", "Main = x"])
    assert _found(src, "invalid-expression-syntax") == []


def test_juxtaposed_value_index() -> None:
    from pyvbaanalysis.diagnostics.rules.expressions import juxtaposed_value_index
    from pyvbaanalysis.diagnostics.walker import raw_expression_tokens

    toks = raw_expression_tokens("asdf qwer")
    assert juxtaposed_value_index(toks, 0) == 1
    assert juxtaposed_value_index(raw_expression_tokens("f(a b) + 1"), 0) == -1


def test_member_call_with_parenthesized_arguments() -> None:
    src = _module(["Dim c As Collection", "c.Calc (1, 2)", "Main = 1"])
    found = _found(src, "call-statement-multi-arg-parens")
    assert len(found) == 1
    assert "'Call c.Calc(...)'" in found[0].message


def test_false_and_zero_conversions_divide_by_zero() -> None:
    for expression in ("1 \\ False", "10 / CLng(0)", "10 / CLng(0.4)", "10 / CDbl(0)"):
        src = _module([f"Main = {expression}"])
        assert _found(src, "division-by-zero") != [], expression
    # `CDbl(0.4)` is no zero: `10 / CDbl(0.4)` runs.
    assert _found(_module(["Main = 10 / CDbl(0.4)"]), "division-by-zero") == []
    # Pinned XLIDE also misses this form (xlide_vscode #898); keep parity.
    assert _found(_module(["Main = 10 / VBA.CDbl(0)"]), "division-by-zero") == []


def test_not_binds_below_the_comparisons() -> None:
    for lines in (
        ["Dim s As String", 's = "x"', 'Main = Not s = "y"'],
        ["Dim s As String", 's = "abc"', 'Main = Not s Like "a*"'],
        ["Dim s As String", 's = "x"', 'If Not s = "Sheet1" Then Main = 1'],
    ):
        assert _found(_module(lines), "string-arithmetic-coercion") == [], lines
    src = _module(["Dim s As String", 's = "x"', "Main = Not s"])
    found = _found(src, "string-arithmetic-coercion")
    assert len(found) == 1
    assert _text(src, found[0]) == "s"
