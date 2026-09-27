"""What XLIDE does to a module's findings before it shows them (module_analysis.py).

XLIDE runs the rules through analyzeVbaModuleSource, which drops a runtime-error
finding under On Error Resume Next and merges findings with one code and span.
analyze_module does the same unless raw_rule_output asks for the rules' own list.
Every expectation here was also run through the pinned upstream wrapper.
"""

from __future__ import annotations

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module


def _lines_with(source: str, code: str, *, raw: bool = False) -> list[int]:
    found = analyze_module(source, AnalyzeModuleOptions(raw_rule_output=raw))
    return [source.count("\n", 0, d.span.start) + 1 for d in found if d.code == code]


def test_a_runtime_error_under_on_error_resume_next_is_not_reported() -> None:
    # `n = UBound(a)` under Resume Next is the usual test for an allocated array.
    source = (
        "Option Explicit\n"
        "Function IsAllocated() As Boolean\n"
        "    Dim a() As Long, n As Long\n"
        "    On Error Resume Next\n"
        "    n = UBound(a)\n"
        "    IsAllocated = (Err.Number = 0)\n"
        "End Function\n"
    )
    assert _lines_with(source, "unallocated-dynamic-array-access") == []
    assert _lines_with(source, "unallocated-dynamic-array-access", raw=True) == [5]


def test_the_stretch_runs_from_resume_next_to_the_next_on_error_statement() -> None:
    source = (
        "Option Explicit\n"
        "Sub Main()\n"
        "    Dim d As Double\n"
        "    d = 1 / 0\n"
        "    On Error Resume Next\n"
        "    d = 2 / 0\n"
        "    On Error GoTo 0\n"
        "    d = 3 / 0\n"
        "End Sub\n"
    )
    assert _lines_with(source, "division-by-zero") == [4, 8]
    assert _lines_with(source, "division-by-zero", raw=True) == [4, 6, 8]


def test_on_local_error_resume_next_counts_too() -> None:
    source = (
        "Option Explicit\n"
        "Sub Main()\n"
        "    Dim d As Double\n"
        "    On Local Error Resume Next\n"
        "    d = 1 / 0\n"
        "End Sub\n"
    )
    assert _lines_with(source, "division-by-zero") == []


def test_resume_next_in_an_if_arm_covers_what_follows_it_in_source_order() -> None:
    # Branches are not modelled: once the arm has run, the handler stays on.
    source = (
        "Option Explicit\n"
        "Sub Main(ByVal quiet As Boolean)\n"
        "    Dim d As Double\n"
        "    If quiet Then\n"
        "        On Error Resume Next\n"
        "    End If\n"
        "    d = 1 / 0\n"
        "End Sub\n"
    )
    assert _lines_with(source, "division-by-zero") == []


def test_the_stretch_ends_with_its_procedure() -> None:
    source = (
        "Option Explicit\n"
        "Sub Quiet()\n"
        "    On Error Resume Next\n"
        "End Sub\n"
        "Sub Loud()\n"
        "    Dim d As Double\n"
        "    d = 1 / 0\n"
        "End Sub\n"
    )
    assert _lines_with(source, "division-by-zero") == [7]


def test_a_compile_error_under_resume_next_is_still_reported() -> None:
    source = (
        "Option Explicit\n"
        "Sub Main()\n"
        "    Dim v As Variant\n"
        "    On Error Resume Next\n"
        "    Set v = 5\n"
        "End Sub\n"
    )
    assert _lines_with(source, "set-requires-object") == [5]


def test_findings_with_one_code_and_span_are_merged_keeping_the_later() -> None:
    # Three string-arithmetic-coercion findings land on the same "abc" in a
    # one-line If; XLIDE shows the last, which names the assignment's target.
    source = (
        "Option Explicit\n"
        "Sub Demo()\n"
        "    Dim x As Long, y As Double\n"
        "    x = 2\n"
        '    If x = 1 Then y = "abc" + 1\n'
        "    Debug.Print y\n"
        "End Sub\n"
    )
    shown = [
        d.message
        for d in analyze_module(source)
        if d.code == "string-arithmetic-coercion"
    ]
    assert len(shown) == 1
    assert shown[0].startswith("Assignment to 'y' expects Double")
    assert len(_lines_with(source, "string-arithmetic-coercion", raw=True)) == 3
