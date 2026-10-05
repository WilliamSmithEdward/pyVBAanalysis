"""Behavioral regressions from the XLIDE 11.0.0 release."""

import pytest

from pyvbaanalysis import analyze_module
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex
from pyvbaanalysis.docs.doc_comment import leading_doc_lines, whole_line_span
from pyvbaanalysis.diagnostics.null_operators import operator_yields_null
from pyvbaanalysis.lexer.tokenize import tokenize
from pyvbaanalysis.constants import resolve_raw_integer_constants
from pyvbaanalysis.conditional import evaluate_conditional_expression


@pytest.mark.parametrize("cycle", [False, True])
def test_deep_constant_dependencies(cycle: bool) -> None:
    count = 2000
    raw = {f"c{i}": f"c{i + 1} + 1" for i in range(count - 1)}
    raw[f"c{count - 1}"] = "c0" if cycle else "1"
    resolved = resolve_raw_integer_constants(raw)
    assert resolved["c0"] == (None if cycle else count)
    assert len(resolved) == count


def test_conditional_parentheses_depth_limit() -> None:
    assert evaluate_conditional_expression("(" * 256 + "1" + ")" * 256) == 1
    assert evaluate_conditional_expression("(" * 257 + "1" + ")" * 257) is None


def test_reference_host_tokens_are_normalized() -> None:
    source = "Sub Main()\nDim xl As Excel.Application\nDim slide As PowerPoint.Slide\nxl.Calculate\nslide.Copy\nEnd Sub"
    found = analyze_module(source, AnalyzeModuleOptions(
        host=" Word ", referenced_hosts=[" EXCEL ", "PowerPoint"], raw_rule_output=True,
    ))
    assert not any(d.code == "missing-library-reference" for d in found)
    assert sum(d.code == "object-variable-not-set" for d in found) == 2


def test_deep_null_expression() -> None:
    assert operator_yields_null(tokenize("Not " * 3000 + "Null"), lambda t: t.raw_text == "Null")


@pytest.mark.parametrize("eol", ["\n", "\r\n", "\r"])
def test_documentation_physical_lines(eol: str) -> None:
    doc = "''' <summary>Example</summary>"
    source = doc + eol + "Sub Main()" + eol + "End Sub"
    lines = leading_doc_lines(source, source.index("Sub"))
    assert len(lines) == 1
    assert lines[0].text == "<summary>Example</summary>"
    assert whole_line_span(source, 0, len(doc)) == (0, len(doc) + len(eol))


@pytest.mark.parametrize("conversion", ["CInt", "CLng", "CByte", "CDbl", "CCur"])
def test_vba_qualified_zero_divisor(conversion: str) -> None:
    source = f"Sub Main()\nDebug.Print 1 / VBA.{conversion}(0)\nEnd Sub"
    findings = analyze_module(source, AnalyzeModuleOptions(raw_rule_output=True))
    assert any(d.code == "division-by-zero" for d in findings)


@pytest.mark.parametrize("eol", ["\n", "\r\n", "\r"])
@pytest.mark.parametrize("mode", ["read", "assigned", "escaped", "constructor"])
def test_bracketed_class_members(eol: str, mode: str) -> None:
    operation = {
        "read": "Debug.Print [c].[member].Count",
        "assigned": "Set [c].[member] = New Collection" + eol + "Debug.Print c.member.Count",
        "escaped": "Take [c]" + eol + "Debug.Print c.member.Count",
        "constructor": "Debug.Print c.member.Count",
    }[mode]
    class_source = "Public member As Object" + eol
    if mode == "constructor":
        class_source += eol.join([
            "Private Sub Class_Initialize()", "Set Me.[member] = New Collection", "End Sub", "",
        ])
    source = eol.join([
        "Option Explicit", "Sub Main()", "Dim c As New C", operation, "End Sub",
        "Sub Take(ByRef value As C)", "Set value.member = New Collection", "End Sub", "",
    ])
    index = ProjectIndex()
    index.set_module(ModuleInput("C", ModuleSymbolKind.CLASS, class_source))
    index.set_module(ModuleInput("M", ModuleSymbolKind.STANDARD, source))
    findings = analyze_module(source, AnalyzeModuleOptions(project_class_members=index.project_class_members()))
    errors = [d for d in findings if d.code == "object-variable-not-set"]
    assert bool(errors) == (mode == "read")


@pytest.mark.parametrize("kind", ["Function", "Property Get"])
@pytest.mark.parametrize("result", ["Object", "Variant"])
def test_bracketed_function_results(kind: str, result: str) -> None:
    assignment = "Set [member] = New Collection" if result == "Object" else "[member] = 1"
    end = "Function" if kind == "Function" else "Property"
    source = f"Public {kind} member() As {result}\n{assignment}\nEnd {end}\n"
    index = ProjectIndex()
    index.set_module(ModuleInput("C", ModuleSymbolKind.CLASS, source))
    member = next(m for c in index.project_class_members() for m in c.members if m.name == "member")
    assert member.known_value == (None if result == "Object" else "scalar")
