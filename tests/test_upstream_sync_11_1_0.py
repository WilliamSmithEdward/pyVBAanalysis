"""Diagnostic contracts from the pinned XLIDE 11.1.0 release."""

import pytest

from pyvbaanalysis import analyze_module
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.diagnostics.call_extraction import unwrap_outer_parens
from pyvbaanalysis.lexer.tokenize import tokenize
from pyvbaanalysis.types.type_names import normalize_type
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex


def _project_findings(source: str, library: str, kind: ModuleSymbolKind = ModuleSymbolKind.STANDARD):
    index = ProjectIndex()
    index.set_module(ModuleInput("Caller", ModuleSymbolKind.STANDARD, source))
    index.set_module(ModuleInput("Library", kind, library))
    failures = []
    findings = analyze_module(source, AnalyzeModuleOptions(
        module_name="Caller",
        project_class_members=index.project_member_surfaces("Caller"),
        project_visible_symbols=index.visible_identifier_symbols("Caller"),
        on_internal_error=lambda error, where: failures.append((error, where)),
    ))
    assert not failures
    return findings


@pytest.mark.parametrize("name", ["VbClass", "VbEnum", "VbClass()"])
def test_declared_vb_prefix_is_preserved(name: str) -> None:
    assert normalize_type(name) == name.removesuffix("()").lower()


def test_complete_parenthesis_groups_are_unwrapped() -> None:
    assert [t.raw_text for t in unwrap_outer_parens(tokenize("(((1 + 2)))"))] == ["1", "+", "2"]
    assert [t.raw_text for t in unwrap_outer_parens(tokenize("((1) + (2))"))] == ["(", "1", ")", "+", "(", "2", ")"]


@pytest.mark.parametrize("declaration", ["Dim x", "Dim x(1 To 2)"])
def test_module_owned_default_types(declaration: str) -> None:
    argument = "x" if "(" not in declaration else "x(1)"
    source = f"Option Explicit\nDefLng X-X\nSub Main()\n{declaration}\nWork {argument}\nEnd Sub\nSub Work(ByRef x)\nEnd Sub\n"
    findings = analyze_module(source, AnalyzeModuleOptions(known_identifiers=set()))
    assert not any(d.code == "byref-argument-type-mismatch" for d in findings)


def test_paramarray_keeps_variant_elements_under_deftype() -> None:
    source = "Option Explicit\nDefLng A-Z\nSub Main(ParamArray x())\nWork x\nEnd Sub\nSub Work(ByRef x() As Long)\nEnd Sub\n"
    findings = analyze_module(source, AnalyzeModuleOptions(known_identifiers=set()))
    assert any(d.code == "argument-shape-mismatch" for d in findings)


@pytest.mark.parametrize("statement, code", [
    ('ws.Range("A1").Height = 20', "host-readonly-value-assignment"),
    ('ws.EnableCalculation = "nonsense"', "assignment-type-mismatch"),
])
def test_host_scalar_setters(statement: str, code: str) -> None:
    source = f"Option Explicit\nSub Main(ByVal ws As Worksheet)\n{statement}\nEnd Sub\n"
    findings = analyze_module(source)
    assert sum(d.code == code for d in findings) == 1


@pytest.mark.parametrize("value, code", [('True', "argument-shape-mismatch"), ('values', "array-target-assignment")])
def test_array_setter_value_contract(value: str, code: str) -> None:
    source = f"Sub Main()\nDim values() As Boolean\nFlags = {value}\nEnd Sub\nProperty Let Flags(ByRef value() As Boolean)\nEnd Property\n"
    findings = analyze_module(source)
    assert any(d.code == code for d in findings)


@pytest.mark.parametrize("expression", ["Library.[values]", "Library.[MakeFlags]()"])
def test_qualified_bracketed_array_values(expression: str) -> None:
    library = "Public values() As Long\nPublic Function MakeFlags() As Long()\nEnd Function\n"
    source = f"Sub T()\nDim copy() As Boolean\ncopy = {expression}\nEnd Sub\n"
    assert any(d.code == "array-target-assignment" for d in _project_findings(source, library))


def test_indexed_setter_reports_one_argument_count() -> None:
    library = "Public Property Get State(ByVal index As Long) As Boolean\nState = True\nEnd Property\nPublic Property Let State(ByVal index As Long, ByVal value As Boolean)\nEnd Property\n"
    source = "Option Explicit\nSub T(ByVal item As Library)\nitem.State = True\nEnd Sub\n"
    findings = _project_findings(source, library, ModuleSymbolKind.CLASS)
    counts = [d for d in findings if d.code == "argument-count"]
    assert len(counts) == 1
    assert counts[0].message == "Argument not optional: property 'State' requires an index."


@pytest.mark.parametrize("address, reported", [("A1", True), ("A0", False), ("XFE1", False)])
def test_range_area_requires_a_valid_grid_address(address: str, reported: bool) -> None:
    findings = analyze_module(f'Sub Main()\nDebug.Print Range("{address}").Areas(2).Address\nEnd Sub\n')
    assert any(d.code == "host-argument-out-of-range" for d in findings) is reported


def test_errors_only_preserves_errors_and_respects_overrides() -> None:
    source = 'Sub Main()\nDim unused As Long\nDim count As Long\ncount = "wrong"\nEnd Sub\n'
    findings = analyze_module(source, AnalyzeModuleOptions(errors_only=True))
    assert findings
    assert all(d.severity.value == "error" for d in findings)
    assert any(d.code == "assignment-type-mismatch" for d in findings)
    overridden = analyze_module(source, AnalyzeModuleOptions(errors_only=True, severity_overrides={"assignment-type-mismatch": "warning"}))
    assert not any(d.code == "assignment-type-mismatch" for d in overridden)


def test_module_collection_state_is_not_assumed_local() -> None:
    source = 'Option Explicit\nDim c As Collection\nSub Main()\nSet c = New Collection\nFill\nDebug.Print c(1)\nEnd Sub\nSub Fill()\nc.Add 1\nEnd Sub\n'
    assert not any(d.code == "collection-index-out-of-range" for d in analyze_module(source))
