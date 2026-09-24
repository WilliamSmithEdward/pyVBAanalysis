"""M8/M10: assignment / Set / missing-return type rules (assignments.ts parity).

M10 slice 3d un-defers member-access assignment typing (`obj.Member = value`):
checkMemberAssignmentTypes resolves the exact member via resolveExactMemberCompletion
(now carrying writable/write_type/returns) and applies the same compatibility rules
as bare assignment, plus read-only and Set-required checks. Only source-backed
project members carry writability, so host members and unresolved receivers stay
silent (the no-FP gate). The two formerly-deferred oracle cases (a non-numeric
string assigned to an Integer property/field) now fire through the shared harness,
which threads project_class_members in.
"""

from __future__ import annotations

import pytest
from oracle_support import (
    accepted_cases,
    assert_oracle_behavior,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, VbaDiagnostic, analyze_module
from pyvbaanalysis.project import analyze_project
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex

# Runtime-error-kind and compile-error-kind codes emitted by these rules.
_CODES = (
    "assignment-type-mismatch",
    "string-arithmetic-coercion",
    "array-assignment-to-scalar",
    "set-required",
    "set-requires-object",
    "missing-return-assignment",
    "readonly-member-assignment",
    "assignment-object-type-mismatch",
)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def _member_diagnostics(
    modules: list[tuple[str, ModuleSymbolKind, str]], target: str
) -> list[VbaDiagnostic]:
    """Diagnostics for `target` with the project-class member surface threaded in."""
    index = ProjectIndex()
    for name, kind, src in modules:
        index.set_module(ModuleInput(name, kind, src))
    target_src = next(src for name, _k, src in modules if name == target)
    opts = AnalyzeModuleOptions(
        module_name=target,
        module_kind=next(k for n, k, _s in modules if n == target),
        project_class_members=index.project_class_members(),
        project_visible_symbols=index.visible_identifier_symbols(target),
    )
    return analyze_module(target_src, opts)


def _member_codes(modules: list[tuple[str, ModuleSymbolKind, str]], target: str) -> set[str]:
    """Codes for `target` with the project-class member surface threaded in."""
    return {d.code for d in _member_diagnostics(modules, target)}


def test_assignment_type_mismatch() -> None:
    assert "assignment-type-mismatch" in _codes('Sub S()\n    Dim n As Long\n    n = "blah"\nEnd Sub')


def test_string_arithmetic_in_assignment() -> None:
    assert "string-arithmetic-coercion" in _codes(
        'Sub S()\n    Dim n As Long\n    n = 1 + "abc"\nEnd Sub'
    )


def test_set_requires_object() -> None:
    assert "set-requires-object" in _codes("Sub S()\n    Dim n As Long\n    Set n = Nothing\nEnd Sub")


def test_object_assignment_requires_set() -> None:
    assert "set-required" in _codes("Sub S()\n    Dim o As Object\n    o = 5\nEnd Sub")


def test_array_assignment_to_scalar() -> None:
    src = "Sub S()\n    Dim a(3) As Long\n    Dim n As Long\n    n = a\nEnd Sub"
    assert "array-assignment-to-scalar" in _codes(src)


def test_missing_return_assignment() -> None:
    assert "missing-return-assignment" in _codes("Function F()\nEnd Function")
    # A function that assigns its return name is silent.
    assert "missing-return-assignment" not in _codes("Function F()\n    F = 1\nEnd Function")


def test_compatible_assignments_silent() -> None:
    src = "Sub S()\n    Dim n As Long\n    n = 5\nEnd Sub"
    assert not (_codes(src) & set(_CODES))


# -- member-access assignment typing (M10 slice 3d) ------------------------


def test_member_scalar_type_mismatch_fires() -> None:
    # p.Age (Integer field) assigned a non-numeric string literal.
    mods = [
        ("Person", ModuleSymbolKind.CLASS, "Public Age As Integer\n"),
        ("M", ModuleSymbolKind.STANDARD,
         'Public Sub S()\n    Dim p As Person\n    Set p = New Person\n    p.Age = "blah"\nEnd Sub\n'),
    ]
    assert "assignment-type-mismatch" in _member_codes(mods, "M")


def test_member_readonly_assignment_fires() -> None:
    # Age is a Get-only property (no Let), so it is read-only.
    mods = [
        ("Person", ModuleSymbolKind.CLASS,
         "Public Property Get Age() As Integer\n    Age = 1\nEnd Property\n"),
        ("M", ModuleSymbolKind.STANDARD,
         "Public Sub S()\n    Dim p As Person\n    Set p = New Person\n    p.Age = 5\nEnd Sub\n"),
    ]
    assert "readonly-member-assignment" in _member_codes(mods, "M")


_READONLY_PERSON = (
    "Person",
    ModuleSymbolKind.CLASS,
    "Public Property Get Age() As Integer\n    Age = 1\nEnd Property\n"
    "Public Property Get Items() As Collection\nEnd Property\n",
)


def test_a_comparison_in_a_condition_is_no_assignment() -> None:
    # The parser keeps `ElseIf p.Age = 2 Then` as a statement of its If block, and
    # a single-line If is one statement, so both carry an `=` after `p.Age`. A
    # keyword stands before the receiver, which makes neither an assignment
    # (XLIDE issue #78).
    body = (
        "Public Sub S()\n    Dim p As Person\n    Set p = New Person\n"
        "    If p.Age = 1 Then\n        Exit Sub\n    ElseIf p.Age = 2 Or p.Age = 3 Then\n"
        "        Exit Sub\n    End If\n    If p.Age = 4 Then Exit Sub\nEnd Sub\n"
    )
    mods = [_READONLY_PERSON, ("M", ModuleSymbolKind.STANDARD, body)]
    assert "readonly-member-assignment" not in _member_codes(mods, "M")


@pytest.mark.parametrize(
    "statement",
    [
        pytest.param("Select Case True\n    Case p.Age = 2\n    End Select", id="a Case expression"),
        pytest.param("Do While p.Age = 2\n    Loop", id="a Do While condition"),
        pytest.param("Debug.Print p.Age = 2", id="a call given the comparison"),
        pytest.param("MsgBox p.Age = 2", id="a procedure called with the comparison"),
    ],
)
def test_a_comparison_passed_on_or_tested_is_no_assignment(statement: str) -> None:
    # A target is one receiver chain ending in the member; anything else before
    # the `=` is another statement comparing it (XLIDE issue #78).
    body = f"Public Sub S()\n    Dim p As Person\n    Set p = New Person\n    {statement}\nEnd Sub\n"
    mods = [_READONLY_PERSON, ("M", ModuleSymbolKind.STANDARD, body)]
    assert "readonly-member-assignment" not in _member_codes(mods, "M")


def test_a_readonly_member_assignment_still_fires_through_any_receiver_chain() -> None:
    for statement in ("p.Age = 5", "With p\n        .Age = 5\n    End With", "Let p.Age = 5"):
        body = f"Public Sub S()\n    Dim p As Person\n    Set p = New Person\n    {statement}\nEnd Sub\n"
        mods = [_READONLY_PERSON, ("M", ModuleSymbolKind.STANDARD, body)]
        assert "readonly-member-assignment" in _member_codes(mods, "M"), statement


def test_a_single_line_if_branch_assignment_is_named_by_itself() -> None:
    # Read whole, the target was 'If True Then p.Age'. The If's branches are
    # statements of their own, as the other assignment rules read them.
    body = "Public Sub S()\n    Dim p As Person\n    Set p = New Person\n    If True Then p.Age = 2\nEnd Sub\n"
    mods = [_READONLY_PERSON, ("M", ModuleSymbolKind.STANDARD, body)]
    hits = [d for d in _member_diagnostics(mods, "M") if d.code == "readonly-member-assignment"]
    assert [d.message for d in hits] == ["Cannot assign to read-only property 'p.Age'."]
    assert [body[d.span.start : d.span.end] for d in hits] == ["Age"]


def test_set_object_type_mismatch_fires() -> None:
    # Set a (ClassA) = New ClassB: incompatible project object types.
    mods = [
        ("ClassA", ModuleSymbolKind.CLASS, "Public Sub A()\nEnd Sub\n"),
        ("ClassB", ModuleSymbolKind.CLASS, "Public Sub B()\nEnd Sub\n"),
        ("M", ModuleSymbolKind.STANDARD,
         "Public Sub S()\n    Dim a As ClassA\n    Set a = New ClassB\nEnd Sub\n"),
    ]
    assert "assignment-object-type-mismatch" in _member_codes(mods, "M")


def test_set_compatible_object_silent() -> None:
    # Set a (ClassA) = New ClassA: same type, no mismatch.
    mods = [
        ("ClassA", ModuleSymbolKind.CLASS, "Public Sub A()\nEnd Sub\n"),
        ("M", ModuleSymbolKind.STANDARD,
         "Public Sub S()\n    Dim a As ClassA\n    Set a = New ClassA\nEnd Sub\n"),
    ]
    assert "assignment-object-type-mismatch" not in _member_codes(mods, "M")


def test_member_object_assignment_requires_set_fires() -> None:
    # p.Pal is typed As a project class (Buddy), so a plain (non-Set) assignment of
    # an object value requires Set.
    mods = [
        ("Buddy", ModuleSymbolKind.CLASS, "Public Sub Greet()\nEnd Sub\n"),
        ("Person", ModuleSymbolKind.CLASS, "Public Pal As Buddy\n"),
        ("M", ModuleSymbolKind.STANDARD,
         "Public Sub S()\n    Dim p As Person\n    Set p = New Person\n    p.Pal = New Buddy\nEnd Sub\n"),
    ]
    assert "set-required" in _member_codes(mods, "M")


def test_member_compatible_assignment_silent() -> None:
    # A numeric literal assigned to an Integer member is fine.
    mods = [
        ("Person", ModuleSymbolKind.CLASS, "Public Age As Integer\n"),
        ("M", ModuleSymbolKind.STANDARD,
         "Public Sub S()\n    Dim p As Person\n    Set p = New Person\n    p.Age = 5\nEnd Sub\n"),
    ]
    assert not (_member_codes(mods, "M") & set(_CODES))


def test_host_member_assignment_silent() -> None:
    # Host members carry no writability proof, so assignment typing stays silent.
    src = 'Public Sub S()\n    ThisWorkbook.Name = "x"\nEnd Sub\n'
    assert not (_codes(src) & set(_CODES))


def _excel_project_codes(source: str) -> set[str]:
    """Codes for one standard module analyzed as a project, which loads the Excel model."""
    results = analyze_project([ModuleInput("M", ModuleSymbolKind.STANDARD, source)])
    return {d.code for d in results["M"]}


def _host_set_source(declared: str, receiver: str, member: str) -> str:
    return (
        f"Public Sub S(ByVal source As {receiver})\n"
        f"    Dim target As {declared}\n"
        f"    Set target = source.{member}\n"
        "End Sub\n"
    )


def test_host_duplicate_of_a_shape_range_stays_silent() -> None:
    # ShapeRange.Duplicate returns a ShapeRange, and the model says so.
    src = _host_set_source("ShapeRange", "ShapeRange", "Duplicate")
    assert "assignment-object-type-mismatch" not in _excel_project_codes(src)


# Each Set below compiles and runs in Excel: the type library declares the member
# As the target type, and TypeName of the value Excel returns is that type (both
# checked 2026-09-23). The model mistyped all four until XLIDE 10.7.1 (#90): the
# hand-written excelObjectModel.ts returned a ShapeRange from Shape.Duplicate and
# a SparkColor from SparklineGroup.SeriesColor, and the model's own Charts and
# Worksheets where the library returns Sheets.
@pytest.mark.parametrize(
    ("declared", "receiver", "member"),
    [
        pytest.param("Shape", "Shape", "Duplicate", id="Shape.Duplicate"),
        pytest.param("FormatColor", "SparklineGroup", "SeriesColor", id="SparklineGroup.SeriesColor"),
        pytest.param("Sheets", "Application", "Charts", id="Application.Charts"),
        pytest.param("Sheets", "Workbook", "Worksheets", id="Workbook.Worksheets"),
    ],
)
def test_host_object_assignment_follows_the_type_library(
    declared: str, receiver: str, member: str
) -> None:
    src = _host_set_source(declared, receiver, member)
    assert "assignment-object-type-mismatch" not in _excel_project_codes(src)


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, _CODES)
        assert not spurious, f"{case.id}: assignment-type false positive {spurious}"
