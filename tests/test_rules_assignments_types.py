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

from oracle_support import (
    accepted_cases,
    assert_oracle_behavior,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module
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


def _member_codes(modules: list[tuple[str, ModuleSymbolKind, str]], target: str) -> set[str]:
    """Codes for `target` with the project-class member surface threaded in."""
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
    return {d.code for d in analyze_module(target_src, opts)}


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


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, _CODES)
        assert not spurious, f"{case.id}: assignment-type false positive {spurious}"
