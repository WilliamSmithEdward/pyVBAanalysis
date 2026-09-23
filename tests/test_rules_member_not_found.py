"""M9: member-not-found rule (undeclared.ts checkMemberNotFound parity).

The oracle sweep runs each case through the shared project-aware harness, which
threads in the project-class member surface the receiver typing needs.
"""

from __future__ import annotations

import dataclasses

from oracle_support import AUDIT, CASES, _kind, case_codes  # type: ignore[attr-defined]

from pyvbaanalysis import analyze_project
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module
from pyvbaanalysis.evidence import OracleCase, OracleModule
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex

_CODE = "member-not-found"


# -- direct unit tests -----------------------------------------------------


def _codes(source: str, opts: AnalyzeModuleOptions | None = None) -> set[str]:
    return {d.code for d in analyze_module(source, opts)}


def test_thisworkbook_alone_falls_back_to_the_open_library_type() -> None:
    """Analyzed with no ThisWorkbook module in sight, ThisWorkbook resolves to
    the library's Workbook. That interface is extensible (XLIDE 10.x, measured
    from the type library's TYPEFLAGS): VBA compiles a member it does not carry
    and asks IDispatch at run time, so nothing is provable there."""
    src = "Public Sub S()\n    ThisWorkbook.AfterSave True\nEnd Sub"
    assert _CODE not in _codes(src)


def _in_a_workbook(entry: str) -> set[str]:
    modules = [
        ModuleInput("ThisWorkbook", ModuleSymbolKind.DOCUMENT, ""),
        ModuleInput("Module1", ModuleSymbolKind.STANDARD, entry),
    ]
    return {d.code for d in analyze_project(modules)["Module1"]}


def test_thisworkbook_is_the_projects_own_closed_class() -> None:
    """In a workbook, ThisWorkbook is the project's document class, and the VBE
    refuses a member it lacks whatever the library's flags say (oracle cases
    workbook_unknown_member_compile and workbook_event_member_call_compile). The
    same typo on a variable declared As Workbook compiles."""
    assert _CODE in _in_a_workbook("Public Sub S()\n    ThisWorkbook.NoSuchMemberXyz\nEnd Sub\n")
    assert _CODE in _in_a_workbook("Public Sub S()\n    ThisWorkbook.AfterSave True\nEnd Sub\n")
    assert _CODE not in _in_a_workbook("Public Sub S()\n    ThisWorkbook.Save\nEnd Sub\n")
    declared = "Public Sub S()\n    Dim wb As Workbook\n    Set wb = ActiveWorkbook\n    wb.NoSuchMemberXyz\nEnd Sub\n"
    assert _CODE not in _in_a_workbook(declared)


def test_a_closed_host_type_still_reports_an_absent_member() -> None:
    """Worksheet IS closed, so a typo on a worksheet variable is a compile error,
    while the same typo on a Range is not."""
    assert _CODE in _codes("Public Sub S()\n    Dim ws As Worksheet\n    ws.NoSuchThing\nEnd Sub")
    assert _CODE not in _codes("Public Sub S()\n    Dim r As Range\n    r.NoSuchThing\nEnd Sub")


def test_a_worksheet_function_on_application_is_ordinary_vba() -> None:
    """VBE-oracle verified (application_worksheet_function_member_compile):
    `Application.Match` is on no interface in the library, and it compiles."""
    src = 'Public Sub S()\n    Dim v As Variant\n    v = Application.Match("a", Range("A1:A9"), 0)\nEnd Sub'
    assert _CODE not in _codes(src)


def test_host_workbook_known_member_is_silent() -> None:
    src = "Public Sub S()\n    ThisWorkbook.Save\nEnd Sub"
    assert _CODE not in _codes(src)


def test_non_exhaustive_host_receiver_is_silent() -> None:
    # Excel.PivotTable is non-exhaustive in the host model: the surface cannot prove
    # absence, so an unknown member never fires.
    src = (
        "Public Sub S()\n    Dim pt As PivotTable\n    Set pt = Nothing\n"
        "    pt.DoesNotExistAtAll\nEnd Sub"
    )
    assert _CODE not in _codes(src)


def test_object_receiver_is_silent() -> None:
    # Object / Variant receivers stay late-bound: no surface, no diagnostic.
    src = (
        "Public Sub S()\n    Dim o As Object\n    Set o = ThisWorkbook\n"
        "    o.DoesNotExist\nEnd Sub"
    )
    assert _CODE not in _codes(src)


def test_unresolved_receiver_is_silent() -> None:
    src = "Public Sub S()\n    foo.Bar\nEnd Sub"
    assert _CODE not in _codes(src)


def test_project_class_unknown_member_fires() -> None:
    person = "Public Sub Save()\nEnd Sub"
    entry = (
        "Public Sub S()\n    Dim p As Person\n    Set p = New Person\n"
        "    p.Delete\nEnd Sub"
    )
    index = ProjectIndex()
    index.set_module(ModuleInput("Person", _kind("class"), person))
    index.set_module(ModuleInput("Module1", _kind("standard"), entry))
    opts = AnalyzeModuleOptions(
        module_name="Module1",
        module_kind=_kind("standard"),
        project_class_members=index.project_class_members(),
    )
    assert _CODE in _codes(entry, opts)


def test_project_class_known_member_is_silent() -> None:
    person = "Public Sub Save()\nEnd Sub"
    entry = (
        "Public Sub S()\n    Dim p As Person\n    Set p = New Person\n"
        "    p.Save\nEnd Sub"
    )
    index = ProjectIndex()
    index.set_module(ModuleInput("Person", _kind("class"), person))
    index.set_module(ModuleInput("Module1", _kind("standard"), entry))
    opts = AnalyzeModuleOptions(
        module_name="Module1",
        module_kind=_kind("standard"),
        project_class_members=index.project_class_members(),
    )
    assert _CODE not in _codes(entry, opts)


def test_project_class_public_field_is_silent() -> None:
    person = "Public Age As Integer"
    entry = (
        "Public Sub S()\n    Dim p As Person\n    Set p = New Person\n"
        "    p.Age = 2\nEnd Sub"
    )
    index = ProjectIndex()
    index.set_module(ModuleInput("Person", _kind("class"), person))
    index.set_module(ModuleInput("Module1", _kind("standard"), entry))
    opts = AnalyzeModuleOptions(
        module_name="Module1",
        module_kind=_kind("standard"),
        project_class_members=index.project_class_members(),
    )
    assert _CODE not in _codes(entry, opts)


# -- oracle sweep ----------------------------------------------------------

# Cases the VBE ran inside a workbook whose ThisWorkbook module the corpus does
# not list. Only the document module makes ThisWorkbook the project's own closed
# class; without it, ThisWorkbook falls back to the library's extensible Workbook
# and nothing is provable.
_IN_A_WORKBOOK: frozenset[str] = frozenset(
    {"workbook_event_member_call_compile", "workbook_unknown_member_compile"}
)

# Asserted cases neither analyzer can meet yet. Empty since XLIDE 10.6.0 closed the
# model's Worksheets the way the library's Sheets is closed (#79), which met
# worksheets_unknown_member_compile.
_SKIP_IDS: frozenset[str] = frozenset()


def _asserted_cases() -> list[OracleCase]:
    return [CASES[i] for i in AUDIT[_CODE].asserted_oracle_cases if i in CASES]


def _as_run(case: OracleCase) -> OracleCase:
    if case.id not in _IN_A_WORKBOOK:
        return case
    this_workbook = OracleModule("ThisWorkbook", "document", "")
    return dataclasses.replace(case, modules=(this_workbook, *case.modules))


def test_oracle_asserted_cases() -> None:
    checked = 0
    for case in _asserted_cases():
        if case.id in _SKIP_IDS:
            continue
        emitted = case_codes(_as_run(case))
        if case.expected == "rejected":
            assert _CODE in emitted, f"{case.id}: expected {_CODE} to fire, got {sorted(emitted)}"
        elif case.expected == "accepted":
            assert _CODE not in emitted, f"{case.id}: {_CODE} must not fire on accepted control"
        checked += 1
    assert checked > 0


def test_no_false_positives_on_accepted_cases() -> None:
    # member-not-found is a compile-equivalent diagnostic, so EVERY accepted case
    # constrains it: it must never fire on compile-valid code.
    for case in CASES.values():
        if case.expected != "accepted":
            continue
        assert _CODE not in case_codes(case), f"{case.id}: {_CODE} false positive"
