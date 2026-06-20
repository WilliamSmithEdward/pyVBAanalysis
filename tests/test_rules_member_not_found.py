"""M9: member-not-found rule (undeclared.ts checkMemberNotFound parity).

The class oracle cases need the project-class member surface, which the shared
oracle harness does not pass; this file builds its own project-aware harness.
"""

from __future__ import annotations

from oracle_support import AUDIT, CASES, _kind  # type: ignore[attr-defined]

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module
from pyvbaanalysis.evidence import OracleCase
from pyvbaanalysis.symbols import ModuleInput, ProjectIndex

_CODE = "member-not-found"


# -- project-aware oracle harness ------------------------------------------


def _case_codes(case: OracleCase) -> set[str]:
    """Codes analyze_module emits across a case's modules, WITH the project-class
    member surface threaded in (the member-not-found receiver typing needs it)."""
    index = ProjectIndex()
    for module in case.modules:
        index.set_module(ModuleInput(module.name, _kind(module.module_type), module.source))
    project_procedures = index.procedure_signatures()
    project_class_members = index.project_class_members()
    out: set[str] = set()
    for module in case.modules:
        opts = AnalyzeModuleOptions(
            module_name=module.name,
            module_kind=_kind(module.module_type),
            project_procedures=project_procedures,
            project_class_members=project_class_members,
            project_integer_constants=index.visible_external_integer_constant_expressions(module.name),
            project_visible_symbols=index.visible_identifier_symbols(module.name),
            project_types=index.visible_type_names(module.name),
            known_procedures=index.visible_procedure_names(module.name),
            known_identifiers=index.visible_identifier_names(module.name),
            known_non_type_names=index.visible_non_type_names(module.name),
        )
        for diag in analyze_module(module.source, opts):
            out.add(diag.code)
    return out


# -- direct unit tests -----------------------------------------------------


def _codes(source: str, opts: AnalyzeModuleOptions | None = None) -> set[str]:
    return {d.code for d in analyze_module(source, opts)}


def test_host_workbook_event_member_fires() -> None:
    # ThisWorkbook -> Excel.Workbook (exhaustive). AfterSave is an event, excluded
    # from the object surface, so it is reported absent.
    src = "Public Sub S()\n    ThisWorkbook.AfterSave True\nEnd Sub"
    assert _CODE in _codes(src)


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

# Every asserted member-not-found case resolves with the host model + the
# project-class member surface this harness builds, so none are skipped.
_SKIP_IDS: frozenset[str] = frozenset()


def _asserted_cases() -> list[OracleCase]:
    return [CASES[i] for i in AUDIT[_CODE].asserted_oracle_cases if i in CASES]


def test_oracle_asserted_cases() -> None:
    checked = 0
    for case in _asserted_cases():
        if case.id in _SKIP_IDS:
            continue
        emitted = _case_codes(case)
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
        assert _CODE not in _case_codes(case), f"{case.id}: {_CODE} false positive"
