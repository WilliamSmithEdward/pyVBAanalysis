"""M8/M10: scalar-member-access (objectState.ts) + For Each loop types (controlFlow.ts).

M10 slice 3a un-defers the user-defined-Type / Enum control-variable arm: a control
variable declared As a project Type or Enum now resolves through resolveTypeName
(project-type registry + host model) and is reported with its specific shape. Like
XLIDE, this needs the project type context threaded in (opts.projectTypes); with no
project context the arm stays silent in both implementations.
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

_CODES = (
    "scalar-member-access",
    "for-each-control-variable-type",
    "for-each-source-type",
)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def _project_codes(source: str) -> set[str]:
    """Codes for a single standard module analyzed WITH its own project-type context
    threaded in (mirrors how the LSP/project flow supplies opts.projectTypes)."""
    index = ProjectIndex()
    index.set_module(ModuleInput("Module1", ModuleSymbolKind.STANDARD, source))
    opts = AnalyzeModuleOptions(
        module_name="Module1",
        module_kind=ModuleSymbolKind.STANDARD,
        project_types=index.visible_type_names("Module1"),
    )
    return {d.code for d in analyze_module(source, opts)}


def test_scalar_member_access() -> None:
    assert "scalar-member-access" in _codes("Sub S()\n    Dim n As Long\n    n.Foo = 1\nEnd Sub")


def test_for_each_scalar_control_variable() -> None:
    src = "Sub S()\n    Dim i As Long\n    Dim c As Collection\n    For Each i In c\n    Next\nEnd Sub"
    assert "for-each-control-variable-type" in _codes(src)


def test_for_each_scalar_source() -> None:
    src = "Sub S()\n    Dim v As Variant\n    Dim n As Long\n    For Each v In n\n    Next\nEnd Sub"
    assert "for-each-source-type" in _codes(src)


def test_for_each_udt_control_variable_fires() -> None:
    src = (
        "Public Type TItem\n    Value As Long\nEnd Type\n\n"
        "Public Sub S()\n    Dim item As TItem\n    For Each item In Array(1, 2, 3)\n"
        "    Next item\nEnd Sub\n"
    )
    assert "for-each-control-variable-type" in _project_codes(src)


def test_for_each_enum_control_variable_fires() -> None:
    src = (
        "Public Enum EColor\n    Red = 1\nEnd Enum\n\n"
        "Public Sub S()\n    Dim c As EColor\n    Dim coll As Collection\n"
        "    For Each c In coll\n    Next c\nEnd Sub\n"
    )
    assert "for-each-control-variable-type" in _project_codes(src)


def test_for_each_udt_control_variable_silent_without_project_context() -> None:
    # Faithful to XLIDE: with no projectTypes threaded in, the UDT arm cannot
    # resolve and the control-variable check stays silent (TItem is not a scalar).
    src = (
        "Public Type TItem\n    Value As Long\nEnd Type\n\n"
        "Public Sub S()\n    Dim item As TItem\n    For Each item In Array(1, 2, 3)\n"
        "    Next item\nEnd Sub\n"
    )
    assert "for-each-control-variable-type" not in _codes(src)


def test_valid_for_each_silent() -> None:
    src = "Sub S()\n    Dim v As Variant\n    Dim c As Collection\n    For Each v In c\n    Next\nEnd Sub"
    assert not (_codes(src) & set(_CODES))


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, _CODES)
        assert not spurious, f"{case.id}: scalar/for-each false positive {spurious}"
