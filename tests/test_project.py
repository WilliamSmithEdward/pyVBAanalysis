"""Project driver: analyze_project, the index-to-options glue, and by-name selection."""

from __future__ import annotations

import pytest

from pyvbaanalysis import ConditionalCompilationEnvironment
from pyvbaanalysis.project import (
    analyze_module_options_for,
    analyze_project,
    build_project_index,
)
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind

_PERSON = "Public Sub Save()\nEnd Sub\n"
_ENTRY = (
    "Public Sub S()\n    Dim p As Person\n    Set p = New Person\n    p.Delete\nEnd Sub\n"
)


def _modules() -> list[ModuleInput]:
    return [
        ModuleInput(module_name="Person", module_kind=ModuleSymbolKind.CLASS, source=_PERSON),
        ModuleInput(module_name="Module1", module_kind=ModuleSymbolKind.STANDARD, source=_ENTRY),
    ]


def test_analyze_project_runs_every_module() -> None:
    result = analyze_project(_modules())
    assert set(result) == {"Person", "Module1"}


def test_cross_module_member_not_found_fires() -> None:
    # Module1 sees the Person class via the shared index, so p.Delete (absent) fires.
    result = analyze_project(_modules())
    assert "member-not-found" in {d.code for d in result["Module1"]}


def test_member_not_found_silent_without_project_context() -> None:
    # Analyzing Module1 alone (no Person in the project) cannot resolve the receiver,
    # so member-not-found must stay quiet: the project context is what enables it.
    lone = [ModuleInput(module_name="Module1", module_kind=ModuleSymbolKind.STANDARD, source=_ENTRY)]
    result = analyze_project(lone)
    assert "member-not-found" not in {d.code for d in result["Module1"]}


def test_only_filters_output_but_keeps_context() -> None:
    result = analyze_project(_modules(), only=["module1"])  # case-insensitive
    assert set(result) == {"Module1"}
    # Context is still the whole project, so the cross-module diagnostic still fires.
    assert "member-not-found" in {d.code for d in result["Module1"]}


def test_options_for_populates_project_fields() -> None:
    index = build_project_index(_modules())
    opts = analyze_module_options_for(index, "Module1", ModuleSymbolKind.STANDARD)
    assert opts.module_name == "Module1"
    assert opts.module_kind is ModuleSymbolKind.STANDARD
    assert opts.project_class_members is not None
    assert any(t.name == "Person" for t in opts.project_class_members)
    assert opts.project_procedures is not None


def test_duplicate_module_names_raise() -> None:
    # Case-insensitive name collision would overwrite a module in the index and
    # silently drop its diagnostics, so it must raise instead.
    modules = [
        ModuleInput("Util", ModuleSymbolKind.STANDARD, "Public Sub A()\nEnd Sub\n"),
        ModuleInput("util", ModuleSymbolKind.STANDARD, "Public Sub B()\nEnd Sub\n"),
    ]
    with pytest.raises(ValueError, match="duplicate module name"):
        analyze_project(modules)


def test_conditional_compilation_baseline_controls_branch_activity() -> None:
    # A #If DebugMode block: setting DebugMode False makes the branch inactive, so the
    # diagnostic inside it is suppressed; the kwarg must reach the diagnostic pass.
    src = 'Sub S()\n#If DebugMode Then\n    Dim x As Long\n    x = "bad"\n#End If\nEnd Sub\n'
    modules = [ModuleInput("M", ModuleSymbolKind.STANDARD, src)]

    def codes(cc: ConditionalCompilationEnvironment | None) -> set[str]:
        return {d.code for d in analyze_project(modules, conditional_compilation=cc)["M"]}

    on = ConditionalCompilationEnvironment(project_constants={"DebugMode": True})
    off = ConditionalCompilationEnvironment(project_constants={"DebugMode": False})
    assert "assignment-type-mismatch" in codes(on)
    assert "assignment-type-mismatch" not in codes(off)


def test_options_for_threads_conditional_compilation() -> None:
    index = build_project_index(_modules())
    env = ConditionalCompilationEnvironment(project_constants={"DebugMode": True})
    opts = analyze_module_options_for(index, "Module1", ModuleSymbolKind.STANDARD, conditional_compilation=env)
    assert opts.conditional_compilation is env


def test_severity_override_off_silences_a_code() -> None:
    # option-explicit-missing permits the "off" override; confirm it threads through.
    src = "Public Sub S()\n    Dim n As Long\n    n = 1\nEnd Sub\n"
    modules = [ModuleInput(module_name="M", module_kind=ModuleSymbolKind.STANDARD, source=src)]
    base = analyze_project(modules)["M"]
    assert "option-explicit-missing" in {d.code for d in base}
    overridden = analyze_project(modules, severity_overrides={"option-explicit-missing": "off"})["M"]
    assert "option-explicit-missing" not in {d.code for d in overridden}
