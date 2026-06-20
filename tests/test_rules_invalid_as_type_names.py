"""M10: invalidAsTypeNames rule (declarations.ts FULL parity).

Every type-name reference is resolved with resolveTypeName over the project-type
registry (opts.project_types) + host model, exactly as XLIDE does. Branches:
  - invalid-as-type-name: an ambiguous project-type name (multiple visible project
    types share it), a reserved VBA identifier, a VBA runtime function, or a known
    project non-type declaration that does not resolve to a real type.
  - invalid-new-type-name: a `New` reference (`As New T` / `New T`) to a type that
    resolves but is NOT creatable and is NOT a host type (only project classes and
    UserForms may be created with New). WithEvents `As New` is exempt.
Qualified references (`Mod.Type`) resolve through the module-qualified candidate
set. The no-false-positive guarantee comes from resolveTypeName itself plus the
project-type context the caller threads in (e.g. `Long`/`Filter` resolve, so they
never fall through to the reserved/runtime branches). The ambiguous and
invalid-new-type-name branches need project context (project_types), so the
positives for them thread a ProjectIndex in.
"""

from __future__ import annotations

from oracle_support import (  # type: ignore[attr-defined]
    accepted_cases,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex

_CODE = "invalid-as-type-name"
_NEW_CODE = "invalid-new-type-name"
_CODES = (_CODE, _NEW_CODE)


_STD = AnalyzeModuleOptions(module_name="Module1", module_kind=ModuleSymbolKind.STANDARD)


def _codes(source: str, opts: AnalyzeModuleOptions | None = None) -> set[str]:
    return {d.code for d in analyze_module(source, opts)}


def _project_codes(modules: list[tuple[str, ModuleSymbolKind, str]], target: str) -> set[str]:
    """Codes for `target` analyzed with the full project context threaded in."""
    index = ProjectIndex()
    for name, kind, src in modules:
        index.set_module(ModuleInput(name, kind, src))
    target_src = next(src for name, _k, src in modules if name == target)
    opts = AnalyzeModuleOptions(
        module_name=target,
        module_kind=next(k for n, k, _s in modules if n == target),
        project_types=index.visible_type_names(target),
        project_class_members=index.project_class_members(),
        known_non_type_names=index.visible_non_type_names(target),
    )
    return {d.code for d in analyze_module(target_src, opts)}


# -- direct positives ------------------------------------------------------


def test_runtime_function_as_type_fires() -> None:
    # `Int` is a VBA runtime function, not a type (the oracle positive).
    assert _CODE in _codes("Public Sub S()\n    Dim x As Int\nEnd Sub", _STD)


def test_reserved_identifier_as_type_fires() -> None:
    # `Loop` is a reserved VBA identifier, not a type.
    assert _CODE in _codes("Public Sub S()\n    Dim x As Loop\nEnd Sub", _STD)


def test_runtime_function_as_parameter_type_fires() -> None:
    assert _CODE in _codes("Public Sub S(ByVal a As Left)\nEnd Sub", _STD)


def test_known_non_type_name_fires() -> None:
    # A project declaration that is not a type (a Sub named like the As-name).
    helper = "Public Sub Widget()\nEnd Sub"
    entry = "Public Sub S()\n    Dim x As Widget\nEnd Sub"
    index = ProjectIndex()
    index.set_module(ModuleInput("Helpers", ModuleSymbolKind.STANDARD, helper))
    index.set_module(ModuleInput("Module1", ModuleSymbolKind.STANDARD, entry))
    opts = AnalyzeModuleOptions(
        module_name="Module1",
        module_kind=ModuleSymbolKind.STANDARD,
        known_non_type_names=index.visible_non_type_names("Module1"),
    )
    assert _CODE in _codes(entry, opts)


# -- controls (must stay silent) -------------------------------------------


def test_primitive_type_is_silent() -> None:
    # `Long` is reserved, but resolves as a primitive type -> no diagnostic.
    assert _CODE not in _codes("Public Sub S()\n    Dim x As Long\nEnd Sub", _STD)


def test_host_type_is_silent() -> None:
    assert _CODE not in _codes("Public Sub S()\n    Dim x As Workbook\nEnd Sub", _STD)


def test_host_type_sharing_runtime_name_is_silent() -> None:
    # `Filter` is both a VBA runtime function AND an Excel host type; the host type
    # wins, so it must not be flagged (FP-hunt regression guard).
    assert _CODE not in _codes("Public Sub S()\n    Dim x As Filter\nEnd Sub", _STD)


def test_same_module_user_type_is_silent() -> None:
    src = (
        "Private Type TPoint\n    X As Long\nEnd Type\n"
        "Public Sub S()\n    Dim p As TPoint\nEnd Sub"
    )
    assert _CODE not in _codes(src, _STD)


def test_same_module_enum_is_silent() -> None:
    src = (
        "Public Enum Color\n    Red\nEnd Enum\n"
        "Public Sub S()\n    Dim c As Color\nEnd Sub"
    )
    assert _CODE not in _codes(src, _STD)


def test_project_class_type_is_silent() -> None:
    person = "Public Sub Save()\nEnd Sub"
    entry = "Public Sub S()\n    Dim p As Person\nEnd Sub"
    index = ProjectIndex()
    index.set_module(ModuleInput("Person", ModuleSymbolKind.CLASS, person))
    index.set_module(ModuleInput("Module1", ModuleSymbolKind.STANDARD, entry))
    opts = AnalyzeModuleOptions(
        module_name="Module1",
        module_kind=ModuleSymbolKind.STANDARD,
        project_class_members=index.project_class_members(),
        known_non_type_names=index.visible_non_type_names("Module1"),
    )
    assert _CODE not in _codes(entry, opts)


def test_qualified_reference_unresolved_is_silent() -> None:
    # `Foo.Bar` resolves to nothing (no project module Foo) and the member `Bar` is
    # not reserved / runtime / a known non-type, so it stays silent — matching XLIDE.
    assert _CODE not in _codes("Public Sub S()\n    Dim x As Foo.Bar\nEnd Sub", _STD)


# -- new branches: ambiguous + invalid-new-type-name -----------------------


def test_ambiguous_project_type_fires() -> None:
    # A class module `Thing` and an Enum `Thing` in another module collide; the bare
    # `As Thing` reference is ambiguous across the project tier.
    mods = [
        ("Thing", ModuleSymbolKind.CLASS, "Public Sub M()\nEnd Sub\n"),
        ("ModB", ModuleSymbolKind.STANDARD, "Public Enum Thing\n    A = 1\nEnd Enum\n"),
        ("ModC", ModuleSymbolKind.STANDARD, "Public Sub S()\n    Dim x As Thing\nEnd Sub\n"),
    ]
    assert _CODE in _project_codes(mods, "ModC")


def test_new_user_type_fires() -> None:
    src = (
        "Public Type TPoint\n    X As Long\nEnd Type\n\n"
        "Public Sub S()\n    Dim p As New TPoint\nEnd Sub\n"
    )
    assert _NEW_CODE in _project_codes([("Module1", ModuleSymbolKind.STANDARD, src)], "Module1")


def test_new_enum_fires() -> None:
    src = (
        "Public Enum EColor\n    Red = 1\nEnd Enum\n\n"
        "Public Sub S()\n    Dim c As New EColor\nEnd Sub\n"
    )
    assert _NEW_CODE in _project_codes([("Module1", ModuleSymbolKind.STANDARD, src)], "Module1")


def test_new_creatable_class_is_silent() -> None:
    # `New Person` where Person is a project class module is legal (creatable).
    mods = [
        ("Person", ModuleSymbolKind.CLASS, "Public Sub Save()\nEnd Sub\n"),
        ("Module1", ModuleSymbolKind.STANDARD, "Public Sub S()\n    Dim p As New Person\nEnd Sub\n"),
    ]
    assert _NEW_CODE not in _project_codes(mods, "Module1")


def test_new_host_type_is_silent() -> None:
    # `New Collection` is a host object-model type; host types are exempt from the
    # New-creatable rule.
    src = "Public Sub S()\n    Dim c As New Collection\nEnd Sub\n"
    assert _NEW_CODE not in _project_codes([("Module1", ModuleSymbolKind.STANDARD, src)], "Module1")


def test_withevents_new_is_silent() -> None:
    # `WithEvents x As New T` is exempt from invalid-new-type-name even though it is
    # a New declaration; a creatable class is the valid target and stays silent.
    mods = [
        ("Sink", ModuleSymbolKind.CLASS, "Public Event Fired()\n"),
        ("Watcher", ModuleSymbolKind.CLASS, "Private WithEvents s As New Sink\n"),
    ]
    assert _NEW_CODE not in _project_codes(mods, "Watcher")


# -- oracle ----------------------------------------------------------------

# The asserted case (Int runtime function) is in the shipped runtime branch.
_SKIP_IDS: frozenset[str] = frozenset()


def test_oracle_asserted_cases() -> None:
    cases = asserted_cases(_CODE)
    assert cases, "expected at least one asserted case"
    checked = 0
    for case in cases:
        if case.id in _SKIP_IDS:
            continue
        emitted: set[str] = set()
        for module in case.modules:
            emitted |= _codes(module.source, _STD)
        if case.expected == "rejected":
            assert _CODE in emitted, f"{case.id}: expected {_CODE}, got {sorted(emitted)}"
        elif case.expected == "accepted":
            assert _CODE not in emitted, f"{case.id}: {_CODE} must not fire on accepted control"
        checked += 1
    assert checked > 0


def test_no_false_positives_on_accepted_cases() -> None:
    # Both invalid-as-type-name and invalid-new-type-name are compile-error
    # diagnostics, so every accepted case constrains them: they must never fire on
    # compile-valid code (evaluated with the full project context the harness threads in).
    for case in accepted_cases():
        offenders = oracle_false_positives(case, _CODES)
        assert not offenders, f"{case.id}: {sorted(offenders)} false positive"
