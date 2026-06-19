"""M9: invalidAsTypeNames rule (declarations.ts parity, SAFE branches only).

Ships the self-contained fallback branches (invalid-as-type-name): a type-position
name that resolves to NO known type AND is a reserved VBA identifier, a VBA runtime
function, or a known project non-type declaration. DEFERS (no-op): the ambiguous
project-type branch, the invalidNewTypeName creatable-type branch, and qualified
references — all need the project-type/external registry the port lacks. A
conservative type-resolution gate (primitives + host aliases + OLE + same-module
Enum/Type + project class/document/userform + project_types when supplied)
suppresses false positives on real types (e.g. `Long`, host `Filter`).

The one asserted oracle case (runtime_function_as_type_name_int -> a `Dim x As Int`
where `Int` is a VBA runtime function) is in the shipped runtime branch, so no
skip_ids. The rule is wired in the real registry, so a plain analyze_module
exercises it.
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


_STD = AnalyzeModuleOptions(module_name="Module1", module_kind=ModuleSymbolKind.STANDARD)


def _codes(source: str, opts: AnalyzeModuleOptions | None = None) -> set[str]:
    return {d.code for d in analyze_module(source, opts)}


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


def test_qualified_reference_is_silent() -> None:
    # Qualified names are deferred (need the registry) -> no diagnostic.
    assert _CODE not in _codes("Public Sub S()\n    Dim x As Foo.Bar\nEnd Sub", _STD)


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
    # invalid-as-type-name is a compile-error diagnostic, so every accepted case
    # constrains it: it must never fire on compile-valid code.
    for case in accepted_cases():
        offenders = oracle_false_positives(case, (_CODE,))
        assert not offenders, f"{case.id}: {_CODE} false positive"
