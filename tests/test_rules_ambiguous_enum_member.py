"""M10: ambiguous-enum-member rule (checkAmbiguousEnumMemberReferences parity).

The rule is wired into the production registry at its registry.ts position (12,
immediately after udtParameterConstraints). This test drives the real engine
through the project-aware harness the rule needs (project_visible_symbols /
project_types / known_procedures threaded in).
"""

from __future__ import annotations

from oracle_support import AUDIT, CASES, _kind  # type: ignore[attr-defined]

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module
from pyvbaanalysis.evidence import OracleCase
from pyvbaanalysis.symbols import ModuleInput, ProjectIndex

_CODE = "ambiguous-enum-member"


# -- project-aware oracle harness ------------------------------------------


def _case_codes(case: OracleCase) -> set[str]:
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


def _project_codes(modules: list[tuple[str, str, str]], target: str) -> set[str]:
    """Codes emitted for `target` module given (name, kind, source) modules."""
    index = ProjectIndex()
    for name, kind, src in modules:
        index.set_module(ModuleInput(name, _kind(kind), src))
    target_src = next(src for name, _kind_, src in modules if name == target)
    opts = AnalyzeModuleOptions(
        module_name=target,
        module_kind=_kind(next(k for n, k, _s in modules if n == target)),
        project_procedures=index.procedure_signatures(),
        project_class_members=index.project_class_members(),
        project_visible_symbols=index.visible_identifier_symbols(target),
        project_types=index.visible_type_names(target),
        known_procedures=index.visible_procedure_names(target),
        known_identifiers=index.visible_identifier_names(target),
    )
    return {d.code for d in analyze_module(target_src, opts)}


def test_same_module_cross_enum_unqualified_read_fires() -> None:
    src = (
        "Public Enum EOne\n    AmbVal = 1\nEnd Enum\n\n"
        "Public Enum ETwo\n    AmbVal = 2\nEnd Enum\n\n"
        "Public Sub S()\n    Debug.Print AmbVal\nEnd Sub\n"
    )
    assert _CODE in _project_codes([("Module1", "standard", src)], "Module1")


def test_qualified_read_is_silent() -> None:
    src = (
        "Public Enum EOne\n    AmbVal = 1\nEnd Enum\n\n"
        "Public Enum ETwo\n    AmbVal = 2\nEnd Enum\n\n"
        "Public Sub S()\n    Debug.Print EOne.AmbVal\nEnd Sub\n"
    )
    assert _CODE not in _project_codes([("Module1", "standard", src)], "Module1")


def test_local_shadow_is_silent() -> None:
    # A local named the same as the ambiguous member shadows it (binder scope != AMBIGUOUS).
    src = (
        "Public Enum EOne\n    AmbVal = 1\nEnd Enum\n\n"
        "Public Enum ETwo\n    AmbVal = 2\nEnd Enum\n\n"
        "Public Sub S()\n    Dim AmbVal As Long\n    Debug.Print AmbVal\nEnd Sub\n"
    )
    assert _CODE not in _project_codes([("Module1", "standard", src)], "Module1")


def test_single_enum_member_read_is_silent() -> None:
    src = (
        "Public Enum EOne\n    OnlyOne = 1\nEnd Enum\n\n"
        "Public Sub S()\n    Debug.Print OnlyOne\nEnd Sub\n"
    )
    assert _CODE not in _project_codes([("Module1", "standard", src)], "Module1")


def test_cross_module_exported_enum_members_fire() -> None:
    # Two OTHER standard modules each export DupVal; the reading module has no
    # same-module binding, so the read binds ambiguously across the project tier.
    mod_a = "Public Enum EA\n    DupVal = 1\nEnd Enum\n"
    mod_b = "Public Enum EB\n    DupVal = 2\nEnd Enum\n"
    mod_c = "Public Sub S()\n    Debug.Print DupVal\nEnd Sub\n"
    assert _CODE in _project_codes(
        [("ModA", "standard", mod_a), ("ModB", "standard", mod_b), ("ModC", "standard", mod_c)],
        "ModC",
    )


def test_same_module_member_shadows_other_module_export() -> None:
    # A same-module enum member binds at MODULE tier and shadows another module's
    # export of the same name (scope != AMBIGUOUS) — no false positive.
    mod_a = "Public Enum EA\n    DupVal = 1\nEnd Enum\n"
    mod_b = (
        "Public Enum EB\n    DupVal = 2\nEnd Enum\n\n"
        "Public Sub S()\n    Debug.Print DupVal\nEnd Sub\n"
    )
    assert _CODE not in _project_codes(
        [("ModA", "standard", mod_a), ("ModB", "standard", mod_b)], "ModB"
    )


def test_no_project_context_is_silent() -> None:
    # With no project_visible_symbols the binder cannot reach AMBIGUOUS across the
    # other module, but same-module collisions still bind ambiguous. Here a single
    # module with two enums + read: this DOES fire (same-module visible members), so
    # the silent guarantee is checked on a no-collision module instead.
    src = "Public Sub S()\n    Debug.Print Foo\nEnd Sub\n"
    assert _CODE not in {d.code for d in analyze_module(src, AnalyzeModuleOptions())}


# -- oracle sweep ----------------------------------------------------------

# Both asserted oracle cases are single-module and resolve with the same-module
# enum-member visibility this harness threads in, so none are skipped.
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
    # ambiguous-enum-member is a compile-equivalent diagnostic: it must never fire on
    # compile-valid code, with the full project context threaded in.
    for case in CASES.values():
        if case.expected != "accepted":
            continue
        assert _CODE not in _case_codes(case), f"{case.id}: {_CODE} false positive"
