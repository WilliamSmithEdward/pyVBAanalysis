"""Project-level orchestration over analyze_module.

The diagnostics engine (analyze_module) is per module: it takes one module's
source plus an AnalyzeModuleOptions carrying the cross-module project context. A
project (a whole VBA project, a workbook, or a folder of exported modules) is a set
of modules that can see each other, so a faithful project pass has to index every
module first and then derive each module's options from that index. This module
provides that glue:

* analyze_module_options_for: turn a populated ProjectIndex into the per-module
  AnalyzeModuleOptions (the project_* / known_* fields the cross-module rules read).
* build_project_index: register a set of modules on a fresh ProjectIndex.
* analyze_project: index every module, then analyze all of them (or a named subset)
  with full cross-module context.

The field mapping in analyze_module_options_for mirrors the wiring exercised by the
oracle test harness, so a project pass sees exactly the cross-module context the
rule suite is validated against.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping

from .conditional import ConditionalCompilationEnvironment
from .diagnostics import AnalyzeModuleOptions, VbaDiagnostic, analyze_module
from .symbols import ModuleInput, ModuleSymbolKind, ProjectIndex, ProjectIndexOptions


def analyze_module_options_for(
    index: ProjectIndex,
    module_name: str,
    module_kind: ModuleSymbolKind,
    *,
    severity_overrides: Mapping[str, str] | None = None,
    conditional_compilation: ConditionalCompilationEnvironment | None = None,
) -> AnalyzeModuleOptions:
    """Build the AnalyzeModuleOptions for one module from a populated ProjectIndex.

    This is the glue that converts the index's cross-module query surface into the
    project_* / known_* option fields analyze_module reads. Every module in the
    project must already be registered on the index (via set_module) so the
    visibility queries see the whole project, not just this module.
    ``conditional_compilation`` sets the #If/#Const baseline for this module's
    diagnostic pass (the index uses it for symbol building separately).
    """
    return AnalyzeModuleOptions(
        module_name=module_name,
        module_kind=module_kind,
        severity_overrides=severity_overrides,
        conditional_compilation=conditional_compilation,
        project_procedures=index.procedure_signatures(),
        project_class_members=index.project_class_members(),
        project_integer_constants=index.visible_external_integer_constant_expressions(module_name),
        project_visible_symbols=index.visible_identifier_symbols(module_name),
        project_types=index.visible_type_names(module_name),
        known_procedures=index.visible_procedure_names(module_name),
        known_identifiers=index.visible_identifier_names(module_name),
        known_non_type_names=index.visible_non_type_names(module_name),
    )


def build_project_index(
    modules: Iterable[ModuleInput], *, options: ProjectIndexOptions | None = None
) -> ProjectIndex:
    """A ProjectIndex with every given module registered.

    ``options`` carries project-wide settings, notably a default
    ConditionalCompilationEnvironment applied to any module that does not set its
    own (per-module ModuleInput.conditional_compilation still wins).
    """
    index = ProjectIndex(options)
    for module in modules:
        index.set_module(module)
    return index


def _ensure_unique_module_names(modules: list[ModuleInput]) -> None:
    """Raise ValueError if two modules share a name (case-insensitively).

    The index keys modules by lowercased name, so duplicates would overwrite each
    other and silently drop a module's diagnostics. VBA module names are unique
    within a project, so a collision means the caller assembled the project wrong.
    """
    seen: set[str] = set()
    duplicates: set[str] = set()
    for module in modules:
        key = module.module_name.lower()
        if key in seen:
            duplicates.add(module.module_name)
        seen.add(key)
    if duplicates:
        raise ValueError(
            "duplicate module name(s) in project: "
            + ", ".join(sorted(duplicates))
            + "; module names must be unique (case-insensitive)."
        )


def analyze_project(
    modules: Iterable[ModuleInput],
    *,
    only: Iterable[str] | None = None,
    severity_overrides: Mapping[str, str] | None = None,
    conditional_compilation: ConditionalCompilationEnvironment | None = None,
) -> dict[str, list[VbaDiagnostic]]:
    """Analyze a whole VBA project with full cross-module context.

    Every module is registered on one shared ProjectIndex so the cross-module rules
    (project procedures, class members, type names, visible identifiers) resolve,
    then each module is analyzed with options derived from that index. Pass ``only``
    to analyze just the named modules (matched case-insensitively) while still
    indexing the whole project for context; ``only=None`` analyzes every module.

    ``conditional_compilation`` sets a project-wide compiler-constant and #Const
    baseline for the #If/#Const directives (a module's own
    ModuleInput.conditional_compilation still wins). Without it the built-in defaults
    apply (VBA7 and Win64 true, Win32 and Mac false).

    Returns a dict mapping module name to that module's diagnostics, preserving the
    input order of the analyzed modules. Module names must be unique within the
    project (case-insensitive); a collision raises ValueError rather than silently
    dropping a module.
    """
    module_list = list(modules)
    _ensure_unique_module_names(module_list)
    options = (
        ProjectIndexOptions(conditional_compilation=conditional_compilation)
        if conditional_compilation is not None
        else None
    )
    index = build_project_index(module_list, options=options)
    selected = None if only is None else {name.lower() for name in only}
    results: dict[str, list[VbaDiagnostic]] = {}
    for module in module_list:
        if selected is not None and module.module_name.lower() not in selected:
            continue
        opts = analyze_module_options_for(
            index,
            module.module_name,
            module.module_kind,
            severity_overrides=severity_overrides,
            conditional_compilation=module.conditional_compilation or conditional_compilation,
        )
        results[module.module_name] = analyze_module(module.source, opts)
    return results
