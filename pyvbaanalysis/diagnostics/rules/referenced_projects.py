"""Rule: a name reached through another VBA project the project references must be
one that project offers.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/referencedProjects.ts (XLIDE
issue #1503). Measured through the VBE oracle in Excel, Word, PowerPoint and Access
16.0, with a referenced project ProbeLibrary whose standard module Service has a
Public Function Value: `ProbeLibrary.Nope`, `ProbeLibrary.Service.Nope()` and, with
Service marked `Option Private Module`, `ProbeLibrary.Service.Value()` and
`ProbeLibrary.Value()` are each "Method or data member not found". So is
`Service.Nope()` while Service is the referenced project's.

Only referenced projects that were read are judged (see
symbols/referenced_projects.py), and a name the project itself declares or sees
comes first: a local, module or project name Service is not the referenced module.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...host import resolve_host_constant, resolve_host_global
from ...host.host_model import HostObjectModel, resolve_host_global_member
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span, VariableGroupNode
from ...runtime import (
    resolve_runtime_constant,
    resolve_runtime_function,
    resolve_runtime_object,
    resolve_vba_library_qualifier,
)
from ...symbols.referenced_projects import (
    ReferencedLookupOrder,
    ReferencedProjectModule,
    ReferencedProjectNames,
    ReferencedProjectSurface,
    referenced_project_names,
    referenced_project_reached,
)
from ...symbols.symbol_model import ModuleSymbolKind, VbaSymbol
from ..context import PushFn, statement_tokens
from ..walker import ProcedureStatementVisitor, absolute_span, for_each_variable_group, token_name
from .declarations import type_name_qualifier_starts


@dataclass(frozen=True, slots=True)
class _ReferencedNames:
    """The referenced projects and the standard modules they show, lowercased, for
    one module's analysis."""

    projects: dict[str, ReferencedProjectSurface]
    # A visible standard module of a referenced project, by name, and the project it is in.
    modules: dict[str, tuple[ReferencedProjectSurface, ReferencedProjectModule]]
    # Every Enum a referenced project offers, lowercased: which bare `X.Y` to look up
    # in reference order.
    enum_names: frozenset[str]


def _referenced_names(surfaces: Sequence[ReferencedProjectSurface]) -> _ReferencedNames:
    projects = {surface.name.lower(): surface for surface in surfaces}
    modules: dict[str, tuple[ReferencedProjectSurface, ReferencedProjectModule]] = {}
    for project in surfaces:
        for module in referenced_project_names(project).modules.values():
            key = module.name.lower()
            if module.kind is ModuleSymbolKind.STANDARD and module.visible and key not in modules:
                modules[key] = (project, module)
    enum_names = frozenset(name for project in surfaces for name in referenced_project_names(project).enums)
    return _ReferencedNames(projects, modules, enum_names)


def _has_member(module: ReferencedProjectModule, name: str) -> bool:
    lower = name.lower()
    return any(member.lower() == lower for member in module.members)


def _offers_bare(lower: str) -> Callable[[ReferencedProjectNames], bool]:
    """Whether a read project offers a bare name, as a value or a type."""
    return lambda offered: lower in offered.bare_names or lower in offered.types


def _is_vba_or_host_value(lower: str, model: HostObjectModel | None) -> bool:
    """A name VBA or the host gives as a value, ahead of any reference: `Strings`,
    Excel's `ActiveWorkbook`."""
    return (
        resolve_runtime_function(lower) is not None
        or resolve_runtime_constant(lower) is not None
        or resolve_runtime_object(lower) is not None
        or resolve_vba_library_qualifier(lower) is not None
        or resolve_host_global(lower, model) is not None
        or resolve_host_global_member(lower, model) is not None
        or resolve_host_constant(lower, model) is not None
    )


def check_referenced_project_names(
    source: str,
    mod: ModuleNode,
    surfaces: Sequence[ReferencedProjectSurface] | None,
    module_symbols: Sequence[VbaSymbol],
    project_visible_symbols: Sequence[VbaSymbol] | None,
    project_module_names: AbstractSet[str] | None,
    order: ReferencedLookupOrder,
    host_model: HostObjectModel | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """`ProbeLibrary.Nope`, `ProbeLibrary.Service.Nope()`, `Service.Nope()` and
    `WidgetColor.Nope`: "Method or data member not found". The port takes the parsed
    module as well, where upstream parses the source again for its type references."""
    if not surfaces:
        return lambda procedure: None
    names = _referenced_names(surfaces)
    # What the project itself names wins over a referenced project's name.
    project_names = {
        *(symbol.name.lower() for symbol in module_symbols),
        *(symbol.name.lower() for symbol in project_visible_symbols or ()),
        *(name.lower() for name in project_module_names or ()),
    }

    def report(text: str, span: Span) -> None:
        push("memberNotFound", f"Method or data member not found: '{text}'.", span)

    # The qualifiers of names in a type position: `TypeOf o Is ProbeLibrary.Widget`.
    type_qualifiers: list[frozenset[int]] = []

    def names_a_type(offset: int) -> bool:
        if not type_qualifiers:
            type_qualifiers.append(type_name_qualifier_starts(source, mod))
        return offset in type_qualifiers[0]

    def factory(procedure: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        own = {param.name.lower() for param in procedure.params}

        def note_group(group: VariableGroupNode) -> None:
            for decl in group.declarations:
                own.add(decl.name.lower())

        for_each_variable_group(procedure.body, note_group)

        def shadowed(lower: str) -> bool:
            return lower in own or lower in project_names

        def visit(stmt: LeafStatementNode) -> None:
            toks = statement_tokens(source, stmt.span)

            def raw_at(index: int) -> str | None:
                return toks[index].raw_text if 0 <= index < len(toks) else None

            def third_name(i: int) -> str | None:
                return token_name(toks[i + 4]) if raw_at(i + 3) == "." and i + 4 < len(toks) else None

            for i in range(len(toks) - 2):
                head = token_name(toks[i])
                if not head or raw_at(i - 1) in (".", "!") or toks[i + 1].raw_text != ".":
                    continue
                second = token_name(toks[i + 2])
                if not second:
                    continue
                lower = head.lower()
                if shadowed(lower):
                    continue
                project = names.projects.get(lower)
                if project is not None:
                    # `ReDim a(1) As ProbeLibrary.Widget` names a type too.
                    previous = raw_at(i - 1)
                    if (previous is not None and previous.lower() == "as") or names_a_type(
                        absolute_span(stmt.span, toks[i]).start
                    ):
                        continue
                    offered = referenced_project_names(project)
                    module = offered.modules.get(second.lower())
                    if module is not None and module.visible:
                        if module.kind is ModuleSymbolKind.CLASS and not module.predeclared:
                            # ProbeLibrary.Widget, a class with no default instance, is no value.
                            report(f"{head}.{second}", absolute_span(stmt.span, toks[i + 2]))
                            continue
                        # ProbeLibrary.Service.Nope(): the module's own members.
                        third = third_name(i)
                        # A document module's members are its host object's too: unjudged.
                        if module.kind is ModuleSymbolKind.STANDARD:
                            judged = _has_member(module, third or "")
                        else:
                            judged = module.kind is not ModuleSymbolKind.CLASS or (
                                third is not None
                                and any(member.name.lower() == third.lower() for member in module.object_members or ())
                            )
                        if third and not judged:
                            report(f"{head}.{second}.{third}", absolute_span(stmt.span, toks[i + 4]))
                        continue
                    # ProbeLibrary.WidgetColor.wcRed: an Enum's own members.
                    enum_decl = offered.enums.get(second.lower())
                    if enum_decl is not None:
                        third = third_name(i)
                        if third and not any(member.lower() == third.lower() for member in enum_decl.members):
                            report(f"{head}.{second}.{third}", absolute_span(stmt.span, toks[i + 4]))
                        continue
                    # ProbeLibrary.Value(): a member of any of its visible standard
                    # modules, or ProbeLibrary.wcRed, an exposed class's Enum member.
                    if second.lower() not in offered.values:
                        report(f"{head}.{second}", absolute_span(stmt.span, toks[i + 2]))
                    continue
                referenced = names.modules.get(lower)
                # WidgetColor.Nope, where the Enum is what the bare name reaches first:
                # ahead of a later project's module of the name, and behind a value VBA
                # or the host gives, or a library that might have it.
                reached = (
                    referenced_project_reached(lower, surfaces, order, _offers_bare(lower))
                    if lower in names.enum_names
                    else None
                )
                enum_decl = (
                    referenced_project_names(reached).enums.get(lower)
                    if reached is not None and (referenced is None or reached is not referenced[0])
                    else None
                )
                if enum_decl is not None:
                    if not _is_vba_or_host_value(lower, host_model) and not any(
                        member.lower() == second.lower() for member in enum_decl.members
                    ):
                        report(f"{head}.{second}", absolute_span(stmt.span, toks[i + 2]))
                    continue
                # Service.Nope(), Service being the referenced project's.
                if referenced is not None and not _has_member(referenced[1], second):
                    report(f"{head}.{second}", absolute_span(stmt.span, toks[i + 2]))

        return visit

    return factory
