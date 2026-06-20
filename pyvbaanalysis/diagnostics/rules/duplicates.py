"""Rule family: duplicate and ambiguous declarations.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/duplicates.ts. Covers the
five duplicate-declaration rules plus the cross-module ambiguous-enum-member rule
(checkAmbiguousEnumMemberReferences): an unqualified read of a member name shared
by more than one visible Enum is the VBA "Ambiguous name detected" compile error.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from collections.abc import Set as AbstractSet

from ...conditional import ConditionalActivity, ConditionalActivityTracker
from ...host import application_member_names, resolve_host_global
from ...parser.nodes import EnumNode, ModuleNode, ProcedureNode, Span, TypeNode
from ...runtime import resolve_runtime_function, resolve_runtime_object
from ...symbols.name_resolution import (
    BareIdentifierContext,
    BareIdentifierResolution,
    BareIdentifierResolutionScope,
)
from ...symbols.symbol_model import (
    ModuleSymbols,
    VbaProcedureSignature,
    VbaProjectClassMembers,
    VbaSymbol,
    VbaSymbolKind,
    is_procedure_kind,
)
from ...types.type_inference import (
    procedure_symbol_for,
    source_identifier_binding,
    source_identifier_bound,
)
from ..callable_signatures import callable_type_signatures_for
from ..context import PushFn
from ..walker import active_module_members
from .shared import (
    declaration_name_hit,
    for_each_undeclared_reference_span,
    value_read_references,
)

_PROPERTY_KINDS = (
    VbaSymbolKind.PROPERTY_GET,
    VbaSymbolKind.PROPERTY_LET,
    VbaSymbolKind.PROPERTY_SET,
)
_LOCAL_DECL_KINDS = (
    VbaSymbolKind.PARAMETER,
    VbaSymbolKind.LOCAL_VARIABLE,
    VbaSymbolKind.CONSTANT,
)


def check_duplicate_procedures(members: Sequence[VbaSymbol], push: PushFn) -> None:
    """A name may be one Sub/Function OR a set of distinct Property accessors."""
    groups: dict[str, list[VbaSymbol]] = {}
    for sym in members:
        if not is_procedure_kind(sym.kind):
            continue
        groups.setdefault(sym.name.lower(), []).append(sym)

    for group in groups.values():
        if len(group) < 2:
            continue
        value_proc_seen = False
        accessor_seen: set[VbaSymbolKind] = set()
        for sym in group:
            is_property = sym.kind in _PROPERTY_KINDS
            if not is_property:
                conflict = value_proc_seen or len(accessor_seen) > 0
                value_proc_seen = True
            else:
                conflict = value_proc_seen or sym.kind in accessor_seen
                accessor_seen.add(sym.kind)
            if conflict:
                push(
                    "duplicateProcedure",
                    f"Ambiguous name detected: '{sym.name}' is already declared in this module.",
                    sym.name_span,
                )


def check_duplicate_declarations(members: Sequence[VbaSymbol], push: PushFn) -> None:
    """Within one procedure, a name is declared once across params/locals/consts."""
    for proc in members:
        if not is_procedure_kind(proc.kind):
            continue
        seen: set[str] = set()
        for child in proc.children or []:
            if child.kind not in _LOCAL_DECL_KINDS:
                continue
            key = child.name.lower()
            if key in seen:
                push(
                    "duplicateDeclaration",
                    f"Duplicate declaration in current scope: '{child.name}'.",
                    child.name_span,
                )
            else:
                seen.add(key)


def check_duplicate_module_members(members: Sequence[VbaSymbol], push: PushFn) -> None:
    """A module-level variable or constant declared more than once."""
    seen: set[str] = set()
    for sym in members:
        if sym.kind not in (VbaSymbolKind.MODULE_VARIABLE, VbaSymbolKind.CONSTANT):
            continue
        key = sym.name.lower()
        if key in seen:
            push(
                "duplicateModuleMember",
                f"Duplicate declaration: '{sym.name}' is already declared at module level.",
                sym.name_span,
            )
        else:
            seen.add(key)


def check_duplicate_enum_members(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """Member names inside one Enum block must be unique."""
    for member in active_module_members(mod, activity):
        if not isinstance(member, EnumNode):
            continue
        seen: set[str] = set()
        for enum_member in member.members:
            # Only provably-active members can collide: an inactive or
            # not-provably-active #If branch member is not guaranteed compiled
            # alongside another same-named member.
            if activity is not None and activity.activity_for_span(enum_member.span) is not ConditionalActivity.ACTIVE:
                continue
            key = enum_member.name.lower()
            hit = declaration_name_hit(source, enum_member.span, enum_member.name)
            if key in seen:
                push(
                    "duplicateEnumMember",
                    f"Duplicate Enum member '{enum_member.name}' in Enum '{member.name}'.",
                    hit.span if hit is not None else enum_member.span,
                )
            else:
                seen.add(key)


def check_duplicate_type_fields(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """Field names inside one Type (UDT) block must be unique."""
    for member in active_module_members(mod, activity):
        if not isinstance(member, TypeNode):
            continue
        seen: set[str] = set()
        for field_node in member.fields:
            if activity is not None and activity.activity_for_span(field_node.span) is not ConditionalActivity.ACTIVE:
                continue
            key = field_node.name.lower()
            hit = declaration_name_hit(source, field_node.span, field_node.name)
            if key in seen:
                push(
                    "duplicateTypeField",
                    f"Duplicate field '{field_node.name}' in Type '{member.name}'.",
                    hit.span if hit is not None else field_node.span,
                )
            else:
                seen.add(key)


def check_ambiguous_enum_member_references(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    module_name: str,
    known_procedures: AbstractSet[str] | None,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_members: Sequence[VbaProjectClassMembers] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> None:
    """An unqualified read of a member name shared by more than one visible Enum is
    rejected as "Ambiguous name detected".

    Same-module bindings take precedence over exported members from other modules,
    and procedure locals/parameters shadow module-level enum members. The
    no-false-positive hinge is the binder: a read that binds to a local/param or any
    non-enumMember symbol resolves with scope != AMBIGUOUS and stays silent; only an
    all-enumMember binding owned by more than one distinct module:container fires.
    """
    visible_enum_members = [
        *_enum_member_symbols(symbols.root.children or []),
        *[
            sym
            for sym in (project_visible_symbols or [])
            if sym.kind is VbaSymbolKind.ENUM_MEMBER
            and sym.module_name.lower() != module_name.lower()
        ],
    ]
    if len(_ambiguous_enum_member_groups(visible_enum_members)) == 0:
        return

    module_signatures = callable_type_signatures_for(symbols, project_procedures)
    app_members = application_member_names()
    known = {name.lower() for name in (known_procedures or ())}

    def is_known_for_skip(name: str, proc_sym: VbaSymbol | None) -> bool:
        lower = name.lower()
        return (
            source_identifier_bound(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.EXPRESSION
            )
            or lower in known
            or lower in app_members
            or resolve_host_global(name) is not None
            or resolve_runtime_object(name) is not None
            or resolve_runtime_function(name) is not None
        )

    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        proc_sym = procedure_symbol_for(symbols, member)
        reported: set[str] = set()

        def visit(
            span: Span,
            proc_sym: VbaSymbol | None = proc_sym,
            reported: set[str] = reported,
        ) -> None:
            def is_skipped(name: str) -> bool:
                return is_known_for_skip(name, proc_sym)

            for ref in value_read_references(
                source,
                span,
                is_skipped,
                module_signatures,
                project_members,
            ):
                binding = source_identifier_binding(
                    symbols,
                    proc_sym,
                    project_visible_symbols,
                    ref.name,
                    BareIdentifierContext.EXPRESSION,
                )
                definitions = _ambiguous_enum_member_definitions(binding)
                if definitions is None:
                    continue
                key = f"{ref.span.start}:{ref.span.end}"
                if key in reported:
                    continue
                reported.add(key)
                owners: list[str] = []
                for definition in definitions:
                    owner = definition.container_name or definition.module_name
                    if owner not in owners:
                        owners.append(owner)
                owner_text = ", ".join(owners[:3])
                detail = f" ({owner_text})" if owner_text else ""
                push(
                    "ambiguousEnumMember",
                    f"Ambiguous Enum member reference: '{ref.name}' is defined by "
                    f"multiple visible Enums{detail}. Qualify the reference with an "
                    f"Enum or module name.",
                    ref.span,
                )

        for_each_undeclared_reference_span(source, member.body, visit, activity)


def _enum_member_symbols(symbols: Sequence[VbaSymbol]) -> list[VbaSymbol]:
    out: list[VbaSymbol] = []
    for symbol in symbols:
        if symbol.kind is VbaSymbolKind.ENUM:
            out.extend(
                child
                for child in (symbol.children or [])
                if child.kind is VbaSymbolKind.ENUM_MEMBER
            )
    return out


def _ambiguous_enum_member_groups(
    symbols: Sequence[VbaSymbol],
) -> dict[str, list[VbaSymbol]]:
    groups: dict[str, dict[str, VbaSymbol]] = {}
    for symbol in symbols:
        if symbol.kind is not VbaSymbolKind.ENUM_MEMBER:
            continue
        key = symbol.name.lower()
        owner_key = f"{symbol.module_name.lower()}:{(symbol.container_name or '').lower()}"
        owners = groups.setdefault(key, {})
        owners.setdefault(owner_key, symbol)
    return {key: list(owners.values()) for key, owners in groups.items() if len(owners) > 1}


def _ambiguous_enum_member_definitions(
    binding: BareIdentifierResolution,
) -> list[VbaSymbol] | None:
    if binding.scope is not BareIdentifierResolutionScope.AMBIGUOUS:
        return None
    if any(d.kind is not VbaSymbolKind.ENUM_MEMBER for d in binding.definitions):
        return None
    owner_keys = {
        f"{d.module_name.lower()}:{(d.container_name or '').lower()}" for d in binding.definitions
    }
    return list(binding.definitions) if len(owner_keys) > 1 else None
