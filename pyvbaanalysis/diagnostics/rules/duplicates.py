"""Rule family: duplicate and ambiguous declarations.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/duplicates.ts. The
cross-module ambiguous-enum-member rule (checkAmbiguousEnumMemberReferences)
depends on type inference, host, and runtime resolution and is deferred to M8/M9;
the five duplicate-declaration rules here are self-contained.
"""

from __future__ import annotations

from collections.abc import Sequence

from ...conditional import ConditionalActivity, ConditionalActivityTracker
from ...parser.nodes import EnumNode, ModuleNode, TypeNode
from ...symbols.symbol_model import VbaSymbol, VbaSymbolKind, is_procedure_kind
from ..context import PushFn
from ..walker import active_module_members
from .shared import declaration_name_hit

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
