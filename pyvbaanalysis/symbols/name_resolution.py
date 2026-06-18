"""Context-aware source resolver for bare identifiers (MS-VBAL 5.3).

Ported from xlide_vscode/src/analyzer/symbols/nameResolution.ts. Ordered
local -> module -> project, with a tri-state ambiguous/unresolved outcome and
property-accessor-family collapsing. The XLIDE synthetic return-variable symbol
carries a `doc`; that is editor-only and omitted here (agent.md Risk 7).
"""

from __future__ import annotations

import enum
from collections.abc import Sequence
from dataclasses import dataclass, field

from .symbol_model import ModuleSymbols, VbaSymbol, VbaSymbolKind


class BareIdentifierContext(str, enum.Enum):
    EXPRESSION = "expression"
    CALL = "call"
    ASSIGNMENT_TARGET = "assignmentTarget"
    MEMBER_RECEIVER = "memberReceiver"
    TYPE_NAME = "typeName"
    NEW_EXPRESSION = "newExpression"


class BareIdentifierResolutionScope(str, enum.Enum):
    LOCAL = "local"
    MODULE = "module"
    PROJECT = "project"
    AMBIGUOUS = "ambiguous"
    UNRESOLVED = "unresolved"


@dataclass(slots=True)
class BareIdentifierResolution:
    name: str
    lower_name: str
    context: BareIdentifierContext
    scope: BareIdentifierResolutionScope
    definitions: list[VbaSymbol]
    reason: str
    tier: BareIdentifierResolutionScope | None = None


@dataclass(slots=True)
class BareIdentifierResolutionInput:
    current_module: ModuleSymbols
    name: str
    context: BareIdentifierContext
    enclosing_procedure: VbaSymbol | None = None
    offset: int | None = None
    project_visible_symbols: list[VbaSymbol] = field(default_factory=list)


def resolve_bare_identifier_binding(
    input: BareIdentifierResolutionInput,
) -> BareIdentifierResolution:
    lower_name = input.name.lower()
    local = local_identifier_matches(input.enclosing_procedure, lower_name, input.context, input.offset)
    if len(local) > 0:
        return _resolution(input, lower_name, _ambiguous_scope(local, BareIdentifierResolutionScope.LOCAL), local)

    module = module_level_identifier_matches(input.current_module, lower_name, input.context)
    if len(module) > 0:
        return _resolution(input, lower_name, _ambiguous_scope(module, BareIdentifierResolutionScope.MODULE), module)

    current_lower = input.current_module.module_name.lower()
    project = [
        symbol
        for symbol in input.project_visible_symbols
        if symbol.module_name.lower() != current_lower
        and symbol.name.lower() == lower_name
        and _symbol_allowed_in_context(symbol, input.context)
    ]
    if len(project) > 0:
        return _resolution(input, lower_name, _ambiguous_scope(project, BareIdentifierResolutionScope.PROJECT), project)

    return BareIdentifierResolution(
        name=input.name,
        lower_name=lower_name,
        context=input.context,
        scope=BareIdentifierResolutionScope.UNRESOLVED,
        definitions=[],
        reason=f"No source-backed {input.context.value} binding found for '{input.name}'.",
    )


def local_identifier_matches(
    procedure: VbaSymbol | None,
    lower_name: str,
    context: BareIdentifierContext,
    offset: int | None = None,
) -> list[VbaSymbol]:
    if (
        procedure is None
        or context is BareIdentifierContext.TYPE_NAME
        or context is BareIdentifierContext.NEW_EXPRESSION
    ):
        return []
    out: list[VbaSymbol] = []
    return_variable = _procedure_return_variable(procedure, context, offset)
    if return_variable is not None and return_variable.name.lower() == lower_name:
        out.append(return_variable)
    for symbol in procedure.children or []:
        if _is_local_identifier_symbol(symbol) and symbol.name.lower() == lower_name:
            out.append(symbol)
    return out


def module_level_identifier_matches(
    mod: ModuleSymbols, lower_name: str, context: BareIdentifierContext
) -> list[VbaSymbol]:
    out: list[VbaSymbol] = []
    for symbol in mod.root.children or []:
        if symbol.name.lower() == lower_name and _symbol_allowed_in_context(symbol, context):
            out.append(symbol)
        if symbol.kind is VbaSymbolKind.ENUM:
            for member in symbol.children or []:
                if member.name.lower() == lower_name and _symbol_allowed_in_context(member, context):
                    out.append(member)
    return out


def source_identifier_names(
    current_module: ModuleSymbols,
    enclosing_procedure: VbaSymbol | None = None,
    project_visible_symbols: Sequence[VbaSymbol] = (),
) -> set[str]:
    out: set[str] = set()
    return_variable = _procedure_return_variable(enclosing_procedure, BareIdentifierContext.EXPRESSION)
    if return_variable is not None:
        out.add(return_variable.name.lower())
    for symbol in (enclosing_procedure.children if enclosing_procedure is not None else None) or []:
        if _is_local_identifier_symbol(symbol):
            out.add(symbol.name.lower())
    for symbol in current_module.root.children or []:
        out.add(symbol.name.lower())
        if symbol.kind is VbaSymbolKind.ENUM:
            for member in symbol.children or []:
                out.add(member.name.lower())
    for symbol in project_visible_symbols:
        out.add(symbol.name.lower())
    return out


def _procedure_return_variable(
    procedure: VbaSymbol | None, context: BareIdentifierContext, offset: int | None = None
) -> VbaSymbol | None:
    if (
        procedure is None
        or not _procedure_returns_through_name(procedure)
        or context is BareIdentifierContext.CALL
    ):
        return None
    if offset is not None and offset <= procedure.name_span.end:
        return None
    return VbaSymbol(
        name=procedure.name,
        kind=VbaSymbolKind.LOCAL_VARIABLE,
        name_span=procedure.name_span,
        full_span=procedure.name_span,
        module_name=procedure.module_name,
        container_name=procedure.name,
        as_type=procedure.as_type,
        is_array=procedure.is_array,
    )


def _procedure_returns_through_name(procedure: VbaSymbol) -> bool:
    return procedure.kind is VbaSymbolKind.FUNCTION or procedure.kind is VbaSymbolKind.PROPERTY_GET


def _is_local_identifier_symbol(symbol: VbaSymbol) -> bool:
    return (
        symbol.kind is VbaSymbolKind.PARAMETER
        or symbol.kind is VbaSymbolKind.LOCAL_VARIABLE
        or (symbol.kind is VbaSymbolKind.CONSTANT and bool(symbol.container_name))
    )


def _symbol_allowed_in_context(symbol: VbaSymbol, context: BareIdentifierContext) -> bool:
    if context is BareIdentifierContext.TYPE_NAME or context is BareIdentifierContext.NEW_EXPRESSION:
        return symbol.kind is VbaSymbolKind.TYPE or symbol.kind is VbaSymbolKind.ENUM
    return True


def _ambiguous_scope(
    definitions: Sequence[VbaSymbol], fallback: BareIdentifierResolutionScope
) -> BareIdentifierResolutionScope:
    if len(definitions) <= 1 or _is_property_accessor_family(definitions):
        return fallback
    return BareIdentifierResolutionScope.AMBIGUOUS


def _is_property_accessor_family(definitions: Sequence[VbaSymbol]) -> bool:
    if len(definitions) <= 1:
        return False
    first = definitions[0]
    return all(
        symbol.module_name.lower() == first.module_name.lower()
        and symbol.name.lower() == first.name.lower()
        and symbol.kind
        in (VbaSymbolKind.PROPERTY_GET, VbaSymbolKind.PROPERTY_LET, VbaSymbolKind.PROPERTY_SET)
        for symbol in definitions
    )


def _resolution(
    input: BareIdentifierResolutionInput,
    lower_name: str,
    scope: BareIdentifierResolutionScope,
    definitions: list[VbaSymbol],
) -> BareIdentifierResolution:
    if scope is BareIdentifierResolutionScope.AMBIGUOUS:
        tier = _definition_tier(input.current_module, definitions)
    elif scope is BareIdentifierResolutionScope.UNRESOLVED:
        tier = None
    else:
        tier = scope
    owner = definitions[0].module_name if definitions else input.current_module.module_name
    label = (
        "ambiguous source-backed"
        if scope is BareIdentifierResolutionScope.AMBIGUOUS
        else f"{scope.value} source-backed"
    )
    return BareIdentifierResolution(
        name=input.name,
        lower_name=lower_name,
        context=input.context,
        scope=scope,
        tier=tier,
        definitions=definitions,
        reason=f"{label} {input.context.value} binding for '{input.name}' in {owner}.",
    )


def _definition_tier(
    current_module: ModuleSymbols, definitions: Sequence[VbaSymbol]
) -> BareIdentifierResolutionScope | None:
    if not definitions:
        return None
    first = definitions[0]
    if _is_local_identifier_symbol(first):
        return BareIdentifierResolutionScope.LOCAL
    return (
        BareIdentifierResolutionScope.MODULE
        if first.module_name.lower() == current_module.module_name.lower()
        else BareIdentifierResolutionScope.PROJECT
    )
