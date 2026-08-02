"""Source-binding shape inference for diagnostics rules.

Ported from the host-free slice of
xlide_vscode/src/analyzer/diagnostics/typeInference.ts (declaredShapeForSourceBinding)
plus procedureSymbolFor from analysisContext.ts. Resolves a bare identifier to the
declared shape (as-type, array-ness, fixed-vs-dynamic) of its source binding using
only the symbol graph. The host/completion-coupled inference lands in M8.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

from ..identity_cache import IdentityLru
from ..parser.nodes import ProcedureNode, ProcKind
from ..symbols.name_resolution import (
    BareIdentifierContext,
    BareIdentifierResolution,
    BareIdentifierResolutionInput,
    BareIdentifierResolutionScope,
    resolve_bare_identifier_binding,
)
from ..symbols.symbol_model import ModuleSymbols, VbaSymbol, VbaSymbolKind

_PROCEDURE_KINDS = frozenset(
    {
        VbaSymbolKind.SUB,
        VbaSymbolKind.FUNCTION,
        VbaSymbolKind.PROPERTY_GET,
        VbaSymbolKind.PROPERTY_LET,
        VbaSymbolKind.PROPERTY_SET,
    }
)


@dataclass(frozen=True, slots=True)
class DeclaredValueShape:
    as_type: str | None
    is_array: bool
    is_fixed_array: bool


@dataclass(frozen=True, slots=True)
class SourceDeclaredShape:
    resolved: bool
    shape: DeclaredValueShape | None = None


@dataclass(frozen=True, slots=True)
class SourceDeclaredType:
    resolved: bool
    as_type: str | None = None


_VALUE_DECLARATION_KINDS = frozenset(
    {
        VbaSymbolKind.PARAMETER,
        VbaSymbolKind.LOCAL_VARIABLE,
        VbaSymbolKind.MODULE_VARIABLE,
        VbaSymbolKind.CONSTANT,
    }
)


# Environments are requested once per rule per procedure, so both the finished
# per-procedure environments and their procedure-independent module-level bases
# are memoized by identity: rebuilding them per request was O(rules x
# procedures x declarations) on large modules. Consumers treat the returned
# dicts as read-only.
_SHAPE_ENV_CACHE = IdentityLru(capacity=64)
_SHAPE_ENV_MODULE_BASE_CACHE = IdentityLru()
_TYPE_ENV_CACHE = IdentityLru(capacity=64)
_TYPE_ENV_MODULE_BASE_CACHE = IdentityLru()


def _shape_env_module_base(symbols: ModuleSymbols) -> dict[str, DeclaredValueShape]:
    cached = _SHAPE_ENV_MODULE_BASE_CACHE.get(symbols)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    base: dict[str, DeclaredValueShape] = {}
    for sym in symbols.root.children or []:
        if sym.kind in _VALUE_DECLARATION_KINDS:
            base[sym.name.lower()] = _shape_of(sym)
    return _SHAPE_ENV_MODULE_BASE_CACHE.put(base, symbols)  # type: ignore[no-any-return]


def declaration_shape_environment_for(
    symbols: ModuleSymbols, proc: ProcedureNode
) -> dict[str, DeclaredValueShape]:
    """Syntactic declared-shape fallback for a procedure's module + local value names.

    The function-return-variable shape (an Erase/assignment target named after the
    procedure) is omitted; that is a precision-only gap, never a false positive.
    """
    cached = _SHAPE_ENV_CACHE.get(symbols, proc)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    out = dict(_shape_env_module_base(symbols))
    proc_sym = procedure_symbol_for(symbols, proc)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if child.kind in _VALUE_DECLARATION_KINDS:
            out[child.name.lower()] = _shape_of(child)
    return _SHAPE_ENV_CACHE.put(out, symbols, proc)  # type: ignore[no-any-return]


def _shape_of(sym: VbaSymbol) -> DeclaredValueShape:
    return DeclaredValueShape(
        as_type=sym.as_type,
        is_array=sym.is_array is True,
        is_fixed_array=sym.array_bounds is not None,
    )


def same_module_type_names(symbols: ModuleSymbols) -> set[str]:
    """Lowercased names of user-defined Type declarations in this module."""
    return {
        sym.name.lower()
        for sym in (symbols.root.children or [])
        if sym.kind is VbaSymbolKind.TYPE
    }


def _return_assignment_type_for(proc: ProcedureNode) -> str | None:
    if proc.proc_kind in (ProcKind.FUNCTION, ProcKind.PROPERTY_GET):
        return proc.return_type
    return None


def _type_env_module_base(symbols: ModuleSymbols) -> dict[str, str]:
    cached = _TYPE_ENV_MODULE_BASE_CACHE.get(symbols)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    base: dict[str, str] = {}
    for sym in symbols.root.children or []:
        if sym.as_type and sym.kind not in _PROCEDURE_KINDS:
            base[sym.name.lower()] = sym.as_type
    return _TYPE_ENV_MODULE_BASE_CACHE.put(base, symbols)  # type: ignore[no-any-return]


def type_environment_for(symbols: ModuleSymbols, proc: ProcedureNode) -> dict[str, str]:
    """Per-procedure {lowercased name -> raw declared as-type} type environment.

    Module-level typed non-procedure symbols first, then the procedure's own
    return binding, then params/locals last (so a local shadowing a module name
    wins). Values are the raw as-type string (callers normalize at comparison).
    """
    cached = _TYPE_ENV_CACHE.get(symbols, proc)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    out = dict(_type_env_module_base(symbols))
    proc_sym = procedure_symbol_for(symbols, proc)
    return_type = _return_assignment_type_for(proc)
    if return_type:
        out[proc.name.lower()] = return_type
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if child.as_type:
            out[child.name.lower()] = child.as_type
    return _TYPE_ENV_CACHE.put(out, symbols, proc)  # type: ignore[no-any-return]


# Procedure symbols are looked up per procedure per rule; the by-start-offset
# index turns each lookup from an O(module members) scan into a dict hit.
_PROCEDURE_SYMBOL_INDEX_CACHE = IdentityLru()


def _procedure_symbol_index(symbols: ModuleSymbols) -> dict[int, VbaSymbol]:
    cached = _PROCEDURE_SYMBOL_INDEX_CACHE.get(symbols)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    index: dict[int, VbaSymbol] = {}
    for sym in symbols.root.children or []:
        if sym.kind in _PROCEDURE_KINDS and sym.full_span.start not in index:
            index[sym.full_span.start] = sym
    return _PROCEDURE_SYMBOL_INDEX_CACHE.put(index, symbols)  # type: ignore[no-any-return]


def procedure_symbol_for(symbols: ModuleSymbols, proc: ProcedureNode) -> VbaSymbol | None:
    """The module symbol for a procedure node, matched by declaration start offset."""
    return _procedure_symbol_index(symbols).get(proc.span.start)


def declared_shape_for_source_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> SourceDeclaredShape:
    """Resolve a bare identifier to its declared shape via the source symbol graph."""
    binding = resolve_bare_identifier_binding(
        BareIdentifierResolutionInput(
            current_module=symbols,
            name=name,
            context=context,
            enclosing_procedure=proc_sym,
            project_visible_symbols=list(project_visible_symbols)
            if project_visible_symbols
            else [],
        )
    )
    if binding.scope in (
        BareIdentifierResolutionScope.UNRESOLVED,
        BareIdentifierResolutionScope.AMBIGUOUS,
    ):
        return SourceDeclaredShape(resolved=binding.scope is BareIdentifierResolutionScope.AMBIGUOUS)
    shaped = next((d for d in binding.definitions if d.as_type or d.is_array), None)
    if shaped is None:
        return SourceDeclaredShape(resolved=True, shape=DeclaredValueShape(None, False, False))
    return SourceDeclaredShape(
        resolved=True,
        shape=DeclaredValueShape(
            as_type=shaped.as_type,
            is_array=shaped.is_array is True,
            is_fixed_array=shaped.array_bounds is not None,
        ),
    )


def is_value_declaration_symbol(sym: VbaSymbol) -> bool:
    return sym.kind in _VALUE_DECLARATION_KINDS


def _source_identifier_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> BareIdentifierResolution:
    return resolve_bare_identifier_binding(
        BareIdentifierResolutionInput(
            current_module=symbols,
            name=name,
            context=context,
            enclosing_procedure=proc_sym,
            project_visible_symbols=list(project_visible_symbols) if project_visible_symbols else [],
        )
    )


def source_identifier_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> BareIdentifierResolution:
    """Resolve a bare identifier to its binding (the public binder seam for rules)."""
    return _source_identifier_binding(symbols, proc_sym, project_visible_symbols, name, context)


def source_identifier_bound(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> bool:
    """Whether a bare identifier resolves to any source binding (scope != UNRESOLVED)."""
    binding = source_identifier_binding(symbols, proc_sym, project_visible_symbols, name, context)
    return binding.scope is not BareIdentifierResolutionScope.UNRESOLVED


def declared_type_for_source_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> SourceDeclaredType:
    binding = _source_identifier_binding(symbols, proc_sym, project_visible_symbols, name, context)
    if binding.scope in (
        BareIdentifierResolutionScope.UNRESOLVED,
        BareIdentifierResolutionScope.AMBIGUOUS,
    ):
        return SourceDeclaredType(resolved=binding.scope is BareIdentifierResolutionScope.AMBIGUOUS)
    typed = next((d for d in binding.definitions if d.as_type), None)
    return SourceDeclaredType(resolved=True, as_type=typed.as_type if typed is not None else None)


def declared_value_type_for_source_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
) -> SourceDeclaredType:
    binding = _source_identifier_binding(
        symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.EXPRESSION
    )
    if binding.scope in (
        BareIdentifierResolutionScope.UNRESOLVED,
        BareIdentifierResolutionScope.AMBIGUOUS,
    ):
        return SourceDeclaredType(resolved=binding.scope is BareIdentifierResolutionScope.AMBIGUOUS)
    value_definitions = [d for d in binding.definitions if is_value_declaration_symbol(d)]
    if not value_definitions:
        return SourceDeclaredType(resolved=False)
    typed = next((d for d in value_definitions if d.as_type), None)
    return SourceDeclaredType(resolved=True, as_type=typed.as_type if typed is not None else None)


def declared_value_type_for_qualified_source_binding(
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    qualifier: str,
    name: str,
) -> SourceDeclaredType:
    qualifier_lower = qualifier.lower()
    name_lower = name.lower()
    candidates: list[VbaSymbol] = []
    if symbols.module_name.lower() == qualifier_lower:
        candidates.extend(symbols.root.children or [])
    candidates.extend(
        s for s in (project_visible_symbols or []) if s.module_name.lower() == qualifier_lower
    )
    if not candidates:
        return SourceDeclaredType(resolved=False)
    matching_values = [
        s for s in candidates if s.name.lower() == name_lower and is_value_declaration_symbol(s)
    ]
    if not matching_values:
        return SourceDeclaredType(resolved=True)
    typed = next((d for d in matching_values if d.as_type), None)
    return SourceDeclaredType(resolved=True, as_type=typed.as_type if typed is not None else None)
