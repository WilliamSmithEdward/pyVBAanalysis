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

from ..parser.nodes import ProcedureNode
from ..symbols.name_resolution import (
    BareIdentifierContext,
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


_VALUE_DECLARATION_KINDS = frozenset(
    {
        VbaSymbolKind.PARAMETER,
        VbaSymbolKind.LOCAL_VARIABLE,
        VbaSymbolKind.MODULE_VARIABLE,
        VbaSymbolKind.CONSTANT,
    }
)


def declaration_shape_environment_for(
    symbols: ModuleSymbols, proc: ProcedureNode
) -> dict[str, DeclaredValueShape]:
    """Syntactic declared-shape fallback for a procedure's module + local value names.

    The function-return-variable shape (an Erase/assignment target named after the
    procedure) is omitted; that is a precision-only gap, never a false positive.
    """
    out: dict[str, DeclaredValueShape] = {}
    for sym in symbols.root.children or []:
        if sym.kind in _VALUE_DECLARATION_KINDS:
            out[sym.name.lower()] = _shape_of(sym)
    proc_sym = procedure_symbol_for(symbols, proc)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if child.kind in _VALUE_DECLARATION_KINDS:
            out[child.name.lower()] = _shape_of(child)
    return out


def _shape_of(sym: VbaSymbol) -> DeclaredValueShape:
    return DeclaredValueShape(
        as_type=sym.as_type,
        is_array=sym.is_array is True,
        is_fixed_array=sym.array_bounds is not None,
    )


def procedure_symbol_for(symbols: ModuleSymbols, proc: ProcedureNode) -> VbaSymbol | None:
    """The module symbol for a procedure node, matched by declaration start offset."""
    for sym in symbols.root.children or []:
        if sym.kind in _PROCEDURE_KINDS and sym.full_span.start == proc.span.start:
            return sym
    return None


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
