"""VBA symbol graph: per-module symbols, name resolution, and the project index."""

from .build_module_symbols import BuildModuleSymbolsOptions, build_module_symbols
from .name_resolution import (
    BareIdentifierContext,
    BareIdentifierResolution,
    BareIdentifierResolutionInput,
    BareIdentifierResolutionScope,
    resolve_bare_identifier_binding,
    source_identifier_names,
)
from .symbol_model import (
    ModuleSymbolKind,
    ModuleSymbols,
    SymbolVisibility,
    VbaProcedureParam,
    VbaProcedureSignature,
    VbaSymbol,
    VbaSymbolAttribute,
    VbaSymbolKind,
    is_bare_callable_kind,
    is_procedure_kind,
    procedure_signature_from_symbol,
    qualified_procedure_key,
)

__all__ = [
    "BuildModuleSymbolsOptions",
    "build_module_symbols",
    "BareIdentifierContext",
    "BareIdentifierResolution",
    "BareIdentifierResolutionInput",
    "BareIdentifierResolutionScope",
    "resolve_bare_identifier_binding",
    "source_identifier_names",
    "ModuleSymbolKind",
    "ModuleSymbols",
    "SymbolVisibility",
    "VbaProcedureParam",
    "VbaProcedureSignature",
    "VbaSymbol",
    "VbaSymbolAttribute",
    "VbaSymbolKind",
    "is_bare_callable_kind",
    "is_procedure_kind",
    "procedure_signature_from_symbol",
    "qualified_procedure_key",
]
