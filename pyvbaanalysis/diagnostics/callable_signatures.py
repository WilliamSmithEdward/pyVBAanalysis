"""Callable-signature tables and source-name scope for diagnostics.

Ported from the call-resolution slice of
xlide_vscode/src/analyzer/diagnostics/typeInference.ts. Builds the module +
project callable signature tables the call/argument rules resolve against, the
source-name shadow scope that suppresses an intrinsic diagnostic when a user
declares the same name, and the scoped integer-constant lookup. Host/runtime
function signatures and external constants are deferred (M9): they resolve to
None, which is precision-only (never a false positive).
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass

from ..constants.integer_constant_expression import IntegerConstantLookup
from ..lexer.token_helpers import match_paren_from
from ..lexer.token_kinds import VbaToken
from ..parser.nodes import ProcedureNode, Span
from ..symbols.name_resolution import (
    BareIdentifierContext,
    BareIdentifierResolutionInput,
    BareIdentifierResolutionScope,
    resolve_bare_identifier_binding,
    source_identifier_names,
)
from ..symbols.symbol_model import (
    ModuleSymbols,
    VbaProcedureSignature,
    VbaSymbol,
    VbaSymbolKind,
    is_bare_callable_kind,
    is_procedure_kind,
    procedure_params_from_symbol,
    qualified_procedure_key,
)
from ..types.type_inference import procedure_symbol_for
from .call_extraction import (
    CallableParamType,
    CallableTypeSignature,
    CallArguments,
    empty_arg_split,
    split_arg_slots,
)
from .context import statement_tokens
from .walker import strip_header_brackets, token_name


# -- signature tables ------------------------------------------------------


def is_by_ref_procedure_param(by_ref: bool | None, by_val: bool | None, param_array: bool) -> bool:
    if param_array:
        return False
    return by_ref is True or by_val is not True


def callable_type_signature_from_symbol(symbol: VbaSymbol) -> CallableTypeSignature:
    params = [
        CallableParamType(
            name=strip_header_brackets(p.name),
            type_=p.type_,
            optional=p.optional,
            param_array=p.param_array,
            is_array=p.is_array,
            by_ref=is_by_ref_procedure_param(p.by_ref, p.by_val, p.param_array),
        )
        for p in procedure_params_from_symbol(symbol, include_passing=True)
    ]
    return CallableTypeSignature(name=symbol.name, params=params, return_type=symbol.as_type)


def build_module_type_signatures(symbols: ModuleSymbols) -> dict[str, CallableTypeSignature]:
    out: dict[str, CallableTypeSignature] = {}
    for symbol in symbols.root.children or []:
        if is_procedure_kind(symbol.kind) or symbol.kind is VbaSymbolKind.DECLARE:
            out[symbol.name.lower()] = callable_type_signature_from_symbol(symbol)
    return out


def same_module_callable_signatures(symbols: ModuleSymbols) -> dict[str, list[CallableTypeSignature]]:
    out: dict[str, list[CallableTypeSignature]] = {}
    for symbol in symbols.root.children or []:
        if not is_bare_callable_kind(symbol.kind):
            continue
        sig = callable_type_signature_from_symbol(symbol)
        out.setdefault(sig.name.lower(), []).append(sig)
    return out


def unique_project_type_signatures(
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
) -> dict[str, CallableTypeSignature]:
    out: dict[str, CallableTypeSignature] = {}
    if not project_procedures:
        return out
    for lower, candidates in project_procedures.items():
        if len(candidates) != 1:
            continue
        candidate = candidates[0]
        params = [
            CallableParamType(
                name=p.name,
                type_=p.type_,
                optional=p.optional,
                param_array=p.param_array,
                is_array=p.is_array,
                by_ref=is_by_ref_procedure_param(p.by_ref, p.by_val, p.param_array),
            )
            for p in candidate.params
        ]
        out[lower] = CallableTypeSignature(
            name=candidate.name, params=params, return_type=candidate.return_type
        )
    return out


def callable_type_signatures_for(
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
) -> dict[str, CallableTypeSignature]:
    out = dict(build_module_type_signatures(symbols))
    for lower, sig in unique_project_type_signatures(project_procedures).items():
        out.setdefault(lower, sig)
    return out


# -- source-name shadow scope ----------------------------------------------


@dataclass(frozen=True, slots=True)
class SourceNameScope:
    callable_shadows: frozenset[str]
    runtime_shadows: frozenset[str]


def is_non_callable_symbol(sym: VbaSymbol) -> bool:
    return sym.kind in (
        VbaSymbolKind.PARAMETER,
        VbaSymbolKind.LOCAL_VARIABLE,
        VbaSymbolKind.MODULE_VARIABLE,
        VbaSymbolKind.CONSTANT,
        VbaSymbolKind.ENUM,
        VbaSymbolKind.ENUM_MEMBER,
        VbaSymbolKind.TYPE,
    )


def module_non_callable_symbols(symbols: ModuleSymbols) -> dict[str, VbaSymbol]:
    out: dict[str, VbaSymbol] = {}
    callable_names = {
        sym.name.lower()
        for sym in (symbols.root.children or [])
        if is_procedure_kind(sym.kind) or sym.kind is VbaSymbolKind.DECLARE
    }
    for sym in symbols.root.children or []:
        if is_non_callable_symbol(sym) and sym.name.lower() not in callable_names:
            out[sym.name.lower()] = sym
        if sym.kind is VbaSymbolKind.ENUM:
            for child in sym.children or []:
                if child.name.lower() not in callable_names:
                    out[child.name.lower()] = child
    return out


def source_name_scope_for(
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    project_visible_symbols: Sequence[VbaSymbol] | None = None,
) -> SourceNameScope:
    callable_shadows = set(module_non_callable_symbols(symbols))
    proc_sym = procedure_symbol_for(symbols, proc)
    runtime_shadows = source_identifier_names(
        symbols, proc_sym, project_visible_symbols if project_visible_symbols is not None else ()
    )
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if is_non_callable_symbol(child):
            callable_shadows.add(child.name.lower())
    return SourceNameScope(
        callable_shadows=frozenset(callable_shadows), runtime_shadows=frozenset(runtime_shadows)
    )


def runtime_callable_source_shadowed(name: str, source_names: SourceNameScope | None) -> bool:
    return source_names is not None and name.lower() in source_names.runtime_shadows


def bare_callable_source_shadowed(name: str, source_names: SourceNameScope | None) -> bool:
    return source_names is not None and name.lower() in source_names.callable_shadows


# -- call resolution -------------------------------------------------------


def callable_signature_for(
    name: str,
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None = None,
) -> CallableTypeSignature | None:
    """The signature a bare callee resolves to, or None.

    Runtime-function signatures are deferred (M9): a bare runtime call resolves to
    None here, which is precision-only (it is never arity/type-checked, so never a
    false positive).
    """
    if bare_callable_source_shadowed(name, source_names):
        return None
    return module_signatures.get(name.lower())


def callable_signature_for_call(
    call: CallArguments,
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None = None,
) -> CallableTypeSignature | None:
    if call.lookup_key:
        return module_signatures.get(call.lookup_key)
    return callable_signature_for(call.name, module_signatures, source_names)


@dataclass(frozen=True, slots=True)
class ParenthesizedCallName:
    name: str
    paren_index: int
    name_end_index: int


def parenthesized_call_name_at(
    toks: Sequence[VbaToken], name_index: int
) -> ParenthesizedCallName | None:
    base_name = token_name(toks[name_index])
    if not base_name:
        return None
    suffix = toks[name_index + 1] if name_index + 1 < len(toks) else None
    after_suffix = toks[name_index + 2] if name_index + 2 < len(toks) else None
    if (
        suffix is not None
        and suffix.raw_text == "$"
        and toks[name_index].end == suffix.start
        and after_suffix is not None
        and after_suffix.raw_text == "("
        and suffix.end == after_suffix.start
    ):
        return ParenthesizedCallName(f"{base_name}$", name_index + 2, name_index + 1)
    if suffix is not None and suffix.raw_text == "(":
        return ParenthesizedCallName(base_name, name_index + 1, name_index)
    return None


def expression_calls(
    source: str,
    span: Span,
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None = None,
) -> list[CallArguments]:
    """Parenthesized current-module / unique-project calls inside an expression."""
    toks = statement_tokens(source, span)
    out: list[CallArguments] = []
    for i in range(len(toks) - 1):
        call_name = parenthesized_call_name_at(toks, i)
        if call_name is None:
            continue
        qualifier = (
            token_name(toks[i - 2]) if i >= 2 and toks[i - 1].raw_text == "." else None
        )
        lookup_key = qualified_procedure_key(qualifier, call_name.name) if qualifier else None
        if qualifier and (lookup_key is None or lookup_key not in module_signatures):
            continue  # host/member calls need receiver binding before checking
        if not qualifier and i > 0 and toks[i - 1].raw_text == ".":
            continue
        if lookup_key is not None:
            if lookup_key not in module_signatures:
                continue
        elif callable_signature_for(call_name.name, module_signatures, source_names) is None:
            continue
        close = match_paren_from(toks, call_name.paren_index)
        if close < 0:
            continue
        inner = list(toks[call_name.paren_index + 1 : close])
        split = empty_arg_split() if not inner else split_arg_slots(inner, span.start)
        out.append(
            CallArguments(
                name=call_name.name,
                qualifier=qualifier,
                lookup_key=lookup_key,
                name_span=Span(
                    span.start + toks[i].start, span.start + toks[call_name.name_end_index].end
                ),
                slots=split.slots,
                slot_spans=split.spans,
                slice_start=span.start,
            )
        )
    return out


# -- scoped integer-constant lookup ----------------------------------------


def is_integer_constant_binding_symbol(symbol: VbaSymbol) -> bool:
    return symbol.kind in (VbaSymbolKind.CONSTANT, VbaSymbolKind.ENUM_MEMBER)


class _ScopedIntegerConstantLookup:
    __slots__ = ("_constants", "_symbols", "_proc_sym", "_project_visible")

    def __init__(
        self,
        constants: Mapping[str, int | None],
        symbols: ModuleSymbols,
        proc_sym: VbaSymbol | None,
        project_visible: Sequence[VbaSymbol] | None,
    ) -> None:
        self._constants = constants
        self._symbols = symbols
        self._proc_sym = proc_sym
        self._project_visible = project_visible

    def get(self, name: str, /) -> int | None:
        key = name.lower()
        if "." in key:
            # External (runtime/host) qualified constants are deferred to M9.
            return self._constants.get(key) if key in self._constants else None
        binding = resolve_bare_identifier_binding(
            BareIdentifierResolutionInput(
                current_module=self._symbols,
                name=name,
                context=BareIdentifierContext.EXPRESSION,
                enclosing_procedure=self._proc_sym,
                project_visible_symbols=list(self._project_visible) if self._project_visible else [],
            )
        )
        if binding.scope is BareIdentifierResolutionScope.UNRESOLVED:
            return self._constants.get(key) if key in self._constants else None
        if binding.scope is BareIdentifierResolutionScope.AMBIGUOUS or any(
            not is_integer_constant_binding_symbol(d) for d in binding.definitions
        ):
            return None
        return self._constants.get(key)


def scoped_integer_constant_lookup(
    constants: Mapping[str, int | None],
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible: Sequence[VbaSymbol] | None,
) -> IntegerConstantLookup:
    return _ScopedIntegerConstantLookup(constants, symbols, proc_sym, project_visible)
