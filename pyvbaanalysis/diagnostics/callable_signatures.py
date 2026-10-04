"""Callable-signature tables and source-name scope for diagnostics.

Ported from the call-resolution slice of
xlide_vscode/src/analyzer/diagnostics/typeInference.ts. Builds the module +
project callable signature tables the call/argument rules resolve against, the
source-name shadow scope that suppresses an intrinsic diagnostic when a user
declares the same name, the expression-level call extraction (expression_calls
and the member-call binders), and the scoped integer-constant lookup, which
falls back to VBA runtime and host constants for a name no source declares.
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, replace

from ..call.call_context import standalone_empty_parenthesized_call_statement
from ..completion.member_access import (
    MemberCompletionContext,
    resolve_exact_member_completion,
)
from ..conditional import ConditionalActivityTracker
from ..constants.integer_constant_expression import IntegerConstantLookup
from ..host.host_model import HostObjectModel
from ..identity_cache import IdentityLru
from ..lexer.token_helpers import match_paren_from
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import ProcedureNode, Span
from ..runtime.vba_runtime import VbaRuntimeFunction, resolve_runtime_function
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
from ..types.type_inference import (
    def_type_of,
    is_member_parenless_argument_start,
    is_property_result_indexing,
    parse_runtime_param_type,
    procedure_symbol_for,
    runtime_signature_parameter_text,
    split_signature_top_level,
)
from .call_extraction import (
    CallableParamType,
    CallableTypeSignature,
    CallArguments,
    empty_arg_split,
    split_arg_slots,
)
from .const_expr import collect_body_literal_integer_constants, external_integer_constant_value
from .context import statement_tokens
from .walker import (
    first_executable_token_index,
    statement_tokens_after_leading_label,
    strip_header_brackets,
    token_name,
    token_text,
    top_level_operator_index,
)


# -- signature tables ------------------------------------------------------


def is_by_ref_procedure_param(by_ref: bool | None, by_val: bool | None, param_array: bool) -> bool:
    if param_array:
        return False
    return by_ref is True or by_val is not True


_VALUED_KINDS = frozenset({VbaSymbolKind.FUNCTION, VbaSymbolKind.PROPERTY_GET})


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
    return CallableTypeSignature(
        name=symbol.name,
        params=params,
        return_type=symbol.as_type,
        valued=True if symbol.kind in _VALUED_KINDS else None,
    )


def _module_signature_from_symbol(symbols: ModuleSymbols, symbol: VbaSymbol) -> CallableTypeSignature:
    """A procedure's signature, its untyped parameters and result typed by the
    module's DefType lines (XLIDE issue #285)."""
    declared = callable_type_signature_from_symbol(symbol)
    if not symbols.def_types or symbol.kind is VbaSymbolKind.DECLARE:
        return declared
    return replace(
        declared,
        params=[
            p if p.type_ or p.param_array else replace(p, type_=def_type_of(symbols, p.name))
            for p in declared.params
        ],
        return_type=declared.return_type
        if declared.return_type is not None
        else (def_type_of(symbols, symbol.name) if symbol.kind in _VALUED_KINDS else None),
    )


# Signature tables are pure functions of the per-pass ModuleSymbols (plus the
# per-pass project procedures), so they are memoized by identity. Results are
# shared: callers must not mutate them.
_MODULE_TYPE_SIGNATURES_CACHE = IdentityLru()
_SAME_MODULE_CALLABLE_SIGNATURES_CACHE = IdentityLru()
_CALLABLE_TYPE_SIGNATURES_CACHE = IdentityLru()
_UNIQUE_PROJECT_TYPE_SIGNATURES_CACHE = IdentityLru()
_EMPTY_PROJECT_TYPE_SIGNATURES: dict[str, CallableTypeSignature] = {}


def build_module_type_signatures(symbols: ModuleSymbols) -> dict[str, CallableTypeSignature]:
    cached = _MODULE_TYPE_SIGNATURES_CACHE.get(symbols)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    out: dict[str, CallableTypeSignature] = {}
    for symbol in symbols.root.children or []:
        if is_procedure_kind(symbol.kind) or symbol.kind is VbaSymbolKind.DECLARE:
            out[symbol.name.lower()] = _module_signature_from_symbol(symbols, symbol)
    return _MODULE_TYPE_SIGNATURES_CACHE.put(out, symbols)  # type: ignore[no-any-return]


def same_module_callable_signatures(symbols: ModuleSymbols) -> dict[str, list[CallableTypeSignature]]:
    cached = _SAME_MODULE_CALLABLE_SIGNATURES_CACHE.get(symbols)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    out: dict[str, list[CallableTypeSignature]] = {}
    for symbol in symbols.root.children or []:
        if not is_bare_callable_kind(symbol.kind):
            continue
        sig = _module_signature_from_symbol(symbols, symbol)
        out.setdefault(sig.name.lower(), []).append(sig)
    return _SAME_MODULE_CALLABLE_SIGNATURES_CACHE.put(out, symbols)  # type: ignore[no-any-return]


def unique_project_type_signatures(
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
) -> dict[str, CallableTypeSignature]:
    if project_procedures is None:
        return _EMPTY_PROJECT_TYPE_SIGNATURES
    cached = _UNIQUE_PROJECT_TYPE_SIGNATURES_CACHE.get(project_procedures)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    out: dict[str, CallableTypeSignature] = {}
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
            name=candidate.name,
            params=params,
            return_type=candidate.return_type,
            valued=True if candidate.kind is VbaSymbolKind.FUNCTION else None,
        )
    return _UNIQUE_PROJECT_TYPE_SIGNATURES_CACHE.put(out, project_procedures)  # type: ignore[no-any-return]


def callable_type_signatures_for(
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
) -> dict[str, CallableTypeSignature]:
    cached = _CALLABLE_TYPE_SIGNATURES_CACHE.get(symbols, project_procedures)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    # Copy: build_module_type_signatures' result is memoized and must stay pure.
    out = dict(build_module_type_signatures(symbols))
    for lower, sig in unique_project_type_signatures(project_procedures).items():
        out.setdefault(lower, sig)
    return _CALLABLE_TYPE_SIGNATURES_CACHE.put(out, symbols, project_procedures)  # type: ignore[no-any-return]


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


# The module-level portion of the scopes below is procedure-independent, and
# every statement/expression rule asks for the scope of the same procedures, so
# both layers are memoized by identity: rebuilding them per rule x procedure
# was the dominant cost of a full pass on large modules (the scans below are
# O(module declarations) each).
_MODULE_NON_CALLABLE_CACHE = IdentityLru()
_SOURCE_NAME_SCOPE_CACHE = IdentityLru(capacity=64)
_NO_PROJECT_SYMBOLS: tuple[VbaSymbol, ...] = ()


def module_non_callable_symbols(symbols: ModuleSymbols) -> dict[str, VbaSymbol]:
    cached = _MODULE_NON_CALLABLE_CACHE.get(symbols)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
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
    return _MODULE_NON_CALLABLE_CACHE.put(out, symbols)  # type: ignore[no-any-return]


def source_name_scope_for(
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    project_visible_symbols: Sequence[VbaSymbol] | None = None,
) -> SourceNameScope:
    project = project_visible_symbols if project_visible_symbols is not None else _NO_PROJECT_SYMBOLS
    cached = _SOURCE_NAME_SCOPE_CACHE.get(symbols, proc, project)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    callable_shadows = set(module_non_callable_symbols(symbols))
    proc_sym = procedure_symbol_for(symbols, proc)
    runtime_shadows = source_identifier_names(symbols, proc_sym, project)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if is_non_callable_symbol(child):
            callable_shadows.add(child.name.lower())
    scope = SourceNameScope(
        callable_shadows=frozenset(callable_shadows), runtime_shadows=frozenset(runtime_shadows)
    )
    return _SOURCE_NAME_SCOPE_CACHE.put(scope, symbols, proc, project)  # type: ignore[no-any-return]


def runtime_callable_source_shadowed(name: str, source_names: SourceNameScope | None) -> bool:
    # A Public Function InStr compiles, and `InStr(0, "abc", "a")` still calls
    # VBA's, which raises 5 (XLIDE issue #280, measured in Excel 16.0).
    lower = name.lower()
    return lower != "instr" and source_names is not None and lower in source_names.runtime_shadows


def bare_callable_source_shadowed(name: str, source_names: SourceNameScope | None) -> bool:
    return source_names is not None and name.lower() in source_names.callable_shadows


# -- call resolution -------------------------------------------------------


def callable_signature_for(
    name: str,
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None = None,
) -> CallableTypeSignature | None:
    """The signature a bare callee resolves to (user module, then VBA runtime), or None."""
    if bare_callable_source_shadowed(name, source_names):
        return None
    user = module_signatures.get(name.lower())
    if user is not None:
        return user
    if runtime_callable_source_shadowed(name, source_names):
        return None
    # The runtime table leaves out the type names; called, `String(n, c)` is
    # String$'s Variant twin and takes its arguments (XLIDE issue #332).
    runtime = resolve_runtime_function(name)
    if runtime is None and name.lower() == "string":
        runtime = resolve_runtime_function("String$")
    if runtime is None:
        return None
    return runtime_type_signature(runtime)


# -- runtime-function signatures -------------------------------------------


def runtime_type_signature(runtime: VbaRuntimeFunction) -> CallableTypeSignature:
    if runtime.params is not None:
        params = [
            CallableParamType(
                name=p.name,
                type_=p.type_,
                optional=p.optional,
                param_array=p.param_array,
                is_array=True if p.is_array else None,
                null_raises=True if p.null_raises else None,
            )
            for p in runtime.params
        ]
        return CallableTypeSignature(name=runtime.name, params=params, return_type=runtime.returns)
    return parse_runtime_display_signature(runtime.name, runtime.signature, runtime.returns)


def runtime_arity_signature(runtime: VbaRuntimeFunction) -> CallableTypeSignature | None:
    if runtime.params is not None or runtime_signature_parameter_text(runtime.signature) is not None:
        return runtime_type_signature(runtime)
    return None


def parse_runtime_display_signature(
    name: str, signature: str, return_type: str | None = None
) -> CallableTypeSignature:
    inner = runtime_signature_parameter_text(signature)
    if inner is None:
        return CallableTypeSignature(name=name, params=[], return_type=return_type)
    params = [
        p for p in (parse_runtime_param_type(s) for s in split_signature_top_level(inner)) if p is not None
    ]
    return CallableTypeSignature(name=name, params=params, return_type=return_type)


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
    first_executable = first_executable_token_index(toks)
    statement_head = token_text(toks[first_executable]) if first_executable < len(toks) else ""
    for i in range(len(toks) - 1):
        call_name = parenthesized_call_name_at(toks, i)
        if call_name is None:
            continue
        # `Take (i)` as a statement passes `(i)` as its one argument: the space
        # makes the parentheses part of the argument, which is then a copy (XLIDE
        # issue #111). `Take(i)` glued, and `x = Take (i)` inside an expression,
        # are calls with a list.
        arguments_parenthesized = (
            i == first_executable and toks[call_name.paren_index].start > toks[call_name.name_end_index].end
        )
        # `ReDim Three(2)` inside Function Three sizes the return array, and a
        # ReDim target anywhere is a variable, never a call (XLIDE issue #115).
        if statement_head == "redim" and _is_redim_target_at(toks, i, first_executable):
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
                arguments_parenthesized=arguments_parenthesized,
            )
        )
    return out


def _is_redim_target_at(toks: Sequence[VbaToken], index: int, first_executable: int) -> bool:
    """Whether the name at `index` is a target of the ReDim statement the tokens
    spell: at depth 0, right after `ReDim`, `Preserve`, or a separating comma."""
    depth = 0
    for k in range(first_executable + 1, index):
        raw = toks[k].raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
    if depth != 0:
        return False
    prev = token_text(toks[index - 1]) if index >= 1 else ""
    return prev == "redim" or prev == "preserve" or prev == ","


# -- member calls ----------------------------------------------------------


@dataclass(frozen=True, slots=True)
class BoundMemberCall:
    """A member call whose receiver the member-completion context bound to a member
    with a verified signature."""

    call: CallArguments
    signature: CallableTypeSignature


_WHITESPACE_RE = re.compile(r"\s")


def member_expression_calls(
    source: str, span: Span, member_ctx: MemberCompletionContext
) -> list[BoundMemberCall]:
    """Parenthesized member calls anywhere in a statement: `ws.Range("A1")`,
    `Application.Calculate()`, `p.Save(1)`."""
    toks = statement_tokens(source, span)
    standalone_empty_call = standalone_empty_parenthesized_call_statement(source, span)
    out: list[BoundMemberCall] = []
    for i in range(1, len(toks) - 1):
        name = token_name(toks[i])
        if not name or toks[i - 1].raw_text != "." or toks[i + 1].raw_text != "(":
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0:
            continue
        member = resolve_exact_member_completion(source, name, span.start + toks[i].end, member_ctx)
        if member is None or not member.signature:
            continue
        inner = list(toks[i + 2 : close])
        call_span = Span(span.start + toks[i].start, span.start + toks[close].end)
        # `obj.Foo()` as a whole statement belongs to the empty-parentheses rule.
        if (
            standalone_empty_call is not None
            and standalone_empty_call.is_member
            and standalone_empty_call.span.start == call_span.start
            and standalone_empty_call.span.end == call_span.end
        ):
            continue
        parsed = parse_runtime_display_signature(member.name, member.signature)
        if is_property_result_indexing(member, parsed, inner):
            continue
        # A Function of a project class gives a value, and `k.Items(1)` may index
        # it (XLIDE issue #609).
        signature = (
            replace(
                parsed,
                valued=True,
                return_type=member.returns if member.returns is not None else parsed.return_type,
            )
            if member.kind == "method" and not member.sub and len(member.definitions or ()) > 0
            else parsed
        )
        split = empty_arg_split() if not inner else split_arg_slots(inner, span.start)
        out.append(
            BoundMemberCall(
                call=CallArguments(
                    name=member.name,
                    name_span=Span(call_span.start, span.start + toks[i].end),
                    slots=split.slots,
                    slot_spans=split.spans,
                    slice_start=span.start,
                ),
                signature=signature,
            )
        )
    return out


def member_statement_calls(
    source: str, span: Span, member_ctx: MemberCompletionContext
) -> list[BoundMemberCall]:
    """The parenless member call a statement makes: `p.Save "x"`, `.Save`,
    `Call Err.Raise`. Parenthesized calls belong to member_expression_calls."""
    toks = statement_tokens_after_leading_label(source, span)
    if not toks or top_level_operator_index(toks, "=") >= 0:
        return []
    explicit_call = token_text(toks[0]) == "call"
    chain_start = 1 if explicit_call else 0
    if chain_start >= len(toks):
        return []
    start_tok = toks[chain_start]
    if not token_name(start_tok) and start_tok.raw_text != ".":
        return []
    first_member_index = chain_start + 1 if start_tok.raw_text == "." else chain_start + 2
    for i in range(first_member_index, len(toks)):
        name = token_name(toks[i])
        if not name or toks[i - 1].raw_text != ".":
            continue
        if not is_member_statement_chain_through(toks, chain_start, i):
            continue
        next_tok = toks[i + 1] if i + 1 < len(toks) else None
        if next_tok is not None and next_tok.raw_text == "(":
            continue
        if explicit_call and next_tok is not None:
            continue  # `Call p.Save arg` is the call-requires-parens syntax error
        if next_tok is not None:
            gap = source[span.start + toks[i].end : span.start + next_tok.start]
            if _WHITESPACE_RE.search(gap) is None or not is_member_parenless_argument_start(next_tok):
                continue
        member = resolve_exact_member_completion(source, name, span.start + toks[i].end, member_ctx)
        if member is None or not member.signature:
            continue
        arg_toks = list(toks[i + 1 :])
        split = empty_arg_split() if not arg_toks else split_arg_slots(arg_toks, span.start)
        return [
            BoundMemberCall(
                call=CallArguments(
                    name=member.name,
                    name_span=Span(span.start + toks[i].start, span.start + toks[i].end),
                    explicit_call=explicit_call,
                    slots=split.slots,
                    slot_spans=split.spans,
                    slice_start=span.start,
                ),
                signature=parse_runtime_display_signature(member.name, member.signature),
            )
        ]
    return []


def is_member_statement_chain_through(
    toks: Sequence[VbaToken], start_idx: int, member_idx: int
) -> bool:
    """True when the tokens from `start_idx` form one receiver chain that reaches
    the member at `member_idx`: names joined by dots, optionally called."""
    if toks[start_idx].raw_text == ".":
        if start_idx + 1 >= len(toks) or not token_name(toks[start_idx + 1]):
            return False
        if start_idx + 1 == member_idx:
            return True
        return is_member_statement_chain_through(toks, start_idx + 1, member_idx)
    if not token_name(toks[start_idx]):
        return False
    # A keyword spaced off a dot is a statement's keyword before a With member,
    # not a receiver: `If .Count = 0 Then .Add 1` (XLIDE issue #198).
    next_tok = toks[start_idx + 1] if start_idx + 1 < len(toks) else None
    if (
        toks[start_idx].kind is TokenKind.KEYWORD
        and next_tok is not None
        and next_tok.raw_text == "."
        and next_tok.start > toks[start_idx].end
    ):
        return False
    i = start_idx + 1
    while i < len(toks):
        raw = toks[i].raw_text
        if raw == "(":
            close = match_paren_from(toks, i)
            if close < 0 or close >= member_idx:
                return False
            i = close + 1
            continue
        if raw != ".":
            return False
        name_idx = i + 1
        if name_idx >= len(toks) or not token_name(toks[name_idx]):
            return False
        if name_idx == member_idx:
            return True
        i = name_idx + 1
    return False


# -- scoped integer-constant lookup ----------------------------------------


def is_integer_constant_binding_symbol(symbol: VbaSymbol) -> bool:
    return symbol.kind in (VbaSymbolKind.CONSTANT, VbaSymbolKind.ENUM_MEMBER)


class _ScopedIntegerConstantLookup:
    __slots__ = ("_constants", "_symbols", "_proc_sym", "_project_visible", "_model")

    def __init__(
        self,
        constants: Mapping[str, float | None],
        symbols: ModuleSymbols,
        proc_sym: VbaSymbol | None,
        project_visible: Sequence[VbaSymbol] | None,
        model: HostObjectModel | None,
    ) -> None:
        self._constants = constants
        self._symbols = symbols
        self._proc_sym = proc_sym
        self._project_visible = project_visible
        self._model = model

    def get(self, name: str, /) -> float | None:
        key = name.lower()
        if "." in key:
            if key in self._constants:
                return self._constants[key]
            return external_integer_constant_value(key, self._model)
        binding = resolve_bare_identifier_binding(
            BareIdentifierResolutionInput(
                current_module=self._symbols,
                name=name,
                context=BareIdentifierContext.EXPRESSION,
                enclosing_procedure=self._proc_sym,
                project_visible_symbols=self._project_visible or (),
            )
        )
        if binding.scope is BareIdentifierResolutionScope.UNRESOLVED:
            if key in self._constants:
                return self._constants[key]
            return external_integer_constant_value(key, self._model)
        if binding.scope is BareIdentifierResolutionScope.AMBIGUOUS or any(
            not is_integer_constant_binding_symbol(d) for d in binding.definitions
        ):
            return None
        return self._constants.get(key)


def scoped_integer_constant_lookup(
    constants: Mapping[str, float | None],
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible: Sequence[VbaSymbol] | None,
    model: HostObjectModel | None = None,
) -> IntegerConstantLookup:
    return _ScopedIntegerConstantLookup(constants, symbols, proc_sym, project_visible, model)


def procedure_integer_constant_lookup(
    member: ProcedureNode,
    module_constants: Mapping[str, float | None],
    symbols: ModuleSymbols,
    project_visible: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    model: HostObjectModel | None = None,
) -> IntegerConstantLookup:
    """The integer-constant lookup for one procedure: its own `Const`s over the
    module's, resolved the way names resolve from inside that procedure."""
    own: dict[str, float | None] = {}
    collect_body_literal_integer_constants(member.body, own, activity)
    return scoped_integer_constant_lookup(
        module_constants if not own else {**module_constants, **own},
        symbols,
        procedure_symbol_for(symbols, member),
        project_visible,
        model,
    )
