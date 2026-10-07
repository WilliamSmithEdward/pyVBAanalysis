"""Rule family: call-argument arity.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/callArity.ts. A call to a
known Sub/Function/Declare must supply an argument count the parameter list
accepts. Same-module procedures come from this module's AST; cross-module checks
use the unique exported project signatures; module-qualified calls resolve through
the named standard module only; bare calls also resolve against the VBA runtime
arity signatures. Object member calls are checked only when the member-completion
context binds a known source or host signature. Ambiguous or unresolved targets
stay silent to remain false-positive-free.
"""

from __future__ import annotations

from ..setter_assignment import source_setter_assignment
from ...types.type_inference import procedure_symbol_for

from collections.abc import Callable, Mapping, Sequence

from ...completion.member_access import MemberCompletionContext, resolve_exact_member_completion
from ...host.host_model import resolve_host_global_member
from ...lexer.token_helpers import split_top_level_token_groups
from ...lexer.token_kinds import VbaToken
from ...parser.nodes import LeafStatementNode, ProcedureNode, Span
from ...runtime.vba_runtime import resolve_runtime_function
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureSignature, VbaSymbol
from ...types.type_inference import type_environment_for
from ...types.type_names import normalize_type
from ..call_extraction import (
    CallableTypeSignature,
    CallArguments,
    extract_call,
    extract_qualified_call,
    validate_arity,
)
from ..callable_signatures import (
    SourceNameScope,
    bare_callable_source_shadowed,
    callable_type_signatures_for,
    expression_calls,
    member_expression_calls,
    member_statement_calls,
    parse_runtime_display_signature,
    runtime_arity_signature,
    runtime_callable_source_shadowed,
    same_module_callable_signatures,
    source_name_scope_for,
    unique_project_type_signatures,
)
from ..context import PushFn, statement_tokens
from ..model import VbaDiagnosticData
from ..walker import (
    ProcedureStatementVisitor,
    match_paren_from,
    statement_and_branch_spans,
    token_name,
)


def check_argument_count(
    source: str,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> ProcedureStatementVisitor:
    same_module_signatures = same_module_callable_signatures(symbols)
    project_signatures = unique_project_type_signatures(project_procedures)
    module_signatures = callable_type_signatures_for(symbols, project_procedures)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        env = type_environment_for(symbols, member)

        def visitor(stmt: LeafStatementNode) -> None:
            for span in statement_and_branch_spans(stmt):
                _check_unmodelled_arity(source, span, env, source_names, member_ctx, push)
            setter_names = {setter.name_span.start for span in statement_and_branch_spans(stmt) if (setter := source_setter_assignment(source, span, symbols, procedure_symbol_for(symbols, member), project_visible_symbols, member_ctx)) is not None}
            project_qualified_call_spans: set[tuple[int, int]] = set()
            statement_call = extract_call(source, stmt.span)
            qualified_statement_call = (
                None if statement_call else extract_qualified_call(source, stmt.span, module_signatures)
            )
            effective = statement_call or qualified_statement_call
            if effective is not None:
                _validate_callable_arity(
                    source, effective, same_module_signatures, project_signatures, source_names, push
                )
                _record_project_qualified_call_span(effective, project_qualified_call_spans)
            for call in expression_calls(source, stmt.span, module_signatures, source_names):
                if call.name_span.start in setter_names or _same_call_target(call, effective):
                    continue
                _validate_callable_arity(
                    source, call, same_module_signatures, project_signatures, source_names, push
                )
                _record_project_qualified_call_span(call, project_qualified_call_spans)
            for member_call in (
                *member_expression_calls(source, stmt.span, member_ctx),
                *member_statement_calls(source, stmt.span, member_ctx),
            ):
                if member_call.call.name_span.start in setter_names:
                    continue
                if _call_target_span_key(
                    member_call.call
                ) in project_qualified_call_spans or _takes_print_list(member_call.signature):
                    continue
                validate_arity(source, member_call.signature, member_call.call, push)
            # A single-line If is one statement, so a CALL STATEMENT it carries,
            # `If ok Then Helper 1, 2, 3`, was never read as one (XLIDE issue #46).
            # Only the statement-call path repeats over the branches: the expression
            # scans above already cover the whole line, and running them again would
            # report the same call twice.
            for branch in statement_and_branch_spans(stmt)[1:]:
                branch_call = extract_call(source, branch) or extract_qualified_call(
                    source, branch, module_signatures
                )
                if (
                    branch_call is not None
                    and _call_target_span_key(branch_call) not in project_qualified_call_spans
                ):
                    _validate_callable_arity(
                        source, branch_call, same_module_signatures, project_signatures,
                        source_names, push,
                    )
                    _record_project_qualified_call_span(branch_call, project_qualified_call_spans)
                for member_call in member_statement_calls(source, branch, member_ctx):
                    if member_call.call.name_span.start in setter_names:
                        continue
                    if _call_target_span_key(
                        member_call.call
                    ) in project_qualified_call_spans or _takes_print_list(member_call.signature):
                        continue
                    validate_arity(source, member_call.signature, member_call.call, push)

        return visitor

    return factory


# VBA's Collection methods, which no host model carries (XLIDE issue #304).
_COLLECTION_SIGNATURES: dict[str, str] = {
    "add": "Add(Item, [Key], [Before], [After])",
    "item": "Item(Index)",
    "count": "Count()",
    "remove": "Remove(Index)",
}

# Excel's Global properties that take an index: Cells reaches Range.Item.
_INDEXED_GLOBALS: dict[str, str] = {
    "cells": "Cells([RowIndex], [ColumnIndex])",
    "range": "Range(Cell1, [Cell2])",
}

_SCALAR_TYPES: frozenset[str] = frozenset(
    {"long", "integer", "byte", "double", "single", "currency", "boolean", "longlong"}
)


def _check_unmodelled_arity(
    source: str,
    span: Span,
    env: Mapping[str, str],
    source_names: SourceNameScope | None,
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> None:
    """The calls the signature tables above do not reach (XLIDE issue #304, each
    measured in Excel 16.0): a Collection's Add, Item, Count and Remove; a bare
    Excel Global method such as Evaluate, Intersect or Union, and Cells or Range
    given more than they take; and a host property that holds a number,
    `Sheets.Count(1)`, given an argument."""
    toks = statement_tokens(source, span)

    def at(index: int) -> VbaToken | None:
        return toks[index] if 0 <= index < len(toks) else None

    def raw_at(index: int) -> str | None:
        tok = at(index)
        return tok.raw_text if tok is not None else None

    def validate(signature: str, display: str, name_index: int, slots: list[list[VbaToken]]) -> None:
        call = CallArguments(
            name=display,
            name_span=Span(span.start + toks[name_index].start, span.start + toks[name_index].end),
            slots=slots,
            slice_start=span.start,
        )
        validate_arity(source, parse_runtime_display_signature(display, signature), call, push)

    def arguments_at(open_index: int) -> list[list[VbaToken]] | None:
        close = match_paren_from(toks, open_index)
        if close < 0:
            return None
        return [] if close == open_index + 1 else split_top_level_token_groups(toks, open_index + 1, ",", close)

    for i, tok in enumerate(toks):
        name = token_name(tok)
        if not name:
            continue
        lower = name.lower()
        member = raw_at(i - 1) == "."
        # `c.Add 1`, `c.Item()`: a local As Collection's own methods.
        receiver_name = token_name(at(i - 2)) if member else None
        receiver = receiver_name.lower() if receiver_name else None
        if (
            member
            and receiver
            and raw_at(i - 3) != "."
            and normalize_type(env.get(receiver)) == "collection"
            and lower in _COLLECTION_SIGNATURES
        ):
            # A project class named Collection keeps its own members.
            own = resolve_exact_member_completion(source, name, span.start + tok.end, member_ctx)
            if own is not None and own.definitions is not None:
                continue
            statement = i == 2 and raw_at(i + 1) != "(" and raw_at(i + 1) != "="
            slots: list[list[VbaToken]] | None
            if raw_at(i + 1) == "(":
                slots = arguments_at(i + 1)
            elif statement:
                slots = split_top_level_token_groups(toks, i + 1, ",", len(toks)) if i + 1 < len(toks) else []
            else:
                slots = None
            if slots is not None:
                validate(_COLLECTION_SIGNATURES[lower], name, i, slots)
            continue
        if raw_at(i + 1) != "(":
            continue
        # `Evaluate()`, `Union(r)`, `Cells(1, 1, 1)`: a bare Excel Global member.
        if (
            not member
            and not bare_callable_source_shadowed(name, source_names)
            and not runtime_callable_source_shadowed(name, source_names)
            and lower not in env
        ):
            global_member = resolve_host_global_member(name, member_ctx.model)
            signature = (
                global_member.get("signature")
                if global_member is not None and global_member.get("kind") == "method"
                else _INDEXED_GLOBALS.get(lower) if global_member is not None else None
            )
            global_slots = arguments_at(i + 1) if signature else None
            if signature and global_slots is not None:
                validate(signature, name, i, global_slots)
            continue
        # `Sheets.Count(1)`: a host property that holds a number takes no argument.
        if member:
            resolved = resolve_exact_member_completion(source, name, span.start + tok.end, member_ctx)
            declared = resolved.declared_type if resolved is not None else None
            value_type = normalize_type(
                declared if declared is not None else resolved.returns if resolved is not None else None
            )
            member_slots = arguments_at(i + 1)
            if (
                resolved is not None
                and resolved.kind == "property"
                and not resolved.signature
                and resolved.definitions is None
                and resolved.let_accessor is None
                and value_type
                and value_type in _SCALAR_TYPES
                and member_slots
            ):
                validate(f"{resolved.name}()", name, i, member_slots)


def _takes_print_list(signature: CallableTypeSignature) -> bool:
    """A host method named Print, a VB6 form's or picture box's (XLIDE issue #358)
    or an Access report's, takes what the Print statement takes:
    `Form1.Print "a"; x`. Its listed signature has no parameters, so its arity is
    not judged."""
    return signature.name.lower() == "print"


def _record_project_qualified_call_span(call: CallArguments, out: set[tuple[int, int]]) -> None:
    if call.lookup_key:
        out.add(_call_target_span_key(call))


def _call_target_span_key(call: CallArguments) -> tuple[int, int]:
    return (call.name_span.start, call.name_span.end)


def _validate_callable_arity(
    source: str,
    call: CallArguments,
    same_module_signatures: Mapping[str, list[CallableTypeSignature]],
    project_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None,
    push: PushFn,
) -> None:
    lower = call.lookup_key or call.name.lower()
    if not call.qualifier and bare_callable_source_shadowed(call.name, source_names):
        return
    candidates = None if call.qualifier else same_module_signatures.get(call.name.lower())
    if candidates is not None:
        if len(candidates) == 1:
            validate_arity(source, candidates[0], call, push)
            return
        # Several same-module signatures share the name. Which one a build
        # compiles is unknown, but when NONE of them accepts this call it is wrong
        # under every build. One declaration per arm of a `#If` chain is the shape
        # that makes this common (XLIDE issue #58).
        rejections = [_arity_rejections(source, signature, call) for signature in candidates]
        if all(rejections):
            rule, message, span, data = rejections[0][0]
            push(rule, message, span, data)
        return
    project_signature = project_signatures.get(lower)
    if project_signature is not None:
        validate_arity(source, project_signature, call, push)
        return
    if not call.qualifier:
        if runtime_callable_source_shadowed(call.name, source_names):
            return
        runtime = resolve_runtime_function(call.name)
        runtime_signature = runtime_arity_signature(runtime) if runtime is not None else None
        if runtime_signature is not None:
            validate_arity(source, runtime_signature, call, push)


_Rejection = tuple[str, str, Span, VbaDiagnosticData | None]


def _arity_rejections(
    source: str, signature: CallableTypeSignature, call: CallArguments
) -> list[_Rejection]:
    """What validate_arity would report for `call` against `signature`."""
    hits: list[_Rejection] = []

    def collect(rule: str, message: str, span: Span, data: VbaDiagnosticData | None = None) -> None:
        hits.append((rule, message, span, data))

    validate_arity(source, signature, call, collect)
    return hits


def _same_call_target(a: CallArguments, b: CallArguments | None) -> bool:
    return (
        b is not None
        and a.name_span.start == b.name_span.start
        and a.name_span.end == b.name_span.end
    )
