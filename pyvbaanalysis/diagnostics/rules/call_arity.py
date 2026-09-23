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

from collections.abc import Callable, Mapping, Sequence

from ...completion.member_access import MemberCompletionContext
from ...parser.nodes import LeafStatementNode, ProcedureNode, Span
from ...runtime.vba_runtime import resolve_runtime_function
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureSignature, VbaSymbol
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
    runtime_arity_signature,
    runtime_callable_source_shadowed,
    same_module_callable_signatures,
    source_name_scope_for,
    unique_project_type_signatures,
)
from ..context import PushFn
from ..model import VbaDiagnosticData
from ..walker import ProcedureStatementVisitor, statement_and_branch_spans


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

        def visitor(stmt: LeafStatementNode) -> None:
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
                if _same_call_target(call, effective):
                    continue
                _validate_callable_arity(
                    source, call, same_module_signatures, project_signatures, source_names, push
                )
                _record_project_qualified_call_span(call, project_qualified_call_spans)
            for member_call in (
                *member_expression_calls(source, stmt.span, member_ctx),
                *member_statement_calls(source, stmt.span, member_ctx),
            ):
                if _call_target_span_key(member_call.call) in project_qualified_call_spans:
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
                    if _call_target_span_key(member_call.call) in project_qualified_call_spans:
                        continue
                    validate_arity(source, member_call.signature, member_call.call, push)

        return visitor

    return factory


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
