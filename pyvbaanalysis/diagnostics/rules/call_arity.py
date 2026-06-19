"""Rule family: call-argument arity.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/callArity.ts. A call to a
known Sub/Function/Declare must supply an argument count the parameter list
accepts. Same-module procedures come from this module's AST; cross-module checks
use the unique exported project signatures; module-qualified calls resolve through
the named standard module only. Ambiguous or unresolved targets stay silent to
remain false-positive-free. The parenthesized object-member calls and runtime
(host) arity signatures need the member-completion / runtime surfaces and are
deferred to M9 — their omission only drops detections, never adds a false one.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence

from ...parser.nodes import LeafStatementNode, ProcedureNode
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
    runtime_arity_signature,
    runtime_callable_source_shadowed,
    same_module_callable_signatures,
    source_name_scope_for,
    unique_project_type_signatures,
)
from ..context import PushFn
from ..walker import ProcedureStatementVisitor


def check_argument_count(
    source: str,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    same_module_signatures = same_module_callable_signatures(symbols)
    project_signatures = unique_project_type_signatures(project_procedures)
    module_signatures = callable_type_signatures_for(symbols, project_procedures)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)

        def visitor(stmt: LeafStatementNode) -> None:
            statement_call = extract_call(source, stmt.span)
            qualified_statement_call = (
                None if statement_call else extract_qualified_call(source, stmt.span, module_signatures)
            )
            effective = statement_call or qualified_statement_call
            if effective is not None:
                _validate_callable_arity(
                    source, effective, same_module_signatures, project_signatures, source_names, push
                )
            for call in expression_calls(source, stmt.span, module_signatures, source_names):
                if _same_call_target(call, effective):
                    continue
                _validate_callable_arity(
                    source, call, same_module_signatures, project_signatures, source_names, push
                )

        return visitor

    return factory


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
        # Skip ambiguous same-module targets where the signature is not unique.
        if len(candidates) == 1:
            validate_arity(source, candidates[0], call, push)
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


def _same_call_target(a: CallArguments, b: CallArguments | None) -> bool:
    return (
        b is not None
        and a.name_span.start == b.name_span.start
        and a.name_span.end == b.name_span.end
    )
