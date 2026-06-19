"""Rule family: call-argument types.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/argumentTypes.ts. When
both a callable parameter type and an argument type are known, flag high-
confidence mismatches: ByRef exact-type mismatches, non-numeric string operands
in a numeric argument, numeric-literal overflow, and scalar/object
incompatibilities. Unknowns and Variant are accepted, and VBA's normal coercions
are allowed. The parenthesized object-member call surface needs the member-
completion context and is deferred to M9 (omitting it only drops detections).
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence

from ...parser.nodes import LeafStatementNode, ProcedureNode
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureSignature, VbaSymbol
from ...types.type_inference import (
    SourceDeclaredType,
    declared_value_type_for_qualified_source_binding,
    declared_value_type_for_source_binding,
    procedure_symbol_for,
    type_environment_for,
)
from ..argument_inference import validate_argument_types
from ..call_extraction import extract_call, extract_qualified_call
from ..callable_signatures import (
    callable_type_signatures_for,
    expression_calls,
    source_name_scope_for,
)
from ..context import PushFn
from ..walker import ProcedureStatementVisitor


def check_argument_types(
    source: str,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    module_signatures = callable_type_signatures_for(symbols, project_procedures)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        proc_sym = procedure_symbol_for(symbols, member)

        def resolve_expression_type(name: str) -> SourceDeclaredType:
            return declared_value_type_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name
            )

        def resolve_qualified_expression_type(qualifier: str, name: str) -> SourceDeclaredType:
            return declared_value_type_for_qualified_source_binding(
                symbols, project_visible_symbols, qualifier, name
            )

        def visitor(stmt: LeafStatementNode) -> None:
            for call in expression_calls(source, stmt.span, module_signatures, source_names):
                validate_argument_types(
                    call, env, module_signatures, source_names, push,
                    resolve_expression_type, resolve_qualified_expression_type,
                )
            statement_call = extract_call(source, stmt.span)
            qualified_statement_call = (
                None if statement_call else extract_qualified_call(source, stmt.span, module_signatures)
            )
            effective = statement_call or qualified_statement_call
            if effective is not None:
                validate_argument_types(
                    effective, env, module_signatures, source_names, push,
                    resolve_expression_type, resolve_qualified_expression_type,
                )

        return visitor

    return factory
