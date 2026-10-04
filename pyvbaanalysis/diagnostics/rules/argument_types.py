"""Rule family: call-argument types.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/argumentTypes.ts. When
both a callable parameter type and an argument type are known, flag high-
confidence mismatches: ByRef exact-type mismatches, non-numeric string operands
in a numeric argument, numeric-literal overflow, and scalar/object
incompatibilities. Unknowns and Variant are accepted, and VBA's normal coercions
are allowed. Member calls are checked against the signature the member-completion
context binds (`ws.Range("A1")`, `p.Save "x"`).
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from typing import Any

from ...completion.member_access import MemberCompletionContext
from ...conditional import ConditionalActivityTracker
from ...constants.date_literal import date_literal_serial
from ...lexer.token_kinds import TokenKind
from ...parser.nodes import LeafStatementNode, ProcedureNode, Span
from ...symbols.symbol_model import (
    ModuleSymbols,
    SymbolVisibility,
    VbaProcedureSignature,
    VbaSymbol,
    VbaSymbolKind,
)
from ...types.type_inference import (
    VALUE_HELD,
    known_local_literal_values_at,
    procedure_symbol_for,
    source_binding_type_resolvers,
    type_environment_for,
)
from ...types.type_names import normalize_type
from ..argument_inference import validate_argument_types, validate_argument_types_for_signature
from ..call_extraction import extract_call, extract_qualified_call
from ..callable_signatures import (
    callable_type_signatures_for,
    expression_calls,
    member_expression_calls,
    member_statement_calls,
    source_name_scope_for,
)
from ..context import PushFn, statement_tokens
from ..held_objects import held_objects_at
from ..model import VbaDiagnosticData
from ..straight_line_values import ReachingAssignments, element_key, known_index, straight_line_assignments
from ..walker import (
    ProcedureStatementVisitor,
    for_each_statement_with_headers,
    raw_expression_tokens,
    statement_and_branch_spans,
    token_name,
    token_text,
)

# JavaScript's `.` stops at a line break and its `$` only at the end.
_ELEMENT_RE = re.compile(r"^([^(]+)\((.+)\)\Z")


def check_argument_types(
    source: str,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    member_ctx: MemberCompletionContext,
    push: PushFn,
    activity: ConditionalActivityTracker | None = None,
) -> ProcedureStatementVisitor:
    module_signatures = callable_type_signatures_for(symbols, project_procedures)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        proc_sym = procedure_symbol_for(symbols, member)
        resolvers = source_binding_type_resolvers(symbols, proc_sym, project_visible_symbols)
        resolve_expression_type = resolvers.resolve_expression_type
        resolve_qualified_expression_type = resolvers.resolve_qualified_expression_type
        children = (proc_sym.children if proc_sym is not None else None) or []
        # What a local holds at the statement (XLIDE issue #246), and a Variant
        # still holding Empty, a number or a String there (issue #410).
        held_at: Any = None
        variant_locals = {
            child.name.lower()
            for child in children
            if child.kind == VbaSymbolKind.LOCAL_VARIABLE
            and child.visibility != SymbolVisibility.STATIC
            and not child.is_array
            and (not child.as_type or child.as_type.lower() == "variant")
        }
        reaching: dict[int, ReachingAssignments] | None = None
        local_values_at: Any = None

        def values_at(node: LeafStatementNode) -> Any:
            nonlocal local_values_at
            if local_values_at is None:
                local_values_at = known_local_literal_values_at(source, member, symbols, activity)
            return local_values_at(node)

        def reaching_at(node: LeafStatementNode) -> ReachingAssignments | None:
            nonlocal reaching
            if reaching is None:
                reaching = straight_line_assignments(source, member.body, activity)
            # The map is keyed by id(statement); the nodes are not hashable.
            return reaching.get(id(node))

        # Where each name is first named: a Variant local is Empty there, though
        # the call may pass it ByRef. A parameter holds what the caller gave it,
        # and a Static local what an earlier call left.
        plain_locals = (
            set()
            if any(word.lower() == "static" for word in member.modifiers)
            else {
                child.name.lower()
                for child in children
                if child.kind == VbaSymbolKind.LOCAL_VARIABLE
                and child.visibility != SymbolVisibility.STATIC
            }
        )
        first_named: dict[str, int] | None = None

        def named_first_at(lower: str, offset: int) -> bool:
            nonlocal first_named
            if lower not in plain_locals:
                return False
            if first_named is None:
                found: dict[str, int] = {}
                first_named = found

                def record(node: LeafStatementNode) -> None:
                    for tok in statement_tokens(source, node.span):
                        name = token_name(tok)
                        lowered = name.lower() if name else None
                        if lowered and lowered not in found:
                            found[lowered] = node.span.start

                for_each_statement_with_headers(source, member.body, record, activity)
            return first_named.get(lower) == offset

        def visitor(stmt: LeafStatementNode) -> None:
            def held_class_of(lower: str) -> str | None:
                nonlocal held_at
                if held_at is None:
                    held_at = held_objects_at(source, member, symbols, activity)
                held: str | None = held_at(stmt).classes.get(lower)
                if held is not None:
                    return held
                if normalize_type(env.get(lower)) == "variant":
                    value = values_at(stmt).get(lower)
                    kind = value.kind if value is not None else ""
                    if kind in ("empty", "number", "string") or named_first_at(lower, stmt.span.start):
                        return VALUE_HELD
                return None

            # A Variant local a straight line has just given Null (issue #324).
            # An element of an array, "a(i)", the subscript naming it here (issue #332).
            def held_null(lower: str) -> bool:
                element = _ELEMENT_RE.match(lower)
                if element is None and lower not in variant_locals:
                    return False
                here = reaching_at(stmt)
                index = (
                    known_index(raw_expression_tokens(element.group(2)), here if here is not None else {})
                    if element is not None
                    else None
                )
                key: str | None
                if element is not None:
                    key = None if index is None else element_key(element.group(1), index)
                else:
                    key = lower
                assigned = here.get(key) if key is not None and here is not None else None
                held = (
                    [tok for tok in assigned if tok.kind is not TokenKind.COMMENT]
                    if assigned is not None
                    else None
                )
                return held is not None and len(held) == 1 and token_text(held[0]) == "null"

            # The number, or for a call to the project's own procedure the String,
            # a local holds here (issues #332 and #558).
            def held_number(lower: str) -> int | float | str | None:
                held = values_at(stmt).get(lower)
                if held is not None and held.kind in ("number", "string") and not held.content_mutated:
                    value: int | float | str = held.value
                    return value
                # A Date local a straight line has just set to a Date literal passes
                # its serial: #1/2/2000# is 36527 (issue #558).
                if normalize_type(env.get(lower)) != "date":
                    return None
                here = reaching_at(stmt)
                assigned = here.get(lower) if here is not None else None
                if assigned is None:
                    return None
                toks = [tok for tok in assigned if tok.kind is not TokenKind.COMMENT]
                if len(toks) == 1 and toks[0].kind is TokenKind.DATE_LITERAL:
                    return date_literal_serial(toks[0].raw_text)
                return None

            # `Call Two(Nothing, 1)` is found both as an expression call and as the
            # statement's call; report each argument once (issue #223).
            reported: set[tuple[int, int, str]] = set()

            def push_once(
                rule: str, message: str, span: Span, data: VbaDiagnosticData | None = None
            ) -> None:
                key = (span.start, span.end, message)
                if key not in reported:
                    reported.add(key)
                    push(rule, message, span, data)

            for call in expression_calls(source, stmt.span, module_signatures, source_names):
                validate_argument_types(
                    call=call,
                    env=env,
                    module_signatures=module_signatures,
                    source_names=source_names,
                    source=source,
                    member_ctx=member_ctx,
                    push=push_once,
                    resolve_expression_type=resolve_expression_type,
                    resolve_qualified_expression_type=resolve_qualified_expression_type,
                    held_class_of=held_class_of,
                    held_null=held_null,
                    held_number=held_number,
                )
            for member_call in (
                *member_expression_calls(source, stmt.span, member_ctx),
                *member_statement_calls(source, stmt.span, member_ctx),
            ):
                validate_argument_types_for_signature(
                    sig=member_call.signature,
                    call=member_call.call,
                    env=env,
                    module_signatures=module_signatures,
                    source_names=source_names,
                    source=source,
                    member_ctx=member_ctx,
                    push=push_once,
                    resolve_expression_type=resolve_expression_type,
                    resolve_qualified_expression_type=resolve_qualified_expression_type,
                    held_class_of=held_class_of,
                    held_null=held_null,
                    held_number=held_number,
                )
            # A single-line If's branch is a statement call too: `If x Then Sl Nothing`
            # (issue #254).
            for span in statement_and_branch_spans(stmt):
                statement_call = extract_call(source, span) or extract_qualified_call(
                    source, span, module_signatures
                )
                if statement_call is not None:
                    validate_argument_types(
                        call=statement_call,
                        env=env,
                        module_signatures=module_signatures,
                        source_names=source_names,
                        source=source,
                        member_ctx=member_ctx,
                        push=push_once,
                        resolve_expression_type=resolve_expression_type,
                        resolve_qualified_expression_type=resolve_qualified_expression_type,
                        held_class_of=held_class_of,
                        held_null=held_null,
                        held_number=held_number,
                    )

        return visitor

    return factory
