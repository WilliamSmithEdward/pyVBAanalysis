"""Rule family: object-variable state.

Ported from the object-variable-not-set rule of
xlide_vscode/src/analyzer/diagnostics/rules/objectState.ts: a local object
variable that is still Nothing when a member is accessed raises Run-time error
'91'. It tracks an unset->set lattice per local over the shared dataflow walk,
falling back to the conservative straight-line walk for procedures with
unstructured flow.

M7 slice: object typing uses the host-free is_known_object_assignment_type
(generic Object only), and the member-surface suppression (hasDefiniteMissingMember)
is deferred to M9; both are sound for the no-false-positive contract (they only
reduce precision, never invent a diagnostic). The scalar-member-access rule in
the same XLIDE family needs the type environment + member surface and lands in M8.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Mapping, Sequence

from ...conditional import ConditionalActivityTracker
from ...flow.procedure_unstructured import procedure_has_unstructured_flow
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import BodyNode, LeafStatementNode, ModuleNode, ProcedureNode, Span, WithBlockNode
from ...symbols.name_resolution import BareIdentifierContext
from ...symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbol, VbaSymbolKind
from ...types.type_inference import (
    SourceDeclaredType,
    declared_type_for_source_binding,
    procedure_symbol_for,
    type_environment_for,
)
from ...types.type_names import is_known_object_assignment_type, is_known_scalar_type, normalize_type
from ..context import PushFn, statement_tokens
from ..dataflow import (
    DataflowHooks,
    Lattice,
    tracked_locals_passed_as_call_arguments,
    walk_branch_merged_body,
    walk_straight_line_body,
)
from ..walker import (
    ProcedureStatementVisitor,
    active_module_members,
    block_header_line_span,
    is_inactive_node,
    set_assignment_target,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

_PROCEDURE_KINDS = frozenset(
    {
        VbaSymbolKind.SUB,
        VbaSymbolKind.FUNCTION,
        VbaSymbolKind.PROPERTY_GET,
        VbaSymbolKind.PROPERTY_LET,
        VbaSymbolKind.PROPERTY_SET,
    }
)


def check_scalar_member_access(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """Member access on a known scalar (`x.Foo` where x As Long) is a VBE compile error."""

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        proc_sym = procedure_symbol_for(symbols, member)

        def resolve_declared_type(name: str) -> SourceDeclaredType:
            return declared_type_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.MEMBER_RECEIVER
            )

        def visitor(stmt: LeafStatementNode) -> None:
            for name, as_type, span, vbe_error in _scalar_member_accesses(
                source, stmt.span, env, resolve_declared_type
            ):
                push(
                    "scalarMemberAccess",
                    f"Member access on '{name}' is invalid because it is declared as {as_type}. "
                    f"This is a VBE compile error: {vbe_error}.",
                    span,
                )

        return visitor

    return factory


def _scalar_member_accesses(
    source: str,
    span: Span,
    env: Mapping[str, str],
    resolve_declared_type: Callable[[str], SourceDeclaredType],
) -> list[tuple[str, str, Span, str]]:
    toks = statement_tokens(source, span)
    out: list[tuple[str, str, Span, str]] = []
    for i in range(len(toks) - 1):
        if toks[i + 1].raw_text != ".":
            continue
        if i > 0 and toks[i - 1].raw_text == ".":
            continue
        name = token_name(toks[i])
        if not name:
            continue
        declared_type = resolve_declared_type(name)
        as_type = declared_type.as_type if declared_type.resolved else env.get(name.lower())
        normalized = normalize_type(as_type)
        if not as_type or not normalized or not is_known_scalar_type(normalized):
            continue
        member_name = token_name(toks[i + 2]) if i + 2 < len(toks) else None
        vbe_error = "Invalid qualifier" if member_name else "Syntax error"
        out.append(
            (name, as_type, Span(span.start + toks[i].start, span.start + toks[i + 1].end), vbe_error)
        )
    return out


def _not_set_message(name: str, context: str) -> str:
    return (
        f"Object variable '{name}' is Nothing before {context}. This will raise "
        "Run-time error '91': Object variable or With block variable not set."
    )


def check_object_variable_not_set(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_procedure(source, member, symbols, activity, push)


def _check_procedure(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    locals_ = _local_object_variables_for(symbols, proc)
    if not locals_:
        return
    state: dict[str, str] = dict.fromkeys(locals_, "unset")

    def on_statement(stmt: LeafStatementNode) -> None:
        _check_statement(source, stmt, locals_, state, push)

    def on_block(node: BodyNode) -> None:
        if not isinstance(node, WithBlockNode):
            return
        receiver = _unset_with_receiver(source, node.span, locals_, state)
        if receiver is not None:
            name, span = receiver
            push("objectVariableNotSet", _not_set_message(name, "With member access"), span)

    def touches(stmt: LeafStatementNode) -> Iterable[str]:
        target = set_assignment_target(source, stmt.span)
        if target is None:
            return ()
        lower = target[0].lower()
        return (lower,) if lower in locals_ else ()

    def demote(lower: str) -> None:
        if state.get(lower) == "unset":
            state[lower] = "unknown"

    def restore(snapshot: Mapping[str, str]) -> None:
        state.clear()
        state.update(snapshot)

    hooks = DataflowHooks(
        on_statement=on_statement,
        touches_in_statement=touches,
        demote_to_unknown=demote,
        on_block=on_block,
        snapshot_state=lambda: dict(state),
        restore_state=restore,
        set_state=lambda key, value: state.__setitem__(key, value),
        lattice=Lattice(init="unset", good="set", unknown="unknown"),
    )
    walk = (
        walk_straight_line_body
        if procedure_has_unstructured_flow(source, proc, activity)
        else walk_branch_merged_body
    )
    walk(proc.body, lambda node: is_inactive_node(activity, node), hooks)


def _check_statement(
    source: str,
    stmt: LeafStatementNode,
    locals_: set[str],
    state: dict[str, str],
    push: PushFn,
) -> None:
    for name, span in _unset_object_member_accesses(source, stmt.span, locals_, state):
        push("objectVariableNotSet", _not_set_message(name, "member access"), span)
    target = set_assignment_target(source, stmt.span)
    if target is not None:
        lower = target[0].lower()
        if lower in locals_:
            state[lower] = "unset" if _set_value_is_nothing(target[2]) else "set"
            return
    toks = statement_tokens_after_leading_label(source, stmt.span)
    for lower in tracked_locals_passed_as_call_arguments(toks, lambda name: name in locals_):
        if state.get(lower) == "unset":
            state[lower] = "unknown"


def _unset_object_member_accesses(
    source: str, span: Span, locals_: set[str], state: Mapping[str, str]
) -> list[tuple[str, Span]]:
    toks = statement_tokens(source, span)
    out: list[tuple[str, Span]] = []
    for i in range(len(toks) - 1):
        if toks[i + 1].raw_text != "." or (i >= 1 and toks[i - 1].raw_text == "."):
            continue
        name = token_name(toks[i])
        if name is None:
            continue
        lower = name.lower()
        if lower not in locals_ or state.get(lower) != "unset":
            continue
        # The member-surface suppression (hasDefiniteMissingMember) is deferred to
        # M9: without the exhaustive surface we never suppress, which is sound for
        # the accepted-case sweep because suppression only avoids double-reporting
        # with the compile-only member-not-found rule.
        out.append((name, Span(span.start + toks[i].start, span.start + toks[i].end)))
    return out


def _set_value_is_nothing(value_tokens: Sequence[VbaToken]) -> bool:
    toks = [
        tok
        for tok in value_tokens
        if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE
    ]
    return len(toks) == 1 and token_text(toks[0]) == "nothing"


def _unset_with_receiver(
    source: str, span: Span, locals_: set[str], state: Mapping[str, str]
) -> tuple[str, Span] | None:
    header = block_header_line_span(source, span)
    toks = statement_tokens_after_leading_label(source, header)
    if len(toks) != 2 or token_text(toks[0]) != "with":
        return None
    name = token_name(toks[1])
    if name is None:
        return None
    lower = name.lower()
    if lower not in locals_ or state.get(lower) != "unset":
        return None
    return (name, Span(header.start + toks[1].start, header.start + toks[1].end))


def _local_object_variables_for(symbols: ModuleSymbols, proc: ProcedureNode) -> set[str]:
    proc_sym = _procedure_symbol_for(symbols, proc)
    if proc_sym is None or proc_sym.children is None:
        return set()
    out: set[str] = set()
    for child in proc_sym.children:
        if (
            child.kind is VbaSymbolKind.LOCAL_VARIABLE
            and child.visibility is not SymbolVisibility.STATIC
            and not child.is_array
            and is_known_object_assignment_type(child.as_type)
        ):
            out.add(child.name.lower())
    return out


def _procedure_symbol_for(symbols: ModuleSymbols, proc: ProcedureNode) -> VbaSymbol | None:
    for sym in symbols.root.children or []:
        if sym.kind in _PROCEDURE_KINDS and sym.full_span.start == proc.span.start:
            return sym
    return None
