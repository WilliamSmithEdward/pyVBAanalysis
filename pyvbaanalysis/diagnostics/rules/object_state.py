"""Rule family: object-variable state.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/objectState.ts: member
access on a variable declared as a scalar, and a local object variable that is
still Nothing when a member is accessed, which raises Run-time error '91'. The
latter tracks an unset->set lattice per local over the shared dataflow walk,
falling back to the conservative straight-line walk for procedures with
unstructured flow.

Object typing uses the host-aware is_known_object_assignment_type (from the
member-completion engine), so host-typed locals like `Dim ws As Worksheet` count as
object variables; and the member-surface suppression (hasDefiniteMissingMember) is
wired to the exhaustive member surface, so an objectVariableNotSet report is
suppressed when the member is provably missing (suppression-only: it can REMOVE a
diagnostic, never add one).
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Mapping, Sequence
from collections.abc import Set as AbstractSet

from ...completion import (
    MemberCompletionContext,
    is_known_object_assignment_type,
)
from ...conditional import ConditionalActivityTracker
from ...flow.procedure_unstructured import procedure_has_unstructured_flow
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ForBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    WithBlockNode,
    iter_body_nodes,
)
from ...symbols.name_resolution import BareIdentifierContext
from ...symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbol, VbaSymbolKind
from ...types.type_inference import (
    SourceDeclaredType,
    declared_type_for_source_binding,
    procedure_symbol_for,
    type_environment_for,
)
from ...types.type_names import is_known_scalar_type, normalize_type
from ..context import PushFn, statement_tokens
from .shared import resolve_exhaustive_member_surface
from ..dataflow import (
    DataflowHooks,
    Lattice,
    tracked_locals_named_whole,
    walk_branch_merged_body,
    walk_straight_line_body,
)
from ..walker import (
    ProcedureStatementVisitor,
    active_module_members,
    bare_assignment_target,
    block_header_line_span,
    for_each_statement,
    is_inactive_node,
    set_assignment_target,
    statement_and_branch_spans,
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
    member_ctx: MemberCompletionContext | None = None,
) -> None:
    # member_ctx is keyword-defaulted so the registry can adopt the new arg without
    # a signature break; the faithful TS order threads ctx.memberCtx here. An empty
    # context still resolves generic Object + host aliases (the model defaults to
    # Excel), losing only project-class precision.
    ctx = member_ctx if member_ctx is not None else MemberCompletionContext()
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_procedure(source, member, symbols, ctx, activity, push)


def _check_procedure(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    locals_ = _local_object_variables_for(symbols, proc, member_ctx)
    if not locals_:
        return
    state: dict[str, str] = dict.fromkeys(locals_, "unset")
    # The locals some statement anywhere in the procedure Sets: a `GoSub` may run any
    # of those statements before control comes back (XLIDE issue #108), so after it
    # none of them is provably still Nothing.
    set_anywhere: set[str] = set()

    def collect_sets(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            target = set_assignment_target(source, span)
            if target is not None and target[0].lower() in locals_:
                set_anywhere.add(target[0].lower())

    for_each_statement(proc.body, collect_sets, activity)

    def on_statement(stmt: LeafStatementNode) -> None:
        _check_statement(source, stmt, locals_, state, set_anywhere, member_ctx, push)

    def on_block(node: BodyNode) -> None:
        # A For Each that runs to its end leaves the control variable Nothing, so an
        # access after the loop is right to report. One the body can leave early -
        # Exit For, or a GoTo out of it - leaves it on the current element, so
        # nothing is proven (XLIDE issue #108: `Exit For` on the first sheet, then
        # `ws.Name`).
        if isinstance(node, ForBlockNode):
            # `For Each x In c` with c still Nothing raises 424, not 91: the loop
            # asks the collection for its enumerator (XLIDE issue #121).
            over = node.source_expression.strip().lower() if node.each and node.source_expression else None
            if over and over in locals_ and state.get(over) == "unset" and node.source_expression_span:
                push(
                    "objectVariableNotSet",
                    f"Object variable '{locals_[over]}' is Nothing when For Each asks it for its "
                    "elements. This will raise Run-time error '424': Object required.",
                    node.source_expression_span,
                )
            lower = node.control_variable.lower() if node.control_variable else None
            if (
                node.each
                and lower
                and lower in locals_
                and state.get(lower) == "unset"
                and _body_can_leave_loop(source, node, activity)
            ):
                state[lower] = "unknown"
            return
        if not isinstance(node, WithBlockNode):
            return
        receiver = _unset_with_receiver(source, node.span, locals_, state)
        if receiver is not None:
            name, span = receiver
            push("objectVariableNotSet", _not_set_message(name, "With member access"), span)

    def touches(stmt: LeafStatementNode) -> Iterable[str]:
        # A local passed whole may have been Set by the callee, inside an If arm
        # as much as outside one, so the branch merge counts it as touched.
        touched = set(_locals_passed_whole(source, stmt.span, locals_))
        # A single-line If's branches Set too.
        for span in statement_and_branch_spans(stmt):
            target = set_assignment_target(source, span)
            if target is not None and target[0].lower() in locals_:
                touched.add(target[0].lower())
        return touched

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


def _body_can_leave_loop(
    source: str, loop: ForBlockNode, activity: ConditionalActivityTracker | None
) -> bool:
    """Whether the loop body can leave the loop before it ends: an `Exit For` at its
    own depth (one inside a nested For leaves that one), or any `GoTo`."""

    def skip(node: BodyNode) -> bool:
        # A nested For's Exit For is its own.
        return is_inactive_node(activity, node) or isinstance(node, ForBlockNode)

    for node in iter_body_nodes(loop.body, skip):
        if isinstance(getattr(node, "body", None), list):
            continue
        for span in statement_and_branch_spans(node):  # type: ignore[arg-type]
            toks = statement_tokens_after_leading_label(source, span)
            head = token_text(toks[0] if toks else None)
            if (head == "exit" and token_text(toks[1] if len(toks) > 1 else None) == "for") or head == "goto":
                return True
    return False


def _nothing_guard_names(condition: Sequence[VbaToken]) -> tuple[set[str], set[str]]:
    """The tracked names a single-line If's condition guards, as (Then arm, Else
    arm): `Not d Is Nothing` guards the Then arm, `d Is Nothing` the Else arm (XLIDE
    issue #108: the block form already read the guard, the one-line form did not)."""
    then_arm: set[str] = set()
    else_arm: set[str] = set()
    for i in range(len(condition) - 2):
        if token_text(condition[i + 1]) != "is" or token_text(condition[i + 2]) != "nothing":
            continue
        name = token_name(condition[i])
        if not name:
            continue
        if token_text(condition[i - 1] if i >= 1 else None) == "not":
            then_arm.add(name.lower())
        else:
            else_arm.add(name.lower())
    return then_arm, else_arm


def _check_statement(
    source: str,
    stmt: LeafStatementNode,
    locals_: Mapping[str, str],
    state: dict[str, str],
    set_anywhere: AbstractSet[str],
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> None:
    toks = statement_tokens_after_leading_label(source, stmt.span)
    head = token_text(toks[0] if toks else None)
    # `GoSub Label` runs the subroutine, which may Set any of the locals, before the
    # statement after it (XLIDE issue #108).
    if head == "gosub" or (head == "on" and any(token_text(tok) == "gosub" for tok in toks)):
        for lower in set_anywhere:
            if state.get(lower) == "unset":
                state[lower] = "unknown"
        return
    # The arms of a single-line If and what its condition proves about them.
    branches = statement_and_branch_spans(stmt)
    then_guards: set[str] = set()
    else_guards: set[str] = set()
    if head == "if" and len(branches) > 1:
        then_index = next(
            (index for index, tok in enumerate(toks) if index > 0 and token_text(tok) == "then"), -1
        )
        if then_index > 0:
            then_guards, else_guards = _nothing_guard_names(toks[1:then_index])

    def guarded_at(name: str, offset: int) -> bool:
        def within(index: int) -> bool:
            return index < len(branches) and branches[index].start <= offset < branches[index].end

        return (name in then_guards and within(1)) or (name in else_guards and within(2))

    # A bare `obj = value` is a Let through the object's default member (XLIDE issue
    # #107), which needs an object to reach: on a variable still Nothing it raises
    # 91, the same as a member access would.
    for span in branches:
        let_target = bare_assignment_target(source, span)
        if let_target is None:
            continue
        lower = let_target[0].lower()
        if (
            lower
            and lower in locals_
            and state.get(lower) == "unset"
            and not guarded_at(lower, let_target[1].start)
        ):
            push("objectVariableNotSet", _not_set_message(let_target[0], "the default-member assignment"), let_target[1])
    passed_whole = _locals_passed_whole(source, stmt.span, locals_)
    for name, span in _unset_object_member_accesses(source, stmt.span, locals_, state, member_ctx):
        # An access after a whole pass in the same statement, as in
        # `If TryGet(obj) Then obj.Name`, runs after the callee had its chance
        # to Set it. One before the pass, as in `Load(obj.Name)`, does not.
        pass_at = passed_whole.get(name.lower())
        if pass_at is not None and span.start > pass_at:
            continue
        if guarded_at(name.lower(), span.start):
            continue
        push("objectVariableNotSet", _not_set_message(name, "member access"), span)
    target = set_assignment_target(source, stmt.span)
    if target is not None:
        lower = target[0].lower()
        if lower in locals_:
            state[lower] = "unset" if _set_value_is_nothing(target[2]) else "set"
            return
    for lower in passed_whole:
        if state.get(lower) == "unset":
            state[lower] = "unknown"
    # A Set in a single-line If's branch runs on one path only, so it moves an
    # unset object to 'unknown' the way a block If without Else does, not to
    # 'set', as unallocated-dynamic-array-access reads a conditional ReDim.
    for branch in statement_and_branch_spans(stmt)[1:]:
        branch_target = set_assignment_target(source, branch)
        branch_lower = branch_target[0].lower() if branch_target is not None else None
        if branch_lower is not None and branch_lower in locals_ and state.get(branch_lower) == "unset":
            state[branch_lower] = "unknown"


# Intrinsics that read an object argument and never Set it.
_OBJECT_READ_ONLY_INTRINSICS: frozenset[str] = frozenset(
    {"typename", "vartype", "isobject", "isnull", "isempty", "ismissing", "objptr"}
)


def _locals_passed_whole(source: str, span: Span, locals_: Mapping[str, str]) -> dict[str, int]:
    return tracked_locals_named_whole(
        statement_tokens_after_leading_label(source, span),
        span.start,
        lambda name: name in locals_,
        _OBJECT_READ_ONLY_INTRINSICS,
    )


def _unset_object_member_accesses(
    source: str,
    span: Span,
    locals_: Mapping[str, str],
    state: Mapping[str, str],
    member_ctx: MemberCompletionContext,
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
        # Suppression-only (M9): when the receiver's member surface is exhaustive and
        # provably lacks this member, the member-not-found compile rule already owns
        # the diagnostic, so do not also report the runtime "object variable not set"
        # (this can only REMOVE a report, never add one).
        member = token_name(toks[i + 2]) if i + 2 < len(toks) else None
        if member is not None and _has_definite_missing_member(
            source, span.start + toks[i + 1].end, member, member_ctx
        ):
            continue
        out.append((name, Span(span.start + toks[i].start, span.start + toks[i].end)))
    return out


def _has_definite_missing_member(
    source: str, dot_end_offset: int, member_name: str, member_ctx: MemberCompletionContext
) -> bool:
    surface = resolve_exhaustive_member_surface(source, dot_end_offset, member_ctx)
    return surface is not None and not surface.has_member(member_name)


def _set_value_is_nothing(value_tokens: Sequence[VbaToken]) -> bool:
    toks = [
        tok
        for tok in value_tokens
        if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE
    ]
    return len(toks) == 1 and token_text(toks[0]) == "nothing"


def _unset_with_receiver(
    source: str, span: Span, locals_: Mapping[str, str], state: Mapping[str, str]
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


def _local_object_variables_for(
    symbols: ModuleSymbols, proc: ProcedureNode, member_ctx: MemberCompletionContext
) -> dict[str, str]:
    """The tracked object locals, lowercased name -> declared name."""
    proc_sym = _procedure_symbol_for(symbols, proc)
    if proc_sym is None or proc_sym.children is None:
        return {}
    out: dict[str, str] = {}
    for child in proc_sym.children:
        if (
            child.kind is VbaSymbolKind.LOCAL_VARIABLE
            and child.visibility is not SymbolVisibility.STATIC
            and not child.is_array
            # `Dim x As New Invoice` is instantiated on ANY access, including the
            # first one and including after `Set x = Nothing`, so it can never be
            # Nothing when a member is touched (XLIDE issue #16).
            and not child.is_auto_instantiated
            and child.as_type
            and is_known_object_assignment_type(child.as_type, member_ctx)
        ):
            out[child.name.lower()] = child.name
    return out


def _procedure_symbol_for(symbols: ModuleSymbols, proc: ProcedureNode) -> VbaSymbol | None:
    for sym in symbols.root.children or []:
        if sym.kind in _PROCEDURE_KINDS and sym.full_span.start == proc.span.start:
            return sym
    return None
