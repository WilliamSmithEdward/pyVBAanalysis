"""Straight-line local-state dataflow shared by diagnostics rules.

Ported from xlide_vscode/src/analyzer/diagnostics/dataflow.ts. The
object-variable-not-set and unallocated-dynamic-array rules both track a small
three-state lattice per procedure local over straight-line statements: every
tracked local starts in the rule's initial state, moves through rule-specific
transitions on plain statements, and demotes to "unknown" when the variable may
be rebound on a path the rule does not model (passed as a bare, potentially
ByRef call argument, or touched anywhere inside a nested runtime block). This
module owns that shared walk and the call-argument escape scan so the escape
analysis cannot drift between rules; each rule supplies its own transitions and
touch detection.

Upstream's walks recurse once per nested block. Here each recursive function is
a generator that yields the sub-walk it would have called and receives its
result, and `_drive` runs them on an explicit stack, so nesting depth is not
bounded by Python's recursion limit.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Generator, Iterable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass
from typing import Any, Generic, TypeVar

from ..flow.procedure_labels import statement_label_declarations, statement_label_references
from ..lexer.token_helpers import token_name, token_word, tokens_without_leading_line_number
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import (
    BodyNode,
    IfBlockNode,
    IfBranchKind,
    IfBranchNode,
    LeafStatementNode,
    SelectBlockNode,
    Span,
    StatementNode,
    WithBlockNode,
    is_leaf_statement,
    iter_body_nodes,
)
from .block_headers import block_header_leaves, is_loop_block, select_arms
from .context import statement_tokens

_NO_NAMES: frozenset[str] = frozenset()

# A sub-walk: yields the sub-walks it calls, receives each one's result.
_Task = Generator[Any, Any, Any]

# Upstream's CalleeKeepsArgument (calleeArguments.ts): whether the module's own
# procedure `callee` keeps the argument at `index`, or the one named `named`.
CalleeKeeps = Callable[[str, int, "str | None"], bool]


def _drive(task: _Task) -> Any:
    """Run a generator walk and every sub-walk it yields on an explicit stack."""
    stack: list[_Task] = [task]
    value: Any = None
    while stack:
        try:
            child = stack[-1].send(value)
        except StopIteration as stop:
            stack.pop()
            value = stop.value
            continue
        stack.append(child)
        value = None
    return value


@dataclass(frozen=True, slots=True)
class Lattice:
    """The rule's good/init/unknown labels, kept rule-agnostic for the merge."""

    init: str
    good: str
    unknown: str


@dataclass(frozen=True, slots=True)
class StraightLineDataflowHooks:
    """Rule-specific hooks driving one straight-line dataflow walk.

    The branch-merge hooks (snapshot_state / restore_state / set_state / lattice)
    are optional: without them an If block is treated conservatively.
    """

    # Applies one straight-line statement's transitions and diagnostics.
    on_statement: Callable[[LeafStatementNode], None]
    # Lowercased tracked names one nested-block statement touches.
    touches_in_statement: Callable[[LeafStatementNode], Iterable[str]]
    # Demotes one tracked name to the rule's "unknown" state.
    demote_to_unknown: Callable[[str], None]
    # Inspects one non-statement node before its body's touch demotion.
    on_block: Callable[[BodyNode], None] | None = None
    # Sets what a loop leaves once its touch demotion is done: a For Each's
    # control variable (XLIDE issue #336).
    after_block: Callable[[BodyNode], None] | None = None
    # Snapshot every tracked name's current state, for forking If arms.
    snapshot_state: Callable[[], dict[str, str]] | None = None
    # Overwrite the live state from a snapshot, restoring it before the next arm.
    restore_state: Callable[[Mapping[str, str]], None] | None = None
    # Write one tracked name's merged post-block state.
    set_state: Callable[[str, str], None] | None = None
    # The rule's good/init labels so the branch merge stays rule-agnostic.
    lattice: Lattice | None = None
    # Drops the rule's findings while true: the GoTo-following walk runs the body
    # more than once to settle what each label is entered with, and reports on
    # its last run only (XLIDE issue #271).
    set_silent: Callable[[bool], None] | None = None
    # What an If condition comes to from the rule's own state, None when that is
    # not certain: `c Is Nothing` with c never Set. An arm the condition rules
    # out is not walked, and one that always leaves ends the path (#273).
    known_condition: Callable[[Sequence[VbaToken]], bool | None] | None = None


# The port's earlier name for the hooks.
DataflowHooks = StraightLineDataflowHooks


def walk_straight_line_body(
    source: str,
    body: Sequence[BodyNode],
    is_inactive: Callable[[BodyNode], bool],
    hooks: StraightLineDataflowHooks,
) -> None:
    """Walk the straight-line statements of a procedure body: plain statements run
    the rule's transitions in order, while nested blocks (If/For/Do/...) are not
    entered - every tracked name touched anywhere inside them is demoted to
    "unknown" instead of guessing which runtime path executes."""
    if not _walk_following_jumps(source, body, is_inactive, hooks):
        _drive(
            _walk_body(
                source,
                body,
                is_inactive,
                hooks,
                _only_error_mode_statements(source, body, is_inactive),
            )
        )


def _single_line_goto_only(
    source: str, leaf: LeafStatementNode, next_node: BodyNode | None
) -> bool:
    """`If cond Then GoTo L`, with no Else and nothing after the GoTo."""
    branches = leaf.single_line_if_branches if isinstance(leaf, StatementNode) else None
    if branches is None or len(branches) != 1 or (
        next_node is not None and _is_single_line_if_tail(next_node)
    ):
        return False
    then = _tokens_after_label(source, branches[0])
    return len(then) == 2 and token_word(then[0]) == "goto"


# Whether the body's only unstructured statements are `On Error Resume Next`,
# `On Error GoTo 0` and `On Error GoTo -1`: no label, no jump to one and no
# Resume. Control then runs in order, a failing statement going on to the next
# under Resume Next, so blocks are entered as in a structured body. What Resume
# Next keeps from raising is dropped from the report where its stretch is known.
_JUMPING_HEADS = frozenset(("resume", "gosub", "return"))


def _only_error_mode_statements(
    source: str, body: Sequence[BodyNode], is_inactive: Callable[[BodyNode], bool]
) -> bool:
    for leaf in _collect_leaves(body, _never):
        if is_inactive(leaf):
            continue
        if statement_label_references(source, leaf.span) or statement_label_declarations(
            source, leaf.span
        ):
            return False
        # A Resume, GoSub or Return statement, alone or after a single-line If's Then.
        toks = _tokens_after_label(source, leaf.span)
        for k, tok in enumerate(toks):
            if token_word(tok) in _JUMPING_HEADS and (
                k == 0
                or toks[k - 1].raw_text == ":"
                or token_word(toks[k - 1]) == "then"
                or token_word(toks[k - 1]) == "else"
            ):
                return False
    return True


def _never(_node: BodyNode) -> bool:
    return False


# The most runs the GoTo-following walk takes to settle its labels.
_MAX_JUMP_PASSES = 6

_CASE_ELSE_RE = re.compile(r"^\s*case\s+else\b", re.IGNORECASE)


def _walk_following_jumps(
    source: str,
    body: Sequence[BodyNode],
    is_inactive: Callable[[BodyNode], bool],
    hooks: StraightLineDataflowHooks,
) -> bool:
    """Walks the top-level statements following GoTo (XLIDE issue #271, measured in
    Excel 16.0). A label is entered with the state that falls into it, if the
    statement before it can, merged with the state at each unconditional
    top-level `GoTo` to it: `GoTo Setup` ... `Setup: Set c = ...: GoTo Use`
    reaches `Use:` with c set, and `GoTo Use` past the Set reaches it with c still
    Nothing. A jump that may or may not run - a GoTo in a single-line If or a
    block, On Error GoTo, GoSub, Resume to a label, On ... GoTo - enters its label
    with nothing known. Code nothing reaches is not checked. The body is run until
    the labels settle, then once more to report. A Resume with no label, which
    returns into the body anywhere, keeps the plain walk: False is returned then,
    and when the rule cannot snapshot its state."""
    snapshot_state = hooks.snapshot_state
    restore_state = hooks.restore_state
    lattice = hooks.lattice
    set_silent = hooks.set_silent
    if snapshot_state is None or restore_state is None or lattice is None or set_silent is None:
        return False
    leaves = _collect_leaves(body, is_inactive)
    # Every label counts here, one that never runs too: the procedure is still
    # walked label by label, and a dead label is skipped in the run. Without it
    # the plain walk took over and entered no block (XLIDE issue #439).
    every_leaf = _collect_leaves(body, _never)
    if not any(
        statement_label_references(source, leaf.span)
        or statement_label_declarations(source, leaf.span)
        for leaf in every_leaf
    ):
        return False
    for leaf in leaves:
        toks = _tokens_after_label(source, leaf.span)
        for k, tok in enumerate(toks):
            if token_word(tok) == "resume" and (
                k + 1 >= len(toks)
                or token_word(toks[k + 1]) == "next"
                or toks[k + 1].kind is TokenKind.COMMENT
            ):
                return False
    initial = snapshot_state()
    unknown_state = {key: lattice.unknown for key in initial}

    def merge(states: Sequence[Mapping[str, str]]) -> dict[str, str]:
        out: dict[str, str] = {}
        for key in initial:
            values = {state.get(key, lattice.unknown) for state in states}
            out[key] = next(iter(values)) if len(values) == 1 else lattice.unknown
        return out

    def run(entries: Mapping[str, Mapping[str, str]]) -> dict[str, dict[str, str]]:
        restore_state(initial)
        reaching: dict[str, list[dict[str, str]]] = {}

        def arrive(key: str, state: Mapping[str, str]) -> None:
            reaching.setdefault(key, []).append(dict(state))

        def run_list(list_: Sequence[BodyNode], reachable_in: bool) -> _Task:
            """Walks one statement list; returns whether its end is reached."""
            reachable = reachable_in
            i = 0
            while i < len(list_):
                node = list_[i]
                if is_inactive(node):
                    i += 1
                    continue
                if is_leaf_statement(node):
                    labels = [label.key for label in statement_label_declarations(source, node.span)]
                    if labels:
                        states: list[Mapping[str, str]] = [snapshot_state()] if reachable else []
                        states.extend(entries[key] for key in labels if key in entries)
                        reachable = len(states) > 0
                        if reachable:
                            restore_state(merge(states))
                if not reachable:
                    i += 1
                    continue
                guard = _known_single_line_if(source, list_, i, hooks)
                if guard is not None and guard.known is False:
                    i += len(guard.group)
                    continue
                if guard is not None and guard.leaves:
                    # It runs, and a GoTo in it arrives with what holds after it.
                    i += len(guard.group)
                    hooks.on_statement(guard.group[0])
                    _walk_single_line_if_tail(guard.group[1:], hooks)
                    for stmt in guard.group:
                        for ref in statement_label_references(source, stmt.span):
                            arrive(
                                ref.key,
                                snapshot_state() if ref.statement_kind == "goto" else unknown_state,
                            )
                    reachable = False
                    continue
                if _is_single_line_if_tail(node):
                    tail: list[LeafStatementNode] = []
                    while i < len(list_):
                        tail_node = list_[i]
                        if not (is_leaf_statement(tail_node) and tail_node.single_line_if_tail):
                            break
                        tail.append(tail_node)
                        i += 1
                    _walk_single_line_if_tail(tail, hooks)
                    for stmt in tail:
                        for ref in statement_label_references(source, stmt.span):
                            arrive(ref.key, unknown_state)
                    continue
                if is_leaf_statement(node):
                    hooks.on_statement(node)
                    toks = _tokens_after_label(source, node.span)
                    head = token_word(_at(toks, 0))
                    # A single-line If starts with If, so only a plain GoTo, Exit,
                    # Resume, Return, End or Err.Raise ends the path here. A one-line
                    # If whose only branch is `GoTo L` changes nothing on the way,
                    # and arrives with what holds (XLIDE issue #614).
                    only_goto = head == "if" and _single_line_goto_only(
                        source, node, list_[i + 1] if i + 1 < len(list_) else None
                    )
                    for ref in statement_label_references(source, node.span):
                        arrive(
                            ref.key,
                            snapshot_state()
                            if ref.statement_kind == "goto" and (head == "goto" or only_goto)
                            else unknown_state,
                        )
                    if leaves_the_list(source, node.span):
                        reachable = False
                    i += 1
                    continue
                if hooks.on_block is not None:
                    hooks.on_block(node)
                child = getattr(node, "body", None)
                if not isinstance(child, list):
                    i += 1
                    continue
                # An If runs one arm, or none without an Else; a Select one Case, or
                # none without Case Else; a With its body once. Each starts from the
                # block's entry state, and what follows merges where they end. A
                # loop may run any number of times, so it forgets what it touches.
                arms: list[list[BodyNode]] | None
                if isinstance(node, IfBlockNode):
                    arms = [branch.body for branch in node.branches]
                elif isinstance(node, SelectBlockNode):
                    arms = select_arms(source, node.body)
                elif isinstance(node, WithBlockNode):
                    arms = [node.body]
                else:
                    arms = None
                if arms is None:
                    # A GoTo in a loop leaves with the state the loop started with,
                    # less whatever the loop may have changed.
                    touched = _block_touches(source, node, is_inactive, hooks.touches_in_statement)
                    leaving = snapshot_state()
                    for lower in touched:
                        leaving[lower] = lattice.unknown
                    for leaf in _collect_leaves(child, is_inactive):
                        for ref in statement_label_references(source, leaf.span):
                            arrive(ref.key, leaving)
                    for lower in touched:
                        hooks.demote_to_unknown(lower)
                    if hooks.after_block is not None:
                        hooks.after_block(node)
                    i += 1
                    continue
                for lower in _header_touches(source, node, hooks.touches_in_statement):
                    hooks.demote_to_unknown(lower)
                # An If runs only the arms its known conditions allow.
                if_arms = (
                    _if_arms_that_may_run(source, node, hooks)
                    if isinstance(node, IfBlockNode)
                    else None
                )
                if if_arms is not None:
                    arms = [branch.body for branch in if_arms[0]]
                entry = snapshot_state()
                ends: list[Mapping[str, str]] = []
                for arm in arms:
                    restore_state(entry)
                    if (yield run_list(arm, True)):
                        ends.append(snapshot_state())
                exhaustive = (
                    isinstance(node, WithBlockNode)
                    or (if_arms is not None and if_arms[1])
                    or (
                        isinstance(node, SelectBlockNode)
                        and any(
                            is_leaf_statement(stmt)
                            and _CASE_ELSE_RE.search(source[stmt.span.start : stmt.span.end])
                            is not None
                            for stmt in node.body
                        )
                    )
                )
                if not exhaustive:
                    ends.append(entry)
                reachable = len(ends) > 0
                if reachable:
                    restore_state(merge(ends))
                i += 1
            return reachable

        _drive(run_list(body, True))
        return {key: merge(states) for key, states in reaching.items()}

    entries: dict[str, dict[str, str]] = {}
    set_silent(True)
    for _pass in range(_MAX_JUMP_PASSES):
        nxt = run(entries)
        settled = len(nxt) == len(entries) and all(
            key in entries and all(entries[key].get(name) == value for name, value in state.items())
            for key, state in nxt.items()
        )
        entries = nxt
        if settled:
            break
    set_silent(False)
    run(entries)
    return True


def _collect_leaves(
    body: Sequence[BodyNode], is_inactive: Callable[[BodyNode], bool]
) -> list[LeafStatementNode]:
    return [node for node in iter_body_nodes(body, is_inactive) if is_leaf_statement(node)]


def _tokens_after_label(source: str, span: Span) -> list[VbaToken]:
    """A statement's tokens after any leading label or line number."""
    toks = tokens_without_leading_line_number(statement_tokens(source, span))
    if statement_label_declarations(source, span) and len(toks) > 1 and toks[1].raw_text == ":":
        toks = toks[2:]
    return toks


def walk_branch_merged_body(
    source: str,
    body: Sequence[BodyNode],
    is_inactive: Callable[[BodyNode], bool],
    hooks: StraightLineDataflowHooks,
) -> None:
    """Like walk_straight_line_body, but intersects the per-branch state of an
    If/ElseIf/Else block instead of blanket-demoting every name it touches. Each
    arm is walked from the block's entry state; a tracked name advances to its
    "good" state after the If only when it reaches "good" on EVERY arm AND a
    syntactic else arm is present, otherwise it follows the conservative
    demotion. Names a balanced If never touches keep their entry state.
    Callers must supply snapshot_state/restore_state/set_state/lattice; without
    them an If is treated conservatively.

    Only sound for procedures WITHOUT unstructured control flow (labels, GoTo, On
    Error, Resume): callers gate on procedure_has_unstructured_flow and fall back
    to walk_straight_line_body when it holds.
    """
    _drive(_walk_body(source, body, is_inactive, hooks, True))


def _walk_body(
    source: str,
    body: Sequence[BodyNode],
    is_inactive: Callable[[BodyNode], bool],
    hooks: StraightLineDataflowHooks,
    merge_if_blocks: bool,
    # Names an enclosing loop changes: a nested block may run on a later pass.
    loop_touched: AbstractSet[str] = _NO_NAMES,
) -> _Task:
    """The walk both entry points share; merging If arms is the one place they
    differ. Returns whether the end of the body is reached: with If arms merged,
    a statement that always leaves ends the path (XLIDE issue #273)."""
    i = 0
    while i < len(body):
        node = body[i]
        if is_inactive(node):
            i += 1
            continue
        # A single-line If runs its statements on some passes only.
        if merge_if_blocks and _is_conditional_leaf(node):
            for lower in loop_touched:
                hooks.demote_to_unknown(lower)
        guard = _known_single_line_if(source, body, i, hooks) if merge_if_blocks else None
        if guard is not None and guard.known is False:
            i += len(guard.group)
            continue
        if guard is not None and guard.leaves:
            hooks.on_statement(guard.group[0])
            _walk_single_line_if_tail(guard.group[1:], hooks)
            return False
        if _is_single_line_if_tail(node):
            tail: list[LeafStatementNode] = []
            while i < len(body):
                tail_node = body[i]
                if not (is_leaf_statement(tail_node) and tail_node.single_line_if_tail):
                    break
                tail.append(tail_node)
                i += 1
            _walk_single_line_if_tail(tail, hooks)
            continue
        if is_leaf_statement(node):
            hooks.on_statement(node)
            if merge_if_blocks and leaves_the_list(source, node.span):
                return False
            i += 1
            continue
        if hooks.on_block is not None:
            hooks.on_block(node)
        if merge_if_blocks:
            for lower in loop_touched:
                hooks.demote_to_unknown(lower)
        if (
            merge_if_blocks
            and isinstance(node, IfBlockNode)
            and hooks.snapshot_state is not None
            and hooks.restore_state is not None
            and hooks.set_state is not None
            and hooks.lattice is not None
        ):
            if not (yield _merge_if_block(source, node, is_inactive, hooks, loop_touched)):
                return False
            i += 1
            continue
        child = getattr(node, "body", None)
        if isinstance(child, list):
            touched = _block_touches(source, node, is_inactive, hooks.touches_in_statement)
            if merge_if_blocks and hooks.snapshot_state is not None and hooks.restore_state is not None:
                yield _walk_block_from_entry(source, node, touched, is_inactive, hooks, loop_touched)
            for lower in touched:
                hooks.demote_to_unknown(lower)
            if hooks.after_block is not None:
                hooks.after_block(node)
        i += 1
    return True


def _walk_block_from_entry(
    source: str,
    node: BodyNode,
    touched: AbstractSet[str],
    is_inactive: Callable[[BodyNode], bool],
    hooks: StraightLineDataflowHooks,
    loop_touched: AbstractSet[str],
) -> _Task:
    """Checks the statements of a For, Do, While, Select or With block with the
    state the block is entered with (XLIDE issue #237), once its own lines have
    run: a block that never touches a name leaves what is known about it as it
    was. A statement directly in a loop's body runs on the first pass as it
    stands; a block nested in the loop may run on a later pass, and forgets what
    the loop changes before it is walked. Each Case of a Select starts from the
    entry state. The state after the block is the caller's to set."""
    snapshot_state = hooks.snapshot_state
    restore_state = hooks.restore_state
    assert snapshot_state is not None and restore_state is not None
    for lower in _header_touches(source, node, hooks.touches_in_statement):
        hooks.demote_to_unknown(lower)
    entry = snapshot_state()
    if isinstance(node, SelectBlockNode):
        for arm in select_arms(source, node.body):
            restore_state(entry)
            yield _walk_body(source, arm, is_inactive, hooks, True, loop_touched)
    else:
        child: list[BodyNode] = getattr(node, "body")
        yield _walk_body(
            source,
            child,
            is_inactive,
            hooks,
            True,
            loop_touched | touched if is_loop_block(node) else loop_touched,
        )
    restore_state(entry)


def _is_conditional_leaf(node: BodyNode) -> bool:
    """A single-line If, or a statement it runs after a colon."""
    return is_leaf_statement(node) and (
        node.single_line_if_tail
        or (isinstance(node, StatementNode) and node.single_line_if_branches is not None)
    )


def _is_single_line_if_tail(node: BodyNode) -> bool:
    return is_leaf_statement(node) and node.single_line_if_tail


# Statement heads after which the rest of the list does not run.
_LIST_LEAVING_HEADS = frozenset(("exit", "goto", "return"))


def leaves_the_list(source: str, span: Span, raise_leaves: bool = True) -> bool:
    """Whether a statement always leaves the list it is in: Exit, GoTo, Return, a
    bare End, and Resume and Err.Raise unless the procedure resumes past errors:
    under On Error Resume Next a Resume with no error pending raises 20, which is
    skipped (XLIDE issues #273, #446)."""
    toks = _tokens_after_label(source, span)
    head = token_word(_at(toks, 0))
    second = _at(toks, 1)
    return (
        head in _LIST_LEAVING_HEADS
        or (head == "end" and len(toks) == 1)
        or (raise_leaves and head == "resume")
        or (
            raise_leaves
            and head == "err"
            and second is not None
            and second.raw_text == "."
            and token_word(_at(toks, 2)) == "raise"
        )
    )


@dataclass(frozen=True, slots=True)
class _KnownGuard:
    group: list[LeafStatementNode]
    known: bool
    leaves: bool


def _known_single_line_if(
    source: str,
    list_: Sequence[BodyNode],
    i: int,
    hooks: StraightLineDataflowHooks,
) -> _KnownGuard | None:
    """A single-line If with no Else and the statements after its colons, when the
    rule knows its condition (XLIDE issue #273): false runs none of them, and true
    runs them all in order, which leaves when one of them does."""
    from .condition_value import if_condition_tokens

    node = list_[i]
    known_condition = hooks.known_condition
    if (
        known_condition is None
        or not isinstance(node, StatementNode)
        or node.single_line_if_branches is None
        or len(node.single_line_if_branches) != 1
    ):
        return None
    condition = if_condition_tokens(_tokens_after_label(source, node.span))
    known = known_condition(condition) if condition is not None else None
    if known is None:
        return None
    group: list[LeafStatementNode] = [node]
    for k in range(i + 1, len(list_)):
        tail_node = list_[k]
        if not (is_leaf_statement(tail_node) and tail_node.single_line_if_tail):
            break
        group.append(tail_node)
    branch = node.single_line_if_branches[0]
    leaves = known and any(
        leaves_the_list(source, branch if k == 0 else stmt.span) for k, stmt in enumerate(group)
    )
    return _KnownGuard(group=group, known=known, leaves=leaves)


def _if_arms_that_may_run(
    source: str,
    if_block: IfBlockNode,
    hooks: StraightLineDataflowHooks,
) -> tuple[list[IfBranchNode], bool]:
    """The arms of an If block that may run, and whether one of them always does
    (XLIDE issue #273). A condition known false drops its arm, and one known true,
    or an Else, drops every arm after it."""
    from .condition_value import if_condition_tokens

    arms: list[IfBranchNode] = []
    for branch in if_block.branches:
        if branch.branch_kind is IfBranchKind.ELSE:
            arms.append(branch)
            return arms, True
        known_condition = hooks.known_condition
        condition = (
            if_condition_tokens(_tokens_after_label(source, branch.header_span))
            if known_condition is not None
            else None
        )
        known = (
            known_condition(condition)
            if known_condition is not None and condition is not None
            else None
        )
        if known is not False:
            arms.append(branch)
        if known is True:
            return arms, True
    return arms, False


def _walk_single_line_if_tail(
    tail: Sequence[LeafStatementNode], hooks: StraightLineDataflowHooks
) -> None:
    """The statements a single-line If runs after a colon, `b` in `If x Then a: b`,
    run only with its branch (MS-VBAL 5.4.2.9). They are checked on that path, and
    afterwards the state is what a block If without Else leaves: as it was before
    them, with every name they touch made unknown."""
    entry = hooks.snapshot_state() if hooks.snapshot_state is not None else None
    if entry is not None and hooks.restore_state is not None:
        for stmt in tail:
            hooks.on_statement(stmt)
        hooks.restore_state(entry)
    for stmt in tail:
        for lower in hooks.touches_in_statement(stmt):
            hooks.demote_to_unknown(lower)


def _merge_if_block(
    source: str,
    if_block: IfBlockNode,
    is_inactive: Callable[[BodyNode], bool],
    hooks: StraightLineDataflowHooks,
    loop_touched: AbstractSet[str],
) -> _Task:
    """Intersects the per-arm state of one If block (see walk_branch_merged_body).
    An arm that always leaves takes no part, and False is returned when no path
    goes past the block (XLIDE issue #273)."""
    snapshot_state = hooks.snapshot_state
    restore_state = hooks.restore_state
    set_state = hooks.set_state
    lattice = hooks.lattice
    assert (
        snapshot_state is not None
        and restore_state is not None
        and set_state is not None
        and lattice is not None
    )
    touched = _block_touches(source, if_block, is_inactive, hooks.touches_in_statement)
    # Each arm is checked from the block's entry state, once its conditions have
    # run: `If TryGet(k, obj) Then` sets obj for the arm (XLIDE issue #237).
    for lower in _header_touches(source, if_block, hooks.touches_in_statement):
        hooks.demote_to_unknown(lower)
    # Only the arms its known conditions allow run; when that is one arm that
    # always runs, the state after the block is the state it ends with.
    arms, exhaustive = _if_arms_that_may_run(source, if_block, hooks)
    if not arms:
        return True
    if len(arms) == 1 and exhaustive:
        reached: bool = yield _walk_body(
            source, arms[0].body, is_inactive, hooks, True, loop_touched
        )
        return reached
    entry = snapshot_state()
    arm_states: list[dict[str, str]] = []
    for branch in arms:
        restore_state(entry)
        if (yield _walk_body(source, branch.body, is_inactive, hooks, True, loop_touched)):
            arm_states.append(snapshot_state())
    restore_state(entry)
    if not arm_states:
        # Every arm leaves: only the path that skips the block goes on.
        return not exhaustive
    if not exhaustive:
        # No else arm: the empty fall-through path keeps the entry state, so a name
        # can only remain "good" after the block if it was already "good".
        # Reproduce the conservative behavior by demoting every touched name.
        for lower in touched:
            hooks.demote_to_unknown(lower)
        return True
    for lower in touched:
        fallback = entry.get(lower, lattice.unknown)
        set_state(lower, _join_branch_states(arm_states, lower, fallback, lattice))
    return True


def _join_branch_states(
    arm_states: Sequence[Mapping[str, str]],
    lower: str,
    fallback: str,
    lattice: Lattice,
) -> str:
    """Meet-toward-unknown join over an If block's arms for one tracked name.

    "good" only when every arm ends "good"; any unknown arm or any disagreement
    collapses to "unknown".
    """
    all_good = True
    all_init = True
    for arm in arm_states:
        state = arm.get(lower, fallback)
        if state == lattice.unknown:
            return lattice.unknown
        if state != lattice.good:
            all_good = False
        if state != lattice.init:
            all_init = False
    if all_good:
        return lattice.good
    return lattice.init if all_init else lattice.unknown


def _header_touches(
    source: str,
    node: BodyNode,
    touches: Callable[[LeafStatementNode], Iterable[str]],
) -> set[str]:
    """The tracked names a block's own lines pass on: `If TryGet(k, obj) Then`."""
    out: set[str] = set()
    for header in block_header_leaves(source, node):
        out.update(touches(header))
    return out


def _block_touches(
    source: str,
    node: BodyNode,
    is_inactive: Callable[[BodyNode], bool],
    touches: Callable[[LeafStatementNode], Iterable[str]],
) -> set[str]:
    """The tracked names a block may change: those its body touches, and those its
    own lines pass on, `If TryGet(k, obj) Then` and a For Each's control variable
    among them (XLIDE issue #237). Upstream recurses through nested blocks; the
    same set comes from one walk over every nested node."""
    child: list[BodyNode] = getattr(node, "body", None) or []
    out: set[str] = set()
    for nested in iter_body_nodes(child, is_inactive):
        if is_leaf_statement(nested):
            out.update(touches(nested))
        elif isinstance(getattr(nested, "body", None), list):
            out.update(_header_touches(source, nested, touches))
    out.update(_header_touches(source, node, touches))
    return out


_S = TypeVar("_S")


@dataclass(frozen=True, slots=True)
class BlockEnteringState(Generic[_S]):
    """The state a rule's own walk keeps, and how a block is allowed to change it."""

    # A copy of the state, to enter each block and each If arm from.
    snapshot: Callable[[], _S]
    # Puts back a copy taken by snapshot.
    restore: Callable[[_S], None]
    # Drops what is known about these tracked names.
    forget: Callable[[AbstractSet[str]], None]
    # The tracked names a statement mentions, which a block may change.
    touches: Callable[[LeafStatementNode], Iterable[str]]
    # Called as a block is entered, before its own lines run: a For Each header
    # reads its source here.
    enter: Callable[[BodyNode], None] | None = None
    # Called once a block is left, after what it touches is forgotten: a rule that
    # ran a loop pass by pass puts back what the loop leaves (XLIDE issue #350).
    exit: Callable[[BodyNode], None] | None = None
    # Keeps what a With body leaves, rather than its entry state less what it
    # names. Only for a rule that reads `.Member` in the body as the subject's
    # (XLIDE issue #584).
    with_body_runs_through: bool = False


def walk_entering_blocks(
    source: str,
    body: Sequence[BodyNode],
    is_inactive: Callable[[BodyNode], bool],
    visit: Callable[[BodyNode], None],
    state: BlockEnteringState[_S],
    loop_touched: AbstractSet[str] = _NO_NAMES,
) -> None:
    """A rule's own statement walk, entering blocks (XLIDE issue #237). A statement
    inside a block is visited with the state the block is entered with, once the
    block's own lines have run: each If arm and each Case from that state, a
    With's body once, and a loop's body as its first pass runs it. A block or a
    single-line If nested in a loop may run on a later pass, so it first forgets
    every name the loop changes (XLIDE issue #238). After a block the state is its
    entry state less every name it touches, so a block that never names a
    variable keeps what is known about it."""
    _drive(_walk_entering_blocks(source, body, is_inactive, visit, state, loop_touched))


def _walk_entering_blocks(
    source: str,
    body: Sequence[BodyNode],
    is_inactive: Callable[[BodyNode], bool],
    visit: Callable[[BodyNode], None],
    state: BlockEnteringState[_S],
    loop_touched: AbstractSet[str],
) -> _Task:
    for node in body:
        if is_inactive(node):
            continue
        child = getattr(node, "body", None)
        if not isinstance(child, list):
            # A single-line If runs its statements on some passes only.
            if _is_conditional_leaf(node):
                state.forget(loop_touched)
            visit(node)
            continue
        touched = _block_touches(source, node, is_inactive, state.touches)
        # An enclosing loop may have changed these on an earlier pass, and the
        # block's own lines run before its body.
        state.forget(loop_touched)
        if state.enter is not None:
            state.enter(node)
        state.forget(_header_touches(source, node, state.touches))
        entry = state.snapshot()
        if isinstance(node, IfBlockNode):
            for branch in node.branches:
                state.restore(entry)
                yield _walk_entering_blocks(
                    source, branch.body, is_inactive, visit, state, loop_touched
                )
        elif isinstance(node, SelectBlockNode):
            for arm in select_arms(source, node.body):
                state.restore(entry)
                yield _walk_entering_blocks(source, arm, is_inactive, visit, state, loop_touched)
        elif isinstance(node, WithBlockNode) and state.with_body_runs_through:
            # A With body runs once, in order: what holds at its end holds after
            # it. `With c` then `.Add 10` leaves c with one element (XLIDE #584).
            yield _walk_entering_blocks(source, child, is_inactive, visit, state, loop_touched)
            if state.exit is not None:
                state.exit(node)
            continue
        else:
            yield _walk_entering_blocks(
                source,
                child,
                is_inactive,
                visit,
                state,
                loop_touched | touched if is_loop_block(node) else loop_touched,
            )
        state.restore(entry)
        state.forget(touched)
        if state.exit is not None:
            state.exit(node)


def tracked_locals_named_whole(
    toks: Sequence[VbaToken],
    span_start: int,
    is_tracked: Callable[[str], bool],
    read_only_intrinsics: AbstractSet[str],
    arrays: AbstractSet[str] = _NO_NAMES,
    keeps: CalleeKeeps | None = None,
) -> dict[str, int]:
    """Every tracked local the statement names whole, bare rather than as a member
    access or indexed, in an argument position: a call statement's argument, an
    argument to a function inside an expression, or an argument to a qualified
    member call.

    VBA passes by reference by default, so the callee may have assigned or
    allocated the caller's variable, and its state is unknown from that point on
    (XLIDE issue #70). A mention the callee provably only reads is left out: the
    operand of `Is`, the argument of an intrinsic in `read_only_intrinsics`, a
    subscript of one of `arrays`, a part of a larger expression, and an argument
    `keeps` says the module's own procedure cannot change. Each name maps to its
    first such mention's absolute offset, so a rule can tell an access before the
    pass from one after it within the same statement. `toks` are the statement's
    significant tokens after any leading label, with offsets relative to
    `span_start`.
    """
    out: dict[str, int] = {}
    if len(toks) < 2:
        return out
    # A one-line If: its condition is an expression, and each arm after Then or
    # Else a statement of its own (XLIDE issue #575).
    then_at = _single_line_if_then(toks)
    if then_at is not None:

        def merge(part: Sequence[VbaToken]) -> None:
            for lower, at in tracked_locals_named_whole(
                part, span_start, is_tracked, read_only_intrinsics, arrays
            ).items():
                if lower not in out or out[lower] > at:
                    out[lower] = at

        merge(toks[:then_at])
        depth = 0
        arm_start = then_at + 1
        for i in range(arm_start, len(toks) + 1):
            raw = toks[i].raw_text if i < len(toks) else None
            if raw in ("(", "["):
                depth += 1
            elif raw in (")", "]"):
                depth -= 1
            elif i == len(toks) or (depth == 0 and token_word(toks[i]) == "else"):
                merge(toks[arm_start:i])
                arm_start = i + 1
        return out
    # A bare mention at the top level is an argument only in a call statement:
    # `Foo x`, `Call Foo(x)`, `obj.Method x`, or `.Method x` inside With. In
    # `Set a = b`, `Dim a As T`, or `If a Is Nothing` it is not. Nor in a control
    # statement's own words: `If a > 1 Then c.Add 5` passes no a, and its Then arm
    # is a statement of its own (XLIDE issue #575).
    head = 1 if token_word(toks[0]) == "call" else 0
    head_tok = _at(toks, head)
    is_call_statement = (
        (token_name(head_tok) is not None or (head_tok is not None and head_tok.raw_text == "."))
        and not (head == 0 and token_word(toks[0]) in _CONTROL_WORDS)
        and not _has_top_level_assignment(toks)
    )
    depth = 0
    # What each open parenthesis follows: a subscript of one of `arrays` passes
    # nothing (XLIDE issue #479: `a(i) = i` leaves i as it was).
    opened: list[str | None] = []
    opened_at: list[int] = []
    for i in range(1, len(toks)):
        raw = toks[i].raw_text
        if raw in ("(", "["):
            depth += 1
            before1 = _at(toks, i - 1)
            before2 = _at(toks, i - 2)
            if (before1 is not None and before1.raw_text == ".") or (
                before2 is not None and before2.raw_text == "."
            ):
                opened.append(None)
            else:
                callee_name = token_name(before1)
                opened.append(callee_name.lower() if callee_name is not None else None)
            opened_at.append(i)
            continue
        if raw in (")", "]"):
            depth -= 1
            if opened:
                opened.pop()
                opened_at.pop()
            continue
        name = token_name(toks[i])
        lower = name.lower() if name else None
        if not lower or not is_tracked(lower) or lower in out:
            continue
        if depth == 0 and not is_call_statement:
            continue
        enclosing = opened[-1] if opened else None
        if depth > 0 and enclosing is not None and enclosing in arrays:
            continue
        prev = toks[i - 1].raw_text
        nxt = _at(toks, i + 1)
        if prev in (".", "!") or (nxt is not None and nxt.raw_text in ("(", ".", "!")):
            continue
        if token_word(nxt) == "is":
            continue
        # Part of a larger expression, `a + 1`, is passed by value (XLIDE issue
        # #665, measured in Excel 16.0: `Cells(a + 1, 1)` leaves a as it was).
        next_word = token_word(nxt) or (nxt.raw_text if nxt is not None else "")
        prev_word = token_word(toks[i - 1]) or prev
        if next_word in _EXPRESSION_OPERATORS or (i > 1 and prev_word in _EXPRESSION_OPERATORS):
            continue
        # An argument of the module's own Function called inside an expression,
        # `b = Twice(a)`, as a call statement's is (XLIDE issue #665).
        if (
            keeps is not None
            and depth > 0
            and enclosing is not None
            and prev in ("(", ",", ":=")
            and opened_at
        ):
            open_ = opened_at[-1]
            index = 0
            inner = 0
            for k in range(open_ + 1, i):
                r = toks[k].raw_text
                inner += 1 if r == "(" else -1 if r == ")" else 0
                if inner == 0 and r == ",":
                    index += 1
            named = token_name(_at(toks, i - 2)) if prev == ":=" else None
            if keeps(enclosing, index, named):
                continue
        if prev == "(":
            callee = token_name(_at(toks, i - 2))
            if (callee.lower() if callee else "") in read_only_intrinsics:
                continue
        # `SetN (n)` and `F a, (n)`: parentheses of their own pass a copy (XLIDE
        # issue #449, measured in Excel 16.0).
        if (
            prev == "("
            and nxt is not None
            and nxt.raw_text == ")"
            and _grouping_paren(toks, i - 1, head, is_call_statement)
        ):
            continue
        if keeps is not None and is_call_statement and _callee_keeps(toks, i, head, keeps):
            continue
        out[lower] = span_start + toks[i].start
    return out


def _grouping_paren(
    toks: Sequence[VbaToken], open_: int, head: int, is_call_statement: bool
) -> bool:
    """Whether the parenthesis at `open_` groups an argument rather than opening a
    call's list: after a comma or another parenthesis, or after the callee of a
    call statement written without Call, where VBA reads `F (n)` and `F(n)` alike
    as F given (n)."""
    before = _at(toks, open_ - 1)
    if before is not None and before.raw_text in (",", "("):
        return True
    after = _at(toks, _matching_close(toks, open_) + 1)
    return (
        is_call_statement
        and head == 0
        and open_ == _callee_end(toks, 0) + 1
        and (after is None or after.raw_text == "," or after.kind is TokenKind.COMMENT)
    )


def _callee_end(toks: Sequence[VbaToken], head: int) -> int:
    """The index of the last name in a call statement's callee chain: `Foo`,
    `obj.Method`."""
    i = head
    while True:
        dot = _at(toks, i + 1)
        if dot is None or dot.raw_text != "." or token_name(_at(toks, i + 2)) is None:
            return i
        i += 2


def _callee_keeps(
    toks: Sequence[VbaToken], i: int, head: int, keeps: CalleeKeeps
) -> bool:
    """Whether the module's own procedure called by the statement keeps the
    argument at `i`: `Touch c`, `Call InitV(c)`, `Touch arg:=c` (XLIDE issue #449).
    Only a bare callee is asked, never a member of an object."""
    callee = token_name(_at(toks, head))
    after_callee = _at(toks, head + 1)
    if not callee or (after_callee is not None and after_callee.raw_text == "."):
        return False
    explicit = head == 1
    # The argument list: after the callee, or inside `Call Foo(...)`.
    from_ = head + 1
    to = len(toks)
    if explicit:
        if after_callee is None or after_callee.raw_text != "(":
            return False
        from_ = head + 2
        to = _matching_close(toks, head + 1)
    if i < from_ or i >= to:
        return False
    depth = 0
    index = 0
    for k in range(from_, i):
        raw = toks[k].raw_text
        depth += 1 if raw == "(" else -1 if raw == ")" else 0
        if raw == "," and depth == 0:
            index += 1
    if depth != 0:
        return False
    before = _at(toks, i - 1)
    named = token_name(_at(toks, i - 2)) if before is not None and before.raw_text == ":=" else None
    return keeps(callee, index, named)


def _matching_close(toks: Sequence[VbaToken], open_: int) -> int:
    depth = 0
    for k in range(open_, len(toks)):
        raw = toks[k].raw_text
        depth += 1 if raw == "(" else -1 if raw == ")" else 0
        if depth == 0:
            return k
    return len(toks)


def _single_line_if_then(toks: Sequence[VbaToken]) -> int | None:
    """The index of a one-line If's top-level Then, when an arm follows it."""
    head = token_word(toks[0])
    if head != "if" and head != "elseif":
        return None
    depth = 0
    for i in range(1, len(toks)):
        raw = toks[i].raw_text
        if raw in ("(", "["):
            depth += 1
        elif raw in (")", "]"):
            depth -= 1
        elif depth == 0 and token_word(toks[i]) == "then":
            return i if i + 1 < len(toks) and toks[i + 1].kind is not TokenKind.COMMENT else None
    return None


# Statement words that read a condition or a value, and open no call.
_CONTROL_WORDS = frozenset(("if", "elseif", "else", "while", "do", "loop", "select", "case"))

# Operators that make a name part of a larger expression, which passes by value.
_EXPRESSION_OPERATORS = frozenset(
    (
        "+", "-", "*", "/", "\\", "^", "&", "mod", "and", "or", "xor", "not", "eqv", "imp",
        "like", "=", "<>", "<", ">", "<=", ">=",
    )
)


def _has_top_level_assignment(toks: Sequence[VbaToken]) -> bool:
    """True when a top-level '=' makes the statement an assignment, not a call."""
    depth = 0
    for tok in toks:
        raw = tok.raw_text
        if raw in ("(", "["):
            depth += 1
        elif raw in (")", "]"):
            depth -= 1
        elif depth == 0 and tok.kind is TokenKind.OPERATOR and raw == "=":
            return True
    return False


def _at(tokens: Sequence[VbaToken], i: int) -> VbaToken | None:
    return tokens[i] if 0 <= i < len(tokens) else None


__all__ = [
    "BlockEnteringState",
    "DataflowHooks",
    "Lattice",
    "StraightLineDataflowHooks",
    "leaves_the_list",
    "tracked_locals_named_whole",
    "walk_branch_merged_body",
    "walk_entering_blocks",
    "walk_straight_line_body",
]
