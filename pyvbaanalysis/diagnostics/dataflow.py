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
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass

from ..lexer.token_helpers import token_name, token_word
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import BodyNode, IfBlockNode, IfBranchKind, LeafStatementNode, is_leaf_statement


@dataclass(frozen=True, slots=True)
class Lattice:
    """The rule's good/init/unknown labels, kept rule-agnostic for the merge."""

    init: str
    good: str
    unknown: str


@dataclass(frozen=True, slots=True)
class DataflowHooks:
    """Rule-specific hooks driving one straight-line dataflow walk.

    The branch-merge hooks (snapshot_state / restore_state / set_state / lattice)
    are optional: without them walk_branch_merged_body treats an If block
    conservatively, exactly like walk_straight_line_body.
    """

    # Applies one straight-line statement's transitions and diagnostics.
    on_statement: Callable[[LeafStatementNode], None]
    # Lowercased tracked names one statement touches (for nested-block demotion).
    touches_in_statement: Callable[[LeafStatementNode], Iterable[str]]
    # Demotes one tracked name to the rule's "unknown" state.
    demote_to_unknown: Callable[[str], None]
    # Inspects one non-statement node before its body's touch demotion.
    on_block: Callable[[BodyNode], None] | None = None
    # Snapshot every tracked name's current state, for forking If arms.
    snapshot_state: Callable[[], dict[str, str]] | None = None
    # Overwrite the live state from a snapshot, restoring it before the next arm.
    restore_state: Callable[[Mapping[str, str]], None] | None = None
    # Write one tracked name's merged post-block state.
    set_state: Callable[[str, str], None] | None = None
    lattice: Lattice | None = None


def walk_straight_line_body(
    body: Sequence[BodyNode],
    is_inactive: Callable[[BodyNode], bool],
    hooks: DataflowHooks,
) -> None:
    """Walk the straight-line statements of a procedure body.

    Plain statements run the rule's transitions in order; nested blocks
    (If/For/Do/...) are not entered. Every tracked name touched anywhere inside a
    block is demoted to "unknown" rather than guessing which runtime path runs.
    """
    for node in body:
        if is_inactive(node):
            continue
        if is_leaf_statement(node):
            hooks.on_statement(node)
            continue
        if hooks.on_block is not None:
            hooks.on_block(node)
        child = getattr(node, "body", None)
        if isinstance(child, list):
            for lower in _collect_nested_touches(child, is_inactive, hooks):
                hooks.demote_to_unknown(lower)


def walk_branch_merged_body(
    body: Sequence[BodyNode],
    is_inactive: Callable[[BodyNode], bool],
    hooks: DataflowHooks,
) -> None:
    """Like walk_straight_line_body, but intersects an If block's per-arm state.

    A tracked name advances to its "good" state after the If only when it reaches
    "good" on EVERY arm AND a syntactic else arm is present, otherwise it follows
    the conservative demotion. Names a balanced If never touches keep their entry
    state (the precision win). For/Do/While/With/Select stay conservative. Only
    sound for procedures WITHOUT unstructured control flow; callers gate on
    procedure_has_unstructured_flow and fall back to walk_straight_line_body.
    """
    for node in body:
        if is_inactive(node):
            continue
        if is_leaf_statement(node):
            hooks.on_statement(node)
            continue
        if hooks.on_block is not None:
            hooks.on_block(node)
        if (
            isinstance(node, IfBlockNode)
            and hooks.snapshot_state is not None
            and hooks.restore_state is not None
            and hooks.set_state is not None
            and hooks.lattice is not None
        ):
            _merge_if_block(node, is_inactive, hooks)
            continue
        child = getattr(node, "body", None)
        if isinstance(child, list):
            for lower in _collect_nested_touches(child, is_inactive, hooks):
                hooks.demote_to_unknown(lower)


def _merge_if_block(
    if_block: IfBlockNode,
    is_inactive: Callable[[BodyNode], bool],
    hooks: DataflowHooks,
) -> None:
    """Intersect the per-arm state of one If block (see walk_branch_merged_body)."""
    snapshot = hooks.snapshot_state
    restore = hooks.restore_state
    set_state = hooks.set_state
    lattice = hooks.lattice
    assert snapshot is not None and restore is not None and set_state is not None and lattice is not None

    touched = _collect_nested_touches(if_block.body, is_inactive, hooks)
    has_else = any(branch.branch_kind is IfBranchKind.ELSE for branch in if_block.branches)
    if not has_else:
        # No else arm: the empty fall-through path keeps the entry state, so a
        # name can only stay "good" if it was already "good". Reproduce the
        # conservative behavior by demoting every touched name.
        for lower in touched:
            hooks.demote_to_unknown(lower)
        return
    entry = snapshot()
    arm_states: list[dict[str, str]] = []
    for branch in if_block.branches:
        restore(entry)
        walk_branch_merged_body(branch.body, is_inactive, hooks)
        arm_states.append(snapshot())
    restore(entry)
    for lower in touched:
        fallback = entry.get(lower, lattice.unknown)
        set_state(lower, _join_branch_states(arm_states, lower, fallback, lattice))


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


def _collect_nested_touches(
    body: Sequence[BodyNode],
    is_inactive: Callable[[BodyNode], bool],
    hooks: DataflowHooks,
) -> set[str]:
    """Recursively collect tracked names touched anywhere inside nested bodies."""
    out: set[str] = set()
    for node in body:
        if is_inactive(node):
            continue
        if is_leaf_statement(node):
            out.update(hooks.touches_in_statement(node))
            continue
        child = getattr(node, "body", None)
        if isinstance(child, list):
            out.update(_collect_nested_touches(child, is_inactive, hooks))
    return out


def tracked_locals_passed_as_call_arguments(
    toks: Sequence[VbaToken],
    is_tracked: Callable[[str], bool],
) -> set[str]:
    """Lowercased tracked locals passed as bare arguments of a call statement.

    Covers `Helper x` and `Call Helper(x)`, where ByRef passing may rebind them.
    `toks` are the statement's significant tokens after any leading label.
    """
    if len(toks) < 2 or _has_top_level_assignment(toks):
        return set()
    start = 1 if token_word(toks[0]) == "call" else 0
    after = _at(toks, start + 1)
    if token_name(_at(toks, start)) is None or (after is not None and after.raw_text == "."):
        return set()
    out: set[str] = set()
    for i in range(start + 1, len(toks)):
        prev = _at(toks, i - 1)
        nxt = _at(toks, i + 1)
        if (prev is not None and prev.raw_text == ".") or (nxt is not None and nxt.raw_text == "."):
            continue
        name = token_name(toks[i])
        if name is not None and is_tracked(name.lower()):
            out.add(name.lower())
    return out


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
