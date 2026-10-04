"""Ported from xlide_vscode/src/analyzer/diagnostics/straightLineValues.ts.

The assignment each statement of a procedure sees for a local (XLIDE issue #180).

The value rules (division-by-zero, string-arithmetic-coercion,
variant-value-misuse, array-subscript-out-of-bounds, runtime-argument-value)
read a local's value from the procedure as a whole: a local every assignment
gives the same literal. A second assignment anywhere turned them off, even one
after the line that fails: `d = 0: x = 10 / d: d = 2`. This walk follows each
statement list in order and records, for every leaf statement, the last
`x = value` that reaches it with nothing between able to change x. The rules
read that first and fall back to the procedure-wide value.

What ends a value: another assignment (the new one replaces it), passing the
name whole to a call (ByRef), a statement that writes it another way (Set,
ReDim, Erase, Input #, Get #, Line Input #, LSet, RSet, Mid =), and any block
that touches the name, since a loop or a branch may or may not have run it. A
label ends every value, because a GoTo may arrive there from anywhere, and so
does a GoSub, which may run any statement of the procedure. A block keeps the
values it never touches, inside and after it.

Port notes. Parser nodes are not hashable, so what upstream keys by node (the
reaching assignments, the statements that never run) is keyed here by
`id(node)`; the walk cache keeps the body, and so its nodes, alive. Upstream's
walk recurses once per nested block; here each statement list and block is a
generator, run on an explicit stack by `_run_walk`, so nesting is not bounded
by Python's recursion limit.
"""

from __future__ import annotations

import math
import re
import struct
from collections.abc import Callable, Generator, Iterable, Mapping, Sequence
from dataclasses import dataclass
from types import MappingProxyType
from typing import TYPE_CHECKING, Any, Generic, Literal, NamedTuple, TypeGuard, TypeVar, Union

from ..conditional import ConditionalActivityTracker
from ..constants.date_literal import date_literal_serial
from ..constants.integer_constant_expression import (
    bankers_round,
    evaluate_integer_constant_expression,
    parse_vba_integer_literal,
)
from ..flow.procedure_labels import (
    jump_target_label_declaration,
    statement_label_declaration,
    statement_label_references,
)
from ..js_compat import JS_WHITESPACE, js_number, js_number_to_string
from ..lexer.token_helpers import match_paren_from, split_top_level_token_groups, token_name
from ..lexer.token_helpers import token_word as token_text
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import (
    BodyNode,
    ConditionalDirectiveNode,
    DoBlockNode,
    ForBlockNode,
    IfBlockNode,
    IfBranchKind,
    LeafStatementNode,
    SelectBlockNode,
    Span,
    StatementNode,
    VariableGroupNode,
    WhileBlockNode,
    WithBlockNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ..runtime.vba_runtime import resolve_runtime_function
from .condition_value import ConditionFacts, condition_value, if_condition_tokens, number_value
from .reaching_snapshot import with_reaching_value
from .walker import (
    bare_assignment_target,
    block_footer_line_span,
    block_header_line_span,
    is_inactive_node,
    raw_expression_tokens,
    statement_and_branch_spans,
    statement_tokens,
    statement_tokens_after_leading_label,
)

if TYPE_CHECKING:
    from .known_string_calls import ModuleCompare

# The value tokens of each local's reaching assignment, by lowercased name, and
# of a local array's element by `element_key`: `a(0) = Null` is under "a(0)".
ReachingAssignments = Mapping[str, Sequence[VbaToken]]

# Never changed: an empty state every walk may share.
_NONE: ReachingAssignments = MappingProxyType({})


def element_key(lower: str, index: float) -> str:
    """The key a local array's element is held under: "a(0)"."""
    return f"{lower}({js_number_to_string(index)})"


# What the walk knows an object local holds, by identity: Nothing, from
# `Set c = Nothing` or a local never set, and a Collection nothing has added to
# yet, from `Set c = New Collection` or `Dim c As New Collection`. A loop over
# the one, or until the other is Nothing, runs no pass (issue #483). Any other
# mention of the name ends it.
OBJECT_NOTHING: Sequence[VbaToken] = tuple(raw_expression_tokens("Nothing"))
EMPTY_COLLECTION: Sequence[VbaToken] = tuple(raw_expression_tokens("New Collection"))
# What a Variant local holds before anything assigns it (issue #691).
VARIANT_EMPTY: Sequence[VbaToken] = tuple(raw_expression_tokens("Empty"))

# Statement heads that write every name they mention.
_WRITING_HEADS: frozenset[str] = frozenset(
    {"set", "redim", "erase", "input", "get", "line", "lset", "rset", "mid", "mid$"}
)

# VBA functions that only read an argument named whole.
_READ_ONLY_INTRINSICS: frozenset[str] = frozenset(
    {"lbound", "ubound", "isarray", "len", "lenb", "isempty", "isnull", "isnumeric", "typename", "vartype"}
)

# What the calls in a statement leave in the names they pass ByRef, by lowercased name.
CallEffects = Callable[[Sequence[VbaToken]], Mapping[str, Sequence[VbaToken]]]

_K = TypeVar("_K")
_V = TypeVar("_V")


class _IdentityTable(Generic[_K, _V]):
    """Values by the identity of a key object, where upstream keeps a WeakMap.

    A state or a body is neither hashable as itself nor weak-referenceable. Each
    entry keeps its key alive, so an id() is never reused while it stands, and
    past `capacity` the oldest entry goes, as identity_cache.IdentityLru bounds
    its entries; a lookup here is one dict probe, where IdentityLru scans."""

    __slots__ = ("_capacity", "_entries")

    def __init__(self, capacity: int) -> None:
        self._capacity = capacity
        self._entries: dict[int, tuple[_K, _V]] = {}

    def get(self, key: _K) -> _V | None:
        entry = self._entries.get(id(key))
        return entry[1] if entry is not None and entry[0] is key else None

    def set(self, key: _K, value: _V) -> None:
        self._entries.pop(id(key), None)
        self._entries[id(key)] = (key, value)
        if len(self._entries) > self._capacity:
            del self._entries[next(iter(self._entries))]


_CALL_EFFECTS: _IdentityTable[ReachingAssignments, CallEffects] = _IdentityTable(1024)


def set_call_effects(start: ReachingAssignments, effects: CallEffects) -> None:
    """Has walks from `start` apply `effects` at each statement (issue #449):
    `ZeroN n`, where ZeroN sets its ByRef parameter to 0, leaves n 0. A start is
    one procedure's, so its walks share the effects and the cache."""
    _CALL_EFFECTS.set(start, effects)


# The effects of the walk running now; set while a walk runs.
_walk_call_effects: CallEffects | None = None


@dataclass(frozen=True, slots=True)
class DeclaredFacts:
    """What a procedure's declarations say of its locals, for the guards (issue #691)."""

    # The declared type, lowercased: "long", "variant", "long()" for an array.
    type: Callable[[str], str | None]
    # A fixed one-dimension array's bounds.
    bounds: Callable[[str], tuple[float, float] | None]
    # The value of a Const or Enum member of the module: `mB` (issue #691).
    constant: Callable[[str], float | None]


_DECLARED_FACTS: _IdentityTable[ReachingAssignments, DeclaredFacts] = _IdentityTable(1024)


def set_declared_facts(start: ReachingAssignments, facts: DeclaredFacts) -> None:
    """Has walks from `start` know what the declarations say."""
    _DECLARED_FACTS.set(start, facts)


# The declarations of the walk running now; set while a walk runs.
_walk_declared: DeclaredFacts | None = None


class StraightLineExit(NamedTuple):
    """What holds as a body runs off its end (None when no path does), with the
    ids of the statements and the one-line If branches the walk found never run."""

    exit: ReachingAssignments | None
    dead: set[int]
    dead_spans: list[Span]


def straight_line_assignments(
    source: str,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    initial: ReachingAssignments = _NONE,
) -> dict[int, ReachingAssignments]:
    """For each statement of the body that some straight-line assignment reaches,
    by id(statement), the reaching assignments; for a block, those that reach its
    header. A statement not in the map has none. `initial` is what holds as the
    procedure starts: each local's declared default, so `x = 1 / x` reads the 0 x
    starts with (issue #259)."""
    return _cached_walk(source, body, activity, initial).result


def straight_line_dead_branches(
    source: str,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    initial: ReachingAssignments = _NONE,
) -> list[Span]:
    """The spans of the one-line If branches that never run, because the walk
    knows the condition (issue #430): the Else of `If x = 2 Then ... Else ...`
    with x still 2. The If itself runs, so it is not in the unreachable set."""
    return _cached_walk(source, body, activity, initial).dead_spans


def straight_line_exit(
    source: str,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    initial: ReachingAssignments,
) -> StraightLineExit:
    """What holds as the body runs off its end, from `initial`, with the
    statements and one-line If branches the walk found never run. The end state
    is None when no path reaches it (issue #562)."""
    walk = _cached_walk(source, body, activity, initial)
    return StraightLineExit(walk.exit, walk.dead, walk.dead_spans)


def straight_line_unreachable(
    source: str,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    initial: ReachingAssignments = _NONE,
) -> set[int]:
    """The ids of the statements of the body that never run, because a guard whose
    value the walk knows decides against them (issue #273): `n = 0: If n > 0 Then
    ...`, the arms after `Case 0` with the selector 0, and everything after
    `If d = 0 Then Exit Function` with d still 0. A label ends it, since a GoTo
    may arrive there. A block's own statements are in it when the whole block
    never runs."""
    return _cached_walk(source, body, activity, initial).dead


@dataclass(slots=True)
class _CachedWalk:
    source: str
    activity: ConditionalActivityTracker | None
    result: dict[int, ReachingAssignments]
    dead: set[int]
    dead_spans: list[Span]
    # What holds as the body runs off its end; None when no path does.
    exit: ReachingAssignments | None


@dataclass(slots=True)
class _WalkOut:
    """What one walk collects: each statement's reaching values, and the statements that never run."""

    out: dict[int, ReachingAssignments]
    dead: set[int]
    # The one-line If branches a known condition decides against.
    dead_spans: list[Span]
    # Whether Err.Raise leaves the list: not when the procedure resumes past errors.
    raise_leaves: bool
    # How many times a GoTo, GoSub, Resume or On ... GoTo names each label, by key.
    referenced: dict[str, int]


# The end of a statement list no path reaches.
_UNREACHED: ReachingAssignments = MappingProxyType({})

_WALKS: _IdentityTable[Sequence[BodyNode], dict[str, _CachedWalk]] = _IdentityTable(512)

# Each start's cache key, by identity: a kept start is asked for by several rules.
_START_KEYS: _IdentityTable[ReachingAssignments, str] = _IdentityTable(1024)


def _js_ws_class() -> str:
    return "[" + "".join(re.escape(ch) for ch in JS_WHITESPACE) + "]"


# `/\)\s*=\s*null\b/i` and `/\bon\s+error\s+resume\s+next\b/i`, with
# JavaScript's `\s` and ASCII word characters.
_ELEMENT_NULL_RE = re.compile(
    r"\)" + _js_ws_class() + "*=" + _js_ws_class() + r"*null\b", re.IGNORECASE | re.ASCII
)
_RESUME_NEXT_RE = re.compile(
    r"\bon" + _js_ws_class() + "+error" + _js_ws_class() + "+resume" + _js_ws_class() + r"+next\b",
    re.IGNORECASE | re.ASCII,
)


def _cached_walk(
    source: str,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    initial: ReachingAssignments,
) -> _CachedWalk:
    global _walk_arrays, _walk_procedures, _walk_elements, _walk_call_effects
    global _walk_declared, _walk_collections
    # Six rules ask for the same procedure in one pass; a parse makes a new
    # body, so the body is the key, with what holds at the start.
    key = _START_KEYS.get(initial)
    if key is None:
        key = "\n".join(
            sorted(f"{name}={' '.join(tok.raw_text for tok in value)}" for name, value in initial.items())
        )
        _START_KEYS.set(initial, key)
    by_start = _WALKS.get(body)
    if by_start is None:
        by_start = {}
        _WALKS.set(body, by_start)
    cached = by_start.get(key)
    if cached is not None and cached.source == source and cached.activity is activity:
        return cached
    out: dict[int, ReachingAssignments] = {}
    dead: set[int] = set()
    # Under On Error Resume Next, Err.Raise goes on to the next line.
    text = source[body[0].span.start : body[-1].span.end] if len(body) > 0 else ""
    dead_spans: list[Span] = []
    # The walk is synchronous, so its arrays can sit beside it for _passed_whole.
    outer = _walk_arrays
    outer_procedures = _walk_procedures
    outer_elements = _walk_elements
    outer_effects = _walk_call_effects
    outer_declared = _walk_declared
    outer_collections = _walk_collections
    _walk_collections = _local_collection_names(body, activity)
    _walk_arrays = _local_array_names(body, activity)
    _walk_call_effects = _CALL_EFFECTS.get(initial)
    _walk_declared = _DECLARED_FACTS.get(initial)
    _walk_procedures = _module_procedure_names(source)
    _walk_elements = _ELEMENT_NULL_RE.search(text) is not None
    try:
        walk = _WalkOut(
            out,
            dead,
            dead_spans,
            _RESUME_NEXT_RE.search(text) is None,
            _referenced_labels(source, body, activity),
        )
        exit_state = _run_walk(_walk_list(source, body, initial, activity, walk))
    finally:
        _walk_arrays = outer
        _walk_procedures = outer_procedures
        _walk_elements = outer_elements
        _walk_call_effects = outer_effects
        _walk_declared = outer_declared
        _walk_collections = outer_collections
    result = _CachedWalk(source, activity, out, dead, dead_spans, None if exit_state is _UNREACHED else exit_state)
    by_start[key] = result
    return result


# A step of the walk: a generator that yields the walk of a nested statement list
# or block, is sent what holds after it, and returns what holds after its own.
_Walk = Generator[Any, Any, ReachingAssignments]


def _run_walk(root: _Walk) -> ReachingAssignments:
    """Runs a walk and every walk it yields on an explicit stack."""
    stack: list[_Walk] = [root]
    sent: ReachingAssignments | None = None
    error: BaseException | None = None
    while True:
        top = stack[-1]
        try:
            if error is not None:
                child = top.throw(error)
            else:
                child = top.send(sent)
        except StopIteration as stop:
            stack.pop()
            if not stack:
                done: ReachingAssignments = stop.value
                return done
            sent, error = stop.value, None
            continue
        except BaseException as exc:
            stack.pop()
            if not stack:
                raise
            sent, error = None, exc
            continue
        error = None
        sent = None
        stack.append(child)


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw(tok: VbaToken | None) -> str | None:
    return tok.raw_text if tok is not None else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name is not None else None


def _body_of(node: BodyNode) -> list[BodyNode] | None:
    """`'body' in node && Array.isArray(node.body)`: a block's statements."""
    body = getattr(node, "body", None)
    return body if isinstance(body, list) else None


def _single_line_if_branches(node: BodyNode) -> list[Span] | None:
    return node.single_line_if_branches if isinstance(node, StatementNode) else None


def _walk_list(
    source: str,
    statements: Sequence[BodyNode],
    entry: ReachingAssignments,
    activity: ConditionalActivityTracker | None,
    walk: _WalkOut,
    case_resets: bool = False,
) -> _Walk:
    """Walks one statement list from `entry` and returns what holds after it. The
    state is replaced, never changed in place, so every statement can keep the one
    it saw."""
    current = entry
    # What holds at each forward `GoTo L` of this list, plain or the whole
    # branch of a one-line If, by the label's key (issue #614).
    jumps: dict[str, list[ReachingAssignments]] = {}
    jump_counts: dict[str, int] = {}
    i = 0
    while i < len(statements):
        node = statements[i]
        if is_inactive_node(activity, node) or isinstance(node, (VariableGroupNode, ConditionalDirectiveNode)):
            i += 1
            continue
        # A label a jump may reach starts over; one nothing names, as when
        # `On Error GoTo EH` is commented out, leaves dead code dead (issue
        # #421), and live code as it was: `a = 1` then `L1:` keeps a 1
        # (issue #665, measured in Excel 16.0). One only forward GoTos of
        # this list reach holds what every way in agrees on (issue #614).
        label = statement_label_declaration(source, node.span) if is_leaf_statement(node) else None
        if label is not None and label.key in walk.referenced:
            ways = jumps.get(label.key, [])
            # Every GoTo to it seen, those that never jump included (issue #673).
            seen = jump_counts.get(label.key, 0)
            if seen > 0 and seen == walk.referenced.get(label.key):
                states = ways if current is _UNREACHED else [current, *ways]
                current = _UNREACHED if len(states) == 0 else _agreed(states)
            else:
                current = _NONE
        if current is not _UNREACHED and is_leaf_statement(node):
            target = _forward_goto(source, node, statements[i + 1] if i + 1 < len(statements) else None)
            if target is not None:
                # `If False Then GoTo L` never jumps, and adds no way into L
                # (issue #673, measured in Excel 16.0).
                condition = (
                    if_condition_tokens(statement_tokens_after_leading_label(source, node.span))
                    if _single_line_if_branches(node) is not None
                    else None
                )
                may_jump = condition_value(condition, _facts_from(current, source)) if condition is not None else True
                jump_counts[target] = jump_counts.get(target, 0) + 1
                if may_jump is not False:
                    jumps[target] = [*jumps.get(target, []), current]
        if current is _UNREACHED:
            # After a guard that always leaves: nothing here runs (issue #273).
            _mark_unreachable(node, walk.dead)
            i += 1
            continue
        if not is_leaf_statement(node):
            # What holds as the block starts: a For reads its bounds here
            # (issue #200).
            _record(walk.out, node, current)
            current = yield _walk_block(source, node, current, activity, walk)
            i += 1
            continue
        if case_resets and token_text(_at(statement_tokens_after_leading_label(source, node.span), 0)) == "case":
            # A Select's arms are exclusive: each starts where the block did.
            current = entry
        if _single_line_if_branches(node) is not None:
            group: list[LeafStatementNode] = [node]
            while i + 1 < len(statements):
                tail = statements[i + 1]
                if not (is_leaf_statement(tail) and tail.single_line_if_tail):
                    break
                group.append(tail)
                i += 1
            current = _walk_single_line_if(source, group, current, walk)
            i += 1
            continue
        _record(walk.out, node, current)
        current = _after_statement(source, node.span, current)
        if _leaves_the_list(source, node.span, walk.raise_leaves):
            current = _UNREACHED
        i += 1
    return current


def _leaves_the_list(source: str, span: Span, raise_leaves: bool) -> bool:
    from .dataflow import leaves_the_list

    return leaves_the_list(source, span, raise_leaves)


def _walk_single_line_if(
    source: str,
    group: Sequence[LeafStatementNode],
    current: ReachingAssignments,
    walk: _WalkOut,
) -> ReachingAssignments:
    """A single-line If and the statements after its colons run only with their
    branch. With the condition known (issue #273), a false one runs nothing, and a
    true one runs its branch in order, which may leave. Otherwise they see what
    held before the If, less whatever the If itself changes, and so does what
    follows."""
    if_stmt = group[0]
    branches = _single_line_if_branches(if_stmt) or []
    condition = (
        if_condition_tokens(statement_tokens_after_leading_label(source, if_stmt.span))
        if len(branches) == 1
        else None
    )
    known = condition_value(condition, _facts_from(current, source)) if condition is not None else None
    if known is False:
        for stmt in group:
            walk.dead.add(id(stmt))
        return current
    if known is True:
        state = current
        for k, stmt in enumerate(group):
            _record(walk.out, stmt, state)
            span = branches[0] if k == 0 else stmt.span
            state = _after_statement(source, span, state)
            if _leaves_the_list(source, span, walk.raise_leaves):
                return _UNREACHED
        return state
    # `If x = 2 Then y = 0 Else y = Sqr(-1)` with x known: the If runs, and
    # the branch its condition decides against does not (issue #430).
    if len(branches) == 2:
        decided = condition_value(
            if_condition_tokens(statement_tokens_after_leading_label(source, if_stmt.span)) or [],
            _facts_from(current, source),
        )
        if decided is not None:
            else_start = branches[1].start
            walk.dead_spans.append(branches[1 if decided else 0])
            for tail in group[1:]:
                if (tail.span.start >= else_start) == decided:
                    walk.dead_spans.append(tail.span)
            # One statement an arm, and no If nested in them: the arm that
            # runs is the If's whole effect, `If n > 0 Then F = 1 Else F = 0`
            # with n known leaving F known (issue #562).
            # The Then arm's span runs through the Else that ends it.
            arm_toks = statement_tokens(source, branches[0 if decided else 1])
            else_tok = next((tok for tok in arm_toks if token_text(tok) == "else"), None) if decided else None
            taken = (
                Span(branches[0].start, branches[0].start + else_tok.start)
                if else_tok is not None
                else branches[0 if decided else 1]
            )
            nested = any(
                k > 0 and token_text(tok) == "if"
                for k, tok in enumerate(statement_tokens_after_leading_label(source, if_stmt.span))
            )
            if len(group) == 1 and not nested:
                _record(walk.out, if_stmt, current)
                if _leaves_the_list(source, taken, walk.raise_leaves):
                    return _UNREACHED
                return _after_statement(source, taken, current)
    touched = _touched_by(source, group)
    after = _without_mentioned_objects(
        _NONE if touched == "all" else _without(current, touched),
        source,
        Span(group[0].span.start, group[-1].span.end),
    )
    for stmt in group:
        _record(walk.out, stmt, after)
    return after


def _walk_block(
    source: str,
    node: BodyNode,
    entry: ReachingAssignments,
    activity: ConditionalActivityTracker | None,
    walk: _WalkOut,
) -> _Walk:
    global _walk_with_collection
    from .block_headers import is_loop_block

    body = _body_of(node)
    if body is None:
        return entry
    # An If or a Select whose outcome is known runs one arm (issue #273).
    chosen = _known_arm(source, node, entry)
    if chosen is not None:
        arms, taken = chosen
        for arm in arms:
            if arm is not taken:
                for stmt in arm:
                    _mark_unreachable(stmt, walk.dead)
        if taken is None:
            return entry
        after_arm: ReachingAssignments = yield _walk_list(source, taken, entry, activity, walk)
        return after_arm
    # A loop known to run no pass runs none of its body (issue #406): `For i
    # = 1 To 0`, `While d <> 0` with d still 0, `For Each x In Array()`. A
    # For counter is left at its start.
    none = _loop_runs_no_pass(source, node, entry)
    if none is not None:
        for stmt in body:
            _mark_unreachable(stmt, walk.dead)
        if none.counter is None:
            return entry
        return with_reaching_value(
            entry, none.counter.name, raw_expression_tokens(js_number_to_string(none.counter.value))
        )
    # `With k`, k a Collection: `.Add a` inside reads a (issue #665).
    outer_with = _walk_with_collection
    if isinstance(node, WithBlockNode):
        header = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
        _walk_with_collection = len(header) == 2 and (_lower_name(header[1]) or "") in _walk_collections
    try:
        # The rest of upstream's walkBlock, its walkBlockBody, with a With
        # block's subject set.
        touched = _loop_touched(source, node, entry, activity)
        after = _without_mentioned_objects(
            _NONE if touched == "all" else _without(entry, touched), source, node.span
        )
        # An If arm, a Case or a With body runs once, from the state the block
        # is entered with; a loop's body may run again with what it changed
        # (issue #259: `If True Then x = 1 / x` reads the 0 x starts with).
        inside = after if is_loop_block(node) or touched == "all" else entry
        if isinstance(node, IfBlockNode):
            for branch in node.branches:
                yield _walk_list(source, branch.body, inside, activity, walk)
        else:
            yield _walk_list(source, body, inside, activity, walk, isinstance(node, SelectBlockNode))
        final = None
        if touched != "all":
            final = _for_counter_final_value(source, node, activity)
            if final is None:
                final = _do_counter_final_value(source, node, entry, activity)
        if final is not None:
            return with_reaching_value(after, final.name, raw_expression_tokens(js_number_to_string(final.value)))
        return after
    finally:
        _walk_with_collection = outer_with


class _Counter(NamedTuple):
    name: str
    value: float


# The most passes the walk runs a Do loop's counter through before giving up.
_DO_COUNTER_PASSES = 100_000


def _is_number(value: object) -> TypeGuard[float]:
    """`typeof value === 'number'`."""
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _do_counter_final_value(
    source: str,
    node: BodyNode,
    entry: ReachingAssignments,
    activity: ConditionalActivityTracker | None,
) -> _Counter | None:
    """What a Do or While loop's counter holds after the loop ends (issue #479,
    measured in Excel 16.0): `i = 1: Do While i <= 5: ...: i = i + 1: Loop`
    leaves i at 6. The body steps the counter by one top-level `i = i + k` or
    `i = i - k` and writes it nowhere else, nothing in it may leave the loop, and
    the condition reads only the counter and names the body leaves alone."""
    if not isinstance(node, (DoBlockNode, WhileBlockNode)):
        return None
    is_do = isinstance(node, DoBlockNode)
    body = [stmt for stmt in node.body if not is_inactive_node(activity, stmt)]
    # The condition: on the header, or on the footer of a Do.
    header = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
    footer = (
        statement_tokens_after_leading_label(source, block_footer_line_span(source, node.span)) if is_do else []
    )
    head_word = token_text(_at(header, 1 if is_do else 0))
    foot_word = token_text(_at(footer, 1))
    at_head = head_word == "while" or head_word == "until"
    at_foot = foot_word == "while" or foot_word == "until"
    if at_head == at_foot:
        return None
    condition = header[2 if is_do else 1 :] if at_head else footer[2:]
    until = (head_word if at_head else foot_word) == "until"
    # The one step: `i = i + k` or `i = i - k`, with k a whole number.
    counter: str | None = None
    step: int | None = None
    step_at: BodyNode | None = None
    for stmt in body:
        bare = (
            bare_assignment_target(source, stmt.span)
            if is_leaf_statement(stmt) and _single_line_if_branches(stmt) is None
            else None
        )
        value = [tok for tok in bare[2] if tok.kind is not TokenKind.COMMENT] if bare is not None else []
        lower = bare[0].lower() if bare is not None else None
        k = (
            _signed_integer([value[2]])
            if len(value) == 3
            and _lower_name(value[0]) == lower
            and (value[1].raw_text == "+" or value[1].raw_text == "-")
            else None
        )
        if lower and k is not None and _literal_of(entry.get(lower)) is not None and lower in _condition_names(condition):
            if counter is not None:
                return None
            counter = lower
            step = k if value[1].raw_text == "+" else -k
            step_at = stmt
    start = _literal_of(entry.get(counter)) if counter is not None else None
    if counter is None or step is None or step == 0 or not _is_number(start):
        return None
    # Nothing else may write the counter or a name the condition reads, or leave.
    rest = [stmt for stmt in body if stmt is not step_at]
    for lower in _condition_names(condition):
        if _loop_body_may_leave_or_write(source, rest, lower, activity):
            return None
    facts = _facts_from(entry, source)
    the_counter = counter
    value_now = float(start)

    def known(current: float) -> bool | None:
        return condition_value(
            condition,
            ConditionFacts(
                value=lambda lower: current if lower == the_counter else facts.value(lower),
                compare=facts.compare,
            ),
        )

    for _pass in range(_DO_COUNTER_PASSES + 1):
        if at_head:
            holds = known(value_now)
            if holds is None:
                return None
            if holds == until:
                return _Counter(the_counter, value_now)
            value_now += step
        else:
            value_now += step
            holds = known(value_now)
            if holds is None:
                return None
            if holds == until:
                return _Counter(the_counter, value_now)
    return None


def _condition_names(condition: Sequence[VbaToken]) -> set[str]:
    """The names a condition reads, lowercased."""
    return _mentioned_names(condition)


def _index_where(toks: Sequence[VbaToken], test: Callable[[VbaToken], bool]) -> int:
    """Array.prototype.findIndex."""
    for i, tok in enumerate(toks):
        if test(tok):
            return i
    return -1


def _for_counter_final_value(
    source: str,
    node: BodyNode,
    activity: ConditionalActivityTracker | None,
) -> _Counter | None:
    """What a For counter holds after a loop of literal bounds runs to its end
    (issue #263, measured in Excel 16.0): one step past its last pass, or the
    start when no pass runs. `For i = 0 To 3 ... Next` leaves i at 4, so `a(i)`
    on a Dim a(3) raises 9. A body that may leave the loop, or that writes the
    counter, keeps it unknown."""
    if not isinstance(node, ForBlockNode) or node.each or not node.control_variable:
        return None
    toks = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
    eq = _index_where(toks, lambda tok: tok.raw_text == "=")
    to = _index_where(toks, lambda tok: token_text(tok) == "to")
    step_at = _index_where(toks, lambda tok: token_text(tok) == "step")
    start = _signed_integer(toks[eq + 1 : to]) if eq > 0 and to > eq else None
    limit = _signed_integer(toks[to + 1 : step_at if step_at > 0 else len(toks)]) if to > 0 else None
    step = _signed_integer(toks[step_at + 1 :]) if step_at > 0 else 1
    if start is None or limit is None or step is None or step == 0:
        return None
    lower = node.control_variable.lower()
    if _loop_body_may_leave_or_write(source, node.body, lower, activity):
        return None
    # JavaScript numbers: the arithmetic is in doubles.
    first, last, by = float(start), float(limit), float(step)
    if by > 0:
        passes = math.floor((last - first) / by) + 1 if first <= last else 0
    else:
        passes = math.floor((first - last) / -by) + 1 if first >= last else 0
    return _Counter(lower, first + passes * by)


class _NoPass(NamedTuple):
    counter: _Counter | None


def _is_empty_array(value: Sequence[VbaToken]) -> bool:
    return (
        len(value) == 3
        and token_text(value[0]) == "array"
        and value[1].raw_text == "("
        and value[2].raw_text == ")"
    )


def _loop_runs_no_pass(source: str, node: BodyNode, entry: ReachingAssignments) -> _NoPass | None:
    """Whether a loop runs no pass, from the state it is entered with: a For whose
    bounds and step are known and pass each other, a `Do While` or `While` whose
    condition is known False, a `Do Until` whose condition is known True, and a For
    Each over `Array()`. None when it may run."""
    from .block_headers import is_loop_block

    if not is_loop_block(node):
        return None
    toks = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
    facts = _facts_from(entry, source)
    if isinstance(node, ForBlockNode) and node.each:
        in_at = _index_where(toks, lambda tok: token_text(tok) == "in")
        group = toks[in_at + 1 :]
        # A local holding `Array()`, or a Collection nothing has added to
        # (issue #483, measured in Excel 16.0).
        held = entry.get(_lower_name(group[0]) or "") if len(group) == 1 else None
        empty = in_at > 0 and (
            _is_empty_array(group)
            or held is EMPTY_COLLECTION
            or (held is not None and _is_empty_array([tok for tok in held if tok.kind is not TokenKind.COMMENT]))
        )
        return _NoPass(None) if empty else None
    if isinstance(node, ForBlockNode):
        eq = _index_where(toks, lambda tok: tok.raw_text == "=")
        to = _index_where(toks, lambda tok: token_text(tok) == "to")
        step_at = _index_where(toks, lambda tok: token_text(tok) == "step")

        def known(part: Sequence[VbaToken]) -> float | None:
            literal = _signed_integer(part)
            value: object = (
                literal
                if literal is not None
                else facts.value(_lower_name(part[0]) or "")
                if len(part) == 1
                else None
            )
            return value if _is_number(value) else None

        start = known(toks[eq + 1 : to]) if eq > 0 and to > eq else None
        limit = known(toks[to + 1 : step_at if step_at > 0 else len(toks)]) if to > 0 else None
        step = known(toks[step_at + 1 :]) if step_at > 0 else 1
        if start is None or limit is None or step is None or step == 0 or not node.control_variable:
            return None
        passed = start > limit if step > 0 else start < limit
        return _NoPass(_Counter(node.control_variable.lower(), start)) if passed else None
    # `Do While c`, `Do Until c` and `While c`; a condition after `Loop` lets one pass run.
    is_do = isinstance(node, DoBlockNode)
    head = token_text(_at(toks, 1 if is_do else 0))
    if head != "while" and head != "until":
        return None
    condition = toks[2 if is_do else 1 :]
    value = condition_value(condition, facts) if len(condition) > 0 else None
    return _NoPass(None) if value is not None and value == (head == "until") else None


def _referenced_labels(
    source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> dict[str, int]:
    """The keys of every label a statement of the body jumps to or resumes at."""
    out: dict[str, int] = {}
    for node in iter_body_nodes(body, lambda node: is_inactive_node(activity, node)):
        if is_leaf_statement(node):
            for ref in statement_label_references(source, node.span):
                out[ref.key] = out.get(ref.key, 0) + 1
    return out


def _signed_integer(toks: Sequence[VbaToken]) -> int | None:
    negative = len(toks) == 2 and toks[0].raw_text == "-"
    literal = _at(toks, 1 if negative else 0)
    if len(toks) != (2 if negative else 1) or literal is None or literal.kind is not TokenKind.INTEGER_LITERAL:
        return None
    value = parse_vba_integer_literal(literal.raw_text)
    return None if value is None else -value if negative else value


# Statement heads that may leave a loop or jump within the procedure.
_LEAVING_HEADS: frozenset[str] = frozenset({"exit", "goto", "gosub", "resume", "return", "on"})


def _loop_body_may_leave_or_write(
    source: str, body: Sequence[BodyNode], lower: str, activity: ConditionalActivityTracker | None
) -> bool:
    """Whether a loop body may leave it early, jump, or write the counter."""
    for node in iter_body_nodes(body, lambda node: is_inactive_node(activity, node)):
        if not is_leaf_statement(node):
            if isinstance(node, ForBlockNode) and node.control_variable is not None and node.control_variable.lower() == lower:
                return True
            header = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
            if lower in _passed_whole(source, header, node.span.start):
                return True
            continue
        if jump_target_label_declaration(source, node.span):
            return True
        for span in statement_and_branch_spans(node):
            toks = statement_tokens_after_leading_label(source, span)
            head = token_text(_at(toks, 0))
            if head in _LEAVING_HEADS or (head == "end" and len(toks) == 1):
                return True
            if head in _WRITING_HEADS and lower in _written_names(toks):
                return True
            bare = bare_assignment_target(source, span)
            if (bare is not None and bare[0].lower() == lower) or lower in _passed_whole(source, toks, span.start):
                return True
    return False


class _HeldIntegers:
    """The whole numbers the reaching assignments give their names, for a sum."""

    __slots__ = ("_before",)

    def __init__(self, before: ReachingAssignments) -> None:
        self._before = before

    def get(self, name: str, /) -> int | None:
        held = _literal_of(self._before.get(name))
        return held if isinstance(held, int) and not isinstance(held, bool) else None


def _is_long(value: object) -> bool:
    """Number.isInteger(value) and within the Long range."""
    if not _is_number(value):
        return False
    number = float(value)
    return math.isfinite(number) and number == math.floor(number) and -2147483648 <= number <= 2147483647


def _known_sum(value: Sequence[VbaToken], before: ReachingAssignments) -> Sequence[VbaToken] | None:
    """An expression of whole numbers and names that hold them here, as the tokens
    of its value, or None: `b + 1` with b at 5 is `6`. Kept to the Long range."""
    result = evaluate_integer_constant_expression(" ".join(tok.raw_text for tok in value), _HeldIntegers(before))
    return raw_expression_tokens(js_number_to_string(result)) if result is not None and _is_long(result) else None


def _known_call(
    source: str, value: Sequence[VbaToken], before: ReachingAssignments, target: str
) -> Sequence[VbaToken] | None:
    """A built-in call the walk can work out, as the tokens of its whole-number
    result: `n = Len(s)` with s at "abcde" is `5` (issue #685). Not into a String,
    which would hold the number's text."""
    if not any(token_name(tok) is not None and _raw(_at(value, i + 1)) == "(" for i, tok in enumerate(value)) or (
        _walk_declared is not None and _walk_declared.type(target) == "string"
    ):
        return None
    result = number_value(value, _facts_from(before, source))
    return raw_expression_tokens(js_number_to_string(result)) if result is not None and _is_long(result) else None


def _forward_goto(source: str, node: LeafStatementNode, following: BodyNode | None) -> str | None:
    """The label key a statement jumps to when it is `GoTo L`, or a one-line If
    whose only branch is `GoTo L`. None for anything else."""
    refs = statement_label_references(source, node.span)
    if len(refs) != 1:
        return None
    branches = _single_line_if_branches(node)
    if branches is not None:
        then = statement_tokens_after_leading_label(source, branches[0]) if len(branches) == 1 else []
        tail = following is not None and is_leaf_statement(following) and following.single_line_if_tail
        return refs[0].key if not tail and len(then) == 2 and token_text(then[0]) == "goto" else None
    toks = statement_tokens_after_leading_label(source, node.span)
    return refs[0].key if len(toks) == 2 and token_text(toks[0]) == "goto" else None


def _tokens_text(value: Sequence[VbaToken]) -> str:
    return " ".join(tok.raw_text for tok in value)


def _agreed(states: Sequence[ReachingAssignments]) -> ReachingAssignments:
    """What every one of the states holds alike."""
    if len(states) == 0:
        return _NONE
    out: dict[str, Sequence[VbaToken]] = {}
    for key, value in states[0].items():
        agree = True
        for state in states:
            other = state.get(key)
            if not (
                other is value
                or (
                    other is not None
                    and value is not OBJECT_NOTHING
                    and value is not EMPTY_COLLECTION
                    and other is not OBJECT_NOTHING
                    and other is not EMPTY_COLLECTION
                    and _tokens_text(other) == _tokens_text(value)
                )
            ):
                agree = False
                break
        if agree:
            out[key] = value
    return out


def _after_statement(source: str, span: Span, before: ReachingAssignments) -> ReachingAssignments:
    """What holds after one plain statement runs."""
    toks = statement_tokens_after_leading_label(source, span)
    head = token_text(_at(toks, 0))
    if head == "gosub":
        return _NONE
    # A mention of an object the walk knows may change it: `c.Add 1`.
    mentioned = _mentioned_names(toks)
    known = [
        lower
        for lower in mentioned
        if before.get(lower) is OBJECT_NOTHING or before.get(lower) is EMPTY_COLLECTION
    ]
    before = _without(before, known)
    before = _without(before, [_elements_of(lower) for lower in mentioned])
    if head in _WRITING_HEADS and not (head == "line" and token_text(_at(toks, 1)) != "input"):
        written = _without(before, _written_names(toks))
        held_object = _set_object_value(toks) if head == "set" else None
        if held_object is None:
            return written
        return with_reaching_value(written, held_object[0], held_object[1])
    after = _without(before, _passed_whole(source, toks, span.start))
    # A name passed ByRef holds what the callee leaves in it (issue #449).
    effects = _walk_call_effects(toks) if _walk_call_effects is not None else None
    if effects is not None and len(effects) > 0:
        for lower, effect in effects.items():
            after = with_reaching_value(after, lower, effect)
    bare = bare_assignment_target(source, span)
    if bare is not None and not identity_assignment(bare[0], bare[2]):
        value = [tok for tok in bare[2] if tok.kind is not TokenKind.COMMENT]
        # `d = a` copies what a holds here: `a = 0: d = a` leaves d 0, and a
        # later change to a leaves d as it was (issue #346).
        copied = _lower_name(value[0]) if len(value) == 1 else None
        # `b = b + 1` with b known: the sum, a whole number (issue #614).
        computed: Sequence[VbaToken] | None = None
        if copied is None and any(token_name(tok) is not None for tok in value):
            computed = _known_sum(value, before)
            if computed is None:
                computed = _known_call(source, value, before, bare[0].lower())
        held = before.get(copied) if copied is not None else None
        if held is None:
            held = computed if computed is not None else value
        after = with_reaching_value(after, bare[0].lower(), held)
    element = _element_assignment(toks, before) if _walk_elements else None
    if element is not None:
        after = with_reaching_value(after, element[0], element[1])
    return after


def _element_assignment(
    toks: Sequence[VbaToken], before: ReachingAssignments
) -> tuple[str, Sequence[VbaToken]] | None:
    """`a(0) = Null` on an array the procedure declares, the index a whole number
    or a local the walk knows holds one: the element's key and its value (issue
    #332). Null is the one value a rule reads from an element."""
    lower = _lower_name(_at(toks, 0))
    if not lower or lower not in _walk_arrays or _raw(_at(toks, 1)) != "(":
        return None
    close = -1
    depth = 1
    for i in range(2, len(toks)):
        depth += 1 if toks[i].raw_text == "(" else -1 if toks[i].raw_text == ")" else 0
        if depth == 0:
            close = i
            break
    if close < 0 or _raw(_at(toks, close + 1)) != "=":
        return None
    index = known_index(toks[2:close], before)
    value = [tok for tok in toks[close + 2 :] if tok.kind is not TokenKind.COMMENT]
    if index is None or len(value) != 1 or token_text(value[0]) != "null":
        return None
    return element_key(lower, index), value


def known_index(subscript: Sequence[VbaToken], held: ReachingAssignments) -> float | None:
    """A subscript's value: a whole-number literal, or a local holding one here."""
    toks = [tok for tok in subscript if tok.kind is not TokenKind.COMMENT]
    if len(toks) != 1:
        return None
    lower = _lower_name(toks[0])
    value = _literal_of(held.get(lower)) if lower is not None else _literal_of(toks)
    return value if _is_number(value) else None


def _elements_of(lower: str) -> str:
    """The marker `_without` reads as every element of the array: "a("."""
    return f"{lower}("


_Touched = Union[set[str], Literal["all"]]


def _touched_in_block(
    source: str,
    block: BodyNode,
    activity: ConditionalActivityTracker | None,
    decided: Callable[[Sequence[VbaToken]], bool | None] = lambda _condition: None,
) -> _Touched:
    """Every name a block may change, header and footer lines included, or 'all'."""
    names: set[str] = set()
    if isinstance(block, ForBlockNode) and block.control_variable:
        names.add(block.control_variable.lower())
    for span in (block_header_line_span(source, block.span), block_footer_line_span(source, block.span)):
        names.update(_passed_whole(source, statement_tokens_after_leading_label(source, span), span.start))
    body = _body_of(block)
    if body is None:
        return names
    # Upstream visits the nested lists recursively and stops at the first that
    # may be entered from anywhere; the names found then go unused, so the
    # lists are visited here in any order from a stack.
    pending: list[Sequence[BodyNode]] = [body]
    while pending:
        statements = pending.pop()
        i = 0
        while i < len(statements):
            node = statements[i]
            if is_inactive_node(activity, node):
                i += 1
                continue
            if is_leaf_statement(node):
                # A label inside the block can be reached from anywhere, and
                # what follows the block then runs with whatever that path held.
                if jump_target_label_declaration(source, node.span):
                    return "all"
                # A one-line If with no Else that is decided False changes
                # only what its condition passes, and its tail never runs.
                if _single_line_if_branches(node) is not None:
                    toks = statement_tokens_after_leading_label(source, node.span)
                    condition = if_condition_tokens(toks)
                    if (
                        condition is not None
                        and not any(token_text(tok) == "else" for tok in toks)
                        and decided(condition) is False
                    ):
                        names.update(_passed_whole(source, [toks[0], *condition], node.span.start))
                        while i + 1 < len(statements):
                            tail = statements[i + 1]
                            if not (is_leaf_statement(tail) and tail.single_line_if_tail):
                                break
                            i += 1
                        i += 1
                        continue
                touched = _touched_by(source, [node])
                if touched == "all":
                    return "all"
                names.update(touched)
                i += 1
                continue
            if isinstance(node, ForBlockNode) and node.control_variable:
                names.add(node.control_variable.lower())
            inner = _body_of(node)
            if inner is not None:
                if isinstance(node, IfBlockNode):
                    # An arm decided against never runs; one decided for is the
                    # last that may (issue #575).
                    for branch in node.branches:
                        names.update(
                            _passed_whole(
                                source,
                                statement_tokens_after_leading_label(source, branch.header_span),
                                branch.header_span.start,
                            )
                        )
                        condition = (
                            None
                            if branch.branch_kind is IfBranchKind.ELSE
                            else if_condition_tokens(statement_tokens_after_leading_label(source, branch.header_span))
                        )
                        verdict = decided(condition) if condition is not None else None
                        if verdict is False:
                            continue
                        pending.append(branch.body)
                        if verdict is True:
                            break
                else:
                    pending.append(inner)
            i += 1
    return names


def _loop_touched(
    source: str,
    node: BodyNode,
    entry: ReachingAssignments,
    activity: ConditionalActivityTracker | None,
) -> _Touched:
    """The names a block may change. In a loop, a name written only in an arm that
    what holds on every pass decides against keeps its value: with a = 0 and
    nothing else writing a, `If a > 1 Then c.Add a` never runs, so a stays 0
    (issue #575). The rounds start from nothing changed and add what the arms
    still live write, until a round adds nothing: every name left out is then
    written only where it never runs."""
    from .block_headers import is_loop_block

    if not is_loop_block(node):
        return _touched_in_block(source, node, activity)
    touched: set[str] = set()
    while True:
        facts = _facts_from(_without(entry, touched), source)
        live = _touched_in_block(source, node, activity, lambda condition: condition_value(condition, facts))
        if live == "all":
            return live
        if all(lower in touched for lower in live):
            return touched
        touched = touched | live


def _touched_by(source: str, stmts: Sequence[LeafStatementNode]) -> _Touched:
    """Every name the statements may change, or 'all' for a GoSub."""
    names: set[str] = set()
    for stmt in stmts:
        for span in statement_and_branch_spans(stmt):
            toks = statement_tokens_after_leading_label(source, span)
            head = token_text(_at(toks, 0))
            if head == "gosub":
                return "all"
            changed = _written_names(toks) if head in _WRITING_HEADS else _passed_whole(source, toks, span.start)
            names.update(changed)
            for lower in _mentioned_names(toks):
                names.add(_elements_of(lower))
            bare = bare_assignment_target(source, span)
            if bare is not None and not identity_assignment(bare[0], bare[2]):
                names.add(bare[0].lower())
    return names


def identity_assignment(name: str, value_tokens: Sequence[VbaToken]) -> bool:
    """`d = d + 0`, `d = d * 1`, `d = 1 * d`, `d = d - 0`, `d = d / 1`: an
    assignment that leaves a number as it was (issue #350)."""
    value = [tok for tok in value_tokens if tok.kind is not TokenKind.COMMENT]
    lower = name.lower()

    def is_self(tok: VbaToken) -> bool:
        return _lower_name(tok) == lower

    def literal(tok: VbaToken) -> str | None:
        if tok.kind is not TokenKind.INTEGER_LITERAL:
            return None
        raw = tok.raw_text
        return raw[:-1] if raw and raw[-1] in "%&^" else raw

    if len(value) != 3:
        return False
    a, op, b = value
    if is_self(a):
        return ((op.raw_text == "+" or op.raw_text == "-") and literal(b) == "0") or (
            (op.raw_text == "*" or op.raw_text == "/") and literal(b) == "1"
        )
    return is_self(b) and ((op.raw_text == "+" and literal(a) == "0") or (op.raw_text == "*" and literal(a) == "1"))


def _without_mentioned_objects(state: ReachingAssignments, source: str, span: Span) -> ReachingAssignments:
    """The state less every object the walk knows that the span names: a block may
    have added to it or set it."""
    known = [lower for lower, value in state.items() if value is OBJECT_NOTHING or value is EMPTY_COLLECTION]
    if len(known) == 0:
        return state
    named = _mentioned_names(statement_tokens(source, span))
    return _without(state, [lower for lower in known if lower in named])


def _set_object_value(toks: Sequence[VbaToken]) -> tuple[str, Sequence[VbaToken]] | None:
    """`Set c = Nothing` and `Set c = New Collection`: the name and what it now holds."""
    name = _lower_name(_at(toks, 1))
    if not name or _raw(_at(toks, 2)) != "=":
        return None
    value = " ".join(tok.raw_text.lower() for tok in toks[3:] if tok.kind is not TokenKind.COMMENT)
    if value == "nothing":
        return name, OBJECT_NOTHING
    if value == "new collection" or value == "new vba . collection":
        return name, EMPTY_COLLECTION
    return None


# The arrays the procedure being walked declares, by lowercased name; set while
# a walk runs.
_walk_arrays: set[str] = set()

# Whether the body being walked sets an element to Null, so the walk keeps
# element keys: elsewhere `_without` need not look for them, which would cost a
# pass over every local at each statement (issue #322).
_walk_elements = False

# A Collection's own methods, which assign none of their arguments.
_COLLECTION_METHODS: frozenset[str] = frozenset({"add", "remove", "item"})

# The Collection locals of the body being walked, by lowercased name (issue
# #665); set while a walk runs.
_walk_collections: set[str] = set()

# Whether the statement being walked sits in `With k`, k one of _walk_collections.
_walk_with_collection = False

_COLLECTION_TYPE_RE = re.compile(r"(vba\.)?collection", re.IGNORECASE | re.ASCII)


def _local_collection_names(body: Sequence[BodyNode], activity: ConditionalActivityTracker | None) -> set[str]:
    """The locals the body declares As Collection or As New Collection."""
    out: set[str] = set()
    for node in iter_body_nodes(body, lambda node: is_inactive_node(activity, node)):
        if isinstance(node, VariableGroupNode) and not node.is_const:
            for decl in node.declarations:
                if not decl.is_array and _COLLECTION_TYPE_RE.fullmatch(decl.as_type or ""):
                    out.add(decl.name.lower())
    return out


def _local_array_names(body: Sequence[BodyNode], activity: ConditionalActivityTracker | None) -> set[str]:
    """The names a procedure's Dim statements declare as arrays: their subscripts pass nothing."""
    out: set[str] = set()
    for node in iter_body_nodes(body, lambda node: is_inactive_node(activity, node)):
        if isinstance(node, VariableGroupNode) and not node.is_const:
            for decl in node.declarations:
                if decl.is_array:
                    out.add(decl.name.lower())
    return out


def _passed_whole(source: str, toks: Sequence[VbaToken], span_start: int) -> list[str]:
    from .callee_arguments import callee_keeps_argument
    from .dataflow import tracked_locals_named_whole

    hits = tracked_locals_named_whole(
        toks,
        span_start,
        lambda _lower: True,
        _READ_ONLY_INTRINSICS,
        _walk_arrays,
        callee_keeps_argument(source),
    )
    # A VBA library function assigns none of its arguments: `Left$("abc", n)`
    # leaves n as it was (issue #565).
    # A Collection's own methods assign none of their arguments: `k.Add a`,
    # and `.Add a` inside `With k` (issue #665, measured in Excel 16.0).
    head = 0 if _raw(_at(toks, 0)) == "." else 1
    collection = (
        token_text(_at(toks, head + 1)) in _COLLECTION_METHODS
        and _raw(_at(toks, head)) == "."
        and (_walk_with_collection if head == 0 else (_lower_name(_at(toks, 0)) or "") in _walk_collections)
    )
    if collection:
        return []
    read_only = _printed_arguments(toks)
    for lower, at in list(hits.items()):
        index = next((i for i, tok in enumerate(toks) if span_start + tok.start == at), -1)
        if index >= 0 and (_library_function_argument(toks, index) or read_only(index)):
            del hits[lower]
    return list(hits)


def _printed_arguments(toks: Sequence[VbaToken]) -> Callable[[int], bool]:
    """The tokens Debug.Print, Debug.Assert, `Print #` and `Write #` read and never
    write, up to the end of the statement or a one-line If's Else: `Debug.Print a`
    leaves a as it was (issue #655, measured in Excel 16.0)."""
    ranges: list[tuple[int, int]] = []
    i = 0
    while i < len(toks):
        word = token_text(toks[i])
        debug = (
            word == "debug"
            and _raw(_at(toks, i + 1)) == "."
            and token_text(_at(toks, i + 2)) in ("print", "assert")
        )
        printed = (
            (word == "print" or word == "write")
            and _raw(_at(toks, i + 1)) == "#"
            and _raw(_at(toks, i - 1)) != "."
        )
        if not debug and not printed:
            i += 1
            continue
        start = i + (3 if debug else 2)
        end = start
        while end < len(toks) and toks[end].raw_text != ":" and token_text(toks[end]) != "else":
            end += 1
        ranges.append((start, end))
        i = end + 1
    return lambda index: any(start <= index < end for start, end in ranges)


def _written_names(toks: Sequence[VbaToken]) -> set[str]:
    """The names a writing statement may change: a ReDim's arrays, not the names
    its bounds read, `ReDim a(n)` leaving n as it was (issue #350); every name of
    any other."""
    if token_text(_at(toks, 0)) != "redim":
        return _mentioned_names(toks)
    start = 2 if token_text(_at(toks, 1)) == "preserve" else 1
    names: set[str] = set()
    for group in split_top_level_token_groups(toks, start, ",", len(toks)):
        lower = _lower_name(next((tok for tok in group if tok.kind is not TokenKind.COMMENT), None))
        if lower:
            names.add(lower)
    return names


# The procedures the module being walked declares, by lowercased name; set
# while a walk runs.
_walk_procedures: frozenset[str] = frozenset()

# `/^[ \t]*(?:...)[ \t]+([A-Za-z_]\w*)/gim`: JavaScript's multiline `^` also
# starts a line after a lone CR, U+2028 and U+2029.
_PROCEDURE_RE = re.compile(
    r"(?:^|(?<=[\n\r\u2028\u2029]))[ \t]*(?:(?:Public|Private|Friend|Static|Global)[ \t]+)*"
    r"(?:Sub|Function|Property[ \t]+(?:Get|Let|Set)|Declare(?:[ \t]+PtrSafe)?[ \t]+(?:Sub|Function))"
    r"[ \t]+([A-Za-z_]\w*)",
    re.IGNORECASE | re.ASCII,
)

_last_procedures: tuple[str, frozenset[str]] | None = None


def _module_procedure_names(source: str) -> frozenset[str]:
    """The Subs, Functions, Properties and Declares a module's source declares: one
    of them hides a VBA function."""
    global _last_procedures
    # One module is walked many times in a row; keeping only the last spares a
    # cache that grows with every edit.
    if _last_procedures is None or _last_procedures[0] != source:
        _last_procedures = (
            source,
            frozenset(match.group(1).lower() for match in _PROCEDURE_RE.finditer(source)),
        )
    return _last_procedures[1]


def _library_function_argument(toks: Sequence[VbaToken], at: int) -> bool:
    """Whether the name at `at` stands in the parentheses of a VBA library function's call."""
    depth = 0
    for j in range(at - 1, 0, -1):
        raw = toks[j].raw_text
        if raw == ")":
            depth += 1
        elif raw == "(":
            outer = depth
            depth -= 1
            if outer != 0:
                continue
            # `Left$(` lexes as Left and a `$`.
            call_at = j - 2 if _raw(_at(toks, j - 1)) == "$" else j - 1
            callee = _lower_name(_at(toks, call_at))
            qualified = _raw(_at(toks, call_at - 1)) == "."
            if (
                not callee
                or (qualified and token_text(_at(toks, call_at - 2)) != "vba")
                or (not qualified and callee in _walk_procedures)
            ):
                return False
            function = resolve_runtime_function(callee) or resolve_runtime_function(f"{callee}$")
            return function is not None and function.kind == "function"
    return False


def _mentioned_names(toks: Iterable[VbaToken]) -> set[str]:
    names: set[str] = set()
    for tok in toks:
        lower = _lower_name(tok)
        if lower:
            names.add(lower)
    return names


def _without(state: ReachingAssignments, names: Iterable[str]) -> ReachingAssignments:
    """The state less each name. A name drops its array's elements too, and a name
    ending "(" drops only the elements: "a(" drops "a(0)"."""
    updated: dict[str, Sequence[VbaToken]] | None = None
    elements = [key for key in state if key.endswith(")")] if _walk_elements else []
    for lower in names:
        if lower in (updated if updated is not None else state):
            if updated is None:
                updated = dict(state)
            del updated[lower]
        if len(elements) > 0:
            prefix = lower if lower.endswith("(") else _elements_of(lower)
            for key in elements:
                if key.startswith(prefix) and key in (updated if updated is not None else state):
                    if updated is None:
                        updated = dict(state)
                    del updated[key]
    return updated if updated is not None else state


def _module_compare(source: str) -> ModuleCompare:
    from .known_string_calls import module_compare

    compare: ModuleCompare = module_compare(source)
    return compare


def _facts_from(current: ReachingAssignments, source: str) -> ConditionFacts:
    """The numbers and strings the reaching assignments give their names, for a condition."""
    declared = _walk_declared

    def is_nothing(lower: str) -> bool | None:
        held = current.get(lower)
        return True if held is OBJECT_NOTHING else False if held is EMPTY_COLLECTION else None

    def is_null(lower: str) -> bool | None:
        held = current.get(lower)
        if held is VARIANT_EMPTY:
            return False
        if held is None:
            return None
        value = [tok for tok in held if tok.kind is not TokenKind.COMMENT]
        return True if len(value) == 1 and token_text(value[0]) == "null" else None

    def is_empty(lower: str) -> bool | None:
        if current.get(lower) is VARIANT_EMPTY:
            return True
        return False if _held_value(current, lower, declared) is not None else None

    return ConditionFacts(
        value=lambda lower: _held_value(current, lower, declared),
        is_nothing=is_nothing,
        range=lambda lower: _date_part_range(current.get(lower)),
        compare=_module_compare(source),
        is_null=is_null,
        type_of=lambda lower: declared.type(lower) if declared is not None else None,
        bounds=lambda lower: declared.bounds(lower) if declared is not None else None,
        is_empty=is_empty,
        count=lambda lower: 0 if current.get(lower) is EMPTY_COLLECTION else None,
    )


# The declared types a fraction is kept whole in, half to even: `n As Long = 2.5` holds 2.
_WHOLE_TYPES: frozenset[str] = frozenset({"byte", "integer", "long", "longlong", "longptr", "boolean"})


def _float_literal_value(raw: str) -> float:
    """`Number(raw.replace(/[!#@]$/, '').replace(/[dD]/, 'e'))`."""
    if raw and raw[-1] in "!#@":
        raw = raw[:-1]
    return js_number(re.sub(r"[dD]", "e", raw, count=1))


def _fround_keeps(number: float) -> bool:
    """`Math.fround(number) === number`: a Single holds the number exactly."""
    try:
        single: float = struct.unpack("<f", struct.pack("<f", number))[0]
    except OverflowError:
        return False
    return single == number


def _held_value(current: ReachingAssignments, lower: str, declared: DeclaredFacts | None) -> float | str | None:
    """The number or string a name holds where the walk knows it (issue #691): a
    literal, as its declared type stores it, a date as its serial, or a Const or
    Enum member of the module, `m = mA` and the bare `mB`. A fraction or a date
    into a type the walk does not know is not followed."""
    toks = current.get(lower)
    literal = _literal_of(toks)
    if literal is not None:
        return literal
    if toks is None:
        return declared.constant(lower) if declared is not None else None
    value = [tok for tok in toks if tok.kind is not TokenKind.COMMENT]
    if len(value) != 1:
        return None
    tok = value[0]
    if tok.kind is TokenKind.IDENTIFIER:
        if toks is VARIANT_EMPTY or declared is None:
            return None
        return declared.constant(_lower_name(tok) or "")
    declared_type = declared.type(lower) if declared is not None else None
    number = (
        _float_literal_value(tok.raw_text)
        if tok.kind is TokenKind.FLOAT_LITERAL
        else date_literal_serial(tok.raw_text)
        if tok.kind is TokenKind.DATE_LITERAL
        else None
    )
    if number is None or not math.isfinite(number) or declared_type is None:
        return None
    if declared_type in _WHOLE_TYPES:
        return bankers_round(number) + 0
    # A Single compares with a literal in ways this does not follow
    # (`f = 0.1` is True for f As Single = 0.1, measured), so only a value
    # a Single holds exactly is used; a Currency only one of four places.
    if declared_type == "single":
        return number if _fround_keeps(number) else None
    if declared_type == "currency":
        scaled = number * 10000
        return number if math.isfinite(scaled) and scaled == math.floor(scaled) else None
    return number if declared_type in ("double", "date", "variant") else None


# What VBA's date-part functions return: Second(Now) is 0 to 59.
_DATE_PART_RANGES: dict[str, tuple[int, int]] = {
    "second": (0, 59),
    "minute": (0, 59),
    "hour": (0, 23),
    "day": (1, 31),
    "month": (1, 12),
    "weekday": (1, 7),
}


def _date_part_range(value: Sequence[VbaToken] | None) -> tuple[float, float] | None:
    """The range a value lies in where it is a date part plus or minus whole
    numbers: `Second(Now) + 1000` is 1000 to 1059 (issue #565, measured in Excel
    16.0: `If b > 5000` is then False on every run)."""
    toks = [tok for tok in (value or ()) if tok.kind is not TokenKind.COMMENT]
    held: tuple[float, float] | None = None
    sign = 1
    i = 0
    while i < len(toks):
        word = token_text(toks[i])
        part: tuple[float, float]
        if toks[i].kind is TokenKind.INTEGER_LITERAL:
            n = parse_vba_integer_literal(toks[i].raw_text)
            if n is None:
                return None
            part = (n, n)
            i += 1
        elif word in _DATE_PART_RANGES and _raw(_at(toks, i + 1)) == "(" and _raw(_at(toks, i - 1)) != ".":
            close = match_paren_from(toks, i + 1)
            if close < 0:
                return None
            part = _DATE_PART_RANGES[word]
            i = close + 1
        else:
            return None
        moved = (part[0], part[1]) if sign > 0 else (-part[1], -part[0])
        held = (held[0] + moved[0], held[1] + moved[1]) if held is not None else moved
        if i >= len(toks):
            break
        if toks[i].raw_text != "+" and toks[i].raw_text != "-":
            return None
        sign = 1 if toks[i].raw_text == "+" else -1
        i += 1
    # A lone number is a literal, which the value holds exactly.
    if held is not None and any(token_text(tok) in _DATE_PART_RANGES for tok in toks):
        return held
    return None


def _literal_of(value: Sequence[VbaToken] | None) -> int | str | None:
    """A value's tokens as one number or string literal, a sign allowed."""
    toks = [tok for tok in (value or ()) if tok.kind is not TokenKind.COMMENT]
    negative = len(toks) == 2 and toks[0].raw_text == "-"
    tok = _at(toks, 1 if negative else 0)
    if tok is None or len(toks) != (2 if negative else 1):
        return None
    if tok.kind is TokenKind.INTEGER_LITERAL:
        number = parse_vba_integer_literal(tok.raw_text)
        return None if number is None else -number if negative else number
    if tok.kind is TokenKind.STRING_LITERAL and not negative:
        return tok.raw_text[1:-1].replace('""', '"')
    return None


def _known_arm(
    source: str, node: BodyNode, entry: ReachingAssignments
) -> tuple[list[list[BodyNode]], list[BodyNode] | None] | None:
    """The arm of an If or a Select that runs when the walk knows its outcome:
    every arm, and the one taken (None when none is). None when the outcome is
    not known."""
    from .block_headers import select_arms

    facts = _facts_from(entry, source)
    if isinstance(node, IfBlockNode):
        arms = [branch.body for branch in node.branches]
        for branch in node.branches:
            if branch.branch_kind is IfBranchKind.ELSE:
                return arms, branch.body
            condition = if_condition_tokens(statement_tokens_after_leading_label(source, branch.header_span))
            known = condition_value(condition, facts) if condition is not None else None
            if known is None:
                return None
            if known:
                return arms, branch.body
        return arms, None
    if not isinstance(node, SelectBlockNode):
        return None
    # `Select Case d` with d known: the first Case whose values match.
    header = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
    selector: float | str | None = None
    if len(header) == 3 and token_text(header[1]) == "case":
        selector = _literal_of(header[2:])
        if selector is None:
            selector = _facts_from(entry, source).value(_lower_name(header[2]) or "")
    if selector is None:
        return None
    # A string selector matches by the module's Option Compare (issue #686).
    compare = _module_compare(source)
    select_arm_lists: list[list[BodyNode]] = select_arms(source, node.body)
    for arm in select_arm_lists:
        case_line = next(
            (
                stmt
                for stmt in arm
                if is_leaf_statement(stmt)
                and token_text(_at(statement_tokens_after_leading_label(source, stmt.span), 0)) == "case"
            ),
            None,
        )
        if case_line is None:
            continue
        matched = _case_matches(statement_tokens_after_leading_label(source, case_line.span), selector, compare)
        if matched is None:
            return None
        if matched:
            return select_arm_lists, arm
    return select_arm_lists, None


def _js_type(value: float | str) -> str:
    """`typeof value` for a literal or a selector."""
    return "string" if isinstance(value, str) else "number"


def _case_matches(toks: Sequence[VbaToken], selector: float | str, compare: ModuleCompare) -> bool | None:
    """Whether a `Case` line's values take the selector: literals, `Is op v`,
    `a To b`, Else. A string compares by the module's Option Compare (issue #686),
    as condition_value decides it."""
    if token_text(_at(toks, 1)) == "else":
        return True

    def literal(value: float | str) -> str:
        return js_number_to_string(value) if not isinstance(value, str) else '"' + value.replace('"', '""') + '"'

    def decide(text: str) -> bool | None:
        return condition_value(raw_expression_tokens(text), ConditionFacts(value=lambda _lower: None, compare=compare))

    any_unknown = False
    for item in split_top_level_token_groups(toks[1:], 0, ","):
        matched: bool | None
        to = _index_where(item, lambda tok: token_text(tok) == "to")
        second = _at(item, 1)
        if token_text(_at(item, 0)) == "is" and second is not None and second.kind is TokenKind.OPERATOR:
            value = _literal_of(item[2:])
            matched = (
                None
                if value is None or _js_type(value) != _js_type(selector)
                else decide(f"{literal(selector)} {second.raw_text} {literal(value)}")
            )
        elif to > 0:
            low = _literal_of(item[:to])
            high = _literal_of(item[to + 1 :])
            matched = (
                None
                if low is None
                or high is None
                or _js_type(low) != _js_type(selector)
                or _js_type(high) != _js_type(selector)
                else decide(f"{literal(selector)} >= {literal(low)} And {literal(selector)} <= {literal(high)}")
            )
        else:
            value = _literal_of(item)
            matched = (
                None
                if value is None or _js_type(value) != _js_type(selector)
                else decide(f"{literal(selector)} = {literal(value)}")
            )
        if matched is True:
            return True
        if matched is None:
            any_unknown = True
    return None if any_unknown else False


def _mark_unreachable(node: BodyNode, dead: set[int]) -> None:
    """Marks a statement, and every statement inside it, as never running."""
    pending: list[BodyNode] = [node]
    while pending:
        current = pending.pop()
        dead.add(id(current))
        if isinstance(current, IfBlockNode):
            for branch in reversed(current.branches):
                pending.extend(reversed(branch.body))
            continue
        body = _body_of(current)
        if body is not None:
            pending.extend(reversed(body))


def _record(out: dict[int, ReachingAssignments], stmt: BodyNode, current: ReachingAssignments) -> None:
    if len(current) > 0:
        out[id(stmt)] = current
