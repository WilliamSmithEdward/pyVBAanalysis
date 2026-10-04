"""What a local of a user-defined type holds in its members, statement by
statement (XLIDE issues #248 and #253, each measured in Excel 16.0): a dynamic
array field has no elements until a ReDim gives it bounds, and none again after
Erase; an object field is Nothing until a Set; a numeric field is 0 until an
assignment stores a literal.

Only a local of a module Type is followed, not Static, from its Dim. A whole use
of the local or of a field that holds more than one value (`Fill t`, `t = u`,
`FillArr t.dyn`, `Take t.o`, `Bump t.a`), LSet, RSet, Input, Get, a label and a
GoSub end what is known. Blocks are entered as issue #237 enters them; entering a
With leaves its subject alone.

Ported from xlide_vscode/src/analyzer/diagnostics/typeMemberState.ts.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING, TypeGuard, Union

from ..conditional import ConditionalActivityTracker
from ..js_compat import js_number
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import BodyNode, LeafStatementNode, ProcedureNode, StatementNode, is_leaf_statement
from ..symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbolKind
from .type_fields import (
    FieldStep,
    ModuleTypes,
    TypeRoot,
    WithSubject,
    field_chain,
    is_fixed_array_field,
    is_leading_dot,
    type_key,
    with_subject,
    with_subjects_in,
)
from .walker import is_inactive_node, statement_tokens_after_leading_label, token_name, token_text

if TYPE_CHECKING:
    from .rules.arrays import FixedArrayBound


@dataclass(frozen=True, slots=True)
class KnownNumber:
    """A numeric field's value."""

    number: int | float


# A dynamic array field with no elements, an object field that is Nothing, a
# dynamic array's bounds, or a number; a String field still "", a Variant field
# still Empty, and a Collection field Set to a New Collection that nothing has
# added to (issue #417). The strings are 'unallocated', 'nothing', 'emptyString',
# 'empty' and 'emptyCollection'.
MemberState = Union[str, "FixedArrayBound", KnownNumber]


def is_array_bounds(state: MemberState | None) -> TypeGuard[FixedArrayBound]:
    return state is not None and not isinstance(state, (str, KnownNumber))


def is_known_number(state: MemberState | None) -> TypeGuard[KnownNumber]:
    return isinstance(state, KnownNumber)


# The member states a statement sees, or a block its header; a single-line If's
# branch sees them less what its condition names.
MemberStatesAt = Callable[[BodyNode, int], Mapping[str, MemberState]]

_NUMBER_TYPES: frozenset[str] = frozenset(
    {"byte", "integer", "long", "longlong", "longptr", "currency", "single", "double", "decimal"}
)

_NONE: Mapping[str, MemberState] = {}

_FORGETTING_HEADS: frozenset[str] = frozenset({"lset", "rset", "input", "get", "line", "mid", "mid$"})


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) out of range."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return None if tok is None else tok.raw_text


def _never_object(_type: str) -> bool:
    return False


@dataclass(frozen=True, slots=True)
class _Chain:
    steps: list[FieldStep]
    root: str | None = None


@dataclass(frozen=True, slots=True)
class _Branch:
    then: int
    states: Mapping[str, MemberState]


def type_member_states_at(
    source: str,
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    types: ModuleTypes,
    activity: ConditionalActivityTracker | None,
    option_base: int,
    is_object_type: Callable[[str], bool] = _never_object,
) -> MemberStatesAt:
    from ..types.type_inference import procedure_symbol_for

    # Keyed by node identity; each entry holds its node so the id stays its own.
    out: dict[int, tuple[BodyNode, Mapping[str, MemberState]]] = {}
    branches: dict[int, tuple[BodyNode, _Branch]] = {}

    def at(stmt: BodyNode, offset: int) -> Mapping[str, MemberState]:
        entry = branches.get(id(stmt))
        branch = entry[1] if entry is not None and entry[0] is stmt else None
        if branch is not None and offset > branch.then:
            return branch.states
        seen = out.get(id(stmt))
        return seen[1] if seen is not None and seen[0] is stmt else _NONE

    roots: dict[str, str] = {}
    states: dict[str, MemberState] = {}
    proc_symbol = procedure_symbol_for(symbols, proc)
    for child in (proc_symbol.children if proc_symbol is not None else None) or []:
        type_ = type_key(child.as_type)
        if (
            child.kind is not VbaSymbolKind.LOCAL_VARIABLE
            or child.visibility is SymbolVisibility.STATIC
            or child.is_array
            or not type_
            or type_ not in types
        ):
            continue
        lower = child.name.lower()
        roots[lower] = type_
        for key, state in _initial_states(types, type_, lower, is_object_type, 0):
            states[key] = state
    if len(states) == 0:
        return at
    subjects = with_subjects_in(source, proc, activity, symbols, types)
    # True while a recorded statement or a snapshot holds the current map, which
    # is then copied before a change.
    shared = False

    def record(node: BodyNode) -> None:
        nonlocal shared
        out[id(node)] = (node, states)
        shared = True

    def set_(key: str, state: MemberState | None) -> None:
        nonlocal states, shared
        if shared:
            states = dict(states)
            shared = False
        if state is None:
            states.pop(key, None)
        else:
            states[key] = state

    def forget(names: Iterable[str]) -> None:
        for name in list(names):
            prefix = f"{name}."
            for key in list(states.keys()):
                if key == name or key.startswith(prefix):
                    set_(key, None)

    def roots_in(toks: Sequence[VbaToken], subject: WithSubject | None) -> set[str]:
        """The tracked locals a statement names, and its With's local where it
        reaches the subject with a `.`."""
        named: set[str] = set()
        subject_root = subject.path.split(".")[0] if subject is not None and subject.path is not None else None
        for i, tok in enumerate(toks):
            name = token_name(tok)
            lower = name.lower() if name is not None else None
            if lower and lower in roots and _raw(toks, i - 1) != "." and _raw(toks, i - 1) != "!":
                named.add(lower)
            if tok.raw_text == "." and subject_root and subject_root in roots and is_leading_dot(toks, i):
                named.add(subject_root)
        return named

    def chain_at(toks: Sequence[VbaToken], i: int, subject: WithSubject | None) -> _Chain | None:
        """The fields a chain at `i` reaches from a tracked local, or from the With's subject."""
        if toks[i].raw_text == ".":
            subject_root = (
                subject.path.split(".")[0] if subject is not None and subject.path is not None else None
            )
            if (
                subject is not None
                and subject.type
                and subject_root
                and subject_root in roots
                and is_leading_dot(toks, i)
            ):
                root = TypeRoot(type=subject.type, path=subject.path, display=subject.display, dot=i)
                return _Chain(steps=field_chain(toks, root, types))
            return None
        name = token_name(toks[i])
        lower = name.lower() if name is not None else None
        type_ = roots.get(lower) if lower else None
        if not lower or not type_ or _raw(toks, i - 1) == "." or _raw(toks, i - 1) == "!":
            return None
        if _raw(toks, i + 1) == ".":
            root = TypeRoot(type=type_, path=lower, display=toks[i].raw_text, dot=i + 1)
            return _Chain(steps=field_chain(toks, root, types), root=lower)
        return _Chain(steps=[], root=lower)

    def last_step(chain: _Chain | None) -> FieldStep | None:
        return chain.steps[-1] if chain is not None and chain.steps else None

    def visit(node: BodyNode) -> None:
        from ..flow.procedure_labels import jump_target_label_declaration
        from .rules.arrays import FixedArrayBound, literal_dimensions

        nonlocal shared
        if not is_leaf_statement(node):
            return
        if jump_target_label_declaration(source, node.span):
            forget(list(roots))
        toks = statement_tokens_after_leading_label(source, node.span)
        subject = subjects.get(node.span.start)
        record(node)
        head = token_text(_at(toks, 0))
        if head == "gosub":
            forget(list(roots))
            return
        # A single-line If runs its branch only sometimes, and its condition may
        # guard it: `If Not t.o Is Nothing Then t.o.Add 1`.
        single_line_if = (
            isinstance(node, StatementNode) and node.single_line_if_branches is not None
        ) or node.single_line_if_tail is True
        if single_line_if:
            then = next((i for i, tok in enumerate(toks) if token_text(tok) == "then"), -1)
            if then > 0:
                forget(roots_in(toks[:then], subject))
                branches[id(node)] = (node, _Branch(then=node.span.start + toks[then].start, states=states))
                shared = True
            forget(roots_in(toks, subject))
            return
        if head == "redim" or head == "erase":
            start = 2 if token_text(_at(toks, 1)) == "preserve" else 1
            for group in _split_groups(toks[start:]):
                step = last_step(chain_at(group, 0, subject))
                tracked = step is not None and step.path is not None and step.path in states
                if (
                    tracked
                    and step is not None
                    and step.path is not None
                    and head == "erase"
                    and step.field.is_array
                    and step.open is None
                    and step.at == len(group) - 1
                ):
                    set_(step.path, "unallocated")
                    continue
                # A literal ReDim of a field under Option Base 0; Option Base 1 was not measured.
                dims = (
                    literal_dimensions(group[step.open + 1 : step.close], 0)
                    if tracked
                    and step is not None
                    and head == "redim"
                    and step.open is not None
                    and option_base == 0
                    else None
                )
                if dims and step is not None and step.path is not None:
                    set_(step.path, FixedArrayBound(name=step.path, dims=tuple(dims), origin="ReDim"))
                    continue
                if tracked and step is not None and step.path is not None:
                    forget([step.path])
                else:
                    forget(roots_in(group, subject))
            return
        if head in _FORGETTING_HEADS:
            forget(roots_in(toks, subject))
            return
        # `Set t.o = ...` and `t.a = 5`: what the target holds after the statement.
        is_set = head == "set"
        target_at = 1 if is_set or head == "let" else 0
        target = last_step(chain_at(toks, target_at, subject)) if _at(toks, target_at) is not None else None
        target_end = (target.close if target.close is not None else target.at) if target is not None else -1
        assigns = target is not None and _raw(toks, target_end + 1) == "=" and target.open is None
        for i in range(target_end + 2 if assigns else 0, len(toks)):
            chain = chain_at(toks, i, subject)
            if chain is None:
                continue
            last = last_step(chain)
            if last is None:
                if chain.root:
                    forget([chain.root])  # the local used whole
                continue
            end = last.close if last.close is not None else last.at
            # `t.c.Add 1` and `For Each`: the Collection is no longer known empty.
            if last.path and states.get(last.path) == "emptyCollection" and _raw(toks, end + 1) != "(":
                forget([last.path])
            whole = (
                last.open is None
                and _raw(toks, end + 1) != "."
                and _raw(toks, end + 1) != "!"
                and _raw(toks, end + 1) != "("
            )
            bound = token_text(_at(toks, i - 2)) in ("ubound", "lbound") and _raw(toks, i - 1) == "("
            is_type_value = (
                not last.field.is_array and last.field.type is not None and last.field.type in types
            )
            if (
                last.path
                and whole
                and not bound
                and (is_type_value or last.field.is_array or _passed_whole(toks, i, end))
            ):
                forget([last.path])
        if (
            not assigns
            or target is None
            or not target.path
            or (target.path not in states and not _is_tracked_scalar(target, is_object_type))
        ):
            return
        value = [tok for tok in toks[target_end + 2 :] if tok.kind is not TokenKind.COMMENT]
        if is_set:
            new_collection = (
                len(value) == 2
                and token_text(value[0]) == "new"
                and token_text(value[1]) == "collection"
                and target.field.type == "collection"
            )
            if (
                len(value) == 1
                and token_text(value[0]) == "nothing"
                and is_object_type(target.field.type or "")
            ):
                set_(target.path, "nothing")
            else:
                set_(target.path, "emptyCollection" if new_collection else None)
            return
        number = (
            _literal_number(value)
            if target.field.type and target.field.type in _NUMBER_TYPES and not target.field.is_array
            else None
        )
        set_(target.path, None if number is None else KnownNumber(number))

    def snapshot() -> dict[str, MemberState]:
        nonlocal shared
        shared = True
        return states

    def restore(saved: Mapping[str, MemberState]) -> None:
        nonlocal states, shared
        states = dict(saved)
        shared = False

    def enter(node: BodyNode) -> None:
        # What a For Each header reads, before its own touches are forgotten.
        record(node)

    def touches(stmt: LeafStatementNode) -> Iterable[str]:
        toks = statement_tokens_after_leading_label(source, stmt.span)
        subject = subjects.get(stmt.span.start)
        # Entering `With t` evaluates t, and changes nothing.
        # `With t.c` may add to the Collection.
        is_with = token_text(_at(toks, 0)) == "with"
        held_step = last_step(chain_at(toks, 1, subject)) if is_with and len(toks) > 1 else None
        held = held_step.path if held_step is not None else None
        if held and states.get(held) == "emptyCollection":
            return [held]
        if is_with and with_subject(toks, symbols, proc, types, subject) is not None:
            return []
        return roots_in(toks, subject)

    from .dataflow import BlockEnteringState, walk_entering_blocks

    walk_entering_blocks(
        source,
        proc.body,
        lambda node: is_inactive_node(activity, node),
        visit,
        BlockEnteringState(
            snapshot=snapshot,
            restore=restore,
            forget=forget,
            enter=enter,
            touches=touches,
        ),
    )
    return at


def _is_tracked_scalar(step: FieldStep, is_object_type: Callable[[str], bool]) -> bool:
    """A scalar or object field an assignment can give a state to."""
    type_ = step.field.type
    return not step.field.is_array and type_ is not None and (type_ in _NUMBER_TYPES or is_object_type(type_))


def _initial_states(
    types: ModuleTypes,
    type_: str,
    prefix: str,
    is_object_type: Callable[[str], bool],
    depth: int,
) -> list[tuple[str, MemberState]]:
    """Each member a local starts with a state for: dynamic arrays, objects and
    numbers, through fields that are not arrays. Recursion is bounded at depth 8."""
    out: list[tuple[str, MemberState]] = []
    for lower, field in (types.get(type_) or {}).items():
        key = f"{prefix}.{lower}"
        if field.is_array:
            if not is_fixed_array_field(field):
                out.append((key, "unallocated"))
        elif field.type and field.type in types:
            if depth < 8:
                out.extend(_initial_states(types, field.type, key, is_object_type, depth + 1))
        elif field.type and field.type in _NUMBER_TYPES:
            out.append((key, KnownNumber(0)))
        elif field.type and is_object_type(field.type):
            out.append((key, "nothing"))
        elif field.type == "string":
            out.append((key, "emptyString"))
        elif not field.type or field.type == "variant":
            out.append((key, "empty"))
    return out


def _passed_whole(toks: Sequence[VbaToken], start: int, end: int) -> bool:
    """Whether the chain from `start` to `end` stands alone in an argument slot,
    where a call may take it ByRef: `Bump t.a`, `Take(t.o)`, `x = F(t.a)`."""
    prev = _at(toks, start - 1)
    nxt = _at(toks, end + 1)
    opens = (
        prev is None
        or prev.raw_text == "("
        or prev.raw_text == ","
        or prev.kind is TokenKind.IDENTIFIER
        or (prev.kind is TokenKind.KEYWORD and token_text(prev) == "call")
    )
    closes = (
        nxt is None
        or nxt.raw_text == ")"
        or nxt.raw_text == ","
        or nxt.raw_text == ":"
        or nxt.kind is TokenKind.COMMENT
    )
    return opens and closes


_TYPE_SUFFIX_RE = re.compile(r"[%&^!#@]$")
_RADIX_PREFIX_RE = re.compile(r"^&[hHoO]")


def _literal_number(value: Sequence[VbaToken]) -> int | float | None:
    """The value of a plain number literal, signed or not, or None."""
    signed = len(value) == 2 and (value[0].raw_text == "-" or value[0].raw_text == "+")
    literal = value[0] if len(value) == 1 else value[1] if signed else None
    if literal is None or (
        literal.kind is not TokenKind.INTEGER_LITERAL and literal.kind is not TokenKind.FLOAT_LITERAL
    ):
        return None
    # JavaScript's `$` does not match before a trailing newline; a token has none.
    raw = _TYPE_SUFFIX_RE.sub("", literal.raw_text, count=1)
    if _RADIX_PREFIX_RE.match(raw):
        return None
    number = js_number(re.sub(r"[dD]", "E", raw, count=1))
    if not math.isfinite(number):
        return None
    signed_number = -number if value[0].raw_text == "-" else number
    return int(signed_number) if signed_number.is_integer() else signed_number


def _split_groups(toks: Sequence[VbaToken]) -> list[list[VbaToken]]:
    """A statement's comma-separated groups outside parentheses."""
    out: list[list[VbaToken]] = [[]]
    depth = 0
    for tok in toks:
        if tok.raw_text == "(":
            depth += 1
        elif tok.raw_text == ")":
            depth -= 1
        elif tok.raw_text == "," and depth == 0:
            out.append([])
            continue
        if tok.kind is not TokenKind.COMMENT:
            out[-1].append(tok)
    return [group for group in out if len(group) > 0]
