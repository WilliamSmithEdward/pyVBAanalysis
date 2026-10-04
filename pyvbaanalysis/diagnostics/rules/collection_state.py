"""Rule family: a VBA.Collection whose contents the code makes plain (XLIDE issue #121).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/collectionState.ts.

Every case was measured in Excel 16.0 (build 20326, 2026-09-26) on a local
`New Collection` with only the Adds shown; each compiles and raises every
time it runs.

 - collection-index-out-of-range: `c(1)`, `c(0)`, `c(-1)` or `c.Remove 1`
   with nothing added -> 5 (Invalid procedure call or argument); `c(0)`,
   `c(-1)`, `c.Item(2)`, `c(3)` after two Adds, `c.Remove 0`, `c.Remove 2`
   after one Add -> 9 (Subscript out of range). Collections are 1-based.
 - collection-key-not-found: `c("nokey")`, `c.Item("nokey")`, `c.Remove "x"`
   when the key was never added, or was removed -> 5. Keys compare without
   case: `c.Add 1, "k"` then `c("K")` runs.
 - collection-key-in-use: `c.Add 1, "k"` then `c.Add 2, "K"` -> 457.
 - collection-add-argument (issue #219): a key that is a number or True,
   `c.Add "x", 5` -> 13, and so is a Variant never assigned, which is Empty;
   Before and After together -> 5. Before or After on an empty collection
   -> 5, and outside 1 to Count -> 9, are collection-index-out-of-range's.
 - array-subscript-out-of-bounds (issue #248): `c.Add Array(1, 2)` then
   `c(1)(5)` indexes past the end of the array the item holds -> 9.

The rule follows a procedure's top-level statements in order, as the file
rule does: a block ends what is known, and any use of the variable other
than Add, Remove, Item, Count and indexing ends it too.
"""

from __future__ import annotations

import json
import math
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field, replace
from typing import AbstractSet, TypeVar, Union

from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import (
    IntegerConstantLookup,
    bankers_round,
    evaluate_integer_constant_expression,
    parse_vba_integer_literal,
    resolve_raw_integer_constants,
)
from ...flow.procedure_labels import jump_target_label_declaration
from ...host.host_model import HostObjectModel
from ...js_compat import js_number_to_string, js_trim
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ForBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    ProcKind,
    Span,
    StatementNode,
    VariableGroupNode,
    WithBlockNode,
    is_leaf_statement,
)
from ...symbols.symbol_model import ModuleSymbols, VbaSymbol, VbaSymbolKind
from ...types.type_inference import (
    known_local_literal_values_at,
    procedure_symbol_for,
    unreachable_statements_in,
    with_known_locals,
)
from ...types.type_names import is_known_scalar_type, normalize_type
from ..call_extraction import string_literal_value
from ..callable_signatures import procedure_integer_constant_lookup
from ..callee_arguments import CalleeMemberCalls, callee_member_calls
from ..const_expr import collect_module_literal_integer_constants
from ..context import PushFn
from ..dataflow import BlockEnteringState, walk_entering_blocks
from ..loop_counters import CounterAtom, LoopCounter, counter_text, loop_counters_at, numeric_counter_passes
from ..walker import (
    active_module_members,
    block_header_line_span,
    for_each_variable_group,
    match_paren_from,
    raw_expression_tokens,
    set_assignment_target,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .arrays import FixedArrayBound, array_value_shape, module_option_base, shape_subscript_violation
from .shared import name_mentions, names_in


@dataclass(slots=True, eq=False)
class _CollectionContents:
    # Keys in element order; None for an element added without a key.
    items: list[str | None] = field(default_factory=list)
    # False once an Add named a key the rule could not read, or ordered by Before/After.
    keys_known: bool = True
    # The bounds of the array each element holds, where an Add gave it
    # `Array(...)`; aligned with items.
    shapes: list[FixedArrayBound | None] = field(default_factory=list)
    # What each element is, where an Add gave it a tracked Collection, which the
    # element then shares, or a number literal (XLIDE issue #452); aligned with items.
    held: list[_Held] = field(default_factory=list)
    # Set once a name for this collection stopped being followed: it may have changed unseen.
    stale: bool = False


_Held = Union[_CollectionContents, str, None]  # a Collection, 'number', 'string', or unknown

_T = TypeVar("_T")


def _empty_contents() -> _CollectionContents:
    return _CollectionContents()


def _held_item_read(
    value: Sequence[VbaToken], states: Mapping[str, _CollectionContents]
) -> tuple[_Held] | None:
    """What `c(1)`, `c("k")` or `c.Item(1)`, the whole of a value, reads from a tracked collection."""
    toks = [tok for tok in value if tok.kind is not TokenKind.COMMENT]
    contents = states.get(_lower_name(_token_at(toks, 0)) or "")
    if _raw_text_at(toks, 1) == "(":
        open_index = 1
    elif (
        _raw_text_at(toks, 1) == "."
        and token_text(_token_at(toks, 2)) == "item"
        and _raw_text_at(toks, 3) == "("
    ):
        open_index = 3
    else:
        open_index = -1
    if contents is None or contents.stale or open_index < 0 or match_paren_from(toks, open_index) != len(toks) - 1:
        return None
    arg = toks[open_index + 1 : len(toks) - 1]
    key = _literal_key_text(arg)
    index = _literal_index(arg)
    if index is None and key is not None and contents.keys_known and key in contents.items:
        index = contents.items.index(key) + 1
    return (contents.held[index - 1],) if index is not None and 1 <= index <= len(contents.held) else None


def _forget_collection(states: dict[str, _CollectionContents], lower: str) -> None:
    """Stops following a name; what it named may now change unseen, so an element sharing it is no longer read."""
    contents = states.get(lower)
    if contents is not None:
        contents.stale = True
    states.pop(lower, None)


# The name a `With New Collection` block's collection is followed under.
_NEW_WITH_SUBJECT = "New Collection"

# A With block's subject: a name, or the tokens of an element, `c ( 1 )`.
_WithSubject = Union[str, Sequence[VbaToken]]

_WITH_NEW_COLLECTION = re.compile(r"\bWith[ \t]+New[ \t]+Collection\b", re.IGNORECASE | re.ASCII)


def check_collection_state(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    symbols: ModuleSymbols | None = None,
    project_integer_constants: Mapping[str, str | None] | None = None,
    project_visible_symbols: Sequence[VbaSymbol] | None = None,
    host_model: HostObjectModel | None = None,
) -> None:
    module_constants = collect_module_literal_integer_constants(
        mod, activity, resolve_raw_integer_constants(project_integer_constants or {}, {})
    )
    callee_calls = callee_member_calls(source)
    option_base = module_option_base(mod, activity)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        _check_procedure(
            source, member, activity, push, symbols, project_visible_symbols, host_model,
            module_constants, callee_calls, option_base,
        )


def _check_procedure(
    source: str,
    member: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    symbols: ModuleSymbols | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    host_model: HostObjectModel | None,
    module_constants: Mapping[str, float | None],
    callee_calls: CalleeMemberCalls,
    option_base: int,
) -> None:
    auto_instanced = _collection_locals(member, activity)
    states: dict[str, _CollectionContents] = {}
    for name in auto_instanced.new_locals:
        states[name] = _empty_contents()
    if (
        len(states) == 0
        and len(auto_instanced.plain_locals) == 0
        and _WITH_NEW_COLLECTION.search(source[member.span.start : member.span.end]) is None
    ):
        return
    # A Variant local named nowhere but one statement is Empty there.
    mentions: list[Mapping[str, int]] = []

    def is_empty(lower: str) -> bool:
        if lower not in auto_instanced.variant_locals:
            return False
        if not mentions:
            mentions.append(name_mentions(source, member, activity))
        return mentions[0].get(lower) == 1

    # An index through a Const or a local with one known value (issue #238).
    constants = (
        procedure_integer_constant_lookup(
            member, module_constants, symbols, project_visible_symbols, activity, host_model
        )
        if symbols is not None
        else None
    )
    values_at = known_local_literal_values_at(source, member, symbols, activity) if symbols is not None else None
    # Code that never runs changes nothing and raises nothing (issue #406).
    unreachable = unreachable_statements_in(source, member, symbols, activity) if symbols is not None else None

    # A scalar or Variant local, or the Function's result, which a Collection's
    # value is Let into (issue #452).
    def holds_value(type_name: str | None) -> bool:
        normalized = normalize_type(type_name)
        return normalized is None or normalized == "variant" or is_known_scalar_type(normalized)

    proc_symbol = procedure_symbol_for(symbols, member) if symbols is not None else None
    local_symbols = (proc_symbol.children if proc_symbol is not None else None) or []
    scalars = {
        child.name.lower()
        for child in local_symbols
        if child.kind is VbaSymbolKind.LOCAL_VARIABLE and not child.is_array and holds_value(child.as_type)
    }
    if member.proc_kind is ProcKind.FUNCTION and holds_value(member.return_type):
        scalars.add(member.name.lower())

    def scalar_local(lower: str) -> bool:
        return lower in scalars

    # The subject of each With block the walk is in: a tracked local's name, its
    # element `c(1)`, `New Collection` for a new one, or None for anything else.
    with_subjects: list[_WithSubject | None] = []

    def current_subject() -> _WithSubject | None:
        return with_subjects[-1] if with_subjects else None

    # Blocks are entered with the state they start with (issue #237).
    def visit(node: BodyNode) -> None:
        if not is_leaf_statement(node) or (unreachable is not None and id(node) in unreachable):
            return  # a Dim inside the body declares, and runs nothing
        if isinstance(node, StatementNode) and node.single_line_if_branches is not None:
            _forget_mentioned(source, node.span, states)
            # `If x Then .Add 20` inside `With c` may change c (issue #584).
            within = current_subject()
            if within is not None and _reaches_subject(statement_tokens_after_leading_label(source, node.span), within):
                _forget_collection(states, _subject_name(within))
            return
        own = statement_tokens_after_leading_label(source, node.span)
        if len(own) == 0:
            return
        # Inside `With c` or `With New Collection`, `.Item(2)` is the subject's
        # (issue #295, measured in Excel 16.0).
        subject = current_subject()
        toks = with_receiver(own, subject) if subject is not None else own
        # A label may be reached from anywhere; a GoSub may run any statement.
        if jump_target_label_declaration(source, node.span) or token_text(toks[0]) == "gosub":
            states.clear()
        # `Set c = New Collection` starts an empty collection. `Set o = c` makes o
        # and c one collection, so they share one state and an Add through either
        # is seen by both (issue #147). Any other Set ends tracking of its target,
        # and of every tracked collection its value names, since the value's new
        # holder can change it unseen.
        target = set_assignment_target(source, node.span)
        if target is not None:
            lower = target[0].lower()
            equals = next((k for k, tok in enumerate(toks) if tok.raw_text == "="), -1)
            value = toks[equals + 1 :]
            is_collection_local = lower in auto_instanced.plain_locals or lower in auto_instanced.new_locals
            aliased = _lower_name(value[0]) if len(value) == 1 else None
            if (
                is_collection_local
                and len(value) == 2
                and token_text(value[0]) == "new"
                and token_text(value[1]) == "collection"
            ):
                states[lower] = _empty_contents()
                return
            if is_collection_local and aliased is not None and aliased in states:
                states[lower] = states[aliased]
                return
            item = _held_item_read(value, states)
            if item is not None and item[0] in ("number", "string"):
                first = value[0]
                last = value[-1]
                push(
                    "variantValueMisuse",
                    f"'{''.join(tok.raw_text for tok in value)}' holds a {item[0]}, not an object, so Set has "
                    "nothing to assign. This will raise Run-time error '424': Object required.",
                    Span(node.span.start + first.start, node.span.start + last.end),
                )
            _forget_collection(states, lower)
            for tok in value:
                mentioned = _lower_name(tok)
                if mentioned and mentioned in states:
                    _forget_collection(states, mentioned)
            return
        known_here = values_at(node) if values_at is not None else None
        lookup = (
            with_known_locals(constants, known_here)
            if constants is not None and known_here is not None
            else None
        )

        # `c(1.6)` rounds to 2, half to even (issue #349, measured in Excel 16.0).
        def index_of(arg: Sequence[VbaToken]) -> int | float | None:
            literal = _literal_index(arg)
            if literal is not None:
                return literal
            if len(arg) == 1 and arg[0].kind is TokenKind.FLOAT_LITERAL:
                number = _float_literal_value(arg[0].raw_text)
                if number is not None:
                    return bankers_round(number)
            return (
                evaluate_integer_constant_expression(" ".join(tok.raw_text for tok in arg), lookup)
                if lookup is not None
                else None
            )

        def key_of(arg: Sequence[VbaToken]) -> str | None:
            name = _lower_name(arg[0]) if len(arg) == 1 else None
            held = known_here.get(name) if name and known_here is not None else None
            literal = _literal_key_text(arg)
            if literal is not None:
                return literal
            if held is not None and held.kind == "string" and not held.content_mutated and isinstance(held.value, str):
                return held.value
            return None

        # `R1 c` where R1 only adds to or removes from its parameter: c changes as
        # those calls change it (issue #685).
        replays = replayed_calls(toks, states, callee_calls)
        if replays is not None:
            for lower, calls in replays.items():
                for call in calls:
                    raised: list[Replayed] = []

                    def capture(
                        rule: str, message: str, span: Span, data: object = None, raised: list[Replayed] = raised
                    ) -> None:
                        if not raised:
                            raised.append(Replayed(rule, message))

                    _check_statement(
                        node.span, call, states, capture, is_empty, index_of, option_base, lookup, scalar_local, key_of
                    )
                    if raised:
                        hit = replayed_diagnostic(node.span, toks, lower, call, raised[0])
                        push(hit.rule, hit.message, hit.span)
                    if raised or lower not in states:
                        _forget_collection(states, lower)
                        break
            return
        _check_statement(node.span, toks, states, push, is_empty, index_of, option_base, lookup, scalar_local, key_of)

    # What a filling loop leaves, by id(loop): the states after it, and the other
    # names for those collections, which are not followed past it.
    leaves: dict[int, tuple[BodyNode, dict[str, _CollectionContents], list[str]]] = {}

    def enter_with(node: BodyNode) -> None:
        header = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
        name = token_name(header[1]) if len(header) == 2 else None
        element = (
            token_name(header[1])
            if len(header) >= 5 and header[2].raw_text == "(" and match_paren_from(header, 2) == len(header) - 1
            else None
        )
        if name and name.lower() in states:
            with_subjects.append(name)
        elif element and element.lower() in states:
            # `With c(1)` on a collection whose element is a collection: `.Item(2)`
            # reads as `c(1).Item(2)` (issue #295).
            with_subjects.append(list(header[1:]))
        elif len(header) == 3 and token_text(header[1]) == "new" and token_text(header[2]) == "collection":
            states[_NEW_WITH_SUBJECT.lower()] = _empty_contents()
            with_subjects.append(_NEW_WITH_SUBJECT)
        else:
            with_subjects.append(None)

    def snapshot() -> dict[str, _CollectionContents]:
        return _clone_states(states)

    def restore(saved: dict[str, _CollectionContents]) -> None:
        states.clear()
        states.update(_clone_states(saved))

    def forget(names: AbstractSet[str]) -> None:
        for lower in names:
            _forget_collection(states, lower)

    # `With c` reads c and changes nothing; its body's lines say what they do, a
    # `.Add` among them naming the subject (issue #584).
    def touches(stmt: LeafStatementNode) -> set[str]:
        toks = statement_tokens_after_leading_label(source, stmt.span)
        # `With c(1)` reads one element, by a literal, and changes nothing.
        if token_text(_token_at(toks, 0)) == "with" and (
            len(toks) == 2
            or (
                len(toks) == 5
                and toks[2].raw_text == "("
                and toks[3].kind in (TokenKind.INTEGER_LITERAL, TokenKind.STRING_LITERAL)
                and toks[4].raw_text == ")"
            )
        ):
            return set()
        names = names_in(source, stmt.span)
        within = current_subject()
        return {*names, _subject_name(within)} if within is not None and _reaches_subject(toks, within) else names

    # A counted loop that removes or reads by its counter (issue #263), or fills
    # or empties a collection (issue #350).
    def enter(node: BodyNode) -> None:
        if isinstance(node, WithBlockNode):
            enter_with(node)
        _check_scalar_elements(source, node, states, local_symbols, push)
        _simulate_counted_loop(source, node, states, push, activity)
        after = (
            None
            if unreachable is not None and id(node) in unreachable
            else _simulate_filling_loop(source, node, states, push, activity)
        )
        if after is not None:
            # Another name for the same collection, `Set o = c` above, is not
            # followed past the loop.
            aliases = [
                lower
                for lower, contents in states.items()
                if lower not in after and any(states.get(name) is contents for name in after)
            ]
            leaves[id(node)] = (node, after, aliases)

    def exit_block(node: BodyNode) -> None:
        if isinstance(node, WithBlockNode):
            popped = with_subjects.pop() if with_subjects else None
            if isinstance(popped, str) and popped == _NEW_WITH_SUBJECT:
                _forget_collection(states, _NEW_WITH_SUBJECT.lower())
        left = leaves.get(id(node))
        if left is not None:
            for lower, contents in left[1].items():
                states[lower] = contents
            for lower in left[2]:
                _forget_collection(states, lower)

    walk_entering_blocks(
        source,
        member.body,
        lambda node: activity is not None and activity.is_inactive(node.span),
        visit,
        BlockEnteringState(
            snapshot=snapshot,
            restore=restore,
            forget=forget,
            touches=touches,
            enter=enter,
            exit=exit_block,
            with_body_runs_through=True,
        ),
    )


def replayed_calls(
    toks: Sequence[VbaToken],
    states: Mapping[str, object],
    callee_calls: CalleeMemberCalls,
) -> Mapping[str, Sequence[Sequence[VbaToken]]] | None:
    """The member calls a call statement's callee makes on the tracked objects
    passed to it, when each tracked name the statement mentions is one of them,
    passed whole. None otherwise, and the statement is judged as any other
    (XLIDE issue #685)."""
    mentioned = [
        tok
        for i, tok in enumerate(toks)
        if (_lower_name(tok) or "") in states and _raw_text_at(toks, i - 1) != "."
    ]
    if len(mentioned) == 0:
        return None
    replays = callee_calls(toks)
    return replays if all((_lower_name(tok) or "") in replays for tok in mentioned) else None


@dataclass(frozen=True, slots=True)
class Replayed:
    """What a replayed callee statement raised."""

    rule: str
    message: str


@dataclass(frozen=True, slots=True)
class ReplayedHit:
    """A Replayed, with the span of the call it is reported at."""

    rule: str
    message: str
    span: Span


def replayed_diagnostic(
    base: Span, toks: Sequence[VbaToken], lower: str, call: Sequence[VbaToken], raised: Replayed
) -> ReplayedHit:
    """A callee statement that raises, reported at the call on the object passed
    to it: `AddK d` twice, AddK doing `p.Add "k", 1`, raises 457 the second time
    (XLIDE issue #685, measured in Excel 16.0)."""
    at = 1 if token_text(_token_at(toks, 0)) == "call" else 0
    arg = next((tok for i, tok in enumerate(toks) if i > at and _lower_name(tok) == lower), toks[at])
    text = ""
    for i, tok in enumerate(call):
        tight = (
            i == 0
            or tok.raw_text in (".", ",", ")")
            or call[i - 1].raw_text == "."
            or call[i - 1].raw_text == "("
        )
        text += ("" if tight else " ") + tok.raw_text
    return ReplayedHit(
        raised.rule,
        f"'{toks[at].raw_text}' runs {text} here. {raised.message}",
        Span(base.start + arg.start, base.start + arg.end),
    )


def _subject_name(subject: _WithSubject) -> str:
    """The lowercased local a With subject is, or is an element of."""
    return (subject if isinstance(subject, str) else subject[0].raw_text).lower()


def _reaches_subject(toks: Sequence[VbaToken], subject: _WithSubject) -> bool:
    """Whether a statement inside `With subject` reaches the subject by a leading dot."""
    significant = [tok for tok in toks if tok.kind is not TokenKind.COMMENT]
    return len(with_receiver(significant, subject)) > len(significant)


def _check_scalar_elements(
    source: str,
    node: BodyNode,
    states: Mapping[str, _CollectionContents],
    locals_: Sequence[VbaSymbol],
    push: PushFn,
) -> None:
    """`For Each v In c` with every element of c a number or a string: v holds
    one, and the body's first line to name v, `v Is Nothing` or `v.Count`, raises
    424 (XLIDE issue #612, measured in Excel 16.0)."""
    if (
        not isinstance(node, ForBlockNode)
        or not node.each
        or not node.control_variable
        or not node.source_expression
    ):
        return
    lower = node.control_variable.lower()
    source_expression = js_trim(node.source_expression)
    contents = states.get(source_expression.lower())
    local = next((child for child in locals_ if child.name.lower() == lower), None)
    type_ = normalize_type(local.as_type if local is not None else None)
    if (
        contents is None
        or contents.stale
        or len(contents.held) == 0
        or not all(isinstance(held, str) for held in contents.held)
        or local is None
        or local.kind is not VbaSymbolKind.LOCAL_VARIABLE
        or (type_ is not None and type_ != "variant")
    ):
        return
    # `new RegExp('\\b' + lower + '\\b', 'i')`: a word boundary in JavaScript's
    # ASCII sense.
    mention = re.compile(
        r"(?<![A-Za-z0-9_])" + re.escape(lower) + r"(?![A-Za-z0-9_])", re.IGNORECASE
    )
    for child in node.body:
        if not is_leaf_statement(child):
            if mention.search(source[child.span.start : child.span.end]) is not None:
                return
            continue
        toks = statement_tokens_after_leading_label(source, child.span)
        at = next(
            (i for i, tok in enumerate(toks) if _lower_name(tok) == lower and _raw_text_at(toks, i - 1) != "."),
            -1,
        )
        if at < 0:
            continue
        is_operand = (
            token_text(_token_at(toks, at + 1)) == "is" and token_text(_token_at(toks, at - 1)) != "typeof"
        ) or (token_text(_token_at(toks, at - 1)) == "is" and token_text(_token_at(toks, at - 2)) != "typeof")
        has_member = _raw_text_at(toks, at + 1) == "." and token_name(_token_at(toks, at + 2)) is not None
        if is_operand or has_member:
            push(
                "variantValueMisuse",
                f"'{toks[at].raw_text}' holds an element of '{source_expression}', each a number or a string, "
                f"not an object{' with members' if has_member else ' for Is to compare'}. This will raise "
                "Run-time error '424': Object required.",
                Span(child.span.start + toks[at].start, child.span.start + toks[at].end),
            )
        return


def with_receiver(toks: Sequence[VbaToken], subject: _WithSubject) -> list[VbaToken]:
    """A statement's tokens with the With subject before each member the block
    reaches by a leading dot: `.Item(2)` reads as `c.Item(2)`."""
    out: list[VbaToken] = []
    for i, tok in enumerate(toks):
        before = toks[i - 1] if i > 0 else None
        leading = tok.raw_text == "." and (
            before is None
            or (
                before.kind is not TokenKind.IDENTIFIER
                and before.kind is not TokenKind.BRACKETED_IDENTIFIER
                and before.raw_text != ")"
                and not (before.kind is TokenKind.KEYWORD and token_text(before) == "me")
            )
        )
        if leading and token_name(_token_at(toks, i + 1)):
            if isinstance(subject, str):
                out.append(
                    replace(tok, kind=TokenKind.IDENTIFIER, raw_text=subject, end=tok.start, canonical_text=None)
                )
            else:
                out.extend(replace(part, start=tok.start, end=tok.start) for part in subject)
        out.append(tok)
    return out


# The most passes a counted loop is run for.
_MAX_SIMULATED_PASSES = 10000

_LEAVING_HEADS = frozenset({"exit", "goto", "gosub", "resume", "return", "end", "on", "stop"})


@dataclass(frozen=True, slots=True)
class _Use:
    name: str
    display: str
    arg: Sequence[VbaToken]
    base: int
    removes: bool


def _simulate_counted_loop(
    source: str,
    node: BodyNode,
    states: Mapping[str, _CollectionContents],
    push: PushFn,
    activity: ConditionalActivityTracker | None,
) -> None:
    """`For i = 1 To c.Count: c.Remove i: Next` on three elements removes 1 and 2,
    then finds no element 3 (XLIDE issue #263, measured in Excel 16.0: error 9;
    and error 5 once the collection is empty). A For loop whose bounds the
    contents decide, and whose body is plain statements that touch the collection
    only by `c.Remove k` and `c(k)`, k the counter, a whole number or the counter
    plus or minus one, is run pass by pass. Anything else in the body that names
    the collection, writes the counter or may leave the pass stops it."""
    if not isinstance(node, ForBlockNode) or node.each or not node.control_variable or len(states) == 0:
        return
    counter = node.control_variable.lower()
    header = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
    eq = next((k for k, tok in enumerate(header) if tok.raw_text == "="), -1)
    to = next((k for k, tok in enumerate(header) if token_text(tok) == "to"), -1)
    step_at = next((k for k, tok in enumerate(header) if token_text(tok) == "step"), -1)

    def bound(toks: Sequence[VbaToken]) -> int | None:
        literal = _literal_index(toks)
        if literal is not None:
            return literal
        # `c.Count`, `c.Count - 1`
        name = _lower_name(_token_at(toks, 0))
        contents = states.get(name) if name else None
        if contents is None or _raw_text_at(toks, 1) != "." or token_text(_token_at(toks, 2)) != "count":
            return None
        if len(toks) == 3:
            return len(contents.items)
        offset = (
            _literal_index([toks[4]]) if len(toks) == 5 and toks[3].raw_text in ("+", "-") else None
        )
        return None if offset is None else len(contents.items) + (-offset if toks[3].raw_text == "-" else offset)

    start = bound(header[eq + 1 : to]) if eq > 0 and to > eq else None
    limit = bound(header[to + 1 : step_at if step_at > 0 else len(header)]) if to > 0 else None
    step = _literal_index(header[step_at + 1 :]) if step_at > 0 else 1
    if start is None or limit is None or step is None or step == 0:
        return
    uses: list[_Use] = []
    for stmt in node.body:
        if activity is not None and activity.is_inactive(stmt.span):
            continue
        if not is_leaf_statement(stmt) or (
            isinstance(stmt, StatementNode) and stmt.single_line_if_branches is not None
        ):
            return
        toks = statement_tokens_after_leading_label(source, stmt.span)
        head = token_text(_token_at(toks, 0))
        if head in _LEAVING_HEADS or jump_target_label_declaration(source, stmt.span):
            return
        # The counter only read: an operand, a whole collection index, or the
        # whole Remove argument. Assigned, passed or printed, it is not followed.
        for i, tok in enumerate(toks):
            if _lower_name(tok) != counter or _raw_text_at(toks, i - 1) == ".":
                continue
            following = _token_at(toks, i + 1)
            operand = (i > 0 and toks[i - 1].kind is TokenKind.OPERATOR) or (
                following is not None
                and following.kind is TokenKind.OPERATOR
                and not (i == 0 and following.raw_text == "=")
            )
            indexes = (
                _raw_text_at(toks, i - 1) == "("
                and _raw_text_at(toks, i + 1) == ")"
                and (_lower_name(_token_at(toks, i - 2)) or "") in states
            )
            removes = (
                i == 3
                and len(toks) == 4
                and token_text(toks[2]) == "remove"
                and (_lower_name(toks[0]) or "") in states
            )
            if not operand and not indexes and not removes:
                return
        i = 0
        while i < len(toks):
            lower = _lower_name(toks[i])
            if not lower or _raw_text_at(toks, i - 1) == "." or lower not in states:
                i += 1
                continue
            if i == 0 and _raw_text_at(toks, 1) == "." and token_text(_token_at(toks, 2)) == "remove" and len(toks) > 3:
                uses.append(_Use(lower, toks[0].raw_text, toks[3:], stmt.span.start, True))
                break
            if _raw_text_at(toks, i + 1) == "(":
                open_index = i + 1
            elif (
                _raw_text_at(toks, i + 1) == "."
                and token_text(_token_at(toks, i + 2)) == "item"
                and _raw_text_at(toks, i + 3) == "("
            ):
                open_index = i + 3
            else:
                return  # Add, Count after a change, a pass or a Set: not followed
            close = match_paren_from(toks, open_index)
            if close < 0:
                # Upstream sets its index to -1 here and rescans the statement from
                # the start, forever; an unclosed paren is not followed instead.
                return
            uses.append(_Use(lower, toks[i].raw_text, toks[open_index + 1 : close], stmt.span.start, False))
            i = close + 1
    if len(uses) == 0:
        return

    def index_at(arg: Sequence[VbaToken], value: int) -> int | None:
        literal = _literal_index(arg)
        if literal is not None:
            return literal
        if _lower_name(_token_at(arg, 0)) != counter:
            return None
        if len(arg) == 1:
            return value
        offset = _literal_index([arg[2]]) if len(arg) == 3 and arg[1].raw_text in ("+", "-") else None
        return None if offset is None else value + (-offset if arg[1].raw_text == "-" else offset)

    if any(index_at(use.arg, start) is None for use in uses):
        return
    counts: dict[str, int] = {}
    for use in uses:
        if use.name not in counts:
            counts[use.name] = len(states[use.name].items)
    passes = 0
    value = start
    while value <= limit if step > 0 else value >= limit:
        passes += 1
        if passes > _MAX_SIMULATED_PASSES:
            return
        for use in uses:
            count = counts[use.name]
            index = index_at(use.arg, value)
            assert index is not None
            if 1 <= index <= count:
                if use.removes:
                    counts[use.name] = count - 1
                continue
            if passes == 1 and _literal_index(use.arg) is not None:
                return  # the walk into the block reports the first pass
            span = Span(use.base + use.arg[0].start, use.base + use.arg[-1].end)
            where = f"On the pass of the For loop where '{node.control_variable}' is {value}"
            if count == 0:
                message = (
                    f"{where}, '{use.display}' holds nothing, so no index reaches an element. This will raise "
                    "Run-time error '5': Invalid procedure call or argument."
                )
            else:
                message = (
                    f"{where}, '{use.display}' holds {count} element{'' if count == 1 else 's'}, indexed 1 to "
                    f"{count}; {index} is outside that. This will raise Run-time error '9': Subscript out of range."
                )
            push("collectionIndexOutOfRange", message, span)
            return
        value += step


@dataclass(frozen=True, slots=True)
class _Added:
    key: str | None
    key_toks: Sequence[VbaToken] | None
    key_built: bool
    key_is_variable: bool


@dataclass(frozen=True, slots=True)
class _Change:
    name: str
    display: str
    base: int
    add: _Added | None = None
    remove: int | None = None


def _simulate_filling_loop(
    source: str,
    node: BodyNode,
    states: Mapping[str, _CollectionContents],
    push: PushFn,
    activity: ConditionalActivityTracker | None,
) -> dict[str, _CollectionContents] | None:
    """`For i = 1 To 3: c.Add i: Next` (XLIDE issue #350, measured in Excel 16.0): a
    For loop with literal bounds whose body is plain statements, each either
    `c.Add item[, key]` or `c.Remove n` on a tracked collection or one that names
    none, is run pass by pass. A literal key added on a second pass raises 457
    there. What the loop leaves, an empty collection after a loop of no pass
    included, is returned for after it; a key the code builds leaves the keys
    unknown. None when the loop cannot be followed."""
    if not isinstance(node, ForBlockNode) or not node.control_variable or len(states) == 0:
        return None
    counter = node.control_variable.lower()
    # The values the loop variable takes, pass by pass: a counted For's, or the
    # elements of a literal Split or Array a For Each steps through.
    values: list[int | str] = []
    if node.each:
        elements = _literal_elements(node.source_expression or "")
        if elements is None:
            return None
        values.extend(elements)
    else:
        header = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
        eq = next((k for k, tok in enumerate(header) if tok.raw_text == "="), -1)
        to = next((k for k, tok in enumerate(header) if token_text(tok) == "to"), -1)
        step_at = next((k for k, tok in enumerate(header) if token_text(tok) == "step"), -1)
        start = _literal_index(header[eq + 1 : to]) if eq > 0 and to > eq else None
        limit = _literal_index(header[to + 1 : step_at if step_at > 0 else len(header)]) if to > 0 else None
        step = _literal_index(header[step_at + 1 :]) if step_at > 0 else 1
        if start is None or limit is None or not step:
            return None
        value = start
        while value <= limit if step > 0 else value >= limit:
            if len(values) >= _MAX_SIMULATED_PASSES:
                return None
            values.append(value)
            value += step
    changes: list[_Change] = []
    for stmt in node.body:
        if activity is not None and activity.is_inactive(stmt.span):
            continue
        if (
            not is_leaf_statement(stmt)
            or (isinstance(stmt, StatementNode) and stmt.single_line_if_branches is not None)
            or jump_target_label_declaration(source, stmt.span)
        ):
            return None
        toks = statement_tokens_after_leading_label(source, stmt.span)
        if token_text(_token_at(toks, 0)) in _LEAVING_HEADS:
            return None
        name = _lower_name(_token_at(toks, 0))
        mentions = any(
            (_lower_name(tok) or "") in states and _raw_text_at(toks, k - 1) != "." for k, tok in enumerate(toks)
        )
        if not name or name not in states or _raw_text_at(toks, 1) != ".":
            if mentions or (len(toks) > 0 and _lower_name(toks[0]) == counter):
                return None
            continue
        member = token_text(_token_at(toks, 2))
        args = _arguments_after(toks, 3)
        if member == "add" and 1 <= len(args) <= 2 and not any(_raw_text_at(arg, 1) == ":=" for arg in args):
            key_toks = args[1] if len(args) > 1 else None
            key = _literal_key(key_toks) if key_toks is not None else None
            # `c.Add p, p` in a For Each: the key is the element of the pass.
            key_is_variable = (
                node.each and key_toks is not None and len(key_toks) == 1 and _lower_name(key_toks[0]) == counter
            )
            changes.append(
                _Change(
                    name,
                    toks[0].raw_text,
                    stmt.span.start,
                    add=_Added(
                        key,
                        key_toks,
                        not key_is_variable and key_toks is not None and len(key_toks) > 0 and key is None,
                        key_is_variable,
                    ),
                )
            )
            continue
        index = _literal_index(args[0]) if member == "remove" and len(args) == 1 else None
        if index is None:
            return None
        changes.append(_Change(name, toks[0].raw_text, stmt.span.start, remove=index))
    if len(changes) == 0:
        return None
    after: dict[str, _CollectionContents] = {}
    for change in changes:
        if change.name not in after:
            contents = states[change.name]
            after[change.name] = _CollectionContents(
                list(contents.items), contents.keys_known, list(contents.shapes), list(contents.held)
            )
    passes = 0
    for pass_value in values:
        passes += 1
        for change in changes:
            contents = after[change.name]
            if change.remove is not None:
                if change.remove < 1 or change.remove > len(contents.items):
                    return None  # the walk into the block reports a first pass
                del contents.items[change.remove - 1]
                del contents.shapes[change.remove - 1 : change.remove]
                del contents.held[change.remove - 1 : change.remove]
                continue
            added = change.add
            assert added is not None
            key = _js_string_of(pass_value).lower() if added.key_is_variable else added.key
            if key is not None and contents.keys_known and key in contents.items:
                if passes > 1 and added.key_toks is not None:
                    added_key = added.key_toks
                    shown = (
                        json.dumps(pass_value, ensure_ascii=False)
                        if isinstance(pass_value, str)
                        else _js_string(pass_value)
                    )
                    push(
                        "collectionKeyInUse",
                        f"On the pass of the For loop where '{node.control_variable}' is {shown}, "
                        f"'{change.display}' already has an element with the key {added_key[0].raw_text} from an "
                        "earlier pass. This will raise Run-time error '457': This key is already associated "
                        "with an element of this collection.",
                        Span(change.base + added_key[0].start, change.base + added_key[-1].end),
                    )
                return None
            contents.items.append(key)
            contents.shapes.append(None)
            contents.held.append(None)
            if added.key_built:
                contents.keys_known = False
    return after


def _literal_elements(expression: str) -> list[str] | None:
    """The elements of `Split("a,b", ",")` or `Array("a", "b")` written with literals, as Strings (issue #350)."""
    toks = [tok for tok in raw_expression_tokens(expression) if tok.kind is not TokenKind.COMMENT]
    callee = token_text(_token_at(toks, 0))
    if (callee != "split" and callee != "array") or _raw_text_at(toks, 1) != "(" or match_paren_from(
        toks, 1
    ) != len(toks) - 1:
        return None
    args = _arguments_after(toks, 1)
    if not all(len(arg) == 1 and arg[0].kind is TokenKind.STRING_LITERAL for arg in args):
        return None
    texts = [string_literal_value(arg[0].raw_text) for arg in args]
    if callee == "array":
        return texts
    if len(texts) < 1 or len(texts) > 2 or (len(texts) > 1 and texts[1] == "") or texts[0] == "":
        return None
    return texts[0].split(texts[1] if len(texts) > 1 else " ")


def _clone_states(states: Mapping[str, _CollectionContents]) -> dict[str, _CollectionContents]:
    """A copy of the states in which two names that shared one collection still do.
    An element that is a collection is copied once, as the name sharing it is."""
    copies: dict[int, tuple[_CollectionContents, _CollectionContents]] = {}
    pending: list[_CollectionContents] = []

    def copy_of(contents: _CollectionContents) -> _CollectionContents:
        entry = copies.get(id(contents))
        if entry is not None:
            return entry[1]
        copy = _CollectionContents(
            list(contents.items), contents.keys_known, list(contents.shapes), [], contents.stale
        )
        copies[id(contents)] = (contents, copy)
        pending.append(contents)
        return copy

    out = {lower: copy_of(contents) for lower, contents in states.items()}
    # The held elements are copied from a work list, not by recursion.
    while pending:
        original = pending.pop()
        copies[id(original)][1].held = [
            copy_of(held) if isinstance(held, _CollectionContents) else held for held in original.held
        ]
    return out


@dataclass(frozen=True, slots=True)
class _CollectionLocals:
    # Lowercased names declared `As New Collection`: a collection from the start.
    new_locals: set[str]
    # Lowercased names declared `As Collection` (or `As Object`): tracked once
    # `Set x = New Collection`.
    plain_locals: set[str]
    # Lowercased Variant locals.
    variant_locals: set[str]


_TYPE_SUFFIX_NAME = re.compile(r"[%&^!#@$]\Z")


def _collection_locals(proc: ProcedureNode, activity: ConditionalActivityTracker | None) -> _CollectionLocals:
    new_locals: set[str] = set()
    plain_locals: set[str] = set()
    variant_locals: set[str] = set()

    def visit(group: VariableGroupNode) -> None:
        if group.is_const or group.modifier.lower() == "static":
            return
        for decl in group.declarations:
            type_ = normalize_type(decl.as_type)
            if (
                not decl.is_array
                and (type_ is None or type_ == "variant")
                and _TYPE_SUFFIX_NAME.search(decl.name) is None
            ):
                variant_locals.add(decl.name.lower())
            # An Object is followed from `Set o = New Collection` on, as a
            # Collection is: `o(1)` on it empty raises 5 (issue #415).
            if decl.is_array or (type_ != "collection" and not (type_ == "object" and not decl.is_new)):
                continue
            (new_locals if decl.is_new else plain_locals).add(decl.name.lower())

    for_each_variable_group(proc.body, visit, activity)
    return _CollectionLocals(new_locals, plain_locals, variant_locals)


def _forget_mentioned(source: str, span: Span, states: dict[str, _CollectionContents]) -> None:
    """Drops every tracked collection a statement names anywhere."""
    for tok in statement_tokens_after_leading_label(source, span):
        lower = _lower_name(tok)
        if lower and lower in states:
            _forget_collection(states, lower)


_IndexOf = Callable[[Sequence[VbaToken]], "int | float | None"]
# The key an argument names: a string literal, or a String local known to hold one (issue #346).
_KeyOf = Callable[[Sequence[VbaToken]], "str | None"]


def _never_empty(_lower: str) -> bool:
    return False


def _no_scalar(_lower: str) -> bool:
    return False


def _check_statement(
    base: Span,
    toks: Sequence[VbaToken],
    states: dict[str, _CollectionContents],
    push: PushFn,
    is_empty: Callable[[str], bool],
    index_of: _IndexOf | None = None,
    option_base: int = 0,
    lookup: IntegerConstantLookup | None = None,
    scalar_local: Callable[[str], bool] = _no_scalar,
    key_of: _KeyOf | None = None,
) -> None:
    index_reader: _IndexOf = index_of if index_of is not None else _literal_index
    key_reader: _KeyOf = key_of if key_of is not None else _literal_key_text

    def at(start: int, end: int) -> Span:
        return Span(base.start + toks[start].start, base.start + toks[end].end)

    # First pass: reads and the recognised forms, in source order. A mention in
    # any other shape ends tracking of that variable after this statement.
    to_forget: list[str] = []
    mutations: list[Callable[[], None]] = []
    i = 1 if token_text(_token_at(toks, 0)) == "call" else 0
    first = i

    # `c.Add inner`: the element shares inner, which the Add leaves as it is.
    def held_of(item: Sequence[VbaToken]) -> _Held:
        named = _lower_name(item[0]) if len(item) == 1 else None
        contents = states.get(named) if named else None
        if contents is not None:
            return contents
        # `c.Add New Collection`: an empty Collection no name holds (issue #306).
        if len(item) == 2 and token_text(item[0]) == "new" and token_text(item[1]) == "collection":
            return _empty_contents()
        if _literal_index(item) is not None or (len(item) == 1 and item[0].kind is TokenKind.FLOAT_LITERAL):
            return "number"
        return "string" if len(item) == 1 and item[0].kind is TokenKind.STRING_LITERAL else None

    adds_whole = (
        _arguments_after(toks, first + 3)
        if (_lower_name(_token_at(toks, first)) or "") in states
        and _raw_text_at(toks, first + 1) == "."
        and token_text(_token_at(toks, first + 2)) == "add"
        else None
    )
    added_item = adds_whole[0] if adds_whole else None
    added_name = added_item[0] if added_item is not None and len(added_item) == 1 and token_name(added_item[0]) else None
    item_ctx = _ItemContext(
        base, toks, push, index_reader, mutations, is_empty, option_base, held_of, scalar_local, first
    )

    def forget_later(lower: str) -> None:
        if lower not in to_forget:
            to_forget.append(lower)

    while i < len(toks):
        lower = _lower_name(toks[i])
        if not lower or lower not in states or _raw_text_at(toks, i - 1) == "." or toks[i] is added_name:
            i += 1
            continue
        state = states[lower]
        following = _raw_text_at(toks, i + 1)
        # `c(index)` or `c("key")`
        if following == "(":
            close = match_paren_from(toks, i + 1)
            if close > i + 2 and _check_read(
                lower, state, toks[i + 2 : close], at(i + 2, close - 1), push, index_reader, key_reader
            ):
                _check_item_array(
                    base, toks, toks[i].raw_text, state, toks[i + 2 : close], close, push, index_reader, lookup
                )
                _use_held_item(item_ctx, i, toks[i].raw_text, state, toks[i + 2 : close], close)
                i += 1
                continue
            forget_later(lower)
            i += 1
            continue
        if following != ".":
            forget_later(lower)
            i += 1
            continue
        member_name = token_text(_token_at(toks, i + 2))
        if member_name == "count":
            i += 1
            continue
        if member_name == "item":
            item_close = match_paren_from(toks, i + 3) if _raw_text_at(toks, i + 3) == "(" else -1
            if item_close > i + 4 and _check_read(
                lower, state, toks[i + 4 : item_close], at(i + 4, item_close - 1), push, index_reader, key_reader
            ):
                _check_item_array(
                    base, toks, toks[i].raw_text, state, toks[i + 4 : item_close], item_close, push,
                    index_reader, lookup,
                )
                _use_held_item(
                    item_ctx, i, f"{toks[i].raw_text}.Item", state, toks[i + 4 : item_close], item_close
                )
                i += 1
                continue
            forget_later(lower)
            i += 1
            continue
        if member_name in ("add", "remove") and i == (1 if token_text(_token_at(toks, 0)) == "call" else 0):
            args = _arguments_after(toks, i + 3)
            if member_name == "add":
                mutations.append(
                    _bind_add(lower, state, args, base, push, is_empty, index_reader, option_base, held_of)
                )
            elif len(args) == 1:
                mutations.append(_bind_remove(lower, state, args[0], base, push, index_reader))
            else:
                forget_later(lower)
            i += 1
            continue
        forget_later(lower)
        i += 1
    for mutation in mutations:
        mutation()
    for lower in to_forget:
        _forget_collection(states, lower)


def _bind_add(
    name: str,
    state: _CollectionContents,
    args: list[list[VbaToken]],
    base: Span,
    push: PushFn,
    is_empty: Callable[[str], bool],
    index_of: _IndexOf,
    option_base: int,
    held_of: Callable[[Sequence[VbaToken]], _Held],
) -> Callable[[], None]:
    """An Add deferred with its own arguments bound now, as upstream's closure captures them."""
    return lambda: _add(name, state, args, base, push, is_empty, index_of, option_base, held_of)


def _bind_remove(
    name: str, state: _CollectionContents, arg: list[VbaToken], base: Span, push: PushFn, index_of: _IndexOf
) -> Callable[[], None]:
    return lambda: _remove(name, state, arg, base, push, index_of)


@dataclass(frozen=True, slots=True)
class _ItemContext:
    base: Span
    toks: Sequence[VbaToken]
    push: PushFn
    index_of: _IndexOf
    mutations: list[Callable[[], None]]
    is_empty: Callable[[str], bool]
    option_base: int
    held_of: Callable[[Sequence[VbaToken]], _Held]
    scalar_local: Callable[[str], bool]
    # The statement's first token after a Call.
    first: int


def _use_held_item(
    ctx: _ItemContext,
    name_at: int,
    display: str,
    state: _CollectionContents,
    arg: Sequence[VbaToken],
    close: int,
) -> None:
    """What follows `c(k)` where element k is a Collection the code added, or a
    number (XLIDE issue #452, measured in Excel 16.0): an index or key into the
    inner Collection, and its Add and Remove, are judged against what it holds; a
    number indexed raises 13 and a member of one 424; a Collection Let into a
    typed local raises 450, its default member Item needing an index. Any other
    use of an inner Collection may change it unseen."""
    toks = ctx.toks
    base = ctx.base
    key = _literal_key(arg)
    index = ctx.index_of(arg)
    if index is None and key is not None and state.keys_known and key in state.items:
        index = state.items.index(key) + 1
    held: _Held = _element_at(state.held, index) if index is not None and index >= 1 else None
    if held is None or (isinstance(held, _CollectionContents) and held.stale):
        return
    shown = f"{display}({''.join(tok.raw_text for tok in arg)})"

    def span_of(start: int, end: int) -> Span:
        # Upstream reads toks[end] unguarded; a missing close paren throws there.
        if end < 0:
            raise IndexError("no closing parenthesis")
        return Span(base.start + toks[start].start, base.start + toks[end].end)

    after = _raw_text_at(toks, close + 1)
    member = token_text(_token_at(toks, close + 2)) if after == "." else None
    # `Set c(1) = o`: Item has no Property Set, so the Set reaches what the item
    # holds: 424 on a value, 438 on an object (issue #306, measured in Excel 16.0).
    if token_text(_token_at(toks, 0)) == "set" and name_at == 1 and after == "=":
        if isinstance(held, str):
            ctx.push(
                "variantValueMisuse",
                f"'{shown}' holds a {held}, and a Collection's item cannot be replaced in place: Item has no "
                f"Property Set, so the Set reaches the {held}. This will raise Run-time error '424': Object "
                "required.",
                span_of(name_at, close),
            )
        else:
            ctx.push(
                "runtimeMemberNotFound",
                f"'{shown}' is read through Item, which a Set cannot write: a Collection's item cannot be "
                "replaced in place. This will raise Run-time error '438': Object doesn't support this property "
                "or method.",
                span_of(name_at, close),
            )
        return
    if isinstance(held, str):
        # `c(1) = 5`: a Let into the item, which only an object's default member
        # could take (issue #305, measured in Excel 16.0).
        if after == "=" and name_at == ctx.first:
            ctx.push(
                "variantValueMisuse",
                f"'{shown}' holds a {held}, and a Collection's item cannot be replaced in place: only an object "
                "item takes a value through its default member. This will raise Run-time error '424': Object "
                "required.",
                span_of(name_at, close),
            )
        elif after == "(":
            ctx.push(
                "variantValueMisuse",
                f"'{shown}' holds a {held}, which takes no index. This will raise Run-time error '13': Type "
                "mismatch.",
                span_of(name_at, match_paren_from(toks, close + 1)),
            )
        elif member:
            ctx.push(
                "variantValueMisuse",
                f"'{shown}' holds a {held}, not an object, so it has no {toks[close + 2].raw_text}. This will "
                "raise Run-time error '424': Object required.",
                span_of(name_at, close + 2),
            )
        return
    if after == "(" or (member == "item" and _raw_text_at(toks, close + 3) == "("):
        open_index = close + 1 if after == "(" else close + 3
        inner_close = match_paren_from(toks, open_index)
        if inner_close > open_index + 1:
            _check_read(
                shown, held, toks[open_index + 1 : inner_close], span_of(open_index + 1, inner_close - 1),
                ctx.push, ctx.index_of,
            )
        return
    if member == "count":
        return
    if member in ("add", "remove") and name_at == ctx.first:
        args = _arguments_after(toks, close + 3)
        if member == "add":
            ctx.mutations.append(
                _bind_add(shown, held, args, base, ctx.push, ctx.is_empty, ctx.index_of, ctx.option_base, ctx.held_of)
            )
        elif len(args) == 1:
            ctx.mutations.append(_bind_remove(shown, held, args[0], base, ctx.push, ctx.index_of))
        else:
            held.stale = True
        return
    if member is not None:
        held.stale = True
        return
    # `v = c(1)` with v a Long or a Variant.
    target = _lower_name(_token_at(toks, ctx.first))
    if (
        target
        and _raw_text_at(toks, ctx.first + 1) == "="
        and name_at == ctx.first + 2
        and close == len(toks) - 1
        and ctx.scalar_local(target)
    ):
        ctx.push(
            "objectDefaultValue",
            f"'{shown}' is a Collection: its default member Item needs an index, so it has no value for "
            f"'{toks[ctx.first].raw_text}' to take. This will raise Run-time error '450': Wrong number of "
            "arguments or invalid property assignment.",
            span_of(name_at, close),
        )


def _check_item_array(
    base: Span,
    toks: Sequence[VbaToken],
    name: str,
    state: _CollectionContents,
    arg: Sequence[VbaToken],
    close: int,
    push: PushFn,
    index_of: _IndexOf,
    lookup: IntegerConstantLookup | None,
) -> None:
    """`c(1)(5)`: the parentheses after an item read, against the array the item holds."""
    index = index_of(arg) if _raw_text_at(toks, close + 1) == "(" else None
    shape = _element_at(state.shapes, index) if index is not None else None
    hit = (
        shape_subscript_violation(base, toks, replace(shape, name=f"{name}({_js_string(index)})"), close + 1, lookup)
        if shape is not None and index is not None
        else None
    )
    if hit is not None:
        push(hit.rule if hit.rule is not None else "arraySubscriptOutOfBounds", hit.message, hit.span)


def _arguments_after(toks: Sequence[VbaToken], start: int) -> list[list[VbaToken]]:
    """The argument groups of a statement-form or parenthesised member call starting at `start`."""
    body = list(toks[start:])
    if len(body) > 0 and body[0].raw_text == "(" and match_paren_from(toks, start) == len(toks) - 1:
        body = body[1:-1]
    out: list[list[VbaToken]] = []
    current: list[VbaToken] = []
    depth = 0
    for tok in body:
        if tok.raw_text == "(":
            depth += 1
        elif tok.raw_text == ")":
            depth -= 1
        if tok.raw_text == "," and depth == 0:
            out.append(current)
            current = []
            continue
        current.append(tok)
    if len(current) > 0 or len(out) > 0:
        out.append(current)
    return out


def _literal_key(arg: Sequence[VbaToken]) -> str | None:
    if len(arg) == 1 and arg[0].kind is TokenKind.STRING_LITERAL:
        return string_literal_value(arg[0].raw_text).lower()
    return None


def _literal_index(arg: Sequence[VbaToken]) -> int | None:
    """The whole-number value of a literal argument, optionally negated."""
    if len(arg) == 1 and arg[0].kind is TokenKind.INTEGER_LITERAL:
        return parse_vba_integer_literal(arg[0].raw_text)
    if len(arg) == 2 and arg[0].raw_text == "-" and arg[1].kind is TokenKind.INTEGER_LITERAL:
        value = parse_vba_integer_literal(arg[1].raw_text)
        return None if value is None else -value
    return None


def _check_read(
    name: str,
    state: _CollectionContents,
    arg: Sequence[VbaToken],
    span: Span,
    push: PushFn,
    index_of: _IndexOf,
    key_of: _KeyOf | None = None,
) -> bool:
    """Judges `c(arg)` or `c.Item(arg)`; false when the argument is not a literal the rule reads."""
    index = index_of(arg)
    if index is not None:
        _report_index(name, state, index, span, push)
        return True
    key = (key_of if key_of is not None else _literal_key_text)(arg)
    if key is not None:
        _report_key(name, state, key, span, push)
        return True
    return False


def _literal_key_text(arg: Sequence[VbaToken]) -> str | None:
    return string_literal_value(arg[0].raw_text) if len(arg) == 1 and arg[0].kind is TokenKind.STRING_LITERAL else None


def _report_index(name: str, state: _CollectionContents, index: int | float, span: Span, push: PushFn) -> None:
    count = len(state.items)
    if count == 0:
        push(
            "collectionIndexOutOfRange",
            f"'{name}' holds nothing here, so no index reaches an element. "
            "This will raise Run-time error '5': Invalid procedure call or argument.",
            span,
        )
        return
    if index < 1 or index > count:
        push(
            "collectionIndexOutOfRange",
            f"'{name}' holds {count} element{'' if count == 1 else 's'} here, "
            f"indexed 1 to {count}; {_js_string(index)} is outside that. "
            "This will raise Run-time error '9': Subscript out of range.",
            span,
        )


def _report_key(name: str, state: _CollectionContents, key: str, span: Span, push: PushFn) -> bool:
    if len(state.items) == 0:
        push(
            "collectionKeyNotFound",
            f"'{name}' holds nothing here, so no key reaches an element. "
            "This will raise Run-time error '5': Invalid procedure call or argument.",
            span,
        )
        return True
    if state.keys_known and key.lower() not in state.items:
        push(
            "collectionKeyNotFound",
            f"No element of '{name}' was added with the key \"{key}\". "
            "This will raise Run-time error '5': Invalid procedure call or argument.",
            span,
        )
        return True
    return False


_ADD_PARAMETERS = ("item", "key", "before", "after")


def _add_arguments(args: Sequence[list[VbaToken]]) -> dict[str, list[VbaToken]] | None:
    """Add's arguments by parameter, positional or named; None when a name is unknown."""
    out: dict[str, list[VbaToken]] = {}
    for k, arg in enumerate(args):
        if _raw_text_at(arg, 1) == ":=":
            param = token_text(arg[0])
            if param not in _ADD_PARAMETERS:
                return None
            out[param] = arg[2:]
        elif len(arg) > 0 and k < len(_ADD_PARAMETERS):
            out[_ADD_PARAMETERS[k]] = arg
    return out


@dataclass(frozen=True, slots=True)
class _Refusal:
    rule: str  # 'collectionAddArgument' | 'collectionIndexOutOfRange'
    message: str
    span: Span


def _add_refusal(
    name: str,
    state: _CollectionContents,
    by_name: Mapping[str, list[VbaToken]],
    base: Span,
    is_empty: Callable[[str], bool],
    index_of: _IndexOf,
) -> _Refusal | None:
    """What Add refuses before it adds (XLIDE issue #219, measured in Excel 16.0):
    a key that is a number or True rather than a string raises 13; Before and
    After together raise 5; either one on an empty collection raises 5; and an
    index outside 1 to Count raises 9 (Before:=0, After:=2 with one element)."""

    def span_of(arg: Sequence[VbaToken]) -> Span:
        return Span(base.start + arg[0].start, base.start + arg[-1].end)

    key = by_name.get("key")
    key_literal = [t for t in key if t.kind is not TokenKind.COMMENT] if key is not None else []
    non_string = len(key_literal) > 0 and (
        _literal_index(key_literal) is not None
        or (len(key_literal) == 1 and key_literal[0].kind is TokenKind.FLOAT_LITERAL)
        or (len(key_literal) == 1 and token_text(key_literal[0]) in ("true", "false"))
    )
    empty_key = len(key_literal) == 1 and is_empty(_lower_name(key_literal[0]) or "")
    if empty_key:
        return _Refusal(
            "collectionAddArgument",
            f"The key of '{name}.Add' is '{key_literal[0].raw_text}', which is never assigned and so is Empty, "
            "not a string. This will raise Run-time error '13': Type mismatch.",
            span_of(key_literal),
        )
    if non_string:
        return _Refusal(
            "collectionAddArgument",
            f"The key of '{name}.Add' is {''.join(t.raw_text for t in key_literal)}, not a string. This will "
            "raise Run-time error '13': Type mismatch.",
            span_of(key_literal),
        )
    # `k = 5` then `c.Add 1, k`: a local known to hold a number (issue #349).
    held_number = index_of(key_literal) if len(key_literal) == 1 and token_name(key_literal[0]) else None
    if held_number is not None:
        return _Refusal(
            "collectionAddArgument",
            f"The key of '{name}.Add' is '{key_literal[0].raw_text}', which holds the number "
            f"{_js_string(held_number)}, not a string. This will raise Run-time error '13': Type mismatch.",
            span_of(key_literal),
        )
    before = by_name.get("before")
    after = by_name.get("after")
    if before is not None and after is not None:
        return _Refusal(
            "collectionAddArgument",
            f"'{name}.Add' is given both Before and After. This will raise Run-time error '5': Invalid procedure "
            "call or argument.",
            span_of(after),
        )
    position = before if before is not None else after
    if position is None:
        return None
    which = "Before" if before is not None else "After"
    if len(state.items) == 0:
        return _Refusal(
            "collectionIndexOutOfRange",
            f"'{name}' holds nothing here, so {which} names no element. This will raise Run-time error '5': "
            "Invalid procedure call or argument.",
            span_of(position),
        )
    # `c.Add 2, "b", "zz"`: Before names a key no element has (issue #349).
    position_key = _literal_key(position)
    if position_key is not None and state.keys_known and position_key not in state.items:
        return _Refusal(
            "collectionAddArgument",
            f"No element of '{name}' was added with the key {position[0].raw_text}, which {which} names. This "
            "will raise Run-time error '5': Invalid procedure call or argument.",
            span_of(position),
        )
    index = index_of(position)
    count = len(state.items)
    if index is not None and (index < 1 or index > count):
        return _Refusal(
            "collectionIndexOutOfRange",
            f"'{name}' holds {count} element{'' if count == 1 else 's'} here, indexed 1 to {count}; {which} is "
            f"{_js_string(index)}. This will raise Run-time error '9': Subscript out of range.",
            span_of(position),
        )
    return None


def _no_held(_item: Sequence[VbaToken]) -> _Held:
    return None


def _add(
    name: str,
    state: _CollectionContents,
    raw_args: list[list[VbaToken]],
    base: Span,
    push: PushFn,
    is_empty: Callable[[str], bool],
    index_of: _IndexOf,
    option_base: int,
    held_of: Callable[[Sequence[VbaToken]], _Held] = _no_held,
) -> None:
    by_name = _add_arguments(raw_args)
    if by_name is None:
        state.items.append(None)
        state.shapes.append(None)
        state.held.append(None)
        state.keys_known = False
        return
    refusal = _add_refusal(name, state, by_name, base, is_empty, index_of)
    if refusal is not None:
        push(refusal.rule, refusal.message, refusal.span)
        return
    args = [by_name.get(param, []) for param in _ADD_PARAMETERS]
    key_arg = args[1]
    key = _literal_key(key_arg) if len(key_arg) > 0 else None
    if len(key_arg) > 0 and key is None:
        state.keys_known = False
    if key is not None and state.keys_known and key in state.items:
        push(
            "collectionKeyInUse",
            f"'{name}' already has an element with the key \"{string_literal_value(key_arg[0].raw_text)}\" (keys "
            "compare without case). This will raise Run-time error '457': This key is already associated with an "
            "element of this collection.",
            Span(base.start + key_arg[0].start, base.start + key_arg[-1].end),
        )
        return
    if any(len(arg) > 0 for arg in args[2:]):
        # Before or After: the position is not followed, the count is.
        state.items.append(key)
        state.keys_known = False
        state.shapes = [None for _ in state.items]
        state.held = [None for _ in state.items]
        return
    state.items.append(key)
    state.shapes.append(array_value_shape(args[0], name, option_base) if len(args[0]) > 0 else None)
    state.held.append(held_of(args[0]))


def _remove(
    name: str, state: _CollectionContents, arg: list[VbaToken], base: Span, push: PushFn, index_of: _IndexOf
) -> None:
    span = Span(base.start + arg[0].start, base.start + arg[-1].end)
    index = index_of(arg)
    if index is not None:
        if len(state.items) == 0 or index < 1 or index > len(state.items):
            _report_index(name, state, index, span, push)
            return
        _splice_one(state, index - 1)
        return
    key = _literal_key(arg)
    if key is None:
        # A variable index or key: one element fewer, which one unknown.
        _forget_order(state)
        return
    if _report_key(name, state, string_literal_value(arg[0].raw_text), span, push):
        return
    if key in state.items:
        _splice_one(state, state.items.index(key))
    else:
        _forget_order(state)


def _splice_one(state: _CollectionContents, position: int | float) -> None:
    """`splice(position, 1)` on the items, shapes and held lists. A fractional
    position (an index rounded from a Double is whole; one from a lookup may not
    be) is truncated as splice truncates it."""
    at = int(position)
    del state.items[at : at + 1]
    del state.shapes[at : at + 1]
    del state.held[at : at + 1]


def _forget_order(state: _CollectionContents) -> None:
    """One element fewer, which one unknown: `items.pop()`, and the shapes and held
    elements no longer line up."""
    if state.items:
        state.items.pop()
    state.keys_known = False
    state.shapes = [None for _ in state.items]
    state.held = [None for _ in state.items]


def check_collection_loop_counters(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """A loop counter indexing a Collection outside 1 to Count (XLIDE issue #200,
    measured in Excel 16.0): `For i = 0 To c.Count - 1` reads c(0) on its first
    pass, and `For i = 1 To c.Count + 1` reads past the last element on its last.
    Collections are 1-based; an index outside raises 9, or 5 when the collection
    is empty. The contents are not tracked into a loop, so the error is 9 only
    where the loop's own bounds show it runs with an element."""
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        locals_ = _collection_locals(member, activity)
        collections = {*locals_.new_locals, *locals_.plain_locals}
        for param in member.params:
            if not param.is_array and normalize_type(param.as_type) == "collection":
                collections.add(param.name.lower())
        if len(collections) == 0:
            continue
        for stmt, counters in loop_counters_at(source, member.body, activity).items():
            toks = statement_tokens_after_leading_label(source, stmt.span)
            for i in range(len(toks) - 1):
                lower = _lower_name(toks[i])
                if not lower or lower not in collections or _raw_text_at(toks, i - 1) == ".":
                    continue
                # `c(i)` or `c.Item(i)`
                if toks[i + 1].raw_text == "(":
                    open_index = i + 1
                elif (
                    toks[i + 1].raw_text == "."
                    and token_text(_token_at(toks, i + 2)) == "item"
                    and _raw_text_at(toks, i + 3) == "("
                ):
                    open_index = i + 3
                else:
                    open_index = -1
                close = -1 if open_index < 0 else match_paren_from(toks, open_index)
                arg = toks[open_index + 1] if close == open_index + 2 else None
                counter = counters.get(_lower_name(arg) or "") if arg is not None else None
                message = (
                    _collection_counter_message(toks[i].raw_text, lower, arg.raw_text, counter)
                    if counter is not None and arg is not None
                    else None
                )
                if arg is not None and message:
                    push(
                        "collectionIndexOutOfRange",
                        message,
                        Span(stmt.span.start + arg.start, stmt.span.start + arg.end),
                    )


def _collection_counter_message(name: str, lower: str, counter_name: str, counter: LoopCounter) -> str | None:
    def is_count(atom: CounterAtom) -> bool:
        return atom.kind == "count" and atom.name == lower

    reached: str | None = None
    for counter_pass in numeric_counter_passes(counter, lambda _atom, _counter: None) or ():
        if counter_pass.value < 1:
            value = _js_string(counter_pass.value)
            reached = (
                f"Counter '{counter_name}' is {value} on its first pass"
                if counter_pass.pass_ == "first"
                else f"Counter '{counter_name}' reaches {value} on its last pass"
            )
            break
    if not reached:
        for which, bound in (("first", counter.first), ("last", counter.last)):
            if bound is not None and bound.atom is not None and is_count(bound.atom) and bound.offset > 0:
                reached = (
                    f"Counter '{counter_name}' is {counter_text(bound)} on its first pass"
                    if which == "first"
                    else f"Counter '{counter_name}' reaches {counter_text(bound)} on its last pass"
                )
                break
    if not reached:
        return None
    error = (
        "This will raise Run-time error '9': Subscript out of range."
        if _runs_with_an_element(counter, is_count)
        else f"This will raise Run-time error '9': Subscript out of range, or '5' if '{name}' is empty."
    )
    return f"{reached}, and '{name}' holds its elements at 1 to {name}.Count. {error}"


def _runs_with_an_element(counter: LoopCounter, is_count: Callable[[CounterAtom], bool]) -> bool:
    """Whether the loop's bounds show it runs only when the collection has an
    element: `For i = 0 To c.Count - 1`."""
    low, high = (counter.first, counter.last) if counter.step > 0 else (counter.last, counter.first)
    # The loop runs when low <= Count + offset, so Count >= low - offset.
    return (
        not (low is not None and low.atom is not None)
        and high is not None
        and high.atom is not None
        and is_count(high.atom)
        and low is not None
        and low.offset - high.offset >= 1
    )


# -- helpers ------------------------------------------------------------------


def _lower_name(tok: VbaToken | None) -> str | None:
    """tokenName(tok)?.toLowerCase()."""
    name = token_name(tok)
    return name.lower() if name is not None else None


def _token_at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """toks[index], or None past either end, as a JavaScript out-of-range read."""
    return toks[index] if 0 <= index < len(toks) else None


def _raw_text_at(toks: Sequence[VbaToken], index: int) -> str | None:
    return toks[index].raw_text if 0 <= index < len(toks) else None


def _element_at(values: Sequence[_T], index: int | float) -> _T | None:
    """`values[index - 1]`, None where a JavaScript array read would be undefined."""
    if isinstance(index, float) and not index.is_integer():
        return None
    position = int(index) - 1
    return values[position] if 0 <= position < len(values) else None


_FLOAT_SUFFIX = re.compile(r"[!#@]\Z")


def _float_literal_value(raw: str) -> float | None:
    """Number() of a float literal without its type suffix, when finite."""
    try:
        value = float(_FLOAT_SUFFIX.sub("", raw))
    except ValueError:
        return None
    return value if math.isfinite(value) else None


def _js_string(value: int | float) -> str:
    return js_number_to_string(value)


def _js_string_of(value: int | str) -> str:
    """String(value)."""
    return value if isinstance(value, str) else js_number_to_string(value)


__all__ = [
    "Replayed",
    "ReplayedHit",
    "check_collection_loop_counters",
    "check_collection_state",
    "replayed_calls",
    "replayed_diagnostic",
    "with_receiver",
]
