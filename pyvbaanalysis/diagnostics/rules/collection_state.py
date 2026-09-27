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

The rule follows a procedure's top-level statements in order, as the file
rule does: a block ends what is known, and any use of the variable other
than Add, Remove, Item, Count and indexing ends it too.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from functools import partial

from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import parse_vba_integer_literal
from ...flow.procedure_labels import statement_label_declaration
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    VariableGroupNode,
    is_leaf_statement,
)
from ...types.type_names import normalize_type
from ..call_extraction import string_literal_value
from ..context import PushFn
from ..walker import (
    active_module_members,
    for_each_variable_group,
    is_inactive_node,
    match_paren_from,
    set_assignment_target,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)


@dataclass(slots=True)
class _CollectionContents:
    # Keys in element order; None for an element added without a key.
    items: list[str | None] = field(default_factory=list)
    # False once an Add named a key the rule could not read, or ordered by Before/After.
    keys_known: bool = True


@dataclass(frozen=True, slots=True)
class _CollectionLocals:
    # Lowercased names declared `As New Collection`: a collection from the start.
    new_locals: set[str]
    # Lowercased names declared `As Collection`: tracked once `Set x = New Collection`.
    plain_locals: set[str]


def check_collection_state(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        auto_instanced = _collection_locals(member, activity)
        states: dict[str, _CollectionContents] = {}
        for name in auto_instanced.new_locals:
            states[name] = _CollectionContents()
        if len(states) == 0 and len(auto_instanced.plain_locals) == 0:
            continue
        for node in member.body:
            if is_inactive_node(activity, node):
                continue
            if isinstance(node, VariableGroupNode):
                continue  # a Dim inside the body declares, and runs nothing
            if not is_leaf_statement(node):
                states.clear()
                continue
            if isinstance(node, StatementNode) and node.single_line_if_branches is not None:
                _forget_mentioned(source, node.span, states)
                continue
            toks = statement_tokens_after_leading_label(source, node.span)
            if len(toks) == 0:
                continue
            # A label may be reached from anywhere; a GoSub may run any statement.
            if (
                statement_label_declaration(source, node.span) is not None
                or token_text(toks[0]) == "gosub"
            ):
                states.clear()
            # `Set c = New Collection` starts an empty collection. `Set o = c` makes
            # o and c one collection, so they share one state and an Add through
            # either is seen by both (XLIDE issue #147). Any other Set ends tracking
            # of its target, and of every tracked collection its value names, since
            # the value's new holder can change it unseen.
            target = set_assignment_target(source, node.span)
            if target is not None:
                lower = target[0].lower()
                equals = next((index for index, tok in enumerate(toks) if tok.raw_text == "="), -1)
                value = toks[equals + 1 :]
                is_collection_local = (
                    lower in auto_instanced.plain_locals or lower in auto_instanced.new_locals
                )
                value_name = token_name(value[0]) if len(value) == 1 else None
                aliased = value_name.lower() if value_name is not None else None
                if (
                    is_collection_local
                    and len(value) == 2
                    and token_text(value[0]) == "new"
                    and token_text(value[1]) == "collection"
                ):
                    states[lower] = _CollectionContents()
                    continue
                if is_collection_local and aliased is not None and aliased in states:
                    states[lower] = states[aliased]
                    continue
                states.pop(lower, None)
                for tok in value:
                    mentioned = token_name(tok)
                    if mentioned is not None and mentioned.lower() in states:
                        del states[mentioned.lower()]
                continue
            _check_statement(node.span, toks, states, push)


def _collection_locals(
    proc: ProcedureNode, activity: ConditionalActivityTracker | None
) -> _CollectionLocals:
    new_locals: set[str] = set()
    plain_locals: set[str] = set()

    def visit(group: VariableGroupNode) -> None:
        if group.is_const or group.modifier.lower() == "static":
            return
        for decl in group.declarations:
            if decl.is_array or normalize_type(decl.as_type) != "collection":
                continue
            (new_locals if decl.is_new else plain_locals).add(decl.name.lower())

    for_each_variable_group(proc.body, visit, activity)
    return _CollectionLocals(new_locals, plain_locals)


def _forget_mentioned(source: str, span: Span, states: dict[str, _CollectionContents]) -> None:
    """Drops every tracked collection a statement names anywhere."""
    for tok in statement_tokens_after_leading_label(source, span):
        name = token_name(tok)
        lower = name.lower() if name is not None else None
        if lower and lower in states:
            del states[lower]


def _check_statement(
    base: Span,
    toks: Sequence[VbaToken],
    states: dict[str, _CollectionContents],
    push: PushFn,
) -> None:
    def at(first: int, last: int) -> Span:
        return Span(base.start + toks[first].start, base.start + toks[last].end)

    # First pass: reads and the recognised forms, in source order. A mention
    # in any other shape ends tracking of that variable after this statement.
    to_forget: set[str] = set()
    # Deferred with each call's own arguments bound now, as the upstream
    # closures capture their block-scoped consts.
    mutations: list[Callable[[], None]] = []
    head_index = 1 if token_text(toks[0]) == "call" else 0
    for i in range(head_index, len(toks)):
        name = token_name(toks[i])
        lower = name.lower() if name is not None else None
        if not lower or lower not in states or (i > 0 and toks[i - 1].raw_text == "."):
            continue
        state = states[lower]
        following = _token_at(toks, i + 1)
        # `c(index)` or `c("key")`
        if following is not None and following.raw_text == "(":
            close = match_paren_from(toks, i + 1)
            if close > i + 2 and _check_read(
                lower, state, toks[i + 2 : close], at(i + 2, close - 1), push
            ):
                continue
            to_forget.add(lower)
            continue
        if following is None or following.raw_text != ".":
            to_forget.add(lower)
            continue
        member_name = token_text(_token_at(toks, i + 2))
        if member_name == "count":
            continue
        if member_name == "item":
            item_open = _token_at(toks, i + 3)
            item_close = (
                match_paren_from(toks, i + 3)
                if item_open is not None and item_open.raw_text == "("
                else -1
            )
            if item_close > i + 4 and _check_read(
                lower, state, toks[i + 4 : item_close], at(i + 4, item_close - 1), push
            ):
                continue
            to_forget.add(lower)
            continue
        if (member_name == "add" or member_name == "remove") and i == head_index:
            args = _arguments_after(toks, i + 3)
            if member_name == "add":
                mutations.append(partial(_add, lower, state, args, base, push))
            elif len(args) == 1:
                mutations.append(partial(_remove, lower, state, args[0], base, push))
            else:
                to_forget.add(lower)
            continue
        to_forget.add(lower)
    for mutation in mutations:
        mutation()
    for lower in to_forget:
        states.pop(lower, None)


def _token_at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """toks[index], or None past either end, as a JavaScript out-of-range read."""
    return toks[index] if 0 <= index < len(toks) else None


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
    name: str, state: _CollectionContents, arg: Sequence[VbaToken], span: Span, push: PushFn
) -> bool:
    """Judges `c(arg)` or `c.Item(arg)`; false when the argument is not a literal the rule reads."""
    index = _literal_index(arg)
    if index is not None:
        _report_index(name, state, index, span, push)
        return True
    if len(arg) == 1 and arg[0].kind is TokenKind.STRING_LITERAL:
        _report_key(name, state, string_literal_value(arg[0].raw_text), span, push)
        return True
    return False


def _report_index(
    name: str, state: _CollectionContents, index: int, span: Span, push: PushFn
) -> None:
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
            f"indexed 1 to {count}; {index} is outside that. "
            "This will raise Run-time error '9': Subscript out of range.",
            span,
        )


def _report_key(
    name: str, state: _CollectionContents, key: str, span: Span, push: PushFn
) -> bool:
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


def _add(
    name: str,
    state: _CollectionContents,
    args: list[list[VbaToken]],
    base: Span,
    push: PushFn,
) -> None:
    key_arg = args[1] if len(args) > 1 else None
    key: str | None = None
    if key_arg is not None and len(key_arg) > 0:
        key = _literal_key(key_arg)
        if key is None:
            state.keys_known = False
    if key is not None and key_arg is not None and state.keys_known and key in state.items:
        push(
            "collectionKeyInUse",
            f"'{name}' already has an element with the key "
            f"\"{string_literal_value(key_arg[0].raw_text)}\" (keys compare without case). "
            "This will raise Run-time error '457': "
            "This key is already associated with an element of this collection.",
            Span(base.start + key_arg[0].start, base.start + key_arg[-1].end),
        )
        return
    if len(args) > 2 and any(len(arg) > 0 for arg in args[2:]):
        # Before or After: the position is not followed, the count is.
        state.items.append(key)
        state.keys_known = False
        return
    state.items.append(key)


def _remove(
    name: str,
    state: _CollectionContents,
    arg: list[VbaToken],
    base: Span,
    push: PushFn,
) -> None:
    span = Span(base.start + arg[0].start, base.start + arg[-1].end)
    index = _literal_index(arg)
    if index is not None:
        if len(state.items) == 0 or index < 1 or index > len(state.items):
            _report_index(name, state, index, span, push)
            return
        del state.items[index - 1]
        return
    key = _literal_key(arg)
    if key is None:
        # A variable index or key: one element fewer, which one unknown.
        _drop_last(state.items)
        state.keys_known = False
        return
    if _report_key(name, state, string_literal_value(arg[0].raw_text), span, push):
        return
    if key in state.items:
        state.items.remove(key)
    else:
        _drop_last(state.items)
        state.keys_known = False


def _drop_last(items: list[str | None]) -> None:
    """Array.prototype.pop: removes the last element, and does nothing on an empty list."""
    if len(items) > 0:
        items.pop()
