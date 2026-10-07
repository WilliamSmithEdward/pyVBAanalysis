"""Rule family: a Scripting.Dictionary whose keys the code makes plain (XLIDE issue
#243). Measured in Excel 16.0 (build 20326, 2026-09-30); each compiles and raises
every time it runs.

 - `d.Add "a", 1` then `d.Add "a", 2` -> 457, the key is in use.
 - `d.Remove "a"` with no such key -> 32811.
 - `d.Keys()(3)` or `d.Items()(3)` past the last key -> 9. They are based at 0.

A Dictionary differs from a Collection, whose state collection_state.py follows:
its keys compare as written, "A" and "a" being two keys, and a number and a
string being two; and reading `d("missing")` runs, adding the key with an Empty
item. The local is followed from `Set d = CreateObject("Scripting.Dictionary")`
or `New Scripting.Dictionary`. Anything else that names it ends what is known.

Issue #349, measured in Excel 16.0 on 2026-10-02: `d.CompareMode = 1` after a
key is in raises 5, and before it makes "k" and "K" one key; `d.Key("a") = "b"`
raises 32811 with no "a" and 457 with "b" in use; an Array as the key of Add
raises 5; and `d("k").Count` with no "k" adds the key with an Empty item, whose
member raises 424.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/dictionaryState.ts.
"""

from __future__ import annotations

import json
import re
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from typing import AbstractSet

from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import jump_target_label_declaration
from ...js_compat import js_number_to_string
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    VariableGroupNode,
    WithBlockNode,
    is_leaf_statement,
)
from ..call_extraction import string_literal_value
from ..callee_arguments import CalleeMemberCalls, callee_member_calls
from ..context import PushFn
from ..dataflow import BlockEnteringState, walk_entering_blocks
from ..walker import (
    active_module_members,
    block_header_line_span,
    match_paren_from,
    set_assignment_target,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .collection_state import Replayed, replayed_calls, replayed_diagnostic, with_receiver
from .shared import names_in

# The name a `With CreateObject("Scripting.Dictionary")` block's Dictionary is followed under.
_NEW_WITH_DICTIONARY = "New Dictionary"


@dataclass(slots=True, eq=False)
class _DictionaryKeys:
    """The keys of one Dictionary in the order they were added, each a typed literal."""

    keys: list[str] = field(default_factory=list)
    # The keys whose item is a number or a string literal the code wrote (issue #306).
    values: dict[str, str] | None = None
    # Set by `CompareMode = 1`: string keys compare without case (issue #349).
    text_compare: bool | None = None


def _copy_keys(state: _DictionaryKeys) -> _DictionaryKeys:
    return _DictionaryKeys(
        list(state.keys), dict(state.values) if state.values is not None else None, state.text_compare
    )


def check_dictionary_state(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    callee_calls = callee_member_calls(source)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        _check_procedure(source, member, activity, push, callee_calls)


def _check_procedure(
    source: str,
    member: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    callee_calls: CalleeMemberCalls,
) -> None:
    if re.search(r"\bon\s+error\b", source[member.span.start:member.span.end], re.IGNORECASE):
        return
    from ..walker import for_each_variable_group
    locals_: set[str] = set()

    def collect(group: VariableGroupNode) -> None:
        if not group.is_const and group.modifier.lower() != "static":
            locals_.update(decl.name.lower() for decl in group.declarations)

    for_each_variable_group(member.body, collect, activity)
    states: dict[str, _DictionaryKeys] = {}
    # The subject of each With block the walk is in, as for Collections.
    with_subjects: list[str | None] = []

    def forget(names: Iterable[str]) -> None:
        for lower in names:
            states.pop(lower, None)

    def visit(node: BodyNode) -> None:
        if not is_leaf_statement(node):
            return
        if jump_target_label_declaration(source, node.span):
            states.clear()
        if isinstance(node, StatementNode) and node.single_line_if_branches is not None:
            forget(names_in(source, node.span))
            return
        own = statement_tokens_after_leading_label(source, node.span)
        # Inside `With d`, `.Add "a", 1` is d's (issue #295, measured in Excel 16.0).
        subject = with_subjects[-1] if with_subjects else None
        toks = with_receiver(own, subject) if subject else own
        if token_text(_token_at(toks, 0)) == "gosub":
            states.clear()
            return
        target = set_assignment_target(source, node.span)
        if target is not None:
            lower = target[0].lower()
            equals = next((k for k, tok in enumerate(toks) if tok.raw_text == "="), -1)
            value = toks[equals + 1 :]
            # `Set x = d("k")` where "k" holds a number: no object to Set (issue
            # #306, measured in Excel 16.0).
            item = _value_read(value, states)
            if item:
                push(
                    "variantValueMisuse",
                    f"'{''.join(tok.raw_text for tok in value)}' holds a {item}, not an object, so Set has nothing "
                    "to assign. This will raise Run-time error '424': Object required.",
                    Span(node.span.start + value[0].start, node.span.start + value[-1].end),
                )
            forget(names_in(source, node.span))
            if lower in locals_ and _creates_dictionary(value):
                states[lower] = _DictionaryKeys()
            return
        # `AddK d` where AddK only adds to or removes from its parameter: d's keys
        # change as those calls change them (issue #685).
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

                    _check_statement(node.span, call, states, capture)
                    if raised:
                        hit = replayed_diagnostic(node.span, toks, lower, call, raised[0])
                        push(hit.rule, hit.message, hit.span)
                    if raised or lower not in states:
                        states.pop(lower, None)
                        break
            return
        _check_statement(node.span, toks, states, push)

    def snapshot() -> dict[str, _DictionaryKeys]:
        return {lower: _copy_keys(state) for lower, state in states.items()}

    def restore(saved: dict[str, _DictionaryKeys]) -> None:
        states.clear()
        for lower, state in saved.items():
            states[lower] = _copy_keys(state)

    def forget_set(names: AbstractSet[str]) -> None:
        forget(names)

    def touches(stmt: LeafStatementNode) -> AbstractSet[str]:
        toks = statement_tokens_after_leading_label(source, stmt.span)
        if token_text(_token_at(toks, 0)) == "with" and (len(toks) == 2 or _creates_dictionary(toks[1:])):
            return set()
        # `For Each x In d` reads d, and changes only x (issue #349).
        if (
            token_text(_token_at(toks, 0)) == "for"
            and token_text(_token_at(toks, 1)) == "each"
            and token_text(_token_at(toks, 3)) == "in"
            and len(toks) == 5
            and token_name(toks[2])
        ):
            return {(token_name(toks[2]) or "").lower()}
        return names_in(source, stmt.span)

    def enter(node: BodyNode) -> None:
        if not isinstance(node, WithBlockNode):
            return
        header = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
        name = token_name(header[1]) if len(header) == 2 else None
        if name and name.lower() in states:
            with_subjects.append(name)
        elif _creates_dictionary(header[1:]):
            states[_NEW_WITH_DICTIONARY.lower()] = _DictionaryKeys()
            with_subjects.append(_NEW_WITH_DICTIONARY)
        else:
            with_subjects.append(None)

    def exit_block(node: BodyNode) -> None:
        if isinstance(node, WithBlockNode):
            popped = with_subjects.pop() if with_subjects else None
            if popped == _NEW_WITH_DICTIONARY:
                states.pop(_NEW_WITH_DICTIONARY.lower(), None)

    walk_entering_blocks(
        source,
        member.body,
        lambda node: activity is not None and activity.is_inactive(node.span),
        visit,
        BlockEnteringState(
            snapshot=snapshot,
            restore=restore,
            forget=forget_set,
            touches=touches,
            enter=enter,
            exit=exit_block,
        ),
    )


def _value_read(value: Sequence[VbaToken], states: Mapping[str, _DictionaryKeys]) -> str | None:
    """What `d("k")` or `d.Item("k")`, the whole of a value, reads where the code wrote a literal there."""
    state = states.get(_lower_name(_token_at(value, 0)) or "")
    if _raw_text_at(value, 1) == "(":
        open_index = 1
    elif (
        _raw_text_at(value, 1) == "."
        and token_text(_token_at(value, 2)) == "item"
        and _raw_text_at(value, 3) == "("
    ):
        open_index = 3
    else:
        open_index = -1
    if state is None or open_index < 0 or match_paren_from(value, open_index) != len(value) - 1:
        return None
    key = _literal_key(value[open_index + 1 : len(value) - 1])
    if not key or state.values is None:
        return None
    return state.values.get(f"s:{key[2:].lower()}" if state.text_compare and key.startswith("s:") else key)


def _remember(state: _DictionaryKeys, key: str, kind: str | None) -> None:
    """Notes what a key's item is now, or forgets it."""
    if kind:
        if state.values is None:
            state.values = {}
        state.values[key] = kind
    elif state.values is not None:
        state.values.pop(key, None)


def _literal_kind(value: Sequence[VbaToken]) -> str | None:
    """The kind of a one-token literal value: a number or a string."""
    if len(value) != 1:
        return None
    if value[0].kind is TokenKind.STRING_LITERAL:
        return "string"
    return "number" if value[0].kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL) else None


_CREATE_DICTIONARY = re.compile(r'(?:vba\.)?createobject\("scripting\.dictionary"\)')


def _creates_dictionary(value: Sequence[VbaToken]) -> bool:
    """`CreateObject("Scripting.Dictionary")`, `VBA.CreateObject(...)` or `New Scripting.Dictionary`."""
    text = "".join(tok.raw_text for tok in value).lower()
    return _CREATE_DICTIONARY.fullmatch(text) is not None or text == "newscripting.dictionary"


_DIGITS = re.compile(r"[0-9]+")


def _literal_key(arg: Sequence[VbaToken]) -> str | None:
    """The literal key an argument names, typed so that "1" and 1 differ, or None."""
    if len(arg) != 1:
        return None
    if arg[0].kind is TokenKind.STRING_LITERAL:
        return f"s:{string_literal_value(arg[0].raw_text)}"
    if arg[0].kind is TokenKind.INTEGER_LITERAL and _DIGITS.fullmatch(arg[0].raw_text) is not None:
        return f"n:{js_number_to_string(int(arg[0].raw_text))}"
    return None


def _shown_key(key: str) -> str:
    return json.dumps(key[2:], ensure_ascii=False) if key.startswith("s:") else key[2:]


def _arguments_of(toks: Sequence[VbaToken], start: int, end: int | None = None) -> list[list[VbaToken]]:
    """The comma-separated arguments from `toks[start]` to the end, or within one pair of parentheses."""
    stop = len(toks) if end is None else end
    out: list[list[VbaToken]] = []
    current: list[VbaToken] = []
    depth = 0
    for i in range(start, stop):
        raw = toks[i].raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
        if raw == "," and depth == 0:
            out.append(current)
            current = []
            continue
        current.append(toks[i])
    out.append(current)
    return out


def _literal_item_argument(value: Sequence[VbaToken]) -> bool:
    if len(value) == 1:
        return value[0].kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL, TokenKind.STRING_LITERAL, TokenKind.DATE_LITERAL) or token_text(value[0]) in ("true", "false", "nothing", "empty", "null")
    return len(value) == 2 and value[0].raw_text in ("+", "-") and value[1].kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL)


def _check_statement(base: Span, toks: Sequence[VbaToken], states: dict[str, _DictionaryKeys], push: PushFn) -> None:
    def at(first: VbaToken, last: VbaToken) -> Span:
        return Span(base.start + first.start, base.start + last.end)

    head = _lower_name(_token_at(toks, 0))
    state = states.get(head) if head else None

    # Under `CompareMode = 1` "k" and "K" are one key.
    def key_of(arg: Sequence[VbaToken], held: _DictionaryKeys) -> str | None:
        key = _literal_key(arg)
        if key and held.text_compare and key.startswith("s:") and re.search(r"[^\x20-\x7e]|i", key[2:], re.IGNORECASE):
            return None
        return f"s:{key[2:].lower()}" if key and held.text_compare and key.startswith("s:") else key

    eq = next((k for k, tok in enumerate(toks) if tok.raw_text == "="), -1)
    # `d.CompareMode = 1`: refused once a key is in (issue #349, measured).
    if state is not None and head and _raw_text_at(toks, 1) == "." and token_text(_token_at(toks, 2)) == "comparemode" and eq == 3:
        mode = token_text(_token_at(toks, 4))
        text = len(toks) == 5 and mode in ("1", "vbtextcompare")
        binary = len(toks) == 5 and mode in ("0", "vbbinarycompare")
        # The mode it already has is no change, and runs (issue #556).
        changes = (text and not state.text_compare) or (binary and state.text_compare is True)
        if len(state.keys) > 0 and changes:
            count = len(state.keys)
            push(
                "collectionAddArgument",
                f"Dictionary '{toks[0].raw_text}' already holds {count} key{'' if count == 1 else 's'}, so its "
                "CompareMode cannot change. This will raise Run-time error '5': Invalid procedure call or argument.",
                at(toks[2], toks[-1]),
            )
            return
        if text or binary:
            state.text_compare = text
        else:
            states.clear()
        return
    # `d.Key("a") = "b"` renames a key: 32811 with no "a", 457 with "b" in use.
    if (
        state is not None
        and head
        and _raw_text_at(toks, 1) == "."
        and token_text(_token_at(toks, 2)) == "key"
        and _raw_text_at(toks, 3) == "("
        and eq == match_paren_from(toks, 3) + 1
    ):
        source_key = key_of(toks[4 : eq - 1], state)
        target_key = key_of(toks[eq + 1 :], state)
        if not source_key or not target_key:
            states.clear()
            return
        if source_key not in state.keys:
            push(
                "collectionKeyNotFound",
                f"Dictionary '{toks[0].raw_text}' holds no key {_shown_key(source_key)} to rename. This will raise "
                "Run-time error '32811': Application-defined or object-defined error.",
                at(toks[4], toks[eq - 2]),
            )
            return
        if target_key in state.keys and target_key != source_key:
            push(
                "collectionKeyInUse",
                f"Dictionary '{toks[0].raw_text}' already holds the key {_shown_key(target_key)}. This will raise "
                "Run-time error '457': This key is already associated with an element of this collection.",
                at(toks[eq + 1], toks[-1]),
            )
            return
        state.keys[state.keys.index(source_key)] = target_key
        moved = state.values.get(source_key) if state.values is not None else None
        if state.values is not None:
            state.values.pop(source_key, None)
        if moved and state.values is not None:
            state.values[target_key] = moved
        return
    # A method called as a statement: `d.Add k, v`, `d.Remove k`, `d.RemoveAll`.
    if state is not None and _raw_text_at(toks, 1) == "." and not any(tok.raw_text == "=" for tok in toks):
        method = token_text(_token_at(toks, 2))
        paren = _raw_text_at(toks, 3) == "(" and match_paren_from(toks, 3) == len(toks) - 1
        if len(toks) > 3:
            args = _arguments_of(toks, 4, len(toks) - 1) if paren else _arguments_of(toks, 3)
        else:
            args = []
        key = key_of(args[0], state) if args else None
        # `d.Add Array(1), 1`: an array is no key (issue #349, measured).
        if (
            method == "add"
            and len(args) == 2
            and token_text(_token_at(args[0], 0)) == "array"
            and _raw_text_at(args[0], 1) == "("
            and match_paren_from(args[0], 1) == len(args[0]) - 1
        ):
            states.clear()
            push(
                "collectionAddArgument",
                f"The key of '{toks[0].raw_text}.Add' is an array, which a Dictionary takes as no key. This will "
                "raise Run-time error '5': Invalid procedure call or argument.",
                at(args[0][0], args[0][-1]),
            )
            return
        if method == "add" and len(args) == 2 and (not key or not _literal_item_argument(args[1])):
            states.clear()
            return
        if method == "add" and len(args) == 2 and key:
            if key in state.keys:
                push(
                    "collectionKeyInUse",
                    f"Dictionary '{toks[0].raw_text}' already holds the key {_shown_key(key)}. This will raise "
                    "Run-time error '457': This key is already associated with an element of this collection.",
                    at(args[0][0], args[0][-1]),
                )
                return
            state.keys.append(key)
            _remember(state, key, _literal_kind(args[1]))
            return
        if method == "remove" and len(args) == 1 and key:
            if key not in state.keys:
                push(
                    "collectionKeyNotFound",
                    f"Dictionary '{toks[0].raw_text}' holds no key {_shown_key(key)} to remove. This will raise "
                    "Run-time error '32811': Application-defined or object-defined error.",
                    at(args[0][0], args[0][-1]),
                )
                return
            state.keys.remove(key)
            if state.values is not None:
                state.values.pop(key, None)
            return
        if method == "removeall" and len(args) == 0:
            state.keys = []
            state.values = None
            return
    # Reads, anywhere in the statement: `d.Keys()(n)`, `d.Items()(n)`, `d(k)`, `d.Item(k)`.
    for i in range(len(toks)):
        lower = _lower_name(toks[i])
        read = states.get(lower) if lower else None
        if read is None or _raw_text_at(toks, i - 1) == "." or _raw_text_at(toks, i - 1) == "!":
            continue
        member = token_text(_token_at(toks, i + 2)) if _raw_text_at(toks, i + 1) == "." else ""
        if (
            member in ("keys", "items")
            and _raw_text_at(toks, i + 3) == "("
            and _raw_text_at(toks, i + 4) == ")"
            and _raw_text_at(toks, i + 5) == "("
        ):
            # `d.Items()(-1)`: no Keys or Items array has a negative index (issue #349).
            close = match_paren_from(toks, i + 5)
            negative = _raw_text_at(toks, i + 6) == "-"
            digits = _token_at(toks, i + (7 if negative else 6))
            index = (
                int(digits.raw_text) * (-1 if negative else 1)
                if close == i + (8 if negative else 7)
                and digits is not None
                and digits.kind is TokenKind.INTEGER_LITERAL
                and _DIGITS.fullmatch(digits.raw_text) is not None
                else None
            )
            if index is not None and digits is not None and (index < 0 or index >= len(read.keys)):
                count = len(read.keys)
                held = (
                    "holds no keys"
                    if count == 0
                    else f"holds {count} key{'' if count == 1 else 's'}, indexed 0 to {count - 1}"
                )
                push(
                    "collectionIndexOutOfRange",
                    f"Dictionary '{toks[i].raw_text}' {held} here; {js_number_to_string(index)} is outside that. "
                    "This will raise Run-time error '9': Subscript out of range.",
                    at(toks[i + 6], digits),
                )
            continue
        # `UBound(d.Keys)`: the arrays read the Dictionary and change nothing.
        if member in ("keys", "items", "count", "exists"):
            continue
        # `d(k)` and `d.Item(k)` add a key they do not find, read or written.
        if _raw_text_at(toks, i + 1) == "(":
            open_index = i + 1
        elif member == "item" and _raw_text_at(toks, i + 3) == "(":
            open_index = i + 3
        else:
            open_index = -1
        close = match_paren_from(toks, open_index) if open_index >= 0 else -1
        key = key_of(toks[open_index + 1 : close], read) if close > open_index + 1 else None
        # `d("k") = 5` writes the item, and `Set d("k") = o` an object.
        written = _raw_text_at(toks, close + 1) == "=" and (
            i == 0 or (i == 1 and token_text(toks[0]) == "set")
        )
        if written and i == 0:
            value = toks[close + 2 :]
            # `d("k") = New Collection`: the Let reads the Collection's value
            # (issue #306, measured in Excel 16.0).
            if len(value) == 2 and token_text(value[0]) == "new" and token_text(value[1]) == "collection":
                push(
                    "objectDefaultValue",
                    f"'New Collection' is assigned to '{''.join(tok.raw_text for tok in toks[: close + 1])}' "
                    "without Set, so its value is read, and a Collection's default member Item needs an index. "
                    "This will raise Run-time error '450': Wrong number of arguments or invalid property "
                    "assignment.",
                    at(value[0], value[1]),
                )
        if key:
            if written:
                _remember(read, key, _literal_kind(toks[close + 2 :]) if i == 0 else None)
            if key not in read.keys:
                # `d("k").Count` with no "k": the read adds it with an Empty item,
                # which has no members (issue #349, measured).
                if _raw_text_at(toks, close + 1) == ".":
                    after_dot = _raw_text_at(toks, close + 2)
                    push(
                        "variantValueMisuse",
                        f"Dictionary '{toks[i].raw_text}' holds no key {_shown_key(key)}, so the read adds it with "
                        f"an Empty item, which has no {after_dot if after_dot is not None else 'member'}. This "
                        "will raise Run-time error '424': Object required.",
                        at(toks[i], toks[close]),
                    )
                read.keys.append(key)
            continue
        # Anything else may change it.
        if lower:
            states.pop(lower, None)


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name is not None else None


def _token_at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    return toks[index] if 0 <= index < len(toks) else None


def _raw_text_at(toks: Sequence[VbaToken], index: int) -> str | None:
    return toks[index].raw_text if 0 <= index < len(toks) else None
