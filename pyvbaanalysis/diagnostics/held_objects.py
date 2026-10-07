"""What object a local is known to hold at a statement, and the classes of the
items a Collection local holds (XLIDE issue #246). Read from `Set x = New C`,
from `Set x = y` with y known, and from `c.Add New C`. Anything else that names a
local whole, which may pass it ByRef or replace it, ends what is known of it; a
member call or an index read (`x.Foo`, `c(1)`, `c.Count`) does not. Blocks are
entered as issue #237 enters them.

Ported from xlide_vscode/src/analyzer/diagnostics/heldObjects.ts.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterable, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ..conditional import ConditionalActivityTracker
from ..lexer.token_helpers import split_top_level_token_groups
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import BodyNode, LeafStatementNode, ProcedureNode, StatementNode, is_leaf_statement
from ..symbols.symbol_model import ModuleSymbols, VbaSymbolKind, SymbolVisibility
from ..types.type_names import normalize_type
from .walker import set_assignment_target, statement_tokens_after_leading_label, token_name, token_text


@dataclass(slots=True)
class HeldObjects:
    """What a statement sees. Consumers read it; the walk owns the dicts."""

    # The class each local holds, by lowercased name, as written after New.
    classes: dict[str, str]
    # The classes of a Collection local's items, in order.
    items: dict[str, list[str]]


_NOTHING_HELD = HeldObjects(classes={}, items={})

# The class recorded for ActiveSheet, which may be a Worksheet or a Chart.
ACTIVE_SHEET_HELD = "Worksheet or Chart"

# The item class recorded for a number or string a Collection holds: no object.
HELD_VALUE = "(value)"


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) out of range."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return None if tok is None else tok.raw_text


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name is not None else None


def _host_object_held(value: Sequence[VbaToken], declared: AbstractSet[str]) -> str | None:
    """`Set x = Application`, `Set x = ActiveWorkbook.Names` and, in Word, `Set x =
    ActiveDocument`: the host's own object, which is never Nothing (issues #415
    and #438)."""
    last = token_text(_at(value, len(value) - 1))
    if len(value) == 1 and last == "application" and "application" not in declared:
        return "Application"
    # Word's own Document (issue #438).
    if len(value) == 1 and (last == "activedocument" or last == "thisdocument") and last not in declared:
        return "Document"
    if last == "names" and (
        "names" not in declared if len(value) == 1 else _raw(value, len(value) - 2) == "."
    ):
        return "Names"
    # Excel's ActiveSheet, whichever kind of sheet it is: a Collection parameter
    # refuses it with 13 (issue #685).
    if len(value) == 1 and last == "activesheet" and last not in declared:
        return ACTIVE_SHEET_HELD
    return None


# The class each ProgID CreateObject makes, by lowercased ProgID (issue #685).
_PROGID_CLASSES: dict[str, str] = {
    "scripting.dictionary": "Scripting.Dictionary",
    "scripting.filesystemobject": "Scripting.FileSystemObject",
}


def _created_by_prog_id(value: Sequence[VbaToken], declared: AbstractSet[str]) -> str | None:
    """`Set d = CreateObject("Scripting.Dictionary")`: a Dictionary, which a
    Collection parameter refuses with 13 (issue #685)."""
    at = 2 if token_text(_at(value, 0)) == "vba" and _raw(value, 1) == "." else 0
    prog_id = _at(value, at + 2)
    if (
        token_text(_at(value, at)) != "createobject"
        or "createobject" in declared
        or _raw(value, at + 1) != "("
        or prog_id is None
        or prog_id.kind is not TokenKind.STRING_LITERAL
        or _raw(value, at + 3) != ")"
        or len(value) != at + 4
    ):
        return None
    return _PROGID_CLASSES.get(prog_id.raw_text[1:-1].lower())


# Members that read a Collection without changing it.
_COLLECTION_READS: frozenset[str] = frozenset({"count", "item"})


def held_objects_at(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    # The class of any other value a Set gives, `Set o = Range("A1").Font`, where
    # the caller can tell (issue #685).
    class_of_value: Callable[[Sequence[VbaToken], int], str | None] | None = None,
) -> Callable[[BodyNode], HeldObjects]:
    """What each statement, and each block as it is entered, sees. A statement the
    walk never reached sees nothing."""
    from ..flow.procedure_labels import jump_target_label_declaration
    from ..types.type_inference import procedure_symbol_for
    from .dataflow import BlockEnteringState, walk_entering_blocks
    from .rules.shared import names_in

    if re.search(r"\bon\s+error\b", source[proc.span.start:proc.span.end], re.IGNORECASE):
        return lambda _: _NOTHING_HELD

    # Keyed by node identity; each entry holds its node so the id stays its own.
    seen: dict[int, tuple[BodyNode, HeldObjects]] = {}
    state = HeldObjects(classes={}, items={})
    proc_symbol = procedure_symbol_for(symbols, proc)
    own = (proc_symbol.children if proc_symbol is not None else None) or []
    static_procedure = re.search(r"\bstatic\s+(?:sub|function|property)\b", source[proc.span.start:proc.body[0].span.start if proc.body else proc.span.end], re.IGNORECASE) is not None
    locals_ = {child.name.lower() for child in own if child.kind is VbaSymbolKind.LOCAL_VARIABLE and child.visibility is not SymbolVisibility.STATIC and not static_procedure}
    locals_.update(param.name.lower() for param in proc.params if param.by_val)
    # The names a local or parameter takes, which hide a host's global.
    declared = {child.name.lower() for child in own} | {param.name.lower() for param in proc.params}
    # `Dim c As New Collection` holds an empty one from the start.
    for child in own:
        if (
            child.name.lower() in locals_
            and child.is_auto_instantiated
            and not child.is_array
            and child.as_type
        ):
            state.classes[child.name.lower()] = child.as_type
            if normalize_type(child.as_type) == "collection":
                state.items[child.name.lower()] = []
    # Maps and item lists are copied; reuse the snapshot until tracked state changes.
    current_snapshot: HeldObjects | None = None

    def snapshot() -> HeldObjects:
        nonlocal current_snapshot
        if current_snapshot is None:
            current_snapshot = HeldObjects(
                classes=dict(state.classes),
                items={key: list(value) for key, value in state.items.items()},
            )
        return current_snapshot

    def drop_items(lower: str) -> None:
        nonlocal current_snapshot
        if state.items.pop(lower, None) is not None:
            current_snapshot = None

    def forget(names: Iterable[str]) -> None:
        nonlocal current_snapshot
        for lower in names:
            removed_class = state.classes.pop(lower, None) is not None
            removed_items = state.items.pop(lower, None) is not None
            if removed_class or removed_items:
                current_snapshot = None

    def add_item(head: str, toks: Sequence[VbaToken], item: str) -> None:
        nonlocal current_snapshot
        items = state.items[head]
        at = _add_position(split_top_level_token_groups(toks, 3, ","), len(items))
        if at is None:
            drop_items(head)
        else:
            current_snapshot = None
            items.insert(at, item)

    def visit(node: BodyNode) -> None:
        nonlocal current_snapshot
        if not is_leaf_statement(node):
            return
        toks = statement_tokens_after_leading_label(source, node.span)
        if jump_target_label_declaration(source, node.span) or token_text(_at(toks, 0)) == "gosub":
            forget([*state.classes.keys(), *state.items.keys()])
        if len(state.classes) > 0 or len(state.items) > 0:
            seen[id(node)] = (node, snapshot())
        if isinstance(node, StatementNode) and node.single_line_if_branches is not None:
            forget(names_in(source, node.span))
            return
        target = set_assignment_target(source, node.span)
        if target is not None:
            lower = target[0].lower()
            equals = next((i for i, tok in enumerate(toks) if tok.raw_text == "="), -1)
            value = toks[equals + 1 :]
            from_ = _lower_name(value[0]) if len(value) == 1 else None
            created = token_name(value[1]) if len(value) == 2 and token_text(value[0]) == "new" else None
            held = created
            if held is None and from_:
                held = state.classes.get(from_)
            if held is None:
                held = _host_object_held(value, declared)
            if held is None:
                held = _created_by_prog_id(value, declared)
            if held is None and class_of_value is not None:
                held = class_of_value(value, node.span.start)
            # The value's own holder may now change it unseen.
            forget([name for name in [lower, *names_in(source, node.span)] if name != from_ or not held])
            if from_ and held:
                drop_items(from_)
            if held and lower in locals_:
                current_snapshot = None
                state.classes[lower] = held
                if created and normalize_type(created) == "collection":
                    state.items[lower] = []
            return
        # `c.Add New Flat1` puts a Flat1 at the end; Before:=1 or After:=1, a whole
        # number, puts it there (issue #356). A key or an expression for either
        # leaves the order unknown.
        head = _lower_name(_at(toks, 0))
        sixth = _at(toks, 5)
        if (
            head
            and head in state.items
            and _raw(toks, 1) == "."
            and token_text(_at(toks, 2)) == "add"
            and token_text(_at(toks, 3)) == "new"
            and token_name(_at(toks, 4))
            and (sixth is None or sixth.raw_text == ",")
        ):
            item_class = token_name(toks[4])
            assert item_class is not None
            add_item(head, toks, item_class)
            return
        # `c.Add 1` or `c.Add "a"`: a value, no object (issue #447).
        literal = _at(toks, 3)
        fifth = _at(toks, 4)
        if (
            head
            and head in state.items
            and _raw(toks, 1) == "."
            and token_text(_at(toks, 2)) == "add"
            and literal is not None
            and literal.kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL, TokenKind.STRING_LITERAL)
            and (fifth is None or fifth.raw_text == ",")
        ):
            add_item(head, toks, HELD_VALUE)
            return
        forget_changed(toks)

    def forget_changed(toks: Sequence[VbaToken]) -> None:
        for i, tok in enumerate(toks):
            lower = _lower_name(tok)
            if (
                not lower
                or (lower not in state.classes and lower not in state.items)
                or _raw(toks, i - 1) == "."
                or _raw(toks, i - 1) == "!"
            ):
                continue
            nxt = _raw(toks, i + 1)
            if nxt == "(":
                continue  # an index read
            if nxt == "." or nxt == "!":
                # A member call leaves the object; only Count and Item leave a
                # Collection's items.
                if token_text(_at(toks, i + 2)) not in _COLLECTION_READS:
                    drop_items(lower)
                continue
            forget([lower])

    def restore(saved: HeldObjects) -> None:
        nonlocal current_snapshot
        state.classes = dict(saved.classes)
        state.items = {key: list(value) for key, value in saved.items.items()}
        current_snapshot = saved

    def touches(stmt: LeafStatementNode) -> Iterable[str]:
        return names_in(source, stmt.span)

    def enter(node: BodyNode) -> None:
        if len(state.classes) > 0 or len(state.items) > 0:
            seen[id(node)] = (node, snapshot())

    walk_entering_blocks(
        source,
        proc.body,
        lambda node: activity is not None and activity.is_inactive(node.span),
        visit,
        BlockEnteringState(
            snapshot=snapshot,
            restore=restore,
            forget=forget,
            touches=touches,
            enter=enter,
        ),
    )

    def at(node: BodyNode) -> HeldObjects:
        entry = seen.get(id(node))
        return entry[1] if entry is not None and entry[0] is node else _NOTHING_HELD

    return at


_WHOLE_NUMBER_RE = re.compile(r"^[0-9]+\Z")


def _add_position(args: Sequence[Sequence[VbaToken]], count: int) -> int | None:
    """Where `c.Add item[, key[, before[, after]]]` puts the item among `count`, as
    a 0-based index, or None when the code does not say: Before:=1 is the front,
    After:=1 the second place, nothing the end."""
    before: Sequence[VbaToken] | None = None
    after: Sequence[VbaToken] | None = None
    for k in range(1, len(args)):
        arg = args[k]
        named = token_text(arg[0]) if _raw(arg, 1) == ":=" else None
        value = arg[2:] if named is not None else arg
        role = named if named is not None else "before" if k == 2 else "after" if k == 3 else "key"
        if len(value) == 0 or role == "key" or role == "item":
            continue
        if role == "before":
            before = value
        elif role == "after":
            after = value
        else:
            return None
    place = before if before is not None else after
    if place is None:
        return count
    if before is not None and after is not None:
        return None
    n = int(place[0].raw_text) if len(place) == 1 and _WHOLE_NUMBER_RE.match(place[0].raw_text) else None
    if n is None or n < 1 or n > count:
        return None
    return n - 1 if before is not None else n
