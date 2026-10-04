"""Rule family: a name a Word document or a PowerPoint presentation already
holds, given again (XLIDE issue #311).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/documentNames.ts.

Measured in Word and PowerPoint 16.0 (2026-10-02); each compiles and raises
every time it runs.

 - Word: `d.Variables.Add "zq", 1` twice raises 5903, `d.Styles.Add "Zq"`
   twice 5173, and `d.CustomDocumentProperties.Add Name:="zq", ...` twice
   -2147467259. The names ignore case.
 - PowerPoint: two slides the procedure added, `Set a = p.Slides.Add(...)`,
   cannot share a name, and slide names ignore case: `a.Name = "Zq"` then
   `b.Name = "zq"` raises -2147188160.

XLIDE issue #610 adds, each measured: a variable added again after its value
was written, or through another name for the document, `Set d2 = d` (5903); a
custom property named "" (-2147418113); and, on a presentation the procedure
made with Presentations.Add, the slides it adds: each is named Slide1, Slide2
in the order it was added, so `a.Name = b.Name` and `a.Name = "Slide2"` raise
-2147188160, as does `p.Slides(2).Name` given a name another slide has; and
`p.Slides.Add 3` or `a.MoveTo 2` past the count raise -2147188160 too.

What is known is followed in a straight line. Any other mention of the
document or slide variable, a label, a block or a call into the project's own
code ends it.
"""

from __future__ import annotations

import math
from collections.abc import Callable, Iterable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass, field

from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import statement_label_declaration
from ...js_compat import js_number, js_number_to_string
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    is_leaf_statement,
)
from ..call_extraction import string_literal_value
from ..context import PushFn
from ..dataflow import BlockEnteringState, walk_entering_blocks
from ..walker import (
    active_module_members,
    match_paren_from,
    set_assignment_target,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .shared import names_in


@dataclass(frozen=True, slots=True)
class _WordNamed:
    error: str
    noun: str


# Word collections whose Add refuses a name already in, by lowercased member.
_WORD_NAMED: Mapping[str, _WordNamed] = {
    "variables": _WordNamed("'5903': The Variable name already exists", "variable"),
    "styles": _WordNamed(
        "'5173': This style name already exists or is reserved for a built-in style", "style"
    ),
    "customdocumentproperties": _WordNamed("'-2147467259': Automation error", "custom property"),
}

# The state key of a slide the procedure added; a variable name holds no `#`.
_SLIDE = "#slide:"


@dataclass(slots=True)
class _PresentationState:
    """A presentation the procedure made, and the slides it holds, in order."""

    # Slide state keys in order; None for a slide not followed.
    order: list[str | None] = field(default_factory=list)
    # How many slides were added, which numbers the next default name.
    added: int = 0


@dataclass(frozen=True, slots=True)
class _VariableUse:
    key: str
    name: str
    deletes: bool


@dataclass(frozen=True, slots=True)
class _WordAdd:
    key: str
    chain: str
    name: str
    name_token: VbaToken


def _at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """`toks[index]`, or None out of range the way a JavaScript index reads undefined."""
    return toks[index] if 0 <= index < len(toks) else None


def _raw(toks: Sequence[VbaToken], index: int) -> str | None:
    tok = _at(toks, index)
    return tok.raw_text if tok is not None else None


def _kind(toks: Sequence[VbaToken], index: int) -> TokenKind | None:
    tok = _at(toks, index)
    return tok.kind if tok is not None else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name is not None else None


def _plural(count: int) -> str:
    return "" if count == 1 else "s"


def check_document_names(
    source: str,
    mod: ModuleNode,
    host: str | None,
    callables: AbstractSet[str],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    if host != "Word" and host != "PowerPoint":
        return
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        _check_procedure(source, member, host, callables, activity, push)


def _check_procedure(
    source: str,
    member: ProcedureNode,
    host: str,
    callables: AbstractSet[str],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    # `d.variables` -> the names added; `#slide:a` -> the name slide a was given ('' none).
    state: dict[str, set[str]] = {}
    # By presentation variable (XLIDE issue #610).
    presentations: dict[str, _PresentationState] = {}
    # The presentation each followed slide is in, by slide state key.
    slide_in: dict[str, str] = {}

    def forget(names: Iterable[str]) -> None:
        gone = set(names)
        for key in list(state.keys()):
            head = key[len(_SLIDE) :] if key.startswith(_SLIDE) else key.split(".")[0]
            if head in gone:
                state.pop(key, None)
                slide_in.pop(key, None)
        for lower in gone:
            presentations.pop(lower, None)

    # The slide `p.Slides(2)` names, by its state key, or None.
    def slide_at(pres: str, index: float) -> str | None:
        held = presentations.get(pres)
        if held is None or math.isnan(index) or not float(index).is_integer():
            return None
        k = int(index) - 1
        return held.order[k] if 0 <= k < len(held.order) else None

    # A name another followed slide holds, in any case: its variable.
    def holder(key: str, lower: str) -> str | None:
        for other, held in state.items():
            if other.startswith(_SLIDE) and other != key and lower in held:
                return other[len(_SLIDE) :]
        return None

    def name_slide(key: str, name: str, name_at: VbaToken, at: Callable[[VbaToken], Span]) -> None:
        lower = name.lower()
        taken = holder(key, lower)
        if taken is not None:
            who = "a slide the code added" if taken.startswith("#") else f"'{taken}', another slide the code added"
            push(
                "hostArgumentOutOfRange",
                f"\"{name}\" is the name of {who}, and slide names ignore case. This will raise "
                "Run-time error '-2147188160': Another slide already has this name.",
                at(name_at),
            )
        state[key] = {lower}

    # `Slides.Add(i, ...)` on a presentation the procedure made: a new slide at
    # i, named Slide<n> by the order it was added.
    def add_slide(
        pres: str,
        index: float | None,
        key: str,
        index_at: VbaToken | None,
        at: Callable[[VbaToken], Span],
    ) -> None:
        held = presentations.get(pres)
        if held is None:
            return
        if index is None or index < 1:
            presentations.pop(pres, None)
            return
        if index > len(held.order) + 1:
            count = len(held.order)
            assert index_at is not None
            push(
                "hostArgumentOutOfRange",
                f"'{pres}' holds {count} slide{_plural(count)}, so a new one goes at 1 to {count + 1}; "
                f"{js_number_to_string(index)} is past that. This will raise Run-time error "
                "'-2147188160': Integer out of range.",
                at(index_at),
            )
            presentations.pop(pres, None)
            return
        held.added += 1
        # Array.prototype.splice reads a NaN start as 0, and truncates a fraction.
        start = 0 if math.isnan(index) else int(index - 1)
        held.order.insert(start, key)
        state[key] = {f"slide{held.added}"}
        slide_in[key] = pres

    def visit(node: BodyNode) -> None:
        if not is_leaf_statement(node):
            return
        toks = statement_tokens_after_leading_label(source, node.span)
        if (
            statement_label_declaration(source, node.span)
            or token_text(_at(toks, 0)) == "gosub"
            or any(token_name(tok) is not None and token_text(tok) in callables for tok in toks)
        ):
            state.clear()
        if isinstance(node, StatementNode) and node.single_line_if_branches:
            forget(names_in(source, node.span))
            return

        def at(tok: VbaToken) -> Span:
            return Span(node.span.start + tok.start, node.span.start + tok.end)

        if host == "Word":
            add = _word_add(toks)
            if add is not None and add.name == "" and add.key.endswith(".customdocumentproperties"):
                push(
                    "hostArgumentOutOfRange",
                    "A custom property needs a name, and \"\" is none. This will raise Run-time error "
                    "'-2147418113': Automation error.",
                    at(add.name_token),
                )
            if add is not None:
                held_names = state.get(add.key)
                if held_names is None:
                    held_names = set()
                name = add.name.lower()
                if name in held_names:
                    kind = _WORD_NAMED[add.key[add.key.rfind(".") + 1 :]]
                    push(
                        "hostArgumentOutOfRange",
                        f"The {kind.noun} \"{add.name}\" was already added to {add.chain}, and the names "
                        f"ignore case. This will raise Run-time error {kind.error}.",
                        at(add.name_token),
                    )
                head = add.key.split(".")[0]
                forget([lower for lower in names_in(source, node.span) if lower != head])
                held_names.add(name)
                state[add.key] = held_names
                return
            # `d.Variables("zq").Value = 3` changes the value, not the names;
            # `.Delete` takes the name out (XLIDE issue #610).
            use = _variable_use(toks)
            if use is not None and use.key in state:
                if use.deletes:
                    state[use.key].discard(use.name)
                head = use.key.split(".")[0]
                forget([lower for lower in names_in(source, node.span) if lower != head])
                return
        set_target = set_assignment_target(source, node.span)
        if set_target is not None:
            lower = set_target[0].lower()
            value = [tok for tok in set_target[2] if tok.kind is not TokenKind.COMMENT]
            # `Set d2 = d`: another name for the document, whose names it shares
            # (XLIDE issue #610).
            alias = _lower_name(value[0]) if len(value) == 1 else None
            shared = (
                [
                    (key, held)
                    for key, held in state.items()
                    if not key.startswith(_SLIDE) and key.split(".")[0] == alias
                ]
                if alias
                else []
            )
            call = next(
                (
                    k
                    for k, tok in enumerate(value)
                    if _raw(value, k - 1) == "."
                    and _at(value, k - 2) is not None
                    and token_text(_at(value, k - 2)) == "slides"
                    and token_text(tok) in ("add", "addslide")
                ),
                -1,
            )
            # `Set a = p.Slides.Add(1, ...)` keeps what is known of p.
            pres = (
                _lower_name(value[0])
                if call == 4 and value[1].raw_text == "." and value[3].raw_text == "."
                else None
            )
            forget([name for name in names_in(source, node.span) if name != pres or name == lower])
            for key, held in shared:
                state[key] = held
                assert alias is not None
                state[lower + key[len(alias) :]] = held
            if (
                host == "PowerPoint"
                and call > 0
                and _raw(value, call + 1) == "("
                and match_paren_from(value, call + 1) == len(value) - 1
            ):
                state[_SLIDE + lower] = set()
                # `p.Slides.Add(1, ...)` on a presentation the procedure made.
                if pres and pres in presentations:
                    index = _at(value, call + 2)
                    add_slide(
                        pres,
                        js_number(index.raw_text)
                        if index is not None and index.kind is TokenKind.INTEGER_LITERAL
                        else None,
                        _SLIDE + lower,
                        index,
                        at,
                    )
            # `Set p = Presentations.Add(...)`: a new presentation, no slides.
            if (
                host == "PowerPoint"
                and token_text(_at(value, 0)) == "presentations"
                and _raw(value, 1) == "."
                and token_text(_at(value, 2)) == "add"
                and (
                    len(value) == 3
                    or (_raw(value, 3) == "(" and match_paren_from(value, 3) == len(value) - 1)
                )
            ):
                presentations[lower] = _PresentationState()
            return
        if host == "PowerPoint":
            # `p.Slides.Add 3, ppLayoutBlank` as a statement.
            pres = _lower_name(_at(toks, 0))
            if (
                pres
                and pres in presentations
                and _raw(toks, 1) == "."
                and token_text(_at(toks, 2)) == "slides"
                and _raw(toks, 3) == "."
                and token_text(_at(toks, 4)) == "add"
                and _raw(toks, 5) != "("
            ):
                index = _at(toks, 5)
                add_slide(
                    pres,
                    js_number(index.raw_text)
                    if index is not None and index.kind is TokenKind.INTEGER_LITERAL
                    else None,
                    f"{_SLIDE}#{node.span.start}",
                    index,
                    at,
                )
                return
            # `p.Slides(2).Name = "zq"`.
            if (
                pres
                and pres in presentations
                and _raw(toks, 1) == "."
                and token_text(_at(toks, 2)) == "slides"
                and _raw(toks, 3) == "("
                and _kind(toks, 4) is TokenKind.INTEGER_LITERAL
                and _raw(toks, 5) == ")"
                and _raw(toks, 6) == "."
                and token_text(_at(toks, 7)) == "name"
                and _raw(toks, 8) == "="
                and len(toks) == 10
                and toks[9].kind is TokenKind.STRING_LITERAL
            ):
                slide_key = slide_at(pres, js_number(toks[4].raw_text))
                if slide_key:
                    name_slide(slide_key, string_literal_value(toks[9].raw_text), toks[9], at)
                    return
            # `a.MoveTo 2` past the slides of the presentation a is in.
            moved = _lower_name(_at(toks, 0))
            moved_in = slide_in.get(_SLIDE + moved) if moved else None
            moved_pres = presentations.get(moved_in) if moved_in else None
            count = len(moved_pres.order) if moved_pres is not None else None
            if (
                count is not None
                and _raw(toks, 1) == "."
                and token_text(_at(toks, 2)) == "moveto"
                and len(toks) == 4
                and toks[3].kind is TokenKind.INTEGER_LITERAL
                and js_number(toks[3].raw_text) > count
            ):
                push(
                    "hostArgumentOutOfRange",
                    f"'{moved_in}' holds {count} slide{_plural(count)}, so {toks[3].raw_text} is past "
                    "the last. This will raise Run-time error '-2147188160': Integer out of range.",
                    at(toks[3]),
                )
                return
        # `a.Name = "Zq"` or `a.Name = b.Name` on a slide the procedure added.
        slide = (
            _lower_name(_at(toks, 0))
            if _raw(toks, 1) == "." and token_text(_at(toks, 2)) == "name" and _raw(toks, 3) == "="
            else None
        )
        if slide is not None and _SLIDE + slide in state and (len(toks) == 5 or len(toks) == 7):
            other = (
                _lower_name(toks[4])
                if len(toks) == 7 and toks[5].raw_text == "." and token_text(toks[6]) == "name"
                else None
            )
            other_name = next(iter(state.get(_SLIDE + other, set())), None) if other is not None else None
            name_value = (
                string_literal_value(toks[4].raw_text)
                if len(toks) == 5 and toks[4].kind is TokenKind.STRING_LITERAL
                else other_name
            )
            if name_value is not None:
                name_slide(_SLIDE + slide, name_value, toks[4], at)
            else:
                state[_SLIDE + slide] = set()
            return
        forget(names_in(source, node.span))

    # A block forgets the presentations; what it names it forgets too.
    def snapshot() -> dict[str, set[str]]:
        return {key: set(held) for key, held in state.items()}

    def restore(saved: Mapping[str, set[str]]) -> None:
        state.clear()
        presentations.clear()
        for key, held in saved.items():
            state[key] = set(held)

    def touches(stmt: LeafStatementNode) -> set[str]:
        return names_in(source, stmt.span)

    walk_entering_blocks(
        source,
        member.body,
        lambda node: activity is not None and activity.is_inactive(node.span),
        visit,
        BlockEnteringState(snapshot=snapshot, restore=restore, forget=forget, touches=touches),
    )


def _variable_use(toks: Sequence[VbaToken]) -> _VariableUse | None:
    """`d.Variables("zq").Value = 3`, `d.Variables("zq") = 3` or
    `d.Variables("zq").Delete` on a receiver of names and dots: the state key and
    the name, lower case, and whether it deletes."""
    at = next(
        (
            k
            for k, tok in enumerate(toks)
            if token_text(tok) == "variables"
            and _raw(toks, k - 1) == "."
            and _raw(toks, k + 1) == "("
            and k + 2 < len(toks)
            and toks[k + 2].kind is TokenKind.STRING_LITERAL
            and _raw(toks, k + 3) == ")"
        ),
        -1,
    )
    if at < 2:
        return None
    start = at - 2
    while start >= 2 and toks[start - 1].raw_text == "." and token_name(toks[start - 2]) is not None:
        start -= 2
    if start != 0 or token_name(toks[0]) is None:
        return None
    rest = [token_text(tok) or tok.raw_text for tok in toks[at + 4 :]]
    writes = (len(rest) >= 3 and rest[0] == "." and rest[1] == "value" and rest[2] == "=") or (
        len(rest) >= 1 and rest[0] == "="
    )
    deletes = len(rest) == 2 and rest[0] == "." and rest[1] == "delete"
    if not writes and not deletes:
        return None
    key = "".join(tok.raw_text for tok in toks[: at + 1]).lower()
    return _VariableUse(key, string_literal_value(toks[at + 2].raw_text).lower(), deletes)


def _word_add(toks: Sequence[VbaToken]) -> _WordAdd | None:
    """`d.Variables.Add "zq", 1`, `d.Styles.Add("Zq")` or
    `d.CustomDocumentProperties.Add Name:="zq", ...` with a literal name, on a
    receiver of names and dots: the state key, the receiver as written, and the
    name."""
    add = next(
        (
            k
            for k, tok in enumerate(toks)
            if token_text(tok) == "add"
            and _raw(toks, k - 1) == "."
            and token_text(_at(toks, k - 2)) in _WORD_NAMED
        ),
        -1,
    )
    if add < 3:
        return None
    # The receiver: names and dots from the statement's start, or after `=`.
    start = add - 2
    while start >= 2 and toks[start - 1].raw_text == "." and token_name(toks[start - 2]) is not None:
        start -= 2
    before = _raw(toks, start - 1)
    if start == add - 2 or (start != 0 and before != "=" and token_text(_at(toks, start - 1)) != "call"):
        return None
    open_index = add + 1 if _raw(toks, add + 1) == "(" else -1
    close = match_paren_from(toks, open_index) if open_index > 0 else len(toks)
    args: list[list[VbaToken]] = [[]]
    depth = 0
    for k in range(open_index + 1 if open_index > 0 else add + 1, close):
        raw = toks[k].raw_text
        depth += 1 if raw == "(" else -1 if raw == ")" else 0
        if raw == "," and depth == 0:
            args.append([])
            continue
        args[-1].append(toks[k])
    named = next(
        (arg for arg in args if _raw(arg, 1) == ":=" and token_text(_at(arg, 0)) == "name"),
        None,
    )
    arg = named[2:] if named is not None else None if _raw(args[0], 1) == ":=" else args[0]
    if arg is None or len(arg) != 1 or arg[0].kind is not TokenKind.STRING_LITERAL:
        return None
    chain = "".join(tok.raw_text for tok in toks[start : add - 1])
    return _WordAdd(chain.lower(), chain[: chain.rfind(".")], string_literal_value(arg[0].raw_text), arg[0])
