"""Rule family: Variant values used as the wrong kind of value.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/variantValues.ts.

Rule: a Variant local whose value the code makes plain, used as something that
value is not (XLIDE issue #121). Measured in Excel 16.0 (build 20326,
2026-09-26); each compiles and raises every time it runs.

 - A scalar used as an object: `v = 5` or `v = "abc"` then `v.Foo` -> 424,
   Object required. An array used as one: `v = Array(1, 2)` then `v.Foo`
   -> 424.
 - A scalar used as an array: `v = 5` or `v = "abc"` then `UBound(v)` -> 13.
 - An array used as a scalar: `v = Array(1, 2)` then `v + 1`, `v - 1`,
   `v & "x"`, `If v = 1 Then` -> 13, Type mismatch.
 - An array a call returns, used the same way (issue #239):
   `Array(1) + 1`, `Split("a") + 1`, `-Array(1)`, `Not Array(1)`.
 - A scalar where only an array will do (issue #219): `v = 5` then
   `Erase v`, `ReDim Preserve v(2)` or `For Each x In v` -> 13. For Each
   over a Variant nothing assigns, which is Empty, raises 13 too.

The values come from the same analysis the division and subscript rules use: a
literal, or an array from Array(), Split() on literals or a Range literal's
Value, that either every assignment in the procedure gives, or the last
assignment before the statement gives with nothing between able to change it
(issue #180).
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from typing import Generic, TypeVar

from ...conditional import ConditionalActivityTracker
from ...js_compat import JS_WHITESPACE, js_number_to_string, js_trim
from ...lexer.token_helpers import match_paren_from, top_level_equals_index
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ForBlockNode,
    IfBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    WithBlockNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ...symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbol, VbaSymbolKind
from ...types.type_inference import (
    known_local_literal_values_at,
    picked_values,
    procedure_symbol_for,
    type_environment_for,
)
from ...types.type_names import normalize_type
from ..callable_signatures import (
    SourceNameScope,
    runtime_callable_source_shadowed,
    source_name_scope_for,
)
from ..context import PushFn, statement_tokens
from ..known_locals import KnownLocalValue
from ..straight_line_values import straight_line_assignments
from ..walker import (
    active_module_members,
    bare_assignment_target,
    block_header_line_span,
    for_each_statement,
    is_inactive_node,
    statement_and_branch_spans,
    token_name,
    token_text,
)
from .arrays import FixedArrayBound, known_array_shapes_at, module_option_base, single_cell_value
from .shared import body_may_leave_loop, is_bare_or_vba_qualified_intrinsic_call, name_mentions

_SCALAR_OPERATORS = frozenset({"=", "<", ">", "<=", ">=", "<>", "+", "-", "*", "/", "\\", "&", "^"})

# The calls that return an array whatever their arguments.
_ARRAY_FUNCTIONS = frozenset({"array", "split"})

_ARRAY_CALL_WORD_OPERATORS = frozenset({"mod", "not", "and", "or", "xor", "like"})

_JS_SPACE = "[" + JS_WHITESPACE + "]"
_DECLARATION_LINE = re.compile("^" + _JS_SPACE + r"*(?:dim|static|const)(?![a-z0-9_])", re.IGNORECASE | re.ASCII)
_PLAIN_NAME = re.compile(r"[a-z_][a-z0-9_]*")

_CELL_VALUE = "one cell's value, which is no array"

_K = TypeVar("_K")
_V = TypeVar("_V")


class _MemoByIdentity(Generic[_K, _V]):
    """`derive` run once per distinct input object. Each key is held with its value,
    so an id cannot be reused while the memo lives."""

    __slots__ = ("_derive", "_cache")

    def __init__(self, derive: Callable[[_K], _V]) -> None:
        self._derive = derive
        self._cache: dict[int, tuple[_K, _V]] = {}

    def __call__(self, key: _K) -> _V:
        hit = self._cache.get(id(key))
        if hit is not None and hit[1] is not None:
            return hit[1]
        value = self._derive(key)
        self._cache[id(key)] = (key, value)
        return value


@dataclass(frozen=True, slots=True)
class _Guard:
    lower: str
    start: int
    end: int


@dataclass(frozen=True, slots=True)
class _AfterEach:
    lower: str
    start: int
    until: int


@dataclass(frozen=True, slots=True)
class _Hit:
    start: int
    end: int
    message: str


def check_variant_value_misuse(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    project_visible_symbols: Sequence[VbaSymbol] | None = None,
) -> None:
    option_base = module_option_base(mod, activity)
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_procedure(source, member, symbols, activity, option_base, push, project_visible_symbols)


def _check_procedure(
    source: str,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    option_base: int,
    push: PushFn,
    project_visible_symbols: Sequence[VbaSymbol] | None,
) -> None:
    env = type_environment_for(symbols, member)

    def is_variant(lower: str) -> bool:
        type_ = normalize_type(env.get(lower))
        return type_ is None or type_ == "variant"

    # What each Variant holds at a statement (issue #180): the last assignment to
    # reach it, or the one value the procedure agrees on.
    values_at = known_local_literal_values_at(source, member, symbols, activity)
    shapes_at = known_array_shapes_at(source, symbols, member, activity, option_base)

    def pick_scalar(lower: str, value: KnownLocalValue) -> str | None:
        if value.origin == "literal" and is_variant(lower):
            return _literal_text(value)
        return None

    scalars_for: _MemoByIdentity[Mapping[str, KnownLocalValue], Mapping[str, str]] = _MemoByIdentity(
        lambda values: picked_values(values, pick_scalar)
    )
    # One cell's value is no array, whatever it holds: `v = Range("A1").Value`
    # then `v(1, 1)` or `UBound(v)` raises 13 (issue #278, measured in Excel 16.0).
    reaching_cache: list[Mapping[int, Mapping[str, Sequence[VbaToken]]]] = []

    def held_tokens(stmt: BodyNode, lower: str) -> list[VbaToken] | None:
        """The tokens the straight line last stored in `lower` before `stmt`,
        comments dropped, or None."""
        if not reaching_cache:
            reaching_cache.append(straight_line_assignments(source, member.body, activity))
        # Upstream's Map keyed by node: the port keys node maps by id().
        held = reaching_cache[0].get(id(stmt))
        toks = held.get(lower) if held is not None else None
        return None if toks is None else [tok for tok in toks if tok.kind is not TokenKind.COMMENT]

    # Asked by name, as the scalars are (issue #322).
    def cell_scalar_at(stmt: BodyNode, lower: str) -> str | None:
        toks = held_tokens(stmt, lower)
        return _CELL_VALUE if toks is not None and is_variant(lower) and single_cell_value(toks) else None

    def derive_arrays(shapes: Mapping[str, FixedArrayBound]) -> dict[str, str]:
        return {lower: shape.origin for lower, shape in shapes.items() if is_variant(lower)}

    arrays_for: _MemoByIdentity[Mapping[str, FixedArrayBound], dict[str, str]] = _MemoByIdentity(derive_arrays)

    # A Variant local is still Empty at a statement that names it first, though
    # Erase and ReDim end what the value analysis knows of it: the only statement
    # to name it, or the first in a procedure with no GoTo, GoSub or Resume to
    # come back to an earlier line. Inside a block too: the first pass through a
    # loop, or the arm that runs, reaches it before anything else names it
    # (issue #237).
    proc_sym = procedure_symbol_for(symbols, member)
    proc_children = (proc_sym.children if proc_sym is not None else None) or []
    mentions_cache: list[Mapping[str, int]] = []
    # The procedure's tokens, each with its offset in the module.
    procedure_tokens: list[tuple[VbaToken, int]] | None = None

    def empty_here(lower: str, offset: int) -> bool:
        nonlocal procedure_tokens
        local = next((child for child in proc_children if child.name.lower() == lower), None)
        if (
            local is None
            or local.kind is not VbaSymbolKind.LOCAL_VARIABLE
            or local.is_array
            or not is_variant(lower)
        ):
            return False
        if not mentions_cache:
            mentions_cache.append(name_mentions(source, member, activity))
        if mentions_cache[0].get(lower) == 1:
            return True
        # A Static keeps what an earlier call left; one no other statement names
        # is Empty on every call (issue #612, measured in Excel 16.0).
        if local.visibility is SymbolVisibility.STATIC:
            return False
        if procedure_tokens is None:
            procedure_tokens = [
                (tok, tok.start + member.span.start) for tok in statement_tokens(source, member.span)
            ]
        if any(token_text(tok) in ("goto", "gosub", "resume") for tok, _ in procedure_tokens):
            return False
        # The first use after the declaration (a Dim names it too).
        first_use = next(
            (
                start
                for tok, start in procedure_tokens
                if (token_name(tok) or "").lower() == lower
                and token_name(tok) is not None
                and not _is_in_declaration(source, start)
            ),
            None,
        )
        return first_use == offset

    source_names = source_name_scope_for(symbols, member, project_visible_symbols)

    # `If IsObject(v) Then`: v is an object in the arm that runs (issue #612). The
    # arm's lines are not judged on v.
    guards: list[_Guard] = []

    def guards_in(condition: Sequence[VbaToken], start: int, end: int) -> None:
        i = 0
        while i + 3 < len(condition):
            name = token_name(condition[i + 2])
            lower = name.lower() if name else None
            if (
                token_text(condition[i]) == "isobject"
                and condition[i + 1].raw_text == "("
                and lower
                and condition[i + 3].raw_text == ")"
                and token_text(_at(condition, i - 1)) != "not"
            ):
                guards.append(_Guard(lower, start, end))
            i += 1

    # Upstream's recursive visitGuards on an explicit stack; the guards are a set
    # asked with `some`, so their order does not matter.
    guard_bodies: list[Sequence[BodyNode]] = [member.body]
    while guard_bodies:
        for node in guard_bodies.pop():
            if isinstance(node, IfBlockNode):
                for k, branch in enumerate(node.branches):
                    following = (
                        node.branches[k + 1].header_span.start if k + 1 < len(node.branches) else node.span.end
                    )
                    guards_in(statement_tokens(source, branch.header_span), branch.header_span.end, following)
                    guard_bodies.append(branch.body)
                continue
            if isinstance(node, StatementNode) and node.single_line_if_branches:
                toks = statement_tokens(source, node.span)
                then = next((k for k, tok in enumerate(toks) if token_text(tok) == "then"), -1)
                if then > 0:
                    first_branch = node.single_line_if_branches[0]
                    guards_in(toks[:then], first_branch.start, first_branch.end)
            body = getattr(node, "body", None)
            if isinstance(body, list):
                guard_bodies.append(body)

    def guarded(lower: str, offset: int) -> bool:
        return any(guard.lower == lower and guard.start <= offset < guard.end for guard in guards)

    # Null or an error value a straight line has just put in v (issue #612).
    def held_special(stmt: BodyNode, lower: str) -> str | None:
        held = held_tokens(stmt, lower)
        if held is None or not is_variant(lower):
            return None
        if len(held) == 1 and token_text(held[0]) == "null":
            return "Null"
        if token_text(_at(held, 0)) == "cverr" and _raw_at(held, 1) == "(":
            return "an error value from CVErr"
        return None

    # `For Each v In c` over a Collection leaves v Empty when it ends (issue #612,
    # measured in Excel 16.0), until the next line that names v.
    after_each: list[_AfterEach] = []
    each_bodies: list[Sequence[BodyNode]] = [member.body]
    while each_bodies:
        body_list = each_bodies.pop()
        for k, node in enumerate(body_list):
            lower = (
                node.control_variable.lower()
                if isinstance(node, ForBlockNode) and node.each and node.control_variable
                else None
            )
            over = (
                js_trim(node.source_expression).lower()
                if isinstance(node, ForBlockNode) and node.source_expression is not None
                else None
            )
            if (
                lower
                and over
                and isinstance(node, ForBlockNode)
                and is_variant(lower)
                and (normalize_type(env.get(over)) or "") in ("collection", "vba.collection")
                and not body_may_leave_loop(source, node.body)
            ):
                mention = re.compile(r"\b" + re.escape(lower) + r"\b", re.IGNORECASE | re.ASCII)
                later = next(
                    (
                        candidate
                        for candidate in body_list[k + 1 :]
                        if mention.search(source[candidate.span.start : candidate.span.end]) is not None
                    ),
                    None,
                )
                after_each.append(
                    _AfterEach(lower, node.span.end, later.span.end if later is not None else member.span.end)
                )
            if isinstance(node, IfBlockNode):
                for branch in node.branches:
                    each_bodies.append(branch.body)
            else:
                body = getattr(node, "body", None)
                if isinstance(body, list):
                    each_bodies.append(body)

    def empty_after_each(lower: str, offset: int) -> bool:
        return any(each.lower == lower and each.start < offset < each.until for each in after_each)

    # `With v` with v a number, a string or Empty, and a member access as the
    # first statement inside (issue #325, measured in Excel 16.0: 424 there; an
    # empty With runs).
    for node in iter_body_nodes(member.body, lambda block: is_inactive_node(activity, block)):
        if not isinstance(node, WithBlockNode):
            continue
        header = block_header_line_span(source, node.span)
        toks = statement_tokens(source, header)
        with_name = token_name(toks[1]) if len(toks) == 2 and token_text(toks[0]) == "with" else None
        lower = with_name.lower() if with_name else None
        first = next(
            (child for child in node.body if not is_inactive_node(activity, child) and is_leaf_statement(child)),
            None,
        )
        member_first = first is not None and _raw_at(statement_tokens(source, first.span), 0) == "."
        if not lower or not is_variant(lower) or not member_first:
            continue
        scalar = scalars_for(values_at(node)).get(lower)
        if scalar is None:
            scalar = cell_scalar_at(node, lower)
        if scalar is None and empty_after_each(lower, header.start):
            scalar = "Empty, as the For Each above left it"
        at = Span(header.start + toks[1].start, header.start + toks[1].end)
        if scalar:
            push(
                "variantValueMisuse",
                f"'{toks[1].raw_text}' holds {scalar} here, not an object for With to reach members of. "
                "This will raise Run-time error '424': Object required.",
                at,
            )
        elif empty_here(lower, at.start):
            push(
                "variantValueMisuse",
                f"'{toks[1].raw_text}' is never assigned, so it is Empty here, not an object for With to "
                "reach members of. This will raise Run-time error '424': Object required.",
                at,
            )

    # `a(v)` with v a Variant the straight line just set to Null: an index must be
    # a number (issue #332, measured in Excel 16.0: 94).
    array_names = {child.name.lower() for child in proc_children if child.is_array}

    def visit(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt) if array_names else []:
            toks = statement_tokens(source, span)
            for i in range(2, len(toks)):
                name = token_name(toks[i])
                lower = name.lower() if name else None
                if not lower or toks[i - 1].raw_text not in ("(", ",") or not is_variant(lower):
                    continue
                open_at = _last_index(toks, i, "(")
                array_name = token_name(_at(toks, open_at - 1))
                array = array_name.lower() if array_name else None
                close = match_paren_from(toks, open_at) if open_at >= 0 else -1
                whole = _raw_at(toks, i + 1) in (")", ",")
                if (
                    not array
                    or array not in array_names
                    or _raw_at(toks, open_at - 2) == "."
                    or close < i
                    or not whole
                ):
                    continue
                held = held_tokens(stmt, lower)
                if held is not None and len(held) == 1 and token_text(held[0]) == "null":
                    push(
                        "variantValueMisuse",
                        f"'{toks[i].raw_text}' holds Null here, and an index of '{toks[open_at - 1].raw_text}' "
                        "must be a number. This will raise Run-time error '94': Invalid use of Null.",
                        Span(span.start + toks[i].start, span.start + toks[i].end),
                    )
        for span in statement_and_branch_spans(stmt):
            for hit in _array_call_operands(statement_tokens(source, span), source_names):
                push("variantValueMisuse", hit.message, Span(span.start + hit.start, span.start + hit.end))
        # `v Is Nothing` with v still Empty (issue #325, measured: 424).
        for span in statement_and_branch_spans(stmt):
            toks = statement_tokens(source, span)
            for i, tok in enumerate(toks):
                name = None if _raw_at(toks, i - 1) == "." else token_name(tok)
                lower = name.lower() if name else None
                beside_is = token_text(_at(toks, i + 1)) == "is" or token_text(_at(toks, i - 1)) == "is"
                offset = span.start + tok.start
                if (
                    lower
                    and beside_is
                    and token_text(_at(toks, i - 1)) != "typeof"
                    and not guarded(lower, offset)
                    and (empty_here(lower, offset) or empty_after_each(lower, offset))
                ):
                    push(
                        "variantValueMisuse",
                        f"'{tok.raw_text}' is never assigned, so it is Empty here, not an object for Is to "
                        "compare. This will raise Run-time error '424': Object required.",
                        Span(offset, span.start + tok.end),
                    )
        for span in statement_and_branch_spans(stmt):
            toks = statement_tokens(source, span)
            head = token_text(_at(toks, 0))
            if head != "erase" and head != "redim":
                continue
            for i in range(1, len(toks)):
                name = token_name(toks[i])
                lower = name.lower() if name else None
                statement = _array_statement_target(toks, i) if lower else None
                if lower and statement and empty_here(lower, span.start + toks[i].start):
                    push(
                        "variantValueMisuse",
                        f"'{toks[i].raw_text}' is never assigned, so it is Empty here, which is not an array "
                        f"for {statement} to act on. This will raise Run-time error '13': Type mismatch.",
                        Span(span.start + toks[i].start, span.start + toks[i].end),
                    )
        # Looked up by the names the statement uses: asking the scalars' size
        # would walk every local (issue #322).
        literals = scalars_for(values_at(stmt))
        arrays = arrays_for(shapes_at(stmt))
        for span in statement_and_branch_spans(stmt):
            _check_span(
                source, span, stmt, is_variant, literals, arrays, cell_scalar_at, held_special, guarded,
                source_names, push,
            )

    for_each_statement(member.body, visit, activity)

    # `For Each x In v` with v a scalar or Empty (issue #219).
    for node in iter_body_nodes(member.body, lambda block: is_inactive_node(activity, block)):
        if not isinstance(node, ForBlockNode) or not node.each:
            continue
        lower = js_trim(node.source_expression).lower() if node.source_expression is not None else None
        if (
            not lower
            or _PLAIN_NAME.fullmatch(lower) is None
            or not is_variant(lower)
            or not node.source_expression_span
        ):
            continue
        value = values_at(node).get(lower)
        if value is not None and value.kind == "empty":
            holds: str | None = "nothing (it is never assigned, so it is Empty)"
        elif value is not None and value.origin == "literal":
            holds = _literal_text(value)
        else:
            holds = cell_scalar_at(node, lower)
        if holds:
            push(
                "variantValueMisuse",
                f"'{js_trim(node.source_expression or '')}' holds {holds} here, which For Each cannot step "
                "through. This will raise Run-time error '13': Type mismatch.",
                node.source_expression_span,
            )


def _check_span(
    source: str,
    span: Span,
    stmt: LeafStatementNode,
    is_variant: Callable[[str], bool],
    literals: Mapping[str, str],
    arrays: Mapping[str, str],
    cell_scalar_at: Callable[[BodyNode, str], str | None],
    held_special: Callable[[BodyNode, str], str | None],
    guarded: Callable[[str, int], bool],
    source_names: SourceNameScope,
    push: PushFn,
) -> None:
    toks = statement_tokens(source, span)
    target = bare_assignment_target(source, span)
    target_index = (
        next((i for i, tok in enumerate(toks) if tok.raw_text == "="), -1) - 1 if target is not None else -1
    )
    for i, tok in enumerate(toks):
        if i == target_index or _raw_at(toks, i - 1) == ".":
            continue
        name = token_name(tok)
        lower = name.lower() if name else None
        if not lower:
            continue
        at = Span(span.start + tok.start, span.start + tok.end)
        # `If IsObject(v) Then` holds an object (issue #612).
        if guarded(lower, at.start):
            continue
        nxt = _at(toks, i + 1)
        # `v Is Nothing` on a number, a string, an array, Null or an error value
        # (issues #325 and #612, measured: 424). `TypeOf v Is Collection` is False
        # on any of them.
        is_operand = (token_text(nxt) == "is" and token_text(_at(toks, i - 1)) != "typeof") or (
            token_text(_at(toks, i - 1)) == "is" and token_text(_at(toks, i - 2)) != "typeof"
        )
        special = held_special(stmt, lower) if is_operand else None
        if special:
            push(
                "variantValueMisuse",
                f"'{tok.raw_text}' holds {special} here, not an object for Is to compare. "
                "This will raise Run-time error '424': Object required.",
                at,
            )
            continue
        cell = cell_scalar_at(stmt, lower)
        scalar = cell if cell is not None else literals.get(lower)
        array = arrays.get(lower)
        if not scalar and not array:
            continue
        if is_operand:
            holds = scalar if scalar is not None else f"an array from {array}"
            push(
                "variantValueMisuse",
                f"'{tok.raw_text}' holds {holds} here, not an object for Is to compare. "
                "This will raise Run-time error '424': Object required.",
                at,
            )
            continue
        if nxt is not None and nxt.raw_text == "." and token_name(_at(toks, i + 2)):
            holds = scalar if scalar is not None else f"an array from {array}"
            push(
                "variantValueMisuse",
                f"'{tok.raw_text}' holds {holds} here, which has no members. "
                "This will raise Run-time error '424': Object required.",
                at,
            )
            continue
        if cell is not None and nxt is not None and nxt.raw_text == "(":
            push(
                "variantValueMisuse",
                f"'{tok.raw_text}' holds {scalar} here, so it has no element to index. "
                "This will raise Run-time error '13': Type mismatch.",
                at,
            )
            continue
        if scalar and _is_bound_argument(toks, i):
            push(
                "variantValueMisuse",
                f"'{tok.raw_text}' holds {scalar} here, which is not an array. "
                "This will raise Run-time error '13': Type mismatch.",
                at,
            )
            continue
        array_statement = _array_statement_target(toks, i) if scalar else None
        if scalar and array_statement:
            push(
                "variantValueMisuse",
                f"'{tok.raw_text}' holds {scalar} here, which is not an array for {array_statement} to act on. "
                "This will raise Run-time error '13': Type mismatch.",
                at,
            )
            continue
        length = _len_call_around(toks, i, i, source_names) if array else None
        if length:
            push(
                "variantValueMisuse",
                f"'{tok.raw_text}' holds an array from {array} here, which {length} cannot measure. "
                "This will raise Run-time error '13': Type mismatch.",
                at,
            )
            continue
        if array and (nxt is None or nxt.raw_text != "("):
            # The operator on either side, never the assignment's own `=`.
            previous = None if i - 1 == target_index + 1 else _at(toks, i - 1)
            operator = next(
                (
                    side
                    for side in (nxt, previous)
                    if side is not None
                    and (
                        (side.kind is TokenKind.OPERATOR and side.raw_text in _SCALAR_OPERATORS)
                        or token_text(side) == "mod"
                    )
                ),
                None,
            )
            if operator is not None:
                push(
                    "variantValueMisuse",
                    f"'{tok.raw_text}' holds an array from {array} here, which "
                    f"'{operator.raw_text}' cannot combine with a scalar. "
                    "This will raise Run-time error '13': Type mismatch.",
                    at,
                )


def _array_call_operands(toks: Sequence[VbaToken], source_names: SourceNameScope) -> list[_Hit]:
    """`Array(1) + 1`, `Split("a") & "x"`, `-Array(1)` and `Not Array(1)`: an array a
    call returns, as the operand of a scalar operator. Each raises 13 (issue #239,
    measured in Excel 16.0). Offsets are the statement's."""
    out: list[_Hit] = []
    # A single-line If's own line is its condition; each branch comes as a span of its own.
    condition = token_text(_at(toks, 0)) == "if"
    end = next((k for k, tok in enumerate(toks) if token_text(tok) == "then"), -1) if condition else len(toks)
    assignment = -1 if condition else top_level_equals_index(toks)
    for i in range(end - 1):
        name = token_text(toks[i])
        if (
            name not in _ARRAY_FUNCTIONS
            or toks[i + 1].raw_text != "("
            or not is_bare_or_vba_qualified_intrinsic_call(toks, i)
        ):
            continue
        qualified = _raw_at(toks, i - 1) == "."
        if not qualified and runtime_callable_source_shadowed(toks[i].raw_text, source_names):
            continue
        close = match_paren_from(toks, i + 1)
        # `Split(s)(0)` indexes the array, and its element is a scalar.
        if close < 0 or _raw_at(toks, close + 1) == "(":
            continue
        first = i - 2 if qualified else i
        # A statement's own `=` assigns; any other `=` compares.
        before = None if first - 1 == assignment else _at(toks, first - 1)
        after = _at(toks, close + 1)
        operator = next(
            (
                side
                for side in (after, before)
                if side is not None
                and (
                    (side.kind is TokenKind.OPERATOR and side.raw_text in _SCALAR_OPERATORS)
                    or token_text(side) in _ARRAY_CALL_WORD_OPERATORS
                )
            ),
            None,
        )
        call = "".join(tok.raw_text for tok in toks[first : close + 1])
        length = None if operator is not None else _len_call_around(toks, first, close, source_names)
        if length:
            out.append(
                _Hit(
                    toks[first].start,
                    toks[close].end,
                    f"{call} returns an array, which {length} cannot measure. "
                    "This will raise Run-time error '13': Type mismatch.",
                )
            )
            continue
        if operator is None:
            continue
        out.append(
            _Hit(
                toks[first].start,
                toks[close].end,
                f"{call} returns an array, which '{operator.raw_text}' cannot use as a scalar. "
                "This will raise Run-time error '13': Type mismatch.",
            )
        )
    return out


def _len_call_around(
    toks: Sequence[VbaToken], first: int, last: int, source_names: SourceNameScope
) -> str | None:
    """The Len or LenB whose one argument runs from `first` to `last`: an array value
    there raises 13 (issue #248, measured in Excel 16.0)."""
    callee = first - 2
    word = token_text(_at(toks, callee))
    if (
        (word != "len" and word != "lenb")
        or _raw_at(toks, first - 1) != "("
        or _raw_at(toks, last + 1) != ")"
        or not is_bare_or_vba_qualified_intrinsic_call(toks, callee)
    ):
        return None
    if _raw_at(toks, callee - 1) == ".":
        qualifier = _at(toks, callee - 2)
        return f"{qualifier.raw_text if qualifier is not None else 'undefined'}.{toks[callee].raw_text}"
    return None if runtime_callable_source_shadowed(toks[callee].raw_text, source_names) else toks[callee].raw_text


def _is_in_declaration(source: str, offset: int) -> bool:
    """Whether the offset is on a Dim, Static or Const line, which declares rather
    than uses."""
    line_start = source.rfind("\n", 0, offset) + 1
    return _DECLARATION_LINE.search(source[line_start:offset]) is not None


def _array_statement_target(toks: Sequence[VbaToken], i: int) -> str | None:
    """The statement a name is the target of, where only an array will do: `Erase v`
    and `ReDim Preserve v(2)` raise 13 on a scalar (issue #219, measured in Excel
    16.0). A plain ReDim makes v an array and runs."""
    head = token_text(_at(toks, 0))
    previous = _raw_at(toks, i - 1)
    if head == "erase" and (i == 1 or previous == ","):
        return "Erase"
    if (
        head == "redim"
        and token_text(_at(toks, 1)) == "preserve"
        and (i == 2 or previous == ",")
        and _raw_at(toks, i + 1) == "("
    ):
        return "ReDim Preserve"
    return None


def _is_bound_argument(toks: Sequence[VbaToken], i: int) -> bool:
    """True when `toks[i]` is the whole first argument of UBound, LBound, Join (issue
    #239) or Filter (issue #476)."""
    name = token_text(_at(toks, i - 2))
    return (
        _raw_at(toks, i - 1) == "("
        and name in ("ubound", "lbound", "join", "filter")
        and _raw_at(toks, i - 3) != "."
        and (_raw_at(toks, i + 1) == ")" or _raw_at(toks, i + 1) == ",")
    )


def _literal_text(value: KnownLocalValue) -> str:
    if value.kind == "string":
        return f'the string "{value.value}"'
    return f"the number {_number_text(value.value)}"


def _number_text(value: int | float | str) -> str:
    # A number-kind value is always numeric; the str arm only satisfies the type.
    return value if isinstance(value, str) else js_number_to_string(value)


def _last_index(toks: Sequence[VbaToken], before: int, raw: str) -> int:
    """`toks.slice(0, before).map(raw).lastIndexOf(raw)`."""
    for k in range(min(before, len(toks)) - 1, -1, -1):
        if toks[k].raw_text == raw:
            return k
    return -1


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None
