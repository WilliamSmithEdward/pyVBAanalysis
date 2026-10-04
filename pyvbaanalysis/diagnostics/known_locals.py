"""Locals whose value the procedure's text fixes, and the straight-line walk starts.

Ported from xlide_vscode/src/analyzer/diagnostics/typeInference.ts:
knownLocalLiteralValues and knownLocalLiteralValuesAt with their helpers (the
module-variable defaults, the Static locals, the per-statement views), the
walk starts the value rules share (declared defaults, Empty Variants, Nothing
objects, Consts, call effects and declared facts), unreachableStatementsIn,
deadBranchSpansIn, defaultedStraightLine, statementMayChangeModuleVariable and
functionResultFor (XLIDE issues #118, #119, #180, #241, #273, #348, #449, #562,
#618 and the others cited below). types/type_inference.py re-exports the public
names at the locations the sync map gives them.

A local nothing ever assigns holds its default, 0 for a number and "" for a
String, and one whose every assignment is the same literal holds that literal.
The overflow, runtime-value, variant-value and expression rules read the map to
prove a value the declared type alone leaves open.

Statements are unhashable dataclasses, so node-keyed results use id(node), as
straight_line_values does.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Iterator, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass, field
from itertools import chain
from typing import Generic, TypeVar

from ..conditional import ConditionalActivityTracker, inactive_node_skip
from ..constants.integer_constant_expression import (
    bankers_round,
    evaluate_integer_constant_expression,
    parse_vba_integer_literal,
)
from ..identity_cache import IdentityLru
from ..lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import (
    BodyNode,
    ForBlockNode,
    IfBlockNode,
    ProcedureNode,
    ProcKind,
    Span,
    StatementNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ..parser.parse_module import parse_module
from ..runtime.vba_runtime import resolve_runtime_constant, resolve_runtime_function
from ..symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbolKind, is_procedure_kind
from ..types.type_inference import def_type_of, procedure_symbol_for
from ..types.type_names import is_known_scalar_type, is_numeric_type, normalize_type
from .call_extraction import string_literal_value, unwrap_outer_parens
from .const_expr import collect_module_literal_integer_constants
from .context import statement_tokens
from .module_state import untouched_module_variables_in, written_names_in
from .straight_line_values import (
    EMPTY_COLLECTION,
    OBJECT_NOTHING,
    VARIANT_EMPTY,
    CallEffects,
    DeclaredFacts,
    ReachingAssignments,
    identity_assignment,
    set_call_effects,
    set_declared_facts,
    straight_line_assignments,
    straight_line_dead_branches,
    straight_line_exit,
    straight_line_unreachable,
)
from .walker import (
    bare_assignment_target,
    block_footer_line_span,
    block_header_line_span,
    first_executable_token_index,
    raw_expression_tokens,
    statement_and_branch_spans,
    token_name,
    token_text,
)

_T = TypeVar("_T")
_V = TypeVar("_V")


@dataclass(frozen=True, slots=True)
class KnownLocalValue:
    """A local whose value the procedure's text fixes: its default, or one literal."""

    # "number" | "string" | "empty". "empty" is a Variant nothing ever assigns,
    # whose value 0 is Empty's as a number.
    kind: str
    value: int | float | str
    # "default" when nothing ever assigns it, "literal" when every assignment is
    # the same literal.
    origin: str
    # A `Mid(x, ...) = ` statement rewrites characters of the value without
    # changing its length, so the length is still known and the characters are not.
    content_mutated: bool = False


# A local's kind for the literal rules: "number", "string", or None for a Variant,
# which takes its kind from the literal it is given.


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name else None


def _per_procedure(store: IdentityLru, symbols: ModuleSymbols) -> dict[int, tuple[ProcedureNode, object]]:
    """Upstream's perProcedureCache: a table per ModuleSymbols, keyed by the
    procedure's id with the procedure kept alongside so the id stays its own."""
    table: dict[int, tuple[ProcedureNode, object]] | None = store.get(symbols)
    if table is None:
        table = store.put({}, symbols)
    return table


def _cached_for(table: dict[int, tuple[ProcedureNode, object]], proc: ProcedureNode) -> object | None:
    entry = table.get(id(proc))
    return entry[1] if entry is not None and entry[0] is proc else None


_MODULE_MEMBER_NAMES_CACHE = IdentityLru()


def _module_member_names(symbols: ModuleSymbols) -> frozenset[str]:
    """The lowercased names the module declares at its top level, which shadow the
    VBA library's."""
    cached = _MODULE_MEMBER_NAMES_CACHE.get(symbols)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    names = frozenset(child.name.lower() for child in symbols.root.children or [])
    return _MODULE_MEMBER_NAMES_CACHE.put(names, symbols)  # type: ignore[no-any-return]


@dataclass(slots=True)
class _Candidate:
    # "number" | "string"; None for a Variant no literal has given a kind yet.
    kind: str | None
    # Numbers are held as numbers, which compare the way upstream's String() forms
    # of them do: 1 and 1.0 are one literal.
    literals: set[int | float | str] = field(default_factory=set)
    mutated: bool = False
    content_mutated: bool = False


_SETTING_HEADS = frozenset({"set", "input", "get", "line", "erase"})


def known_local_literal_values(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> dict[str, KnownLocalValue]:
    """The locals of a procedure whose value is plain from the text, keyed by
    lowercased name.

    Anything that could change a local another way - passing it whole to a call
    (ByRef), a For counter, `Input #`/`Get #`/`Line Input #`, `Mid(x, ...) =`,
    ReDim, `Set` - drops it from the map, as does any assignment whose value is
    not a plain literal. Object locals are left out. A Variant nothing assigns is
    Empty, and module variables nothing writes hold their defaults.
    """
    # A Variant (or untyped) local has no kind until a literal gives it one;
    # literals of two kinds, or none, leave it unknown (XLIDE issue #121).
    candidates: dict[str, _Candidate] = {
        lower: _Candidate(kind) for lower, kind in _literal_value_locals(proc, symbols).items()
    }
    if not candidates:
        return _module_variable_defaults(source, proc, symbols)
    bytes_ = _byte_variables(proc, symbols)
    wholes = _whole_number_variables(proc, symbols)
    member_names = _module_member_names(symbols)

    def mutate(lower: str | None) -> None:
        entry = candidates.get(lower) if lower else None
        if entry is not None:
            entry.mutated = True

    def library_function_argument(toks: Sequence[VbaToken], at: int) -> bool:
        """A VBA library function assigns none of its arguments: `n = CLng(s)`."""
        depth = 0
        for j in range(at - 1, 0, -1):
            raw = toks[j].raw_text
            if raw == ")":
                depth += 1
            elif raw == "(":
                if depth != 0:
                    depth -= 1
                    continue
                # `Left$(` lexes as Left and a `$` (XLIDE issue #334).
                callee_at = j - 2 if toks[j - 1].raw_text == "$" else j - 1
                callee = _lower_name(_at(toks, callee_at))
                before = _at(toks, callee_at - 1)
                qualified = before is not None and before.raw_text == "."
                if (
                    not callee
                    or (qualified and token_text(_at(toks, callee_at - 2)) != "vba")
                    or (not qualified and callee in member_names)
                ):
                    return False
                # The library knows String only as String$.
                runtime = resolve_runtime_function(callee) or resolve_runtime_function(f"{callee}$")
                return runtime is not None and runtime.kind == "function"
        return False

    def mutate_whole_arguments(toks: Sequence[VbaToken], start: int, in_value: bool) -> None:
        """A whole name passed to any call may be ByRef: `Take d`, `Take(d)`,
        `Call Take(d)`, `x = Take(d)`. Only a name standing alone in an argument
        slot counts; `Take(d + 1)` copies. In a value a call takes parentheses:
        `x = 10 Mod d` passes nothing."""
        for i in range(start, len(toks)):
            name = _lower_name(toks[i])
            if not name or name not in candidates:
                continue
            prev = _at(toks, i - 1)
            nxt = _at(toks, i + 1)
            # `SetN n:=n` passes n too, by name (XLIDE issue #449).
            opens_slot = (
                (prev is not None and prev.raw_text in ("(", ",", ":="))
                or (
                    not in_value
                    and (prev is None or prev.kind is TokenKind.IDENTIFIER or prev.kind is TokenKind.KEYWORD)
                )
            )
            closes_slot = (
                nxt is None or nxt.raw_text in (")", ",", ":") or nxt.kind is TokenKind.COMMENT
            )
            if (
                opens_slot
                and closes_slot
                and not (prev is not None and prev.kind is TokenKind.OPERATOR and prev.raw_text != ":=")
                and not (nxt is not None and nxt.kind is TokenKind.OPERATOR)
                and not (in_value and library_function_argument(toks, i))
            ):
                mutate(name)

    for node in iter_body_nodes(proc.body, inactive_node_skip(activity)):
        if isinstance(node, ForBlockNode):
            mutate(node.control_variable.lower() if node.control_variable else None)
        if isinstance(getattr(node, "body", None), list):
            # A header passes a name ByRef as a value does: `If Take(1, d) = 0
            # Then`, `Loop While Take(d)`. An ElseIf line is a statement of the
            # If's flat body.
            for span in (block_header_line_span(source, node.span), block_footer_line_span(source, node.span)):
                mutate_whole_arguments(statement_tokens(source, span), 0, True)
            continue
        if not is_leaf_statement(node):
            continue
        for span in statement_and_branch_spans(node):
            toks = statement_tokens(source, span)
            first = first_executable_token_index(toks)
            head = token_text(_at(toks, first))
            bare = bare_assignment_target(source, span)
            # `d = d + 0` leaves d as it was (XLIDE issue #350).
            if bare is not None and identity_assignment(bare[0], bare[2]):
                continue
            if bare is not None:
                lower = bare[0].lower()
                entry = candidates.get(lower)
                if entry is not None:
                    value = [tok for tok in toks[first + 2 :] if tok.kind is not TokenKind.COMMENT]
                    kind = entry.kind
                    if kind is None:
                        unwrapped = unwrap_outer_parens(value)
                        kind = "string" if unwrapped and unwrapped[0].kind is TokenKind.STRING_LITERAL else "number"
                    literal = _plain_literal(value, kind, entry.kind is not None, lower in bytes_, lower in wholes)
                    if literal is None or (entry.kind is not None and entry.kind != kind):
                        entry.mutated = True
                    else:
                        entry.kind = kind
                        entry.literals.add(literal)
                # The value may pass a name ByRef: `w = Take(d)` (XLIDE issue #238).
                mutate_whole_arguments(toks, first + 2, True)
                continue
            # A ReDim writes its arrays alone: `ReDim a(n)` reads n (XLIDE #350).
            if head == "redim":
                start = first + 2 if token_text(_at(toks, first + 1)) == "preserve" else first + 1
                for group in split_top_level_token_groups(toks, start, ",", len(toks)):
                    mutate(_lower_name(next((tok for tok in group if tok.kind is not TokenKind.COMMENT), None)))
                continue
            if head in _SETTING_HEADS:
                for tok in toks:
                    mutate(_lower_name(tok))
                continue
            # `Mid$` lexes as Mid and a `$` of its own (XLIDE issue #327).
            after_head = _at(toks, first + 1)
            mid_open = first + 2 if after_head is not None and after_head.raw_text == "$" else first + 1
            opener = _at(toks, mid_open)
            if head in ("mid", "mid$") and opener is not None and opener.raw_text == "(":
                # `Mid(x, start, len) = value` rewrites characters of x and keeps
                # its length; anything else named in it is read.
                target = candidates.get(_lower_name(_at(toks, mid_open + 1)) or "")
                if target is not None:
                    target.content_mutated = True
                continue
            if head in ("lset", "rset"):
                for tok in toks:
                    mutate(_lower_name(tok))
                continue
            # A Case line reads values: `Case w` passes w to nothing, and only a
            # call in it can (XLIDE issue #268).
            mutate_whole_arguments(toks, first + 1 if head == "case" else 0, head == "case")

    out: dict[str, KnownLocalValue] = {}
    for lower, entry in candidates.items():
        if entry.mutated:
            continue
        if entry.kind is None:
            # A Variant nothing assigns is Empty, which divides as 0: `5 / v`
            # raises 11 and `v / v` 6 (XLIDE issue #219, measured in Excel 16.0).
            if not entry.content_mutated:
                out[lower] = KnownLocalValue("empty", 0, "default")
            continue
        if len(entry.literals) == 0:
            out[lower] = KnownLocalValue(
                entry.kind, 0 if entry.kind == "number" else "", "default", entry.content_mutated
            )
        elif len(entry.literals) == 1:
            (only,) = entry.literals
            out[lower] = KnownLocalValue(entry.kind, only, "literal", entry.content_mutated)
    for lower, default in _module_variable_defaults(source, proc, symbols).items():
        out.setdefault(lower, default)
    return out


def _module_variable_defaults(
    source: str, proc: ProcedureNode, symbols: ModuleSymbols
) -> dict[str, KnownLocalValue]:
    """The initial value of each module variable nothing writes, as a procedure
    sees it (XLIDE issue #241): 0 for a number, "" for a String, Empty for a
    Variant. A local or parameter of the same name hides it."""
    out: dict[str, KnownLocalValue] = {}
    for lower, variable in untouched_module_variables_in(source, symbols, proc).items():
        if variable.is_array:
            continue
        type_ = normalize_type(variable.as_type)
        if type_ is None or type_ == "variant":
            out[lower] = KnownLocalValue("empty", 0, "default")
        elif is_numeric_type(type_):
            out[lower] = KnownLocalValue("number", 0, "default")
        elif type_ == "string":
            out[lower] = KnownLocalValue("string", "", "default")
    return out


def _local_kind(type_: str | None) -> str | None:
    """'number', 'string', None for a Variant, or 'other'."""
    if type_ is None or type_ == "variant":
        return None
    if is_numeric_type(type_) or type_ in ("boolean", "date"):
        return "number"
    return "string" if type_ == "string" else "other"


def _literal_value_locals(proc: ProcedureNode, symbols: ModuleSymbols) -> dict[str, str | None]:
    """The locals whose literal value the rules follow, with their kind: 'number' or
    'string' from the declared type, None for a Variant, which takes its kind from
    the literal."""
    out: dict[str, str | None] = {}
    proc_sym = procedure_symbol_for(symbols, proc)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if (
            child.kind is not VbaSymbolKind.LOCAL_VARIABLE
            or child.is_array
            or child.visibility is SymbolVisibility.STATIC
        ):
            continue
        kind = _local_kind(normalize_type(child.as_type))
        if kind == "other" or child.fixed_length is not None:
            continue
        out[child.name.lower()] = kind
    return out


def _static_value_locals(proc: ProcedureNode, symbols: ModuleSymbols) -> dict[str, str | None]:
    """The Static locals _literal_value_locals leaves out, with the same kinds."""
    out: dict[str, str | None] = {}
    proc_sym = procedure_symbol_for(symbols, proc)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if (
            child.kind is not VbaSymbolKind.LOCAL_VARIABLE
            or child.is_array
            or child.visibility is not SymbolVisibility.STATIC
            or child.fixed_length is not None
        ):
            continue
        kind = _local_kind(normalize_type(child.as_type))
        if kind != "other":
            out[child.name.lower()] = kind
    return out


_WHOLE = "whole"


class _StatementValues(Mapping[str, KnownLocalValue]):
    """The values one statement sees: the procedure-wide values, with each local a
    straight-line assignment reaches read from that assignment instead. A name is
    worked out when it is first asked for, and the whole map only when something
    walks it (XLIDE issue #322)."""

    __slots__ = ("_whole", "_assignments", "_derive", "_keeps_whole", "_known", "_full")

    def __init__(
        self,
        whole: Mapping[str, KnownLocalValue],
        assignments: ReachingAssignments,
        derive: Callable[[str, Sequence[VbaToken]], KnownLocalValue | str | None],
        keeps_whole: Callable[[str], bool],
    ) -> None:
        self._whole = whole
        self._assignments = assignments
        # The value an assignment gives a name: None for none, "whole" for the
        # procedure-wide one.
        self._derive = derive
        # Whether a name no assignment reaches keeps its procedure-wide value.
        self._keeps_whole = keeps_whole
        self._known: dict[str, KnownLocalValue | None] = {}
        self._full: dict[str, KnownLocalValue] | None = None

    def get(self, lower: str, default: object = None) -> KnownLocalValue | None:  # type: ignore[override]
        if self._full is not None:
            return self._full.get(lower)
        if lower in self._known:
            return self._known[lower]
        # A Variant's starting Empty is for the guards (XLIDE issue #691): these
        # values keep reading what the procedure as a whole says of it.
        raw = self._assignments.get(lower)
        value = None if raw is VARIANT_EMPTY else raw
        derived: KnownLocalValue | str | None
        if value is not None:
            derived = self._derive(lower, value)
        else:
            derived = _WHOLE if self._keeps_whole(lower) else None
        out = self._whole.get(lower) if derived == _WHOLE else derived
        assert out is None or isinstance(out, KnownLocalValue)
        self._known[lower] = out
        return out

    def __getitem__(self, lower: str) -> KnownLocalValue:
        value = self.get(lower)
        if value is None:
            raise KeyError(lower)
        return value

    def __contains__(self, lower: object) -> bool:
        return isinstance(lower, str) and self.get(lower) is not None

    def __iter__(self) -> Iterator[str]:
        return iter(self._all())

    def __len__(self) -> int:
        return len(self._all())

    def _all(self) -> dict[str, KnownLocalValue]:
        if self._full is None:
            full = dict(self._whole)
            for lower in self._whole:
                reached = self._assignments.get(lower)
                if (reached is None or reached is VARIANT_EMPTY) and not self._keeps_whole(lower):
                    del full[lower]
            for lower, value in self._assignments.items():
                if value is VARIANT_EMPTY:
                    continue
                derived = self._derive(lower, value)
                if derived == _WHOLE:
                    continue
                if derived is None:
                    full.pop(lower, None)
                else:
                    assert isinstance(derived, KnownLocalValue)
                    full[lower] = derived
            self._full = full
        return self._full


class _PickedValues(Mapping[str, _T], Generic[_T]):
    """A view of a statement's values through `pick`, worked out per name on first
    use (XLIDE issue #322)."""

    __slots__ = ("_base", "_pick", "_known", "_full")

    def __init__(
        self, base: Mapping[str, KnownLocalValue], pick: Callable[[str, KnownLocalValue], _T | None]
    ) -> None:
        self._base = base
        self._pick = pick
        self._known: dict[str, _T | None] = {}
        self._full: dict[str, _T] | None = None

    def get(self, lower: str, default: object = None) -> _T | None:  # type: ignore[override]
        if lower in self._known:
            return self._known[lower]
        value = self._base.get(lower)
        out = None if value is None else self._pick(lower, value)
        self._known[lower] = out
        return out

    def __getitem__(self, lower: str) -> _T:
        value = self.get(lower)
        if value is None:
            raise KeyError(lower)
        return value

    def __contains__(self, lower: object) -> bool:
        return isinstance(lower, str) and self.get(lower) is not None

    def __iter__(self) -> Iterator[str]:
        return iter(self._all())

    def __len__(self) -> int:
        return len(self._all())

    def _all(self) -> dict[str, _T]:
        if self._full is None:
            full: dict[str, _T] = {}
            for lower, value in self._base.items():
                picked = self._pick(lower, value)
                if picked is not None:
                    full[lower] = picked
            self._full = full
        return self._full


def picked_values(
    values: Mapping[str, KnownLocalValue], pick: Callable[[str, KnownLocalValue], _T | None]
) -> Mapping[str, _T]:
    """A view of a statement's values through `pick`, worked out per name on first
    use: the rules that keep only a statement's strings or literals look up the
    names the statement uses (XLIDE issue #322)."""
    return _PickedValues(values, pick)


LocalValuesAt = Callable[[BodyNode | None], Mapping[str, KnownLocalValue]]

# Rules share one view of a procedure's values. Symbols own the cache so a
# changed module/project context cannot reuse the old analysis.
_LOCAL_VALUES_AT = IdentityLru()


@dataclass(frozen=True, slots=True)
class _SourceKeyed(Generic[_V]):
    source: str
    activity: ConditionalActivityTracker | None
    value: _V


def known_local_literal_values_at(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> LocalValuesAt:
    """known_local_literal_values at each statement (XLIDE issue #180). Where the
    last assignment to reach a statement in a straight line is a literal, the
    statement sees that literal, though other assignments in the procedure
    disagree with it: `d = 0: x = 10 / d: d = 2` divides by 0. Elsewhere it sees
    the procedure-wide value. With no statement (a block header) it is the
    procedure-wide map."""
    table = _per_procedure(_LOCAL_VALUES_AT, symbols)
    cached = _cached_for(table, proc)
    if isinstance(cached, _SourceKeyed) and cached.source == source and cached.activity is activity:
        return cached.value  # type: ignore[no-any-return]
    values_at = _build_known_local_literal_values_at(source, proc, symbols, activity)
    table[id(proc)] = (proc, _SourceKeyed(source, activity, values_at))
    return values_at


def _build_known_local_literal_values_at(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> LocalValuesAt:
    whole = known_local_literal_values(source, proc, symbols, activity)
    start_locals = _literal_value_locals(proc, symbols)
    # A Static local holds what the last call left as the procedure starts, but
    # what the straight line assigns it after that (XLIDE issue #685): it has no
    # start of its own and no procedure-wide value.
    locals_ = {**start_locals, **_static_value_locals(proc, symbols)}
    module_variables = _followed_module_variables(proc, symbols)
    bytes_ = _byte_variables(proc, symbols)
    wholes = _whole_number_variables(proc, symbols)
    writes: dict[str, list[BodyNode]] | None = None
    # The same start as unreachable_statements_in, so the two share one walk.
    reaching: Mapping[int, ReachingAssignments] = (
        {}
        if not locals_ and not module_variables
        else straight_line_assignments(
            source, proc.body, activity, _walk_start_with_effects(source, symbols, proc, start_locals, activity)
        )
    )
    # Statements in a run share one reaching map, so they share one result.
    results: dict[int, tuple[ReachingAssignments, Mapping[str, KnownLocalValue]]] = {}
    # A block's opening line, given as a statement of its own, sees what reaches
    # the block: `x = "abc"` then `While x` (XLIDE issue #424).
    blocks_by_start: dict[int, BodyNode] | None = None

    def block_at(stmt: BodyNode) -> BodyNode | None:
        nonlocal blocks_by_start
        if blocks_by_start is None:
            blocks_by_start = {
                node.span.start: node
                for node in iter_body_nodes(proc.body)
                if not is_leaf_statement(node) and id(node) in reaching
            }
        return blocks_by_start.get(stmt.span.start)

    # One assignment reaches many statements, its tokens shared by all of them:
    # what it gives a name is worked out once (XLIDE issue #322).
    derived: dict[int, tuple[Sequence[VbaToken], dict[str, KnownLocalValue | str | None]]] = {}

    def derive(lower: str, value: Sequence[VbaToken]) -> KnownLocalValue | str | None:
        entry = derived.get(id(value))
        if entry is None or entry[0] is not value:
            entry = (value, {})
            derived[id(value)] = entry
        by_name = entry[1]
        if lower in by_name:
            return by_name[lower]
        out: KnownLocalValue | str | None
        if lower not in locals_:
            out = _WHOLE
        else:
            declared = locals_.get(lower)
            kind = declared if declared is not None else _kind_of_value(value)
            literal = _plain_literal(value, kind, declared is not None, lower in bytes_, lower in wholes)
            origin = "default" if value is DEFAULT_NUMBER or value is DEFAULT_STRING else "literal"
            out = None if literal is None else KnownLocalValue(kind, literal, origin)
        by_name[lower] = out
        return out

    def keeps_whole(lower: str) -> bool:
        if locals_.get(lower) is None:
            return True
        held = whole.get(lower)
        return held is None or held.origin != "literal"

    def values_at(stmt: BodyNode | None) -> Mapping[str, KnownLocalValue]:
        nonlocal writes
        # No statement: a block header, which sees the procedure-wide values.
        assignments: ReachingAssignments | None = None
        if stmt is not None:
            assignments = reaching.get(id(stmt))
            if assignments is None and isinstance(stmt, StatementNode):
                block = block_at(stmt)
                assignments = reaching.get(id(block)) if block is not None else None
        if assignments is None or stmt is None:
            return whole
        kept = results.get(id(assignments))
        if kept is not None and kept[0] is assignments:
            result = kept[1]
        else:
            # A typed local starts the walk with its default, so one it no
            # longer holds may have kept that default on some path: with `If b >
            # 5000 Then a = 4` undecided, a is 0 or 4 after it, not the 4 it
            # holds wherever it is assigned (XLIDE issue #565).
            result = _StatementValues(whole, assignments, derive, keeps_whole)
            results[id(assignments)] = (assignments, result)
        if not module_variables:
            return result
        # A module variable written in the straight line holds the value while
        # nothing between could run other code (XLIDE issue #348).
        with_module: dict[str, KnownLocalValue] | None = None
        for lower, value in assignments.items():
            if lower not in module_variables or lower in locals_:
                continue
            declared = module_variables.get(lower)
            kind = declared if declared is not None else _kind_of_value(value)
            literal = _plain_literal(value, kind, declared is not None, lower in bytes_, lower in wholes)
            # The reaching write is the last one that runs before the statement:
            # a later one in a block would have ended the value.
            dead = unreachable_statements_in(source, proc, symbols, activity)
            if writes is None:
                writes = _module_variable_writes(source, proc, activity)
            write = next(
                (
                    node
                    for node in reversed(writes.get(lower, []))
                    if node.span.end <= stmt.span.start and id(node) not in dead
                ),
                None,
            )
            if literal is None or write is None:
                continue
            # A call to a procedure of this module that leaves the variable alone,
            # and runs no code but the module's own, keeps its value (#618).
            variable = lower
            if _code_may_run(
                statement_tokens(source, Span(write.span.end, stmt.span.end)),
                proc,
                symbols,
                lambda name: _callee_leaves_alone(source, symbols, name, variable),
            ):
                continue
            if with_module is None:
                with_module = dict(result)
            with_module[lower] = KnownLocalValue(kind, literal, "literal")
        return with_module if with_module is not None else result

    return values_at


def _kind_of_value(value: Sequence[VbaToken]) -> str:
    unwrapped = unwrap_outer_parens(value)
    return "string" if unwrapped and unwrapped[0].kind is TokenKind.STRING_LITERAL else "number"


class _NamesTest:
    """A has(lower) test over a procedure's locals, falling back to the module's."""

    __slots__ = ("_locals", "_module", "_test")

    def __init__(
        self,
        locals_: Mapping[str, object],
        module: AbstractSet[str],
        test: Callable[[object], bool],
    ) -> None:
        self._locals = locals_
        self._module = module
        self._test = test

    def __contains__(self, lower: object) -> bool:
        if not isinstance(lower, str):
            return False
        local = self._locals.get(lower)
        return self._test(local) if local is not None else lower in self._module


_MODULE_BYTE_VARIABLES = IdentityLru()
_MODULE_WHOLE_VARIABLES = IdentityLru()

# The whole-number types, which store a fraction rounded half to even: `a As Long
# = 4.4` holds 4.
_WHOLE_NUMBER_TYPES = frozenset({"byte", "integer", "long", "longlong", "longptr"})


def _procedure_children_by_name(proc: ProcedureNode, symbols: ModuleSymbols) -> dict[str, object]:
    proc_sym = procedure_symbol_for(symbols, proc)
    return {child.name.lower(): child for child in (proc_sym.children if proc_sym is not None else None) or []}


def _byte_variables(proc: ProcedureNode, symbols: ModuleSymbols) -> _NamesTest:
    """The locals and module variables a procedure sees that are declared As Byte,
    by lowercased name."""
    module: frozenset[str] | None = _MODULE_BYTE_VARIABLES.get(symbols)
    if module is None:
        module = _MODULE_BYTE_VARIABLES.put(
            frozenset(
                sym.name.lower()
                for sym in symbols.root.children or []
                if sym.kind is VbaSymbolKind.MODULE_VARIABLE and normalize_type(sym.as_type) == "byte"
            ),
            symbols,
        )
    assert module is not None
    return _NamesTest(
        _procedure_children_by_name(proc, symbols),
        module,
        lambda local: getattr(local, "kind", None) is VbaSymbolKind.LOCAL_VARIABLE
        and normalize_type(getattr(local, "as_type", None)) == "byte",
    )


def _whole_number_variables(proc: ProcedureNode, symbols: ModuleSymbols) -> _NamesTest:
    """The locals and module variables a procedure sees that are declared a
    whole-number type (XLIDE issue #685)."""
    module: frozenset[str] | None = _MODULE_WHOLE_VARIABLES.get(symbols)
    if module is None:
        module = _MODULE_WHOLE_VARIABLES.put(
            frozenset(
                sym.name.lower()
                for sym in symbols.root.children or []
                if sym.kind is VbaSymbolKind.MODULE_VARIABLE
                and (normalize_type(sym.as_type) or "") in _WHOLE_NUMBER_TYPES
            ),
            symbols,
        )
    assert module is not None
    return _NamesTest(
        _procedure_children_by_name(proc, symbols),
        module,
        lambda local: getattr(local, "kind", None) is VbaSymbolKind.LOCAL_VARIABLE
        and (normalize_type(getattr(local, "as_type", None)) or "") in _WHOLE_NUMBER_TYPES,
    )


def _followed_module_variables(proc: ProcedureNode, symbols: ModuleSymbols) -> dict[str, str | None]:
    """The module variables a procedure may follow through its straight line, with
    their kind as _literal_value_locals gives a local's. A local or parameter of
    the same name hides one."""
    proc_sym = procedure_symbol_for(symbols, proc)
    hidden = {
        name.lower()
        for name in chain(
            (proc.name,),
            (param.name for param in proc.params),
            (child.name for child in (proc_sym.children if proc_sym is not None else None) or []),
        )
    }
    out: dict[str, str | None] = {}
    for child in symbols.root.children or []:
        if (
            child.kind is not VbaSymbolKind.MODULE_VARIABLE
            or child.is_array
            or child.is_auto_instantiated
            or child.fixed_length is not None
            or child.name.lower() in hidden
        ):
            continue
        type_ = normalize_type(child.as_type)
        kind = (
            None
            if type_ is None or type_ == "variant"
            else "number"
            if is_numeric_type(type_)
            else "string"
            if type_ == "string"
            else "other"
        )
        if kind != "other":
            out[child.name.lower()] = kind
    return out


def _module_variable_writes(
    source: str, proc: ProcedureNode, activity: ConditionalActivityTracker | None
) -> dict[str, list[BodyNode]]:
    """The statements of a procedure that assign a name, `x = value`, in source
    order, by lowercased name. An If is read through its arms, not its flat body."""
    out: dict[str, list[BodyNode]] = {}
    stack: list[Iterator[BodyNode]] = [iter(proc.body)]
    while stack:
        for node in stack[-1]:
            if activity is not None and activity.is_inactive(node.span):
                continue
            if isinstance(node, IfBlockNode):
                stack.append(chain.from_iterable(branch.body for branch in node.branches))
                break
            body = getattr(node, "body", None)
            if isinstance(body, list):
                stack.append(iter(body))
                break
            bare = bare_assignment_target(source, node.span) if is_leaf_statement(node) else None
            if bare is not None:
                out.setdefault(bare[0].lower(), []).append(node)
        else:
            stack.pop()
    return out


# VBA functions that wait for the user or call by name, so other code may run
# meanwhile.
_CODE_RUNNING_FUNCTIONS = frozenset({"callbyname", "doevents", "inputbox", "msgbox"})


def _code_may_run(
    toks: Sequence[VbaToken],
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    leaves_alone: Callable[[str], bool] | None = None,
) -> bool:
    """Whether the tokens may run code other than the procedure's own, which could
    write a module variable: a call to a procedure, a member of an object, `New`
    of a class, RaiseEvent, or a ByRef parameter, which may be the module variable
    itself. The procedure's locals, its ByVal parameters, the module's variables
    and Consts, and the VBA library's functions and constants run nothing."""
    proc_sym = procedure_symbol_for(symbols, proc)
    safe = {child.name.lower() for child in (proc_sym.children if proc_sym is not None else None) or []}
    for param in proc.params:
        if not param.by_val:
            safe.discard(param.name.lower())
    return _code_may_run_with(toks, proc.name.lower(), safe, symbols, leaves_alone)


def _code_may_run_with(
    toks: Sequence[VbaToken],
    own_name: str,
    own_safe: AbstractSet[str],
    symbols: ModuleSymbols,
    leaves_alone: Callable[[str], bool] | None = None,
) -> bool:
    """Whether these tokens may run code that changes what the rules follow: `safe`
    names are this procedure's own, and a call `leaves_alone` clears is one the
    rules have read through."""
    safe = set(own_safe)
    for child in symbols.root.children or []:
        if child.kind in (VbaSymbolKind.MODULE_VARIABLE, VbaSymbolKind.CONSTANT):
            safe.add(child.name.lower())
    member_names = _module_member_names(symbols)
    for i, tok in enumerate(toks):
        if tok.kind is TokenKind.COMMENT:
            continue
        prev = _at(toks, i - 1)
        name = _lower_name(tok)
        if prev is not None and prev.raw_text in (".", "!"):
            # `Debug.Print` and `Err.Number` run nothing of the project's.
            owner = token_text(_at(toks, i - 2))
            if owner in ("debug", "err") and name != "raise":
                continue
            return True
        if not name:
            continue
        if name == "raiseevent":
            return True
        if token_text(prev) == "new":
            # `Dim c As New Collection` makes nothing until c is used.
            if name != "collection" and token_text(_at(toks, i - 2)) != "as":
                return True
            continue
        if tok.kind is TokenKind.KEYWORD and resolve_runtime_function(name) is None:
            continue
        if token_text(prev) == "as" or name in safe or resolve_runtime_constant(name) is not None:
            continue
        # The procedure's own name is its result; with an argument list it is a call.
        after = _at(toks, i + 1)
        if name == own_name and (after is None or after.raw_text != "("):
            continue
        fn = resolve_runtime_function(name) or resolve_runtime_function(f"{name}$")
        if fn is not None and name not in _CODE_RUNNING_FUNCTIONS and name not in member_names:
            continue
        if leaves_alone is not None and leaves_alone(name):
            continue
        return True
    return False


def statement_may_change_module_variable(
    source: str, symbols: ModuleSymbols, proc: ProcedureNode, span: Span, variable: str
) -> bool:
    """Whether a statement may run code that changes the module variable: a call to
    a procedure of the module that writes it, or any code outside the module's own
    (XLIDE issue #618)."""
    return _code_may_run(
        statement_tokens(source, span),
        proc,
        symbols,
        lambda name: _callee_leaves_alone(source, symbols, name, variable),
    )


_LEAVES_ALONE = IdentityLru()


def _callee_leaves_alone(source: str, symbols: ModuleSymbols, name: str, variable: str) -> bool:
    """Whether a call to the module's procedure `name` leaves the module variable
    alone: neither it nor any procedure of the module it calls writes the variable,
    and none runs code outside the module's own (XLIDE issue #618, measured in
    Excel 16.0). Kept per module and asked by name."""
    cache: dict[str, bool] | None = _LEAVES_ALONE.get(symbols)
    if cache is None:
        cache = _LEAVES_ALONE.put({}, symbols)
    assert cache is not None
    known_cache = cache
    visiting: set[str] = set()

    def check(callee: str) -> bool:
        key = f"{callee}|{variable}"
        known = known_cache.get(key)
        if known is not None:
            return known
        if callee in visiting:
            return True  # a cycle adds nothing the other procedures do not
        procedure = [
            child
            for child in symbols.root.children or []
            if is_procedure_kind(child.kind) and child.name.lower() == callee
        ]
        if not procedure:
            return False
        visiting.add(callee)
        alone = True
        for symbol in procedure:
            text = source[symbol.full_span.start : symbol.full_span.end]
            own = {child.name.lower() for child in symbol.children or []}
            # The callee's own lines, its header left out.
            body = [
                tok
                for tok in statement_tokens(source, symbol.full_span)
                if tok.start >= symbol.name_span.end - symbol.full_span.start
            ]
            if variable in written_names_in(text) or _code_may_run_with(body, callee, own, symbols, check):
                alone = False
                break
        visiting.discard(callee)
        known_cache[key] = alone
        return alone

    return check(name)


# What a local of a number type and a String hold before anything assigns them.
DEFAULT_NUMBER: Sequence[VbaToken] = tuple(raw_expression_tokens("0"))
DEFAULT_STRING: Sequence[VbaToken] = tuple(raw_expression_tokens('""'))
_CONSTANT_TRUE: Sequence[VbaToken] = tuple(raw_expression_tokens("-1"))


def _declared_defaults(locals_: Mapping[str, str | None]) -> dict[str, Sequence[VbaToken]]:
    """A typed local holds its declared default until something assigns it, the
    statement that does included: `x = 1 / x` divides by 0 (XLIDE issue #259)."""
    return {
        lower: DEFAULT_NUMBER if kind == "number" else DEFAULT_STRING
        for lower, kind in locals_.items()
        if kind is not None
    }


def _variant_starts(symbols: ModuleSymbols, proc: ProcedureNode) -> dict[str, Sequence[VbaToken]]:
    """Each Variant local, `As Variant` or with no type and no DefType for its
    letter, holds Empty until assigned: `If Not IsEmpty(v) Then` never runs its arm
    (XLIDE issue #691)."""
    out: dict[str, Sequence[VbaToken]] = {}
    proc_sym = procedure_symbol_for(symbols, proc)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if (
            child.kind is VbaSymbolKind.LOCAL_VARIABLE
            and not child.is_array
            and child.visibility is not SymbolVisibility.STATIC
        ):
            def_type = def_type_of(symbols, child.name)
            type_ = normalize_type(child.as_type) or (def_type.lower() if def_type else None) or "variant"
            if type_ == "variant":
                out[child.name.lower()] = VARIANT_EMPTY
    return out


_UNREACHABLE = IdentityLru()


def unreachable_statements_in(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> AbstractSet[int]:
    """The ids of the statements of a procedure that never run, because a guard
    whose value the straight-line walk knows decides against them (XLIDE issue
    #273)."""
    # Seven rules share these facts within one symbol/analysis context.
    table = _per_procedure(_UNREACHABLE, symbols)
    kept = _cached_for(table, proc)
    if isinstance(kept, _SourceKeyed) and kept.source == source and kept.activity is activity:
        return kept.value  # type: ignore[no-any-return]
    # Walked even with no value known: `GoTo Done` leaves whatever the locals hold.
    dead = straight_line_unreachable(
        source,
        proc.body,
        activity,
        _walk_start_with_effects(source, symbols, proc, _literal_value_locals(proc, symbols), activity),
    )
    table[id(proc)] = (proc, _SourceKeyed(source, activity, dead))
    return dead


def dead_branch_spans_in(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> Sequence[Span]:
    """The one-line If branches of a procedure that never run (XLIDE issue #430)."""
    return straight_line_dead_branches(
        source,
        proc.body,
        activity,
        _walk_start_with_effects(source, symbols, proc, _literal_value_locals(proc, symbols), activity),
    )


def defaulted_straight_line(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> Mapping[int, ReachingAssignments]:
    """The straight-line walk from what holds as the procedure starts: each local's
    default, and `Dim c As New Collection` empty (XLIDE issue #614). The same walk
    the value rules share, so it costs nothing more. Keyed by id(statement)."""
    return straight_line_assignments(
        source, proc.body, activity, _walk_start(symbols, proc, _literal_value_locals(proc, symbols))
    )


_WALK_STARTS = IdentityLru()


def _walk_start(
    symbols: ModuleSymbols, proc: ProcedureNode, locals_: Mapping[str, str | None]
) -> ReachingAssignments:
    """What holds as a procedure starts: its Consts' values, and each local's
    declared default. Several rules ask for each procedure, so it is kept, and the
    walk's cache finds it by identity."""
    table = _per_procedure(_WALK_STARTS, symbols)
    start = _cached_for(table, proc)
    if start is None:
        start = {
            **_condition_constants(symbols, proc),
            **_declared_defaults(locals_),
            **_variant_starts(symbols, proc),
            **_object_starts(symbols, proc),
        }
        table[id(proc)] = (proc, start)
    return start  # type: ignore[return-value]


# Statement heads after which a Function may end before its last line.
_RESULT_LEAVING_HEADS = frozenset({"exit", "goto", "gosub", "return", "resume", "on", "stop", "error", "end"})


def _may_leave_early(
    source: str,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    dead: AbstractSet[int],
    dead_spans: Sequence[Span],
) -> bool:
    """Whether a statement that may leave the procedure early still runs: a path
    that leaves returns what it held then, which the end state does not show. One
    the walk found never runs may stay."""

    def never_runs(span: Span) -> bool:
        return any(span.start >= dead_span.start and span.end <= dead_span.end for dead_span in dead_spans)

    def skip(node: BodyNode) -> bool:
        return (activity is not None and activity.is_inactive(node.span)) or id(node) in dead

    for node in iter_body_nodes(body, skip):
        if not is_leaf_statement(node):
            continue
        for span in statement_and_branch_spans(node):
            toks = statement_tokens(source, span)
            first = first_executable_token_index(toks)
            head = token_text(_at(toks, first))
            if not never_runs(span) and (
                head in _RESULT_LEAVING_HEADS
                or (head == "err" and token_text(_at(toks, first + 2)) == "raise")
            ):
                return True
    return False


_CALL_EFFECT_READERS = IdentityLru()


def _call_effects_for(
    source: str, symbols: ModuleSymbols, activity: ConditionalActivityTracker | None
) -> CallEffects:
    """What the module's calls leave in the names they pass ByRef (XLIDE issue
    #449): `ZeroN n`, where ZeroN's body runs through to `n = 0`, leaves the
    caller's n 0. A parameter ByVal, an array, one the callee may leave early or
    assign on one path only, and an argument in parentheses of its own, which
    passes a copy, leave nothing known."""
    kept = _CALL_EFFECT_READERS.get(symbols)
    if isinstance(kept, _SourceKeyed) and kept.source == source and kept.activity is activity:
        return kept.value  # type: ignore[no-any-return]
    procedures: dict[str, ProcedureNode | None] | None = None

    def procedure_named(lower: str) -> ProcedureNode | None:
        nonlocal procedures
        if procedures is None:
            procedures = {}
            for member in parse_module(source).members:
                if (
                    isinstance(member, ProcedureNode)
                    and not (activity is not None and activity.is_inactive(member.span))
                    and member.proc_kind in (ProcKind.SUB, ProcKind.FUNCTION)
                ):
                    key = member.name.lower()
                    procedures[key] = None if key in procedures else member
        return procedures.get(lower)

    left: dict[str, Sequence[VbaToken] | None] = {}

    def left_in(proc: ProcedureNode, index: int) -> Sequence[VbaToken] | None:
        key = f"{proc.name.lower()}|{index}"
        if key not in left:
            param = proc.params[index] if index < len(proc.params) else None
            value: Sequence[VbaToken] | None = None
            if (
                param is not None
                and not param.by_val
                and not param.param_array
                and not param.is_array
                and not any(word.lower() == "static" for word in proc.modifiers)
            ):
                # From a start of its own, which no effects ride: a callee that
                # calls itself is not followed into.
                initial = dict(_walk_start(symbols, proc, _literal_value_locals(proc, symbols)))
                initial.pop(param.name.lower(), None)
                walked = straight_line_exit(source, proc.body, activity, initial)
                held = (
                    walked.exit.get(param.name.lower())
                    if walked.exit is not None
                    and not _may_leave_early(source, proc.body, activity, walked.dead, walked.dead_spans)
                    else None
                )
                value = (
                    held
                    if held
                    and held is not OBJECT_NOTHING
                    and held is not EMPTY_COLLECTION
                    and all(token_name(tok) is None or token_text(tok) in ("true", "false") for tok in held)
                    else None
                )
            left[key] = value
        return left[key]

    def effects(toks: Sequence[VbaToken]) -> Mapping[str, Sequence[VbaToken]]:
        out: dict[str, Sequence[VbaToken]] = {}

        def apply(proc: ProcedureNode | None, args: Sequence[Sequence[VbaToken]]) -> None:
            if proc is None:
                return
            for index, arg in enumerate(args):
                parts = [tok for tok in arg if tok.kind is not TokenKind.COMMENT]
                if any(tok.raw_text == ":=" for tok in parts):
                    return
                name = _lower_name(parts[0]) if len(parts) == 1 else None
                value = left_in(proc, index) if name else None
                if name and value:
                    out[name] = value

        first = first_executable_token_index(toks)
        head = token_text(_at(toks, first))
        callee = _at(toks, first + 1)
        opener = _at(toks, first + 2)
        after_first = _at(toks, first + 1)
        if head == "call" and token_name(callee) and opener is not None and opener.raw_text == "(":
            close = match_paren_from(toks, first + 2)
            apply(
                procedure_named((token_name(callee) or "").lower()),
                split_top_level_token_groups(toks, first + 3, ",", close) if close > first + 3 else [],
            )
        elif token_name(_at(toks, first)) and (
            after_first is None or after_first.raw_text not in ("=", ".", "(")
        ):
            apply(
                procedure_named((token_name(toks[first]) or "").lower()),
                split_top_level_token_groups(toks, first + 1, ",", len(toks)) if len(toks) > first + 1 else [],
            )
        for i in range(first + 1, len(toks) - 1):
            name = _lower_name(toks[i])
            before = _at(toks, i - 1)
            if (
                not name
                or toks[i + 1].raw_text != "("
                or (before is not None and before.raw_text == ".")
                or (head == "call" and i == first + 1)
            ):
                continue
            proc = procedure_named(name)
            if proc is not None and proc.proc_kind is ProcKind.FUNCTION:
                close = match_paren_from(toks, i + 1)
                apply(proc, split_top_level_token_groups(toks, i + 2, ",", close) if close > i + 2 else [])
        return out

    _CALL_EFFECT_READERS.put(_SourceKeyed(source, activity, effects), symbols)
    return effects


def _walk_start_with_effects(
    source: str,
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    locals_: Mapping[str, str | None],
    activity: ConditionalActivityTracker | None,
) -> ReachingAssignments:
    """A procedure's start, with the effects of the module's calls riding its walks."""
    start = _walk_start(symbols, proc, locals_)
    set_call_effects(start, _call_effects_for(source, symbols, activity))
    set_declared_facts(start, _declared_facts_for(source, symbols, proc))
    return start


# A bare upper bound's lower bound: 1 under Option Base 1, else 0.
_OPTION_BASE_ONE = re.compile(r"^[ \t]*Option[ \t]+Base[ \t]+1\b", re.IGNORECASE | re.MULTILINE | re.ASCII)
_FIXED_BOUNDS = re.compile(r"^\s*(?:(-?[0-9]+)\s+To\s+)?(-?[0-9]+)\s*$", re.IGNORECASE)


def _declared_facts_for(source: str, symbols: ModuleSymbols, proc: ProcedureNode) -> DeclaredFacts:
    """What the declarations say of each local, for the guards (XLIDE issue #691):
    its declared type ("long", "long()" for an array), and a fixed one-dimension
    array's bounds."""
    types: dict[str, str] = {}
    bounds: dict[str, tuple[float, float]] = {}
    base: int | None = None
    proc_sym = procedure_symbol_for(symbols, proc)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if (
            child.kind is not VbaSymbolKind.LOCAL_VARIABLE
            or child.visibility is SymbolVisibility.STATIC
            or child.fixed_length is not None
        ):
            continue
        lower = child.name.lower()
        def_type = def_type_of(symbols, child.name)
        type_ = normalize_type(child.as_type) or (def_type.lower() if def_type else None) or "variant"
        types[lower] = f"{type_}()" if child.is_array else type_
        fixed = _FIXED_BOUNDS.match(child.array_bounds or "") if child.is_array else None
        if fixed is not None:
            if fixed.group(1) is not None:
                lower_bound = int(fixed.group(1))
            else:
                if base is None:
                    base = 1 if _OPTION_BASE_ONE.search(source) else 0
                lower_bound = base
            bounds[lower] = (lower_bound, int(fixed.group(2)))
    # The module's Consts and Enum members, which a local or parameter of the same
    # name hides.
    constants: dict[str, float | None] | None = None
    params = {param.name.lower() for param in proc.params}

    def constant(lower: str) -> float | None:
        nonlocal constants
        if lower in types or lower in params:
            return None
        if constants is None:
            constants = collect_module_literal_integer_constants(parse_module(source), None)
        return constants.get(lower)

    return DeclaredFacts(type=types.get, bounds=bounds.get, constant=constant)


_CALL_RESULTS = IdentityLru()


def function_result_for(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    args: Sequence[Sequence[VbaToken] | None],
    object_result: bool = False,
) -> Sequence[VbaToken] | None:
    """What a Function of the module returns for one call's arguments, as the tokens
    of the value it last assigns its name (XLIDE issue #562): `Sign1(-1)` runs `If n
    > 0 Then Sign1 = 1 Else Sign1 = 0` with n = -1, so 0. An omitted Optional takes
    its default. None where the result depends on more than the walk follows: a
    ParamArray, a Static Function, a statement that may leave early and still runs,
    or a value the walk does not know."""
    if (
        proc.proc_kind is not ProcKind.FUNCTION
        or any(word.lower() == "static" for word in proc.modifiers)
        or len(args) > len(proc.params)
    ):
        return None
    # Calls share results only within the same source, symbols and active branch.
    key = (
        f"{'true' if object_result else 'false'}|"
        + ",".join(" ".join(tok.raw_text for tok in arg) if arg is not None else "" for arg in args)
    )
    table = _per_procedure(_CALL_RESULTS, symbols)
    kept = _cached_for(table, proc)
    if not (isinstance(kept, _SourceKeyed) and kept.source == source and kept.activity is activity):
        kept = _SourceKeyed(source, activity, {})
        table[id(proc)] = (proc, kept)
    calls: dict[str, Sequence[VbaToken] | None] = kept.value
    if key in calls:
        return calls[key]
    result = _run_function_for(source, proc, symbols, activity, args, object_result)
    calls[key] = result
    return result


class _NameLookup:
    """An IntegerConstantLookup over a function."""

    __slots__ = ("_get",)

    def __init__(self, get: Callable[[str], float | None]) -> None:
        self._get = get

    def get(self, name: str, /) -> float | None:
        return self._get(name)


_NO_CONSTANTS = _NameLookup(lambda _name: None)


def _run_function_for(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    args: Sequence[Sequence[VbaToken] | None],
    object_result: bool,
) -> Sequence[VbaToken] | None:
    # An object Function's result starts Nothing; the caller knows its type is an
    # object's.
    type_ = normalize_type(proc.return_type if proc.return_type is not None else "Variant")
    result_default: Sequence[VbaToken] | None
    if object_result:
        result_default = OBJECT_NOTHING
    elif type_ == "string":
        result_default = DEFAULT_STRING
    elif type_ is not None and (is_numeric_type(type_) or type_ == "boolean"):
        result_default = DEFAULT_NUMBER
    else:
        result_default = None
    if result_default is None:
        return None
    initial: dict[str, Sequence[VbaToken]] = dict(_walk_start(symbols, proc, _literal_value_locals(proc, symbols)))
    # A module variable nothing writes holds its default: `GetMod = m`.
    for name, held in _module_variable_defaults(source, proc, symbols).items():
        if name not in initial and held.kind in ("number", "string"):
            initial[name] = DEFAULT_NUMBER if held.kind == "number" else DEFAULT_STRING
    for index, param in enumerate(proc.params):
        argument = args[index] if index < len(args) else None
        value: Sequence[VbaToken] | None = (
            argument
            if argument is not None
            else (
                raw_expression_tokens(param.default_raw)
                if param.optional and param.default_raw is not None
                else None
            )
        )
        if param.param_array or value is None:
            return None
        initial[param.name.lower()] = value
    lower = proc.name.lower()
    initial[lower] = result_default
    walked = straight_line_exit(source, proc.body, activity, initial)
    exit_state = walked.exit
    if exit_state is None:
        return None
    value_held = (
        None
        if _may_leave_early(source, proc.body, activity, walked.dead, walked.dead_spans)
        else exit_state.get(lower)
    )
    if (
        value_held is None
        or value_held is OBJECT_NOTHING
        or value_held is EMPTY_COLLECTION
        or not any(token_name(tok) is not None for tok in value_held)
    ):
        return value_held

    # `Twice = n * 2`: a name nothing reassigns holds what the call gave it, so the
    # value folds with it.
    def held_value(name: str) -> float | None:
        held = exit_state.get(name.lower())
        if held is not None and held is initial.get(name.lower()) and all(token_name(tok) is None for tok in held):
            return evaluate_integer_constant_expression(" ".join(tok.raw_text for tok in held), _NO_CONSTANTS)
        return None

    folded = evaluate_integer_constant_expression(
        " ".join(tok.raw_text for tok in value_held), _NameLookup(held_value)
    )
    return None if folded is None else raw_expression_tokens(str(folded))


def _object_starts(symbols: ModuleSymbols, proc: ProcedureNode) -> dict[str, Sequence[VbaToken]]:
    """What an object local holds as the procedure starts (XLIDE issue #483):
    Nothing for one never set, and an empty Collection for `Dim c As New
    Collection`."""
    out: dict[str, Sequence[VbaToken]] = {}
    proc_sym = procedure_symbol_for(symbols, proc)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        type_ = normalize_type(child.as_type)
        if (
            child.kind is not VbaSymbolKind.LOCAL_VARIABLE
            or child.visibility is SymbolVisibility.STATIC
            or child.is_array
            or type_ is None
            or type_ in ("variant", "string")
            or is_known_scalar_type(type_)
        ):
            continue
        if not child.is_auto_instantiated:
            out[child.name.lower()] = OBJECT_NOTHING
        elif type_ in ("collection", "vba.collection"):
            out[child.name.lower()] = EMPTY_COLLECTION
    return out


_DIGITS_RE = re.compile(r"^[0-9]+$")


def _condition_constants(symbols: ModuleSymbols, proc: ProcedureNode) -> dict[str, Sequence[VbaToken]]:
    """The Consts a procedure sees whose value is one literal, as the walk reads a
    value: `Const DEBUGGING = False` holds 0, so `If DEBUGGING Then` never runs its
    arm (XLIDE issue #406). A local or parameter of the same name hides a
    module's."""
    proc_sym = procedure_symbol_for(symbols, proc)
    children = (proc_sym.children if proc_sym is not None else None) or []
    child_ids = {id(child) for child in children}
    out: dict[str, Sequence[VbaToken]] = {}
    for symbol in [*(symbols.root.children or []), *children]:
        toks = (
            [tok for tok in raw_expression_tokens(symbol.default_raw) if tok.kind is not TokenKind.COMMENT]
            if symbol.kind is VbaSymbolKind.CONSTANT and symbol.default_raw is not None
            else []
        )
        word = token_text(toks[0]) if len(toks) == 1 else ""
        value: Sequence[VbaToken] | None
        if word == "true":
            value = _CONSTANT_TRUE
        elif word == "false":
            value = DEFAULT_NUMBER
        elif len(toks) == 1 and (toks[0].kind is TokenKind.STRING_LITERAL or _DIGITS_RE.match(toks[0].raw_text)):
            value = toks
        else:
            value = None
        if value is not None:
            out[symbol.name.lower()] = value
        elif id(symbol) in child_ids:
            out.pop(symbol.name.lower(), None)
    return out


# -- literal values --------------------------------------------------------

_FLOAT_SUFFIX_RE = re.compile(r"[!#@]$")
# Integers past this lose precision as a JavaScript number, so they stay floats.
_EXACT_INTEGER_LIMIT = 2**53


def _plain_literal(
    value: Sequence[VbaToken],
    kind: str,
    typed: bool = False,
    byte: bool = False,
    whole: bool = False,
) -> int | float | str | None:
    """Port of plainLiteralText: the literal a plain `x = literal` assigns, or None
    for any other value. Numbers are returned as numbers (upstream's text read back
    through Number())."""
    literal = _plain_literal_as_written(value, kind, typed, byte)
    # A whole-number type keeps a fraction rounded half to even: `a As Long = 4.4`
    # holds 4, so `r(a)` on `Dim r(3)` raises 9 (XLIDE issue #685).
    if whole and kind == "number" and isinstance(literal, float) and not literal.is_integer():
        return _as_js_number(bankers_round(literal) + 0)
    return literal


def _plain_literal_as_written(
    value: Sequence[VbaToken], kind: str, typed: bool, byte: bool
) -> int | float | str | None:
    toks = unwrap_outer_parens(value)
    if kind == "string":
        if len(toks) == 1 and toks[0].kind is TokenKind.STRING_LITERAL:
            return string_literal_value(toks[0].raw_text)
        return None
    # A typed number or Boolean stores True as -1 and False as 0 (XLIDE issue
    # #491), and a Byte stores True as 255 (#556); a Variant keeps a Boolean, which
    # is no number here.
    word = token_text(toks[0]) if len(toks) == 1 and toks[0].kind is TokenKind.KEYWORD else ""
    if typed and word in ("true", "false"):
        return 0 if word == "false" else 255 if byte else -1
    sign = 1
    rest = toks
    if rest and rest[0].raw_text in ("-", "+"):
        sign = -1 if rest[0].raw_text == "-" else 1
        rest = rest[1:]
    if len(rest) != 1:
        # Whole numbers and operators only: `x = 1 \ 3` stores 0 (XLIDE issue #286).
        if len(toks) > 1 and all(
            tok.kind is TokenKind.INTEGER_LITERAL or tok.kind is TokenKind.OPERATOR or token_text(tok) == "mod"
            for tok in toks
        ):
            return evaluate_integer_constant_expression(" ".join(tok.raw_text for tok in toks), _NO_CONSTANTS)
        return None
    if rest[0].kind is TokenKind.INTEGER_LITERAL:
        parsed = parse_vba_integer_literal(rest[0].raw_text)
        return None if parsed is None else sign * parsed
    if rest[0].kind is TokenKind.FLOAT_LITERAL:
        from ..types.type_inference import float_literal_value

        parsed_float = float_literal_value(rest[0].raw_text)
        if not math.isfinite(parsed_float):
            return None
        return _as_js_number(sign * parsed_float)
    return None


def _as_js_number(value: float) -> int | float:
    """An integral float as an int, the way a JavaScript number prints: no `.0`, and
    equal as a literal to the same value written without a fraction."""
    if value.is_integer() and abs(value) < _EXACT_INTEGER_LIMIT:
        return int(value)
    return value
