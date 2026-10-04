"""Rule family: array declarations, ReDim, subscripts, Erase, and allocation.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/arrays.ts: ReDim
target and bounds validation, unallocated dynamic-array access, Erase targets,
LBound/UBound argument checks, and subscripts proven outside an array's known
bounds. The array-shape helpers here (known_array_shapes_at, redim_shapes_at,
array_value_shape, shape_subscript_violation, ...) are shared with other rules.

Maps upstream keys by a parse node (a statement, a reaching-assignment map, a
token) are keyed here by the object's id(), with the object kept alive beside
the value where the map outlives it: the nodes and tokens are mutable
dataclasses, which are not hashable.
"""

from __future__ import annotations

import dataclasses
import math
import re
from collections.abc import Callable, Iterable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass
from typing import Any

from ...conditional import ConditionalActivityTracker, inactive_node_skip
from ...constants.integer_constant_expression import (
    IntegerConstantLookup,
    bankers_round,
    evaluate_integer_constant_expression,
    parse_vba_integer_literal,
    resolve_raw_integer_constants,
    safe_integer,
)
from ...flow.procedure_labels import jump_target_label_declaration
from ...flow.procedure_unstructured import procedure_has_unstructured_flow
from ...host.host_model import HostObjectModel
from ...identity_cache import IdentityLru
from ...js_compat import JS_WHITESPACE, js_number, js_number_to_string, js_trim, utf16_length
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups, top_level_equals_index
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize
from ...parser.expression_limits import MAX_EXPRESSION_DEPTH
from ...parser.nodes import (
    BodyNode,
    DoBlockNode,
    ForBlockNode,
    LeafStatementNode,
    ModuleNode,
    OptionNode,
    ProcedureNode,
    ProcKind,
    SelectBlockNode,
    Span,
    StatementNode,
    TypeFieldNode,
    TypeNode,
    VariableDeclNode,
    VariableGroupNode,
    WhileBlockNode,
    is_leaf_statement,
    iter_body_nodes_in_context,
)
from ...symbols.name_resolution import BareIdentifierContext, BareIdentifierResolutionScope
from ...symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbol, VbaSymbolKind
from ...types.type_inference import (
    DeclaredValueShape,
    SourceDeclaredShape,
    declaration_shape_environment_for,
    declared_shape_for_source_binding,
    known_local_literal_values_at,
    procedure_symbol_for,
    source_identifier_binding,
    string_constants_in_scope,
    unreachable_statements_in,
    with_known_locals,
)
from ...types.type_names import is_known_scalar_type, normalize_type
from ..block_headers import block_header_statements
from ..call_extraction import split_arg_slots, string_literal_value
from ..callable_signatures import (
    procedure_integer_constant_lookup,
    runtime_callable_source_shadowed,
    scoped_integer_constant_lookup,
    source_name_scope_for,
)
from ..const_expr import collect_module_literal_integer_constants
from ..context import PushFn, statement_tokens
from ..dataflow import (
    BlockEnteringState,
    Lattice,
    StraightLineDataflowHooks,
    walk_branch_merged_body,
    walk_entering_blocks,
    walk_straight_line_body,
)
from ..known_string_calls import ModuleCompare, StringFoldContext, fold_string_expression, module_compare
from ..loop_counters import CountersAt, CounterValue, counter_text, loop_counters_at, numeric_counter_passes
from ..module_state import untouched_module_variables_in
from ..straight_line_values import straight_line_assignments
from ..string_conversion import numeric_string_readings
from ..walker import (
    ProcedureStatementVisitor,
    absolute_span,
    active_module_members,
    bare_assignment_target,
    for_each_statement,
    for_each_statement_with_headers,
    for_each_variable_group,
    is_inactive_node,
    locals_named_whole,
    pluralize_count,
    raw_expression_tokens,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .shared import is_bare_or_vba_qualified_intrinsic_call, source_expression_syntax_problem

# -- node-keyed lookups ----------------------------------------------------
#
# Upstream's node sets use either identity keys or a node-aware collection.


def _node_in(collection: Any, node: object) -> bool:
    """`collection.has(node)`: a set or dict of ids, or a node-keyed collection."""
    if isinstance(collection, (set, frozenset, dict)):
        return id(node) in collection
    return node in collection


class StatementShapes:
    """Upstream's `Map<LeafStatementNode, ReadonlyMap<string, FixedArrayBound>>`.
    Statement nodes are not hashable, so this keys them by identity and keeps
    them alive."""

    __slots__ = ("_by_id",)

    def __init__(self) -> None:
        self._by_id: dict[int, tuple[object, Mapping[str, FixedArrayBound]]] = {}

    def set(self, stmt: object, shapes: Mapping[str, FixedArrayBound]) -> None:
        self._by_id[id(stmt)] = (stmt, shapes)

    def get(self, stmt: object) -> Mapping[str, FixedArrayBound] | None:
        entry = self._by_id.get(id(stmt))
        return entry[1] if entry is not None else None

    def __len__(self) -> int:
        return len(self._by_id)


# -- small token helpers ---------------------------------------------------


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name else None


def _without_comments(toks: Iterable[VbaToken]) -> list[VbaToken]:
    return [tok for tok in toks if tok.kind is not TokenKind.COMMENT]


def _fmt(value: float) -> str:
    """A number as JavaScript's template literal prints it."""
    return js_number_to_string(value)


# Integers below this are exact as a JavaScript number, so they print as Python
# prints them.
_EXACT_INTEGER_LIMIT = 2**53


def _as_js_number(value: float) -> int | float:
    """An integral double as an int, the way a JavaScript number prints: no `.0`."""
    if math.isfinite(value) and value.is_integer() and abs(value) < _EXACT_INTEGER_LIMIT:
        return int(value)
    return value


# -- checkArrayBoundIntrinsicArguments -------------------------------------


def check_array_bound_intrinsic_arguments(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """Per-statement rule: rides the shared procedure-statement walk (audit #0)."""

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        shapes = declaration_shape_environment_for(symbols, member)
        proc_sym = procedure_symbol_for(symbols, member)

        def resolve_shape(name: str) -> SourceDeclaredShape:
            return declared_shape_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.EXPRESSION
            )

        def visitor(stmt: LeafStatementNode) -> None:
            for name, as_type, span in _scalar_locals_indexed(source, stmt.span, member, shapes):
                push(
                    "scalarIndexed",
                    f"'{name}' is declared As {as_type}, which is no array to index. "
                    "This is a VBE compile error: Expected array.",
                    span,
                )
            for span, message in _array_bound_syntax_problems(source, stmt.span):
                push("malformedStatement", message, span)
            for function_name, name, span, as_type in _array_bound_scalar_arguments(
                source, stmt.span, shapes, resolve_shape
            ):
                push(
                    "arrayBoundRequiresArray",
                    f"{function_name} requires an array argument, but '{name}' is declared As {as_type}.",
                    span,
                )

        return visitor

    return factory


_SCALAR_INDEXED_SKIPPED_LEADS = frozenset(
    {"dim", "redim", "static", "const", "private", "public", "global", "erase"}
)


def _scalar_locals_indexed(
    source: str,
    span: Span,
    proc: ProcedureNode,
    shapes: Mapping[str, DeclaredValueShape],
) -> list[tuple[str, str, Span]]:
    """`Dim f As Long: Main = f(1)`: a number, string or date variable given
    a subscript is the VBE compile error Expected array (issue #417, Excel
    16.0). The procedure's own name is a call, and declarations are left to
    the declaration rules."""
    toks = statement_tokens_after_leading_label(source, span)
    if token_text(_at(toks, 0)) in _SCALAR_INDEXED_SKIPPED_LEADS:
        return []
    out: list[tuple[str, str, Span]] = []
    for i in range(len(toks) - 1):
        name = token_name(toks[i])
        if (
            not name
            or toks[i + 1].raw_text != "("
            or _raw_at(toks, i - 1) in (".", "!")
            or name.lower() == proc.name.lower()
        ):
            continue
        shape = shapes.get(name.lower())
        normalized = (
            normalize_type(shape.as_type) if shape is not None and not shape.is_array and shape.as_type else None
        )
        if not normalized or not is_known_scalar_type(normalized):
            continue
        assert shape is not None and shape.as_type is not None
        out.append((name, shape.as_type, Span(span.start + toks[i].start, span.start + toks[i].end)))
    return out


def _array_bound_syntax_problems(source: str, span: Span) -> list[tuple[Span, str]]:
    """`UBound(5)`, `LBound("abc")`, `UBound(-5)` and `UBound(5 + 1)`: what
    UBound and LBound read is an array variable, a member or a call, and
    anything else is a Syntax error (measured in Excel 16.0). `VBA.UBound` is
    left alone: UBound is no member of VBA, which the VBE reports first."""
    toks = statement_tokens(source, span)
    out: list[tuple[Span, str]] = []
    for i in range(len(toks) - 2):
        word = token_text(toks[i])
        if word not in ("ubound", "lbound") or toks[i + 1].raw_text != "(" or _raw_at(toks, i - 1) == ".":
            continue
        close = match_paren_from(toks, i + 1)
        split = None if close < 0 else split_arg_slots(toks[i + 2 : close], span.start)
        first = (
            [tok for tok in split.slots[0] if tok.kind not in (TokenKind.COMMENT, TokenKind.NEWLINE)]
            if split is not None and split.slots
            else []
        )
        # `UBound((a))` is reported as an array in parentheses already.
        if not first or first[0].raw_text == "(":
            continue
        at = Span(span.start + first[0].start, span.start + first[-1].end)
        what = source_expression_syntax_problem(source[at.start : at.end])
        if what:
            out.append(
                (
                    at,
                    f"{toks[i].raw_text} takes an array variable, a member or a call, and {what} "
                    "is none of them. This is a VBE compile error: Syntax error.",
                )
            )
    return out


def _array_bound_scalar_arguments(
    source: str,
    span: Span,
    shapes: Mapping[str, DeclaredValueShape],
    resolve_shape: Callable[[str], SourceDeclaredShape] | None = None,
) -> list[tuple[str, str, Span, str]]:
    toks = statement_tokens(source, span)
    hits: list[tuple[str, str, Span, str]] = []
    for i in range(len(toks) - 2):
        function_name = token_name(toks[i])
        lower = function_name.lower() if function_name else None
        if lower not in ("lbound", "ubound"):
            continue
        if _raw_at(toks, i + 1) != "(" or not is_bare_or_vba_qualified_intrinsic_call(toks, i):
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0:
            continue
        inner = toks[i + 2 : close]
        if not inner:
            continue
        split = split_arg_slots(inner, span.start)
        first_slot = split.slots[0] if split.slots else []
        if len(first_slot) != 1:
            continue
        arg_name = token_name(first_slot[0])
        if not arg_name:
            continue
        resolved = resolve_shape(arg_name) if resolve_shape is not None else None
        shape = resolved.shape if resolved is not None and resolved.resolved else shapes.get(arg_name.lower())
        if shape is None or shape.is_array or not shape.as_type:
            continue
        # Only a Variant can hold an array: a Collection, Object or Type is
        # refused as well (issue #417, Excel 16.0).
        normalized = normalize_type(shape.as_type)
        if not normalized or normalized in ("variant", "any"):
            continue
        assert function_name is not None
        hit_span = (
            split.spans[0]
            if split.spans
            else Span(span.start + first_slot[0].start, span.start + first_slot[0].end)
        )
        hits.append((function_name, arg_name, hit_span, shape.as_type))
    return hits


# -- ReDim target parsing ----------------------------------------------------


@dataclass(frozen=True, slots=True)
class _RedimBlockedDeclaration:
    name: str
    span: Span
    kind: str  # "fixedArray" | "scalar"


@dataclass(frozen=True, slots=True)
class _RedimDimension:
    span: Span
    key: str | None = None
    lower_key: str | None = None
    # Whether the bound writes a lower, `1 To n`; without one it takes the Option Base.
    lower_written: bool | None = None
    lower_value: int | None = None
    upper_value: int | None = None


@dataclass(frozen=True, slots=True)
class _RedimElementType:
    """The element type an `As` after the bounds names, with its tokens' span."""

    name: str
    span: Span


@dataclass(frozen=True, slots=True)
class _RedimTarget:
    name: str
    span: Span
    preserve: bool
    dimensions: list[_RedimDimension]
    as_type: _RedimElementType | None = None


def check_invalid_redim_targets(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """Rule: ReDim can allocate dynamic arrays, but it cannot resize a variable
    that was already declared as a scalar or as a fixed-size array."""
    module_declarations = _redim_blocked_declarations_for_module(mod, activity)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        local_declarations = _redim_blocked_declarations_for_body(member.body, activity)
        local_names = _declaration_names_for_body(member.body, activity)
        proc_sym = procedure_symbol_for(symbols, member)

        def visitor(stmt: LeafStatementNode) -> None:
            for target in _redim_statement_targets(source, stmt.span):
                lower = target.name.lower()
                # A Const or Enum member is no array (issue #255): "Expected array".
                binding = source_identifier_binding(
                    symbols, proc_sym, project_visible_symbols, target.name, BareIdentifierContext.EXPRESSION
                )
                if (
                    binding.scope is not BareIdentifierResolutionScope.AMBIGUOUS
                    and len(binding.definitions) > 0
                    and all(
                        definition.kind in (VbaSymbolKind.CONSTANT, VbaSymbolKind.ENUM_MEMBER)
                        for definition in binding.definitions
                    )
                ):
                    push(
                        "scalarRedim",
                        f"'{target.name}' is a constant, which ReDim cannot resize. "
                        "This is a VBE compile error: Expected array.",
                        target.span,
                    )
                    continue
                resolved = declared_shape_for_source_binding(
                    symbols, proc_sym, project_visible_symbols, target.name,
                    BareIdentifierContext.ASSIGNMENT_TARGET,
                )
                declaration: _RedimBlockedDeclaration | None
                if resolved.resolved:
                    declaration = _redim_blocked_declaration_for_shape(target.name, target.span, resolved.shape)
                else:
                    declaration = local_declarations.get(lower)
                    if declaration is None and lower not in local_names:
                        declaration = module_declarations.get(lower)
                if declaration is None:
                    continue
                if declaration.kind == "scalar":
                    push(
                        "scalarRedim",
                        f"Scalar variable '{target.name}' cannot be resized with ReDim; "
                        "declare it as a dynamic array first.",
                        target.span,
                    )
                    continue
                push(
                    "fixedArrayRedim",
                    f"Fixed-size array '{target.name}' cannot be resized with ReDim.",
                    target.span,
                )

        return visitor

    return factory


def _redim_blocked_declaration_for_shape(
    name: str, span: Span, shape: DeclaredValueShape | None
) -> _RedimBlockedDeclaration | None:
    if shape is None:
        return None
    if not shape.is_array:
        if _is_variant_like_redim_target_type(shape.as_type):
            return None
        return _RedimBlockedDeclaration(name=name, span=span, kind="scalar")
    if shape.is_fixed_array:
        return _RedimBlockedDeclaration(name=name, span=span, kind="fixedArray")
    return None


def _is_variant_like_redim_target_type(as_type: str | None) -> bool:
    return not as_type or normalize_type(as_type) == "variant"


def _redim_blocked_declarations_for_module(
    mod: ModuleNode, activity: ConditionalActivityTracker | None
) -> dict[str, _RedimBlockedDeclaration]:
    out: dict[str, _RedimBlockedDeclaration] = {}
    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            _add_redim_blocked_declarations(member, out)
    return out


def _redim_blocked_declarations_for_body(
    body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> dict[str, _RedimBlockedDeclaration]:
    out: dict[str, _RedimBlockedDeclaration] = {}
    for_each_variable_group(body, lambda group: _add_redim_blocked_declarations(group, out), activity)
    return out


def _declaration_names_for_body(
    body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> set[str]:
    names: set[str] = set()
    for_each_variable_group(
        body, lambda group: names.update(decl.name.lower() for decl in group.declarations), activity
    )
    return names


def _add_redim_blocked_declarations(
    group: VariableGroupNode, out: dict[str, _RedimBlockedDeclaration]
) -> None:
    for decl in group.declarations:
        kind = _redim_blocked_declaration_kind(decl)
        if kind is None:
            continue
        lower = decl.name.lower()
        if lower not in out:
            out[lower] = _RedimBlockedDeclaration(name=decl.name, span=decl.span, kind=kind)


def _redim_blocked_declaration_kind(decl: VariableDeclNode) -> str | None:
    if not decl.is_array:
        if _is_variant_like_redim_target_type(decl.as_type):
            return None
        return "scalar"
    return "fixedArray" if decl.array_bounds else None


def _redim_statement_targets(source: str, span: Span) -> list[_RedimTarget]:
    return _redim_targets_from_tokens(span, statement_tokens_after_leading_label(source, span))


def _redim_targets_from_tokens(span: Span, toks: Sequence[VbaToken]) -> list[_RedimTarget]:
    if token_text(_at(toks, 0)) != "redim":
        return []
    preserve = token_text(_at(toks, 1)) == "preserve"
    start = 2 if preserve else 1
    out: list[_RedimTarget] = []
    for group in split_top_level_token_groups(toks, start, ","):
        target = _redim_target_from_group(span, group, preserve)
        if target is not None:
            out.append(target)
    return out


def _single_line_if_redim_targets(source: str, span: Span) -> list[_RedimTarget]:
    """ReDim allocations embedded in a single-line ``If cond Then ReDim a(...)``
    (and its ``Else ReDim b(...)`` arm). The parser keeps single-line If
    statements as one leaf, so _redim_statement_targets - which requires
    ``redim`` as the first token - never sees them, and the generic index-access
    scan would otherwise flag the ReDim's own target as an unallocated access."""
    toks = statement_tokens_after_leading_label(source, span)
    if token_text(_at(toks, 0)) != "if":
        return []
    out: list[_RedimTarget] = []
    for i in range(1, len(toks) - 1):
        word = token_text(toks[i])
        if word not in ("then", "else") or token_text(toks[i + 1]) != "redim":
            continue
        end = len(toks)
        for j in range(i + 2, len(toks)):
            if token_text(toks[j]) == "else":
                end = j
                break
        out.extend(_redim_targets_from_tokens(span, toks[i + 1 : end]))
    return out


def _redim_target_from_group(
    base: Span, group: Sequence[VbaToken], preserve: bool
) -> _RedimTarget | None:
    content = _without_comments(group)
    name_tok = _at(content, 0)
    name = token_name(name_tok)
    if not name or name_tok is None:
        return None
    # A qualified ReDim target (`ReDim x.arr(...)` or `ReDim x!arr(...)`) resizes
    # a member array, not the base variable. The scalar/fixed-array shape checks
    # only apply to a simple local/module variable, and the member's declared
    # shape is not resolvable here, so skip qualified targets rather than mistake
    # the container for the array being resized.
    if _raw_at(content, 1) in (".", "!"):
        return None
    dimensions: list[_RedimDimension] = []
    as_type: _RedimElementType | None = None
    if _raw_at(content, 1) == "(":
        close = match_paren_from(content, 1)
        type_tokens = (
            content[close + 2 :] if close > 1 and token_text(_at(content, close + 1)) == "as" else []
        )
        if type_tokens and all(token_name(tok) or tok.raw_text == "." for tok in type_tokens):
            as_type = _RedimElementType(
                name="".join(tok.raw_text for tok in type_tokens),
                span=_token_group_span(base, type_tokens),
            )
        if close > 1:
            for part in split_top_level_token_groups(content, 2, ",", close):
                dim_tokens = _without_comments(part)
                if not dim_tokens:
                    continue
                key, lower_key, lower_value, upper_value = _comparable_array_bound_key(dim_tokens)
                dimensions.append(
                    _RedimDimension(
                        span=_token_group_span(base, dim_tokens),
                        key=key,
                        lower_key=lower_key,
                        lower_written=any(token_text(tok) == "to" for tok in dim_tokens),
                        lower_value=lower_value,
                        upper_value=upper_value,
                    )
                )
    return _RedimTarget(
        name=name,
        span=absolute_span(base, name_tok),
        preserve=preserve,
        dimensions=dimensions,
        as_type=as_type,
    )


# -- checkRedimTypeChange ----------------------------------------------------

# The type a suffix character gives a name: `x$` is a String.
_SUFFIX_TYPES: Mapping[str, str] = {
    "%": "integer",
    "&": "long",
    "^": "longlong",
    "@": "currency",
    "!": "single",
    "#": "double",
    "$": "string",
}

_JS_WHITESPACE_RUN_RE = re.compile(f"[{JS_WHITESPACE}]+")
_STAR_RE = re.compile(r"\*")


def _dynamic_array_element_type(decl: VariableDeclNode) -> str | None:
    """The element type a dynamic array was declared with, lowercased: `Dim x()
    As String` is string and `Dim v()` Variant. None for anything that is not a
    dynamic array, and for a fixed-length string, whose ReDim is not judged."""
    if not decl.is_array or js_trim(decl.array_bounds or "") != "" or decl.fixed_length is not None:
        return None
    if decl.as_type:
        return _JS_WHITESPACE_RUN_RE.sub("", decl.as_type).lower()
    return _SUFFIX_TYPES.get(decl.type_suffix) if decl.type_suffix else "variant"


def check_redim_type_change(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """Rule: a ReDim may not give a dynamic array another element type. `Dim x()
    As String` then `ReDim x(1) As Long` is "Can't change data types of array
    elements" in the VBE, and so is `Dim v()` then `ReDim v(1) As Long`, with
    Preserve or without; a Variant that is not an array takes any ReDim (issue
    #212, measured in Excel 16.0)."""
    module_arrays: dict[str, str | None] = {}
    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode) and not member.is_const:
            for decl in member.declarations:
                key = decl.name.lower()
                module_arrays[key] = None if key in module_arrays else _dynamic_array_element_type(decl)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        local_arrays: dict[str, str | None] = {}

        def visit_group(group: VariableGroupNode) -> None:
            for decl in group.declarations:
                key = decl.name.lower()
                local_arrays[key] = None if key in local_arrays else _dynamic_array_element_type(decl)

        for_each_variable_group(member.body, visit_group, activity)
        params = {param.name.lower() for param in member.params}

        def visitor(stmt: LeafStatementNode) -> None:
            for target in _redim_statement_targets(source, stmt.span):
                if target.as_type is None or _STAR_RE.search(target.as_type.name):
                    continue
                key = target.name.lower()
                if key in params:
                    continue
                declared = local_arrays.get(key) if key in local_arrays else module_arrays.get(key)
                if declared is None or declared == target.as_type.name.lower():
                    continue
                push(
                    "redimTypeChange",
                    f"'{target.name}' was declared with elements of another type, and a ReDim cannot "
                    f"change it to {target.as_type.name}. This is a VBE compile error: Can't change "
                    "data types of array elements.",
                    target.as_type.span,
                )

        return visitor

    return factory


# -- comparable literal-bound folding ---------------------------------------


def _comparable_array_bound_key(
    toks: Sequence[VbaToken],
) -> tuple[str | None, str | None, int | None, int | None]:
    """Returns (key, lower_key, lower_value, upper_value) for one dimension."""
    to_index = next((i for i, tok in enumerate(toks) if token_text(tok) == "to"), -1)
    if to_index > 0:
        lower_key = _comparable_array_bound_expression_key(toks[:to_index])
        lower_value = _comparable_array_bound_expression_value(toks[:to_index])
        upper_key = _comparable_array_bound_expression_key(toks[to_index + 1 :])
        upper_value = _comparable_array_bound_expression_value(toks[to_index + 1 :])
        key = f"{lower_key}to{upper_key}" if lower_key and upper_key else None
        return (key, lower_key, lower_value, upper_value)
    upper_key = _comparable_array_bound_expression_key(toks)
    return (upper_key, None, None, _comparable_array_bound_expression_value(toks))


def _comparable_array_bound_expression_key(toks: Sequence[VbaToken]) -> str | None:
    parts: list[str] = []
    for tok in toks:
        word = token_text(tok)
        if tok.kind is TokenKind.INTEGER_LITERAL or tok.raw_text in ("+", "-") or word == "to":
            parts.append(word if word else tok.raw_text.lower())
            continue
        return None
    return "".join(parts) if parts else None


def _comparable_array_bound_expression_value(toks: Sequence[VbaToken]) -> int | None:
    """Folds a `lower To upper` array bound that is built only from signed
    integer literals (`-3`, `1 + 2`). Intentionally a literal-only subset of
    evaluate_integer_constant_expression."""
    value = 0
    sign = 1
    expecting_value = True
    saw_value = False
    for tok in toks:
        if expecting_value:
            if tok.raw_text in ("+", "-"):
                sign *= -1 if tok.raw_text == "-" else 1
                continue
            if tok.kind is not TokenKind.INTEGER_LITERAL:
                return None
            parsed = parse_vba_integer_literal(tok.raw_text)
            if parsed is None:
                return None
            next_value = value + sign * parsed
            if safe_integer(next_value) is None:
                return None
            value = next_value
            sign = 1
            expecting_value = False
            saw_value = True
            continue
        if tok.raw_text in ("+", "-"):
            sign = -1 if tok.raw_text == "-" else 1
            expecting_value = True
            continue
        return None
    return value if saw_value and not expecting_value else None


# -- checkRedimImpossibleBounds ----------------------------------------------

# VBA allows at most 60 array dimensions.
_MAX_ARRAY_DIMENSIONS = 60


def check_redim_impossible_bounds(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_integer_constants: Mapping[str, str | None] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    host_model: HostObjectModel | None = None,
) -> ProcedureStatementVisitor:
    """Rule: ReDim lower bounds must not be greater than their upper bounds, a
    bound being a literal or a constant expression (issue #209): `ReDim a(LO To
    HI)` with LO = 5 and HI = 1 raises error 9. Also: a ReDim of more than 60
    dimensions, which the VBE refuses as a syntax error (issue #209)."""
    module_declarations = _redim_blocked_declarations_for_module(mod, activity)
    option_base = module_option_base(mod, activity)
    module_constants = _module_integer_constants(mod, project_integer_constants, activity)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        local_declarations = _redim_blocked_declarations_for_body(member.body, activity)
        local_names = _declaration_names_for_body(member.body, activity)
        constants: list[IntegerConstantLookup] = []

        def constants_lookup() -> IntegerConstantLookup:
            if not constants:
                constants.append(
                    procedure_integer_constant_lookup(
                        member, module_constants, symbols, project_visible_symbols, activity, host_model
                    )
                )
            return constants[0]

        values_at = known_local_literal_values_at(source, member, symbols, activity)

        def visitor(stmt: LeafStatementNode) -> None:
            # `zz = -1` then `ReDim a(zz)` (issue #238).
            def lookup() -> IntegerConstantLookup:
                return with_known_locals(constants_lookup(), values_at(stmt))

            for target in _redim_statement_targets(source, stmt.span):
                if len(target.dimensions) > _MAX_ARRAY_DIMENSIONS:
                    push(
                        "tooManyArrayDimensions",
                        f"ReDim of '{target.name}' has {len(target.dimensions)} dimensions; VBA allows "
                        f"at most {_MAX_ARRAY_DIMENSIONS}. This is a VBE compile error: Syntax error.",
                        target.span,
                    )
                lower_name = target.name.lower()
                blocked = local_declarations.get(lower_name)
                if blocked is None and lower_name not in local_names:
                    blocked = module_declarations.get(lower_name)
                if blocked is not None:
                    continue
                for index, dimension in enumerate(target.dimensions):
                    # `ReDim a(-1)`: the lower bound is Option Base, 0 by default,
                    # and an upper bound below it is the same impossibility as
                    # `ReDim a(5 To 1)` (issue #120, measured in Excel 16.0).
                    bounds = _impossible_bounds(source, dimension.span, lookup(), option_base)
                    if bounds is None:
                        continue
                    lower_text, upper = bounds
                    push(
                        "redimImpossibleBounds",
                        f"ReDim lower bound {lower_text} is greater than upper bound {_fmt(upper)} for "
                        f"dimension {index + 1} of '{target.name}'; this will raise Run-time error '9': "
                        "Subscript out of range.",
                        dimension.span,
                    )

        return visitor

    return factory


# -- checkArrayDeclarationBounds ---------------------------------------------


def check_array_declaration_bounds(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_integer_constants: Mapping[str, str | None] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    host_model: HostObjectModel | None = None,
) -> None:
    """Rule family on `Dim`/`Static`/`Private`/`Public` array declarations and
    on Type members:
     - `array-declaration-impossible-bounds`: a dimension whose lower bound is
       above its upper one, "Range has no values" in the VBE. A bound is a
       literal or a constant expression the module, the procedure, another
       module's Public Const, an Enum or the host defines (issue #209). A bound
       that names a variable stays quiet.
     - `too-many-array-dimensions`: more than 60 dimensions, in a Type member too.
    ReDim is covered separately by check_redim_impossible_bounds."""
    option_base = module_option_base(mod, activity)
    module_constants = _module_integer_constants(mod, project_integer_constants, activity)
    module_lookup: list[IntegerConstantLookup] = []

    def module_scope() -> IntegerConstantLookup:
        if not module_lookup:
            module_lookup.append(
                scoped_integer_constant_lookup(
                    module_constants, symbols, None, project_visible_symbols, host_model
                )
            )
        return module_lookup[0]

    def inspect_group(group: VariableGroupNode, lookup: Callable[[], IntegerConstantLookup]) -> None:
        for decl in group.declarations:
            if not decl.is_array or decl.array_bounds is None or is_inactive_node(activity, decl):
                continue
            _inspect_array_declaration(source, decl, lookup, option_base, push)

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            inspect_group(member, module_scope)
        elif isinstance(member, TypeNode):
            for type_field in member.fields:
                if type_field.is_array and not is_inactive_node(activity, type_field):
                    _inspect_array_declaration(source, type_field, module_scope, option_base, push)
        elif isinstance(member, ProcedureNode):
            _inspect_procedure_declarations(
                member, module_constants, symbols, project_visible_symbols, activity, host_model, inspect_group
            )


def _inspect_procedure_declarations(
    member: ProcedureNode,
    module_constants: Mapping[str, float | None],
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    host_model: HostObjectModel | None,
    inspect_group: Callable[[VariableGroupNode, Callable[[], IntegerConstantLookup]], None],
) -> None:
    constants: list[IntegerConstantLookup] = []

    def procedure_scope() -> IntegerConstantLookup:
        if not constants:
            constants.append(
                procedure_integer_constant_lookup(
                    member, module_constants, symbols, project_visible_symbols, activity, host_model
                )
            )
        return constants[0]

    for_each_variable_group(member.body, lambda group: inspect_group(group, procedure_scope), activity)


def _module_integer_constants(
    mod: ModuleNode,
    project_integer_constants: Mapping[str, str | None] | None,
    activity: ConditionalActivityTracker | None,
) -> dict[str, float | None]:
    """The module's integer constants and Enum members, over the project's Public ones."""
    project_constants = resolve_raw_integer_constants(project_integer_constants or {}, {})
    return collect_module_literal_integer_constants(mod, activity, project_constants)


_FLOAT_TYPE_SUFFIX_RE = re.compile(r"[!#@]\Z")
_DOUBLE_EXPONENT_RE = re.compile(r"[dD]")


def _impossible_bounds(
    source: str, span: Span, lookup: IntegerConstantLookup, option_base: int
) -> tuple[str, int | float] | None:
    """A dimension's bounds when the lower is above the upper, folded through
    constants. With no `To`, the lower bound is Option Base. A bound that is a
    lone float literal is rounded half to even, as VBA rounds it: `1.5 To 1` is
    refused and `2.5 To 2` is not (issue #209). None when either bound is
    unknown or the bounds are fine. Returns (lower text, upper)."""
    text = source[span.start : span.end]
    toks = [tok for tok in tokenize(text) if tok.kind not in (TokenKind.COMMENT, TokenKind.NEWLINE)]
    depth = 0
    to = -1
    for i, tok in enumerate(toks):
        if to >= 0:
            break
        raw = tok.raw_text
        depth += 1 if raw == "(" else -1 if raw == ")" else 0
        if depth == 0 and token_text(tok) == "to":
            to = i

    def side(part: Sequence[VbaToken]) -> int | float | None:
        if not part:
            return None
        # A decimal, signed too: `ReDim d(-0.6)` rounds to -1 (issue #286).
        negative = len(part) == 2 and part[0].raw_text == "-"
        decimal = part[1 if negative else 0]
        if len(part) == (2 if negative else 1) and decimal.kind is TokenKind.FLOAT_LITERAL:
            value = js_number(_DOUBLE_EXPONENT_RE.sub("e", _FLOAT_TYPE_SUFFIX_RE.sub("", decimal.raw_text), count=1))
            if not math.isfinite(value):
                return None
            rounded: int | float = bankers_round(-value if negative else value)
            return rounded + 0
        return evaluate_integer_constant_expression(text[part[0].start : part[-1].end], lookup)

    upper = side(toks if to < 0 else toks[to + 1 :])
    lower = option_base if to < 0 else side(toks[:to])
    if upper is None or lower is None or lower <= upper:
        return None
    lower_text = f"{_fmt(lower)} (Option Base {_fmt(lower)})" if to < 0 else _fmt(lower)
    return (lower_text, upper)


def _inspect_array_declaration(
    source: str,
    decl: VariableDeclNode | TypeFieldNode,
    lookup: Callable[[], IntegerConstantLookup],
    option_base: int,
    push: PushFn,
) -> None:
    toks = statement_tokens(source, decl.span)
    open_index = next((i for i, tok in enumerate(toks) if tok.raw_text == "("), -1)
    if open_index < 0:
        return
    close = match_paren_from(toks, open_index)
    if close < 0:
        return
    dims = [
        dim_tokens
        for dim_tokens in (
            _without_comments(part) for part in split_top_level_token_groups(toks, open_index + 1, ",", close)
        )
        if dim_tokens
    ]
    if len(dims) > _MAX_ARRAY_DIMENSIONS:
        push(
            "tooManyArrayDimensions",
            f"Array '{decl.name}' has {len(dims)} dimensions; VBA allows at most {_MAX_ARRAY_DIMENSIONS}.",
            decl.name_span if decl.name_span is not None else decl.span,
        )
    for index, dim_tokens in enumerate(dims):
        span = _token_group_span(decl.span, dim_tokens)
        bounds = _impossible_bounds(source, span, lookup(), option_base)
        if bounds is None:
            continue
        lower_text, upper = bounds
        push(
            "arrayDeclarationImpossibleBounds",
            f"Array '{decl.name}' lower bound {lower_text} is greater than upper bound {_fmt(upper)} for "
            f"dimension {index + 1}; this is not a valid array bound.",
            span,
        )


# -- checkRedimPreserveDimensions --------------------------------------------


def check_redim_preserve_dimensions(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """Rule: ReDim Preserve may only resize the last dimension of an already
    allocated dynamic array. This tracks simple, active ReDim shapes in a
    conservative per-body flow so nested branch updates do not leak outward."""
    option_base: int | None = None
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        if option_base is None:
            option_base = module_option_base(mod, activity)
        _check_redim_preserve_dimensions_in_body(source, member.body, {}, activity, push, option_base)


def _check_redim_preserve_dimensions_in_body(
    source: str,
    body: Sequence[BodyNode],
    initial_shapes: Mapping[str, _RedimTarget],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    base: int,
) -> None:
    # Copy-down, no-leak-up: a block's body starts from a copy of the shapes
    # learned before the block, and what it learns stays inside it.
    for node, shapes in iter_body_nodes_in_context(
        body, dict(initial_shapes), lambda _block, outer: dict(outer), inactive_node_skip(activity)
    ):
        if not isinstance(node, StatementNode):
            continue
        # After Erase a dynamic array has no bounds, and ReDim Preserve sets
        # them afresh, the first dimension included (issue #420).
        toks = statement_tokens_after_leading_label(source, node.span)
        if token_text(_at(toks, 0)) == "erase":
            for tok in toks[1:]:
                erased = token_name(tok)
                if erased:
                    shapes.pop(erased.lower(), None)
            continue
        for target in _redim_statement_targets(source, node.span):
            if target.preserve:
                previous = shapes.get(target.name.lower())
                reason = (
                    _redim_preserve_dimension_mismatch(previous, target, base) if previous is not None else None
                )
                if reason:
                    push(
                        "redimPreserveDimensionChange",
                        f"ReDim Preserve can only resize the last dimension of '{target.name}'. {reason}",
                        target.span,
                    )
            if target.dimensions:
                shapes[target.name.lower()] = target


def _redim_preserve_dimension_mismatch(
    previous: _RedimTarget, current: _RedimTarget, base: int
) -> str | None:
    # A bound written with no lower takes the Option Base: after
    # `ReDim a(1 To 3)`, `ReDim Preserve a(UBound(a) + 1)` moves the lower
    # bound to 0, which raises 9 (issue #342, measured in Excel 16.0).
    def moves_lower(before: _RedimDimension, after: _RedimDimension) -> bool:
        return (
            before.lower_written is True
            and after.lower_written is False
            and before.lower_value is not None
            and before.lower_value != base
        ) or (
            before.lower_written is False
            and after.lower_written is True
            and after.lower_value is not None
            and after.lower_value != base
        )

    prev_len = len(previous.dimensions)
    cur_len = len(current.dimensions)
    if prev_len == cur_len:
        for i in range(cur_len):
            if moves_lower(previous.dimensions[i], current.dimensions[i]):
                lower = previous.dimensions[i].lower_value
                now = current.dimensions[i].lower_value
                return (
                    f"The lower bound of dimension {i + 1} changes under Preserve, from "
                    f"{base if lower is None else lower} to {base if now is None else now}: "
                    "a bound written without one takes the Option Base."
                )
    if prev_len > 0 and cur_len > 0 and prev_len != cur_len:
        return (
            f"Previous ReDim has {pluralize_count(prev_len, 'dimension')}, "
            f"but this ReDim Preserve has {cur_len}."
        )
    comparable_count = min(prev_len, cur_len) - 1
    for i in range(comparable_count):
        before = previous.dimensions[i].key
        after = current.dimensions[i].key
        if before and after and before != after:
            return f"Dimension {i + 1} changes before the final dimension."
    final_index = min(prev_len, cur_len) - 1
    if final_index >= 0:
        before_lower = previous.dimensions[final_index].lower_key
        after_lower = current.dimensions[final_index].lower_key
        if before_lower and after_lower and before_lower != after_lower:
            return f"The lower bound of dimension {final_index + 1} changes under Preserve."
    return None


# -- checkUnallocatedDynamicArrayAccess --------------------------------------


@dataclass(frozen=True, slots=True)
class _DynamicArrayDeclaration:
    name: str
    span: Span
    # A Variant local that takes a copy of a dynamic array (issue #342).
    variant: bool | None = None


def check_unallocated_dynamic_array_access(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Rule: a local dynamic array declared as `Dim values() As T` has no storage
    until ReDim allocates it. This tracks only straight-line local state; nested
    runtime blocks and helper calls make the state unknown instead of guessed."""
    unset_functions = _array_functions_never_set(source, mod, activity)
    erased_by_call = _arrays_erased_by_calls(source, mod, activity)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        if unset_functions:
            _check_unset_array_results(source, member, symbols, unset_functions, activity, push)
        _check_unallocated_procedure(source, member, symbols, activity, push, erased_by_call)


def _check_unallocated_procedure(
    source: str,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    erased_by_call: Callable[[Sequence[VbaToken]], AbstractSet[str]],
) -> None:
    arrays = _local_dynamic_array_declarations_for_body(member.body, activity)
    state: dict[str, str] = dict.fromkeys(arrays, "unallocated")
    # A Variant that takes a copy of one: `v = a` with a unallocated leaves v
    # an array with no storage (issue #342, measured in Excel 16.0). So does
    # one given Array or Split and then erased (issue #420).
    fixed_arrays = _fixed_array_locals(member.body, activity)
    for lower, decl in _variant_array_copies(source, member.body, arrays, activity, fixed_arrays).items():
        arrays[lower] = decl
        state[lower] = "unknown"
    if not arrays:
        return
    # The GoTo-following walk runs the body until its labels settle, and
    # reports on its last run (issue #271).
    silent = [False]

    def report(rule: str, message: str, span: Span, data: Any = None) -> None:
        if not silent[0]:
            push(rule, message, span, data)

    def set_silent(quiet: bool) -> None:
        silent[0] = quiet

    walk = walk_straight_line_body if procedure_has_unstructured_flow(source, member, activity) else walk_branch_merged_body
    # A statement a known guard keeps from running (issue #273).
    unreachable = unreachable_statements_in(source, member, symbols, activity)

    def on_statement(stmt: LeafStatementNode) -> None:
        _check_unallocated_statement(source, stmt, arrays, state, report, erased_by_call, fixed_arrays)

    def on_block(node: BodyNode) -> None:
        # The header runs as the block is entered: `For i = 0 To UBound(a)`,
        # `Do While i <= UBound(a)`, `Select Case UBound(a)` (issue #342,
        # measured in Excel 16.0).
        if isinstance(node, (SelectBlockNode, DoBlockNode, WhileBlockNode)) or (
            isinstance(node, ForBlockNode) and not node.each
        ):
            before = block_header_statements(source, node).before
            if before is not None:
                _check_unallocated_statement(source, before, arrays, state, report, None, fixed_arrays)
        # `For Each x In a` over an array with no storage raises 92, For loop
        # not initialized, not 9 (issue #181, measured in Excel 16.0).
        if isinstance(node, ForBlockNode) and node.each:
            over = js_trim(node.source_expression).lower() if node.source_expression else None
            if over and node.source_expression_span is not None and state.get(over) == "unallocated":
                report(
                    "unallocatedDynamicArrayAccess",
                    f"Dynamic array '{arrays[over].name}' is not allocated when For Each asks it for its "
                    "elements. This will raise Run-time error '92': For loop not initialized.",
                    node.source_expression_span,
                )

    def demote(lower: str) -> None:
        state[lower] = "unknown"

    def restore(snapshot: Mapping[str, str]) -> None:
        state.clear()
        state.update(snapshot)

    def set_state(key: str, value: str) -> None:
        state[key] = value

    hooks = StraightLineDataflowHooks(
        on_statement=on_statement,
        on_block=on_block,
        touches_in_statement=lambda stmt: _dynamic_array_touches_in_statement(source, stmt, arrays),
        demote_to_unknown=demote,
        snapshot_state=lambda: dict(state),
        restore_state=restore,
        set_state=set_state,
        lattice=Lattice(init="unallocated", good="allocated", unknown="unknown"),
        set_silent=set_silent,
    )
    walk(
        source,
        member.body,
        lambda node: is_inactive_node(activity, node) or _node_in(unreachable, node),
        hooks,
    )


# Statement heads after which a procedure may end before its last line.
_ERASE_LEAVING_HEADS = frozenset({"exit", "goto", "gosub", "return", "end", "resume", "on", "stop", "error"})

_EMPTY_NAMES: frozenset[str] = frozenset()


def _arrays_erased_by_calls(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None
) -> Callable[[Sequence[VbaToken]], AbstractSet[str]]:
    """The names a statement's calls pass to an array parameter the callee ends
    by erasing: the callee names the parameter only at its top level, the last
    of those is `Erase p`, and nothing in it may leave early (issue #449,
    measured in Excel 16.0: `Free a` then `UBound(a)` raises 9)."""
    erasing: dict[str, set[int]] = {}
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode) or member.proc_kind not in (ProcKind.SUB, ProcKind.FUNCTION):
            continue
        params = [(index, param) for index, param in enumerate(member.params) if param.is_array and not param.by_val]
        if not params:
            continue
        leaves = [False]

        def scan(stmt: LeafStatementNode, leaves: list[bool] = leaves) -> None:
            if leaves[0]:
                return
            for span in statement_and_branch_spans(stmt):
                toks = [tok for tok in statement_tokens(source, span) if tok.kind is not TokenKind.INTEGER_LITERAL]
                head = token_text(_at(toks, 0))
                if head in _ERASE_LEAVING_HEADS or (head == "err" and token_text(_at(toks, 2)) == "raise"):
                    leaves[0] = True

        for_each_statement(member.body, scan, activity)
        if leaves[0]:
            continue
        relevant = {param.name.lower() for _, param in params}
        last_named: dict[str, BodyNode] = {}
        nested: set[str] = set()
        for node in member.body:
            leaf = is_leaf_statement(node)
            active = not is_inactive_node(activity, node)
            # The legacy nested-name check includes inactive non-leaf nodes.
            if not active and leaf:
                continue
            for tok in statement_tokens(source, node.span):
                lower = _lower_name(tok)
                if not lower or lower not in relevant:
                    continue
                if active:
                    last_named[lower] = node
                if not leaf:
                    nested.add(lower)
        for index, param in params:
            lower = param.name.lower()
            last = last_named.get(lower)
            last_toks = statement_tokens(source, last.span) if last is not None and is_leaf_statement(last) else []
            if (
                lower not in nested
                and len(last_toks) == 2
                and token_text(last_toks[0]) == "erase"
                and _lower_name(last_toks[1]) == lower
                and not (isinstance(last, StatementNode) and last.single_line_if_branches)
            ):
                erasing.setdefault(member.name.lower(), set()).add(index)

    def erased(toks: Sequence[VbaToken]) -> AbstractSet[str]:
        out: set[str] = set()
        if not erasing:
            return out
        call = 1 if token_text(_at(toks, 0)) == "call" else 0
        indexes = erasing.get(_lower_name(_at(toks, call)) or "")
        if not indexes or _raw_at(toks, call + 1) in ("=", "."):
            return out
        parens = _raw_at(toks, call + 1) == "(" and match_paren_from(toks, call + 1) == len(toks) - 1
        args = (
            split_top_level_token_groups(toks[call + 2 : len(toks) - 1], 0, ",")
            if parens
            else split_top_level_token_groups(toks[call + 1 :], 0, ",")
        )
        for index in indexes:
            arg = args[index] if index < len(args) else None
            name = _lower_name(arg[0]) if arg is not None and len(arg) == 1 else None
            if name:
                out.add(name)
        return out

    return erased


def _check_unallocated_statement(
    source: str,
    stmt: LeafStatementNode,
    arrays: Mapping[str, _DynamicArrayDeclaration],
    state: dict[str, str],
    push: PushFn,
    erased_by_call: Callable[[Sequence[VbaToken]], AbstractSet[str]] | None = None,
    fixed_arrays: AbstractSet[str] = _EMPTY_NAMES,
) -> None:
    redimmed = _redim_statement_targets(source, stmt.span)
    if redimmed:
        for target in redimmed:
            lower = target.name.lower()
            if lower in arrays and target.dimensions:
                state[lower] = "allocated"
        return
    erased = _erase_statement_simple_targets(source, stmt.span)
    if erased:
        for lower in erased:
            if lower in arrays:
                # An erased Variant may hold a fixed array, which Erase clears and
                # keeps; one known to hold a dynamic array, from Array, Split, ReDim
                # or a copy, is left with none (issue #420, measured in Excel 16.0).
                # One a first Erase already emptied stays empty (issue #624).
                variant = arrays[lower].variant
                before = state.get(lower)
                state[lower] = (
                    "unknown" if variant and before != "allocated" and before != "unallocated" else "unallocated"
                )
        return
    conditional_redims = _single_line_if_redim_targets(source, stmt.span)
    passed_whole: Mapping[str, int] = locals_named_whole(source, stmt.span, arrays, _ARRAY_READ_ONLY_INTRINSICS)

    def follows_pass(name: str, hit_span: Span) -> bool:
        # An access that follows a whole-array pass in the same statement, as in
        # `If Load(a) Then Debug.Print a(0)`, runs after the callee had its chance
        # to allocate. One that precedes it, as in `Load(a(0))`, does not.
        pass_at = passed_whole.get(name.lower())
        return pass_at is not None and hit_span.start > pass_at

    for name, hit_span in _unallocated_dynamic_array_index_accesses(source, stmt.span, arrays, state):
        # The "access" may be the target of a ReDim embedded in a single-line
        # If...Then - the allocation itself, not a read. Suppress exactly that
        # target's name-token span; bounds expressions still report.
        if any(t.span.start == hit_span.start and t.span.end == hit_span.end for t in conditional_redims):
            continue
        if follows_pass(name, hit_span):
            continue
        push(
            "unallocatedDynamicArrayAccess",
            f"Dynamic array '{name}' is not allocated before indexed access. "
            "This will raise Run-time error '9': Subscript out of range.",
            hit_span,
        )
    for function_name, name, hit_span in _unallocated_dynamic_array_bound_calls(source, stmt.span, arrays, state):
        if follows_pass(name, hit_span):
            continue
        push(
            "unallocatedDynamicArrayAccess",
            f"Dynamic array '{name}' is not allocated before {function_name}. "
            "This will raise Run-time error '9': Subscript out of range.",
            hit_span,
        )
    assignment = bare_assignment_target(source, stmt.span)
    assignment_lower = assignment[0].lower() if assignment is not None else None
    if assignment is not None and assignment_lower and assignment_lower in arrays:
        # `a = b` copies b's storage, or its lack of it (issue #342).
        value = _without_comments(assignment[2])
        source_name = _lower_name(value[0]) if len(value) == 1 else None
        if source_name and source_name in arrays:
            state[assignment_lower] = state.get(source_name) or "unknown"
        elif _dynamic_array_call(value) or (
            # A Variant's copy of a fixed array is a dynamic array with its
            # storage (issue #685, measured in Excel 16.0: Erase then empties it).
            source_name is not None and source_name in fixed_arrays and arrays[assignment_lower].variant
        ):
            state[assignment_lower] = "allocated"
        else:
            state[assignment_lower] = "unknown"
    # `Free a`, whose callee ends by erasing its parameter, leaves a unallocated
    # (issue #449, measured in Excel 16.0).
    erased_there: AbstractSet[str] = (
        erased_by_call(statement_tokens(source, stmt.span))
        if passed_whole and erased_by_call is not None
        else _EMPTY_NAMES
    )
    for lower in passed_whole:
        declaration = arrays.get(lower)
        if lower in erased_there and (declaration is None or declaration.variant is not True):
            state[lower] = "unallocated"
        elif state.get(lower) == "unallocated":
            state[lower] = "unknown"
    # A conditional (single-line If) ReDim allocates only on one path, so move
    # the array to 'unknown' - mirroring how block-If allocations degrade - not
    # 'allocated'. The 'unallocated' guard keeps an already-allocated array precise.
    for target in conditional_redims:
        lower = target.name.lower()
        if lower in arrays and target.dimensions and state.get(lower) == "unallocated":
            state[lower] = "unknown"
    # `If L > 0 Then tb = txt` gives the array storage on one path only.
    for branch in statement_and_branch_spans(stmt)[1:]:
        branch_assignment = bare_assignment_target(source, branch)
        lower_branch = branch_assignment[0].lower() if branch_assignment is not None else None
        if lower_branch and lower_branch in arrays and state.get(lower_branch) == "unallocated":
            state[lower_branch] = "unknown"


def _local_dynamic_array_declarations_for_body(
    body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> dict[str, _DynamicArrayDeclaration]:
    out: dict[str, _DynamicArrayDeclaration] = {}

    def visit(group: VariableGroupNode) -> None:
        if group.is_const or group.modifier == "Static":
            return
        for decl in group.declarations:
            if not decl.is_array or decl.array_bounds:
                continue
            lower = decl.name.lower()
            if lower not in out:
                out[lower] = _DynamicArrayDeclaration(name=decl.name, span=decl.span)

    for_each_variable_group(body, visit, activity)
    return out


def _variant_array_copies(
    source: str,
    body: Sequence[BodyNode],
    arrays: Mapping[str, _DynamicArrayDeclaration],
    activity: ConditionalActivityTracker | None,
    fixed_arrays: AbstractSet[str] = _EMPTY_NAMES,
) -> dict[str, _DynamicArrayDeclaration]:
    """The Variant locals some statement assigns one of the dynamic arrays whole: `v = a`."""
    variants: dict[str, _DynamicArrayDeclaration] = {}

    def visit_group(group: VariableGroupNode) -> None:
        if group.is_const or group.modifier == "Static":
            return
        for decl in group.declarations:
            type_ = normalize_type(decl.as_type)
            if not decl.is_array and not decl.type_suffix and (type_ is None or type_ == "variant"):
                variants[decl.name.lower()] = _DynamicArrayDeclaration(name=decl.name, span=decl.span, variant=True)

    for_each_variable_group(body, visit_group, activity)
    out: dict[str, _DynamicArrayDeclaration] = {}
    if not variants:
        return out

    def visit(stmt: LeafStatementNode) -> None:
        bare = bare_assignment_target(source, stmt.span)
        value = _without_comments(bare[2]) if bare is not None else []
        lower = bare[0].lower() if bare is not None else ""
        source_name = (_lower_name(value[0]) or "") if len(value) == 1 else ""
        if lower in variants and (source_name in arrays or source_name in fixed_arrays or _dynamic_array_call(value)):
            out[lower] = variants[lower]

    for_each_statement(body, visit, activity)
    return out


def _fixed_array_locals(body: Sequence[BodyNode], activity: ConditionalActivityTracker | None) -> set[str]:
    """The fixed-size array locals of a body, lowercased (issue #685)."""
    out: set[str] = set()

    def visit(group: VariableGroupNode) -> None:
        if group.is_const:
            return
        for decl in group.declarations:
            if decl.is_array and decl.array_bounds is not None:
                out.add(decl.name.lower())

    for_each_variable_group(body, visit, activity)
    return out


def _dynamic_array_call(value: Sequence[VbaToken]) -> bool:
    """`Array(1, 2)` or `Split(s, ",")`, whole: a dynamic array (issue #420)."""
    at = 2 if token_text(_at(value, 0)) == "vba" and _raw_at(value, 1) == "." else 0
    name = token_text(_at(value, at))
    return (
        name in ("array", "split")
        and _raw_at(value, at + 1) == "("
        and match_paren_from(value, at + 1) == len(value) - 1
    )


def _unallocated_dynamic_array_index_accesses(
    source: str,
    span: Span,
    arrays: Mapping[str, _DynamicArrayDeclaration],
    state: Mapping[str, str],
) -> list[tuple[str, Span]]:
    toks = statement_tokens_after_leading_label(source, span)
    out: list[tuple[str, Span]] = []
    for i in range(len(toks) - 1):
        if toks[i + 1].raw_text != "(" or _raw_at(toks, i - 1) in (".", "!"):
            continue
        name = token_name(toks[i])
        lower = name.lower() if name else None
        if not name or not lower or lower not in arrays or state.get(lower) != "unallocated":
            continue
        if match_paren_from(toks, i + 1) <= i + 1:
            continue
        out.append((name, Span(span.start + toks[i].start, span.start + toks[i].end)))
    return out


_EMPTY_PARENS_RETURN_RE = re.compile(f"\\([{JS_WHITESPACE}]*\\)[{JS_WHITESPACE}]*\\Z")


def _array_functions_never_set(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None
) -> dict[str, ProcedureNode]:
    """The module's Functions declared to return an array, `As Long()`, whose
    body never names the result: each returns an array with no storage, so
    `UBound(F())` raises 9 (issue #448, measured in Excel 16.0)."""
    out: dict[str, ProcedureNode] = {}
    seen: set[str] = set()
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        lower = member.name.lower()
        if lower in seen:
            out.pop(lower, None)
            continue
        seen.add(lower)
        if member.proc_kind is not ProcKind.FUNCTION or not _EMPTY_PARENS_RETURN_RE.search(member.return_type or ""):
            continue
        # The header names it once; any other mention may set it.
        toks = statement_tokens(source, member.span)
        named = sum(
            1 for i, tok in enumerate(toks) if _lower_name(tok) == lower and _raw_at(toks, i - 1) != "."
        )
        if named == 1:
            out[lower] = member
    return out


def _check_unset_array_results(
    source: str,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    functions: Mapping[str, ProcedureNode],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """`UBound(F())` or `LBound(F)` on a Function that never sets its array result."""
    proc_sym = procedure_symbol_for(symbols, member)
    own = {
        member.name.lower(),
        *(param.name.lower() for param in member.params),
        *(child.name.lower() for child in ((proc_sym.children if proc_sym is not None else None) or [])),
    }

    def visit(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            toks = statement_tokens(source, span)
            for i in range(len(toks) - 2):
                bound = token_text(toks[i])
                if (
                    bound not in ("lbound", "ubound")
                    or toks[i + 1].raw_text != "("
                    or not is_bare_or_vba_qualified_intrinsic_call(toks, i)
                ):
                    continue
                lower = _lower_name(toks[i + 2]) or ""
                fn = None if lower in own else functions.get(lower)
                end = i + 4 if _raw_at(toks, i + 3) == "(" and _raw_at(toks, i + 4) == ")" else i + 2
                if (
                    fn is None
                    or (end == i + 2 and len(fn.params) > 0)
                    or (_raw_at(toks, end + 1) != ")" and _raw_at(toks, end + 1) != ",")
                ):
                    continue
                push(
                    "unallocatedDynamicArrayAccess",
                    f"Function '{fn.name}' never sets its result, so it returns an array with no storage, "
                    f"and {toks[i].raw_text} has no bounds to read. This will raise Run-time error '9': "
                    "Subscript out of range.",
                    Span(span.start + toks[i + 2].start, span.start + toks[end].end),
                )

    for_each_statement(member.body, visit, activity)


def _unallocated_dynamic_array_bound_calls(
    source: str,
    span: Span,
    arrays: Mapping[str, _DynamicArrayDeclaration],
    state: Mapping[str, str],
) -> list[tuple[str, str, Span]]:
    toks = statement_tokens(source, span)
    out: list[tuple[str, str, Span]] = []
    for i in range(len(toks) - 2):
        function_name = token_name(toks[i])
        lower_function = function_name.lower() if function_name else None
        if lower_function not in ("lbound", "ubound"):
            continue
        if _raw_at(toks, i + 1) != "(" or not is_bare_or_vba_qualified_intrinsic_call(toks, i):
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0:
            continue
        split = split_arg_slots(toks[i + 2 : close], span.start)
        first_slot = split.slots[0] if split.slots else []
        if len(first_slot) != 1:
            continue
        name = token_name(first_slot[0])
        lower = name.lower() if name else None
        if not name or not lower or lower not in arrays or state.get(lower) != "unallocated":
            continue
        assert function_name is not None
        hit_span = (
            split.spans[0]
            if split.spans
            else Span(span.start + first_slot[0].start, span.start + first_slot[0].end)
        )
        out.append((function_name, name, hit_span))
    return out


def _dynamic_array_touches_in_statement(
    source: str, stmt: LeafStatementNode, arrays: Mapping[str, _DynamicArrayDeclaration]
) -> set[str]:
    out: set[str] = set()
    for target in (
        *_redim_statement_targets(source, stmt.span),
        *_single_line_if_redim_targets(source, stmt.span),
    ):
        lower = target.name.lower()
        if lower in arrays:
            out.add(lower)
    for lower in _erase_statement_simple_targets(source, stmt.span):
        if lower in arrays:
            out.add(lower)
    assignment = bare_assignment_target(source, stmt.span)
    assignment_lower = assignment[0].lower() if assignment is not None else None
    if assignment_lower and assignment_lower in arrays:
        out.add(assignment_lower)
    for branch in statement_and_branch_spans(stmt)[1:]:
        branch_assignment = bare_assignment_target(source, branch)
        lower_branch = branch_assignment[0].lower() if branch_assignment is not None else None
        if lower_branch and lower_branch in arrays:
            out.add(lower_branch)
    out.update(locals_named_whole(source, stmt.span, arrays, _ARRAY_READ_ONLY_INTRINSICS).keys())
    return out


# Intrinsics that read an array argument and allocate nothing.
_ARRAY_READ_ONLY_INTRINSICS: frozenset[str] = frozenset({"lbound", "ubound", "isarray"})


def _erase_statement_simple_targets(source: str, span: Span) -> set[str]:
    toks = statement_tokens_after_leading_label(source, span)
    if token_text(_at(toks, 0)) != "erase":
        return set()
    out: set[str] = set()
    for group in split_top_level_token_groups(toks, 1, ","):
        content = _without_comments(group)
        if len(content) != 1:
            continue
        name = token_name(content[0])
        if name:
            out.add(name.lower())
    return out


# -- checkEraseTargets -------------------------------------------------------


def check_erase_targets(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """Rule: Erase targets must be variable/array target names, not arbitrary
    expressions. This intentionally stays syntax-shaped: array-ness/type
    resolution is a separate binder-backed slice."""

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        shapes = declaration_shape_environment_for(symbols, member)
        proc_sym = procedure_symbol_for(symbols, member)

        def resolve_shape(name: str) -> SourceDeclaredShape:
            return declared_shape_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.ASSIGNMENT_TARGET
            )

        def visitor(stmt: LeafStatementNode) -> None:
            for span in _invalid_erase_targets(source, stmt.span):
                push(
                    "invalidEraseTarget",
                    "Erase target must be a variable or array name, not an arbitrary expression.",
                    span,
                )
            for name, span, as_type in _erase_scalar_targets(source, stmt.span, shapes, resolve_shape):
                push(
                    "eraseRequiresArray",
                    f"Erase target '{name}' must be an array or Variant, but it is declared As {as_type}.",
                    span,
                )

        return visitor

    return factory


def _invalid_erase_targets(source: str, span: Span) -> list[Span]:
    toks = statement_tokens_after_leading_label(source, span)
    if token_text(_at(toks, 0)) != "erase":
        return []
    out: list[Span] = []
    for group in split_top_level_token_groups(toks, 1, ","):
        content = _without_comments(group)
        if not content:
            continue
        if _erase_target_looks_variable_like(content):
            continue
        out.append(_token_group_span(span, content))
    return out


def _erase_scalar_targets(
    source: str,
    span: Span,
    shapes: Mapping[str, DeclaredValueShape],
    resolve_shape: Callable[[str], SourceDeclaredShape] | None = None,
) -> list[tuple[str, Span, str]]:
    toks = statement_tokens_after_leading_label(source, span)
    if token_text(_at(toks, 0)) != "erase":
        return []
    out: list[tuple[str, Span, str]] = []
    for group in split_top_level_token_groups(toks, 1, ","):
        content = _without_comments(group)
        if len(content) != 1:
            continue
        name = token_name(content[0])
        if not name:
            continue
        resolved = resolve_shape(name) if resolve_shape is not None else None
        shape = resolved.shape if resolved is not None and resolved.resolved else shapes.get(name.lower())
        if shape is None or shape.is_array or not shape.as_type:
            continue
        normalized = normalize_type(shape.as_type)
        if not normalized or normalized == "variant":
            continue
        if normalized == "object" or is_known_scalar_type(normalized):
            out.append((name, _token_group_span(span, content), shape.as_type))
    return out


def _erase_target_looks_variable_like(toks: Sequence[VbaToken]) -> bool:
    if not token_name(toks[0]):
        return False
    if any(tok.raw_text in _ERASE_EXPRESSION_OPERATORS for tok in toks):
        return False
    return toks[0].raw_text != "("


_ERASE_EXPRESSION_OPERATORS = frozenset({"+", "-", "*", "/", "\\", "^", "&", "=", "<", ">", "<=", ">=", "<>"})


def _token_group_span(base: Span, toks: Sequence[VbaToken]) -> Span:
    first = toks[0] if toks else None
    last = toks[-1] if toks else None
    return Span(base.start + (first.start if first else 0), base.start + (last.end if last else 0))


# -- array shapes --------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class ArrayDimensionBound:
    """One dimension of an array whose bounds the code fixes."""

    lower: int | float
    upper: int | float
    # False when the lower bound came from Option Base rather than the text.
    explicit_lower: bool


@dataclass(frozen=True, slots=True)
class FixedArrayBound:
    """An array whose every dimension's bounds are known from the text."""

    name: str
    dims: Sequence[ArrayDimensionBound]
    # Where the bounds came from, for the message: 'Dim', 'Array(...)',
    # 'Split(...)', 'Range(...).Value'.
    origin: str
    # The arrays its elements hold, by position, where Array(...) of Array(...) built it.
    elements: Sequence[FixedArrayBound | None] | None = None
    # The literal each element holds, by position, where it is known (issue #260).
    values: Sequence[str | int | None] | None = None


# The String a name holds at the statement being read, or None (issue #260).
StringValueOf = Callable[[VbaToken], "str | None"]


@dataclass(frozen=True, slots=True)
class SubscriptHit:
    """A subscript the array cannot take; a count the compiler refuses has its own rule."""

    span: Span
    message: str
    rule: str | None = None  # 'wrongNumberOfDimensions'


def parse_fixed_array_bounds_for_decl(
    source: str, decl: Any, option_base: int
) -> list[ArrayDimensionBound] | None:
    """Parses the literal bounds of a fixed-size array declaration (anything with
    a `span`), one entry per dimension. None unless every dimension's upper bound
    (and any explicit lower bound) folds to a literal integer. A dimension with
    no `To` takes Option Base as its lower bound (issue #120: `Option Base 1`
    then `Dim a(3)` refuses `a(0)`)."""
    toks = statement_tokens(source, decl.span)
    open_index = next((i for i, tok in enumerate(toks) if tok.raw_text == "("), -1)
    if open_index < 0:
        return None
    close = match_paren_from(toks, open_index)
    if close < 0:
        return None
    return literal_dimensions(toks[open_index + 1 : close], option_base)


def literal_dimensions(inner: Sequence[VbaToken], option_base: int) -> list[ArrayDimensionBound] | None:
    """The bounds a parenthesised bounds list states, `2, 1 To 3`, when every one is a literal."""
    dims = [
        dim_tokens
        for dim_tokens in (_without_comments(part) for part in split_top_level_token_groups(inner, 0, ","))
        if dim_tokens
    ]
    if not dims:
        return None
    out: list[ArrayDimensionBound] = []
    for dim in dims:
        _key, _lower_key, lower_value, upper_value = _comparable_array_bound_key(dim)
        if upper_value is None:
            return None  # a Const or variable bound is not statically known
        has_to = any(token_text(tok) == "to" for tok in dim)
        if has_to and lower_value is None:
            return None
        out.append(
            ArrayDimensionBound(
                lower=lower_value if lower_value is not None else option_base,
                upper=upper_value,
                explicit_lower=lower_value is not None,
            )
        )
    return out


def local_fixed_arrays(
    source: str, proc: ProcedureNode, activity: ConditionalActivityTracker | None, option_base: int
) -> dict[str, FixedArrayBound]:
    """A procedure's local fixed arrays whose bounds are literals, by lowercased name."""
    return _local_fixed_array_declarations_for_body(source, proc.body, activity, option_base)


def _local_fixed_array_declarations_for_body(
    source: str,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    option_base: int,
) -> dict[str, FixedArrayBound]:
    """Local, statically-bounded fixed arrays in a procedure body."""
    out: dict[str, FixedArrayBound] = {}

    def visit(group: VariableGroupNode) -> None:
        if group.is_const:
            return
        for decl in group.declarations:
            if not decl.is_array or not decl.array_bounds:
                continue  # dynamic arrays take their bounds from ReDim or a value
            lower = decl.name.lower()
            if lower in out:
                continue
            dims = parse_fixed_array_bounds_for_decl(source, decl, option_base)
            if dims is not None:
                out[lower] = FixedArrayBound(name=decl.name, dims=dims, origin="Dim")

    for_each_variable_group(body, visit, activity)
    return out


# `Option Base 0|1` in an Option directive's text; `\s` and `\b` as JavaScript
# reads them.
_OPTION_BASE_RE = re.compile(r"base[" + JS_WHITESPACE + r"]+([01])\b", re.IGNORECASE | re.ASCII)


def module_option_base(mod: ModuleNode, activity: ConditionalActivityTracker | None) -> int:
    """The module's `Option Base`, 0 when absent."""
    for member in active_module_members(mod, activity):
        if isinstance(member, OptionNode):
            match = _OPTION_BASE_RE.match(js_trim(member.option_text))
            if match:
                return int(match.group(1))
    return 0


# Heads of statements that can reshape an array or Variant named anywhere in them.
_SHAPE_SPOILING_HEADS = frozenset({"redim", "erase", "set", "input", "get", "line"})


def known_array_shapes(
    source: str,
    body: Sequence[BodyNode],
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    option_base: int,
) -> dict[str, FixedArrayBound]:
    """Dynamic-array and Variant locals whose bounds a value fixes (issue #120):
    the local's ONLY assignment is `Array(...)`, `VBA.Array(...)`,
    `Split(literal, literal[, limit])` or `Range("A1:B2").Value`, and nothing
    else touches it (no ReDim, Erase, whole pass to a call, or Set).

     - `Array(a, b)` is based at Option Base; `VBA.Array` ignores Option Base and
       is based at 0 (both measured in Excel 16.0).
     - `Array()` has UBound -1: every index is out of range.
     - `Split` is always 0-based and yields one part per delimiter plus one;
       Split("abc", ",") is one element, Split("a,b", ",") two. Split("") is
       empty, UBound -1, like `Array()` (issue #181).
     - A Range's `.Value` over a multi-cell address literal is a 1-based
       two-dimensional array of the address's rows and columns.
    """
    candidates = _array_value_locals(symbols, proc)
    if not candidates:
        return {}
    strings_at = _string_values_at(source, symbols, proc, activity)
    compare = module_compare(source)
    assignments: dict[str, list[FixedArrayBound]] = {}
    spoiled: set[str] = set()

    def spoil(lower: str | None) -> None:
        if lower and lower in candidates:
            spoiled.add(lower)

    def visit(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            toks = statement_tokens_after_leading_label(source, span)
            head = token_text(_at(toks, 0))
            bare = bare_assignment_target(source, span)
            if bare is not None:
                target_name, _target_span, value_tokens = bare
                lower = target_name.lower()
                if lower not in candidates:
                    continue
                shape = array_value_shape(value_tokens, target_name, option_base, strings_at(stmt), compare)
                if shape is None:
                    spoil(lower)
                else:
                    assignments.setdefault(lower, []).append(shape)
                continue
            if head in _SHAPE_SPOILING_HEADS:
                for tok in toks:
                    spoil(_lower_name(tok))
                continue
            # Passed whole to a call: `Fill v`, `Fill(v)`, `x = Fill(v)`.
            for i, tok in enumerate(toks):
                name = _lower_name(tok)
                if not name or name not in candidates:
                    continue
                prev = _at(toks, i - 1)
                nxt = _at(toks, i + 1)
                if nxt is not None and nxt.raw_text == "(":
                    continue  # an index, not a whole pass
                opens_slot = (
                    prev is None
                    or prev.raw_text in ("(", ",")
                    or prev.kind is TokenKind.IDENTIFIER
                    or prev.kind is TokenKind.KEYWORD
                )
                closes_slot = nxt is None or nxt.raw_text in (")", ",", ":") or nxt.kind is TokenKind.COMMENT
                if (
                    opens_slot
                    and closes_slot
                    and not (prev is not None and prev.kind is TokenKind.OPERATOR)
                    and not (nxt is not None and nxt.kind is TokenKind.OPERATOR)
                    and token_text(prev) != "in"
                ):
                    spoil(name)

    for_each_statement(body, visit, activity)
    out: dict[str, FixedArrayBound] = {}
    for lower, shapes in assignments.items():
        if lower not in spoiled and len(shapes) == 1:
            out[lower] = shapes[0]
    return out


def _string_values_at(
    source: str,
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    activity: ConditionalActivityTracker | None,
) -> Callable[[BodyNode], StringValueOf]:
    """What a String local or a String Const holds as each statement starts, for
    the text and delimiter of `Split(s, d)` and the match of `Filter` (issue
    #260). A local is read where the array is built, not where it is indexed."""
    values_at: list[Any] = []
    consts: list[Mapping[str, str]] = []

    def at(stmt: BodyNode) -> StringValueOf:
        def value_of(tok: VbaToken) -> str | None:
            lower = _lower_name(tok)
            if not lower:
                return None
            if not values_at:
                values_at.append(known_local_literal_values_at(source, proc, symbols, activity))
            local = values_at[0](stmt).get(lower)
            if local:
                return local.value if local.kind == "string" and not local.content_mutated else None
            if not consts:
                consts.append(string_constants_in_scope(symbols, proc))
            return consts[0].get(lower)

        return value_of

    return at


def _array_value_locals(symbols: ModuleSymbols, proc: ProcedureNode) -> dict[str, str]:
    """The dynamic-array and Variant locals a value can give bounds to, lowercased to declared name."""
    out: dict[str, str] = {}
    proc_sym = procedure_symbol_for(symbols, proc)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if child.kind is not VbaSymbolKind.LOCAL_VARIABLE or child.visibility is SymbolVisibility.STATIC:
            continue
        type_ = normalize_type(child.as_type)
        if (child.array_bounds is None) if child.is_array else (type_ is None or type_ == "variant"):
            out[child.name.lower()] = child.name
    return out


def known_array_shapes_at(
    source: str,
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    option_base: int,
) -> Callable[[LeafStatementNode], Mapping[str, FixedArrayBound]]:
    """known_array_shapes at each statement (issue #180). Where the last
    assignment to reach a statement in a straight line builds an array, the
    statement sees its bounds, though the local is assigned again elsewhere:
    `v = Array(1, 2): Debug.Print v(2): v = Array(1, 2, 3)` reads past the end.
    Where it reaches with any other value, the local is not known to be an
    array there."""
    whole = known_array_shapes(source, proc.body, symbols, proc, activity, option_base)
    compare = module_compare(source)
    locals_ = _array_value_locals(symbols, proc)
    reaching: Any = {} if not locals_ else straight_line_assignments(source, proc.body, activity)
    # Each assignment's shape, built with the Strings its own statement sees and
    # keyed by its first value token, which the reaching value shares: by the
    # token's id(), with the token kept beside the shape.
    built: dict[int, tuple[VbaToken, FixedArrayBound | None]] = {}
    if locals_:
        strings_at = _string_values_at(source, symbols, proc, activity)

        def visit(stmt: LeafStatementNode) -> None:
            bare = bare_assignment_target(source, stmt.span)
            if bare is None:
                return
            first = next((tok for tok in bare[2] if tok.kind is not TokenKind.COMMENT), None)
            lower = bare[0].lower()
            if first is not None and lower in locals_:
                built[id(first)] = (
                    first,
                    array_value_shape(bare[2], locals_[lower], option_base, strings_at(stmt), compare),
                )

        for_each_statement(proc.body, visit, activity)
    # By the reaching map's id(), with the map kept beside the result.
    results: dict[int, tuple[Any, Mapping[str, FixedArrayBound]]] = {}

    def at(stmt: LeafStatementNode) -> Mapping[str, FixedArrayBound]:
        assignments = reaching.get(id(stmt))
        if assignments is None:
            return whole
        entry = results.get(id(assignments))
        if entry is not None and entry[0] is assignments:
            return entry[1]
        next_shapes = dict(whole)
        for lower, value in assignments.items():
            if lower not in locals_:
                continue
            first_built = built.get(id(value[0])) if value else None
            if first_built is not None and first_built[0] is value[0]:
                shape = first_built[1]
            else:
                shape = array_value_shape(value, locals_[lower], option_base, None, compare)
            if shape is not None:
                next_shapes[lower] = shape
            else:
                next_shapes.pop(lower, None)
        results[id(assignments)] = (assignments, next_shapes)
        return next_shapes

    return at


def redim_shapes_at(
    source: str,
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    option_base: int,
) -> StatementShapes:
    """The bounds the last `ReDim` gave a dynamic array or Variant local, at each
    statement it reaches (issue #238): `ReDim a(3)` then `a(7)` is error 9. A
    ReDim whose bounds are not literals, any other whole mention of the name
    (`Erase a`, `a = b`, `Fill a`), a ReDim inside a single-line If, and a label
    or GoSub end what is known. Blocks are entered as issue #237 enters them."""
    out = StatementShapes()
    keep_alive: list[object] = []
    locals_ = _array_value_locals(symbols, proc)
    if not locals_:
        return out
    current: list[dict[str, FixedArrayBound]] = [{}]
    # Only output maps and branch snapshots need copying before mutation: the
    # ids of the retained maps, each also held by `out` or a snapshot.
    retained: set[int] = set()

    def retain(shapes: dict[str, FixedArrayBound]) -> None:
        if id(shapes) not in retained:
            retained.add(id(shapes))
            keep_alive.append(shapes)

    def changed() -> None:
        if id(current[0]) in retained:
            current[0] = dict(current[0])

    def forget(names: Iterable[str]) -> None:
        copied = False
        for lower in names:
            if lower in current[0]:
                if not copied:
                    changed()
                    copied = True
                del current[0][lower]

    # A bound computed from what is known: `n = 2 + 1: ReDim a(n)`,
    # `ReDim a(Len(s) - 1)` with s known, `ReDim a(UBound(b))` with b a fixed
    # local array (issue #350, measured in Excel 16.0).
    values_at: list[Any] = []
    fixed_locals: list[dict[str, FixedArrayBound]] = []

    def computed_value(node: LeafStatementNode, toks: Sequence[VbaToken]) -> float | None:
        if not values_at:
            values_at.append(known_local_literal_values_at(source, proc, symbols, activity))
        known = values_at[0](node)
        parts: list[str] = []
        i = 0
        while i < len(toks):
            word = token_text(toks[i])
            close = match_paren_from(toks, i + 1) if _raw_at(toks, i + 1) == "(" else -1
            arg = _lower_name(_at(toks, i + 2)) if close == i + 3 else None
            if arg and word in ("len", "ubound", "lbound"):
                held = known.get(arg)
                value: int | float | None
                if word == "len":
                    value = utf16_length(held.value) if held is not None and held.kind == "string" else None
                else:
                    if not fixed_locals:
                        fixed_locals.append(local_fixed_arrays(source, proc, activity, option_base))
                    fixed = fixed_locals[0].get(arg)
                    dim = fixed.dims[0] if fixed is not None and fixed.dims else None
                    value = (dim.upper if word == "ubound" else dim.lower) if dim is not None else None
                if value is None:
                    return None
                parts.append(_fmt(value))
                i = close + 1
                continue
            parts.append(toks[i].raw_text)
            i += 1
        no_constants: dict[str, float | None] = {}
        return evaluate_integer_constant_expression(" ".join(parts), with_known_locals(no_constants, known))

    def computed_dimension(node: LeafStatementNode, span: Span) -> ArrayDimensionBound | None:
        toks = _without_comments(raw_expression_tokens(source[span.start : span.end]))
        to = next((i for i, tok in enumerate(toks) if token_text(tok) == "to"), -1)
        upper = computed_value(node, toks if to < 0 else toks[to + 1 :])
        lower_value = None if to < 0 else computed_value(node, toks[:to])
        if upper is None or (to >= 0 and lower_value is None):
            return None
        return ArrayDimensionBound(
            lower=lower_value if lower_value is not None else option_base,
            upper=upper,
            explicit_lower=to >= 0,
        )

    def visit(node: BodyNode) -> None:
        if not is_leaf_statement(node):
            return
        if jump_target_label_declaration(source, node.span):
            forget(list(current[0]))
        toks = statement_tokens_after_leading_label(source, node.span)
        # A ReDim's bounds are not subscripts: `ReDim Preserve a(5)` reads no a(5).
        if current[0] and not any(token_text(tok) == "redim" for tok in toks):
            retain(current[0])
            out.set(node, current[0])
        if token_text(_at(toks, 0)) == "gosub":
            forget(list(current[0]))
            return
        # A ReDim a single-line If runs may not run: it only ends what is known.
        conditional = node.single_line_if_tail is True or (
            isinstance(node, StatementNode) and node.single_line_if_branches is not None
        )
        redims = [] if conditional else _redim_statement_targets(source, node.span)
        reshaped = {target.name.lower() for target in redims}
        forget([lower for lower in _shape_touches(source, node) if lower not in reshaped])
        for target in redims:
            lower = target.name.lower()
            name = locals_.get(lower)
            dims: list[ArrayDimensionBound | None] = []
            for dim in target.dimensions:
                if dim.upper_value is not None and (dim.lower_key is None or dim.lower_value is not None):
                    dims.append(
                        ArrayDimensionBound(
                            lower=dim.lower_value if dim.lower_value is not None else option_base,
                            upper=dim.upper_value,
                            explicit_lower=dim.lower_value is not None,
                        )
                    )
                else:
                    dims.append(computed_dimension(node, dim.span))
            changed()
            if name and dims and all(dim is not None for dim in dims):
                current[0][lower] = FixedArrayBound(
                    name=name, dims=[dim for dim in dims if dim is not None], origin="ReDim"
                )
            else:
                current[0].pop(lower, None)

    def snapshot() -> dict[str, FixedArrayBound]:
        retain(current[0])
        return current[0]

    def restore(saved: dict[str, FixedArrayBound]) -> None:
        current[0] = saved
        retain(saved)

    walk_entering_blocks(
        source,
        proc.body,
        lambda node: is_inactive_node(activity, node),
        visit,
        BlockEnteringState(
            snapshot=snapshot,
            restore=restore,
            forget=forget,
            touches=lambda stmt: _shape_touches(source, stmt),
        ),
    )
    return out


def _shape_touches(source: str, stmt: LeafStatementNode) -> set[str]:
    """The names a statement may reshape: every name in a ReDim, and a name used
    whole rather than indexed, as `Erase a` and `Fill a` use it. `a(1) = 2`
    leaves a's bounds alone."""
    toks = statement_tokens_after_leading_label(source, stmt.span)
    redim = any(token_text(tok) == "redim" for tok in toks)
    out: set[str] = set()
    for i, tok in enumerate(toks):
        lower = _lower_name(tok)
        if not lower or _raw_at(toks, i - 1) in (".", "!"):
            continue
        if redim or _raw_at(toks, i + 1) != "(":
            out.add(lower)
    return out


def array_value_shape(
    value_tokens: Sequence[VbaToken],
    name: str,
    option_base: int,
    strings: StringValueOf | None = None,
    compare: ModuleCompare | None = None,
) -> FixedArrayBound | None:
    """The bounds of the array `Array(...)`, `Split(...)`, `Filter(...)` or
    `Range(...).Value` builds, or None. `strings` gives the String a name holds
    where the value is read, so `Split(s, ",")` is known when s is (issue #260)."""
    return _array_value_shape_at_depth(value_tokens, name, option_base, strings, compare, 0)


def _array_value_shape_at_depth(
    value_tokens: Sequence[VbaToken],
    name: str,
    option_base: int,
    strings: StringValueOf | None,
    compare: ModuleCompare | None,
    nesting: int,
) -> FixedArrayBound | None:
    if nesting >= MAX_EXPRESSION_DEPTH:
        return None
    toks = _without_comments(value_tokens)
    if not toks:
        return None
    index = 0
    vba_qualified = False
    if token_text(toks[0]) == "vba" and _raw_at(toks, 1) == ".":
        vba_qualified = True
        index = 2
    callee = token_text(_at(toks, index))
    if callee in ("array", "split", "filter") and _raw_at(toks, index + 1) == "(":
        close = match_paren_from(toks, index + 1)
        if close != len(toks) - 1:
            return None
        inner = toks[index + 2 : close]
        if callee == "array":
            groups = [] if not inner else split_top_level_token_groups(inner, 0, ",")
            lower = 0 if vba_qualified else option_base
            elements = [
                _array_value_shape_at_depth(group, name, option_base, strings, compare, nesting + 1)
                for group in groups
            ]
            values = [_literal_element_value(group) for group in groups]
            return FixedArrayBound(
                name=name,
                dims=[ArrayDimensionBound(lower=lower, upper=lower + len(groups) - 1, explicit_lower=True)],
                origin="VBA.Array(...)" if vba_qualified else "Array(...)",
                elements=elements if any(element is not None for element in elements) else None,
                values=values if any(value is not None for value in values) else None,
            )
        args = split_top_level_token_groups(inner, 0, ",")
        if callee == "filter":
            return _filter_shape(args, name, option_base, strings, compare, nesting)
        if len(args) < 1 or len(args) > 4:
            return None
        text = _string_argument(args[0], strings)
        delimiter = _string_argument(args[1], strings) if len(args) >= 2 and args[1] else " "
        # Limit -1 keeps every part, 0 none, n at most n; any other negative is
        # error 5, which this leaves alone.
        limit = _signed_integer_argument(args[2]) if len(args) >= 3 and args[2] else -1
        # An empty delimiter splits nothing: the whole text is the one element
        # (issue #685, measured in Excel 16.0: `Split("a,b", "")(1)` raises 9,
        # and `Split("", "")(0)` is "").
        if text is not None and delimiter == "" and limit is not None and limit >= -1:
            parts: list[str | int | None] = [] if limit == 0 else [text]
            return FixedArrayBound(
                name=name,
                dims=[ArrayDimensionBound(lower=0, upper=len(parts) - 1, explicit_lower=True)],
                origin="Split(...)",
                values=parts if parts else None,
            )
        text_compare = _compare_argument(args[3]) if len(args) == 4 else _default_compare(compare, delimiter)
        if (
            text is None
            or delimiter is None
            or len(delimiter) == 0
            or limit is None
            or limit < -1
            or text_compare is None
        ):
            return None
        split_values = [] if len(text) == 0 or limit == 0 else _split_parts(text, delimiter, limit, text_compare)
        if split_values is None:
            return None
        return FixedArrayBound(
            name=name,
            dims=[ArrayDimensionBound(lower=0, upper=len(split_values) - 1, explicit_lower=True)],
            origin="Split(...)",
            values=list(split_values) if split_values else None,
        )
    # `Range("A1:B2").Value` and `Worksheets(1).Range("A1:B2").Value`; Value2 too (issue #278).
    block = range_value_block(toks)
    if block is not None and (block[0] > 1 or block[1] > 1):
        return FixedArrayBound(
            name=name,
            dims=[
                ArrayDimensionBound(lower=1, upper=block[0], explicit_lower=True),
                ArrayDimensionBound(lower=1, upper=block[1], explicit_lower=True),
            ],
            origin="Range(...).Value",
        )
    # `Application.Transpose(Range("A1:A3").Value)`: one column becomes a 1-D
    # array from 1, a row or a block a 2-D one turned over (issue #278,
    # measured in Excel 16.0).
    transposed = _transpose_argument(toks)
    inner_block = range_value_block(transposed) if transposed is not None else None
    if inner_block is not None and (inner_block[0] > 1 or inner_block[1] > 1):
        rows, cols = inner_block
        return FixedArrayBound(
            name=name,
            dims=(
                [ArrayDimensionBound(lower=1, upper=rows, explicit_lower=True)]
                if cols == 1
                else [
                    ArrayDimensionBound(lower=1, upper=cols, explicit_lower=True),
                    ArrayDimensionBound(lower=1, upper=rows, explicit_lower=True),
                ]
            ),
            origin="Transpose(...)",
        )
    return None


# A single-cell or block A1-style address literal: `A1`, `A1:B2`. `\d` as
# JavaScript reads it, ASCII digits only.
_CELL_BLOCK_RE = re.compile(r"([A-Za-z]{1,3})([0-9]+)(?::([A-Za-z]{1,3})([0-9]+))?")


def _column_number(letters: str) -> int:
    """excelAddress.ts columnToIndex."""
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def range_value_block(toks: Sequence[VbaToken]) -> tuple[int | float, int] | None:
    """The (rows, cols) of a literal `[...]Range("A1:B2").Value` or `.Value2`, the whole of `toks`."""
    member = token_text(_at(toks, len(toks) - 1))
    if member not in ("value", "value2") or _raw_at(toks, len(toks) - 2) != "." or _raw_at(toks, len(toks) - 3) != ")":
        return None
    close = len(toks) - 3
    open_index = next(
        (i for i, tok in enumerate(toks) if tok.raw_text == "(" and match_paren_from(toks, i) == close), -1
    )
    if (
        open_index <= 0
        or token_text(toks[open_index - 1]) != "range"
        or close != open_index + 2
        or toks[open_index + 1].kind is not TokenKind.STRING_LITERAL
    ):
        return None
    text = toks[open_index + 1].raw_text[1:-1]
    address = _CELL_BLOCK_RE.fullmatch(text)
    if address is None:
        return None
    last_row = address.group(4) if address.group(4) is not None else address.group(2)
    last_col = address.group(3) if address.group(3) is not None else address.group(1)
    # Number() of the row digits, in doubles as upstream computes them.
    rows = _as_js_number(abs(float(last_row) - float(address.group(2))) + 1)
    cols = abs(_column_number(last_col) - _column_number(address.group(1))) + 1
    return (rows, cols)


def single_cell_value(toks: Sequence[VbaToken]) -> bool:
    """Whether `toks` read one cell's value: `Range("A1")`, `Cells(1, 2)`,
    either with `.Value` or `.Value2`, after any receiver (issue #278)."""
    member = token_text(_at(toks, len(toks) - 1))
    end = (
        len(toks) - 3
        if member in ("value", "value2") and _raw_at(toks, len(toks) - 2) == "."
        else len(toks) - 1
    )
    if _raw_at(toks, end) != ")":
        return False
    open_index = next(
        (i for i, tok in enumerate(toks) if tok.raw_text == "(" and match_paren_from(toks, i) == end), -1
    )
    callee = token_text(toks[open_index - 1]) if open_index > 0 else ""
    if open_index <= 0 or (open_index > 1 and _raw_at(toks, open_index - 2) != "."):
        return False
    args = split_top_level_token_groups(toks[open_index + 1 : end], 0, ",")
    if callee == "range":
        block = (
            _CELL_BLOCK_RE.fullmatch(args[0][0].raw_text[1:-1])
            if len(args) == 1 and len(args[0]) == 1 and args[0][0].kind is TokenKind.STRING_LITERAL
            else None
        )
        return block is not None and (
            block.group(3) is None
            or (block.group(3).lower() == block.group(1).lower() and block.group(4) == block.group(2))
        )
    return (
        callee == "cells"
        and len(args) == 2
        and all(len(arg) == 1 and arg[0].kind is TokenKind.INTEGER_LITERAL for arg in args)
    )


_TRANSPOSE_RECEIVERS = frozenset({"application", "worksheetfunction", "application.worksheetfunction"})


def _transpose_argument(toks: Sequence[VbaToken]) -> list[VbaToken] | None:
    """The argument of `Application.Transpose(...)` or `WorksheetFunction.Transpose(...)` that is the whole of `toks`."""
    at = next(
        (
            i
            for i, tok in enumerate(toks)
            if token_text(tok) == "transpose" and _raw_at(toks, i + 1) == "(" and _raw_at(toks, i - 1) == "."
        ),
        -1,
    )
    receiver = "".join(token_text(tok) for tok in toks[: at - 1]) if at > 0 else ""
    if at < 0 or receiver not in _TRANSPOSE_RECEIVERS or match_paren_from(toks, at + 1) != len(toks) - 1:
        return None
    return list(toks[at + 2 : len(toks) - 1])


def _filter_shape(
    args: Sequence[Sequence[VbaToken]],
    name: str,
    option_base: int,
    strings: StringValueOf | None,
    compare: ModuleCompare | None,
    nesting: int,
) -> FixedArrayBound | None:
    """`Filter(source, match[, include[, compare]])` over an array whose every
    element is a known string or whole number: the elements whose text holds
    `match` (or lacks it, with include False), 0-based whatever Option Base
    says (measured in Excel 16.0, issue #260). An empty match keeps every
    element."""
    if len(args) < 2 or len(args) > 4:
        return None
    source = _array_value_shape_at_depth(args[0], name, option_base, strings, compare, nesting + 1)
    match = _string_argument(args[1], strings)
    include = _boolean_argument(args[2]) if len(args) >= 3 and args[2] else True
    text_compare = _compare_argument(args[3]) if len(args) == 4 else _default_compare(compare, match)
    if source is None or len(source.dims) != 1 or match is None or include is None or text_compare is None:
        return None
    texts: list[str] = []
    count = source.dims[0].upper - source.dims[0].lower
    k = 0
    while k <= count:
        value = _item(source.values, k)
        if value is None or _item(source.elements, k) is not None:
            return None
        texts.append(value if isinstance(value, str) else _fmt(value))
        k += 1
    kept: list[str | int | None] = []
    for text in texts:
        holds: int | None
        if text_compare:
            search = _caseless_search(text, match)
            holds = search(0) if search is not None else None
        else:
            holds = text.find(match)
        if holds is None:
            return None
        if (holds >= 0) == include:
            kept.append(text)
    return FixedArrayBound(
        name=name,
        dims=[ArrayDimensionBound(lower=0, upper=len(kept) - 1, explicit_lower=True)],
        origin="Filter(...)",
        values=kept if kept else None,
    )


def _item(items: Sequence[Any] | None, index: int | float) -> Any:
    """`items?.[index]`: None past either end, for a fractional or negative index too."""
    if items is None or not isinstance(index, int) or index < 0 or index >= len(items):
        return None
    return items[index]


def elements_written_in(
    source: str, proc: ProcedureNode, activity: ConditionalActivityTracker | None
) -> set[str]:
    """The names whose elements a statement of the procedure may write, so the
    values an Array or Split gave them are not trusted (issue #260): `v(1) = 2`,
    a writing statement that mentions them (Set, LSet, Mid, Input #, Get #,
    ReDim, Erase), and an element passed alone to a call, ByRef by default."""
    out: set[str] = set()

    def visit(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            toks = statement_tokens_after_leading_label(source, span)
            head = token_text(_at(toks, 0))
            if head in _ELEMENT_WRITING_HEADS:
                for tok in toks:
                    lower = _lower_name(tok)
                    if lower:
                        out.add(lower)
                continue
            call = top_level_equals_index(toks) < 0
            for i in range(len(toks) - 1):
                lower = _lower_name(toks[i])
                if not lower or toks[i + 1].raw_text != "(":
                    continue
                close = match_paren_from(toks, i + 1)
                nxt = _at(toks, close + 1)
                prev = _at(toks, i - 1)
                target = i == (1 if head == "let" else 0) and nxt is not None and nxt.raw_text == "="
                opens_slot = (
                    (prev is not None and prev.raw_text in ("(", ","))
                    or (
                        call
                        and i > 0
                        and prev is not None
                        and prev.kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD)
                    )
                )
                closes_slot = nxt is None or nxt.raw_text in (")", ",", ":")
                if target or (opens_slot and closes_slot):
                    out.add(lower)

    for_each_statement(proc.body, visit, activity)
    return out


_ELEMENT_WRITING_HEADS = frozenset({"set", "lset", "rset", "mid", "mid$", "input", "get", "line", "redim", "erase"})


@dataclass(frozen=True, slots=True)
class ElementOperand:
    """An operand that reads one element of an array whose values are known."""

    # The operand's first and last token.
    first: int
    last: int
    value: str | int


def element_operand_ending_at(
    toks: Sequence[VbaToken], end: int, shapes: Mapping[str, FixedArrayBound], option_base: int
) -> ElementOperand | None:
    """The element `v(1)` or `Split("1 b")(1)` reads, where the operand ends at
    token `end` (issue #260). `shapes` holds the arrays the locals are known to be."""
    if _raw_at(toks, end) != ")":
        return None
    depth = 0
    for open_index in range(end, -1, -1):
        raw = toks[open_index].raw_text
        if raw == ")":
            depth += 1
        elif raw == "(":
            depth -= 1
            if depth == 0:
                return _element_at(toks, open_index, end, shapes, option_base)
    return None


def element_operand_starting_at(
    toks: Sequence[VbaToken], start: int, shapes: Mapping[str, FixedArrayBound], option_base: int
) -> ElementOperand | None:
    """The element an operand starting at token `start` reads, as element_operand_ending_at."""
    if not token_name(_at(toks, start)) or _raw_at(toks, start + 1) != "(":
        return None
    close = match_paren_from(toks, start + 1)
    if close > 0 and _raw_at(toks, close + 1) == "(":
        close = match_paren_from(toks, close + 1)
    element = element_operand_ending_at(toks, close, shapes, option_base) if close > 0 else None
    nxt = _raw_at(toks, close + 1)
    return element if element is not None and element.first == start and nxt not in ("(", ".", "!") else None


def _element_at(
    toks: Sequence[VbaToken],
    open_index: int,
    close: int,
    shapes: Mapping[str, FixedArrayBound],
    option_base: int,
) -> ElementOperand | None:
    shape: FixedArrayBound | None
    first = open_index - 1
    if _raw_at(toks, first) == ")":
        # `Split(...)(k)`, `Array(...)(k)`, `Filter(...)(k)`.
        depth = 0
        call_open = first
        while call_open >= 0:
            raw = toks[call_open].raw_text
            if raw == ")":
                depth += 1
            elif raw == "(":
                depth -= 1
                if depth == 0:
                    break
            call_open -= 1
        first = call_open - 1
        if first >= 2 and toks[first - 1].raw_text == "." and token_text(toks[first - 2]) == "vba":
            first -= 2
        callee = token_text(_at(toks, call_open - 1))
        if call_open < 1 or callee not in ("split", "array", "filter"):
            return None
        shape = array_value_shape(toks[first:open_index], "", option_base)
    else:
        lower = _lower_name(_at(toks, first))
        shape = shapes.get(lower) if lower else None
    before = _raw_at(toks, first - 1)
    if shape is None or len(shape.dims) != 1 or before in (".", "!"):
        return None
    index = _signed_integer_argument(_without_comments(toks[open_index + 1 : close]))
    position = -1 if index is None else index - shape.dims[0].lower
    value = _item(shape.values, position)
    if value is None or _item(shape.elements, position) is not None:
        return None
    return ElementOperand(first=first, last=close, value=value)


def _split_parts(text: str, delimiter: str, limit: int, text_compare: bool) -> list[str] | None:
    """The parts `Split` returns for a non-empty text, at most `limit` of them unless it is -1."""
    parts: list[str] = []
    start = 0
    # A limit of one returns the original text without comparing a delimiter.
    search: Callable[[int], int] | None
    if text_compare and limit != 1:
        search = _caseless_search(text, delimiter)
    else:

        def search(at: int) -> int:
            return text.find(delimiter, at)

    if search is None:
        return None
    while limit == -1 or len(parts) < limit - 1:
        at = search(start)
        if at < 0:
            break
        parts.append(text[start:at])
        start = at + len(delimiter)
    parts.append(text[start:])
    return parts


def _caseless_search(text: str, needle: str) -> Callable[[int], int] | None:
    """Prepare a vbTextCompare search for ASCII text, whose case folding is certain."""
    if any(ord(ch) > 0x7F for ch in text + needle):
        return None
    folded_text = text.lower()
    folded_needle = needle.lower()
    return lambda start: folded_text.find(folded_needle, start)


def _string_argument(arg: Sequence[VbaToken], strings: StringValueOf | None) -> str | None:
    """A string literal, or a name `strings` knows, as an argument."""
    if len(arg) != 1:
        # `Split(LCase("a,b"), ",")`, `Split("ab" & "c", ",")` (issue #509).
        def name_value(tok: VbaToken) -> str | None:
            return strings(tok) if tok.kind is TokenKind.IDENTIFIER and strings is not None else None

        def integer_value(toks: Sequence[VbaToken]) -> int | None:
            return _signed_integer_argument(_without_comments(toks))

        folded: str | None = fold_string_expression(
            arg, StringFoldContext(name_value=name_value, integer_value=integer_value)
        )
        return folded
    if arg[0].kind is TokenKind.STRING_LITERAL:
        return string_literal_value(arg[0].raw_text)
    return strings(arg[0]) if arg[0].kind is TokenKind.IDENTIFIER and strings is not None else None


def _signed_integer_argument(arg: Sequence[VbaToken]) -> int | None:
    """An integer literal, with a leading minus or not."""
    negative = len(arg) == 2 and arg[0].raw_text == "-"
    if len(arg) != (2 if negative else 1) or arg[-1].kind is not TokenKind.INTEGER_LITERAL:
        return None
    value = parse_vba_integer_literal(arg[-1].raw_text)
    return None if value is None else (-value if negative else value)


def _boolean_argument(arg: Sequence[VbaToken]) -> bool | None:
    """True, False or an integer literal, as Filter's include."""
    word = token_text(arg[0]) if len(arg) == 1 else ""
    if word in ("true", "false"):
        return word == "true"
    value = _signed_integer_argument(arg)
    return None if value is None else value != 0


def _default_compare(compare: ModuleCompare | None, text: str | None) -> bool | None:
    """Whether Split or Filter with no compare argument compares as text: it
    takes the module's Option Compare, and Text and Database both ignore case
    (issues #353 and #405, measured in Excel 16.0 and Access 16.0). Where the
    module's setting is not known, only a delimiter or match with no cased
    letter is settled, since both comparisons then agree."""
    if compare is not None:
        return compare != "binary"
    return False if text is not None and text.lower() == text.upper() else None


def _compare_argument(arg: Sequence[VbaToken]) -> bool | None:
    """Whether a compare argument asks for vbTextCompare; None for anything else than the two."""
    word = token_text(arg[0]) if len(arg) == 1 else ""
    if word in ("vbbinarycompare", "vbtextcompare"):
        return word == "vbtextcompare"
    value = _signed_integer_argument(arg)
    return value == 1 if value in (0, 1) else None


def _literal_element_value(group: Sequence[VbaToken]) -> str | int | None:
    """The string or whole number one `Array(...)` element is written as."""
    toks = _without_comments(group)
    if len(toks) == 1 and toks[0].kind is TokenKind.STRING_LITERAL:
        return string_literal_value(toks[0].raw_text)
    return _signed_integer_argument(toks)


def _redim_target_names_in_body(
    source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> set[str]:
    """Names that are ReDim targets anywhere in the body (defensive exclusion)."""
    out: set[str] = set()

    def visit(stmt: LeafStatementNode) -> None:
        for target in _redim_statement_targets(source, stmt.span):
            out.add(target.name.lower())

    for_each_statement(body, visit, activity)
    return out


# The whole numbers a conversion of a string literal gives under each decimal
# point: CInt("3.5") is 4 or 35.
_WHOLE_STRING_CONVERSIONS: Mapping[str, tuple[int, int]] = {
    "cint": (-32768, 32767),
    "clng": (-2147483648, 2147483647),
    "cbyte": (0, 255),
}


def _whole_conversion_readings(slot: Sequence[VbaToken]) -> tuple[int | float, int | float] | None:
    toks = _without_comments(slot)
    at = 2 if token_text(_at(toks, 0)) == "vba" and _raw_at(toks, 1) == "." else 0
    bounds = _WHOLE_STRING_CONVERSIONS.get(token_text(_at(toks, at)))
    if (
        bounds is None
        or len(toks) != at + 4
        or toks[at + 1].raw_text != "("
        or toks[at + 2].kind is not TokenKind.STRING_LITERAL
        or toks[at + 3].raw_text != ")"
    ):
        return None
    readings_raw = numeric_string_readings(string_literal_value(toks[at + 2].raw_text))
    if readings_raw is None:
        return None
    readings: list[int | float] = [bankers_round(value) for value in readings_raw]
    if all(bounds[0] <= value <= bounds[1] for value in readings):
        return (readings[0], readings[1])
    return None


def _subscript_detail(value: int | float, dim: ArrayDimensionBound, index: int, dims: int) -> str | None:
    """Whether `value` is outside `dim`, with the words for the message when it is."""
    if dim.lower <= value <= dim.upper:
        return None
    which = f" in dimension {index + 1}" if dims > 1 else ""
    if dim.upper < dim.lower:
        return f"has no element to reach{which}: the array is empty (UBound {_fmt(dim.upper)})"
    if value > dim.upper:
        return f"is above the upper bound {_fmt(dim.upper)}{which}"
    lower = _fmt(dim.lower)
    if dim.explicit_lower:
        return f"is below the lower bound {lower}{which}"
    return f"is below the lower bound {lower}{which} (Option Base {lower})"


def _fixed_array_subscript_violations(
    source: str,
    span: Span,
    fixed: Mapping[str, FixedArrayBound],
    excluded: AbstractSet[str],
    counters: CountersAt | None,
    lookup: IntegerConstantLookup | None = None,
    entry_lookup: Callable[[BodyNode], IntegerConstantLookup] | None = None,
) -> list[SubscriptHit]:
    """Literal-subscript accesses of a tracked array that fall outside its
    bounds, and a loop counter's subscript on its first or last pass (issue
    #200): a number, or UBound or LBound of the array it indexes, which a bound
    like `UBound(a) + 1` passes whatever the array holds."""
    toks = statement_tokens_after_leading_label(source, span)
    out: list[SubscriptHit] = []
    for i in range(len(toks) - 1):
        if toks[i + 1].raw_text != "(" or _raw_at(toks, i - 1) in (".", "!"):
            continue
        name = token_name(toks[i])
        lower = name.lower() if name else None
        if not name or not lower:
            continue
        decl = fixed.get(lower) if lower in fixed and lower not in excluded else None
        if decl is None and (counters is None or len(counters) == 0):
            continue
        close = match_paren_from(toks, i + 1)
        if close <= i + 1:
            continue
        arg_toks = _without_comments(toks[i + 2 : close])
        slots = split_top_level_token_groups(arg_toks, 0, ",")
        if any(len(slot) == 0 for slot in slots):
            continue
        if decl is None:
            # UBound(x) or LBound(x) in the loop's bounds says x is an array.
            symbolic = _symbolic_counter_subscript(span, name, lower, slots, counters)
            if symbolic is not None:
                out.append(symbolic)
            continue
        if len(slots) != len(decl.dims):
            out.append(dimension_count_violation(span, toks[i], toks[close], decl, len(slots)))
            continue
        # One report per access: the first dimension that is out of range.
        hit: SubscriptHit | None = None
        for index, slot in enumerate(slots):
            hit = subscript_violation(span, decl, fixed, slot, index, counters, lookup, entry_lookup)
            if hit is not None:
                break
        if hit is None:
            hit = _element_subscript_violation(span, toks, decl, slots, close, lookup)
        if hit is not None:
            out.append(hit)
    return out


def dimension_count_violation(
    span: Span, first: VbaToken, last: VbaToken, shape: FixedArrayBound, given: int
) -> SubscriptHit:
    """`g(1)` on `Dim g(2, 2)`: a subscript count other than the array's
    dimensions (issue #248, measured in Excel 16.0). The compiler knows a Dim's
    dimensions and refuses the line; bounds a ReDim or a value set are found
    when it runs, error 9."""
    at = Span(span.start + first.start, span.start + last.end)
    counts = (
        f"has {pluralize_count(len(shape.dims), 'dimension')}, and "
        f"{pluralize_count(given, 'subscript')} {'is' if given == 1 else 'are'} given here"
    )
    if shape.origin == "Dim":
        return SubscriptHit(
            span=at,
            rule="wrongNumberOfDimensions",
            message=f"Array '{shape.name}' {counts}. This is a VBE compile error: Wrong number of dimensions.",
        )
    return SubscriptHit(
        span=at,
        message=(
            f"Array '{shape.name}' ({shape.origin}) {counts}. "
            "This will raise Run-time error '9': Subscript out of range."
        ),
    )


_NO_SHAPES: Mapping[str, FixedArrayBound] = {}


def shape_subscript_violation(
    span: Span,
    toks: Sequence[VbaToken],
    shape: FixedArrayBound,
    open_index: int,
    lookup: IntegerConstantLookup | None = None,
) -> SubscriptHit | None:
    """The subscripts in the parentheses at `open_index` against an array whose
    bounds are known, and on through the arrays its elements hold where
    Array(...) of Array(...) built it (issue #248, measured in Excel 16.0):
    `v(0)(5)`, `c(1)(5)`."""
    # The element arrays chain one parenthesis group to the next: a loop, so a
    # long chain is not bounded by the recursion limit.
    while True:
        close = match_paren_from(toks, open_index)
        if close <= open_index + 1:
            return None
        slots = split_top_level_token_groups(_without_comments(toks[open_index + 1 : close]), 0, ",")
        if any(len(slot) == 0 for slot in slots):
            return None
        if len(slots) != len(shape.dims):
            return dimension_count_violation(span, toks[open_index], toks[close], shape, len(slots))
        for index, slot in enumerate(slots):
            hit = subscript_violation(span, shape, _NO_SHAPES, slot, index, None, lookup)
            if hit is not None:
                return hit
        step = _element_step(toks, shape, slots, close, lookup)
        if step is None:
            return None
        shape, open_index = step


def _element_subscript_violation(
    span: Span,
    toks: Sequence[VbaToken],
    shape: FixedArrayBound,
    slots: Sequence[Sequence[VbaToken]],
    close: int,
    lookup: IntegerConstantLookup | None = None,
) -> SubscriptHit | None:
    """`v(0)(5)`: the next parentheses, against the array element `v(0)` holds."""
    step = _element_step(toks, shape, slots, close, lookup)
    return shape_subscript_violation(span, toks, step[0], step[1], lookup) if step is not None else None


def _element_step(
    toks: Sequence[VbaToken],
    shape: FixedArrayBound,
    slots: Sequence[Sequence[VbaToken]],
    close: int,
    lookup: IntegerConstantLookup | None,
) -> tuple[FixedArrayBound, int] | None:
    """The array element `v(k)` holds and the parentheses that index it next."""
    if shape.elements is None or len(slots) != 1 or _raw_at(toks, close + 1) != "(":
        return None
    value: float | None = _comparable_array_bound_expression_value(slots[0])
    if value is None and lookup is not None:
        value = evaluate_integer_constant_expression(" ".join(tok.raw_text for tok in slots[0]), lookup)
    index = None if value is None else safe_integer(value - shape.dims[0].lower)
    element = None if index is None else _item(shape.elements, index)
    if element is None:
        return None
    return (dataclasses.replace(element, name=f"{shape.name}({value})"), close + 1)


@dataclass(frozen=True, slots=True)
class _ReturnShapesEntry:
    source: str
    activity: ConditionalActivityTracker | None
    option_base: int
    result: Mapping[str, FixedArrayBound]


# Parse nodes survive analysis passes; facts also depend on the active branch.
_RETURN_SHAPES = IdentityLru(capacity=8)

# Words that may leave a Function before its one return assignment runs.
_RETURN_SKIPPING_WORDS = frozenset({"exit", "goto", "gosub", "return", "resume", "on", "raise", "error", "stop"})

_RETURN_END_PAIRS = frozenset({"function", "if", "select", "with"})


def _function_return_shapes(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, option_base: int
) -> Mapping[str, FixedArrayBound]:
    """The bounds of the array each Function of the module returns, where they
    are known from its body (issue #240, measured in Excel 16.0): one top-level
    `F = r` with r a fixed local array, or `F = Array(1, 2)`, the only statement
    that names F, in a body nothing can leave early. So `F()(5)` reads past the
    end."""
    cached: _ReturnShapesEntry | None = _RETURN_SHAPES.get(mod)
    if (
        cached is not None
        and cached.source == source
        and cached.activity is activity
        and cached.option_base == option_base
    ):
        return cached.result
    out: dict[str, FixedArrayBound] = {}
    _RETURN_SHAPES.put(_ReturnShapesEntry(source, activity, option_base, out), mod)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode) or member.proc_kind is not ProcKind.FUNCTION:
            continue
        lower = member.name.lower()
        body = statement_tokens(source, member.span)
        named = sum(1 for tok in body if _lower_name(tok) == lower)
        if named != 2 or any(
            token_text(tok) in _RETURN_SKIPPING_WORDS
            or (token_text(tok) == "end" and token_text(_at(body, i + 1)) not in _RETURN_END_PAIRS)
            for i, tok in enumerate(body)
        ):
            continue  # the header, and one assignment
        assignment = next(
            (
                node
                for node in member.body
                if is_leaf_statement(node)
                and not (isinstance(node, StatementNode) and node.single_line_if_branches)
                and _bare_assignment_lower(source, node.span) == lower
            ),
            None,
        )
        target = bare_assignment_target(source, assignment.span) if assignment is not None else None
        if target is None:
            continue
        value = _without_comments(target[2])
        local = _lower_name(value[0]) if len(value) == 1 else None
        shape = (
            _local_fixed_array_declarations_for_body(source, member.body, activity, option_base).get(local)
            if local
            else array_value_shape(value, member.name, option_base)
        )
        if shape is not None:
            out[lower] = dataclasses.replace(shape, name=f"{member.name}()", origin=f"returned by {member.name}")
    return out


def _bare_assignment_lower(source: str, span: Span) -> str | None:
    target = bare_assignment_target(source, span)
    return target[0].lower() if target is not None else None


_EMPTY_NAME_SET: frozenset[str] = frozenset()


def _returned_array_subscript_violations(
    source: str,
    span: Span,
    returned: Mapping[str, FixedArrayBound],
    lookup: IntegerConstantLookup,
    parameterless: AbstractSet[str] = _EMPTY_NAME_SET,
) -> list[SubscriptHit]:
    """`F()(5)` and `F(1)(5)`: a subscript on the array a Function returns."""
    if not returned:
        return []
    toks = statement_tokens_after_leading_label(source, span)
    out: list[SubscriptHit] = []
    for i in range(len(toks) - 1):
        name_lower = _lower_name(toks[i]) or ""
        shape = returned.get(name_lower)
        if shape is None or toks[i + 1].raw_text != "(" or _raw_at(toks, i - 1) in (".", "!"):
            continue
        call = match_paren_from(toks, i + 1)
        # A Function of no parameters takes `Arr(5)` as a subscript on what it
        # returns (issue #685, measured in Excel 16.0).
        direct = name_lower in parameterless and call > i + 1 and _raw_at(toks, call + 1) != "("
        if call < 0 or (not direct and _raw_at(toks, call + 1) != "("):
            continue
        open_index = i + 1 if direct else call + 1
        close = match_paren_from(toks, open_index)
        slots = (
            [] if close < 0 else split_top_level_token_groups(_without_comments(toks[open_index + 1 : close]), 0, ",")
        )
        if not slots or any(len(slot) == 0 for slot in slots):
            continue
        if len(slots) != len(shape.dims):
            out.append(
                SubscriptHit(
                    span=Span(span.start + toks[open_index + 1].start, span.start + toks[close - 1].end),
                    message=(
                        f"{shape.name} returns an array of {len(shape.dims)} dimension(s), and {len(slots)} "
                        "subscripts are given. This will raise Run-time error '9': Subscript out of range."
                    ),
                )
            )
            continue
        for index, slot in enumerate(slots):
            hit = subscript_violation(span, shape, returned, slot, index, None, lookup)
            if hit is not None:
                out.append(hit)
                break
    return out


def _module_fixed_array_declarations(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, option_base: int
) -> dict[str, FixedArrayBound]:
    """A module's fixed arrays, by lowercased name."""
    groups: list[BodyNode] = [member for member in active_module_members(mod, activity) if isinstance(member, VariableGroupNode)]
    return _local_fixed_array_declarations_for_body(source, groups, activity, option_base)


def _hidden_in(symbols: ModuleSymbols, proc: ProcedureNode) -> set[str]:
    """The names a procedure's own locals and parameters take, which hide module variables."""
    proc_sym = procedure_symbol_for(symbols, proc)
    out = {child.name.lower() for child in ((proc_sym.children if proc_sym is not None else None) or [])}
    for param in proc.params:
        out.add(param.name.lower())
    return out


def _unallocated_module_array_uses(
    source: str, span: Span, arrays: Mapping[str, VbaSymbol]
) -> list[SubscriptHit]:
    """`a(0)`, `UBound(a)` and `LBound(a)` on a module's dynamic array that
    nothing ReDims: it has no elements, and each raises 9 (issue #241,
    measured in Excel 16.0)."""
    toks = statement_tokens_after_leading_label(source, span)
    out: list[SubscriptHit] = []
    for i, tok in enumerate(toks):
        variable = arrays.get(_lower_name(tok) or "")
        if variable is None or _raw_at(toks, i - 1) in (".", "!"):
            continue
        bound = (
            _raw_at(toks, i - 1) == "("
            and token_text(_at(toks, i - 2)) in ("ubound", "lbound")
            and _raw_at(toks, i - 3) != "."
            and _raw_at(toks, i + 1) in (")", ",")
        )
        if not bound and _raw_at(toks, i + 1) != "(":
            continue
        scope = (
            "the project"
            if variable.visibility in (SymbolVisibility.PUBLIC, SymbolVisibility.GLOBAL)
            else "this module"
        )
        what = (
            f"{toks[i - 2].raw_text} reads the bounds of '{tok.raw_text}'"
            if bound
            else f"'{tok.raw_text}' is indexed"
        )
        out.append(
            SubscriptHit(
                span=Span(span.start + tok.start, span.start + tok.end),
                message=(
                    f"{what}, a dynamic array nothing in {scope} ever ReDims, so it has no elements. "
                    "This will raise Run-time error '9': Subscript out of range."
                ),
            )
        )
    return out


def _rounded_decimal_literal(slot: Sequence[VbaToken]) -> tuple[str, int | float] | None:
    """A signed decimal literal, `2.6` or `-0.6`, and the whole number VBA rounds it to."""
    toks = _without_comments(slot)
    negative = len(toks) == 2 and toks[0].raw_text == "-"
    literal = _at(toks, 1 if negative else 0)
    if len(toks) != (2 if negative else 1) or literal is None or literal.kind is not TokenKind.FLOAT_LITERAL:
        return None
    read = js_number(_FLOAT_TYPE_SUFFIX_RE.sub("", literal.raw_text))
    if not math.isfinite(read):
        return None
    value = -read if negative else read
    whole: int | float = bankers_round(value)
    return ("".join(tok.raw_text for tok in toks), whole + 0)


_SUBSCRIPT_ERROR = "This will raise Run-time error '9': Subscript out of range."


def subscript_violation(
    span: Span,
    decl: FixedArrayBound,
    fixed: Mapping[str, FixedArrayBound],
    slot: Sequence[VbaToken],
    index: int,
    counters: CountersAt | None,
    lookup: IntegerConstantLookup | None = None,
    entry_lookup: Callable[[BodyNode], IntegerConstantLookup] | None = None,
) -> SubscriptHit | None:
    """One subscript of an access against the array's bounds. `entry_lookup`
    gives the values as a loop starts, which its For bounds read (issue #685)."""
    dim = decl.dims[index]
    slot_span = Span(span.start + slot[0].start, span.start + slot[-1].end)
    origin = "" if decl.origin == "Dim" else f" ({decl.origin})"
    value = _comparable_array_bound_expression_value(slot)
    if value is not None:
        detail = _subscript_detail(value, dim, index, len(decl.dims))
        return (
            SubscriptHit(
                span=slot_span,
                message=f"Subscript {value} for array '{decl.name}'{origin} {detail}. {_SUBSCRIPT_ERROR}",
            )
            if detail
            else None
        )
    # `a(2.6)` rounds half to even, to 3 (issue #286, measured in Excel 16.0).
    decimal = _rounded_decimal_literal(slot)
    if decimal is not None:
        decimal_text, whole = decimal
        detail = _subscript_detail(whole, dim, index, len(decl.dims))
        return (
            SubscriptHit(
                span=slot_span,
                message=(
                    f"Subscript {decimal_text} rounds to {_fmt(whole)}, which for array '{decl.name}'{origin} "
                    f"{detail}. {_SUBSCRIPT_ERROR}"
                ),
            )
            if detail
            else None
        )
    # `a(i)`, and `a(i + 1)` or `a(i - 1)` a whole number off it (issue #263).
    offset_literal = (
        parse_vba_integer_literal(slot[2].raw_text)
        if len(slot) == 3 and slot[1].raw_text in ("+", "-") and slot[2].kind is TokenKind.INTEGER_LITERAL
        else None
    )
    offset = 0 if offset_literal is None else (-offset_literal if slot[1].raw_text == "-" else offset_literal)
    counter = (
        counters.get(_lower_name(slot[0]) or "")
        if counters is not None and (len(slot) == 1 or offset_literal is not None)
        else None
    )
    if counter is None:
        # A Const, or a local with one known value here (issue #238).
        text = " ".join(tok.raw_text for tok in slot)
        known = evaluate_integer_constant_expression(text, lookup) if lookup is not None else None
        detail = None if known is None else _subscript_detail(known, dim, index, len(decl.dims))
        if detail:
            return SubscriptHit(
                span=slot_span,
                message=(
                    f"Subscript {text} is {known} here, which for array '{decl.name}'{origin} {detail}. "
                    f"{_SUBSCRIPT_ERROR}"
                ),
            )
        # `a(CInt("3.5"))`: 4 where "." is the decimal point, 35 where it
        # groups thousands, and out of bounds either way (issue #703).
        readings = _whole_conversion_readings(slot) if known is None else None
        if readings is not None:
            details = [_subscript_detail(reading, dim, index, len(decl.dims)) for reading in readings]
            if all(one is not None for one in details):
                why = details[0] if details[0] == details[1] else "is outside its bounds"
                return SubscriptHit(
                    span=slot_span,
                    message=(
                        f'Subscript {text} is {_fmt(readings[0])} where "." is the decimal point and '
                        f'{_fmt(readings[1])} where "," is, and either for array \'{decl.name}\'{origin} '
                        f"{why}. {_SUBSCRIPT_ERROR}"
                    ),
                )
        return None

    # `a(i)` inside `For i = 0 To 3`: the counter's first and last passes.
    def atom_value(atom: Any, _counter: Any = None) -> int | float | None:
        # `For i = s To 2` with s a local known to hold 1 (issue #346), read as
        # the loop starts, the body may write s after (issue #685).
        if atom.kind == "local":
            at = entry_lookup(counter.loop_node) if entry_lookup is not None else None
            if at is None and counter.loop == "Do":
                at = lookup
            return evaluate_integer_constant_expression(atom.name, at) if at is not None else None
        array = fixed.get(atom.name)
        shape = _item(array.dims, atom.dimension - 1) if array is not None else None
        if atom.kind == "ubound":
            return shape.upper if shape is not None else None
        if atom.kind == "lbound":
            return shape.lower if shape is not None else None
        return None

    for counter_pass in numeric_counter_passes(counter, atom_value):
        detail = _subscript_detail(counter_pass.value + offset, dim, index, len(decl.dims))
        if detail:
            reached = (
                f"Counter '{slot[0].raw_text}' is {_fmt(counter_pass.value)} on its first pass"
                if counter_pass.pass_ == "first"
                else f"Counter '{slot[0].raw_text}' reaches {_fmt(counter_pass.value)} on its last pass"
            )
            subscript = (
                ""
                if offset == 0
                else f", so {' '.join(tok.raw_text for tok in slot)} is {_fmt(counter_pass.value + offset)}"
            )
            return SubscriptHit(
                span=slot_span,
                message=f"{reached}{subscript}, which for array '{decl.name}'{origin} {detail}. {_SUBSCRIPT_ERROR}",
            )
    if offset == 0:
        return _symbolic_counter_subscript(span, decl.name, decl.name.lower(), [slot], counters, index)
    return None


def _symbolic_counter_subscript(
    span: Span,
    name: str,
    lower: str,
    slots: Sequence[Sequence[VbaToken]],
    counters: CountersAt | None,
    only_index: int | None = None,
) -> SubscriptHit | None:
    """`a(i)` where the counter runs past UBound(a) or below LBound(a) of the
    same dimension: `For i = 1 To UBound(a) + 1`."""
    for index, slot in enumerate(slots):
        if only_index is not None and index != 0:
            break
        dimension = (only_index if only_index is not None else index) + 1
        counter = counters.get(_lower_name(slot[0]) or "") if counters is not None and len(slot) == 1 else None
        if counter is None:
            continue
        passes: list[tuple[str, CounterValue | None]] = [("first", counter.first), ("last", counter.last)]
        for which, value in passes:
            atom = value.atom if value is not None else None
            if value is None or atom is None or atom.name != lower or atom.dimension != dimension:
                continue
            if atom.kind == "ubound" and value.offset > 0:
                past = "above its upper bound"
            elif atom.kind == "lbound" and value.offset < 0:
                past = "below its lower bound"
            else:
                continue
            reached = (
                f"is {counter_text(value)} on its first pass"
                if which == "first"
                else f"reaches {counter_text(value)} on its last pass"
            )
            return SubscriptHit(
                span=Span(span.start + slot[0].start, span.start + slot[-1].end),
                message=(
                    f"Counter '{slot[0].raw_text}' {reached}, which for array '{name}' is {past}. "
                    f"{_SUBSCRIPT_ERROR}"
                ),
            )
    return None


def _bound_intrinsic_dimension_violations(
    source: str, span: Span, fixed: Mapping[str, FixedArrayBound], excluded: AbstractSet[str]
) -> list[SubscriptHit]:
    """`UBound(a, 2)` / `LBound(a, 2)` on an array with fewer dimensions raises 9
    (issue #120, measured in Excel 16.0)."""
    toks = statement_tokens_after_leading_label(source, span)
    out: list[SubscriptHit] = []
    for i in range(len(toks) - 1):
        callee = token_text(toks[i])
        if (
            callee not in ("ubound", "lbound")
            or toks[i + 1].raw_text != "("
            or not is_bare_or_vba_qualified_intrinsic_call(toks, i)
        ):
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0:
            continue
        args = split_top_level_token_groups(_without_comments(toks[i + 2 : close]), 0, ",")
        if len(args) != 2 or len(args[0]) != 1:
            continue
        lower = _lower_name(args[0][0])
        decl = fixed.get(lower) if lower else None
        if decl is None or not lower or lower in excluded:
            continue
        dimension = _comparable_array_bound_expression_value(args[1])
        if dimension is None or 1 <= dimension <= len(decl.dims):
            continue
        out.append(
            SubscriptHit(
                span=Span(span.start + args[1][0].start, span.start + args[1][-1].end),
                message=(
                    f"{toks[i].raw_text} asks for dimension {dimension} of '{decl.name}', which has "
                    f"{pluralize_count(len(decl.dims), 'dimension')}. {_SUBSCRIPT_ERROR}"
                ),
            )
        )
    return out


def _inline_split_index_violations(
    source: str, span: Span, shadowed: Callable[[str], bool], strings: StringValueOf | None = None
) -> list[SubscriptHit]:
    """`Split("abc", ",")(1)`: indexing the result of Split on literals, whose one
    element sits at 0 (issue #120), and of Filter (issue #260), and on a String
    local known to hold its text (issue #559)."""
    toks = statement_tokens_after_leading_label(source, span)
    out: list[SubscriptHit] = []
    for i in range(len(toks) - 1):
        callee = token_text(toks[i])
        # A project procedure named Split takes the call (issue #280).
        if (
            callee not in ("split", "filter")
            or toks[i + 1].raw_text != "("
            or not is_bare_or_vba_qualified_intrinsic_call(toks, i)
            or (_raw_at(toks, i - 1) != "." and shadowed(toks[i].raw_text))
        ):
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0 or _raw_at(toks, close + 1) != "(":
            continue
        index_close = match_paren_from(toks, close + 1)
        if index_close < 0:
            continue
        shape = array_value_shape(toks[i : close + 1], "Split(...)", 0, strings, module_compare(source))
        index_toks = _without_comments(toks[close + 2 : index_close])
        value = _comparable_array_bound_expression_value(index_toks)
        if shape is None or value is None:
            continue
        detail = _subscript_detail(value, shape.dims[0], 0, 1)
        if detail:
            out.append(
                SubscriptHit(
                    span=Span(span.start + index_toks[0].start, span.start + index_toks[-1].end),
                    message=(
                        f"Subscript {value} for the array {'Split' if callee == 'split' else 'Filter'} "
                        f"returns here {detail}. {_SUBSCRIPT_ERROR}"
                    ),
                )
            )
    return out


def check_fixed_array_subscript_bounds(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    project_integer_constants: Mapping[str, str | None] | None = None,
    project_visible_symbols: Sequence[VbaSymbol] | None = None,
    host_model: HostObjectModel | None = None,
) -> None:
    """Rule: a constant subscript proven outside an array's known bounds raises
    Run-time error '9' (oracle-verified `runtime006_*`): a local or module fixed
    array, an array a value or a ReDim gave bounds, the array a Function
    returns, and a module dynamic array nothing ReDims."""
    option_base = module_option_base(mod, activity)
    module_constants = _module_integer_constants(mod, project_integer_constants, activity)
    module_fixed: list[dict[str, FixedArrayBound]] = []
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        if not module_fixed:
            module_fixed.append(_module_fixed_array_declarations(source, mod, activity, option_base))
        _check_fixed_array_subscript_bounds_procedure(
            source, mod, member, symbols, activity, push, option_base, module_constants, module_fixed[0],
            project_visible_symbols, host_model,
        )


def _check_fixed_array_subscript_bounds_procedure(
    source: str,
    mod: ModuleNode,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    option_base: int,
    module_constants: Mapping[str, float | None],
    module_fixed: Mapping[str, FixedArrayBound],
    project_visible_symbols: Sequence[VbaSymbol] | None,
    host_model: HostObjectModel | None,
) -> None:
    # A module's fixed arrays keep their bounds, and its dynamic arrays that
    # nothing ReDims have none (issue #241).
    module_variables: Mapping[str, VbaSymbol] = untouched_module_variables_in(source, symbols, member)
    # Local/parameter names hide every module array using the same scope.
    hidden = _hidden_in(symbols, member) if module_fixed else set()
    declared: dict[str, FixedArrayBound] = {
        lower: shape for lower, shape in module_fixed.items() if lower not in hidden
    }
    declared.update(_local_fixed_array_declarations_for_body(source, member.body, activity, option_base))
    unallocated = {
        lower: variable
        for lower, variable in module_variables.items()
        if variable.is_array and variable.array_bounds is None
    }
    shapes_at = known_array_shapes_at(source, symbols, member, activity, option_base)
    returned = _function_return_shapes(source, mod, activity, option_base)
    hidden_here = _hidden_in(symbols, member)
    parameterless = {
        one.name.lower()
        for one in active_module_members(mod, activity)
        if isinstance(one, ProcedureNode) and one.proc_kind is ProcKind.FUNCTION and len(one.params) == 0
    } - hidden_here
    redimmed = redim_shapes_at(source, symbols, member, activity, option_base)
    # By the ids of the value-shape map and the ReDim map a statement sees, with
    # both kept beside the merged map.
    merged: dict[tuple[int, int], tuple[object, object, Mapping[str, FixedArrayBound]]] = {}

    def fixed_at(stmt: LeafStatementNode) -> Mapping[str, FixedArrayBound]:
        shapes = shapes_at(stmt)
        reshaped = redimmed.get(stmt)
        key = (id(shapes), id(reshaped))
        entry = merged.get(key)
        if entry is not None and entry[0] is shapes and entry[1] is reshaped:
            return entry[2]
        next_fixed = dict(declared)
        for lower, shape in [*shapes.items(), *(reshaped.items() if reshaped is not None else [])]:
            if lower not in declared:
                next_fixed[lower] = shape
        merged[key] = (shapes, reshaped, next_fixed)
        return next_fixed

    redim_targets = _redim_target_names_in_body(source, member.body, activity)
    exclusions: dict[int, tuple[object, AbstractSet[str]]] = {}

    def excluded_at(stmt: LeafStatementNode) -> AbstractSet[str]:
        reshaped = redimmed.get(stmt)
        if reshaped is None:
            return redim_targets
        entry = exclusions.get(id(reshaped))
        if entry is not None and entry[0] is reshaped:
            return entry[1]
        excluded = {lower for lower in redim_targets if lower not in reshaped}
        exclusions[id(reshaped)] = (reshaped, excluded)
        return excluded

    counters = loop_counters_at(source, member.body, activity)
    # A subscript through a Const or a local with one known value (issue #238).
    constants = procedure_integer_constant_lookup(
        member, module_constants, symbols, project_visible_symbols, activity, host_model
    )
    values_at = known_local_literal_values_at(source, member, symbols, activity)
    # `Split(s, ",")(3)` with s a String local known to hold "q,r" (issue #559).
    strings_at = _string_values_at(source, symbols, member, activity)
    # Code that never runs, after `GoTo Done` or in a loop of no pass, raises
    # nothing, whatever state it builds (issue #406).
    unreachable = unreachable_statements_in(source, member, symbols, activity)
    source_names: list[Any] = []

    def shadowed(name: str) -> bool:
        if not source_names:
            source_names.append(source_name_scope_for(symbols, member, project_visible_symbols))
        result: bool = runtime_callable_source_shadowed(name, source_names[0])
        return result

    def entry_lookup(loop: BodyNode) -> IntegerConstantLookup:
        result: IntegerConstantLookup = with_known_locals(constants, values_at(loop))
        return result

    # Headers too: `For i = 1 To a(5)`, `Select Case a(5)` (issue #233).
    def visit(stmt: LeafStatementNode) -> None:
        if _node_in(unreachable, stmt):
            return
        for hit in _inline_split_index_violations(source, stmt.span, shadowed, strings_at(stmt)):
            push("arraySubscriptOutOfBounds", hit.message, hit.span)
        for hit in _unallocated_module_array_uses(source, stmt.span, unallocated) if unallocated else []:
            push("arraySubscriptOutOfBounds", hit.message, hit.span)
        if returned:
            for hit in _returned_array_subscript_violations(
                source, stmt.span, returned, with_known_locals(constants, values_at(stmt)), parameterless
            ):
                push("arraySubscriptOutOfBounds", hit.message, hit.span)
        fixed = fixed_at(stmt)
        stmt_counters = counters.get(stmt)
        if not fixed and stmt_counters is None:
            return
        excluded = excluded_at(stmt)
        for hit in _fixed_array_subscript_violations(
            source, stmt.span, fixed, excluded, stmt_counters, with_known_locals(constants, values_at(stmt)),
            entry_lookup,
        ):
            push(hit.rule or "arraySubscriptOutOfBounds", hit.message, hit.span)
        for hit in _bound_intrinsic_dimension_violations(source, stmt.span, fixed, excluded):
            push("arraySubscriptOutOfBounds", hit.message, hit.span)

    for_each_statement_with_headers(source, member.body, visit, activity)
