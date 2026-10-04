"""Rule family: array declarations, ReDim, subscripts, Erase, and allocation.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/arrays.ts. This module
owns the shared ReDim-target parser and the comparable literal-bound folding that
several array rules reuse; rules are added incrementally (M7).

The bound folding is deliberately literal-only: a dimension bound contributes a
value only when it reduces to a signed sum of integer literals, so variable- and
Const-backed bounds stay quiet (the no-false-positive contract).
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker, inactive_node_skip
from ...constants.integer_constant_expression import parse_vba_integer_literal, safe_integer
from ...flow.procedure_unstructured import procedure_has_unstructured_flow
from ...js_compat import JS_WHITESPACE, js_number_to_string, js_trim
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ForBlockNode,
    LeafStatementNode,
    ModuleNode,
    OptionNode,
    ProcedureNode,
    Span,
    StatementNode,
    VariableDeclNode,
    VariableGroupNode,
    is_leaf_statement,
    iter_body_nodes_in_context,
)
from ...symbols.name_resolution import BareIdentifierContext
from ...symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbol, VbaSymbolKind
from ...types.type_inference import (
    DeclaredValueShape,
    declaration_shape_environment_for,
    declared_shape_for_source_binding,
    procedure_symbol_for,
)
from ...types.type_names import is_known_scalar_type, normalize_type
from ..context import PushFn, statement_tokens
from ..dataflow import (
    DataflowHooks,
    Lattice,
    tracked_locals_named_whole,
    walk_branch_merged_body,
    walk_straight_line_body,
)
from ..walker import (
    ProcedureStatementVisitor,
    absolute_span,
    active_module_members,
    bare_assignment_target,
    for_each_statement,
    for_each_variable_group,
    is_inactive_node,
    pluralize_count,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .shared import is_bare_or_vba_qualified_intrinsic_call


@dataclass(frozen=True, slots=True)
class _RedimBlockedDeclaration:
    name: str
    span: Span
    kind: str  # "scalar" | "fixedArray"


@dataclass(frozen=True, slots=True)
class _RedimDimension:
    span: Span
    key: str | None = None
    lower_key: str | None = None
    lower_value: int | None = None
    upper_value: int | None = None


@dataclass(frozen=True, slots=True)
class _RedimTarget:
    name: str
    span: Span
    preserve: bool
    dimensions: list[_RedimDimension]


# -- shared ReDim-target parsing + literal-bound folding -------------------


def _token_group_span(base: Span, tokens: Sequence[VbaToken]) -> Span:
    return Span(base.start + tokens[0].start, base.start + tokens[-1].end)


def _comparable_array_bound_expression_value(toks: Sequence[VbaToken]) -> int | None:
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


def _comparable_array_bound_expression_key(toks: Sequence[VbaToken]) -> str | None:
    parts: list[str] = []
    for tok in toks:
        word = token_text(tok)
        if tok.kind is TokenKind.INTEGER_LITERAL or tok.raw_text in ("+", "-") or word == "to":
            parts.append(word if word else tok.raw_text.lower())
            continue
        return None
    return "".join(parts) if parts else None


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


def _redim_target_from_group(
    base: Span, group: Sequence[VbaToken], preserve: bool
) -> _RedimTarget | None:
    content = [tok for tok in group if tok.kind is not TokenKind.COMMENT]
    name_tok = content[0] if content else None
    name = token_name(name_tok) if name_tok is not None else None
    if name_tok is None or name is None:
        return None
    # A qualified ReDim target (`ReDim x.arr(...)` / `ReDim x!arr(...)`) resizes a member
    # array, not the base variable. The scalar/fixed-array shape checks apply only to a
    # simple variable, and the member's declared shape is not resolvable here, so skip
    # qualified targets rather than mistake the container for the array being resized.
    if len(content) > 1 and content[1].raw_text in (".", "!"):
        return None
    dimensions: list[_RedimDimension] = []
    if len(content) > 1 and content[1].raw_text == "(":
        close = match_paren_from(content, 1)
        if close > 1:
            for part in split_top_level_token_groups(content, 2, ",", close):
                dim_tokens = [tok for tok in part if tok.kind is not TokenKind.COMMENT]
                if not dim_tokens:
                    continue
                key, lower_key, lower_value, upper_value = _comparable_array_bound_key(dim_tokens)
                dimensions.append(
                    _RedimDimension(
                        span=_token_group_span(base, dim_tokens),
                        key=key,
                        lower_key=lower_key,
                        lower_value=lower_value,
                        upper_value=upper_value,
                    )
                )
    return _RedimTarget(name=name, span=absolute_span(base, name_tok), preserve=preserve, dimensions=dimensions)


def _redim_statement_targets(source: str, span: Span) -> list[_RedimTarget]:
    return _redim_targets_from_tokens(span, statement_tokens_after_leading_label(source, span))


def _redim_targets_from_tokens(span: Span, toks: Sequence[VbaToken]) -> list[_RedimTarget]:
    if not toks or token_text(toks[0]) != "redim":
        return []
    preserve = len(toks) > 1 and token_text(toks[1]) == "preserve"
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
    if not toks or token_text(toks[0]) != "if":
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
        out.extend(_redim_targets_from_tokens(span, list(toks[i + 1 : end])))
    return out


# -- shared blocked-declaration maps (scalar / fixed-size ReDim targets) ----


def _is_variant_like_redim_target_type(as_type: str | None) -> bool:
    return not as_type or normalize_type(as_type) == "variant"


def _redim_blocked_declaration_kind(is_array: bool, as_type: str | None, array_bounds: str | None) -> str | None:
    if not is_array:
        if _is_variant_like_redim_target_type(as_type):
            return None
        return "scalar"
    return "fixedArray" if array_bounds else None


def _add_redim_blocked_declarations(
    group: VariableGroupNode, out: dict[str, _RedimBlockedDeclaration]
) -> None:
    for decl in group.declarations:
        kind = _redim_blocked_declaration_kind(decl.is_array, decl.as_type, decl.array_bounds)
        if kind is None:
            continue
        lower = decl.name.lower()
        if lower not in out:
            out[lower] = _RedimBlockedDeclaration(name=decl.name, span=decl.span, kind=kind)


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


# -- checkRedimImpossibleBounds --------------------------------------------


def check_redim_impossible_bounds(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> ProcedureStatementVisitor:
    module_declarations = _redim_blocked_declarations_for_module(mod, activity)
    option_base = module_option_base(mod, activity)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        local_declarations = _redim_blocked_declarations_for_body(member.body, activity)
        local_names = _declaration_names_for_body(member.body, activity)

        def visitor(stmt: LeafStatementNode) -> None:
            for target in _redim_statement_targets(source, stmt.span):
                lower_name = target.name.lower()
                blocked = local_declarations.get(lower_name)
                if blocked is None and lower_name not in local_names:
                    blocked = module_declarations.get(lower_name)
                if blocked is not None:
                    # A scalar / fixed-size ReDim target is a compile error reported
                    # by invalidRedimTargets; do not also flag the runtime bound.
                    continue
                for index, dimension in enumerate(target.dimensions):
                    if dimension.upper_value is None:
                        continue
                    # `ReDim a(-1)`: the lower bound is Option Base, 0 by default,
                    # and an upper bound below it is the same impossibility as
                    # `ReDim a(5 To 1)` (XLIDE issue #120, measured in Excel 16.0).
                    lower: int | None
                    if dimension.lower_value is not None:
                        lower = dimension.lower_value
                    else:
                        lower = option_base if dimension.lower_key is None else None
                    if lower is None or lower <= dimension.upper_value:
                        continue
                    lower_text = (
                        f"{lower} (Option Base {lower})"
                        if dimension.lower_value is None
                        else str(lower)
                    )
                    push(
                        "redimImpossibleBounds",
                        f"ReDim lower bound {lower_text} is greater than upper bound "
                        f"{dimension.upper_value} for dimension {index + 1} of '{target.name}'; "
                        "this will raise Run-time error '9': Subscript out of range.",
                        dimension.span,
                    )

        return visitor

    return factory


# -- checkInvalidRedimTargets ----------------------------------------------


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


def check_invalid_redim_targets(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    module_declarations = _redim_blocked_declarations_for_module(mod, activity)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        local_declarations = _redim_blocked_declarations_for_body(member.body, activity)
        local_names = _declaration_names_for_body(member.body, activity)
        proc_sym = procedure_symbol_for(symbols, member)

        def visitor(stmt: LeafStatementNode) -> None:
            for target in _redim_statement_targets(source, stmt.span):
                lower = target.name.lower()
                resolved = declared_shape_for_source_binding(
                    symbols, proc_sym, project_visible_symbols, target.name,
                    BareIdentifierContext.ASSIGNMENT_TARGET,
                )
                if resolved.resolved:
                    declaration = _redim_blocked_declaration_for_shape(
                        target.name, target.span, resolved.shape
                    )
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


# -- checkArrayDeclarationBounds -------------------------------------------

# VBA allows at most 60 array dimensions.
_MAX_ARRAY_DIMENSIONS = 60


def check_array_declaration_bounds(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    def inspect_group(group: VariableGroupNode) -> None:
        for decl in group.declarations:
            if not decl.is_array or decl.array_bounds is None or is_inactive_node(activity, decl):
                continue
            _inspect_array_declaration(source, decl, push)

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            inspect_group(member)
        elif isinstance(member, ProcedureNode):
            for_each_variable_group(member.body, inspect_group, activity)


def _inspect_array_declaration(source: str, decl: VariableDeclNode, push: PushFn) -> None:
    toks = statement_tokens(source, decl.span)
    open_index = next((i for i, tok in enumerate(toks) if tok.raw_text == "("), -1)
    if open_index < 0:
        return
    close = match_paren_from(toks, open_index)
    if close < 0:
        return
    dims = [
        [tok for tok in part if tok.kind is not TokenKind.COMMENT]
        for part in split_top_level_token_groups(toks, open_index + 1, ",", close)
    ]
    dims = [dim_tokens for dim_tokens in dims if dim_tokens]
    if len(dims) > _MAX_ARRAY_DIMENSIONS:
        push(
            "tooManyArrayDimensions",
            f"Array '{decl.name}' has {len(dims)} dimensions; "
            f"VBA allows at most {_MAX_ARRAY_DIMENSIONS}.",
            decl.name_span if decl.name_span is not None else decl.span,
        )
    for index, dim_tokens in enumerate(dims):
        _key, _lower_key, lower_value, upper_value = _comparable_array_bound_key(dim_tokens)
        if lower_value is None or upper_value is None or lower_value <= upper_value:
            continue
        push(
            "arrayDeclarationImpossibleBounds",
            f"Array '{decl.name}' lower bound {lower_value} is greater than upper bound "
            f"{upper_value} for dimension {index + 1}; this is not a valid array bound.",
            _token_group_span(decl.span, dim_tokens),
        )


# -- checkFixedArraySubscriptBounds ----------------------------------------


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
    dims: tuple[ArrayDimensionBound, ...]
    # Where the bounds came from, for the message: 'Dim', 'Array(...)', 'Split(...)',
    # 'Range(...).Value'.
    origin: str


@dataclass(frozen=True, slots=True)
class _ForCounter:
    """A For counter in force at a statement: the value its last pass has."""

    last: int | float
    span: Span


# Integers below this are exact as a JavaScript number, so they print as Python
# prints them.
_EXACT_INTEGER_LIMIT = 2**53

# Heads of statements that can reshape an array or Variant named anywhere in them.
_SHAPE_SPOILING_HEADS = frozenset({"redim", "erase", "set", "input", "get", "line"})

# `Option Base 0|1` in an Option directive's text; `\s` and `\b` as JavaScript
# reads them.
_OPTION_BASE_RE = re.compile(r"base[" + JS_WHITESPACE + r"]+([01])\b", re.IGNORECASE | re.ASCII)

# A multi-cell A1-style address literal: `A1:B2`.
_RANGE_ADDRESS_RE = re.compile(r"([A-Za-z]{1,3})([0-9]+):([A-Za-z]{1,3})([0-9]+)")


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name else None


def _as_js_number(value: float) -> int | float:
    """An integral double as an int, the way a JavaScript number prints: no `.0`."""
    if value.is_integer() and abs(value) < _EXACT_INTEGER_LIMIT:
        return int(value)
    return value


def _parse_fixed_array_bounds_for_decl(
    source: str, decl: VariableDeclNode, option_base: int
) -> list[ArrayDimensionBound] | None:
    """Parses the literal bounds of a fixed-size array declaration, one entry per
    dimension. None unless every dimension's upper bound (and any explicit lower
    bound) folds to a literal integer. A dimension with no `To` takes Option Base
    as its lower bound (XLIDE issue #120: `Option Base 1` then `Dim a(3)` refuses
    `a(0)`)."""
    toks = statement_tokens(source, decl.span)
    open_index = next((i for i, tok in enumerate(toks) if tok.raw_text == "("), -1)
    if open_index < 0:
        return None
    close = match_paren_from(toks, open_index)
    if close < 0:
        return None
    dims = [
        [tok for tok in part if tok.kind is not TokenKind.COMMENT]
        for part in split_top_level_token_groups(toks, open_index + 1, ",", close)
    ]
    dims = [dim_tokens for dim_tokens in dims if dim_tokens]
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
            dims = _parse_fixed_array_bounds_for_decl(source, decl, option_base)
            if dims is not None:
                out[lower] = FixedArrayBound(name=decl.name, dims=tuple(dims), origin="Dim")

    for_each_variable_group(body, visit, activity)
    return out


def module_option_base(mod: ModuleNode, activity: ConditionalActivityTracker | None) -> int:
    """The module's `Option Base`, 0 when absent."""
    for member in active_module_members(mod, activity):
        if isinstance(member, OptionNode):
            match = _OPTION_BASE_RE.match(js_trim(member.option_text))
            if match:
                return int(match.group(1))
    return 0


def known_array_shapes(
    source: str,
    body: Sequence[BodyNode],
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    option_base: int,
) -> dict[str, FixedArrayBound]:
    """Dynamic-array and Variant locals whose bounds a value fixes (XLIDE issue
    #120): the local's ONLY assignment is `Array(...)`, `VBA.Array(...)`,
    `Split(literal, literal[, limit])` or `Range("A1:B2").Value`, and nothing else
    touches it (no ReDim, Erase, whole pass to a call, or Set).

     - `Array(a, b)` is based at Option Base; `VBA.Array` ignores Option Base and
       is based at 0 (both measured in Excel 16.0).
     - `Array()` has UBound -1: every index is out of range.
     - `Split` is always 0-based and yields one part per delimiter plus one;
       Split("abc", ",") is one element, Split("a,b", ",") two.
     - A Range's `.Value` over a multi-cell address literal is a 1-based
       two-dimensional array of the address's rows and columns.

    Keyed by lowercased name, in the order the locals are first assigned.
    """
    proc_sym = procedure_symbol_for(symbols, proc)
    candidates: set[str] = set()
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if child.kind is not VbaSymbolKind.LOCAL_VARIABLE or child.visibility is SymbolVisibility.STATIC:
            continue
        type_ = normalize_type(child.as_type)
        if (child.array_bounds is None) if child.is_array else (type_ is None or type_ == "variant"):
            candidates.add(child.name.lower())
    if not candidates:
        return {}
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
                shape = _array_value_shape(value_tokens, target_name, option_base)
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
                closes_slot = (
                    nxt is None
                    or nxt.raw_text in (")", ",", ":")
                    or nxt.kind is TokenKind.COMMENT
                )
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


def _array_value_shape(
    value_tokens: Sequence[VbaToken], name: str, option_base: int
) -> FixedArrayBound | None:
    """The bounds of the array `Array(...)`, `Split(...)` or `Range(...).Value`
    builds, or None."""
    toks = [tok for tok in value_tokens if tok.kind is not TokenKind.COMMENT]
    if not toks:
        return None
    index = 0
    vba_qualified = False
    if token_text(toks[0]) == "vba" and _raw_at(toks, 1) == ".":
        vba_qualified = True
        index = 2
    callee = token_text(_at(toks, index))
    if callee in ("array", "split") and _raw_at(toks, index + 1) == "(":
        close = match_paren_from(toks, index + 1)
        if close != len(toks) - 1:
            return None
        inner = toks[index + 2 : close]
        if callee == "array":
            count = 0 if not inner else len(split_top_level_token_groups(inner, 0, ","))
            lower = 0 if vba_qualified else option_base
            return FixedArrayBound(
                name=name,
                dims=(ArrayDimensionBound(lower=lower, upper=lower + count - 1, explicit_lower=True),),
                origin="VBA.Array(...)" if vba_qualified else "Array(...)",
            )
        args = split_top_level_token_groups(inner, 0, ",")
        if (
            len(args) < 1
            or len(args) > 2
            or len(args[0]) != 1
            or args[0][0].kind is not TokenKind.STRING_LITERAL
        ):
            return None
        if len(args) == 2 and (len(args[1]) != 1 or args[1][0].kind is not TokenKind.STRING_LITERAL):
            return None
        text = args[0][0].raw_text[1:-1].replace('""', '"')
        delimiter = args[1][0].raw_text[1:-1].replace('""', '"') if len(args) == 2 else " "
        if len(delimiter) == 0:
            return None
        parts = 1 if len(text) == 0 else len(text.split(delimiter))
        return FixedArrayBound(
            name=name,
            dims=(ArrayDimensionBound(lower=0, upper=parts - 1, explicit_lower=True),),
            origin="Split(...)",
        )
    # `Range("A1:B2").Value` and `Worksheets(1).Range("A1:B2").Value`.
    last = len(toks) - 1
    if (
        token_text(toks[last]) == "value"
        and _raw_at(toks, last - 1) == "."
        and _raw_at(toks, last - 2) == ")"
    ):
        close = last - 2
        open_index = next(
            (
                i
                for i, tok in enumerate(toks)
                if tok.raw_text == "(" and match_paren_from(toks, i) == close
            ),
            -1,
        )
        if (
            open_index > 0
            and token_text(toks[open_index - 1]) == "range"
            and close == open_index + 2
            and toks[open_index + 1].kind is TokenKind.STRING_LITERAL
        ):
            address = _RANGE_ADDRESS_RE.fullmatch(toks[open_index + 1].raw_text[1:-1])
            if address is not None:
                # Number() of the row digits, in doubles as upstream computes them.
                rows = _as_js_number(abs(float(address.group(4)) - float(address.group(2))) + 1)
                cols = abs(_column_number(address.group(3)) - _column_number(address.group(1))) + 1
                if rows > 1 or cols > 1:
                    return FixedArrayBound(
                        name=name,
                        dims=(
                            ArrayDimensionBound(lower=1, upper=rows, explicit_lower=True),
                            ArrayDimensionBound(lower=1, upper=cols, explicit_lower=True),
                        ),
                        origin="Range(...).Value",
                    )
    return None


def _column_number(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _redim_target_names_in_body(
    source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> set[str]:
    out: set[str] = set()

    def visit(stmt: LeafStatementNode) -> None:
        for target in _redim_statement_targets(source, stmt.span):
            out.add(target.name.lower())

    for_each_statement(body, visit, activity)
    return out


def _subscript_detail(
    value: int | float, dim: ArrayDimensionBound, index: int, dims: int
) -> str | None:
    """Whether `value` is outside `dim`, with the words for the message when it is."""
    if value >= dim.lower and value <= dim.upper:
        return None
    which = f" in dimension {index + 1}" if dims > 1 else ""
    if dim.upper < dim.lower:
        return (
            f"has no element to reach{which}: the array is empty "
            f"(UBound {js_number_to_string(dim.upper)})"
        )
    if value > dim.upper:
        return f"is above the upper bound {js_number_to_string(dim.upper)}{which}"
    lower = js_number_to_string(dim.lower)
    if dim.explicit_lower:
        return f"is below the lower bound {lower}{which}"
    return f"is below the lower bound {lower}{which} (Option Base {lower})"


def _fixed_array_subscript_violations(
    source: str,
    span: Span,
    fixed: Mapping[str, FixedArrayBound],
    excluded: set[str],
    counters: Mapping[str, _ForCounter] | None = None,
) -> list[tuple[Span, str]]:
    """Literal-subscript accesses of a tracked array that fall outside its bounds."""
    toks = statement_tokens_after_leading_label(source, span)
    out: list[tuple[Span, str]] = []
    for i in range(len(toks) - 1):
        if toks[i + 1].raw_text != "(" or (i >= 1 and toks[i - 1].raw_text in (".", "!")):
            continue
        name = token_name(toks[i])
        lower = name.lower() if name else None
        if not name or not lower or lower not in fixed or lower in excluded:
            continue
        close = match_paren_from(toks, i + 1)
        if close <= i + 1:
            continue
        arg_toks = [tok for tok in toks[i + 2 : close] if tok.kind is not TokenKind.COMMENT]
        slots = split_top_level_token_groups(arg_toks, 0, ",")
        decl = fixed[lower]
        if len(slots) != len(decl.dims) or any(len(slot) == 0 for slot in slots):
            continue  # the dimension count is the compiler's business, not this rule's
        # One report per access: the first dimension that is out of range.
        for index, slot in enumerate(slots):
            dim = decl.dims[index]
            value: int | float | None = _comparable_array_bound_expression_value(slot)
            via_counter: _ForCounter | None = None
            if value is None and len(slot) == 1:
                # `a(i)` inside `For i = 0 To 3`: the counter's last pass.
                via_counter = (counters or {}).get(_lower_counter_key(slot[0]))
                value = via_counter.last if via_counter is not None else None
            if value is None:
                continue  # a variable, Const or member chain: not provable
            detail = _subscript_detail(value, dim, index, len(decl.dims))
            if detail is None:
                continue
            origin = "" if decl.origin == "Dim" else f" ({decl.origin})"
            reached = (
                f"Counter '{slot[0].raw_text}' reaches {js_number_to_string(value)} "
                "on its last pass, which"
                if via_counter is not None
                else f"Subscript {js_number_to_string(value)}"
            )
            out.append(
                (
                    Span(span.start + slot[0].start, span.start + slot[-1].end),
                    f"{reached} for array '{decl.name}'{origin} {detail}. "
                    "This will raise Run-time error '9': Subscript out of range.",
                )
            )
            break
    return out


def _lower_counter_key(tok: VbaToken) -> str:
    # tokenName(...)?.toLowerCase() ?? '': an empty name stays empty.
    name = token_name(tok)
    return name.lower() if name is not None else ""


def _bound_intrinsic_dimension_violations(
    source: str, span: Span, fixed: Mapping[str, FixedArrayBound], excluded: set[str]
) -> list[tuple[Span, str]]:
    """`UBound(a, 2)` / `LBound(a, 2)` on an array with fewer dimensions raises 9
    (XLIDE issue #120, measured in Excel 16.0)."""
    toks = statement_tokens_after_leading_label(source, span)
    out: list[tuple[Span, str]] = []
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
        args = split_top_level_token_groups(
            [tok for tok in toks[i + 2 : close] if tok.kind is not TokenKind.COMMENT], 0, ","
        )
        if len(args) != 2 or len(args[0]) != 1:
            continue
        lower = _lower_name(args[0][0])
        decl = fixed.get(lower) if lower else None
        if decl is None or not lower or lower in excluded:
            continue
        dimension = _comparable_array_bound_expression_value(args[1])
        if dimension is None or (dimension >= 1 and dimension <= len(decl.dims)):
            continue
        out.append(
            (
                Span(span.start + args[1][0].start, span.start + args[1][-1].end),
                f"{toks[i].raw_text} asks for dimension {dimension} of '{decl.name}', which has "
                f"{pluralize_count(len(decl.dims), 'dimension')}. "
                "This will raise Run-time error '9': Subscript out of range.",
            )
        )
    return out


def _inline_split_index_violations(source: str, span: Span) -> list[tuple[Span, str]]:
    """`Split("abc", ",")(1)`: indexing the result of Split on literals, whose one
    element sits at 0 (XLIDE issue #120)."""
    toks = statement_tokens_after_leading_label(source, span)
    out: list[tuple[Span, str]] = []
    for i in range(len(toks) - 1):
        if (
            token_text(toks[i]) != "split"
            or toks[i + 1].raw_text != "("
            or not is_bare_or_vba_qualified_intrinsic_call(toks, i)
        ):
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0 or _raw_at(toks, close + 1) != "(":
            continue
        index_close = match_paren_from(toks, close + 1)
        if index_close < 0:
            continue
        shape = _array_value_shape(toks[i : close + 1], "Split(...)", 0)
        index_toks = [
            tok for tok in toks[close + 2 : index_close] if tok.kind is not TokenKind.COMMENT
        ]
        value = _comparable_array_bound_expression_value(index_toks)
        if shape is None or value is None:
            continue
        detail = _subscript_detail(value, shape.dims[0], 0, 1)
        if detail is not None:
            out.append(
                (
                    Span(span.start + index_toks[0].start, span.start + index_toks[-1].end),
                    f"Subscript {value} for the array Split returns here {detail}. "
                    "This will raise Run-time error '9': Subscript out of range.",
                )
            )
    return out


def _for_counter_last_values(
    source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> dict[int, dict[str, _ForCounter]]:
    """The For counters in force at each statement, with the last value each
    reaches: `For i = 0 To 3` (no Step, or a positive literal Step) ends its last
    pass at 3, so `a(i)` inside it indexes 3 on that pass (XLIDE issue #120).

    Keyed by the statement node's id(): the nodes are not hashable, and every one
    outlives the procedure's pass."""
    out: dict[int, dict[str, _ForCounter]] = {}

    def enter(block: BodyNode, counters: dict[str, _ForCounter]) -> dict[str, _ForCounter]:
        if not isinstance(block, ForBlockNode):
            return counters
        inner = dict(counters)
        header = _for_header_literal_range(source, block)
        if header is not None:
            inner[header[0]] = header[1]
        elif block.control_variable:
            inner.pop(block.control_variable.lower(), None)
        return inner

    empty: dict[str, _ForCounter] = {}
    for node, counters in iter_body_nodes_in_context(body, empty, enter, inactive_node_skip(activity)):
        if is_leaf_statement(node) and counters:
            out[id(node)] = counters
    return out


def _for_header_literal_range(source: str, node: ForBlockNode) -> tuple[str, _ForCounter] | None:
    """`For i = <literal> To <literal> [Step <positive literal>]`: the counter and
    the value its last pass has."""
    if node.each or not node.control_variable:
        return None
    header_end = source.find("\n", node.span.start)
    header = Span(
        node.span.start, node.span.end if header_end < 0 else min(header_end, node.span.end)
    )
    toks = statement_tokens_after_leading_label(source, header)
    eq = next((i for i, tok in enumerate(toks) if tok.raw_text == "="), -1)
    to = next((i for i, tok in enumerate(toks) if token_text(tok) == "to"), -1)
    if eq < 0 or to < eq:
        return None
    step = next((i for i, tok in enumerate(toks) if token_text(tok) == "step"), -1)
    start = _comparable_array_bound_expression_value(toks[eq + 1 : to])
    up_to = _comparable_array_bound_expression_value(toks[to + 1 : step if step > 0 else len(toks)])
    step_value = _comparable_array_bound_expression_value(toks[step + 1 :]) if step > 0 else 1
    if start is None or up_to is None or step_value is None or step_value <= 0 or up_to < start:
        return None
    # The last pass runs at the highest from + k*step not above upTo. Computed in
    # doubles, as upstream's Math.floor((upTo - from) / step) * step is.
    quotient = math.floor(float(up_to - start) / float(step_value))
    last = _as_js_number(float(start) + float(quotient) * float(step_value))
    counter_span = node.control_variable_span if node.control_variable_span is not None else header
    return (node.control_variable.lower(), _ForCounter(last=last, span=counter_span))


def check_fixed_array_subscript_bounds(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Rule: a constant subscript proven outside a LOCAL fixed-size array's declared
    bounds raises Run-time error '9' (oracle-verified `runtime006_*`). No-FP scope:
    only local, single-dimension fixed arrays with a literal upper bound, accessed
    with a literal (or folded signed-integer) subscript, are checked. Dynamic /
    ReDim'd arrays, variable/Const subscripts, multi-dimension arrays, parameters,
    and the Option-Base-dependent lower region of single-bound `Dim a(n)` decls
    stay quiet. Flags subscripts above the upper bound, below an explicit literal
    lower bound, or negative."""
    option_base = module_option_base(mod, activity)
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_fixed_array_subscript_bounds_procedure(
                source, member, symbols, activity, option_base, push
            )


def _check_fixed_array_subscript_bounds_procedure(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    option_base: int,
    push: PushFn,
) -> None:
    fixed = _local_fixed_array_declarations_for_body(source, proc.body, activity, option_base)
    for lower, shape in known_array_shapes(
        source, proc.body, symbols, proc, activity, option_base
    ).items():
        if lower not in fixed:
            fixed[lower] = shape
    excluded = _redim_target_names_in_body(source, proc.body, activity)
    counters = _for_counter_last_values(source, proc.body, activity)

    def visit(stmt: LeafStatementNode) -> None:
        for span, message in _inline_split_index_violations(source, stmt.span):
            push("arraySubscriptOutOfBounds", message, span)
        if not fixed:
            return
        for span, message in _fixed_array_subscript_violations(
            source, stmt.span, fixed, excluded, counters.get(id(stmt))
        ):
            push("arraySubscriptOutOfBounds", message, span)
        for span, message in _bound_intrinsic_dimension_violations(
            source, stmt.span, fixed, excluded
        ):
            push("arraySubscriptOutOfBounds", message, span)

    for_each_statement(proc.body, visit, activity)


# -- checkRedimPreserveDimensions ------------------------------------------


def check_redim_preserve_dimensions(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_redim_preserve_dimensions_in_body(source, member.body, {}, activity, push)


def _check_redim_preserve_dimensions_in_body(
    source: str,
    body: Sequence[BodyNode],
    initial_shapes: Mapping[str, _RedimTarget],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    # Copy-down, no-leak-up: a shape learned inside a nested block is visible
    # deeper in that block but does not propagate back to the enclosing body. A
    # block's body starts from a copy of the shapes learned before the block.
    for node, shapes in iter_body_nodes_in_context(
        body, dict(initial_shapes), lambda _block, outer: dict(outer), inactive_node_skip(activity)
    ):
        if isinstance(node, StatementNode):
            for target in _redim_statement_targets(source, node.span):
                if target.preserve:
                    previous = shapes.get(target.name.lower())
                    reason = (
                        _redim_preserve_dimension_mismatch(previous, target)
                        if previous is not None
                        else None
                    )
                    if reason is not None:
                        push(
                            "redimPreserveDimensionChange",
                            f"ReDim Preserve can only resize the last dimension of "
                            f"'{target.name}'. {reason}",
                            target.span,
                        )
                if target.dimensions:
                    shapes[target.name.lower()] = target


def _redim_preserve_dimension_mismatch(
    previous: _RedimTarget, current: _RedimTarget
) -> str | None:
    prev_len = len(previous.dimensions)
    cur_len = len(current.dimensions)
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


# -- checkEraseTargets -----------------------------------------------------

_ERASE_EXPRESSION_OPERATORS = frozenset(
    {"+", "-", "*", "/", "\\", "^", "&", "=", "<", ">", "<=", ">=", "<>"}
)


def check_erase_targets(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        shapes = declaration_shape_environment_for(symbols, member)
        proc_sym = procedure_symbol_for(symbols, member)

        def visitor(stmt: LeafStatementNode) -> None:
            for span in _invalid_erase_targets(source, stmt.span):
                push(
                    "invalidEraseTarget",
                    "Erase target must be a variable or array name, not an arbitrary expression.",
                    span,
                )
            for name, span, as_type in _erase_scalar_targets(
                source, stmt.span, shapes, symbols, proc_sym, project_visible_symbols
            ):
                push(
                    "eraseRequiresArray",
                    f"Erase target '{name}' must be an array or Variant, "
                    f"but it is declared As {as_type}.",
                    span,
                )

        return visitor

    return factory


def _invalid_erase_targets(source: str, span: Span) -> list[Span]:
    toks = statement_tokens_after_leading_label(source, span)
    if not toks or token_text(toks[0]) != "erase":
        return []
    out: list[Span] = []
    for group in split_top_level_token_groups(toks, 1, ","):
        content = [tok for tok in group if tok.kind is not TokenKind.COMMENT]
        if not content:
            continue
        if _erase_target_looks_variable_like(content):
            continue
        out.append(_token_group_span(span, content))
    return out


def _erase_target_looks_variable_like(toks: Sequence[VbaToken]) -> bool:
    if token_name(toks[0]) is None:
        return False
    if any(tok.raw_text in _ERASE_EXPRESSION_OPERATORS for tok in toks):
        return False
    return toks[0].raw_text != "("


def _erase_scalar_targets(
    source: str,
    span: Span,
    shapes: dict[str, DeclaredValueShape],
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
) -> list[tuple[str, Span, str]]:
    toks = statement_tokens_after_leading_label(source, span)
    if not toks or token_text(toks[0]) != "erase":
        return []
    out: list[tuple[str, Span, str]] = []
    for group in split_top_level_token_groups(toks, 1, ","):
        content = [tok for tok in group if tok.kind is not TokenKind.COMMENT]
        if len(content) != 1:
            continue
        name = token_name(content[0])
        if name is None:
            continue
        resolved = declared_shape_for_source_binding(
            symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.ASSIGNMENT_TARGET
        )
        shape = resolved.shape if resolved.resolved else shapes.get(name.lower())
        if shape is None:
            continue
        as_type = shape.as_type
        if shape.is_array or not as_type:
            continue
        normalized = normalize_type(as_type)
        if not normalized or normalized == "variant":
            continue
        if normalized == "object" or is_known_scalar_type(normalized):
            out.append((name, _token_group_span(span, content), as_type))
    return out


# -- checkUnallocatedDynamicArrayAccess ------------------------------------


def check_unallocated_dynamic_array_access(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_unallocated_dynamic_array_access_procedure(source, member, activity, push)


def _check_unallocated_dynamic_array_access_procedure(
    source: str, proc: ProcedureNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    arrays = _local_dynamic_array_declarations_for_body(proc.body, activity)
    if not arrays:
        return
    state: dict[str, str] = dict.fromkeys(arrays, "unallocated")

    def on_statement(stmt: LeafStatementNode) -> None:
        _check_unallocated_statement(source, stmt, arrays, state, push)

    def touches(stmt: LeafStatementNode) -> set[str]:
        return _dynamic_array_touches_in_statement(source, stmt, arrays)

    def demote(lower: str) -> None:
        state[lower] = "unknown"

    def restore(snapshot: Mapping[str, str]) -> None:
        state.clear()
        state.update(snapshot)

    hooks = DataflowHooks(
        on_statement=on_statement,
        touches_in_statement=touches,
        demote_to_unknown=demote,
        snapshot_state=lambda: dict(state),
        restore_state=restore,
        set_state=lambda key, value: state.__setitem__(key, value),
        lattice=Lattice(init="unallocated", good="allocated", unknown="unknown"),
    )
    walk = (
        walk_straight_line_body
        if procedure_has_unstructured_flow(source, proc, activity)
        else walk_branch_merged_body
    )
    walk(proc.body, lambda node: is_inactive_node(activity, node), hooks)


def _check_unallocated_statement(
    source: str, stmt: LeafStatementNode, arrays: set[str], state: dict[str, str], push: PushFn
) -> None:
    redimmed = _redim_statement_targets(source, stmt.span)
    if redimmed:
        for target in redimmed:
            if target.name.lower() in arrays and target.dimensions:
                state[target.name.lower()] = "allocated"
        return
    erased = _erase_statement_simple_targets(source, stmt.span)
    if erased:
        for lower in erased:
            if lower in arrays:
                state[lower] = "unallocated"
        return
    conditional_redims = _single_line_if_redim_targets(source, stmt.span)
    passed_whole = tracked_locals_named_whole(
        statement_tokens_after_leading_label(source, stmt.span),
        stmt.span.start,
        lambda name: name in arrays,
        _ARRAY_READ_ONLY_INTRINSICS,
    )

    def follows_pass(name: str, hit_span: Span) -> bool:
        # An access that follows a whole-array pass in the same statement, as in
        # `If Load(a) Then Debug.Print a(0)`, runs after the callee had its chance
        # to allocate. One that precedes it, as in `Load(a(0))`, does not.
        pass_at = passed_whole.get(name.lower())
        return pass_at is not None and hit_span.start > pass_at

    for name, hit_span in _unallocated_index_accesses(source, stmt.span, arrays, state):
        # The "access" may be the target of a ReDim embedded in a single-line
        # If...Then - the allocation itself, not a read. Suppress exactly that
        # target's name-token span; bounds expressions still report.
        if any(t.span == hit_span for t in conditional_redims):
            continue
        if follows_pass(name, hit_span):
            continue
        push(
            "unallocatedDynamicArrayAccess",
            f"Dynamic array '{name}' is not allocated before indexed access. "
            "This will raise Run-time error '9': Subscript out of range.",
            hit_span,
        )
    for function_name, name, hit_span in _unallocated_bound_calls(source, stmt.span, arrays, state):
        if follows_pass(name, hit_span):
            continue
        push(
            "unallocatedDynamicArrayAccess",
            f"Dynamic array '{name}' is not allocated before {function_name}. "
            "This will raise Run-time error '9': Subscript out of range.",
            hit_span,
        )
    assignment = bare_assignment_target(source, stmt.span)
    if assignment is not None and assignment[0].lower() in arrays:
        state[assignment[0].lower()] = "unknown"
    for lower in passed_whole:
        if state.get(lower) == "unallocated":
            state[lower] = "unknown"
    # A conditional (single-line If) ReDim allocates only on one path, so move
    # the array to 'unknown' - mirroring how block-If allocations degrade - not
    # 'allocated'. The 'unallocated' guard keeps an already-allocated array precise.
    for target in conditional_redims:
        lower = target.name.lower()
        if lower in arrays and target.dimensions and state.get(lower) == "unallocated":
            state[lower] = "unknown"


def _local_dynamic_array_declarations_for_body(
    body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> set[str]:
    out: set[str] = set()

    def visit(group: VariableGroupNode) -> None:
        if group.is_const or group.modifier.lower() == "static":
            return
        for decl in group.declarations:
            if decl.is_array and not decl.array_bounds:
                out.add(decl.name.lower())

    for_each_variable_group(body, visit, activity)
    return out


def _erase_statement_simple_targets(source: str, span: Span) -> set[str]:
    toks = statement_tokens_after_leading_label(source, span)
    if not toks or token_text(toks[0]) != "erase":
        return set()
    out: set[str] = set()
    for group in split_top_level_token_groups(toks, 1, ","):
        content = [tok for tok in group if tok.kind is not TokenKind.COMMENT]
        if len(content) != 1:
            continue
        name = token_name(content[0])
        if name is not None:
            out.add(name.lower())
    return out


def _unallocated_index_accesses(
    source: str, span: Span, arrays: set[str], state: Mapping[str, str]
) -> list[tuple[str, Span]]:
    toks = statement_tokens_after_leading_label(source, span)
    out: list[tuple[str, Span]] = []
    for i in range(len(toks) - 1):
        if toks[i + 1].raw_text != "(" or (i >= 1 and toks[i - 1].raw_text in (".", "!")):
            continue
        name = token_name(toks[i])
        if name is None:
            continue
        lower = name.lower()
        if lower not in arrays or state.get(lower) != "unallocated":
            continue
        if match_paren_from(toks, i + 1) <= i + 1:
            continue
        out.append((name, Span(span.start + toks[i].start, span.start + toks[i].end)))
    return out


def _unallocated_bound_calls(
    source: str, span: Span, arrays: set[str], state: Mapping[str, str]
) -> list[tuple[str, str, Span]]:
    toks = statement_tokens(source, span)
    out: list[tuple[str, str, Span]] = []
    for i in range(len(toks) - 2):
        function_name = token_name(toks[i])
        if function_name is None or function_name.lower() not in ("lbound", "ubound"):
            continue
        if toks[i + 1].raw_text != "(" or not is_bare_or_vba_qualified_intrinsic_call(toks, i):
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0:
            continue
        slots = split_top_level_token_groups(toks, i + 2, ",", close)
        first_slot = slots[0] if slots else []
        if len(first_slot) != 1:
            continue
        name = token_name(first_slot[0])
        if name is None:
            continue
        lower = name.lower()
        if lower not in arrays or state.get(lower) != "unallocated":
            continue
        out.append(
            (function_name, name, Span(span.start + first_slot[0].start, span.start + first_slot[0].end))
        )
    return out


def _dynamic_array_touches_in_statement(
    source: str, stmt: LeafStatementNode, arrays: set[str]
) -> set[str]:
    out: set[str] = set()
    for target in (
        *_redim_statement_targets(source, stmt.span),
        *_single_line_if_redim_targets(source, stmt.span),
    ):
        if target.name.lower() in arrays:
            out.add(target.name.lower())
    for lower in _erase_statement_simple_targets(source, stmt.span):
        if lower in arrays:
            out.add(lower)
    assignment = bare_assignment_target(source, stmt.span)
    if assignment is not None and assignment[0].lower() in arrays:
        out.add(assignment[0].lower())
    out.update(
        tracked_locals_named_whole(
            statement_tokens_after_leading_label(source, stmt.span),
            stmt.span.start,
            lambda name: name in arrays,
            _ARRAY_READ_ONLY_INTRINSICS,
        )
    )
    return out


# Intrinsics that read an array argument and allocate nothing.
_ARRAY_READ_ONLY_INTRINSICS: frozenset[str] = frozenset({"lbound", "ubound", "isarray"})


# -- checkArrayBoundIntrinsicArguments -------------------------------------


def check_array_bound_intrinsic_arguments(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        shapes = declaration_shape_environment_for(symbols, member)
        proc_sym = procedure_symbol_for(symbols, member)

        def visitor(stmt: LeafStatementNode) -> None:
            for function_name, name, hit_span, as_type in _array_bound_scalar_arguments(
                source, stmt.span, shapes, symbols, proc_sym, project_visible_symbols
            ):
                push(
                    "arrayBoundRequiresArray",
                    f"{function_name} requires an array argument, "
                    f"but '{name}' is declared As {as_type}.",
                    hit_span,
                )

        return visitor

    return factory


def _array_bound_scalar_arguments(
    source: str,
    span: Span,
    shapes: dict[str, DeclaredValueShape],
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
) -> list[tuple[str, str, Span, str]]:
    toks = statement_tokens(source, span)
    hits: list[tuple[str, str, Span, str]] = []
    for i in range(len(toks) - 2):
        function_name = token_name(toks[i])
        if function_name is None or function_name.lower() not in ("lbound", "ubound"):
            continue
        if toks[i + 1].raw_text != "(" or not is_bare_or_vba_qualified_intrinsic_call(toks, i):
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0 or close <= i + 2:
            continue
        slots = split_top_level_token_groups(toks, i + 2, ",", close)
        first_slot = slots[0] if slots else []
        if len(first_slot) != 1:
            continue
        arg_name = token_name(first_slot[0])
        if arg_name is None:
            continue
        resolved = declared_shape_for_source_binding(
            symbols, proc_sym, project_visible_symbols, arg_name, BareIdentifierContext.EXPRESSION
        )
        shape = resolved.shape if resolved.resolved else shapes.get(arg_name.lower())
        if shape is None:
            continue
        as_type = shape.as_type
        if shape.is_array or not as_type:
            continue
        normalized = normalize_type(as_type)
        if not normalized or not is_known_scalar_type(normalized):
            continue
        hits.append(
            (
                function_name,
                arg_name,
                Span(span.start + first_slot[0].start, span.start + first_slot[0].end),
                as_type,
            )
        )
    return hits


# --- sync stubs (2f49b93): replaced as each group is ported ---


def check_redim_type_change(*args: object, **kwargs: object) -> None:
    return None


StringValueOf = object


class SubscriptHit:
    pass


def parse_fixed_array_bounds_for_decl(*args: object, **kwargs: object) -> None:
    return None


def literal_dimensions(*args: object, **kwargs: object) -> None:
    return None


def local_fixed_arrays(*args: object, **kwargs: object) -> None:
    return None


def known_array_shapes_at(*args: object, **kwargs: object) -> None:
    return None


def redim_shapes_at(*args: object, **kwargs: object) -> None:
    return None


def array_value_shape(*args: object, **kwargs: object) -> None:
    return None


def range_value_block(*args: object, **kwargs: object) -> None:
    return None


def single_cell_value(*args: object, **kwargs: object) -> None:
    return None


def elements_written_in(*args: object, **kwargs: object) -> None:
    return None


class ElementOperand:
    pass


def element_operand_ending_at(*args: object, **kwargs: object) -> None:
    return None


def element_operand_starting_at(*args: object, **kwargs: object) -> None:
    return None


def dimension_count_violation(*args: object, **kwargs: object) -> None:
    return None


def shape_subscript_violation(*args: object, **kwargs: object) -> None:
    return None


def subscript_violation(*args: object, **kwargs: object) -> None:
    return None
