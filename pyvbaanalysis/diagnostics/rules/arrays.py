"""Rule family: array declarations, ReDim, subscripts, Erase, and allocation.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/arrays.ts. This module
owns the shared ReDim-target parser and the comparable literal-bound folding that
several array rules reuse; rules are added incrementally (M7).

The bound folding is deliberately literal-only: a dimension bound contributes a
value only when it reduces to a signed sum of integer literals, so variable- and
Const-backed bounds stay quiet (the no-false-positive contract).
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import parse_vba_integer_literal, safe_integer
from ...flow.procedure_unstructured import procedure_has_unstructured_flow
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    VariableDeclNode,
    VariableGroupNode,
)
from ...symbols.name_resolution import BareIdentifierContext
from ...symbols.symbol_model import ModuleSymbols, VbaSymbol
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
    tracked_locals_passed_as_call_arguments,
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
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)


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
    toks = statement_tokens_after_leading_label(source, span)
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

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        local_declarations = _redim_blocked_declarations_for_body(member.body, activity)
        local_names = _declaration_names_for_body(member.body, activity)

        def visitor(stmt: LeafStatementNode) -> None:
            for target in _redim_statement_targets(source, stmt.span):
                lower = target.name.lower()
                blocked = local_declarations.get(lower)
                if blocked is None and lower not in local_names:
                    blocked = module_declarations.get(lower)
                if blocked is not None:
                    # A scalar / fixed-size ReDim target is a compile error reported
                    # by invalidRedimTargets; do not also flag the runtime bound.
                    continue
                for index, dimension in enumerate(target.dimensions):
                    if (
                        dimension.lower_value is None
                        or dimension.upper_value is None
                        or dimension.lower_value <= dimension.upper_value
                    ):
                        continue
                    push(
                        "redimImpossibleBounds",
                        f"ReDim lower bound {dimension.lower_value} is greater than upper bound "
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
class _FixedArrayBound:
    name: str
    upper_value: int
    has_explicit_lower: bool
    lower_value: int | None = None


def _parse_fixed_array_bounds_for_decl(
    source: str, decl: VariableDeclNode
) -> tuple[int | None, int, bool] | None:
    """Returns (lower_value, upper_value, has_explicit_lower) for a single-dim fixed array."""
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
    if len(dims) != 1:
        return None  # multi-dimension subscript matching is out of scope
    _key, _lower_key, lower_value, upper_value = _comparable_array_bound_key(dims[0])
    if upper_value is None:
        return None  # non-literal upper bound is not statically known
    return (lower_value, upper_value, lower_value is not None)


def _local_fixed_array_declarations_for_body(
    source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> dict[str, _FixedArrayBound]:
    out: dict[str, _FixedArrayBound] = {}

    def visit(group: VariableGroupNode) -> None:
        if group.is_const:
            return
        for decl in group.declarations:
            if not decl.is_array or not decl.array_bounds:
                continue  # dynamic arrays (no static bounds) are out of scope
            lower = decl.name.lower()
            if lower in out:
                continue
            bounds = _parse_fixed_array_bounds_for_decl(source, decl)
            if bounds is not None:
                lower_value, upper_value, has_explicit_lower = bounds
                out[lower] = _FixedArrayBound(
                    name=decl.name,
                    upper_value=upper_value,
                    has_explicit_lower=has_explicit_lower,
                    lower_value=lower_value,
                )

    for_each_variable_group(body, visit, activity)
    return out


def _redim_target_names_in_body(
    source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> set[str]:
    out: set[str] = set()

    def visit(stmt: LeafStatementNode) -> None:
        for target in _redim_statement_targets(source, stmt.span):
            out.add(target.name.lower())

    for_each_statement(body, visit, activity)
    return out


def _fixed_array_subscript_violations(
    source: str, span: Span, fixed: dict[str, _FixedArrayBound], excluded: set[str]
) -> list[tuple[Span, str]]:
    toks = statement_tokens_after_leading_label(source, span)
    out: list[tuple[Span, str]] = []
    for i in range(len(toks) - 1):
        if toks[i + 1].raw_text != "(" or (i >= 1 and toks[i - 1].raw_text in (".", "!")):
            continue
        name = token_name(toks[i])
        lower = name.lower() if name is not None else None
        if name is None or lower is None or lower not in fixed or lower in excluded:
            continue
        close = match_paren_from(toks, i + 1)
        if close <= i + 1:
            continue
        arg_toks = [tok for tok in toks[i + 2 : close] if tok.kind is not TokenKind.COMMENT]
        slots = split_top_level_token_groups(arg_toks, 0, ",")
        if len(slots) != 1 or len(slots[0]) == 0:
            continue  # not a single-subscript index access (multi-dim/empty)
        value = _comparable_array_bound_expression_value(slots[0])
        if value is None:
            continue  # non-literal subscript -> not statically provable
        decl = fixed[lower]
        low_gate = decl.lower_value if (decl.has_explicit_lower and decl.lower_value is not None) else 0
        if value <= decl.upper_value and value >= low_gate:
            continue
        slot = slots[0]
        if value > decl.upper_value:
            detail = f"is above the array's declared upper bound {decl.upper_value}"
        elif decl.has_explicit_lower:
            detail = f"is below the array's declared lower bound {decl.lower_value}"
        else:
            detail = "is negative and out of range"
        out.append(
            (
                Span(span.start + slot[0].start, span.start + slot[-1].end),
                f"Subscript {value} for array '{decl.name}' {detail}. "
                "This will raise Run-time error '9': Subscript out of range.",
            )
        )
    return out


def check_fixed_array_subscript_bounds(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_fixed_array_subscript_bounds_procedure(source, member, activity, push)


def _check_fixed_array_subscript_bounds_procedure(
    source: str, proc: ProcedureNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    fixed = _local_fixed_array_declarations_for_body(source, proc.body, activity)
    if not fixed:
        return
    excluded = _redim_target_names_in_body(source, proc.body, activity)

    def visit(stmt: LeafStatementNode) -> None:
        for span, message in _fixed_array_subscript_violations(source, stmt.span, fixed, excluded):
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
    # deeper in that block but does not propagate back to the enclosing body.
    shapes = dict(initial_shapes)
    for node in body:
        if is_inactive_node(activity, node):
            continue
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
            continue
        child = getattr(node, "body", None)
        if isinstance(child, list):
            _check_redim_preserve_dimensions_in_body(source, child, shapes, activity, push)


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
    for name, hit_span in _unallocated_index_accesses(source, stmt.span, arrays, state):
        push(
            "unallocatedDynamicArrayAccess",
            f"Dynamic array '{name}' is not allocated before indexed access. "
            "This will raise Run-time error '9': Subscript out of range.",
            hit_span,
        )
    for function_name, name, hit_span in _unallocated_bound_calls(source, stmt.span, arrays, state):
        push(
            "unallocatedDynamicArrayAccess",
            f"Dynamic array '{name}' is not allocated before {function_name}. "
            "This will raise Run-time error '9': Subscript out of range.",
            hit_span,
        )
    assignment = bare_assignment_target(source, stmt.span)
    if assignment is not None and assignment[0].lower() in arrays:
        state[assignment[0].lower()] = "unknown"
    toks = statement_tokens_after_leading_label(source, stmt.span)
    for lower in tracked_locals_passed_as_call_arguments(toks, lambda name: name in arrays):
        if state.get(lower) == "unallocated":
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
        if toks[i + 1].raw_text != "(" or not _is_bare_or_vba_qualified_intrinsic_call(toks, i):
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


def _is_bare_or_vba_qualified_intrinsic_call(toks: Sequence[VbaToken], name_index: int) -> bool:
    if name_index < 1 or toks[name_index - 1].raw_text != ".":
        return True
    qualifier = token_name(toks[name_index - 2]) if name_index >= 2 else None
    return (
        qualifier is not None
        and qualifier.lower() == "vba"
        and (name_index < 3 or toks[name_index - 3].raw_text != ".")
    )


def _dynamic_array_touches_in_statement(
    source: str, stmt: LeafStatementNode, arrays: set[str]
) -> set[str]:
    out: set[str] = set()
    for target in _redim_statement_targets(source, stmt.span):
        if target.name.lower() in arrays:
            out.add(target.name.lower())
    for lower in _erase_statement_simple_targets(source, stmt.span):
        if lower in arrays:
            out.add(lower)
    assignment = bare_assignment_target(source, stmt.span)
    if assignment is not None and assignment[0].lower() in arrays:
        out.add(assignment[0].lower())
    toks = statement_tokens_after_leading_label(source, stmt.span)
    out.update(tracked_locals_passed_as_call_arguments(toks, lambda name: name in arrays))
    return out
