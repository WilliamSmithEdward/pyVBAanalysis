"""Rule family: array declarations, ReDim, subscripts, Erase, and allocation.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/arrays.ts. This module
owns the shared ReDim-target parser and the comparable literal-bound folding that
several array rules reuse; rules are added incrementally (M7).

The bound folding is deliberately literal-only: a dimension bound contributes a
value only when it reduces to a signed sum of integer literals, so variable- and
Const-backed bounds stay quiet (the no-false-positive contract).
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import parse_vba_integer_literal, safe_integer
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    VariableDeclNode,
    VariableGroupNode,
)
from ...types.type_names import normalize_type
from ..context import PushFn, statement_tokens
from ..walker import (
    ProcedureStatementVisitor,
    absolute_span,
    active_module_members,
    for_each_variable_group,
    is_inactive_node,
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
