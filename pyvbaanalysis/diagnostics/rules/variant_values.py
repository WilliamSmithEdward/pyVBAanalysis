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

The values come from the same analysis the division and subscript rules use: a
local every assignment gives the same literal, or an array from Array(), Split()
on literals or a Range literal's Value, with nothing else able to change it.
"""

from __future__ import annotations

from collections.abc import Sequence

from ...conditional import ConditionalActivityTracker
from ...js_compat import js_number_to_string
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...symbols.symbol_model import ModuleSymbols
from ...types.type_inference import type_environment_for
from ...types.type_names import normalize_type
from ..context import PushFn, statement_tokens
from ..known_locals import known_local_literal_values
from ..walker import (
    active_module_members,
    bare_assignment_target,
    for_each_statement,
    statement_and_branch_spans,
    token_name,
    token_text,
)
from .arrays import known_array_shapes, module_option_base

_SCALAR_OPERATORS = frozenset({"=", "<", ">", "<=", ">=", "<>", "+", "-", "*", "/", "\\", "&", "^"})


def check_variant_value_misuse(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    option_base = module_option_base(mod, activity)
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_procedure(source, member, symbols, activity, option_base, push)


def _check_procedure(
    source: str,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    option_base: int,
    push: PushFn,
) -> None:
    env = type_environment_for(symbols, member)

    def is_variant(lower: str) -> bool:
        type_ = normalize_type(env.get(lower))
        return type_ is None or type_ == "variant"

    scalars: dict[str, str] = {}
    for lower, value in known_local_literal_values(source, member, symbols, activity).items():
        if value.origin == "literal" and is_variant(lower):
            scalars[lower] = (
                f'the string "{value.value}"'
                if value.kind == "string"
                else f"the number {_number_text(value.value)}"
            )
    arrays: dict[str, str] = {}
    for lower, shape in known_array_shapes(
        source, member.body, symbols, member, activity, option_base
    ).items():
        if is_variant(lower):
            arrays[lower] = shape.origin
    if not scalars and not arrays:
        return

    def visit(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            _check_span(source, span, scalars, arrays, push)

    for_each_statement(member.body, visit, activity)


def _check_span(
    source: str,
    span: Span,
    scalars: dict[str, str],
    arrays: dict[str, str],
    push: PushFn,
) -> None:
    toks = statement_tokens(source, span)
    target = bare_assignment_target(source, span)
    target_index = (
        next((i for i, tok in enumerate(toks) if tok.raw_text == "="), -1) - 1
        if target is not None
        else -1
    )
    for i, tok in enumerate(toks):
        if i == target_index or _raw_at(toks, i - 1) == ".":
            continue
        name = token_name(tok)
        lower = name.lower() if name else None
        if not lower:
            continue
        at = Span(span.start + tok.start, span.start + tok.end)
        scalar = scalars.get(lower)
        array = arrays.get(lower)
        if not scalar and not array:
            continue
        nxt = _at(toks, i + 1)
        if nxt is not None and nxt.raw_text == "." and token_name(_at(toks, i + 2)):
            holds = scalar if scalar is not None else f"an array from {array}"
            push(
                "variantValueMisuse",
                f"'{tok.raw_text}' holds {holds} here, which has no members. "
                "This will raise Run-time error '424': Object required.",
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


def _is_bound_argument(toks: Sequence[VbaToken], i: int) -> bool:
    """True when `toks[i]` is the whole first argument of UBound or LBound."""
    name = token_text(_at(toks, i - 2))
    return (
        _raw_at(toks, i - 1) == "("
        and (name == "ubound" or name == "lbound")
        and _raw_at(toks, i - 3) != "."
        and (_raw_at(toks, i + 1) == ")" or _raw_at(toks, i + 1) == ",")
    )


def _number_text(value: int | float | str) -> str:
    # A number-kind value is always numeric; the str arm only satisfies the type.
    return value if isinstance(value, str) else js_number_to_string(value)


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None
