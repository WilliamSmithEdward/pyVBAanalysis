"""Rule family: optional-parameter default values.

Ported from checkParameterDefaultValues (declarations.ts): an Optional parameter's
default value must be type-compatible with its declared type (a string literal for
a numeric/Boolean parameter is a VBE compile error; an object parameter's default
must be Nothing; an array parameter's default cannot be scalar). Object typing uses
the host-free is_known_object_assignment_type (generic Object); host-class object
parameters resolve to None (M9), which is precision-only.
"""

from __future__ import annotations

from ...conditional import ConditionalActivityTracker
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize
from ...parser.nodes import ModuleNode, ParameterNode, ProcedureNode, Span
from ...types.type_names import is_known_object_assignment_type, is_known_scalar_type, normalize_type
from ..argument_inference import incompatibility_reason, infer_argument_type
from ..call_extraction import InferredArgumentType
from ..context import PushFn
from ..walker import active_module_members, span_for_tokens, top_level_operator_index


def check_parameter_default_values(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        for param in member.params:
            if not param.default_raw or not param.as_type:
                continue
            default = _value_tokens_after_equals(source, param.span)
            if default is None:
                continue
            tokens, span = default
            actual = infer_argument_type(tokens, param.span.start, {}, {})
            if actual is None:
                continue
            reason = _parameter_default_incompatibility_reason(param, actual)
            if reason is None:
                continue
            push(
                "parameterDefaultTypeMismatch",
                f"Optional parameter '{param.name}' expects {_parameter_default_expected_label(param)}, "
                f"but its default value is {actual.label}. {reason}",
                span,
            )


def _value_tokens_after_equals(source: str, span: Span) -> tuple[list[VbaToken], Span] | None:
    toks = [
        t
        for t in tokenize(source[span.start : span.end])
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]
    eq = top_level_operator_index(toks, "=")
    if eq < 0 or eq + 1 >= len(toks):
        return None
    tokens = toks[eq + 1 :]
    return tokens, span_for_tokens(tokens, span.start)


def _parameter_default_incompatibility_reason(
    param: ParameterNode, actual: InferredArgumentType
) -> str | None:
    if param.is_array and _is_known_scalar_default_type(actual.type_):
        return "Optional array parameter defaults cannot be scalar values."
    expected_raw = param.as_type
    if not expected_raw:
        return None
    if is_known_object_assignment_type(expected_raw):
        if normalize_type(actual.type_) == "nothing":
            return None
        return "Optional object parameter defaults must be Nothing."
    reason = incompatibility_reason(expected_raw, actual)
    if not reason or "string literal" not in actual.label.lower():
        return None
    return "This is a VBE compile error: Type mismatch."


def _parameter_default_expected_label(param: ParameterNode) -> str:
    base = param.as_type if param.as_type else "Variant"
    return f"{base}()" if param.is_array else base


def _is_known_scalar_default_type(type_: str | None) -> bool:
    normalized = normalize_type(type_)
    return normalized is not None and is_known_scalar_type(normalized)
