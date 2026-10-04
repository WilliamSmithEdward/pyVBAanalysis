"""Argument / expression type inference and type-compatibility checking.

Ported from the inference engine of
xlide_vscode/src/analyzer/diagnostics/typeInference.ts: inferExpressionType and
its atomic, arithmetic and concatenation typers, the host-global call type,
member-expression typing through the member-completion context, the
string-arithmetic operand check, incompatibilityReason, the object/value
argument checks, and validateArgumentTypes(ForSignature). The pure helpers
these read (operand splitters, literal and overflow typing, ByRef exactness)
are in types/type_inference.py.

The member-completion paths run only when the caller passes the source and a
member context; without them an expression that needs one resolves to None, and
an unresolved argument type is simply not checked.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, replace

from ..completion.member_access import (
    MemberCompletionContext,
    MemberCompletionEntry,
    is_known_object_assignment_type,
    resolve_exact_member_completion,
)
from ..constants.integer_constant_expression import parse_decimal_integer_literal
from ..host.host_model import resolve_host_global, resolve_host_global_member
from ..js_compat import js_number_to_string
from ..lexer.token_helpers import match_paren_from
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import Span
from ..symbols.symbol_model import qualified_procedure_key
from ..types.type_inference import (
    VALUE_HELD,
    ByRefMismatch,
    SourceDeclaredType,
    arithmetic_of_scalars,
    by_ref_variable_type_mismatch,
    final_member_token_in_expression,
    find_nonnumeric_string_in_arithmetic_expression,
    float_literal_value,
    has_top_level_operator,
    infer_bare_external_constant_expression_type,
    infer_bare_external_object_expression_type,
    infer_intrinsic_cverr_error_variant,
    infer_qualified_external_constant_expression_type,
    infer_signed_numeric_literal,
    js_string_literal,
    member_accepts_zero_arguments,
    member_expression_return_type,
    numeric_literal_overflow_reason,
    object_value_needs_index,
    parameterless_value_signature,
    sheets_from_collection_property,
    split_top_level_arithmetic_operands,
    split_top_level_operands,
)
from ..types.type_names import (
    is_known_scalar_type,
    is_numeric_type,
    is_string_concatenation_operand_type,
    normalize_type,
    numeric_literal_bounds,
)
from .call_extraction import (
    CallableParamType,
    CallableTypeSignature,
    CallArguments,
    InferredArgumentType,
    named_argument_slot,
    split_arg_slots,
    string_literal_value,
    unwrap_outer_parens,
)
from .callable_signatures import (
    SourceNameScope,
    bare_callable_source_shadowed,
    callable_signature_for,
    callable_signature_for_call,
    parenthesized_call_name_at,
    runtime_callable_source_shadowed,
)
from .context import PushFn
from .null_operators import operator_yields_null
from .string_conversion import is_invalid_boolean_string, is_invalid_date_string, numeric_string_verdict
from .walker import span_for_tokens, strip_header_brackets, token_name, token_text

SourceDeclaredTypeResolver = Callable[[str], SourceDeclaredType]
SourceQualifiedDeclaredTypeResolver = Callable[[str, str], SourceDeclaredType]

_SignatureMap = Mapping[str, CallableTypeSignature]


def _significant(toks: Sequence[VbaToken]) -> list[VbaToken]:
    return [t for t in toks if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE]


# -- expression type inference ---------------------------------------------


def infer_argument_type(
    slot: Sequence[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    module_signatures: _SignatureMap,
    source_names: SourceNameScope | None = None,
    resolve_expression_type: SourceDeclaredTypeResolver | None = None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None = None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> InferredArgumentType | None:
    return infer_expression_type(
        _significant(slot),
        slice_start,
        env,
        module_signatures,
        source_names,
        resolve_expression_type,
        resolve_qualified_expression_type,
        source=source,
        member_ctx=member_ctx,
    )


# Runtime functions that return Null for a Null first argument, where the $
# spellings raise 94 at the call (XLIDE issue #184, measured in Excel 16.0).
_NULL_PROPAGATING_RUNTIME_FUNCTIONS = frozenset(
    {"len", "lenb", "left", "right", "mid", "ucase", "lcase", "trim", "ltrim", "rtrim"}
)


def infer_expression_type(
    toks: list[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    module_signatures: _SignatureMap,
    source_names: SourceNameScope | None = None,
    resolve_expression_type: SourceDeclaredTypeResolver | None = None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None = None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> InferredArgumentType | None:
    if not toks:
        return None
    # Upstream recurses once per enclosing pair of parentheses; unwrap them all.
    while True:
        unwrapped = unwrap_outer_parens(toks)
        if len(unwrapped) == len(toks):
            break
        toks = unwrapped
    if not toks:
        return None
    signed = infer_signed_numeric_literal(toks, slice_start)
    if signed is not None:
        return signed
    concatenation = infer_string_concatenation_expression_type(
        toks, slice_start, env, module_signatures, source_names,
        resolve_expression_type, resolve_qualified_expression_type,
        source=source, member_ctx=member_ctx,
    )
    if concatenation is not None:
        return concatenation
    arithmetic = infer_arithmetic_expression_type(
        toks, slice_start, env, module_signatures, source_names,
        resolve_expression_type, resolve_qualified_expression_type,
        source=source, member_ctx=member_ctx,
    )
    if arithmetic is not None:
        return arithmetic
    return infer_atomic_expression_type(
        toks, slice_start, env, module_signatures, source_names,
        resolve_expression_type, resolve_qualified_expression_type,
        source=source, member_ctx=member_ctx,
    )


def infer_atomic_expression_type(
    toks: list[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    module_signatures: _SignatureMap,
    source_names: SourceNameScope | None = None,
    resolve_expression_type: SourceDeclaredTypeResolver | None = None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None = None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> InferredArgumentType | None:
    if not toks:
        return None
    first = toks[0]
    span = Span(slice_start + first.start, slice_start + first.end)
    if len(toks) == 1:
        literal = _infer_atomic_literal(first, span)
        if literal is not None:
            return literal

    name = token_name(first)
    if name and len(toks) == 1:
        declared_type = resolve_expression_type(name) if resolve_expression_type else None
        # A Const that holds a string literal converts as the literal does (XLIDE
        # issue #255).
        if declared_type is not None and declared_type.resolved and declared_type.string_value is not None:
            return InferredArgumentType(
                type_="String",
                label=f"constant '{name}' ({js_string_literal(declared_type.string_value)})",
                span=span,
                string_value=declared_type.string_value,
            )
        type_ = (
            declared_type.as_type
            if declared_type is not None and declared_type.resolved
            else env.get(name.lower())
        )
        if type_:
            return InferredArgumentType(type_=type_, label=f"{name} As {type_}", span=span)
        sig = parameterless_value_signature(name, module_signatures, source_names)
        if sig is not None and sig.return_type:
            return InferredArgumentType(
                type_=sig.return_type, label=f"{name} As {sig.return_type}", span=span
            )
        external_object = infer_bare_external_object_expression_type(name, span, source_names, member_ctx)
        if external_object is not None:
            return external_object
        return infer_bare_external_constant_expression_type(
            name, span, source_names, member_ctx.model if member_ctx is not None else None
        )

    if token_text(first) == "new" and len(toks) == 2:
        type_name = token_name(toks[1])
        if type_name:
            return InferredArgumentType(
                type_=type_name,
                label=f"New {type_name}",
                span=Span(slice_start + toks[1].start, slice_start + toks[1].end),
            )

    if name:
        call_name = parenthesized_call_name_at(toks, 0)
        error_variant = (
            infer_intrinsic_cverr_error_variant(toks, slice_start, module_signatures, source_names)
            if call_name is not None and call_name.paren_index == 1
            else None
        )
        if error_variant is not None:
            return error_variant
        # `Len(Null)` and `UCase(Null)` return Null, whatever type they otherwise
        # return (XLIDE issue #184, measured in Excel 16.0).
        if (
            call_name is not None
            and call_name.name.lower() in _NULL_PROPAGATING_RUNTIME_FUNCTIONS
            and call_name.name.lower() not in module_signatures
            and not bare_callable_source_shadowed(call_name.name, source_names)
            and not runtime_callable_source_shadowed(call_name.name, source_names)
            and match_paren_from(toks, call_name.paren_index) == len(toks) - 1
        ):
            slots = split_arg_slots(toks[call_name.paren_index + 1 : -1], slice_start).slots
            first_slot = slots[0] if slots else []
            if len(first_slot) == 1 and token_text(first_slot[0]) == "null":
                return InferredArgumentType(
                    type_="Null",
                    label=f"{call_name.name}(Null), which is Null",
                    span=Span(span.start, slice_start + toks[-1].end),
                )
        if call_name is not None:
            sig = callable_signature_for(call_name.name, module_signatures, source_names)
            if (
                sig is not None
                and sig.return_type
                and match_paren_from(toks, call_name.paren_index) == len(toks) - 1
            ):
                return InferredArgumentType(
                    type_=sig.return_type,
                    label=f"{call_name.name}(...) As {sig.return_type}",
                    span=Span(span.start, slice_start + toks[call_name.name_end_index].end),
                )
            host_global = _infer_bare_host_global_call_type(
                toks, call_name.name, call_name.paren_index, slice_start, module_signatures,
                source_names, member_ctx,
            )
            if host_global is not None:
                return host_global

    if name and len(toks) > 1 and toks[1].raw_text == ".":
        member = token_name(toks[2]) if len(toks) > 2 else None
        error_variant = infer_intrinsic_cverr_error_variant(
            toks, slice_start, module_signatures, source_names
        )
        if error_variant is not None:
            return error_variant
        if member and len(toks) == 3:
            member_span = Span(slice_start + toks[2].start, slice_start + toks[2].end)
            lookup_key = qualified_procedure_key(name, member)
            sig = parameterless_value_signature(lookup_key, module_signatures)
            if sig is not None and sig.return_type:
                return InferredArgumentType(
                    type_=sig.return_type,
                    label=f"{name}.{member} As {sig.return_type}",
                    span=member_span,
                )
            declared_type = (
                resolve_qualified_expression_type(name, member)
                if resolve_qualified_expression_type
                else None
            )
            if declared_type is not None and declared_type.resolved:
                if declared_type.as_type:
                    return InferredArgumentType(
                        type_=declared_type.as_type,
                        label=f"{name}.{member} As {declared_type.as_type}",
                        span=member_span,
                    )
                return None
            external = infer_qualified_external_constant_expression_type(
                name, member, member_span, member_ctx.model if member_ctx is not None else None
            )
            if external is not None:
                return external
        if (
            member
            and len(toks) > 3
            and toks[3].raw_text == "("
            and match_paren_from(toks, 3) == len(toks) - 1
        ):
            lookup_key = qualified_procedure_key(name, member)
            sig = module_signatures.get(lookup_key)
            if sig is not None and sig.return_type:
                return InferredArgumentType(
                    type_=sig.return_type,
                    label=f"{name}.{member}(...) As {sig.return_type}",
                    span=Span(slice_start + toks[2].start, slice_start + toks[2].end),
                )
    if source is not None and member_ctx is not None:
        return infer_member_expression_type(source, toks, slice_start, member_ctx)
    return None


def _infer_bare_host_global_call_type(
    toks: Sequence[VbaToken],
    call_name: str,
    paren_index: int,
    slice_start: int,
    module_signatures: _SignatureMap,
    source_names: SourceNameScope | None,
    member_ctx: MemberCompletionContext | None,
) -> InferredArgumentType | None:
    """`Range("A1")`, `Cells(1, 1)`, `Names(1)`: a member of the host's hidden Global
    interface called bare, which the member chains never see because nothing
    precedes it. Its arguments index what it returns the way any member call's do
    (XLIDE issue #202)."""
    if (
        member_ctx is None
        or paren_index != 1
        or match_paren_from(toks, 1) != len(toks) - 1
        or call_name.lower() in module_signatures
        or bare_callable_source_shadowed(call_name, source_names)
        or runtime_callable_source_shadowed(call_name, source_names)
    ):
        return None
    member = resolve_host_global_member(call_name, member_ctx.model)
    if member is None or not member.get("returns"):
        return None
    owner = (member_ctx.model.get("globalType") if member_ctx.model is not None else None) or ""
    entry = MemberCompletionEntry(
        name=member["name"],
        kind=member.get("kind", "property"),
        returns=member.get("returns"),
        signature=member.get("signature"),
        owner=owner,
    )
    type_ = member_expression_return_type(entry, toks[2:-1], member_ctx)
    return InferredArgumentType(
        type_=type_,
        label=f"{call_name}(...) As {type_}",
        span=Span(slice_start + toks[0].start, slice_start + toks[-1].end),
    )


def _infer_atomic_literal(first: VbaToken, span: Span) -> InferredArgumentType | None:
    if first.kind is TokenKind.STRING_LITERAL:
        value = string_literal_value(first.raw_text)
        return InferredArgumentType(
            type_="String", label=f"String literal {first.raw_text}", span=span, string_value=value
        )
    if first.kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL):
        numeric_value = (
            parse_decimal_integer_literal(first.raw_text)
            if first.kind is TokenKind.INTEGER_LITERAL
            else None
        )
        float_value = (
            float_literal_value(first.raw_text) if first.kind is TokenKind.FLOAT_LITERAL else None
        )
        return InferredArgumentType(
            type_="Double",
            label=f"numeric literal {first.raw_text}",
            span=span,
            numeric_value=numeric_value,
            numeric_text=first.raw_text,
            float_value=float_value if float_value is not None and math.isfinite(float_value) else None,
        )
    if first.kind is TokenKind.DATE_LITERAL:
        return InferredArgumentType(type_="Date", label="Date literal", span=span)
    if first.kind is TokenKind.KEYWORD:
        word = first.raw_text.lower()
        if word in ("true", "false"):
            return InferredArgumentType(type_="Boolean", label="Boolean literal", span=span)
        if word == "nothing":
            return InferredArgumentType(type_="Nothing", label="Nothing", span=span)
        if word == "null":
            return InferredArgumentType(type_="Null", label="Null", span=span)
    return None


# -- member-expression typing ----------------------------------------------


def infer_member_expression_type(
    source: str,
    toks: Sequence[VbaToken],
    slice_start: int,
    member_ctx: MemberCompletionContext,
) -> InferredArgumentType | None:
    """The type a member-access expression yields (`ActiveSheet.Range("A1")`,
    `p.Name`), from the member the completion context resolves."""
    if has_top_level_operator(toks):
        return None
    resolved = final_member_token_in_expression(toks)
    if resolved is None:
        return None
    member = resolve_exact_member_completion(
        source, resolved.name, slice_start + resolved.token.end, member_ctx
    )
    if member is None or not member.returns:
        return None
    if not resolved.called and member.kind == "method" and not member_accepts_zero_arguments(member):
        return None
    return_type = member_expression_return_type(member, resolved.argument_tokens, member_ctx)
    label_start = toks[0].start if toks else resolved.token.start
    label_end = toks[-1].end if resolved.called else resolved.token.end
    label_text = source[slice_start + label_start : slice_start + label_end].strip()
    return InferredArgumentType(
        type_=return_type,
        label=f"{label_text} As {return_type}",
        span=Span(slice_start + resolved.token.start, slice_start + resolved.token.end),
    )


# -- operator expressions --------------------------------------------------


def infer_arithmetic_expression_type(
    toks: list[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    module_signatures: _SignatureMap,
    source_names: SourceNameScope | None = None,
    resolve_expression_type: SourceDeclaredTypeResolver | None = None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None = None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> InferredArgumentType | None:
    parts = split_top_level_arithmetic_operands(toks)
    if len(parts) < 2:
        return None
    for part in parts:
        inferred = infer_expression_type(
            part, slice_start, env, module_signatures, source_names,
            resolve_expression_type, resolve_qualified_expression_type,
            source=source, member_ctx=member_ctx,
        )
        normalized = normalize_type(inferred.type_ if inferred is not None else None)
        if not normalized or not is_numeric_type(normalized):
            return None
    return InferredArgumentType(
        type_="Double", label="numeric expression", span=span_for_tokens(toks, slice_start)
    )


def infer_string_concatenation_expression_type(
    toks: list[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    module_signatures: _SignatureMap,
    source_names: SourceNameScope | None = None,
    resolve_expression_type: SourceDeclaredTypeResolver | None = None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None = None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> InferredArgumentType | None:
    parts = split_top_level_operands(toks, "&")
    if len(parts) < 2:
        return None
    for part in parts:
        inferred = infer_expression_type(
            part, slice_start, env, module_signatures, source_names,
            resolve_expression_type, resolve_qualified_expression_type,
            source=source, member_ctx=member_ctx,
        )
        normalized = normalize_type(inferred.type_ if inferred is not None else None)
        if not normalized or not is_string_concatenation_operand_type(normalized):
            return None
    return InferredArgumentType(
        type_="String",
        label="string concatenation expression",
        span=span_for_tokens(toks, slice_start),
    )


# -- string-arithmetic operand check ---------------------------------------


def nonnumeric_string_arithmetic_operand(
    expected_raw: str, slot: Sequence[VbaToken], slice_start: int
) -> InferredArgumentType | None:
    expected = normalize_type(expected_raw)
    if not expected or not is_numeric_type(expected):
        return None
    return find_nonnumeric_string_in_arithmetic_expression(_significant(slot), slice_start)


# -- ByRef variable type mismatch ------------------------------------------


def byref_variable_type_mismatch(
    param: CallableParamType,
    slot: Sequence[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    resolve_expression_type: SourceDeclaredTypeResolver | None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None,
) -> ByRefMismatch | None:
    """by_ref_variable_type_mismatch (types/type_inference.py), under the name the
    port's rules have imported it by."""
    return by_ref_variable_type_mismatch(
        param, slot, slice_start, env, resolve_expression_type, resolve_qualified_expression_type
    )


# -- type compatibility ----------------------------------------------------


def incompatibility_reason(expected_raw: str, actual: InferredArgumentType) -> str | None:
    expected = normalize_type(expected_raw)
    actual_type = normalize_type(actual.type_)
    if not expected or not actual_type or expected == "variant" or actual_type == "variant":
        return None
    if actual_type == "error" and is_known_scalar_type(expected):
        return (
            "An Error Variant cannot be coerced to this scalar type. "
            "This will raise Run-time error '13': Type mismatch."
        )
    if actual_type == "null" and is_known_scalar_type(expected):
        return (
            "Null cannot be coerced to this scalar type. "
            "This will raise Run-time error '94': Invalid use of Null."
        )
    if expected == "object":
        if actual_type in ("nothing", "object") or not is_known_scalar_type(actual_type):
            return None
        return "An object parameter requires an object value."
    if is_numeric_type(expected):
        overflow = numeric_literal_overflow_reason(expected, actual)
        if overflow is not None:
            return overflow
        if is_numeric_type(actual_type) or actual_type == "boolean":
            return None
        if actual_type == "string" and actual.string_value is not None:
            # The string's number where every locale reads it alike: "&H10000" is
            # 65536 and overflows an Integer (XLIDE issue #188).
            verdict = numeric_string_verdict(actual.string_value)
            if verdict.kind == "invalid":
                if actual.held_by is not None:
                    return (
                        f"'{actual.held_by}' holds {js_string_literal(actual.string_value)} here, "
                        "which converts to no number. This will raise Run-time error '13': "
                        "Type mismatch."
                    )
                return (
                    "This string literal cannot be converted to a numeric value. "
                    "This will raise Run-time error '13': Type mismatch."
                )
            bounds = numeric_literal_bounds(expected) if verdict.value is not None else None
            if (
                bounds is not None
                and verdict.value is not None
                and (verdict.value < bounds.min or verdict.value > bounds.max)
            ):
                return (
                    f"The string {js_string_literal(actual.string_value)} converts to "
                    f"{js_number_to_string(verdict.value)}, outside the {bounds.label} range "
                    f"{bounds.min} to {bounds.max}. This will raise Run-time error '6': Overflow."
                )
        return None
    if expected == "boolean":
        if actual_type == "boolean" or is_numeric_type(actual_type):
            return None
        # A String whose value is not known may be "True" or "5", which convert
        # (measured in Excel 16.0).
        if (
            actual_type == "string"
            and actual.string_value is not None
            and is_invalid_boolean_string(actual.string_value)
        ):
            return (
                "This string literal cannot be converted to Boolean. "
                "This will raise Run-time error '13': Type mismatch."
            )
        return None
    if expected == "date" and actual_type == "string" and actual.string_value is not None:
        return (
            "This string literal cannot be converted to a Date. "
            "This will raise Run-time error '13': Type mismatch."
            if is_invalid_date_string(actual.string_value)
            else None
        )
    return None  # String accepts any stringifiable scalar; do not warn.


# -- object and value arguments --------------------------------------------


@dataclass(frozen=True, slots=True)
class _ObjectValueProblem:
    # "argumentObjectTypeMismatch" | "argumentTypeMismatch"
    rule: str
    what: str
    reason: str
    tokens: Sequence[VbaToken]


_VALUE_LITERAL_KINDS = frozenset(
    {TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL, TokenKind.STRING_LITERAL, TokenKind.DATE_LITERAL}
)
_TYPE_MISMATCH_COMPILE = "An object parameter takes an object. This is a VBE compile error: Type mismatch."
_RAISES_13 = "This will raise Run-time error '13': Type mismatch."


def _object_value_argument_problem(
    expected: str,
    slot: Sequence[VbaToken],
    actual: InferredArgumentType | None,
    member_ctx: MemberCompletionContext,
    source_names: SourceNameScope | None,
    is_declared: Callable[[str], bool],
    held_class_of: Callable[[str], str | None] | None = None,
    by_value: bool = False,
    env: Mapping[str, str] | None = None,
) -> _ObjectValueProblem | None:
    """An object where a parameter takes a value, or a value where it takes an
    object (XLIDE issue #223, measured in Excel 16.0):

    - Nothing into a Long or String parameter: "Invalid use of object".
    - New Collection there: "Argument not optional", since its default member
      Item needs an index.
    - A literal, True, False, a date, or an expression of a known scalar type into
      a Collection or other known object parameter: "Type mismatch". So is a
      scalar variable passed by value (#410). Each is a compile error.
    - A Variant holding no object, passed by value to one: 424 when the call
      runs (#410).
    - Array(...) or Split(...) into a Long or String parameter: an array, which
      raises 13 when the call runs.
    """
    from .rules.type_of_is import object_assignment_incompatibility_reason

    declared_env: Mapping[str, str] = env if env is not None else {}
    toks = unwrap_outer_parens(
        [tok for tok in slot if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE]
    )
    if not toks:
        return None
    expected_type = normalize_type(expected)
    if expected_type and is_known_scalar_type(expected_type):
        if len(toks) == 1 and token_text(toks[0]) == "nothing":
            return _ObjectValueProblem(
                "argumentObjectTypeMismatch",
                "Nothing",
                "This is a VBE compile error: Invalid use of object.",
                toks,
            )
        if len(toks) == 2 and token_text(toks[0]) == "new" and object_value_needs_index(toks[1].raw_text, member_ctx):
            return _ObjectValueProblem(
                "argumentObjectTypeMismatch",
                f"New {toks[1].raw_text}, whose default member Item needs an index",
                "This is a VBE compile error: Argument not optional.",
                toks,
            )
        # A Collection variable passed by value is read for its value, the same
        # way (XLIDE issue #647). ByRef it is byref-argument-type-mismatch's.
        passed = token_name(toks[0]) if len(toks) == 1 else None
        passed_type = declared_env.get(passed.lower()) if passed else None
        if by_value and passed_type and is_declared(toks[0].raw_text) and object_value_needs_index(passed_type, member_ctx):
            return _ObjectValueProblem(
                "argumentObjectTypeMismatch",
                f"'{toks[0].raw_text}', declared {passed_type}, whose default member Item needs an index",
                "This is a VBE compile error: Argument not optional.",
                toks,
            )
        callee = token_text(toks[0])
        if (
            callee in ("array", "split")
            and len(toks) > 1
            and toks[1].raw_text == "("
            and match_paren_from(toks, 1) == len(toks) - 1
            and not runtime_callable_source_shadowed(toks[0].raw_text, source_names)
        ):
            return _ObjectValueProblem(
                "argumentTypeMismatch", f"{toks[0].raw_text}(...), an array", _RAISES_13, toks
            )
        return None
    # A literal, True, False or a date, and an expression of a known scalar type
    # such as `v + 0` (XLIDE issue #410).
    atom = toks[1:] if toks[0].raw_text == "-" else toks
    literal = len(atom) == 1 and (
        atom[0].kind in _VALUE_LITERAL_KINDS
        or (len(toks) == 1 and token_text(atom[0]) in ("true", "false"))
    )
    # `b + 0` and `d + 0` with b a Boolean and d a Date too (XLIDE issue #647).
    scalar_expression = (
        len(toks) > 1
        and not literal
        and (
            (actual is not None and is_known_scalar_type(normalize_type(actual.type_) or ""))
            or arithmetic_of_scalars(toks, declared_env)
        )
    )
    if (
        (literal or scalar_expression)
        and expected_type != "object"
        and is_known_object_assignment_type(expected, member_ctx)
    ):
        return _ObjectValueProblem(
            "argumentObjectTypeMismatch",
            actual.label if actual is not None else " ".join(tok.raw_text for tok in toks),
            _TYPE_MISMATCH_COMPILE,
            toks,
        )
    # An object of another class, as a Set of it would be: TakeWs(Range("A1")) and
    # TakeWs(ThisWorkbook) into a Worksheet raise 13 when the call runs (#223). A
    # declared variable is judged by what it holds, not its declared class (#246).
    first_name = token_name(toks[0])
    declared_name = len(toks) == 1 and first_name is not None and is_declared(toks[0].raw_text)
    held = (
        held_class_of(first_name.lower())
        if declared_name and held_class_of is not None and first_name is not None
        else None
    )
    # A scalar variable passed by value is a value, as a literal is; ByRef it is
    # byref-argument-type-mismatch's. A Variant holding Empty or a value raises
    # 424 (XLIDE issue #410).
    if (
        declared_name
        and by_value
        and expected_type != "object"
        and is_known_object_assignment_type(expected, member_ctx)
        and first_name is not None
    ):
        declared_raw = declared_env.get(first_name.lower())
        declared = normalize_type(declared_raw)
        if declared and is_known_scalar_type(declared):
            return _ObjectValueProblem(
                "argumentObjectTypeMismatch",
                f"'{toks[0].raw_text}', declared {declared_raw}",
                _TYPE_MISMATCH_COMPILE,
                toks,
            )
        if held == VALUE_HELD:
            return _ObjectValueProblem(
                "argumentTypeMismatch",
                f"'{toks[0].raw_text}', a Variant that holds no object here",
                "An object parameter takes an object. This will raise Run-time error '424': "
                "Object required.",
                toks,
            )
    if held and held != VALUE_HELD and expected_type != "object" and is_known_object_assignment_type(expected, member_ctx):
        holding = InferredArgumentType(
            type_=held,
            label=f"'{toks[0].raw_text}', which holds a {held} here",
            span=Span(toks[0].start, toks[0].end),
        )
        reason = object_assignment_incompatibility_reason(expected, holding, member_ctx)
        if reason:
            return _ObjectValueProblem("argumentTypeMismatch", holding.label, f"{reason} {_RAISES_13}", toks)
    # `TakeC(ActiveSheet)` into a Collection, as the Object holding it (XLIDE #685).
    if (
        not declared_name
        and len(toks) == 1
        and token_text(toks[0]) == "activesheet"
        and not is_declared(toks[0].raw_text)
        and resolve_host_global("ActiveSheet", member_ctx.model) is not None
        and expected_type != "object"
        and is_known_object_assignment_type(expected, member_ctx)
    ):
        sheet = InferredArgumentType(
            type_="Worksheet or Chart",
            label=f"'{toks[0].raw_text}', a Worksheet or a Chart",
            span=Span(toks[0].start, toks[0].end),
        )
        reason = object_assignment_incompatibility_reason(expected, sheet, member_ctx)
        if reason:
            return _ObjectValueProblem("argumentTypeMismatch", sheet.label, f"{reason} {_RAISES_13}", toks)
    # `TakeW(Worksheets)` into a parameter As Worksheets (XLIDE issue #404).
    sheets = (
        sheets_from_collection_property(toks, expected, source_names, member_ctx)
        if not declared_name and source_names is not None
        else None
    )
    if sheets is not None:
        return _ObjectValueProblem(
            "argumentTypeMismatch",
            f"'{sheets.text}', which returns a Sheets object",
            f"Excel's Worksheets and Charts properties return a Sheets object, never a "
            f"{sheets.collection} one. {_RAISES_13}",
            toks,
        )
    if (
        actual is not None
        and not declared_name
        and expected_type != "object"
        and is_known_object_assignment_type(expected, member_ctx)
        and not is_known_scalar_type(normalize_type(actual.type_) or "")
    ):
        reason = object_assignment_incompatibility_reason(expected, actual, member_ctx)
        if reason:
            return _ObjectValueProblem("argumentTypeMismatch", actual.label, f"{reason} {_RAISES_13}", toks)
    return None


# -- argument-type validation ----------------------------------------------


def validate_argument_types(
    call: CallArguments,
    env: Mapping[str, str],
    module_signatures: _SignatureMap,
    source_names: SourceNameScope | None,
    push: PushFn,
    resolve_expression_type: SourceDeclaredTypeResolver | None = None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None = None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
    held_class_of: Callable[[str], str | None] | None = None,
    held_null: Callable[[str], bool] | None = None,
    held_number: Callable[[str], int | float | str | None] | None = None,
) -> None:
    sig = callable_signature_for_call(call, module_signatures, source_names)
    if sig is None or not sig.params:
        return
    validate_argument_types_for_signature(
        sig, call, env, module_signatures, source_names, push,
        resolve_expression_type, resolve_qualified_expression_type,
        source=source, member_ctx=member_ctx,
        held_class_of=held_class_of, held_null=held_null, held_number=held_number,
    )


_PAREN_SPACING_RE = re.compile(r" ?([()]) ?")


def _is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _is_whole(value: int | float) -> bool:
    return isinstance(value, int) or (math.isfinite(value) and value.is_integer())


def validate_argument_types_for_signature(
    sig: CallableTypeSignature,
    call: CallArguments,
    env: Mapping[str, str],
    module_signatures: _SignatureMap,
    source_names: SourceNameScope | None,
    push: PushFn,
    resolve_expression_type: SourceDeclaredTypeResolver | None = None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None = None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
    held_class_of: Callable[[str], str | None] | None = None,
    held_null: Callable[[str], bool] | None = None,
    held_number: Callable[[str], int | float | str | None] | None = None,
) -> None:
    if not sig.params:
        return
    params_by_name = {strip_header_brackets(p.name).lower(): p for p in sig.params}

    def holds_null(lower: str) -> bool:
        return held_null is not None and held_null(lower) is True

    def is_declared(name: str) -> bool:
        return name.lower() in env or (
            resolve_expression_type is not None and resolve_expression_type(name).resolved is True
        )

    positional_index = 0
    for slot in call.slots:
        named = named_argument_slot(slot)
        param: CallableParamType | None
        if named is not None:
            param = params_by_name.get(named[0].lower())
            value_slot = named[1]
        else:
            param = sig.params[min(positional_index, len(sig.params) - 1)]
            if positional_index >= len(sig.params) and not param.param_array:
                continue
            positional_index += 1
            value_slot = slot
        if param is None:
            continue
        expected = param.type_
        if not expected:
            continue
        byref_mismatch = (
            None
            if call.arguments_parenthesized
            else by_ref_variable_type_mismatch(
                param, value_slot, call.slice_start, env,
                resolve_expression_type, resolve_qualified_expression_type,
            )
        )
        if byref_mismatch is not None:
            push(
                "byRefArgumentTypeMismatch",
                f"ByRef argument '{param.name}' of '{sig.name}' expects {expected}, but "
                f"'{byref_mismatch.name}' is declared as {byref_mismatch.actual}. This is a "
                "VBE compile error: ByRef argument type mismatch.",
                byref_mismatch.span,
            )
            continue
        # An array parameter takes an array variable; anything else is
        # argument-shape-mismatch's, not a value to convert (XLIDE issue #410).
        if param.is_array:
            continue
        # `TakeL(1 + Null)`: an operator on Null gives Null, which a typed
        # parameter refuses (XLIDE issue #324, measured in Excel 16.0).
        null_slot = [tok for tok in value_slot if tok.kind is not TokenKind.COMMENT]
        scalar_expected = normalize_type(expected)
        if (
            len(null_slot) > 1
            and scalar_expected
            and scalar_expected != "variant"
            and is_known_scalar_type(scalar_expected)
            and operator_yields_null(
                null_slot,
                lambda tok: token_text(tok) == "null" or holds_null((token_name(tok) or "").lower()),
            )
        ):
            text = _PAREN_SPACING_RE.sub(r"\1", " ".join(tok.raw_text for tok in null_slot))
            push(
                "argumentTypeMismatch",
                f"Argument '{param.name}' of '{sig.name}' expects {expected}, but '{text}' is "
                "Null: an operator on Null gives Null. Null cannot be coerced to this scalar "
                "type. This will raise Run-time error '94': Invalid use of Null.",
                Span(call.slice_start + null_slot[0].start, call.slice_start + null_slot[-1].end),
            )
            continue
        string_arithmetic = nonnumeric_string_arithmetic_operand(
            expected, value_slot, call.slice_start
        )
        if string_arithmetic is not None:
            push(
                "stringArithmeticCoercion",
                f"Argument '{param.name}' of '{sig.name}' expects {expected}, but this numeric "
                f"expression contains {string_arithmetic.label}. This will raise Run-time error "
                "'13': Type mismatch.",
                string_arithmetic.span,
            )
            continue
        actual = infer_argument_type(
            value_slot, call.slice_start, env, module_signatures, source_names,
            resolve_expression_type, resolve_qualified_expression_type,
            source=source, member_ctx=member_ctx,
        )
        kind_problem = (
            _object_value_argument_problem(
                expected,
                value_slot,
                actual,
                member_ctx,
                source_names,
                is_declared,
                held_class_of,
                param.by_ref is False or call.arguments_parenthesized,
                env,
            )
            if member_ctx is not None
            else None
        )
        if kind_problem is not None:
            push(
                kind_problem.rule,
                f"Argument '{param.name}' of '{sig.name}' expects {expected}, but got "
                f"{kind_problem.what}. {kind_problem.reason}",
                Span(
                    call.slice_start + kind_problem.tokens[0].start,
                    call.slice_start + kind_problem.tokens[-1].end,
                ),
            )
            continue
        # A Variant local a straight line has just set to Null is Null here, for
        # both checks below, whether or not it was given a type, and so is an
        # element `a(0)` or `a(i)` (XLIDE issue #332, measured in Excel 16.0).
        element = (
            token_name(value_slot[0])
            if len(value_slot) == 4 and value_slot[1].raw_text == "(" and value_slot[3].raw_text == ")"
            else None
        )
        single = token_name(value_slot[0]) if len(value_slot) == 1 else None
        held_name = (
            (single.lower() if single is not None else None)
            if len(value_slot) == 1
            else f"{element.lower()}({value_slot[2].raw_text.lower()})"
            if element is not None
            else None
        )
        # Mid hands a Null string back without reading its Length: `Mid(n0, 1, n2)`
        # with both Null runs (XLIDE issue #664, measured in Excel 16.0).
        first = [
            tok
            for tok in (call.slots[0] if call.slots else [])
            if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE
        ]
        null_string = (
            sig.name.lower() == "mid"
            and param.name.lower() == "length"
            and len(first) == 1
            and (token_text(first[0]) == "null" or holds_null((token_name(first[0]) or "").lower()))
        )
        held_null_here = (
            not null_string
            and held_name is not None
            and holds_null(held_name)
            and normalize_type(actual.type_ if actual is not None else None) in (None, "variant")
        )
        if held_null_here:
            text = "".join(tok.raw_text for tok in value_slot)
            label = f"'{text}', which holds Null here"
            if actual is None:
                actual = InferredArgumentType(
                    type_="Null",
                    label=label,
                    span=Span(call.slice_start + value_slot[0].start, call.slice_start + value_slot[-1].end),
                )
            else:
                actual = replace(actual, type_="Null", label=label)
        if actual is None:
            continue
        # A local known to hold a number that is not whole is range-checked as that
        # number: `c = 922337203685477.5807@: Space(c)` overflows (XLIDE issue #332).
        # A whole one past the Long range is runtime-argument-value's (#336).
        held_value = (
            held_number(held_name)
            if held_number is not None
            and held_name is not None
            and actual.numeric_value is None
            and actual.float_value is None
            else None
        )
        if isinstance(held_value, (int, float)) and _is_number(held_value) and not _is_whole(held_value):
            actual = replace(actual, held_by=value_slot[0].raw_text, float_value=float(held_value))
        # A local's whole number or String passed by value to the project's own
        # procedure converts as a literal does: `v = -3` then `S v` with `ByVal p
        # As Byte` raises 6, and "abc" into an Integer 13 (XLIDE issue #558). `S
        # (v)` passes a copy whatever p is.
        in_parens_name = (
            token_name(value_slot[1])
            if len(value_slot) == 3 and value_slot[0].raw_text == "(" and value_slot[2].raw_text == ")"
            else None
        )
        in_parens = in_parens_name.lower() if in_parens_name is not None else None
        copied_name = (
            held_name
            if (param.by_ref is False or call.arguments_parenthesized) and held_name is not None
            else in_parens
        )
        own_procedure = module_signatures.get(call.lookup_key or call.name.lower()) is sig
        # A Boolean's True is no -1 here: into a Byte it is 255 (XLIDE issue #664).
        boolean = copied_name is not None and normalize_type(env.get(copied_name)) == "boolean"
        copied = (
            held_number(copied_name)
            if held_number is not None
            and own_procedure
            and copied_name is not None
            and not boolean
            and actual.numeric_value is None
            and actual.float_value is None
            and actual.string_value is None
            else None
        )
        holder = (
            value_slot[1].raw_text
            if in_parens is not None and copied_name == in_parens
            else value_slot[0].raw_text
        )
        if isinstance(copied, (int, float)) and _is_number(copied):
            actual = (
                replace(actual, held_by=holder, numeric_value=copied)
                if _is_whole(copied)
                else replace(actual, held_by=holder, float_value=float(copied))
            )
        elif isinstance(copied, str):
            actual = replace(actual, held_by=holder, type_="String", string_value=copied)
        # A Variant parameter the function still refuses Null for: CStr(Null),
        # Chr(Null), Asc(Null) raise 94 where Left(Null, 1) hands Null back (XLIDE
        # issue #104).
        if param.null_raises and normalize_type(actual.type_) == "null":
            and_label = f", and {actual.label}" if held_null_here else ""
            push(
                "argumentTypeMismatch",
                f"Argument '{param.name}' of '{sig.name}' cannot be Null{and_label}. This will "
                "raise Run-time error '94': Invalid use of Null.",
                actual.span,
            )
            continue
        reason = incompatibility_reason(expected, actual)
        if not reason:
            continue
        rule = (
            "argumentObjectTypeMismatch"
            if normalize_type(expected) == "object"
            else "argumentTypeMismatch"
        )
        push(
            rule,
            f"Argument '{param.name}' of '{sig.name}' expects {expected}, but got "
            f"{actual.label}. {reason}",
            actual.span,
        )
