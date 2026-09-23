"""Argument / expression type inference and type-compatibility checking.

Ported from the inference engine of
xlide_vscode/src/analyzer/diagnostics/typeInference.ts: inferExpressionType and
its operand splitters, the host-global, runtime-object and external-constant
value types, member-expression typing through the member-completion context, the
ByRef / string-arithmetic / numeric-overflow operand checks,
incompatibilityReason, and validateArgumentTypes(ForSignature).

The member-completion paths run only when the caller passes the source and a
member context; without them an expression that needs one resolves to None, and
an unresolved argument type is simply not checked.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, replace

from ..completion.member_access import (
    MemberCompletionContext,
    MemberCompletionEntry,
    is_explicit_element_accessor,
    resolve_exact_member_completion,
)
from ..constants.integer_constant_expression import parse_decimal_integer_literal
from ..host.host_model import (
    HostConstant,
    HostObjectModel,
    get_host_members,
    resolve_host_constant,
    resolve_host_global,
)
from ..lexer.token_helpers import match_paren_from
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import Span
from ..runtime.vba_runtime import VbaRuntimeConstant, resolve_runtime_constant, resolve_runtime_object
from ..symbols.symbol_model import qualified_procedure_key
from ..types.type_inference import SourceDeclaredType
from ..types.type_names import (
    is_boolean_string,
    is_known_scalar_type,
    is_numeric_type,
    is_provably_non_numeric_string,
    is_string_concatenation_operand_type,
    normalize_type,
    numeric_literal_bounds,
)
from .call_extraction import (
    CallableParamType,
    CallableTypeSignature,
    CallArguments,
    InferredArgumentType,
    callable_accepts_zero_arguments,
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
    parse_runtime_display_signature,
    runtime_callable_source_shadowed,
)
from .const_expr import numeric_external_constant_value
from .context import PushFn
from .walker import span_for_tokens, strip_header_brackets, token_name, token_text

SourceDeclaredTypeResolver = Callable[[str], SourceDeclaredType]
SourceQualifiedDeclaredTypeResolver = Callable[[str, str], SourceDeclaredType]

_SignatureMap = Mapping[str, CallableTypeSignature]


def _significant(toks: list[VbaToken]) -> list[VbaToken]:
    return [t for t in toks if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE]


# -- expression type inference ---------------------------------------------


def infer_argument_type(
    slot: list[VbaToken],
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
    unwrapped = unwrap_outer_parens(toks)
    if len(unwrapped) != len(toks):
        return infer_expression_type(
            unwrapped, slice_start, env, module_signatures, source_names,
            resolve_expression_type, resolve_qualified_expression_type,
            source=source, member_ctx=member_ctx,
        )
    signed = _infer_signed_numeric_literal(toks, slice_start)
    if signed is not None:
        return signed
    concatenation = _infer_string_concatenation_expression_type(
        toks, slice_start, env, module_signatures, source_names,
        resolve_expression_type, resolve_qualified_expression_type,
        source=source, member_ctx=member_ctx,
    )
    if concatenation is not None:
        return concatenation
    arithmetic = _infer_arithmetic_expression_type(
        toks, slice_start, env, module_signatures, source_names,
        resolve_expression_type, resolve_qualified_expression_type,
        source=source, member_ctx=member_ctx,
    )
    if arithmetic is not None:
        return arithmetic
    return _infer_atomic_expression_type(
        toks, slice_start, env, module_signatures, source_names,
        resolve_expression_type, resolve_qualified_expression_type,
        source=source, member_ctx=member_ctx,
    )


def _infer_signed_numeric_literal(toks: list[VbaToken], slice_start: int) -> InferredArgumentType | None:
    if len(toks) != 2 or toks[0].kind is not TokenKind.OPERATOR:
        return None
    sign = toks[0].raw_text
    if sign not in ("+", "-"):
        return None
    literal = toks[1]
    if literal.kind is not TokenKind.INTEGER_LITERAL:
        return None
    value = parse_decimal_integer_literal(literal.raw_text)
    if value is None:
        return None
    signed = -value if sign == "-" else value
    text = f"{sign}{literal.raw_text}"
    return InferredArgumentType(
        type_="Double",
        label=f"numeric literal {text}",
        span=Span(slice_start + toks[0].start, slice_start + literal.end),
        numeric_value=signed,
        numeric_text=text,
    )


def _infer_atomic_expression_type(
    toks: list[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    module_signatures: _SignatureMap,
    source_names: SourceNameScope | None,
    resolve_expression_type: SourceDeclaredTypeResolver | None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> InferredArgumentType | None:
    first = toks[0]
    span = Span(slice_start + first.start, slice_start + first.end)
    if len(toks) == 1:
        literal = _infer_atomic_literal(first, span)
        if literal is not None:
            return literal

    model = member_ctx.model if member_ctx is not None else None
    name = token_name(first)
    if name and len(toks) == 1:
        declared_type = resolve_expression_type(name) if resolve_expression_type else None
        type_ = (
            declared_type.as_type
            if declared_type is not None and declared_type.resolved
            else env.get(name.lower())
        )
        if type_:
            return InferredArgumentType(type_=type_, label=f"{name} As {type_}", span=span)
        sig = _parameterless_value_signature(name, module_signatures, source_names)
        if sig is not None and sig.return_type:
            return InferredArgumentType(
                type_=sig.return_type, label=f"{name} As {sig.return_type}", span=span
            )
        external_object = _infer_bare_external_object_expression_type(name, span, source_names, model)
        if external_object is not None:
            return external_object
        return _infer_bare_external_constant_expression_type(name, span, source_names, model)

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
            _infer_intrinsic_cverr_error_variant(toks, slice_start, module_signatures, source_names)
            if call_name is not None and call_name.paren_index == 1
            else None
        )
        if error_variant is not None:
            return error_variant
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

    if name and len(toks) > 1 and toks[1].raw_text == ".":
        member = token_name(toks[2]) if len(toks) > 2 else None
        error_variant = _infer_intrinsic_cverr_error_variant(
            toks, slice_start, module_signatures, source_names
        )
        if error_variant is not None:
            return error_variant
        if member and len(toks) == 3:
            member_span = Span(slice_start + toks[2].start, slice_start + toks[2].end)
            lookup_key = qualified_procedure_key(name, member)
            sig = _parameterless_value_signature(lookup_key, module_signatures)
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
            external = _infer_qualified_external_constant_expression_type(
                name, member, member_span, model
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


def _infer_bare_external_object_expression_type(
    name: str,
    span: Span,
    source_names: SourceNameScope | None,
    model: HostObjectModel | None,
) -> InferredArgumentType | None:
    """A host global (`Application`, `ActiveCell`) or runtime object (`Err`) named
    bare, unless source declares the name."""
    if runtime_callable_source_shadowed(name, source_names):
        return None
    host_type = resolve_host_global(name, model)
    if host_type:
        return InferredArgumentType(type_=host_type, label=f"{name} As {host_type}", span=span)
    runtime_object = resolve_runtime_object(name)
    if runtime_object is not None:
        runtime_type = runtime_object.get("type", "")
        return InferredArgumentType(type_=runtime_type, label=f"{name} As {runtime_type}", span=span)
    return None


def _infer_bare_external_constant_expression_type(
    name: str,
    span: Span,
    source_names: SourceNameScope | None,
    model: HostObjectModel | None,
) -> InferredArgumentType | None:
    """A VBA runtime or host constant named bare. A name both define stays unknown."""
    if runtime_callable_source_shadowed(name, source_names):
        return None
    candidates = [
        candidate
        for candidate in (
            _inferred_external_constant(name, resolve_runtime_constant(name)),
            _inferred_external_constant(name, resolve_host_constant(name, model)),
        )
        if candidate is not None
    ]
    if len(candidates) != 1:
        return None
    return replace(candidates[0], span=span)


# Qualifiers that name a host's own constant library, resolved against the CURRENT
# host's model: `Word.wdRed` answers in a Word module and misses in an Excel one
# (XLIDE issue #24).
_HOST_CONSTANT_QUALIFIERS = frozenset({"excel", "word", "powerpoint", "access", "office"})


def _infer_qualified_external_constant_expression_type(
    qualifier: str,
    name: str,
    span: Span,
    model: HostObjectModel | None,
) -> InferredArgumentType | None:
    lower = qualifier.lower()
    if lower == "vba":
        inferred = _inferred_external_constant(f"{qualifier}.{name}", resolve_runtime_constant(name))
        return replace(inferred, span=span) if inferred is not None else None
    if lower in _HOST_CONSTANT_QUALIFIERS:
        inferred = _inferred_external_constant(f"{qualifier}.{name}", resolve_host_constant(name, model))
        return replace(inferred, span=span) if inferred is not None else None
    return None


def _inferred_external_constant(
    display_name: str, constant: VbaRuntimeConstant | HostConstant | None
) -> InferredArgumentType | None:
    if constant is None:
        return None
    declared_type = constant.get("type")
    numeric_value = numeric_external_constant_value(constant.get("value"))
    if numeric_value is not None:
        return InferredArgumentType(
            type_="Long",
            label=f"{display_name} As {declared_type if declared_type is not None else 'Long'}",
            span=Span(0, 0),
            numeric_value=numeric_value,
            numeric_text=display_name,
            # A named-constant origin, so an overflow reads as the constant's value
            # rather than as a "numeric literal".
            numeric_constant_name=display_name,
        )
    if normalize_type(declared_type) == "string":
        return InferredArgumentType(
            type_="String",
            label=f"{display_name} As {declared_type if declared_type is not None else 'String'}",
            span=Span(0, 0),
        )
    return None


# -- member-expression typing ----------------------------------------------


@dataclass(frozen=True, slots=True)
class _FinalMemberToken:
    name: str
    token: VbaToken
    called: bool
    argument_tokens: Sequence[VbaToken] | None = None


def infer_member_expression_type(
    source: str,
    toks: Sequence[VbaToken],
    slice_start: int,
    member_ctx: MemberCompletionContext,
) -> InferredArgumentType | None:
    """The type a member-access expression yields (`ActiveSheet.Range("A1")`,
    `p.Name`), from the member the completion context resolves."""
    if _has_top_level_operator(toks):
        return None
    resolved = _final_member_token_in_expression(toks)
    if resolved is None:
        return None
    member = resolve_exact_member_completion(
        source, resolved.name, slice_start + resolved.token.end, member_ctx
    )
    if member is None or not member.returns:
        return None
    if not resolved.called and member.kind == "method" and not _member_accepts_zero_arguments(member):
        return None
    return_type = _member_expression_return_type(member, resolved.argument_tokens, member_ctx)
    label_start = toks[0].start if toks else resolved.token.start
    label_end = toks[-1].end if resolved.called else resolved.token.end
    label_text = source[slice_start + label_start : slice_start + label_end].strip()
    return InferredArgumentType(
        type_=return_type,
        label=f"{label_text} As {return_type}",
        span=Span(slice_start + resolved.token.start, slice_start + resolved.token.end),
    )


def _member_expression_return_type(
    member: MemberCompletionEntry,
    argument_tokens: Sequence[VbaToken] | None,
    member_ctx: MemberCompletionContext,
) -> str:
    # Calling a member with arguments indexes into it. When the member returns a
    # host collection (one whose Item resolves to an element type), the call yields
    # that element: ws.ChartObjects(1) is a ChartObject, not the collection. A
    # concrete-typed call keeps its declared type (ws.Range("A1") stays Range), and
    # Item/_Default/Add already return the resolved element.
    if member.returns and argument_tokens and not is_explicit_element_accessor(member.name):
        element = _default_host_item_return_type(member.returns, member_ctx)
        return element if element is not None else member.returns
    return member.returns if member.returns else "Variant"


def _default_host_item_return_type(type_name: str, member_ctx: MemberCompletionContext) -> str | None:
    item = next(
        (m for m in get_host_members(type_name, member_ctx.model) if m["name"].lower() == "item"),
        None,
    )
    if item is None:
        return None
    if item.get("returns"):
        return item["returns"]
    # A mixed-element collection (Sheets, whose Item is a Worksheet OR a Chart)
    # carries returnsAnyOf instead. Its indexed element is a late-bound Object, which
    # any specific object target accepts, so `Set ws = ThisWorkbook.Sheets("x")` is
    # not a mismatch while single-typed collections stay strict.
    if item.get("returnsAnyOf"):
        return "Object"
    return None


def _final_member_token_in_expression(toks: Sequence[VbaToken]) -> _FinalMemberToken | None:
    if not toks:
        return None
    last = toks[-1]
    last_name = token_name(last)
    if last_name and len(toks) >= 2 and toks[-2].raw_text == ".":
        return _FinalMemberToken(last_name, last, called=False)
    if last.raw_text != ")":
        return None
    open_index = _matching_open_paren_index(toks, len(toks) - 1)
    if open_index < 2:
        return None
    member = toks[open_index - 1]
    member_name = token_name(member)
    if not member_name or toks[open_index - 2].raw_text != ".":
        return None
    return _FinalMemberToken(
        member_name, member, called=True, argument_tokens=toks[open_index + 1 : -1]
    )


def _matching_open_paren_index(toks: Sequence[VbaToken], close: int) -> int:
    depth = 0
    for i in range(close, -1, -1):
        raw = toks[i].raw_text
        if raw == ")":
            depth += 1
        elif raw == "(":
            depth -= 1
            if depth == 0:
                return i
    return -1


def _has_top_level_operator(toks: Sequence[VbaToken]) -> bool:
    depth = 0
    for tok in toks:
        raw = tok.raw_text
        if raw in ("(", "["):
            depth += 1
        elif raw in (")", "]"):
            depth -= 1
        elif depth == 0 and tok.kind is TokenKind.OPERATOR:
            return True
    return False


def _member_accepts_zero_arguments(member: MemberCompletionEntry) -> bool:
    if not member.signature:
        return False
    return callable_accepts_zero_arguments(parse_runtime_display_signature(member.name, member.signature))


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
        return InferredArgumentType(
            type_="Double",
            label=f"numeric literal {first.raw_text}",
            span=span,
            numeric_value=numeric_value,
            numeric_text=first.raw_text,
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


def _infer_intrinsic_cverr_error_variant(
    toks: list[VbaToken],
    slice_start: int,
    module_signatures: _SignatureMap,
    source_names: SourceNameScope | None,
) -> InferredArgumentType | None:
    first_name = token_name(toks[0])
    if not first_name:
        return None
    paren_index = -1
    display_name = ""
    if first_name.lower() == "cverr" and len(toks) > 1 and toks[1].raw_text == "(":
        if (
            first_name.lower() in module_signatures
            or bare_callable_source_shadowed(first_name, source_names)
            or runtime_callable_source_shadowed(first_name, source_names)
        ):
            return None
        paren_index = 1
        display_name = first_name
    elif (
        first_name.lower() == "vba"
        and len(toks) > 3
        and toks[1].raw_text == "."
        and (token_name(toks[2]) or "").lower() == "cverr"
        and toks[3].raw_text == "("
    ):
        paren_index = 3
        display_name = f"{first_name}.{toks[2].raw_text}"
    if paren_index < 0:
        return None
    close = match_paren_from(toks, paren_index)
    if close != len(toks) - 1:
        return None
    inner = toks[paren_index + 1 : close]
    if not inner:
        return None
    split = split_arg_slots(inner, slice_start)
    if len(split.slots) != 1 or not split.slots[0]:
        return None
    return InferredArgumentType(
        type_="Error",
        label=f"{display_name}(...) Error Variant",
        span=span_for_tokens(toks, slice_start),
    )


def _infer_arithmetic_expression_type(
    toks: list[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    module_signatures: _SignatureMap,
    source_names: SourceNameScope | None,
    resolve_expression_type: SourceDeclaredTypeResolver | None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> InferredArgumentType | None:
    parts = _split_top_level_arithmetic_operands(toks)
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


def _infer_string_concatenation_expression_type(
    toks: list[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    module_signatures: _SignatureMap,
    source_names: SourceNameScope | None,
    resolve_expression_type: SourceDeclaredTypeResolver | None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> InferredArgumentType | None:
    parts = _split_top_level_operands(toks, ("&",))
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


def _split_top_level_arithmetic_operands(toks: list[VbaToken]) -> list[list[VbaToken]]:
    parts = _split_top_level_operands(toks, ("+", "-", "*", "/", "\\", "^"))
    return parts if len(parts) >= 2 else []


def _split_top_level_operands(toks: list[VbaToken], operators: tuple[str, ...]) -> list[list[VbaToken]]:
    allowed = set(operators)
    parts: list[list[VbaToken]] = []
    start = 0
    depth = 0
    for i, tok in enumerate(toks):
        raw = tok.raw_text
        if raw in ("(", "["):
            depth += 1
            continue
        if raw in (")", "]"):
            depth -= 1
            continue
        if depth != 0:
            continue
        if tok.kind is not TokenKind.OPERATOR or raw not in allowed:
            if tok.kind is TokenKind.OPERATOR:
                return []
            continue
        # A +/- at the start of an operand is a unary sign (e.g. `2 * -3`,
        # `x + -1`); fold it into the following operand rather than treating it
        # as a separator. Any other operator at the operand start is malformed.
        if i == start:
            if raw in ("+", "-"):
                continue
            return []
        if i == len(toks) - 1:
            return []
        parts.append(toks[start:i])
        start = i + 1
    if not parts:
        return []
    parts.append(toks[start:])
    return parts


def _parameterless_value_signature(
    name: str, module_signatures: _SignatureMap, source_names: SourceNameScope | None = None
) -> CallableTypeSignature | None:
    sig = callable_signature_for(name, module_signatures, source_names)
    if sig is not None and sig.return_type and callable_accepts_zero_arguments(sig):
        return sig
    return None


# -- string-arithmetic operand check ---------------------------------------


def nonnumeric_string_arithmetic_operand(
    expected_raw: str, slot: list[VbaToken], slice_start: int
) -> InferredArgumentType | None:
    expected = normalize_type(expected_raw)
    if not expected or not is_numeric_type(expected):
        return None
    return _find_nonnumeric_string_in_arithmetic_expression(_significant(slot), slice_start)


def _find_nonnumeric_string_in_arithmetic_expression(
    toks: list[VbaToken], slice_start: int
) -> InferredArgumentType | None:
    unwrapped = unwrap_outer_parens(toks)
    if len(unwrapped) != len(toks):
        return _find_nonnumeric_string_in_arithmetic_expression(unwrapped, slice_start)
    parts = _split_top_level_arithmetic_operands(toks)
    if len(parts) < 2:
        return None
    for part in parts:
        nested = _find_nonnumeric_string_in_arithmetic_expression(part, slice_start)
        if nested is not None:
            return nested
        operand = unwrap_outer_parens(part)
        if len(operand) == 1 and operand[0].kind is TokenKind.STRING_LITERAL:
            value = string_literal_value(operand[0].raw_text)
            if is_provably_non_numeric_string(value):
                return InferredArgumentType(
                    type_="String",
                    label=f"nonnumeric string literal {operand[0].raw_text}",
                    span=Span(slice_start + operand[0].start, slice_start + operand[0].end),
                    string_value=value,
                )
    return None


# -- ByRef variable type mismatch ------------------------------------------


@dataclass(frozen=True, slots=True)
class _ByRefMismatch:
    name: str
    actual: str
    span: Span


def _is_known_by_ref_exact_type(type_: str | None) -> bool:
    if not type_ or type_ == "variant":
        return False
    return type_ == "object" or is_known_scalar_type(type_)


def byref_variable_type_mismatch(
    param: CallableParamType,
    slot: list[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    resolve_expression_type: SourceDeclaredTypeResolver | None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None,
) -> _ByRefMismatch | None:
    if not param.by_ref or not param.type_:
        return None
    expected = normalize_type(param.type_)
    if not _is_known_by_ref_exact_type(expected):
        return None
    toks = _significant(slot)
    name: str | None = None
    actual_raw: str | None = None
    span: Span | None = None
    if len(toks) == 1:
        name = token_name(toks[0])
        if not name:
            return None
        declared_type = resolve_expression_type(name) if resolve_expression_type else None
        actual_raw = (
            declared_type.as_type
            if declared_type is not None and declared_type.resolved
            else env.get(name.lower())
        )
        span = Span(slice_start + toks[0].start, slice_start + toks[0].end)
    elif len(toks) == 3 and toks[1].raw_text == ".":
        qualifier = token_name(toks[0])
        member = token_name(toks[2])
        if not qualifier or not member:
            return None
        declared_type = (
            resolve_qualified_expression_type(qualifier, member)
            if resolve_qualified_expression_type
            else None
        )
        if declared_type is None or not declared_type.resolved:
            return None
        name = f"{qualifier}.{member}"
        actual_raw = declared_type.as_type
        span = Span(slice_start + toks[0].start, slice_start + toks[2].end)
    else:
        return None
    actual = normalize_type(actual_raw)
    if not _is_known_by_ref_exact_type(actual) or actual == expected:
        return None
    return _ByRefMismatch(name=name, actual=actual_raw if actual_raw is not None else name, span=span)


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
        overflow = _numeric_literal_overflow_reason(expected, actual)
        if overflow is not None:
            return overflow
        if is_numeric_type(actual_type) or actual_type == "boolean":
            return None
        if actual_type == "string":
            return (
                "This string literal cannot be converted to a numeric value. "
                "This will raise Run-time error '13': Type mismatch."
                if actual.string_value is not None
                and is_provably_non_numeric_string(actual.string_value)
                else None
            )
        return None
    if expected == "boolean":
        if actual_type == "boolean" or is_numeric_type(actual_type):
            return None
        if actual_type == "string":
            return (
                None
                if actual.string_value is not None and is_boolean_string(actual.string_value)
                else "This string literal cannot be converted to Boolean. "
                "This will raise Run-time error '13': Type mismatch."
            )
        return None
    return None  # String accepts any stringifiable scalar; do not warn.


def _format_numeric(value: float) -> str:
    """Format a numeric value the way JS String() would: no '.0' on an integer."""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _numeric_literal_overflow_reason(expected: str, actual: InferredArgumentType) -> str | None:
    if actual.numeric_value is None:
        return None
    bounds = numeric_literal_bounds(expected)
    if bounds is None:
        return None
    if bounds.min <= actual.numeric_value <= bounds.max:
        return None
    # A resolved named constant must not be described as a "numeric literal"; name
    # the constant and show its value instead.
    if actual.numeric_constant_name is not None:
        return (
            f"The value of constant '{actual.numeric_constant_name}' "
            f"({_format_numeric(actual.numeric_value)}) is outside the {bounds.label} range "
            f"{bounds.min} to {bounds.max}. This will raise Run-time error '6': Overflow."
        )
    literal = actual.numeric_text if actual.numeric_text is not None else str(actual.numeric_value)
    return (
        f"The numeric literal {literal} is outside the {bounds.label} range "
        f"{bounds.min} to {bounds.max}. This will raise Run-time error '6': Overflow."
    )


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
) -> None:
    sig = callable_signature_for_call(call, module_signatures, source_names)
    if sig is None or not sig.params:
        return
    validate_argument_types_for_signature(
        sig, call, env, module_signatures, source_names, push,
        resolve_expression_type, resolve_qualified_expression_type,
        source=source, member_ctx=member_ctx,
    )


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
) -> None:
    if not sig.params:
        return
    params_by_name = {strip_header_brackets(p.name).lower(): p for p in sig.params}
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
        byref_mismatch = byref_variable_type_mismatch(
            param, value_slot, call.slice_start, env,
            resolve_expression_type, resolve_qualified_expression_type,
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
        if actual is None:
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
