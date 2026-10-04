"""Rule: argument-shape-mismatch.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/argumentShape.ts. A bare
array variable, or a same-module user-defined Type value, passed where a parameter
is a scalar, or a scalar/Variant passed where a parameter is an array, is a VBE
compile error. Decides purely on declared SHAPE (array-ness / UDT-ness), never on
element-type coercion. Fires only on a single bare identifier argument whose
declared shape resolves to a provable array or same-module Type; quiet on Variant
parameters, matching array/UDT parameters, ParamArray, indexed/member/call/
parenthesized arguments, and unresolved names. Disjoint from
byref-argument-type-mismatch: defers to it whenever that rule owns the slot.

Issue #253, measured in Excel 16.0: a same-module Type variable passed to a
Variant parameter, Optional and ParamArray included, is "Only user-defined types
defined in public object modules can be coerced to or from a variant or passed to
late-bound functions". A Type parameter takes only a variable of its own Type: a
Variant, a number, an object, a literal, a member that holds a number, or another
Type is "ByRef argument type mismatch", and the Type in parentheses, `TakeT (t)` or
`Call TakeT((t))`, is "Variable required - can't assign to this expression". VBA's
own functions with a Variant parameter refuse a Type the same way, but Len, LenB
and VarPtr take one, and CStr's error is "Type mismatch".
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...completion.member_access import MemberCompletionContext
from ...conditional import ConditionalActivityTracker
from ...lexer.token_helpers import match_paren_from
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...symbols.name_resolution import BareIdentifierContext
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureSignature, VbaSymbol, VbaSymbolKind
from ...types.type_inference import (
    SourceDeclaredShape,
    SourceDeclaredType,
    by_ref_variable_type_mismatch,
    declared_shape_for_source_binding,
    declared_value_type_for_qualified_source_binding,
    declared_value_type_for_source_binding,
    procedure_symbol_for,
    same_module_type_names,
    type_environment_for,
)
from ...types.type_names import is_known_scalar_type, normalize_type
from ..call_extraction import (
    CallableParamType,
    CallableTypeSignature,
    CallArguments,
    extract_call,
    extract_qualified_call,
    named_argument_slot,
)
from ..callable_signatures import (
    callable_signature_for_call,
    callable_type_signatures_for,
    expression_calls,
    source_name_scope_for,
)
from ..context import PushFn
from ..model import VbaDiagnosticData
from ..type_fields import ModuleTypes, field_chain, module_types, type_key, variable_root, variable_symbol_in
from ..walker import ProcedureStatementVisitor, strip_header_brackets, token_name

_ShapeResolver = Callable[[str], SourceDeclaredShape]
_TypeResolver = Callable[[str], SourceDeclaredType]
_QualifiedTypeResolver = Callable[[str, str], SourceDeclaredType]


@dataclass(frozen=True, slots=True)
class _MemberType:
    """A member's type: lowercased, and as written."""

    type: str
    name: str


_MemberTypeOf = Callable[[Sequence[VbaToken]], "_MemberType | None"]


def _no_member_type(_toks: Sequence[VbaToken]) -> _MemberType | None:
    return None


def check_argument_shape(
    source: str,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    _member_ctx: MemberCompletionContext,
    push: PushFn,
    mod: ModuleNode | None = None,
    activity: ConditionalActivityTracker | None = None,
) -> ProcedureStatementVisitor:
    """Per-statement rule: rides the shared procedure-statement walk (audit #0)."""
    module_signatures = callable_type_signatures_for(symbols, project_procedures)
    udt_names = same_module_type_names(symbols)
    types: ModuleTypes = module_types(source, mod, activity) if mod is not None else {}

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        # `t.a`: the type of the member a chain of fields ends at.
        def member_type(toks: Sequence[VbaToken]) -> _MemberType | None:
            name = token_name(toks[0]) if toks else None
            lower = name.lower() if name else None
            root = variable_root(toks, 0, variable_symbol_in(symbols, member, lower), types) if lower else None
            steps = field_chain(toks, root, types) if root is not None else []
            last = steps[-1] if steps else None
            if (
                last is not None
                and (last.close if last.close is not None else last.at) == len(toks) - 1
                and (not last.field.is_array or last.open is not None)
            ):
                return _MemberType(
                    type=last.field.type if last.field.type is not None else "variant",
                    name=last.field.type_name if last.field.type_name is not None else "Variant",
                )
            return None

        env = type_environment_for(symbols, member)
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        proc_sym = procedure_symbol_for(symbols, member)

        def resolve_type(name: str) -> SourceDeclaredType:
            return declared_value_type_for_source_binding(symbols, proc_sym, project_visible_symbols, name)

        def resolve_qualified_type(qualifier: str, name: str) -> SourceDeclaredType:
            return declared_value_type_for_qualified_source_binding(
                symbols, project_visible_symbols, qualifier, name
            )

        def resolve_shape(name: str) -> SourceDeclaredShape:
            return declared_shape_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.EXPRESSION
            )

        def visitor(stmt: LeafStatementNode) -> None:
            # `Call PutD((d))` is found both as an expression call and as the
            # statement's call; report each argument once.
            reported: set[tuple[int, int, str]] = set()

            def push_once(rule: str, message: str, span: Span, data: VbaDiagnosticData | None = None) -> None:
                key = (span.start, span.end, message)
                if key not in reported:
                    reported.add(key)
                    push(rule, message, span, data)

            def check_call(call: CallArguments) -> None:
                sig = callable_signature_for_call(call, module_signatures, source_names)
                if sig is None or not sig.params:
                    return
                _validate_argument_shapes(
                    sig, call, udt_names, env, resolve_type, resolve_qualified_type, resolve_shape,
                    push_once, member_type,
                )

            for call in expression_calls(source, stmt.span, module_signatures, source_names):
                check_call(call)
            statement_call = extract_call(source, stmt.span) or extract_qualified_call(
                source, stmt.span, module_signatures
            )
            if statement_call is not None:
                check_call(statement_call)

        return visitor

    return factory


def _validate_argument_shapes(
    sig: CallableTypeSignature,
    call: CallArguments,
    udt_names: AbstractSet[str],
    env: Mapping[str, str],
    resolve_type: _TypeResolver,
    resolve_qualified_type: _QualifiedTypeResolver,
    resolve_shape: _ShapeResolver,
    push: PushFn,
    member_type: _MemberTypeOf = _no_member_type,
) -> None:
    # Slot -> parameter pairing mirrors validate_argument_types_for_signature
    # (named argument -> by name, otherwise positional by index). Kept local
    # because the shape rule treats ParamArray as always-accepting.
    params_by_name = {strip_header_brackets(p.name).lower(): p for p in sig.params}
    positional_index = 0
    for slot in call.slots:
        named = named_argument_slot(slot)
        param: CallableParamType | None
        value_slot = slot
        if named is not None:
            param = params_by_name.get(named[0].lower())
            value_slot = named[1]
        else:
            param = sig.params[min(positional_index, len(sig.params) - 1)]
            if positional_index >= len(sig.params) and not param.param_array:
                continue
            positional_index += 1
        type_problem = (
            _type_argument_problem(sig.name, param, call, value_slot, udt_names, resolve_shape, member_type)
            if param is not None
            else None
        )
        if type_problem is not None:
            code, message, span = type_problem
            push(code, message, span)
            continue
        if param is None or param.param_array:
            continue  # ParamArray parameters are Variant and accept any shape
        # Defer to byref-argument-type-mismatch when it owns this slot.
        if by_ref_variable_type_mismatch(
            param, value_slot, call.slice_start, env, resolve_type, resolve_qualified_type
        ) is not None:
            continue
        if param.is_array:
            # `PutD (d)` as a statement passes `(d)`, a value (issue #218).
            problem = (
                _parenthesized_argument(value_slot, call.slice_start)
                if call.arguments_parenthesized and len(call.slots) == 1
                else _array_argument_problem(value_slot, call.slice_start, param, resolve_shape, resolve_type)
            )
            if problem is not None:
                what, span = problem
                element = param.type_ if param.type_ is not None else "Variant"
                push(
                    "argumentShapeMismatch",
                    f"{what}, but parameter '{param.name}' of '{sig.name}' is an array of {element}. "
                    "This is a VBE compile error: Type mismatch: array or user-defined type expected.",
                    span,
                )
                continue
        ident = _sole_identifier(value_slot, call.slice_start)
        if ident is None:
            continue
        name, ident_span = ident
        shape = resolve_shape(name)
        if not shape.resolved or shape.shape is None:
            continue  # unresolved / ambiguous -> quiet
        if shape.shape.is_array:
            if not param.is_array and _param_is_known_scalar(param):
                push("argumentShapeMismatch", _array_to_scalar_message(name, param, sig.name), ident_span)
            continue
        as_type = shape.shape.as_type
        if as_type and as_type.lower() in udt_names:
            if not param.is_array and _param_is_known_scalar(param):
                push(
                    "argumentShapeMismatch",
                    _udt_to_scalar_message(name, as_type, param, sig.name),
                    ident_span,
                )
            continue
        # A Collection is no array either (issue #410, measured in Excel 16.0).
        if param.is_array and as_type and (_is_scalar_or_variant(as_type) or normalize_type(as_type) == "collection"):
            push("argumentShapeMismatch", _scalar_to_array_message(name, param, sig.name), ident_span)


def _significant(slot: Sequence[VbaToken]) -> list[VbaToken]:
    return [t for t in slot if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE]


_LITERAL_KINDS = frozenset(
    {TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL, TokenKind.STRING_LITERAL, TokenKind.DATE_LITERAL}
)
_SPLIT_OR_ARRAY_RE = re.compile(r"^(split|array)$", re.IGNORECASE)


def _array_argument_problem(
    slot: Sequence[VbaToken],
    slice_start: int,
    param: CallableParamType,
    resolve_shape: _ShapeResolver,
    resolve_type: _TypeResolver | None = None,
) -> tuple[str, Span] | None:
    """An array parameter takes an array variable of its own element type, and
    nothing else (issue #216, measured in Excel 16.0): a function's result,
    `Split(...)` or `Array(...)`, is refused, and so is an array of String for an
    array of Variant."""
    toks = _significant(slot)
    # `(d)` is a value, even around an array of the right type: IRR((d)) with d a
    # Double array is refused (issue #218, measured).
    if toks and toks[0].raw_text == "(" and match_paren_from(toks, 0) == len(toks) - 1:
        return _parenthesized_argument(toks[1:-1], slice_start)
    # `TArr -1`, `TArr "a"`, `TArr Null`: a literal is no array (issues #410 and
    # #556, measured in Excel 16.0).
    literal = (
        toks[1]
        if len(toks) == 2 and toks[0].raw_text in ("-", "+")
        else toks[0]
        if len(toks) == 1
        else None
    )
    if literal is not None and (
        literal.kind in _LITERAL_KINDS or (len(toks) == 1 and literal.raw_text.lower() in ("null", "empty"))
    ):
        return (
            f"{''.join(t.raw_text for t in toks)} is a literal, not an array",
            Span(slice_start + toks[0].start, slice_start + toks[-1].end),
        )
    # `d + 0`: an expression's value is no array variable (issue #647, measured
    # in Excel 16.0).
    if len(toks) > 2 and _top_level_operator(toks):
        return (
            f"'{' '.join(t.raw_text for t in toks)}' is an expression, not an array",
            Span(slice_start + toks[0].start, slice_start + toks[-1].end),
        )
    name = token_name(toks[0]) if toks else None
    if not name:
        return None
    span = Span(slice_start + toks[0].start, slice_start + toks[-1].end)
    shape = resolve_shape(name)
    if len(toks) > 1 and toks[1].raw_text == "(" and toks[-1].raw_text == ")":
        # The result of VBA's Split or Array, which no project name shadows.
        if not shape.resolved and _SPLIT_OR_ARRAY_RE.match(name):
            return (f"'{name}(...)' is a function's result, not an array variable", span)
        # `a(0)` of an array variable is one element; `a()` is the array (issue
        # #223, measured in Excel 16.0). A Function returning an array is called
        # here, not indexed.
        kind = resolve_type(name).kind if resolve_type is not None else None
        variable = kind in (VbaSymbolKind.LOCAL_VARIABLE, VbaSymbolKind.MODULE_VARIABLE, VbaSymbolKind.PARAMETER)
        if variable and shape.resolved and shape.shape is not None and shape.shape.is_array and len(toks) > 3:
            return (
                f"'{''.join(t.raw_text for t in toks)}' is one element of the array '{name}', not the array",
                span,
            )
        return None
    if len(toks) != 1 or not shape.resolved or shape.shape is None or not shape.shape.is_array:
        return None
    element = normalize_type(shape.shape.as_type) or "variant"
    expected = normalize_type(param.type_) or "variant"
    if (
        element != expected
        and (is_known_scalar_type(element) or element == "variant")
        and (is_known_scalar_type(expected) or expected == "variant")
    ):
        as_written = shape.shape.as_type if shape.shape.as_type is not None else "Variant"
        return (f"'{name}' is an array of {as_written}", span)
    return None


_BINARY_OPERATORS = frozenset(
    {"+", "-", "*", "/", "\\", "^", "&", "mod", "and", "or", "xor", "=", "<>", "<", ">", "<=", ">="}
)


def _top_level_operator(toks: Sequence[VbaToken]) -> bool:
    """Whether an operator joins two operands outside any parentheses: `d + 0`, `a & b`."""
    depth = 0
    for i, tok in enumerate(toks):
        raw = tok.raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
        elif depth == 0 and i > 0 and (raw in _BINARY_OPERATORS or raw.lower() in _BINARY_OPERATORS):
            return True
    return False


# VBA functions this rule leaves alone: Len, LenB and VarPtr take a user-defined
# type (issue #253, measured), and MsgBox is statement_types.py's.
_LEFT_ALONE = frozenset({"len", "lenb", "varptr", "msgbox"})

# The conversion functions; only CVar and CStr were measured with a Type.
_CONVERSIONS = frozenset(
    {
        "cbool", "cbyte", "ccur", "cdate", "cdbl", "cdec", "cint", "clng", "clnglng", "clngptr", "csng",
        "cstr", "cvar", "cvdate", "cverr",
    }
)

_TYPE_PROBLEM_LITERAL_KINDS = frozenset({TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL, TokenKind.STRING_LITERAL})


def _type_argument_problem(
    callee: str,
    param: CallableParamType,
    call: CallArguments,
    slot: Sequence[VbaToken],
    udt_names: AbstractSet[str],
    resolve_shape: _ShapeResolver,
    member_type: _MemberTypeOf,
) -> tuple[str, str, Span] | None:
    """A Type value into a Variant parameter, or anything but a variable of the
    Type into a Type parameter (issue #253). Returns (code, message, span)."""
    toks = _significant(slot)
    if not toks:
        return None
    span = Span(call.slice_start + toks[0].start, call.slice_start + toks[-1].end)
    text = "".join(t.raw_text for t in toks)
    # `((t))`: the parentheses inside the slot make the Type a value.
    wrapped = toks[0].raw_text == "(" and match_paren_from(toks, 0) == len(toks) - 1
    inner = toks[1:-1] if wrapped else toks
    name = token_name(inner[0]) if len(inner) == 1 else None
    shape = resolve_shape(name) if name else None
    variable = (
        shape.shape
        if shape is not None and shape.resolved and shape.shape is not None and not shape.shape.is_array
        else None
    )
    argument_type = type_key(variable.as_type) if variable is not None else None
    is_type = argument_type is not None and argument_type in udt_names
    if param.param_array or (not param.is_array and (normalize_type(param.type_) or "variant") == "variant"):
        fn = callee.lower()
        if not is_type or fn in _LEFT_ALONE or (fn in _CONVERSIONS and fn != "cvar" and fn != "cstr"):
            return None
        if fn == "cstr":
            return (
                "udtValueMismatch",
                f"'{name}' is a user-defined type, which {callee} cannot convert to a String. "
                "This is a VBE compile error: Type mismatch.",
                span,
            )
        return (
            "udtVariantCoercion",
            f"'{name}' is a user-defined type, and parameter '{param.name}' of '{callee}' is a Variant, which "
            "cannot hold one declared in a standard module or a private class. This is a VBE compile error: "
            "Only user-defined types defined in public object modules can be coerced to or from a variant or "
            "passed to late-bound functions.",
            span,
        )
    param_type = type_key(param.type_)
    if param.is_array or not param_type or param_type not in udt_names:
        return None
    if is_type and argument_type == param_type and (wrapped or (call.arguments_parenthesized and len(call.slots) == 1)):
        shown = f"({text})" if call.arguments_parenthesized and not wrapped else text
        return (
            "variableRequired",
            f"'{shown}' is in parentheses, which make the user-defined type a value, and parameter "
            f"'{param.name}' of '{callee}' takes a variable. This is a VBE compile error: Variable required - "
            "can't assign to this expression.",
            Span(call.slice_start + inner[0].start, call.slice_start + inner[-1].end),
        )
    if wrapped:
        return None
    what: str | None
    if variable is not None:
        what = (
            None
            if argument_type == param_type
            else f"'{name}' is declared As {variable.as_type if variable.as_type is not None else 'Variant'}"
        )
    elif len(toks) == 1 and toks[0].kind in _TYPE_PROBLEM_LITERAL_KINDS:
        what = f"{text} is a literal"
    else:
        member = member_type(toks)
        what = (
            f"'{text}' is a member declared As {member.name}"
            if member is not None and member.type != param_type
            else None
        )
    if not what:
        return None
    return (
        "byRefArgumentTypeMismatch",
        f"{what}, but parameter '{param.name}' of '{callee}' is declared As {param.type_}, a user-defined type, "
        "which takes only a variable of that Type. This is a VBE compile error: ByRef argument type mismatch.",
        span,
    )


def _parenthesized_argument(inner: Sequence[VbaToken], slice_start: int) -> tuple[str, Span] | None:
    """The problem with an argument in parentheses, given the tokens inside them."""
    toks = _significant(inner)
    if not toks:
        return None
    return (
        f"'({''.join(t.raw_text for t in toks)})' is in parentheses, which pass a value rather than an array "
        "variable",
        Span(slice_start + toks[0].start, slice_start + toks[-1].end),
    )


def _sole_identifier(slot: Sequence[VbaToken], slice_start: int) -> tuple[str, Span] | None:
    """A single bare identifier argument (not indexed / member / call / expression)."""
    toks = _significant(slot)
    if len(toks) != 1:
        return None
    name = token_name(toks[0])
    if not name:
        return None
    return (name, Span(slice_start + toks[0].start, slice_start + toks[0].end))


def _param_is_known_scalar(param: CallableParamType) -> bool:
    """True when a parameter is a known scalar (excludes Variant, array, object, UDT)."""
    norm = normalize_type(param.type_)
    return norm is not None and is_known_scalar_type(norm)


def _is_scalar_or_variant(as_type: str) -> bool:
    """True when a declared type is a known scalar or `Variant` (the array-param reject set)."""
    norm = normalize_type(as_type)
    return norm is not None and (is_known_scalar_type(norm) or norm == "variant")


def _vbe_scalar_error(param: CallableParamType) -> str:
    return "ByRef argument type mismatch" if param.by_ref else "Type mismatch"


def _array_to_scalar_message(name: str, param: CallableParamType, callee: str) -> str:
    return (
        f"Argument '{name}' is declared as an array, but parameter '{param.name}' of '{callee}' "
        f"expects a scalar {param.type_}. This is a VBE compile error: {_vbe_scalar_error(param)}."
    )


def _udt_to_scalar_message(name: str, as_type: str, param: CallableParamType, callee: str) -> str:
    return (
        f"Argument '{name}' is declared As {as_type} (a user-defined Type), but parameter "
        f"'{param.name}' of '{callee}' expects a scalar {param.type_}. This is a VBE compile "
        f"error: {_vbe_scalar_error(param)}."
    )


def _scalar_to_array_message(name: str, param: CallableParamType, callee: str) -> str:
    return (
        f"Argument '{name}' is a scalar, but parameter '{param.name}' of '{callee}' is declared "
        "as an array. This is a VBE compile error: Type mismatch: array or user-defined type expected."
    )
