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
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence

from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ProcedureNode, Span
from ...symbols.name_resolution import BareIdentifierContext
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureSignature, VbaSymbol
from ...types.type_inference import (
    SourceDeclaredShape,
    SourceDeclaredType,
    declared_shape_for_source_binding,
    declared_value_type_for_qualified_source_binding,
    declared_value_type_for_source_binding,
    procedure_symbol_for,
    same_module_type_names,
    type_environment_for,
)
from ...types.type_names import is_known_scalar_type, normalize_type
from ..argument_inference import byref_variable_type_mismatch
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
from ..walker import ProcedureStatementVisitor, strip_header_brackets, token_name

_ShapeResolver = Callable[[str], SourceDeclaredShape]
_TypeResolver = Callable[[str], SourceDeclaredType]
_QualifiedTypeResolver = Callable[[str, str], SourceDeclaredType]


def check_argument_shape(
    source: str,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    module_signatures = callable_type_signatures_for(symbols, project_procedures)
    udt_names = same_module_type_names(symbols)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
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

        def check_call(call: CallArguments) -> None:
            sig = callable_signature_for_call(call, module_signatures, source_names)
            if sig is None or not sig.params:
                return
            _validate_argument_shapes(
                sig, call, udt_names, env, resolve_type, resolve_qualified_type, resolve_shape, push
            )

        def visitor(stmt: LeafStatementNode) -> None:
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
    udt_names: set[str],
    env: Mapping[str, str],
    resolve_type: _TypeResolver,
    resolve_qualified_type: _QualifiedTypeResolver,
    resolve_shape: _ShapeResolver,
    push: PushFn,
) -> None:
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
        if param is None or param.param_array:
            continue  # ParamArray params are Variant and accept any shape
        # Defer to byref-argument-type-mismatch when it owns this slot.
        if byref_variable_type_mismatch(
            param, value_slot, call.slice_start, env, resolve_type, resolve_qualified_type
        ) is not None:
            continue
        ident = _sole_identifier(value_slot, call.slice_start)
        if ident is None:
            continue
        name, span = ident
        shape = resolve_shape(name)
        if not shape.resolved or shape.shape is None:
            continue
        if shape.shape.is_array:
            if not param.is_array and _param_is_known_scalar(param):
                push("argumentShapeMismatch", _array_to_scalar_message(name, param, sig.name), span)
            continue
        as_type = shape.shape.as_type
        if as_type and as_type.lower() in udt_names:
            if not param.is_array and _param_is_known_scalar(param):
                push(
                    "argumentShapeMismatch",
                    _udt_to_scalar_message(name, as_type, param, sig.name),
                    span,
                )
            continue
        if param.is_array and as_type and _is_scalar_or_variant(as_type):
            push("argumentShapeMismatch", _scalar_to_array_message(name, param, sig.name), span)


def _sole_identifier(slot: list[VbaToken], slice_start: int) -> tuple[str, Span] | None:
    toks = [t for t in slot if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE]
    if len(toks) != 1:
        return None
    name = token_name(toks[0])
    if not name:
        return None
    return (name, Span(slice_start + toks[0].start, slice_start + toks[0].end))


def _param_is_known_scalar(param: CallableParamType) -> bool:
    norm = normalize_type(param.type_)
    return norm is not None and is_known_scalar_type(norm)


def _is_scalar_or_variant(as_type: str) -> bool:
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
