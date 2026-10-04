"""Source-binding shape inference for diagnostics rules.

Ported from the host-free slice of
xlide_vscode/src/analyzer/diagnostics/typeInference.ts (declaredShapeForSourceBinding)
plus procedureSymbolFor from analysisContext.ts. Resolves a bare identifier to the
declared shape (as-type, array-ness, fixed-vs-dynamic) of its source binding using
only the symbol graph. Expression typing through the host model and the member
resolver lives in diagnostics/argument_inference.py.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

from ..identity_cache import IdentityLru
from ..parser.nodes import ProcedureNode, ProcKind
from ..symbols.name_resolution import (
    BareIdentifierContext,
    BareIdentifierResolution,
    BareIdentifierResolutionInput,
    BareIdentifierResolutionScope,
    resolve_bare_identifier_binding,
)
from ..symbols.symbol_model import ModuleSymbols, VbaSymbol, VbaSymbolKind

_PROCEDURE_KINDS = frozenset(
    {
        VbaSymbolKind.SUB,
        VbaSymbolKind.FUNCTION,
        VbaSymbolKind.PROPERTY_GET,
        VbaSymbolKind.PROPERTY_LET,
        VbaSymbolKind.PROPERTY_SET,
    }
)


@dataclass(frozen=True, slots=True)
class DeclaredValueShape:
    as_type: str | None
    is_array: bool
    is_fixed_array: bool


@dataclass(frozen=True, slots=True)
class SourceDeclaredShape:
    resolved: bool
    shape: DeclaredValueShape | None = None


@dataclass(frozen=True, slots=True)
class SourceDeclaredType:
    resolved: bool
    as_type: str | None = None
    # What the resolved binding is: a variable, a constant, a parameter, a procedure.
    kind: VbaSymbolKind | None = None


_VALUE_DECLARATION_KINDS = frozenset(
    {
        VbaSymbolKind.PARAMETER,
        VbaSymbolKind.LOCAL_VARIABLE,
        VbaSymbolKind.MODULE_VARIABLE,
        VbaSymbolKind.CONSTANT,
    }
)


# Environments are requested once per rule per procedure, so both the finished
# per-procedure environments and their procedure-independent module-level bases
# are memoized by identity: rebuilding them per request was O(rules x
# procedures x declarations) on large modules. Consumers treat the returned
# dicts as read-only.
_SHAPE_ENV_CACHE = IdentityLru(capacity=64)
_SHAPE_ENV_MODULE_BASE_CACHE = IdentityLru()
_TYPE_ENV_CACHE = IdentityLru(capacity=64)
_TYPE_ENV_MODULE_BASE_CACHE = IdentityLru()


def _shape_env_module_base(symbols: ModuleSymbols) -> dict[str, DeclaredValueShape]:
    cached = _SHAPE_ENV_MODULE_BASE_CACHE.get(symbols)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    base: dict[str, DeclaredValueShape] = {}
    for sym in symbols.root.children or []:
        if sym.kind in _VALUE_DECLARATION_KINDS:
            base[sym.name.lower()] = _shape_of(sym)
    return _SHAPE_ENV_MODULE_BASE_CACHE.put(base, symbols)  # type: ignore[no-any-return]


def declaration_shape_environment_for(
    symbols: ModuleSymbols, proc: ProcedureNode
) -> dict[str, DeclaredValueShape]:
    """Syntactic declared-shape fallback for a procedure's module + local value names.

    The function-return-variable shape (an Erase/assignment target named after the
    procedure) is omitted; that is a precision-only gap, never a false positive.
    """
    cached = _SHAPE_ENV_CACHE.get(symbols, proc)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    out = dict(_shape_env_module_base(symbols))
    proc_sym = procedure_symbol_for(symbols, proc)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if child.kind in _VALUE_DECLARATION_KINDS:
            out[child.name.lower()] = _shape_of(child)
    return _SHAPE_ENV_CACHE.put(out, symbols, proc)  # type: ignore[no-any-return]


def _shape_of(sym: VbaSymbol) -> DeclaredValueShape:
    return DeclaredValueShape(
        as_type=sym.as_type,
        is_array=sym.is_array is True,
        is_fixed_array=sym.array_bounds is not None,
    )


def same_module_type_names(symbols: ModuleSymbols) -> set[str]:
    """Lowercased names of user-defined Type declarations in this module."""
    return {
        sym.name.lower()
        for sym in (symbols.root.children or [])
        if sym.kind is VbaSymbolKind.TYPE
    }


def _return_assignment_type_for(proc: ProcedureNode) -> str | None:
    if proc.proc_kind in (ProcKind.FUNCTION, ProcKind.PROPERTY_GET):
        return proc.return_type
    return None


def _type_env_module_base(symbols: ModuleSymbols) -> dict[str, str]:
    cached = _TYPE_ENV_MODULE_BASE_CACHE.get(symbols)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    base: dict[str, str] = {}
    for sym in symbols.root.children or []:
        if sym.as_type and sym.kind not in _PROCEDURE_KINDS:
            base[sym.name.lower()] = sym.as_type
    return _TYPE_ENV_MODULE_BASE_CACHE.put(base, symbols)  # type: ignore[no-any-return]


def type_environment_for(symbols: ModuleSymbols, proc: ProcedureNode) -> dict[str, str]:
    """Per-procedure {lowercased name -> raw declared as-type} type environment.

    Module-level typed non-procedure symbols first, then the procedure's own
    return binding, then params/locals last (so a local shadowing a module name
    wins). Values are the raw as-type string (callers normalize at comparison).
    """
    cached = _TYPE_ENV_CACHE.get(symbols, proc)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    out = dict(_type_env_module_base(symbols))
    proc_sym = procedure_symbol_for(symbols, proc)
    return_type = _return_assignment_type_for(proc)
    if return_type:
        out[proc.name.lower()] = return_type
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if child.as_type:
            out[child.name.lower()] = child.as_type
    return _TYPE_ENV_CACHE.put(out, symbols, proc)  # type: ignore[no-any-return]


# Procedure symbols are looked up per procedure per rule; the by-start-offset
# index turns each lookup from an O(module members) scan into a dict hit.
_PROCEDURE_SYMBOL_INDEX_CACHE = IdentityLru()


def _procedure_symbol_index(symbols: ModuleSymbols) -> dict[int, VbaSymbol]:
    cached = _PROCEDURE_SYMBOL_INDEX_CACHE.get(symbols)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    index: dict[int, VbaSymbol] = {}
    for sym in symbols.root.children or []:
        if sym.kind in _PROCEDURE_KINDS and sym.full_span.start not in index:
            index[sym.full_span.start] = sym
    return _PROCEDURE_SYMBOL_INDEX_CACHE.put(index, symbols)  # type: ignore[no-any-return]


def procedure_symbol_for(symbols: ModuleSymbols, proc: ProcedureNode) -> VbaSymbol | None:
    """The module symbol for a procedure node, matched by declaration start offset."""
    return _procedure_symbol_index(symbols).get(proc.span.start)


def declared_shape_for_source_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> SourceDeclaredShape:
    """Resolve a bare identifier to its declared shape via the source symbol graph."""
    binding = resolve_bare_identifier_binding(
        BareIdentifierResolutionInput(
            current_module=symbols,
            name=name,
            context=context,
            enclosing_procedure=proc_sym,
            project_visible_symbols=project_visible_symbols or (),
        )
    )
    if binding.scope in (
        BareIdentifierResolutionScope.UNRESOLVED,
        BareIdentifierResolutionScope.AMBIGUOUS,
    ):
        return SourceDeclaredShape(resolved=binding.scope is BareIdentifierResolutionScope.AMBIGUOUS)
    shaped = next((d for d in binding.definitions if d.as_type or d.is_array), None)
    if shaped is None:
        return SourceDeclaredShape(resolved=True, shape=DeclaredValueShape(None, False, False))
    return SourceDeclaredShape(
        resolved=True,
        shape=DeclaredValueShape(
            as_type=shaped.as_type,
            is_array=shaped.is_array is True,
            is_fixed_array=shaped.array_bounds is not None,
        ),
    )


def is_value_declaration_symbol(sym: VbaSymbol) -> bool:
    return sym.kind in _VALUE_DECLARATION_KINDS


def _source_identifier_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> BareIdentifierResolution:
    return resolve_bare_identifier_binding(
        BareIdentifierResolutionInput(
            current_module=symbols,
            name=name,
            context=context,
            enclosing_procedure=proc_sym,
            project_visible_symbols=project_visible_symbols or (),
        )
    )


def source_identifier_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> BareIdentifierResolution:
    """Resolve a bare identifier to its binding (the public binder seam for rules)."""
    return _source_identifier_binding(symbols, proc_sym, project_visible_symbols, name, context)


def source_identifier_bound(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> bool:
    """Whether a bare identifier resolves to any source binding (scope != UNRESOLVED)."""
    binding = source_identifier_binding(symbols, proc_sym, project_visible_symbols, name, context)
    return binding.scope is not BareIdentifierResolutionScope.UNRESOLVED


def declared_type_for_source_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> SourceDeclaredType:
    binding = _source_identifier_binding(symbols, proc_sym, project_visible_symbols, name, context)
    if binding.scope in (
        BareIdentifierResolutionScope.UNRESOLVED,
        BareIdentifierResolutionScope.AMBIGUOUS,
    ):
        return SourceDeclaredType(resolved=binding.scope is BareIdentifierResolutionScope.AMBIGUOUS)
    typed = next((d for d in binding.definitions if d.as_type), None)
    return SourceDeclaredType(resolved=True, as_type=typed.as_type if typed is not None else None)


def declared_value_type_for_source_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
) -> SourceDeclaredType:
    binding = _source_identifier_binding(
        symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.EXPRESSION
    )
    if binding.scope in (
        BareIdentifierResolutionScope.UNRESOLVED,
        BareIdentifierResolutionScope.AMBIGUOUS,
    ):
        return SourceDeclaredType(resolved=binding.scope is BareIdentifierResolutionScope.AMBIGUOUS)
    value_definitions = [d for d in binding.definitions if is_value_declaration_symbol(d)]
    if not value_definitions:
        return SourceDeclaredType(resolved=False)
    typed = next((d for d in value_definitions if d.as_type), None)
    return SourceDeclaredType(
        resolved=True,
        as_type=typed.as_type if typed is not None else None,
        kind=(typed if typed is not None else value_definitions[0]).kind,
    )


def declared_value_type_for_qualified_source_binding(
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    qualifier: str,
    name: str,
) -> SourceDeclaredType:
    qualifier_lower = qualifier.lower()
    name_lower = name.lower()
    candidates: list[VbaSymbol] = []
    if symbols.module_name.lower() == qualifier_lower:
        candidates.extend(symbols.root.children or [])
    candidates.extend(
        s for s in (project_visible_symbols or []) if s.module_name.lower() == qualifier_lower
    )
    if not candidates:
        return SourceDeclaredType(resolved=False)
    matching_values = [
        s for s in candidates if s.name.lower() == name_lower and is_value_declaration_symbol(s)
    ]
    if not matching_values:
        return SourceDeclaredType(resolved=True)
    typed = next((d for d in matching_values if d.as_type), None)
    return SourceDeclaredType(resolved=True, as_type=typed.as_type if typed is not None else None)


# --- sync stubs (2f49b93): replaced as each group is ported ---


def def_type_of(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("defTypeOf not ported yet")


def constant_string_value(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("constantStringValue not ported yet")


def string_constants_in_scope(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("stringConstantsInScope not ported yet")


def source_binding_type_resolvers(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("sourceBindingTypeResolvers not ported yet")


def type_field_declared_type(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("typeFieldDeclaredType not ported yet")


def with_known_locals(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("withKnownLocals not ported yet")


def infer_bare_external_constant_expression_type(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("inferBareExternalConstantExpressionType not ported yet")


def infer_bare_external_object_expression_type(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("inferBareExternalObjectExpressionType not ported yet")


def infer_qualified_external_constant_expression_type(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("inferQualifiedExternalConstantExpressionType not ported yet")


def inferred_external_constant(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("inferredExternalConstant not ported yet")


def return_assignment_type_for(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("returnAssignmentTypeFor not ported yet")


def return_assignment_is_array(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("returnAssignmentIsArray not ported yet")


def is_property_result_indexing(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("isPropertyResultIndexing not ported yet")


def is_member_parenless_argument_start(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("isMemberParenlessArgumentStart not ported yet")


def sheets_from_collection_property(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("sheetsFromCollectionProperty not ported yet")


VALUE_HELD: object = None


def arithmetic_of_scalars(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("arithmeticOfScalars not ported yet")


def by_ref_variable_type_mismatch(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("byRefVariableTypeMismatch not ported yet")


def is_known_by_ref_exact_type(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("isKnownByRefExactType not ported yet")


def runtime_signature_parameter_text(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("runtimeSignatureParameterText not ported yet")


def parse_runtime_param_type(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("parseRuntimeParamType not ported yet")


def split_signature_top_level(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("splitSignatureTopLevel not ported yet")


def infer_signed_numeric_literal(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("inferSignedNumericLiteral not ported yet")


def infer_atomic_expression_type(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("inferAtomicExpressionType not ported yet")


def infer_intrinsic_cverr_error_variant(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("inferIntrinsicCverrErrorVariant not ported yet")


def member_expression_return_type(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("memberExpressionReturnType not ported yet")


def default_host_item_return_type(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("defaultHostItemReturnType not ported yet")


def final_member_token_in_expression(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("finalMemberTokenInExpression not ported yet")


def matching_open_paren_index(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("matchingOpenParenIndex not ported yet")


def has_top_level_operator(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("hasTopLevelOperator not ported yet")


def member_accepts_zero_arguments(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("memberAcceptsZeroArguments not ported yet")


def parameterless_value_signature(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("parameterlessValueSignature not ported yet")


def infer_arithmetic_expression_type(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("inferArithmeticExpressionType not ported yet")


def find_nonnumeric_string_in_arithmetic_expression(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("findNonnumericStringInArithmeticExpression not ported yet")


def infer_string_concatenation_expression_type(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("inferStringConcatenationExpressionType not ported yet")


def split_top_level_arithmetic_operands(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("splitTopLevelArithmeticOperands not ported yet")


def split_top_level_operands(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("splitTopLevelOperands not ported yet")


def numeric_literal_overflow_reason(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("numericLiteralOverflowReason not ported yet")


def create_object_assignment_type_resolver(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("createObjectAssignmentTypeResolver not ported yet")


def dao_whole_value_error(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("daoWholeValueError not ported yet")


def object_holding_default(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("objectHoldingDefault not ported yet")


def read_only_host_default(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("readOnlyHostDefault not ported yet")


def argumentless_host_default(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("argumentlessHostDefault not ported yet")


def object_value_needs_index(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("objectValueNeedsIndex not ported yet")


def picked_values(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("pickedValues not ported yet")


def known_local_literal_values_at(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("knownLocalLiteralValuesAt not ported yet")


def statement_may_change_module_variable(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("statementMayChangeModuleVariable not ported yet")


def unreachable_statements_in(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("unreachableStatementsIn not ported yet")


def dead_branch_spans_in(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("deadBranchSpansIn not ported yet")


def defaulted_straight_line(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("defaultedStraightLine not ported yet")


def function_result_for(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("functionResultFor not ported yet")


SCALAR_OBJECT_ASSIGNMENT_REASON: object = None


def create_project_interface_sharing_lookup(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("createProjectInterfaceSharingLookup not ported yet")


def implements_object_type(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("implementsObjectType not ported yet")
