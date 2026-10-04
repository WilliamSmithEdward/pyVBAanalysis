"""Source-binding shapes, type environments and the pure typing helpers.

Ported from xlide_vscode/src/analyzer/diagnostics/typeInference.ts (plus
procedureSymbolFor from analysisContext.ts). The port splits that file: this
module holds the type environments and declared shapes, the source-binding
resolvers, and the host-free typing helpers (operand splitters, literal and
overflow typing, signature-text parsing, ByRef exactness, the DAO and library
default-member tables, object-assignment compatibility lookups). The expression
typing engine is diagnostics/argument_inference.py, the signature tables and
call extraction diagnostics/callable_signatures.py, the known local values and
the straight-line walk starts diagnostics/known_locals.py, and member and
object-assignment resolution completion/member_access.py.

Some of upstream's exports live naturally in those modules and would import
this one back; they are re-exported here by thin functions that import their
home module when called, so each keeps the name and module the sync map gives
it.
"""

from __future__ import annotations

import json
import math
import re
from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Protocol, TypeVar

from ..completion.member_access import (
    KnownObjectAssignmentType,
    MemberCompletionContext,
    MemberCompletionEntry,
    ProjectTypeLookup,
    is_explicit_element_accessor,
    library_object_type,
    member_takes_own_arguments,
    resolve_known_object_assignment_type,
    simple_type_name_for_assignment,
)
from ..constants.integer_constant_expression import (
    IntegerConstantLookup,
    bankers_round,
    parse_decimal_integer_literal,
)
from ..host.host_default_members import HOST_DEFAULT_MEMBERS
from ..host.host_model import (
    HostConstant,
    HostMember,
    HostObjectModel,
    get_host_members,
    get_host_type,
    resolve_host_alias,
    resolve_host_constant,
    resolve_host_global,
    resolve_host_global_member,
)
from ..host.type_extensibility import host_type_resolves_when_compiling
from ..identity_cache import IdentityLru
from ..js_compat import js_number, js_number_to_string, js_trim
from ..lexer.token_helpers import match_paren_from
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import ProcedureNode, ProcKind, Span
from ..runtime.vba_runtime import (
    VbaRuntimeConstant,
    resolve_runtime_constant,
    resolve_runtime_function,
    resolve_runtime_object,
)
from ..symbols.name_resolution import (
    BareIdentifierContext,
    BareIdentifierResolution,
    BareIdentifierResolutionInput,
    BareIdentifierResolutionScope,
    resolve_bare_identifier_binding,
)
from ..symbols.symbol_model import ModuleSymbols, VbaProjectClassMembers, VbaSymbol, VbaSymbolKind
from .type_names import (
    is_known_scalar_type,
    is_provably_non_numeric_string,
    normalize_type,
    numeric_literal_bounds,
)

if TYPE_CHECKING:
    from ..conditional import ConditionalActivityTracker
    from ..diagnostics.call_extraction import CallableParamType, CallableTypeSignature, InferredArgumentType
    from ..diagnostics.callable_signatures import SourceNameScope
    from ..diagnostics.known_locals import KnownLocalValue
    from ..diagnostics.straight_line_values import ReachingAssignments
    from ..parser.nodes import BodyNode

_T = TypeVar("_T")

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
    # Whether the binding is an array, whose element type `as_type` then is.
    is_array: bool | None = None
    # A Const's value, when it is one string literal: `Const K = "abc"` (XLIDE #255).
    string_value: str | None = None


_VALUE_DECLARATION_KINDS = frozenset(
    {
        VbaSymbolKind.PARAMETER,
        VbaSymbolKind.LOCAL_VARIABLE,
        VbaSymbolKind.MODULE_VARIABLE,
        VbaSymbolKind.CONSTANT,
    }
)


def _layered(base: Mapping[str, _T], own: Mapping[str, _T]) -> dict[str, _T]:
    """A procedure's view of a module-level table: its own entries over the
    module's, in upstream's LayeredMap order (the module entries the procedure
    does not shadow, then the procedure's)."""
    if not own:
        return dict(base)
    out = {key: value for key, value in base.items() if key not in own}
    out.update(own)
    return out


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
    """The declared shape of each module-level and procedure value name, and of
    the procedure's own return binding (an assignment target named after it)."""
    cached = _SHAPE_ENV_CACHE.get(symbols, proc)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    own: dict[str, DeclaredValueShape] = {}
    proc_sym = procedure_symbol_for(symbols, proc)
    return_type = return_assignment_type_for(proc)
    if return_type:
        own[proc.name.lower()] = DeclaredValueShape(
            as_type=return_type, is_array=return_assignment_is_array(proc), is_fixed_array=False
        )
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if child.kind in _VALUE_DECLARATION_KINDS:
            own[child.name.lower()] = _shape_of(child)
    out = _layered(_shape_env_module_base(symbols), own)
    return _SHAPE_ENV_CACHE.put(out, symbols, proc)  # type: ignore[no-any-return]


def _shape_of(sym: VbaSymbol) -> DeclaredValueShape:
    return DeclaredValueShape(
        as_type=sym.as_type,
        is_array=sym.is_array is True,
        is_fixed_array=sym.array_bounds is not None,
    )


_SAME_MODULE_TYPE_NAMES_CACHE = IdentityLru()


def same_module_type_names(symbols: ModuleSymbols) -> frozenset[str]:
    """Lowercased names of user-defined Type declarations in this module, memoized
    per parse."""
    cached = _SAME_MODULE_TYPE_NAMES_CACHE.get(symbols)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    names = frozenset(
        sym.name.lower() for sym in (symbols.root.children or []) if sym.kind is VbaSymbolKind.TYPE
    )
    return _SAME_MODULE_TYPE_NAMES_CACHE.put(names, symbols)  # type: ignore[no-any-return]


def _type_env_module_base(symbols: ModuleSymbols) -> dict[str, str]:
    cached = _TYPE_ENV_MODULE_BASE_CACHE.get(symbols)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    base: dict[str, str] = {}
    for sym in symbols.root.children or []:
        type_ = sym.as_type or (
            def_type_of(symbols, sym.name) if sym.kind is VbaSymbolKind.MODULE_VARIABLE else None
        )
        if type_ and sym.kind not in _PROCEDURE_KINDS:
            base[sym.name.lower()] = type_
    return _TYPE_ENV_MODULE_BASE_CACHE.put(base, symbols)  # type: ignore[no-any-return]


_DEF_TYPE_NAME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]*$")


def def_type_of(symbols: ModuleSymbols, name: str) -> str | None:
    """The type a DefType line gives a name declared with no type and no type
    character: `DefInt A-Z` then `Dim i` is an Integer (XLIDE issue #285). A
    Variant is no type to report on, so it gives None, as no DefType does."""
    def_types = symbols.def_types
    type_ = (
        def_types.get(name[0].lower())
        if def_types is not None and _DEF_TYPE_NAME_RE.match(name) is not None
        else None
    )
    return None if type_ in ("Variant", "Decimal") else type_


def type_environment_for(symbols: ModuleSymbols, proc: ProcedureNode) -> dict[str, str]:
    """Per-procedure {lowercased name -> raw declared as-type} type environment.

    Module-level typed non-procedure symbols first, then the procedure's own
    return binding, then params/locals last (so a local shadowing a module name
    wins), then the names the procedure assigns undeclared under a DefType line.
    Values are the raw as-type string (callers normalize at comparison).
    """
    cached = _TYPE_ENV_CACHE.get(symbols, proc)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    own: dict[str, str] = {}
    proc_sym = procedure_symbol_for(symbols, proc)
    return_type = return_assignment_type_for(proc) or (
        def_type_of(symbols, proc.name)
        if proc.proc_kind in (ProcKind.FUNCTION, ProcKind.PROPERTY_GET) and not proc.type_suffix
        else None
    )
    if return_type:
        own[proc.name.lower()] = return_type
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        type_ = child.as_type or (
            def_type_of(symbols, child.name)
            if child.kind in (VbaSymbolKind.LOCAL_VARIABLE, VbaSymbolKind.PARAMETER)
            else None
        )
        if type_:
            own[child.name.lower()] = type_
    # A name the procedure assigns with no declaration is a local the DefType
    # types (XLIDE issue #285), unless it is a host global such as StatusBar or a
    # VBA function such as Mid.
    implicit_locals = symbols.implicit_locals
    for name in (implicit_locals.get(proc.span.start) if implicit_locals is not None else None) or ():
        type_ = def_type_of(symbols, name)
        if (
            type_
            and name not in own
            and not resolve_host_global(name, None)
            and resolve_host_global_member(name, None) is None
            and resolve_runtime_function(name) is None
        ):
            own[name] = type_
    out = _layered(_type_env_module_base(symbols), own)
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


def is_value_declaration_symbol(sym: VbaSymbol) -> bool:
    return sym.kind in _VALUE_DECLARATION_KINDS


def constant_string_value(symbol: VbaSymbol) -> str | None:
    """The value of a Const that is one string literal, or None."""
    if symbol.kind is not VbaSymbolKind.CONSTANT or symbol.default_raw is None:
        return None
    from ..diagnostics.call_extraction import string_literal_value
    from ..diagnostics.walker import raw_expression_tokens

    toks = [tok for tok in raw_expression_tokens(symbol.default_raw) if tok.kind is not TokenKind.COMMENT]
    return (
        string_literal_value(toks[0].raw_text)
        if len(toks) == 1 and toks[0].kind is TokenKind.STRING_LITERAL
        else None
    )


def string_constants_in_scope(symbols: ModuleSymbols, proc: ProcedureNode) -> dict[str, str]:
    """The String Consts a procedure sees, each one string literal, by lowercased
    name (XLIDE issue #255). A local or parameter of the same name hides a
    module's."""
    proc_sym = procedure_symbol_for(symbols, proc)
    children = (proc_sym.children if proc_sym is not None else None) or []
    child_ids = {id(child) for child in children}
    out: dict[str, str] = {}
    for symbol in [*(symbols.root.children or []), *children]:
        value = constant_string_value(symbol)
        if value is not None:
            out[symbol.name.lower()] = value
        elif id(symbol) in child_ids:
            out.pop(symbol.name.lower(), None)
    return out


def declared_type_for_source_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> SourceDeclaredType:
    binding = source_identifier_binding(symbols, proc_sym, project_visible_symbols, name, context)
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
    binding = source_identifier_binding(
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
    chosen = typed if typed is not None else value_definitions[0]
    string_value = constant_string_value(chosen) if len(value_definitions) == 1 else None
    return SourceDeclaredType(
        resolved=True,
        as_type=typed.as_type if typed is not None else None,
        kind=chosen.kind,
        is_array=chosen.is_array is True,
        string_value=string_value,
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


@dataclass(frozen=True, slots=True)
class SourceBindingTypeResolvers:
    """The two lookups an expression-typing pass needs, closed over one
    procedure's scope: a bare name's declared value type, and a
    `Qualifier.Name` member's."""

    resolve_expression_type: Callable[[str], SourceDeclaredType]
    resolve_qualified_expression_type: Callable[[str, str], SourceDeclaredType]


def source_binding_type_resolvers(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
) -> SourceBindingTypeResolvers:
    def resolve_expression_type(name: str) -> SourceDeclaredType:
        return declared_value_type_for_source_binding(symbols, proc_sym, project_visible_symbols, name)

    def resolve_qualified_expression_type(qualifier: str, name: str) -> SourceDeclaredType:
        qualified = declared_value_type_for_qualified_source_binding(
            symbols, project_visible_symbols, qualifier, name
        )
        if qualified.resolved:
            return qualified
        field = type_field_declared_type(symbols, proc_sym, project_visible_symbols, qualifier, name)
        return field if field is not None else qualified

    return SourceBindingTypeResolvers(resolve_expression_type, resolve_qualified_expression_type)


def type_field_declared_type(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    variable: str,
    field: str,
) -> SourceDeclaredType | None:
    """The declared type of a field of a user-defined type a variable holds: `t.i`
    with `Dim t As T1` and `i As Integer` in T1 (XLIDE issue #369). An array field
    is left to the shape rules."""
    holder = declared_value_type_for_source_binding(symbols, proc_sym, project_visible_symbols, variable)
    type_name = (
        js_trim(holder.as_type.split(".")[-1]).lower()
        if holder.resolved and holder.as_type is not None
        else None
    )
    if not type_name or holder.kind is VbaSymbolKind.CONSTANT:
        return None
    user_type = next(
        (
            symbol
            for symbol in [*(symbols.root.children or []), *(project_visible_symbols or [])]
            if symbol.kind is VbaSymbolKind.TYPE and symbol.name.lower() == type_name
        ),
        None,
    )
    field_lower = field.lower()
    member = next(
        (
            child
            for child in ((user_type.children if user_type is not None else None) or [])
            if child.kind is VbaSymbolKind.TYPE_FIELD and child.name.lower() == field_lower
        ),
        None,
    )
    return (
        SourceDeclaredType(resolved=True, as_type=member.as_type)
        if member is not None and not member.is_array
        else None
    )


def declared_shape_for_source_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> SourceDeclaredShape:
    """Resolve a bare identifier to its declared shape via the source symbol graph."""
    binding = source_identifier_binding(symbols, proc_sym, project_visible_symbols, name, context)
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


def source_identifier_binding(
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    name: str,
    context: BareIdentifierContext,
) -> BareIdentifierResolution:
    """Resolve a bare identifier to its binding (the public binder seam for rules)."""
    return resolve_bare_identifier_binding(
        BareIdentifierResolutionInput(
            current_module=symbols,
            name=name,
            context=context,
            enclosing_procedure=proc_sym,
            project_visible_symbols=project_visible_symbols or (),
        )
    )


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


# -- known locals as integer constants -------------------------------------


class _WithKnownLocals:
    __slots__ = ("_constants", "_known")

    def __init__(self, constants: IntegerConstantLookup, known: Mapping[str, KnownLocalValue]) -> None:
        self._constants = constants
        self._known = known

    def get(self, name: str, /) -> int | None:
        constant = self._constants.get(name)
        if constant is not None:
            return constant
        local = self._known.get(name.lower())
        if local is None or local.kind != "number":
            return None
        value = local.value
        if isinstance(value, int):
            return value
        return int(value) if isinstance(value, float) and value.is_integer() else None


def with_known_locals(
    constants: IntegerConstantLookup, known: Mapping[str, KnownLocalValue]
) -> IntegerConstantLookup:
    """`constants`, then the whole-number value a local holds at the statement
    (XLIDE issue #238): `zz = 12` then `arr(zz)`, `zz = -1` then `ReDim a(zz)`."""
    return _WithKnownLocals(constants, known)


# -- external constants and objects ----------------------------------------


def infer_bare_external_constant_expression_type(
    name: str,
    span: Span,
    source_names: SourceNameScope | None = None,
    model: HostObjectModel | None = None,
) -> InferredArgumentType | None:
    """A VBA runtime or host constant named bare. A name both define stays unknown."""
    from ..diagnostics.callable_signatures import runtime_callable_source_shadowed

    if runtime_callable_source_shadowed(name, source_names):
        return None
    candidates = [
        candidate
        for candidate in (
            inferred_external_constant(name, resolve_runtime_constant(name)),
            inferred_external_constant(name, resolve_host_constant(name, model)),
        )
        if candidate is not None
    ]
    if len(candidates) != 1:
        return None
    return replace(candidates[0], span=span)


def infer_bare_external_object_expression_type(
    name: str,
    span: Span,
    source_names: SourceNameScope | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> InferredArgumentType | None:
    """A host global (`Application`, `ActiveCell`) or runtime object (`Err`) named
    bare, unless source declares the name."""
    from ..diagnostics.call_extraction import InferredArgumentType
    from ..diagnostics.callable_signatures import runtime_callable_source_shadowed

    if runtime_callable_source_shadowed(name, source_names):
        return None
    host_type = resolve_host_global(name, member_ctx.model if member_ctx is not None else None)
    if host_type:
        return InferredArgumentType(type_=host_type, label=f"{name} As {host_type}", span=span)
    runtime_object = resolve_runtime_object(name)
    if runtime_object is not None:
        runtime_type = runtime_object.get("type", "")
        return InferredArgumentType(type_=runtime_type, label=f"{name} As {runtime_type}", span=span)
    return None


# Qualifiers that name a host's own constant library, resolved against the CURRENT
# host's model: `Word.wdRed` answers in a Word module and misses in an Excel one
# (XLIDE issue #24).
_HOST_CONSTANT_QUALIFIERS = frozenset({"excel", "word", "powerpoint", "access", "office"})


def infer_qualified_external_constant_expression_type(
    qualifier: str,
    name: str,
    span: Span,
    model: HostObjectModel | None = None,
) -> InferredArgumentType | None:
    lower = qualifier.lower()
    if lower == "vba":
        inferred = inferred_external_constant(f"{qualifier}.{name}", resolve_runtime_constant(name))
        return replace(inferred, span=span) if inferred is not None else None
    if lower in _HOST_CONSTANT_QUALIFIERS:
        inferred = inferred_external_constant(f"{qualifier}.{name}", resolve_host_constant(name, model))
        return replace(inferred, span=span) if inferred is not None else None
    return None


def inferred_external_constant(
    display_name: str, constant: VbaRuntimeConstant | HostConstant | None
) -> InferredArgumentType | None:
    from ..diagnostics.call_extraction import InferredArgumentType
    from ..diagnostics.const_expr import numeric_external_constant_value

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


# -- return bindings -------------------------------------------------------


def return_assignment_type_for(proc: ProcedureNode) -> str | None:
    if proc.proc_kind in (ProcKind.FUNCTION, ProcKind.PROPERTY_GET):
        return proc.return_type
    return None


_EMPTY_PARENS_END_RE = re.compile(r"\(\s*\)\s*$")


def return_assignment_is_array(proc: ProcedureNode) -> bool:
    return _EMPTY_PARENS_END_RE.search(return_assignment_type_for(proc) or "") is not None


# -- member calls ----------------------------------------------------------


def is_property_result_indexing(
    member: MemberCompletionEntry, signature: CallableTypeSignature, inner: Sequence[VbaToken]
) -> bool:
    """`obj.Items(1)` on a parameterless property indexes its result; it is no call."""
    return member.kind == "property" and not signature.params and len(inner) > 0


_PARENLESS_ARGUMENT_KINDS = frozenset(
    {
        TokenKind.IDENTIFIER,
        TokenKind.KEYWORD,
        TokenKind.BRACKETED_IDENTIFIER,
        TokenKind.STRING_LITERAL,
        TokenKind.DATE_LITERAL,
        TokenKind.INTEGER_LITERAL,
        TokenKind.FLOAT_LITERAL,
    }
)


def is_member_parenless_argument_start(tok: VbaToken) -> bool:
    return tok.kind in _PARENLESS_ARGUMENT_KINDS or tok.raw_text in (",", "+", "-")


# -- Worksheets and Charts as Sheets ---------------------------------------

# The workbooks a Worksheets or Charts property is read from: `ThisWorkbook.Worksheets`.
_SHEETS_RECEIVERS = frozenset({"application", "thisworkbook", "activeworkbook"})


@dataclass(frozen=True, slots=True)
class SheetsFromCollection:
    text: str
    # "Worksheets" | "Charts"
    collection: str


def sheets_from_collection_property(
    value: Sequence[VbaToken],
    expected: str | None,
    source_names: SourceNameScope,
    member_ctx: MemberCompletionContext,
) -> SheetsFromCollection | None:
    """Excel's Worksheets and Charts properties, bare, on Application or a workbook,
    or given an Array of names, hand back a Sheets object (XLIDE issue #404):
    TypeName(Worksheets) is "Sheets", and Set into a variable As Worksheets or As
    Charts raises 13. This names the property when the target is one of those two
    collections."""
    from ..diagnostics.walker import token_text

    alias = resolve_host_alias(expected or "", member_ctx.model)
    target = alias.lower() if alias else None
    collection = "Worksheets" if target == "excel.worksheets" else "Charts" if target == "excel.charts" else None
    if collection is None or len(value) == 0:
        return None
    # `Worksheets(Array("Sheet1"))` is a Sheets object too; `Worksheets(1)` is a sheet.
    end = len(value) - 1
    if value[end].raw_text == ")":
        open_index = next(
            (
                i
                for i, tok in enumerate(value)
                if tok.raw_text == "(" and match_paren_from(value, i) == end
            ),
            -1,
        )
        if (
            open_index < 1
            or token_text(value[open_index + 1] if open_index + 1 < len(value) else None) != "array"
            or open_index + 2 >= len(value)
            or value[open_index + 2].raw_text != "("
            or match_paren_from(value, open_index + 2) != end - 1
        ):
            return None
        end = open_index - 1
    property_ = token_text(value[end])
    if property_ not in ("worksheets", "charts"):
        return None
    qualifier = value[:end]
    bare = len(qualifier) == 0 and property_ not in source_names.runtime_shadows
    on_workbook = (
        len(qualifier) == 2 and qualifier[1].raw_text == "." and token_text(qualifier[0]) in _SHEETS_RECEIVERS
    )
    return SheetsFromCollection("".join(tok.raw_text for tok in value), collection) if bare or on_workbook else None


# What a held-class lookup gives for a Variant known to hold Empty, a number or a
# String: no object.
VALUE_HELD = "(value)"

_ARITHMETIC_OPERATORS = frozenset({"+", "-", "*", "/", "\\", "^", "mod"})


def arithmetic_of_scalars(toks: Sequence[VbaToken], env: Mapping[str, str]) -> bool:
    """Whether the tokens are an arithmetic expression of number literals and
    locals declared a scalar type, `b + 0` or `d * 2`, with no call, member or
    parenthesis in it: its value is a scalar (XLIDE issue #647)."""
    from ..diagnostics.walker import token_name, token_text

    operand = True
    operators = 0
    for tok in toks:
        if operand:
            name = token_name(tok)
            lower = name.lower() if name else None
            declared = (
                normalize_type(env.get(lower))
                if lower is not None and tok.kind is TokenKind.IDENTIFIER
                else None
            )
            if (
                tok.kind is not TokenKind.INTEGER_LITERAL
                and tok.kind is not TokenKind.FLOAT_LITERAL
                and not (declared and is_known_scalar_type(declared))
            ):
                return False
        elif (token_text(tok) or tok.raw_text) not in _ARITHMETIC_OPERATORS:
            return False
        else:
            operators += 1
        operand = not operand
    return not operand and operators > 0


# -- ByRef exactness -------------------------------------------------------


@dataclass(frozen=True, slots=True)
class ByRefMismatch:
    name: str
    actual: str
    span: Span


# The bindings a ByRef argument passes as the variable itself, not a copy.
_BYREF_VARIABLE_KINDS = frozenset(
    {VbaSymbolKind.LOCAL_VARIABLE, VbaSymbolKind.MODULE_VARIABLE, VbaSymbolKind.PARAMETER}
)


def by_ref_variable_type_mismatch(
    param: CallableParamType,
    slot: Sequence[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    resolve_expression_type: Callable[[str], SourceDeclaredType] | None = None,
    resolve_qualified_expression_type: Callable[[str, str], SourceDeclaredType] | None = None,
) -> ByRefMismatch | None:
    from ..diagnostics.walker import token_name

    if not param.by_ref or not param.type_:
        return None
    expected = normalize_type(param.type_)
    if not _by_ref_exact(expected):
        return None
    toks = [t for t in slot if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE]
    name: str | None
    actual_raw: str | None
    span: Span
    if len(toks) == 1:
        name = token_name(toks[0])
        if not name:
            return None
        declared_type = resolve_expression_type(name) if resolve_expression_type else None
        # A Const is passed as a temporary copy, so its type never has to match
        # (XLIDE issue #111: `Take(K)` with K an Integer Const compiles).
        if declared_type is not None and declared_type.resolved and declared_type.kind is VbaSymbolKind.CONSTANT:
            return None
        # A non-array into an array parameter is argument-shape-mismatch's.
        if param.is_array and declared_type is not None and declared_type.resolved and not declared_type.is_array:
            return None
        actual_raw = (
            declared_type.as_type
            if declared_type is not None and declared_type.resolved
            else env.get(name.lower())
        )
        span = Span(slice_start + toks[0].start, slice_start + toks[0].end)
        # A VARIABLE declared Variant (or with no type) passed ByRef to a typed
        # parameter is the compile error itself (XLIDE issue #111): the VBE refuses
        # `Take v` with `Dim v As Variant` for `x As Long`, `x As String` and
        # `x As Object` alike. Only a variable or parameter: a parameterless
        # Function's name here is a call result, which passes as a copy.
        if (
            not param.is_array
            and declared_type is not None
            and declared_type.resolved
            and declared_type.kind in _BYREF_VARIABLE_KINDS
            and (normalize_type(declared_type.as_type) or "variant") == "variant"
        ):
            return ByRefMismatch(
                name=name,
                actual=declared_type.as_type if declared_type.as_type is not None else "Variant",
                span=span,
            )
    elif (
        len(toks) >= 4
        and toks[1].raw_text == "("
        and toks[-1].raw_text == ")"
        and match_paren_from(toks, 1) == len(toks) - 1
    ):
        # An element of an array variable passes ByRef as the variable itself
        # does: `Take a(1)` with `Dim a(1) As Variant` for `n As Long` is "ByRef
        # argument type mismatch" (XLIDE issue #216, measured in Excel 16.0).
        name = token_name(toks[0])
        declared_type = resolve_expression_type(name) if name and resolve_expression_type else None
        if (
            not name
            or declared_type is None
            or not declared_type.resolved
            or not declared_type.is_array
            or declared_type.kind not in _BYREF_VARIABLE_KINDS
        ):
            return None
        span = Span(slice_start + toks[0].start, slice_start + toks[-1].end)
        element = normalize_type(declared_type.as_type) or "variant"
        if element == "variant" and not param.is_array:
            return ByRefMismatch(name=f"{name}(...)", actual="Variant", span=span)
        actual_raw = declared_type.as_type
        name = f"{name}(...)"
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
    if not _by_ref_exact(actual) or _same_by_ref_type(actual, expected):
        return None
    # An Object passes ByRef to `c As Collection`, and a Collection to `o As
    # Object`: both compile, and the wrong object raises 13 at run time (XLIDE
    # issue #343, measured in Excel 16.0).
    if "object" in (actual, expected) and "collection" in (actual, expected):
        return None
    return ByRefMismatch(name=name, actual=actual_raw if actual_raw is not None else name, span=span)


def _by_ref_exact(type_: str | None) -> bool:
    """A type a ByRef argument must match exactly: a scalar, Object, or a
    Collection (XLIDE issue #410)."""
    return is_known_by_ref_exact_type(type_) or type_ == "collection"


def _same_by_ref_type(actual: str | None, expected: str | None) -> bool:
    """Whether a ByRef argument's type is the parameter's. LongPtr is LongLong in
    64-bit Office, the platform the analyzer assumes."""

    def widen(type_: str | None) -> str | None:
        return "longlong" if type_ == "longptr" else type_

    return widen(actual) == widen(expected)


def is_known_by_ref_exact_type(type_: str | None) -> bool:
    if not type_ or type_ == "variant":
        return False
    return type_ == "object" or is_known_scalar_type(type_)


# -- runtime display signatures --------------------------------------------


def runtime_signature_parameter_text(signature: str) -> str | None:
    """The text of a display signature's parameter list: from its first `(` to the
    `)` that closes it.

    Not to the LAST `)`. A signature can go on past its parameter list with a
    return type that has parentheses of its own, `Values() As Long()` for a Function
    returning an array, and reading to the last one took `) As Long(` for the
    parameters: one required parameter named `As`. A `)` inside a quoted default
    value closes nothing.
    """
    open_index = signature.find("(")
    if open_index < 0:
        return None
    depth = 0
    quoted = False
    for i in range(open_index, len(signature)):
        ch = signature[i]
        if ch == '"':
            quoted = not quoted
        elif quoted:
            continue
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                return signature[open_index + 1 : i]
    return None


_LEADING_BRACKET = re.compile(r"^\[")
_TRAILING_BRACKET = re.compile(r"\]$")
_PARAM_ARRAY = re.compile(r"^ParamArray\b", re.IGNORECASE | re.ASCII)
_PARAM_ARRAY_PREFIX = re.compile(r"^ParamArray\b\s*", re.IGNORECASE | re.ASCII)
_PASSING_PREFIX = re.compile(r"^(?:ByVal|ByRef)\b\s*", re.IGNORECASE | re.ASCII)
_DEFAULT_SUFFIX = re.compile(r"\s*=\s*.*$")
# `\bAs\s+` with JavaScript's ASCII word boundary.
_AS_KEYWORD = re.compile(r"(?<![A-Za-z0-9_])As\s+", re.IGNORECASE)


def _identifier_at(text: str, start: int) -> str | None:
    """The identifier (`[\\p{L}_][\\p{L}\\p{M}\\p{N}_]*`) starting at `start`."""
    from ..lexer.token_helpers import is_identifier

    if start >= len(text) or not is_identifier(text[start]):
        return None
    end = start + 1
    while end < len(text) and is_identifier("_" + text[end]):
        end += 1
    return text[start:end]


def _first_identifier(text: str) -> str | None:
    for i in range(len(text)):
        found = _identifier_at(text, i)
        if found is not None:
            return found
    return None


def parse_runtime_param_type(raw: str) -> CallableParamType | None:
    from ..diagnostics.call_extraction import CallableParamType

    text = js_trim(raw)
    if not text:
        return None
    optional = text.startswith("[") and text.endswith("]")
    text = js_trim(_TRAILING_BRACKET.sub("", _LEADING_BRACKET.sub("", text)))
    param_array = _PARAM_ARRAY.match(text) is not None
    text = _PARAM_ARRAY_PREFIX.sub("", text, count=1)
    text = _PASSING_PREFIX.sub("", text, count=1)
    text = js_trim(_DEFAULT_SUFFIX.sub("", text, count=1))
    as_type: str | None = None
    for as_match in _AS_KEYWORD.finditer(text):
        identifier = _identifier_at(text, as_match.end())
        if identifier is not None:
            after = as_match.end() + len(identifier)
            as_type = identifier + ("()" if text.startswith("()", after) else "")
            break
    first = _first_identifier(text)
    if first is None:
        return None
    return CallableParamType(name=first, type_=as_type, optional=optional, param_array=param_array)


def split_signature_top_level(text: str) -> list[str]:
    """A parameter list split at its top-level commas. Quoted text is opaque, as in
    runtime_signature_parameter_text: a default of `")"` read as a bracket left the
    depth unbalanced, so `F([s As String = ")"], [n As Long])` came out as one
    parameter."""
    out: list[str] = []
    depth = 0
    start = 0
    quoted = False
    for i, ch in enumerate(text):
        if ch == '"':
            quoted = not quoted
        elif quoted:
            continue
        elif ch in ("(", "["):
            depth += 1
        elif ch in (")", "]"):
            depth -= 1
        elif ch == "," and depth == 0:
            out.append(text[start:i])
            start = i + 1
    out.append(text[start:])
    return out


# -- literals and operand splitting ----------------------------------------

_FLOAT_SUFFIX_RE = re.compile(r"[!#@]$")


def float_literal_value(raw: str) -> float:
    """Number() of a float literal's text without its type character: NaN or
    Infinity where JavaScript gives one (a `1D3` exponent is not a number)."""
    return js_number(_FLOAT_SUFFIX_RE.sub("", raw))


def infer_signed_numeric_literal(toks: Sequence[VbaToken], slice_start: int) -> InferredArgumentType | None:
    from ..diagnostics.call_extraction import InferredArgumentType

    if len(toks) != 2 or toks[0].kind is not TokenKind.OPERATOR:
        return None
    sign = toks[0].raw_text
    if sign not in ("+", "-"):
        return None
    literal = toks[1]
    if literal.kind is TokenKind.FLOAT_LITERAL:
        # `-2147483649#` overflows a Long as its unsigned twin does (XLIDE #223).
        magnitude = float_literal_value(literal.raw_text)
        if not math.isfinite(magnitude):
            return None
        return InferredArgumentType(
            type_="Double",
            label=f"numeric literal {sign}{literal.raw_text}",
            span=Span(slice_start + toks[0].start, slice_start + literal.end),
            numeric_text=f"{sign}{literal.raw_text}",
            float_value=-magnitude if sign == "-" else magnitude,
        )
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


def infer_intrinsic_cverr_error_variant(
    toks: Sequence[VbaToken],
    slice_start: int,
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None = None,
) -> InferredArgumentType | None:
    from ..diagnostics.call_extraction import InferredArgumentType, split_arg_slots
    from ..diagnostics.callable_signatures import (
        bare_callable_source_shadowed,
        runtime_callable_source_shadowed,
    )
    from ..diagnostics.walker import span_for_tokens, token_name

    first_name = token_name(toks[0]) if toks else None
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
    inner = list(toks[paren_index + 1 : close])
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


def split_top_level_arithmetic_operands(toks: Sequence[VbaToken]) -> list[list[VbaToken]]:
    parts = split_top_level_operands(toks, "+", "-", "*", "/", "\\", "^")
    return parts if len(parts) >= 2 else []


def split_top_level_operands(toks: Sequence[VbaToken], *operators: str) -> list[list[VbaToken]]:
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
        parts.append(list(toks[start:i]))
        start = i + 1
    if not parts:
        return []
    parts.append(list(toks[start:]))
    return parts


def find_nonnumeric_string_in_arithmetic_expression(
    toks: Sequence[VbaToken], slice_start: int
) -> InferredArgumentType | None:
    from ..diagnostics.call_extraction import (
        InferredArgumentType,
        string_literal_value,
        unwrap_outer_parens,
    )

    current = list(toks)
    while True:
        unwrapped = unwrap_outer_parens(current)
        if len(unwrapped) == len(current):
            break
        current = unwrapped
    parts = split_top_level_arithmetic_operands(current)
    if len(parts) < 2:
        return None
    for part in parts:
        nested = find_nonnumeric_string_in_arithmetic_expression(part, slice_start)
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


# -- overflow --------------------------------------------------------------


def _number_text(value: float) -> str:
    """String(value) for a JavaScript number."""
    return js_number_to_string(value)


def numeric_literal_overflow_reason(expected: str, actual: InferredArgumentType) -> str | None:
    if actual.numeric_value is None:
        return _float_literal_overflow_reason(expected, actual)
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
            f"({_number_text(actual.numeric_value)}) is outside the {bounds.label} range "
            f"{bounds.min} to {bounds.max}. This will raise Run-time error '6': Overflow."
        )
    literal = actual.numeric_text if actual.numeric_text is not None else _number_text(actual.numeric_value)
    return (
        f"{_held_or_literal(actual, literal)} is outside the {bounds.label} range "
        f"{bounds.min} to {bounds.max}. This will raise Run-time error '6': Overflow."
    )


_TWO_POW_63 = float(2**63)


def _float_literal_overflow_reason(expected: str, actual: InferredArgumentType) -> str | None:
    """A float literal passed to a Byte, Integer or Long: VBA rounds it half to even
    and raises 6 when the result is out of range (XLIDE issue #203). Currency keeps
    four decimal places and is left out, as for whole numbers."""
    float_value = actual.float_value
    shown_text = actual.numeric_text if actual.numeric_text is not None else None
    if float_value is not None and expected in ("longlong", "longptr"):
        # Past LongLong's range a LongPtr overflows too, whatever its width (#232).
        rounded = bankers_round(float_value)
        if -_TWO_POW_63 <= rounded < _TWO_POW_63:
            return None
        label = "LongPtr" if expected == "longptr" else "LongLong"
        at_most = ", at most" if expected == "longptr" else ""
        text = shown_text if shown_text is not None else _number_text(float_value)
        return (
            f"{_held_or_literal(actual, text)} is outside the {label} range{at_most} "
            "-9223372036854775808 to 9223372036854775807. This will raise Run-time error "
            "'6': Overflow."
        )
    if float_value is None or expected not in ("byte", "integer", "long"):
        return None
    bounds = numeric_literal_bounds(expected)
    assert bounds is not None
    rounded = bankers_round(float_value)
    if bounds.min <= rounded <= bounds.max:
        return None
    shown = "" if rounded == float_value else f", which VBA rounds to {_number_text(rounded)},"
    text = shown_text if shown_text is not None else _number_text(float_value)
    return (
        f"{_held_or_literal(actual, text)}{shown} is outside the {bounds.label} range "
        f"{bounds.min} to {bounds.max}. This will raise Run-time error '6': Overflow."
    )


def _held_or_literal(actual: InferredArgumentType, text: str) -> str:
    """"The numeric literal 5", or "The value 5 that 'a' holds here" for a local's
    known value."""
    return (
        f"The value {text} that '{actual.held_by}' holds here"
        if actual.held_by is not None
        else f"The numeric literal {text}"
    )


def js_string_literal(value: str) -> str:
    """JSON.stringify of a string, as upstream's messages quote one."""
    return json.dumps(value, ensure_ascii=False)


# -- member expressions ----------------------------------------------------


@dataclass(frozen=True, slots=True)
class FinalMemberToken:
    name: str
    token: VbaToken
    called: bool
    argument_tokens: Sequence[VbaToken] | None = None


def final_member_token_in_expression(toks: Sequence[VbaToken]) -> FinalMemberToken | None:
    from ..diagnostics.walker import token_name

    if not toks:
        return None
    last = toks[-1]
    last_name = token_name(last)
    if last_name and len(toks) >= 2 and toks[-2].raw_text == ".":
        return FinalMemberToken(last_name, last, called=False)
    if last.raw_text != ")":
        return None
    open_index = matching_open_paren_index(toks, len(toks) - 1)
    if open_index < 2:
        return None
    member = toks[open_index - 1]
    member_name = token_name(member)
    if not member_name or toks[open_index - 2].raw_text != ".":
        return None
    return FinalMemberToken(member_name, member, called=True, argument_tokens=toks[open_index + 1 : -1])


def matching_open_paren_index(toks: Sequence[VbaToken], close: int) -> int:
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


def has_top_level_operator(toks: Sequence[VbaToken]) -> bool:
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


def member_accepts_zero_arguments(member: MemberCompletionEntry) -> bool:
    from ..diagnostics.call_extraction import callable_accepts_zero_arguments
    from ..diagnostics.callable_signatures import parse_runtime_display_signature

    if not member.signature:
        return False
    return callable_accepts_zero_arguments(parse_runtime_display_signature(member.name, member.signature))


def parameterless_value_signature(
    name: str,
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None = None,
) -> CallableTypeSignature | None:
    from ..diagnostics.call_extraction import callable_accepts_zero_arguments
    from ..diagnostics.callable_signatures import callable_signature_for

    sig = callable_signature_for(name, module_signatures, source_names)
    if sig is not None and sig.return_type and callable_accepts_zero_arguments(sig):
        return sig
    return None


def member_expression_return_type(
    member: MemberCompletionEntry,
    argument_tokens: Sequence[VbaToken] | None,
    member_ctx: MemberCompletionContext,
) -> str:
    # Calling a member with arguments indexes into it. When the member returns a
    # host collection (one whose Item resolves to an element type), the call yields
    # that element: ws.ChartObjects(1) is a ChartObject, not the collection. A
    # concrete-typed call keeps its declared type (ws.Range("A1") stays Range), and
    # Item/_Default/Add already return the resolved element. A member that takes
    # an argument of its own is what it returns: Shapes.Range(Array("A")) is a
    # ShapeRange, not a Shape (XLIDE issue #197).
    if (
        member.returns
        and argument_tokens
        and not is_explicit_element_accessor(member.name)
        and not member_takes_own_arguments(member.signature)
    ):
        # VBA's Collection holds Variants: `acc.children(i)` may be anything.
        if normalize_type(member.returns) == "collection":
            return "Variant"
        element = default_host_item_return_type(member.returns, member_ctx)
        return element if element is not None else member.returns
    return member.returns if member.returns else "Variant"


_AS_OBJECT_SIGNATURE_RE = re.compile(r"\bAs Object\s*$", re.IGNORECASE | re.ASCII)


def default_host_item_return_type(type_name: str, member_ctx: MemberCompletionContext) -> str | None:
    members = get_host_members(type_name, member_ctx.model)
    item = next((m for m in members if m["name"].lower() == "item"), None)
    if item is not None and item.get("returns"):
        # The library declares most Item accessors `As Object` and the model
        # repairs the type from the reference prose, which is right for
        # completion and chaining. It is not a compile-time binding: the VBE
        # compiles `Worksheets(1).NoSuchMember` (XLIDE issue #114), so the item's
        # members are late bound. A one-part union carries the type without
        # closing it. The hand-written collections carry the repaired type on
        # Item, so the library's word is read off `_Default` too.
        default_member = next((m for m in members if m["name"] == "_Default"), None)
        declared_object = any(
            candidate is not None
            and (
                candidate.get("declaredType") == "Object"
                or _AS_OBJECT_SIGNATURE_RE.search(candidate.get("signature") or "") is not None
            )
            for candidate in (item, default_member)
        )
        return f"union:{item['returns']}" if declared_object else item["returns"]
    # A mixed-element collection (Sheets, whose Item is a Worksheet OR a Chart)
    # carries returnsAnyOf instead. Its indexed element is a late-bound Object.
    if item is not None and item.get("returnsAnyOf"):
        return "Object"
    return None


# -- expression typing (diagnostics/argument_inference.py) ------------------


def infer_atomic_expression_type(
    toks: list[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None = None,
    resolve_expression_type: Callable[[str], SourceDeclaredType] | None = None,
    resolve_qualified_expression_type: Callable[[str, str], SourceDeclaredType] | None = None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> InferredArgumentType | None:
    from ..diagnostics import argument_inference

    return argument_inference.infer_atomic_expression_type(
        toks, slice_start, env, module_signatures, source_names,
        resolve_expression_type, resolve_qualified_expression_type,
        source=source, member_ctx=member_ctx,
    )


def infer_arithmetic_expression_type(
    toks: list[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None = None,
    resolve_expression_type: Callable[[str], SourceDeclaredType] | None = None,
    resolve_qualified_expression_type: Callable[[str, str], SourceDeclaredType] | None = None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> InferredArgumentType | None:
    from ..diagnostics import argument_inference

    return argument_inference.infer_arithmetic_expression_type(
        toks, slice_start, env, module_signatures, source_names,
        resolve_expression_type, resolve_qualified_expression_type,
        source=source, member_ctx=member_ctx,
    )


def infer_string_concatenation_expression_type(
    toks: list[VbaToken],
    slice_start: int,
    env: Mapping[str, str],
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None = None,
    resolve_expression_type: Callable[[str], SourceDeclaredType] | None = None,
    resolve_qualified_expression_type: Callable[[str, str], SourceDeclaredType] | None = None,
    *,
    source: str | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> InferredArgumentType | None:
    from ..diagnostics import argument_inference

    return argument_inference.infer_string_concatenation_expression_type(
        toks, slice_start, env, module_signatures, source_names,
        resolve_expression_type, resolve_qualified_expression_type,
        source=source, member_ctx=member_ctx,
    )


# -- object values ----------------------------------------------------------


class ObjectAssignmentTypeResolver(Protocol):
    """What create_object_assignment_type_resolver returns. Upstream calls it with
    the type alone; the member context is accepted so it also fits
    type_of_is.ObjectTypeResolver, and is ignored."""

    def __call__(
        self, type_name: str | None, member_ctx: MemberCompletionContext | None = None, /
    ) -> KnownObjectAssignmentType | None: ...


def create_object_assignment_type_resolver(
    member_ctx: MemberCompletionContext,
) -> ObjectAssignmentTypeResolver:
    """Resolve object types against metadata stable for one analysis pass. Project
    indexing stays lazy behind the normal host/library resolution priority. A
    fresh resolver observes metadata changes on the next pass.

    The resolver takes (type_name, member_ctx=None), the shape of
    resolve_known_object_assignment_type, so either can be passed where a rule
    takes a resolver; the context it was made for is the one it uses."""
    queries: dict[str | None, KnownObjectAssignmentType | None] = {}
    project_types: dict[str, VbaProjectClassMembers | None] | None = None

    def project_type(lower: str) -> VbaProjectClassMembers | None:
        nonlocal project_types
        if project_types is None:
            project_types = {}
            for candidate in member_ctx.project_class_members or []:
                if candidate.kind in ("userType", "enum", "standardModule"):
                    continue
                name = candidate.name.lower()
                # Every eligible duplicate makes the name ambiguous, even the same
                # object twice.
                project_types[name] = None if name in project_types else candidate
        return project_types.get(lower)

    lookup: ProjectTypeLookup = project_type

    def resolve(
        type_name: str | None, _member_ctx: MemberCompletionContext | None = None
    ) -> KnownObjectAssignmentType | None:
        # Raw keys preserve the display spelling used for generic and host types.
        if type_name not in queries:
            queries[type_name] = resolve_known_object_assignment_type(type_name, member_ctx, lookup)
        return queries[type_name]

    return resolve


# The run-time error DAO raises where an object of this type is read whole as a
# value, `v = rs`: its default member, or the one that holds, needs an index DAO
# checks for itself. Measured in Access 16.0 with the database held (XLIDE issue
# #464); a type not measured is None.
_DAO_WHOLE_VALUE_ERRORS: Mapping[str, str] = {
    "DAO.Database": "'3001': Invalid argument",
    "DAO.Fields": "'3001': Invalid argument",
    "DAO.Properties": "'3001': Invalid argument",
    "DAO.QueryDefs": "'3001': Invalid argument",
    "DAO.Recordset": "'3001': Invalid argument",
    "DAO.Recordset2": "'3001': Invalid argument",
    "DAO.TableDef": "'450': Wrong number of arguments or invalid property assignment",
    "DAO.TableDefs": "'3001': Invalid argument",
    "DAO.Workspace": "'3001': Invalid argument",
}


def dao_whole_value_error(type_name: str | None) -> str | None:
    library = library_object_type(type_name)
    return _DAO_WHOLE_VALUE_ERRORS.get(library) if library else None


def _library_key(type_name: str | None, member_ctx: MemberCompletionContext) -> str:
    return (
        resolve_host_alias(type_name or "", member_ctx.model)
        or library_object_type(type_name)
        or type_name
        or ""
    )


@dataclass(frozen=True, slots=True)
class ObjectHoldingDefault:
    name: str
    returns: str


def object_holding_default(
    type_name: str | None, member_ctx: MemberCompletionContext
) -> ObjectHoldingDefault | None:
    """A Word, PowerPoint or Access type whose default member is a property that
    holds an object, as a Paragraph's Range does: its name and type. Read whole into
    a Variant it gives that object's value, but the VBE refuses a Let to it, an
    operator on it and a Let of it into a typed value while compiling (XLIDE issue
    #462, measured in Word 16.0)."""
    found = HOST_DEFAULT_MEMBERS.get(_library_key(type_name, member_ctx))
    return (
        ObjectHoldingDefault(found.name, found.returns)
        if found is not None
        and found.kind == "property"
        and found.required == 0
        and found.returns in HOST_DEFAULT_MEMBERS
        else None
    )


def read_only_host_default(type_name: str | None, member_ctx: MemberCompletionContext) -> str | None:
    """A Word, PowerPoint or Access type whose default member is a property no Let
    reaches, as a Document's Name: `x = 5` does not compile, "Invalid use of
    property" (XLIDE issue #438). One holding an object is object_holding_default's."""
    found = HOST_DEFAULT_MEMBERS.get(_library_key(type_name, member_ctx))
    return (
        found.name
        if found is not None
        and found.kind == "property"
        and found.required == 0
        and not found.writable
        and found.returns not in HOST_DEFAULT_MEMBERS
        else None
    )


def argumentless_host_default(type_name: str | None, member_ctx: MemberCompletionContext) -> str | None:
    """The default members a Word, PowerPoint or Access type reaches through,
    `Range.Text` for a Paragraph, when none of them takes an argument: then `x(1)`
    does not compile, "Wrong number of arguments or invalid property assignment"
    (XLIDE issue #438)."""
    key = _library_key(type_name, member_ctx)
    names: list[str] = []
    for _depth in range(4):
        found = HOST_DEFAULT_MEMBERS.get(key)
        if found is None or found.kind != "property" or found.params > 0:
            return None
        names.append(found.name)
        if found.returns not in HOST_DEFAULT_MEMBERS:
            # A Variant or an object it gives may still take an index.
            returns = normalize_type(found.returns)
            return (
                ".".join(names)
                if returns is not None and returns != "variant" and is_known_scalar_type(returns)
                else None
            )
        key = found.returns
    return None


def library_default_verdict(qualified: str, depth: int = 0) -> str | None:
    """The verdict from a Word, PowerPoint or Access type's default member as its
    type library gives it (DISPID 0, XLIDE issue #438): "lets", "argument", or None
    where the table has no entry. Port of typeInference.ts' private
    libraryDefaultVerdict, which objectLetAssignmentVerdict reads."""
    while True:
        found = HOST_DEFAULT_MEMBERS.get(qualified)
        if found is None:
            return None
        if found.kind == "method" or found.required > 0:
            return "argument"
        if depth < 4 and found.returns in HOST_DEFAULT_MEMBERS:
            qualified = found.returns
            depth += 1
            continue
        return "lets"


def host_default_member(qualified: str, member_ctx: MemberCompletionContext) -> HostMember | None:
    """A host type's default member (DISPID 0, `_Default` in the model), if any.
    Port of typeInference.ts' private hostDefaultMember."""
    return next((m for m in get_host_members(qualified, member_ctx.model) if m["name"] == "_Default"), None)


_EXCEL_QUALIFIED_RE = re.compile(r"^excel\.", re.IGNORECASE)


def host_type_has_no_default(resolved: str, member_ctx: MemberCompletionContext) -> bool:
    """Whether a host type provably has no default member: its member list is
    complete, and either the type library resolves members while compiling or the
    type is Excel's (XLIDE issue #221). Port of typeInference.ts' private
    hostTypeHasNoDefault."""
    host_type = get_host_type(resolved, member_ctx.model)
    if host_type is None or host_type.get("exhaustive") is not True:
        return False
    return host_type_resolves_when_compiling(resolved) or _EXCEL_QUALIFIED_RE.match(resolved) is not None


_REQUIRED_PARAMETER_RE = re.compile(r"\((?!\s*\[)[^)]")


def object_value_needs_index(type_name: str | None, member_ctx: MemberCompletionContext) -> bool:
    """Whether reading an object of this type as a value raises because its default
    member needs an index: 450 for a Collection, a host default property typed as
    an element (Hyperlinks), or a host default method with a required parameter
    (Shapes) (XLIDE issue #221). Names, whose default method takes only optional
    parameters, raises 449 and is not judged."""
    from ..diagnostics.rules.type_of_is import object_let_assignment_verdict

    if normalize_type(type_name) == "collection":
        return True
    if object_let_assignment_verdict(type_name, member_ctx) != "argument":
        return False
    resolved = resolve_host_alias(type_name or "", member_ctx.model) or library_object_type(type_name)
    # Word's Paragraphs and Tables, PowerPoint's Slides: Item(Index) raises 450 read
    # as a value (XLIDE issue #462).
    library = HOST_DEFAULT_MEMBERS.get(resolved) if resolved else None
    if library is not None:
        return library.required > 0
    default_member = host_default_member(resolved, member_ctx) if resolved else None
    if default_member is None:
        return False
    if default_member.get("kind") == "method":
        return _REQUIRED_PARAMETER_RE.search(default_member.get("signature") or "") is not None
    return True


# The reason a scalar value cannot be Set: a compile error, where every other
# reason is an object of the wrong class, which compiles and raises 13 when the
# Set runs (XLIDE issue #202).
SCALAR_OBJECT_ASSIGNMENT_REASON = "An object assignment requires an object value."


def create_project_interface_sharing_lookup(
    member_ctx: MemberCompletionContext,
) -> Callable[[str, str], bool]:
    """Find common implementers against metadata stable for one analysis pass.
    Retain only examined memberships and queried pairs, rather than all pairs that
    a surface with many implemented interfaces could form."""
    memberships: dict[str, set[int]] = {}
    pairs: dict[str, dict[str, bool]] = {}
    next_index = 0

    def share(expected: str, actual: str) -> bool:
        nonlocal next_index
        # Shared-implementer membership is symmetric; direct casts stay outside.
        first = expected if expected < actual else actual
        second = actual if expected < actual else expected
        answers = pairs.get(first)
        if answers is not None and second in answers:
            return answers[second]
        if answers is None:
            answers = {}
            pairs[first] = answers
        left = memberships.get(first)
        right = memberships.get(second)
        if left is not None and right is not None:
            smaller, larger = (left, right) if len(left) <= len(right) else (right, left)
            if any(index in larger for index in smaller):
                answers[second] = True
                return True
        surfaces = member_ctx.project_class_members or []
        while next_index < len(surfaces):
            index = next_index
            next_index += 1
            # Keep every kind: the original shared-interface scan did not filter kinds.
            for name in surfaces[index].implements or []:
                memberships.setdefault(name.lower(), set()).add(index)
            if index in memberships.get(first, ()) and index in memberships.get(second, ()):
                answers[second] = True
                return True
        answers[second] = False
        return False

    return share


def project_types_can_share_instance(
    expected: KnownObjectAssignmentType,
    actual: KnownObjectAssignmentType,
    member_ctx: MemberCompletionContext,
    share_interfaces: Callable[[str, str], bool] | None = None,
) -> bool:
    """Whether one project class can carry a value declared as the other: the
    expected class implements the actual type (a cast from an interface back to the
    class), or some project class implements both. Port of typeInference.ts'
    private projectTypesCanShareInstance."""
    if implements_object_type(expected, actual):
        return True
    if share_interfaces is not None:
        return share_interfaces(expected.key, actual.key)
    wanted = {expected.key, actual.key}
    for project_type in member_ctx.project_class_members or []:
        implemented = {name.lower() for name in project_type.implements or []}
        if wanted <= implemented:
            return True
    return False


def implements_object_type(actual: KnownObjectAssignmentType, expected: KnownObjectAssignmentType) -> bool:
    expected_names = {expected.key}
    simple = simple_type_name_for_assignment(expected.display)
    if simple:
        expected_names.add(simple.lower())
    expected_last_segment = expected.key.split(".")[-1]
    if expected_last_segment:
        expected_names.add(expected_last_segment)
    for implemented in actual.implements:
        lower = implemented.lower()
        if lower in expected_names or f"excel.{lower}" in expected_names:
            return True
    return False


# -- known local values (diagnostics/known_locals.py) -----------------------


def picked_values(
    values: Mapping[str, KnownLocalValue],
    pick: Callable[[str, KnownLocalValue], _T | None],
) -> Mapping[str, _T]:
    """A view of a statement's values through `pick`, worked out per name on first
    use (XLIDE issue #322)."""
    from ..diagnostics import known_locals

    return known_locals.picked_values(values, pick)


def known_local_literal_values_at(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> Callable[[BodyNode | None], Mapping[str, KnownLocalValue]]:
    """known_local_literal_values at each statement (XLIDE issue #180)."""
    from ..diagnostics import known_locals

    return known_locals.known_local_literal_values_at(source, proc, symbols, activity)


def statement_may_change_module_variable(
    source: str, symbols: ModuleSymbols, proc: ProcedureNode, span: Span, variable: str
) -> bool:
    """Whether a statement may run code that changes the module variable (XLIDE
    issue #618)."""
    from ..diagnostics import known_locals

    return known_locals.statement_may_change_module_variable(source, symbols, proc, span, variable)


def unreachable_statements_in(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> AbstractSet[int]:
    """The statements of a procedure that never run (XLIDE issue #273)."""
    from ..diagnostics import known_locals

    return known_locals.unreachable_statements_in(source, proc, symbols, activity)


def dead_branch_spans_in(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> Sequence[Span]:
    """The one-line If branches of a procedure that never run (XLIDE issue #430)."""
    from ..diagnostics import known_locals

    return known_locals.dead_branch_spans_in(source, proc, symbols, activity)


def defaulted_straight_line(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> Mapping[int, ReachingAssignments]:
    """The straight-line walk from what holds as the procedure starts (XLIDE
    issue #614)."""
    from ..diagnostics import known_locals

    return known_locals.defaulted_straight_line(source, proc, symbols, activity)


def function_result_for(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    args: Sequence[Sequence[VbaToken] | None],
    object_result: bool = False,
) -> Sequence[VbaToken] | None:
    """What a Function of the module returns for one call's arguments, as the
    tokens of the value it last assigns its name (XLIDE issue #562)."""
    from ..diagnostics import known_locals

    return known_locals.function_result_for(source, proc, symbols, activity, args, object_result)
