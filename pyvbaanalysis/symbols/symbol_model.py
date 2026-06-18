"""VBA symbol model (Project-Wide Symbol Graph).

Ported from xlide_vscode/src/analyzer/symbols/symbolModel.ts. A thin, host-agnostic
projection of the parser AST into named declarations with scope and span info.

The XLIDE model carries an inline `doc` (XML `'''` documentation) on symbols and
signatures; that is editor-only (hover) per agent.md Risk 7 and no diagnostic rule
reads it, so it is intentionally omitted from this port.
"""

from __future__ import annotations

import enum
from dataclasses import dataclass, field

from ..parser.nodes import Span


class ModuleSymbolKind(str, enum.Enum):
    """The workbook-project role of a module (MS-VBAL 4.2 plus host kinds)."""

    STANDARD = "standard"
    CLASS = "class"
    DOCUMENT = "document"
    USERFORM = "userform"


class VbaSymbolKind(str, enum.Enum):
    """Every kind of named symbol the index can produce."""

    MODULE = "module"
    SUB = "sub"
    FUNCTION = "function"
    PROPERTY_GET = "propertyGet"
    PROPERTY_LET = "propertyLet"
    PROPERTY_SET = "propertySet"
    PARAMETER = "parameter"
    LOCAL_VARIABLE = "localVariable"
    MODULE_VARIABLE = "moduleVariable"
    CONSTANT = "constant"
    ENUM = "enum"
    ENUM_MEMBER = "enumMember"
    TYPE = "type"
    TYPE_FIELD = "typeField"
    EVENT = "event"
    DECLARE = "declare"


class SymbolVisibility(str, enum.Enum):
    """Declaration visibility (MS-VBAL 5.2.3.1 / 5.3.1.1)."""

    PUBLIC = "Public"
    PRIVATE = "Private"
    FRIEND = "Friend"
    GLOBAL = "Global"
    DIM = "Dim"
    STATIC = "Static"


class VbaProjectTypeKind(str, enum.Enum):
    """Project-defined type-name categories visible in As type positions."""

    CLASS = "class"
    DOCUMENT = "document"
    USERFORM = "userform"
    ENUM = "enum"
    USER_TYPE = "userType"


@dataclass(frozen=True, slots=True)
class VbaSymbolAttribute:
    """Exported VBA attribute attached to a module or member declaration."""

    name: str
    value_raw: str
    name_span: Span
    full_span: Span
    # Member target before the dot in `Attribute Value.VB_UserMemId = 0`.
    target_name: str | None = None


@dataclass(slots=True)
class VbaSymbol:
    """A single named declaration discovered in a module."""

    name: str
    kind: VbaSymbolKind
    # Span of the declared identifier (selection range for go-to-def).
    name_span: Span
    # Full span of the declaration.
    full_span: Span
    # Owning module's name (VB component name).
    module_name: str
    container_name: str | None = None
    visibility: SymbolVisibility | None = None
    as_type: str | None = None
    fixed_length: str | None = None
    default_raw: str | None = None
    optional: bool | None = None
    param_array: bool | None = None
    by_val: bool | None = None
    by_ref: bool | None = None
    is_array: bool | None = None
    array_bounds: str | None = None
    # External Declare statements are Function or Sub callables.
    declare_kind: str | None = None
    ptr_safe: bool | None = None
    lib_name: str | None = None
    alias_name: str | None = None
    children: list[VbaSymbol] | None = None
    attributes: list[VbaSymbolAttribute] | None = None


@dataclass(slots=True)
class VbaProcedureParam:
    """Parameter shape used by project-wide callable signature diagnostics."""

    name: str
    optional: bool
    param_array: bool
    is_array: bool
    type_: str | None = None
    default_raw: str | None = None
    by_val: bool | None = None
    by_ref: bool | None = None


@dataclass(slots=True)
class VbaProcedureSignature:
    """Exported callable signature collected from the project symbol graph."""

    name: str
    module_name: str
    kind: VbaSymbolKind  # only SUB or FUNCTION
    params: list[VbaProcedureParam] = field(default_factory=list)
    return_type: str | None = None
    signature: str | None = None
    visibility: SymbolVisibility | None = None
    external: bool | None = None
    ptr_safe: bool | None = None
    lib_name: str | None = None
    alias_name: str | None = None


@dataclass(slots=True)
class VbaProjectTypeName:
    """A project-defined type name visible to a module."""

    name: str
    kind: VbaProjectTypeKind
    module_name: str
    name_span: Span | None = None
    full_span: Span | None = None
    visibility: SymbolVisibility | None = None


@dataclass(frozen=True, slots=True)
class VbaProjectClassMemberDefinition:
    """Source declaration location for a project object member."""

    module_name: str
    name_span: Span
    full_span: Span


@dataclass(slots=True)
class VbaProjectClassMember:
    """A source-declared member of a project object module or user-defined Type."""

    name: str
    kind: str  # "property" | "method" | "event"
    module_name: str
    returns: str | None = None
    signature: str | None = None
    writable: bool | None = None
    write_type: str | None = None
    visibility: SymbolVisibility | None = None
    definitions: list[VbaProjectClassMemberDefinition] | None = None
    default_member: bool | None = None
    attributes: list[VbaSymbolAttribute] | None = None


@dataclass(slots=True)
class VbaProjectClassMembers:
    """Public member surface for an object type, standard module, or user Type."""

    name: str
    # "class" | "document" | "userform" | "userType" | "standardModule"
    kind: str
    module_name: str
    members: list[VbaProjectClassMember] = field(default_factory=list)
    # Interfaces named by module-level Implements statements.
    implements: list[str] | None = None
    # True when the member list is complete enough to prove absence.
    exhaustive: bool | None = None


@dataclass(slots=True)
class ModuleSymbols:
    """The symbol view of a single module."""

    module_name: str
    module_kind: ModuleSymbolKind
    # The module itself as a top-level symbol (with all members as children).
    root: VbaSymbol
    # Flat list of every symbol in the module, including nested ones.
    all: list[VbaSymbol]


def procedure_kind_keyword(kind: VbaSymbolKind) -> str:
    return "Function" if kind is VbaSymbolKind.FUNCTION else "Sub"


def _quote_vba_string(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def format_procedure_param_label(param: VbaProcedureParam, include_passing: bool = False) -> str:
    label = param.name
    if param.is_array:
        label += "()"
    if param.type_:
        label += f" As {param.type_}"
    if param.default_raw:
        label += f" = {param.default_raw}"
    if param.param_array:
        label = f"ParamArray {label}"
    if include_passing and not param.param_array:
        if param.by_val:
            label = f"ByVal {label}"
        elif param.by_ref:
            label = f"ByRef {label}"
    return f"[{label}]" if param.optional else label


def procedure_signature_label(procedure: VbaProcedureSignature) -> str:
    params = ", ".join(
        format_procedure_param_label(p, include_passing=bool(procedure.external))
        for p in procedure.params
    )
    returns = (
        f" As {procedure.return_type}"
        if (procedure.kind is VbaSymbolKind.FUNCTION and procedure.return_type)
        else ""
    )
    return f"{procedure.name}({params}){returns}"


def procedure_declaration_signature(procedure: VbaProcedureSignature) -> str:
    if procedure.external:
        keyword = procedure_kind_keyword(procedure.kind)
        ptr_safe = "PtrSafe " if procedure.ptr_safe else ""
        lib = f" Lib {_quote_vba_string(procedure.lib_name)}" if procedure.lib_name else ""
        alias = f" Alias {_quote_vba_string(procedure.alias_name)}" if procedure.alias_name else ""
        external_target = f"{lib}{alias} " if (lib or alias) else ""
        params = ", ".join(
            format_procedure_param_label(p, include_passing=True) for p in procedure.params
        )
        returns = (
            f" As {procedure.return_type}"
            if (procedure.kind is VbaSymbolKind.FUNCTION and procedure.return_type)
            else ""
        )
        return f"Declare {ptr_safe}{keyword} {procedure.name}{external_target}({params}){returns}"
    return f"{procedure_kind_keyword(procedure.kind)} {procedure_signature_label(procedure)}"


def qualified_procedure_key(module_name: str, name: str) -> str:
    """Lowercased key used for module-qualified procedure lookups."""
    return f"{module_name.lower()}.{name.lower()}"


def is_bare_callable_kind(kind: VbaSymbolKind) -> bool:
    """True for bare callables: Sub, Function, and external Declare statements."""
    return kind is VbaSymbolKind.SUB or kind is VbaSymbolKind.FUNCTION or kind is VbaSymbolKind.DECLARE


def procedure_params_from_symbol(
    symbol: VbaSymbol, include_passing: bool = False
) -> list[VbaProcedureParam]:
    """Converts a symbol's parameter children into the shared callable param model."""
    out: list[VbaProcedureParam] = []
    for child in symbol.children or []:
        if child.kind is not VbaSymbolKind.PARAMETER:
            continue
        param = VbaProcedureParam(
            name=child.name,
            type_=child.as_type,
            optional=child.optional or False,
            param_array=child.param_array or False,
            is_array=child.is_array or False,
        )
        if child.default_raw is not None:
            param.default_raw = child.default_raw
        if include_passing:
            if child.by_val:
                param.by_val = True
            elif child.by_ref:
                param.by_ref = True
        out.append(param)
    return out


def _callable_kind_for_symbol(symbol: VbaSymbol) -> VbaSymbolKind:
    if symbol.kind is VbaSymbolKind.DECLARE:
        return VbaSymbolKind.FUNCTION if symbol.declare_kind == "Function" else VbaSymbolKind.SUB
    return symbol.kind  # already SUB or FUNCTION


def procedure_signature_from_symbol(symbol: VbaSymbol) -> VbaProcedureSignature | None:
    """Converts a Sub/Function/Declare symbol into the shared callable signature."""
    if not is_bare_callable_kind(symbol.kind):
        return None
    external = symbol.kind is VbaSymbolKind.DECLARE
    kind = _callable_kind_for_symbol(symbol)
    signature = VbaProcedureSignature(
        name=symbol.name,
        module_name=symbol.module_name,
        kind=kind,
        params=procedure_params_from_symbol(symbol, include_passing=True),
        return_type=symbol.as_type,
        visibility=symbol.visibility,
        external=external or None,
        ptr_safe=symbol.ptr_safe,
        lib_name=symbol.lib_name,
        alias_name=symbol.alias_name,
    )
    signature.signature = (
        procedure_declaration_signature(signature)
        if external
        else procedure_signature_label(signature)
    )
    return signature


def is_procedure_kind(kind: VbaSymbolKind) -> bool:
    """True for the five source procedure body symbol kinds."""
    return kind in (
        VbaSymbolKind.SUB,
        VbaSymbolKind.FUNCTION,
        VbaSymbolKind.PROPERTY_GET,
        VbaSymbolKind.PROPERTY_LET,
        VbaSymbolKind.PROPERTY_SET,
    )


# Re-exported convenience for callers iterating params.
__all__ = [
    "ModuleSymbolKind",
    "VbaSymbolKind",
    "SymbolVisibility",
    "VbaProjectTypeKind",
    "VbaSymbol",
    "VbaSymbolAttribute",
    "VbaProcedureParam",
    "VbaProcedureSignature",
    "VbaProjectTypeName",
    "VbaProjectClassMemberDefinition",
    "VbaProjectClassMember",
    "VbaProjectClassMembers",
    "ModuleSymbols",
    "procedure_kind_keyword",
    "format_procedure_param_label",
    "procedure_signature_label",
    "procedure_declaration_signature",
    "qualified_procedure_key",
    "is_bare_callable_kind",
    "procedure_params_from_symbol",
    "procedure_signature_from_symbol",
    "is_procedure_kind",
]
