"""Builds the symbol view of a single module from its parser AST.

Ported from xlide_vscode/src/analyzer/symbols/buildModuleSymbols.ts. Pure
AST -> symbol projection. The XLIDE builder also attaches XML `'''` doc comments;
that is editor-only (agent.md Risk 7) and omitted here.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from ..conditional import (
    ConditionalActivityTracker,
    ConditionalCompilationEnvironment,
    create_conditional_activity_tracker,
    inactive_node_skip,
)
from ..parser.nodes import (
    AttributeNode,
    BodyNode,
    DeclareNode,
    EnumNode,
    EventNode,
    ModuleNode,
    ParameterNode,
    ProcedureNode,
    ProcKind,
    Span,
    TypeNode,
    VariableGroupNode,
    iter_body_nodes,
)
from ..js_compat import JS_WHITESPACE
from ..parser.parse_module import parse_module
from .symbol_model import (
    ModuleSymbolKind,
    ModuleSymbols,
    SymbolVisibility,
    VbaSymbol,
    VbaSymbolAttribute,
    VbaSymbolKind,
)

_RETURN_ARRAY_RE = re.compile(r"\(\s*\)\s*$")

_VISIBILITY_BY_WORD: dict[str, SymbolVisibility] = {
    "public": SymbolVisibility.PUBLIC,
    "private": SymbolVisibility.PRIVATE,
    "friend": SymbolVisibility.FRIEND,
    "global": SymbolVisibility.GLOBAL,
    "dim": SymbolVisibility.DIM,
    "static": SymbolVisibility.STATIC,
}

_PROC_SYMBOL_KIND: dict[ProcKind, VbaSymbolKind] = {
    ProcKind.SUB: VbaSymbolKind.SUB,
    ProcKind.FUNCTION: VbaSymbolKind.FUNCTION,
    ProcKind.PROPERTY_GET: VbaSymbolKind.PROPERTY_GET,
    ProcKind.PROPERTY_LET: VbaSymbolKind.PROPERTY_LET,
    ProcKind.PROPERTY_SET: VbaSymbolKind.PROPERTY_SET,
}


@dataclass(slots=True)
class BuildModuleSymbolsOptions:
    """Inputs to symbol building: conditional-compilation env and a parsed AST."""

    conditional_compilation: ConditionalCompilationEnvironment | None = None
    parsed_module: ModuleNode | None = None


def _is_inactive(activity: ConditionalActivityTracker | None, span: Span) -> bool:
    return activity is not None and activity.is_inactive(span)


def _to_visibility(word: str | None) -> SymbolVisibility | None:
    return _VISIBILITY_BY_WORD.get((word or "").lower())


def _proc_visibility(modifiers: list[str]) -> SymbolVisibility | None:
    for m in modifiers:
        v = _to_visibility(m)
        if v is not None and v is not SymbolVisibility.DIM and v is not SymbolVisibility.STATIC:
            return v
    return None


def _symbol_attribute(node: AttributeNode) -> VbaSymbolAttribute:
    dot = node.name.find(".")
    if dot > 0 and dot + 1 < len(node.name):
        return VbaSymbolAttribute(
            name=node.name[dot + 1 :],
            target_name=node.name[:dot],
            value_raw=node.value_raw,
            name_span=node.name_span,
            full_span=node.span,
        )
    return VbaSymbolAttribute(
        name=node.name, value_raw=node.value_raw, name_span=node.name_span, full_span=node.span
    )


def _attach_member_attributes(
    symbols: list[VbaSymbol], attributes: list[VbaSymbolAttribute]
) -> None:
    if not attributes:
        return
    by_name: dict[str, list[VbaSymbol]] = {}
    for symbol in symbols:
        by_name.setdefault(symbol.name.lower(), []).append(symbol)
    for attr in attributes:
        if not attr.target_name:
            continue
        # Every accessor/duplicate receives the attribute, in source order.
        for symbol in by_name.get(attr.target_name.lower(), ()):
            symbol.attributes = [*(symbol.attributes or []), attr]


def _collect_locals(
    body: list[BodyNode],
    module_name: str,
    container_name: str,
    out: list[VbaSymbol],
    activity: ConditionalActivityTracker | None,
) -> None:
    for node in iter_body_nodes(body, inactive_node_skip(activity)):
        if isinstance(node, VariableGroupNode):
            for decl in node.declarations:
                out.append(
                    VbaSymbol(
                        name=decl.name,
                        kind=VbaSymbolKind.CONSTANT if node.is_const else VbaSymbolKind.LOCAL_VARIABLE,
                        name_span=decl.name_span or decl.span,
                        full_span=decl.span,
                        module_name=module_name,
                        container_name=container_name,
                        visibility=_to_visibility(node.modifier),
                        as_type=decl.as_type,
                        fixed_length=decl.fixed_length,
                        default_raw=decl.default_raw,
                        is_array=decl.is_array,
                        is_auto_instantiated=decl.is_new,
                        array_bounds=decl.array_bounds,
                    )
                )


def _build_parameter_symbol(
    param: ParameterNode, module_name: str, container_name: str
) -> VbaSymbol:
    return VbaSymbol(
        name=param.name,
        kind=VbaSymbolKind.PARAMETER,
        name_span=param.name_span or param.span,
        full_span=param.span,
        module_name=module_name,
        container_name=container_name,
        as_type=param.as_type,
        optional=param.optional,
        param_array=param.param_array,
        by_val=param.by_val,
        by_ref=param.by_ref,
        is_array=param.is_array,
        default_raw=param.default_raw,
    )


def _procedure_return_is_array(return_type: str | None) -> bool:
    return _RETURN_ARRAY_RE.search(return_type or "") is not None


def _build_procedure(
    proc: ProcedureNode,
    module_name: str,
    flat: list[VbaSymbol],
    activity: ConditionalActivityTracker | None,
) -> VbaSymbol:
    children: list[VbaSymbol] = []
    symbol = VbaSymbol(
        name=proc.name,
        kind=_PROC_SYMBOL_KIND[proc.proc_kind],
        name_span=proc.name_span or proc.span,
        full_span=proc.span,
        module_name=module_name,
        visibility=_proc_visibility(proc.modifiers),
        as_type=proc.return_type,
        is_array=_procedure_return_is_array(proc.return_type),
        attributes=[_symbol_attribute(a) for a in proc.attributes] if proc.attributes else None,
        children=children,
    )
    for param in proc.params:
        param_symbol = _build_parameter_symbol(param, module_name, proc.name)
        children.append(param_symbol)
        flat.append(param_symbol)
    locals_: list[VbaSymbol] = []
    _collect_locals(proc.body, module_name, proc.name, locals_, activity)
    for local in locals_:
        children.append(local)
        flat.append(local)
    return symbol


def _build_declare(declare: DeclareNode, module_name: str, flat: list[VbaSymbol]) -> VbaSymbol:
    children: list[VbaSymbol] = []
    symbol = VbaSymbol(
        name=declare.name,
        kind=VbaSymbolKind.DECLARE,
        name_span=declare.name_span or declare.span,
        full_span=declare.span,
        module_name=module_name,
        visibility=_to_visibility(declare.visibility),
        as_type=declare.return_type,
        declare_kind="Function" if declare.is_function else "Sub",
        ptr_safe=declare.ptr_safe,
        lib_name=declare.lib_name,
        alias_name=declare.alias_name,
        children=children,
    )
    for param in declare.params:
        param_symbol = _build_parameter_symbol(param, module_name, declare.name)
        children.append(param_symbol)
        flat.append(param_symbol)
    return symbol


def _build_event(event: EventNode, module_name: str, flat: list[VbaSymbol]) -> VbaSymbol:
    children: list[VbaSymbol] = []
    symbol = VbaSymbol(
        name=event.name,
        kind=VbaSymbolKind.EVENT,
        name_span=event.name_span or event.span,
        full_span=event.span,
        module_name=module_name,
        visibility=_to_visibility(event.visibility),
        children=children,
    )
    for param in event.params:
        param_symbol = _build_parameter_symbol(param, module_name, event.name)
        children.append(param_symbol)
        flat.append(param_symbol)
    return symbol


def _build_type(node: TypeNode, module_name: str, flat: list[VbaSymbol]) -> VbaSymbol:
    children: list[VbaSymbol] = []
    symbol = VbaSymbol(
        name=node.name,
        kind=VbaSymbolKind.TYPE,
        name_span=node.name_span or node.span,
        full_span=node.span,
        module_name=module_name,
        visibility=_to_visibility(node.visibility),
        children=children,
    )
    for field_node in node.fields:
        field_symbol = VbaSymbol(
            name=field_node.name,
            kind=VbaSymbolKind.TYPE_FIELD,
            name_span=field_node.name_span or field_node.span,
            full_span=field_node.span,
            module_name=module_name,
            container_name=node.name,
            as_type=field_node.as_type,
            fixed_length=field_node.fixed_length,
            is_array=True if field_node.is_array else None,
        )
        children.append(field_symbol)
        flat.append(field_symbol)
    return symbol


def _build_enum(node: EnumNode, module_name: str, flat: list[VbaSymbol]) -> VbaSymbol:
    children: list[VbaSymbol] = []
    symbol = VbaSymbol(
        name=node.name,
        kind=VbaSymbolKind.ENUM,
        name_span=node.name_span or node.span,
        full_span=node.span,
        module_name=module_name,
        visibility=_to_visibility(node.visibility),
        children=children,
    )
    for member in node.members:
        member_symbol = VbaSymbol(
            name=member.name,
            kind=VbaSymbolKind.ENUM_MEMBER,
            name_span=member.name_span or member.span,
            full_span=member.span,
            module_name=module_name,
            container_name=node.name,
            default_raw=member.value_raw,
        )
        children.append(member_symbol)
        flat.append(member_symbol)
    return symbol


def _build_module_variables(
    group: VariableGroupNode,
    module_name: str,
    root_children: list[VbaSymbol],
    flat: list[VbaSymbol],
) -> None:
    for decl in group.declarations:
        symbol = VbaSymbol(
            name=decl.name,
            kind=VbaSymbolKind.CONSTANT if group.is_const else VbaSymbolKind.MODULE_VARIABLE,
            name_span=decl.name_span or decl.span,
            full_span=decl.span,
            module_name=module_name,
            visibility=_to_visibility(group.modifier),
            as_type=decl.as_type,
            fixed_length=decl.fixed_length,
            default_raw=decl.default_raw,
            is_array=decl.is_array,
            is_auto_instantiated=decl.is_new,
            array_bounds=decl.array_bounds,
        )
        root_children.append(symbol)
        flat.append(symbol)


def build_module_symbols(
    module_name: str,
    module_kind: ModuleSymbolKind,
    source: str,
    options: BuildModuleSymbolsOptions | None = None,
) -> ModuleSymbols:
    """Build the ModuleSymbols view of a module from its source text."""
    options = options or BuildModuleSymbolsOptions()
    module = options.parsed_module if options.parsed_module is not None else parse_module(source)
    activity = create_conditional_activity_tracker(module, options.conditional_compilation)
    root_children: list[VbaSymbol] = []
    flat: list[VbaSymbol] = []
    module_attributes: list[VbaSymbolAttribute] = []
    member_attributes: list[VbaSymbolAttribute] = []

    for member in module.members:
        if _is_inactive(activity, member.span):
            continue
        if isinstance(member, AttributeNode):
            attr = _symbol_attribute(member)
            if attr.target_name:
                member_attributes.append(attr)
            else:
                module_attributes.append(attr)
        elif isinstance(member, ProcedureNode):
            proc = _build_procedure(member, module_name, flat, activity)
            root_children.append(proc)
            flat.append(proc)
        elif isinstance(member, TypeNode):
            type_symbol = _build_type(member, module_name, flat)
            root_children.append(type_symbol)
            flat.append(type_symbol)
        elif isinstance(member, EnumNode):
            enum_symbol = _build_enum(member, module_name, flat)
            root_children.append(enum_symbol)
            flat.append(enum_symbol)
        elif isinstance(member, VariableGroupNode):
            _build_module_variables(member, module_name, root_children, flat)
        elif isinstance(member, DeclareNode):
            declare_symbol = _build_declare(member, module_name, flat)
            root_children.append(declare_symbol)
            flat.append(declare_symbol)
        elif isinstance(member, EventNode):
            event_symbol = _build_event(member, module_name, flat)
            root_children.append(event_symbol)
            flat.append(event_symbol)

    _attach_member_attributes(root_children, member_attributes)

    root = VbaSymbol(
        name=module_name,
        kind=VbaSymbolKind.MODULE,
        name_span=Span(module.span.start, module.span.start),
        full_span=module.span,
        module_name=module_name,
        children=root_children,
        attributes=module_attributes,
    )

    def_types = _module_def_types(source)
    implicit_locals = (
        _module_implicit_locals(source, module, root_children)
        if def_types and _OPTION_EXPLICIT_RE.search(source) is None
        else None
    )
    return ModuleSymbols(
        module_name=module_name,
        module_kind=module_kind,
        root=root,
        all=flat,
        def_types=def_types or None,
        implicit_locals=implicit_locals or None,
    )


_OPTION_EXPLICIT_RE = re.compile(r"^[ \t]*Option[ \t]+Explicit\b", re.IGNORECASE | re.MULTILINE | re.ASCII)
_IMPLICIT_ASSIGNMENT_RE = re.compile(
    r"^[ \t]*(?:\d+[ \t]+)?(?:Let[ \t]+|For[ \t]+)?([A-Za-z]\w*)[ \t]*=(?!=)",
    re.IGNORECASE | re.MULTILINE | re.ASCII,
)
_NOT_IMPLICIT_LOCAL_RE = re.compile(
    r"(let|set|for|if|elseif|while|until|case|call|dim|redim|static|const|private|public|global|end|exit"
    r"|on|resume|mid|lset|rset)"
)


def _module_implicit_locals(
    source: str, module: ModuleNode, root_children: list[VbaSymbol]
) -> dict[int, frozenset[str]]:
    """The names each procedure assigns, ``i = 40000`` or ``For i = 1 To 3``,
    that nothing in the procedure or the module declares: with no Option
    Explicit VBA makes each a local, and a DefType line types it (XLIDE issue
    #285)."""
    module_names = {symbol.name.lower() for symbol in root_children}
    by_name_start: dict[int, VbaSymbol] = {}
    for symbol in root_children:
        # Preserve the first-match contract of the previous Array.find.
        by_name_start.setdefault(symbol.name_span.start, symbol)
    out: dict[int, frozenset[str]] = {}
    for member in module.members:
        if not isinstance(member, ProcedureNode):
            continue
        declared = {name.lower() for name in [member.name, *(param.name for param in member.params)]}
        proc_symbol = by_name_start.get((member.name_span or member.span).start)
        for child in (proc_symbol.children if proc_symbol is not None else None) or []:
            declared.add(child.name.lower())
        names: set[str] = set()
        text = source[member.span.start : member.span.end]
        for match in _IMPLICIT_ASSIGNMENT_RE.finditer(text):
            lower = match.group(1).lower()
            if (
                lower not in declared
                and lower not in module_names
                and _NOT_IMPLICIT_LOCAL_RE.fullmatch(lower) is None
            ):
                names.add(lower)
        if names:
            out[member.span.start] = frozenset(names)
    return out


_DEF_TYPE_NAMES: dict[str, str] = {
    "bool": "Boolean", "byte": "Byte", "int": "Integer", "lng": "Long", "lnglng": "LongLong",
    "lngptr": "LongPtr", "cur": "Currency", "sng": "Single", "dbl": "Double", "dec": "Decimal",
    "date": "Date", "str": "String", "obj": "Object", "var": "Variant",
}
_DEF_TYPE_RE = re.compile(
    r"^[ \t]*Def(Bool|Byte|Int|LngLng|LngPtr|Lng|Cur|Sng|Dbl|Dec|Date|Str|Obj|Var)[ \t]+([A-Za-z][A-Za-z \t,-]*)",
    re.IGNORECASE | re.MULTILINE | re.ASCII,
)
_WS = "[" + re.escape(JS_WHITESPACE) + "]"
_DEF_TYPE_RANGE_RE = re.compile(r"^" + _WS + r"*([A-Za-z])" + _WS + r"*(?:-" + _WS + r"*([A-Za-z]))?" + _WS + r"*\Z")


def _module_def_types(source: str) -> dict[str, str]:
    """The type each first letter gives a name declared with no type, from the
    module's DefType lines: ``DefInt A-Z``, ``DefStr S, T-U`` (MS-VBAL 5.2.2;
    XLIDE issue #285, measured in Excel 16.0)."""
    out: dict[str, str] = {}
    for match in _DEF_TYPE_RE.finditer(source):
        type_ = _DEF_TYPE_NAMES[match.group(1).lower()]
        for range_ in match.group(2).split(","):
            ends = _DEF_TYPE_RANGE_RE.match(range_)
            if ends is None:
                continue
            start = ord(ends.group(1).lower())
            stop = ord((ends.group(2) or ends.group(1)).lower())
            for code in range(start, stop + 1):
                out[chr(code)] = type_
    return out
