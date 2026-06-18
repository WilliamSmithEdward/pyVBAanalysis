"""Project-wide symbol graph (cross-module index).

Ported from xlide_vscode/src/analyzer/symbols/projectIndex.ts. Aggregates
per-module symbol views into a workbook-project index and provides the visibility,
signature, type-name, member-surface, and bare-identifier resolution surfaces the
diagnostics engine consumes. Name resolution order follows MS-VBAL 5.3 / 4.2 /
5.2.3.1. The editor-only `doc` field is omitted throughout (agent.md Risk 7).
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from typing import TypeVar, cast

from ..conditional import ConditionalCompilationEnvironment
from ..constants.integer_constant_expression import (
    enum_member_raw_expression,
    parse_vba_integer_literal,
    resolve_raw_integer_constants,
)
from ..parser.nodes import Span
from .build_module_symbols import BuildModuleSymbolsOptions, build_module_symbols
from .name_resolution import (
    BareIdentifierContext,
    BareIdentifierResolution,
    BareIdentifierResolutionInput,
    BareIdentifierResolutionScope,
    resolve_bare_identifier_binding,
)
from .symbol_model import (
    ModuleSymbolKind,
    ModuleSymbols,
    SymbolVisibility,
    VbaProcedureSignature,
    VbaProjectClassMember,
    VbaProjectClassMemberDefinition,
    VbaProjectClassMembers,
    VbaProjectTypeKind,
    VbaProjectTypeName,
    VbaSymbol,
    VbaSymbolAttribute,
    VbaSymbolKind,
    format_procedure_param_label,
    is_bare_callable_kind,
    is_procedure_kind,
    procedure_params_from_symbol,
    procedure_signature_from_symbol,
    procedure_signature_label,
    qualified_procedure_key,
)

_T = TypeVar("_T")


@dataclass(slots=True)
class ModuleInput:
    """Source text + workbook role for one module fed into the index."""

    module_name: str
    module_kind: ModuleSymbolKind
    source: str
    conditional_compilation: ConditionalCompilationEnvironment | None = None


@dataclass(slots=True)
class ProjectIndexOptions:
    """Project-wide options shared by every indexed module."""

    conditional_compilation: ConditionalCompilationEnvironment | None = None


@dataclass(slots=True)
class ShadowedSpan:
    """A procedure span (within a named module) that shadows a name with a local."""

    module_name: str
    span: Span


@dataclass(slots=True)
class ReferenceScope:
    """The binding scope of an identifier, restricting reference/rename search."""

    kind: str  # "local" | "module" | "project"
    definitions: list[VbaSymbol]
    search_modules: list[str]
    shadowed_spans: list[ShadowedSpan]
    procedure_span: Span | None = None


@dataclass(slots=True)
class _ModuleLevelBinding:
    symbol: VbaSymbol
    exported: bool


# --- module-level helpers (free functions in the TS) -----------------------


def _is_exported(symbol: VbaSymbol, module_kind: ModuleSymbolKind | None = None) -> bool:
    if symbol.visibility is SymbolVisibility.PUBLIC or symbol.visibility is SymbolVisibility.GLOBAL:
        return True
    if symbol.visibility:
        return False
    return module_kind is ModuleSymbolKind.STANDARD and is_procedure_kind(symbol.kind)


def _add_procedure_signature(
    signatures: dict[str, list[VbaProcedureSignature]], key: str, sig: VbaProcedureSignature
) -> None:
    existing = signatures.get(key)
    if existing is not None:
        existing.append(sig)
    else:
        signatures[key] = [sig]


def _module_kind_as_type_name(kind: ModuleSymbolKind) -> VbaProjectTypeKind | None:
    if kind is ModuleSymbolKind.CLASS:
        return VbaProjectTypeKind.CLASS
    if kind is ModuleSymbolKind.DOCUMENT:
        return VbaProjectTypeKind.DOCUMENT
    if kind is ModuleSymbolKind.USERFORM:
        return VbaProjectTypeKind.USERFORM
    return None


def _project_type_kind(symbol: VbaSymbol) -> VbaProjectTypeKind | None:
    if symbol.kind is VbaSymbolKind.ENUM:
        return VbaProjectTypeKind.ENUM
    if symbol.kind is VbaSymbolKind.TYPE:
        return VbaProjectTypeKind.USER_TYPE
    return None


def _is_type_exported(symbol: VbaSymbol) -> bool:
    return symbol.visibility is not SymbolVisibility.PRIVATE


def _is_enum_member_exported(enum_symbol: VbaSymbol, module_kind: ModuleSymbolKind | None) -> bool:
    return module_kind is ModuleSymbolKind.STANDARD and _is_type_exported(enum_symbol)


def _module_raw_integer_constant_expressions(mod: ModuleSymbols) -> dict[str, str | None]:
    out: dict[str, str | None] = {}
    seen: set[str] = set()
    module_lower = mod.module_name.lower()

    def add(name: str, raw: str | None) -> None:
        key = name.lower()
        if key in seen:
            out[key] = None
            out[f"{module_lower}.{key}"] = None
            return
        seen.add(key)
        out[key] = raw
        out[f"{module_lower}.{key}"] = raw

    for symbol in mod.root.children or []:
        if symbol.kind is VbaSymbolKind.CONSTANT:
            add(symbol.name, symbol.default_raw)
            continue
        if symbol.kind is VbaSymbolKind.ENUM:
            previous_name: str | None = None
            for member in symbol.children or []:
                add(member.name, enum_member_raw_expression(member.default_raw, previous_name))
                previous_name = member.name
    return out


def _is_visible_project_object_member(symbol: VbaSymbol) -> bool:
    if is_procedure_kind(symbol.kind):
        return symbol.visibility is not SymbolVisibility.PRIVATE
    if symbol.kind is VbaSymbolKind.EVENT:
        return symbol.visibility is not SymbolVisibility.PRIVATE
    if symbol.kind is VbaSymbolKind.MODULE_VARIABLE:
        return symbol.visibility is SymbolVisibility.PUBLIC or symbol.visibility is SymbolVisibility.GLOBAL
    return False


def _project_object_member_kind(symbol: VbaSymbol) -> str | None:
    if symbol.kind in (VbaSymbolKind.SUB, VbaSymbolKind.FUNCTION, VbaSymbolKind.DECLARE):
        return "method"
    if symbol.kind is VbaSymbolKind.EVENT:
        return "event"
    if symbol.kind in (
        VbaSymbolKind.PROPERTY_GET,
        VbaSymbolKind.PROPERTY_LET,
        VbaSymbolKind.PROPERTY_SET,
        VbaSymbolKind.MODULE_VARIABLE,
        VbaSymbolKind.CONSTANT,
        VbaSymbolKind.ENUM,
        VbaSymbolKind.ENUM_MEMBER,
    ):
        return "property"
    return None


def _project_object_member_writable(symbol: VbaSymbol) -> bool | None:
    if symbol.kind in (VbaSymbolKind.PROPERTY_LET, VbaSymbolKind.PROPERTY_SET, VbaSymbolKind.MODULE_VARIABLE):
        return True
    if symbol.kind in (VbaSymbolKind.PROPERTY_GET, VbaSymbolKind.CONSTANT, VbaSymbolKind.ENUM, VbaSymbolKind.ENUM_MEMBER):
        return False
    return None


def _last_parameter(symbol: VbaSymbol) -> VbaSymbol | None:
    params = [child for child in (symbol.children or []) if child.kind is VbaSymbolKind.PARAMETER]
    return params[-1] if params else None


def _project_object_member_write_type(symbol: VbaSymbol) -> str | None:
    if symbol.kind in (VbaSymbolKind.PROPERTY_LET, VbaSymbolKind.PROPERTY_SET):
        last = _last_parameter(symbol)
        return last.as_type if last is not None else None
    if symbol.kind is VbaSymbolKind.MODULE_VARIABLE:
        return symbol.as_type
    return None


def _enum_container_for_member(mod: ModuleSymbols, member: VbaSymbol) -> VbaSymbol | None:
    lower = member.container_name.lower() if member.container_name else None
    if not lower:
        return None
    return next(
        (
            symbol
            for symbol in (mod.root.children or [])
            if symbol.kind is VbaSymbolKind.ENUM and symbol.name.lower() == lower
        ),
        None,
    )


def _is_visible_standard_module_member(symbol: VbaSymbol, mod: ModuleSymbols, same_module: bool) -> bool:
    if not _project_object_member_kind(symbol):
        return False
    if same_module:
        return True
    if symbol.kind is VbaSymbolKind.ENUM:
        return _is_type_exported(symbol)
    if symbol.kind is VbaSymbolKind.ENUM_MEMBER:
        container = _enum_container_for_member(mod, symbol)
        return _is_enum_member_exported(container, mod.module_kind) if container is not None else False
    return _is_exported(symbol, mod.module_kind)


def _project_object_member_return_type(symbol: VbaSymbol) -> str | None:
    if symbol.kind is VbaSymbolKind.ENUM_MEMBER:
        return symbol.container_name
    return symbol.as_type


def _project_object_member_definition(symbol: VbaSymbol) -> VbaProjectClassMemberDefinition:
    return VbaProjectClassMemberDefinition(
        module_name=symbol.module_name, name_span=symbol.name_span, full_span=symbol.full_span
    )


# OLE Automation DISPID for a type's default member (DISPID_VALUE).
_DISPID_VALUE = 0


def _is_default_member_attribute(attr: VbaSymbolAttribute) -> bool:
    return attr.name.lower() == "vb_usermemid" and parse_vba_integer_literal(attr.value_raw) == _DISPID_VALUE


def _is_default_project_object_member(symbol: VbaSymbol) -> bool:
    return any(_is_default_member_attribute(a) for a in (symbol.attributes or []))


def _merge_member_attributes(
    existing: list[VbaSymbolAttribute] | None, incoming: list[VbaSymbolAttribute] | None
) -> list[VbaSymbolAttribute] | None:
    if not incoming:
        return list(existing) if existing else None
    out = list(existing or [])
    seen = {f"{a.full_span.start}:{a.full_span.end}" for a in out}
    for attr in incoming:
        key = f"{attr.full_span.start}:{attr.full_span.end}"
        if key in seen:
            continue
        seen.add(key)
        out.append(attr)
    return out


_IMPLEMENTS_RE = re.compile(
    r"^\s*Implements\s+([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?)\b", re.IGNORECASE
)


def _module_implements(source: str) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for line in re.split(r"\r?\n", source):
        match = _IMPLEMENTS_RE.match(line)
        if match is None:
            continue
        name = match.group(1)
        lower = name.lower()
        if lower not in seen:
            seen.add(lower)
            out.append(name)
    return out


def _project_object_member_signature(symbol: VbaSymbol) -> str | None:
    procedure = procedure_signature_from_symbol(symbol)
    if procedure is not None:
        return procedure_signature_label(procedure)
    if symbol.kind is VbaSymbolKind.EVENT:
        params = ", ".join(format_procedure_param_label(p) for p in procedure_params_from_symbol(symbol))
        return f"{symbol.name}({params})"
    if symbol.kind is not VbaSymbolKind.PROPERTY_GET:
        return None
    params = ", ".join(format_procedure_param_label(p) for p in procedure_params_from_symbol(symbol))
    returns = f" As {symbol.as_type}" if symbol.as_type else ""
    return f"{symbol.name}({params}){returns}"


def _project_member_candidate_symbols(mod: ModuleSymbols, include_enum_members: bool) -> list[VbaSymbol]:
    out: list[VbaSymbol] = []
    for symbol in mod.root.children or []:
        out.append(symbol)
        if include_enum_members and symbol.kind is VbaSymbolKind.ENUM:
            out.extend(symbol.children or [])
    return out


def _user_type_field_signature(symbol: VbaSymbol) -> str:
    fixed_length = f" * {symbol.fixed_length}" if symbol.fixed_length else ""
    as_clause = f" As {symbol.as_type}{fixed_length}" if symbol.as_type else ""
    return f"{symbol.name}{as_clause}"


class ProjectIndex:
    """A project-wide symbol index built from a set of module sources."""

    __slots__ = (
        "_options",
        "_modules",
        "_module_sources",
        "_module_resolved_constants",
        "_module_implements_lists",
        "_query_cache",
    )

    def __init__(self, options: ProjectIndexOptions | None = None) -> None:
        self._options = options or ProjectIndexOptions()
        self._modules: dict[str, ModuleSymbols] = {}
        self._module_sources: dict[str, str] = {}
        self._module_resolved_constants: dict[str, dict[str, int | None]] = {}
        self._module_implements_lists: dict[str, list[str]] = {}
        self._query_cache: dict[str, object] = {}

    # --- mutation ---------------------------------------------------------

    def set_module(self, input: ModuleInput) -> None:
        """Adds or replaces a module in the index."""
        symbols = build_module_symbols(
            input.module_name,
            input.module_kind,
            input.source,
            BuildModuleSymbolsOptions(
                conditional_compilation=input.conditional_compilation
                or self._options.conditional_compilation
            ),
        )
        key = input.module_name.lower()
        self._modules[key] = symbols
        self._module_sources[key] = input.source
        self._invalidate(key)

    def remove_module(self, module_name: str) -> None:
        """Removes a module from the index."""
        key = module_name.lower()
        self._modules.pop(key, None)
        self._module_sources.pop(key, None)
        self._invalidate(key)

    def _invalidate(self, key: str) -> None:
        self._module_resolved_constants.pop(key, None)
        self._module_implements_lists.pop(key, None)
        self._query_cache.clear()

    def _cached(self, key: str, compute: Callable[[], _T]) -> _T:
        if key in self._query_cache:
            return cast(_T, self._query_cache[key])
        value = compute()
        self._query_cache[key] = value
        return value

    def _module_integer_constants(self, mod: ModuleSymbols) -> Mapping[str, int | None]:
        key = mod.module_name.lower()
        resolved = self._module_resolved_constants.get(key)
        if resolved is None:
            resolved = resolve_raw_integer_constants(_module_raw_integer_constant_expressions(mod))
            self._module_resolved_constants[key] = resolved
        return resolved

    def _module_implements_for(self, mod: ModuleSymbols) -> list[str]:
        key = mod.module_name.lower()
        items = self._module_implements_lists.get(key)
        if items is None:
            items = _module_implements(self._module_sources.get(key) or "")
            self._module_implements_lists[key] = items
        return items

    # --- queries ----------------------------------------------------------

    def module_names(self) -> list[str]:
        return [m.module_name for m in self._modules.values()]

    def visible_procedure_names(self, module_name: str) -> set[str]:
        current_lower = module_name.lower()

        def compute() -> set[str]:
            names: set[str] = set()
            for mod in self._modules.values():
                same_module = mod.module_name.lower() == current_lower
                for symbol in mod.root.children or []:
                    if not is_bare_callable_kind(symbol.kind):
                        continue
                    if same_module or (
                        mod.module_kind is ModuleSymbolKind.STANDARD and _is_exported(symbol, mod.module_kind)
                    ):
                        names.add(symbol.name.lower())
            return names

        return set(self._cached(f"procedureNames:{current_lower}", compute))

    def visible_procedure_signatures(self, module_name: str) -> list[VbaProcedureSignature]:
        current_lower = module_name.lower()

        def compute() -> list[VbaProcedureSignature]:
            out: list[VbaProcedureSignature] = []
            for mod in self._modules.values():
                same_module = mod.module_name.lower() == current_lower
                for symbol in mod.root.children or []:
                    if not is_bare_callable_kind(symbol.kind):
                        continue
                    if not same_module and (
                        mod.module_kind is not ModuleSymbolKind.STANDARD
                        or not _is_exported(symbol, mod.module_kind)
                    ):
                        continue
                    signature = procedure_signature_from_symbol(symbol)
                    if signature is not None:
                        out.append(signature)
            return out

        return list(self._cached(f"procedureSignatures:{current_lower}", compute))

    def visible_identifier_names(self, module_name: str) -> set[str]:
        current_lower = module_name.lower()

        def compute() -> set[str]:
            names: set[str] = set()
            for mod in self._modules.values():
                same_module = mod.module_name.lower() == current_lower
                if mod.module_kind is ModuleSymbolKind.DOCUMENT or mod.module_kind is ModuleSymbolKind.USERFORM:
                    names.add(mod.module_name.lower())
                for symbol in self._visible_module_level_identifier_symbols(mod, same_module):
                    names.add(symbol.name.lower())
            return names

        return set(self._cached(f"identifierNames:{current_lower}", compute))

    def visible_identifier_symbols(self, module_name: str) -> list[VbaSymbol]:
        current_lower = module_name.lower()

        def compute() -> list[VbaSymbol]:
            out: list[VbaSymbol] = []
            for mod in self._modules.values():
                same_module = mod.module_name.lower() == current_lower
                out.extend(self._visible_module_level_identifier_symbols(mod, same_module))
            return out

        return list(self._cached(f"identifierSymbols:{current_lower}", compute))

    def visible_external_integer_constant_expressions(self, module_name: str) -> dict[str, str | None]:
        current_lower = module_name.lower()

        def compute() -> dict[str, str | None]:
            out: dict[str, str | None] = {}
            seen: set[str] = set()

            def add(name: str, raw: str | None) -> None:
                key = name.lower()
                if key in seen:
                    out[key] = None
                    return
                seen.add(key)
                out[key] = raw

            def add_qualified(mod: ModuleSymbols, name: str, raw: str | None) -> None:
                out[f"{mod.module_name.lower()}.{name.lower()}"] = raw

            def resolved_raw(resolved: Mapping[str, int | None], key: str, fallback: str | None) -> str | None:
                value = resolved.get(key.lower())
                return fallback if value is None else str(value)

            for mod in self._modules.values():
                if mod.module_name.lower() == current_lower or mod.module_kind is not ModuleSymbolKind.STANDARD:
                    continue
                module_resolved = self._module_integer_constants(mod)
                for symbol in mod.root.children or []:
                    if symbol.kind is VbaSymbolKind.CONSTANT and _is_exported(symbol, mod.module_kind):
                        raw = resolved_raw(module_resolved, symbol.name, symbol.default_raw)
                        add(symbol.name, raw)
                        add_qualified(mod, symbol.name, raw)
                        continue
                    if symbol.kind is VbaSymbolKind.ENUM and _is_enum_member_exported(symbol, mod.module_kind):
                        previous_name: str | None = None
                        for member in symbol.children or []:
                            fallback = enum_member_raw_expression(member.default_raw, previous_name)
                            raw = resolved_raw(module_resolved, member.name, fallback)
                            add(member.name, raw)
                            add_qualified(mod, member.name, raw)
                            previous_name = member.name
            return out

        return dict(self._cached(f"integerConstants:{current_lower}", compute))

    def visible_non_type_names(self, module_name: str) -> set[str]:
        current_lower = module_name.lower()

        def compute() -> set[str]:
            names: set[str] = set()
            for mod in self._modules.values():
                same_module = mod.module_name.lower() == current_lower
                for symbol in self._visible_module_level_identifier_symbols(mod, same_module):
                    if _project_type_kind(symbol):
                        continue
                    names.add(symbol.name.lower())
            return names

        return set(self._cached(f"nonTypeNames:{current_lower}", compute))

    def procedure_signatures(self) -> dict[str, list[VbaProcedureSignature]]:
        def compute() -> dict[str, list[VbaProcedureSignature]]:
            signatures: dict[str, list[VbaProcedureSignature]] = {}
            for mod in self._modules.values():
                for symbol in mod.root.children or []:
                    if (
                        not is_bare_callable_kind(symbol.kind)
                        or mod.module_kind is not ModuleSymbolKind.STANDARD
                        or not _is_exported(symbol, mod.module_kind)
                    ):
                        continue
                    sig = procedure_signature_from_symbol(symbol)
                    if sig is not None:
                        _add_procedure_signature(signatures, symbol.name.lower(), sig)
                        _add_procedure_signature(
                            signatures, qualified_procedure_key(symbol.module_name, symbol.name), sig
                        )
            return signatures

        return dict(self._cached("procedureSignaturesByKey", compute))

    def visible_type_names(self, module_name: str) -> list[VbaProjectTypeName]:
        current_lower = module_name.lower()

        def compute() -> list[VbaProjectTypeName]:
            out: list[VbaProjectTypeName] = []
            for mod in self._modules.values():
                same_module = mod.module_name.lower() == current_lower
                module_type_kind = _module_kind_as_type_name(mod.module_kind)
                if module_type_kind is not None:
                    out.append(
                        VbaProjectTypeName(
                            name=mod.module_name,
                            kind=module_type_kind,
                            module_name=mod.module_name,
                            name_span=mod.root.name_span,
                            full_span=mod.root.full_span,
                        )
                    )
                for symbol in mod.root.children or []:
                    kind = _project_type_kind(symbol)
                    if kind is None:
                        continue
                    if not same_module and not _is_type_exported(symbol):
                        continue
                    out.append(
                        VbaProjectTypeName(
                            name=symbol.name,
                            kind=kind,
                            module_name=mod.module_name,
                            name_span=symbol.name_span,
                            full_span=symbol.full_span,
                            visibility=symbol.visibility,
                        )
                    )
            return out

        return list(self._cached(f"typeNames:{current_lower}", compute))

    def resolve_type_definitions(self, module_name: str, name: str) -> list[VbaProjectTypeName]:
        lower = name.lower()
        return [t for t in self.visible_type_names(module_name) if t.name.lower() == lower]

    def project_class_members(self) -> list[VbaProjectClassMembers]:
        def compute() -> list[VbaProjectClassMembers]:
            out: list[VbaProjectClassMembers] = []
            for mod in self._modules.values():
                kind = _module_kind_as_type_name(mod.module_kind)
                if kind not in (VbaProjectTypeKind.CLASS, VbaProjectTypeKind.DOCUMENT, VbaProjectTypeKind.USERFORM):
                    continue
                members = self._visible_object_members(mod)
                out.append(
                    VbaProjectClassMembers(
                        name=mod.module_name,
                        kind=kind.value,
                        module_name=mod.module_name,
                        implements=self._module_implements_for(mod),
                        exhaustive=kind is VbaProjectTypeKind.CLASS,
                        members=members,
                    )
                )
            return out

        return list(self._cached("projectClassMembers", compute))

    def project_standard_module_members(self, module_name: str) -> list[VbaProjectClassMembers]:
        current_lower = module_name.lower()
        out: list[VbaProjectClassMembers] = []
        for mod in self._modules.values():
            if mod.module_kind is not ModuleSymbolKind.STANDARD:
                continue
            same_module = mod.module_name.lower() == current_lower
            out.append(
                VbaProjectClassMembers(
                    name=mod.module_name,
                    kind="standardModule",
                    module_name=mod.module_name,
                    exhaustive=True,
                    members=self._visible_standard_module_members(mod, same_module),
                )
            )
        return out

    def project_member_surfaces(self, module_name: str) -> list[VbaProjectClassMembers]:
        current_lower = module_name.lower()

        def compute() -> list[VbaProjectClassMembers]:
            return [
                *self.project_class_members(),
                *self.project_standard_module_members(module_name),
                *self._project_user_type_members(module_name),
            ]

        return list(self._cached(f"memberSurfaces:{current_lower}", compute))

    def get_module(self, module_name: str) -> ModuleSymbols | None:
        return self._modules.get(module_name.lower())

    def document_symbols(self, module_name: str) -> VbaSymbol | None:
        mod = self._modules.get(module_name.lower())
        return mod.root if mod is not None else None

    def workspace_symbols(self, query: str | None = None) -> list[VbaSymbol]:
        needle = query.strip().lower() if query is not None else None
        out: list[VbaSymbol] = []
        for mod in self._modules.values():
            for symbol in mod.all:
                if not needle or needle in symbol.name.lower():
                    out.append(symbol)
        return out

    def resolve_definition(self, module_name: str, name: str, offset: int) -> list[VbaSymbol]:
        return list(
            self.resolve_bare_identifier(module_name, name, offset, BareIdentifierContext.EXPRESSION).definitions
        )

    def resolve_bare_identifier(
        self, module_name: str, name: str, offset: int, context: BareIdentifierContext
    ) -> BareIdentifierResolution:
        home = self._modules.get(module_name.lower())
        if home is None:
            return BareIdentifierResolution(
                name=name,
                lower_name=name.lower(),
                context=context,
                scope=BareIdentifierResolutionScope.UNRESOLVED,
                definitions=[],
                reason=f"Module '{module_name}' is not indexed.",
            )
        return resolve_bare_identifier_binding(
            BareIdentifierResolutionInput(
                current_module=home,
                name=name,
                context=context,
                enclosing_procedure=self._enclosing_procedure(home, offset),
                offset=offset,
                project_visible_symbols=self.visible_identifier_symbols(module_name),
            )
        )

    def resolve_qualified_definition(self, qualifier: str, name: str) -> list[VbaSymbol]:
        mod = self._modules.get(qualifier.lower())
        if mod is None:
            return []
        return self._exported_module_level_matches(mod, name.lower())

    def reference_scope(self, module_name: str, name: str, offset: int) -> ReferenceScope:
        lower = name.lower()
        home = self._modules.get(module_name.lower())
        resolved = (
            self.resolve_bare_identifier(module_name, name, offset, BareIdentifierContext.EXPRESSION)
            if home is not None
            else None
        )

        if home is not None and resolved is not None:
            if resolved.scope is BareIdentifierResolutionScope.LOCAL:
                enclosing = self._enclosing_procedure(home, offset)
                return ReferenceScope(
                    kind="local",
                    definitions=list(resolved.definitions),
                    search_modules=[home.module_name],
                    procedure_span=enclosing.full_span if enclosing is not None else None,
                    shadowed_spans=[],
                )
            if resolved.scope is BareIdentifierResolutionScope.MODULE or (
                resolved.scope is BareIdentifierResolutionScope.AMBIGUOUS
                and resolved.tier is BareIdentifierResolutionScope.MODULE
            ):
                exported_home_hits = self._exported_module_level_matches(home, lower)
                if len(exported_home_hits) > 0:
                    return self._project_scope(lower, exported_home_hits)
                return ReferenceScope(
                    kind="module",
                    definitions=list(resolved.definitions),
                    search_modules=[home.module_name],
                    shadowed_spans=self._local_shadow_spans(home, lower),
                )
            if resolved.scope is BareIdentifierResolutionScope.PROJECT:
                return self._project_scope(lower, list(resolved.definitions))
            if (
                resolved.scope is BareIdentifierResolutionScope.AMBIGUOUS
                and resolved.tier is BareIdentifierResolutionScope.PROJECT
            ):
                return self._project_scope(lower, list(resolved.definitions))

        return ReferenceScope(
            kind="module",
            definitions=[],
            search_modules=[home.module_name] if home is not None else [module_name],
            shadowed_spans=self._local_shadow_spans(home, lower) if home is not None else [],
        )

    def duplicate_procedures(self, module_name: str) -> list[VbaSymbol]:
        mod = self._modules.get(module_name.lower())
        if mod is None:
            return []
        seen: dict[str, list[VbaSymbol]] = {}
        for symbol in mod.root.children or []:
            if is_procedure_kind(symbol.kind):
                key = symbol.name.lower()
                seen.setdefault(key, []).append(symbol)
        dupes: list[VbaSymbol] = []
        for items in seen.values():
            if len(items) > 1:
                dupes.extend(items)
        return dupes

    # --- private helpers --------------------------------------------------

    def _enclosing_procedure(self, mod: ModuleSymbols, offset: int) -> VbaSymbol | None:
        return next(
            (
                c
                for c in (mod.root.children or [])
                if is_procedure_kind(c.kind) and c.full_span.start <= offset <= c.full_span.end
            ),
            None,
        )

    def _project_scope(self, lower: str, definitions: list[VbaSymbol]) -> ReferenceScope:
        search_modules: list[str] = []
        shadowed_spans: list[ShadowedSpan] = []
        for mod in self._modules.values():
            module_hits = self._module_level_bindings(mod, lower)
            privately_shadowed = len(module_hits) > 0 and not any(b.exported for b in module_hits)
            if privately_shadowed:
                continue
            search_modules.append(mod.module_name)
            shadowed_spans.extend(self._local_shadow_spans(mod, lower))
        return ReferenceScope(
            kind="project", definitions=definitions, search_modules=search_modules, shadowed_spans=shadowed_spans
        )

    def _local_shadow_spans(self, mod: ModuleSymbols, lower: str) -> list[ShadowedSpan]:
        spans: list[ShadowedSpan] = []
        for symbol in mod.root.children or []:
            if not is_procedure_kind(symbol.kind):
                continue
            shadows = any(
                c.kind in (VbaSymbolKind.PARAMETER, VbaSymbolKind.LOCAL_VARIABLE, VbaSymbolKind.CONSTANT)
                and c.name.lower() == lower
                for c in (symbol.children or [])
            )
            if shadows:
                spans.append(ShadowedSpan(module_name=mod.module_name, span=symbol.full_span))
        return spans

    def _is_bare_identifier_visible(self, symbol: VbaSymbol, mod: ModuleSymbols, same_module: bool) -> bool:
        if same_module:
            return True
        if mod.module_kind is not ModuleSymbolKind.STANDARD:
            return False
        if symbol.kind is VbaSymbolKind.ENUM or symbol.kind is VbaSymbolKind.TYPE:
            return _is_type_exported(symbol)
        return _is_exported(symbol, mod.module_kind)

    def _visible_module_level_identifier_symbols(self, mod: ModuleSymbols, same_module: bool) -> list[VbaSymbol]:
        out: list[VbaSymbol] = []
        for symbol in mod.root.children or []:
            if not self._is_bare_identifier_visible(symbol, mod, same_module):
                continue
            out.append(symbol)
            if symbol.kind is VbaSymbolKind.ENUM:
                out.extend(symbol.children or [])
        return out

    def _exported_module_level_matches(self, mod: ModuleSymbols, lower: str) -> list[VbaSymbol]:
        return [b.symbol for b in self._module_level_bindings(mod, lower) if b.exported]

    def _module_level_bindings(self, mod: ModuleSymbols, lower: str) -> list[_ModuleLevelBinding]:
        hits: list[_ModuleLevelBinding] = []
        for symbol in mod.root.children or []:
            if symbol.name.lower() == lower:
                hits.append(_ModuleLevelBinding(symbol=symbol, exported=_is_exported(symbol, mod.module_kind)))
            if symbol.kind is VbaSymbolKind.ENUM:
                exported = _is_enum_member_exported(symbol, mod.module_kind)
                for member in symbol.children or []:
                    if member.name.lower() == lower:
                        hits.append(_ModuleLevelBinding(symbol=member, exported=exported))
        return hits

    def _visible_object_members(self, mod: ModuleSymbols) -> list[VbaProjectClassMember]:
        return self._visible_project_members(mod, _is_visible_project_object_member)

    def _visible_standard_module_members(self, mod: ModuleSymbols, same_module: bool) -> list[VbaProjectClassMember]:
        return self._visible_project_members(
            mod, lambda symbol: _is_visible_standard_module_member(symbol, mod, same_module), include_enum_members=True
        )

    def _visible_project_members(
        self,
        mod: ModuleSymbols,
        is_visible: Callable[[VbaSymbol], bool],
        include_enum_members: bool = False,
    ) -> list[VbaProjectClassMember]:
        by_name: dict[str, VbaProjectClassMember] = {}
        for symbol in _project_member_candidate_symbols(mod, include_enum_members):
            if not is_visible(symbol):
                continue
            kind = _project_object_member_kind(symbol)
            if not kind:
                continue
            key = symbol.name.lower()
            existing = by_name.get(key)
            if existing is not None:
                returns = _project_object_member_return_type(symbol)
                if not existing.returns and returns:
                    existing.returns = returns
                writable = _project_object_member_writable(symbol)
                if writable is True:
                    existing.writable = True
                elif existing.writable is None and writable is False:
                    existing.writable = False
                if not existing.write_type:
                    existing.write_type = _project_object_member_write_type(symbol)
                if not existing.signature:
                    existing.signature = _project_object_member_signature(symbol)
                if _is_default_project_object_member(symbol):
                    existing.default_member = True
                existing.attributes = _merge_member_attributes(existing.attributes, symbol.attributes)
                existing.definitions = [
                    *(existing.definitions or []),
                    _project_object_member_definition(symbol),
                ]
                continue
            by_name[key] = VbaProjectClassMember(
                name=symbol.name,
                kind=kind,
                returns=_project_object_member_return_type(symbol),
                signature=_project_object_member_signature(symbol),
                writable=_project_object_member_writable(symbol),
                write_type=_project_object_member_write_type(symbol),
                module_name=mod.module_name,
                visibility=symbol.visibility,
                definitions=[_project_object_member_definition(symbol)],
                default_member=True if _is_default_project_object_member(symbol) else None,
                attributes=_merge_member_attributes(None, symbol.attributes),
            )
        return list(by_name.values())

    def _project_user_type_members(self, module_name: str) -> list[VbaProjectClassMembers]:
        current_lower = module_name.lower()
        out: list[VbaProjectClassMembers] = []
        for mod in self._modules.values():
            same_module = mod.module_name.lower() == current_lower
            for symbol in mod.root.children or []:
                if symbol.kind is not VbaSymbolKind.TYPE:
                    continue
                if not same_module and not _is_type_exported(symbol):
                    continue
                out.append(
                    VbaProjectClassMembers(
                        name=symbol.name,
                        kind="userType",
                        module_name=mod.module_name,
                        exhaustive=True,
                        members=self._user_type_field_members(symbol),
                    )
                )
        return out

    def _user_type_field_members(self, symbol: VbaSymbol) -> list[VbaProjectClassMember]:
        out: list[VbaProjectClassMember] = []
        for field_symbol in symbol.children or []:
            if field_symbol.kind is not VbaSymbolKind.TYPE_FIELD:
                continue
            out.append(
                VbaProjectClassMember(
                    name=field_symbol.name,
                    kind="property",
                    returns=field_symbol.as_type,
                    signature=_user_type_field_signature(field_symbol),
                    writable=True,
                    write_type=field_symbol.as_type,
                    module_name=field_symbol.module_name,
                    definitions=[_project_object_member_definition(field_symbol)],
                )
            )
        return out
