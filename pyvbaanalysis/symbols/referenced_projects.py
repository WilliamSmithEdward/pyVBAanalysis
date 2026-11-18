"""What another VBA project offers the project that references it.

Ported from xlide_vscode/src/analyzer/symbols/referencedProjects.ts (XLIDE issue
#1503). Measured through the VBE oracle in Excel, Word, PowerPoint and Access 16.0
with a referenced project ProbeLibrary whose standard module Service has a Public
Function Value:

- `ProbeLibrary.Service.Value()`, `ProbeLibrary.Value()`, `Service.Value()` and
  `Value()` all compile, in any case and bracketed.
- With `Option Private Module` in Service, none does: the qualified forms are
  "Method or data member not found", `Service.Value()` "Variable not defined" and
  `Value()` "Sub or Function not defined". A Private Function is the same as a
  private module for it.
- A class of the referenced project is no type the referring one can name unless
  the class is exposed (VB_Exposed, Instancing PublicNotCreatable): `Dim t As
  ProbeLibrary.Thing` is "User-defined type not defined".
- The referring project's own names come first: its own Value() wins.
- An exposed class is a type another project can name but not create: `New
  ProbeLibrary.Widget` is "Invalid use of New keyword". Its name is a value only as
  a default instance (VB_PredeclaredId): otherwise `v = ProbeLibrary.Widget` is
  "Method or data member not found" and `v = Widget` "Variable not defined" (XLIDE
  issue #1602).

The port reads no referenced file itself: a caller that has the referenced
project's modules builds its surface with referenced_project_surface and passes it
as AnalyzeModuleOptions.referenced_project_surfaces.
"""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field
from typing import Literal

from ..host.library_type_names import library_type_names
from ..identity_cache import IdentityLru
from ..lexer.token_helpers import token_name
from ..lexer.token_kinds import TokenKind
from ..lexer.tokenize import tokenize
from .project_index import ModuleInput, ProjectIndex, is_default_project_object_member
from .symbol_model import (
    ModuleSymbolKind,
    SymbolVisibility,
    VbaSymbol,
    VbaSymbolKind,
    procedure_signature_from_symbol,
)

ReferencedObjectMemberKind = Literal["method", "property"]


@dataclass(eq=False, slots=True)
class ReferencedObjectMember:
    """A Public member of an exposed class, or of a visible standard module, of a
    referenced project."""

    name: str
    kind: ReferencedObjectMemberKind
    # Whether it declares parameters, so arguments in a call are its own.
    takes_arguments: bool = False
    # The exposed class of the same project it returns, `Public Child As Widget`.
    returns_class: str | None = None


@dataclass(frozen=True, eq=False, slots=True)
class ReferencedEnum:
    """A Public Enum of a referenced project's module."""

    name: str
    members: tuple[str, ...]


@dataclass(frozen=True, eq=False, slots=True)
class ReferencedProjectModule:
    """One module of a referenced project, as the referring project sees it."""

    name: str
    kind: ModuleSymbolKind
    # Whether the referring project can name the module at all.
    visible: bool
    # The names a visible standard module offers, its procedures, variables,
    # constants and enum members.
    members: tuple[str, ...] = ()
    # The types it offers by name: its Public Types and Enums, or the class itself
    # when exposed.
    types: tuple[str, ...] = ()
    # Its Public Enums with their members: `WidgetColor.wcRed`, and in a class
    # `ProbeLibrary.wcRed`.
    enums: tuple[ReferencedEnum, ...] | None = None
    # Whether an exposed class's name is its default instance.
    predeclared: bool | None = None
    # What an object of an exposed class offers another project.
    object_members: tuple[ReferencedObjectMember, ...] | None = None
    # A visible standard module's Public members, with the exposed class each
    # returns (XLIDE issue #1623).
    module_members: tuple[ReferencedObjectMember, ...] | None = None
    # Whether the class has a default member, which an index into an object reaches.
    default_member: bool | None = None
    # The exposed class of the same project the default member returns, where it
    # takes the index's arguments itself: `ws(1)` on a Widgets whose default
    # `Item(ByVal i As Long) As Widget` is a Widget (XLIDE issue #1623).
    default_member_class: str | None = None


@dataclass(frozen=True, eq=False, slots=True)
class ReferencedProjectSurface:
    """A referenced VBA project that was read, under the name the reference gives it.
    Surfaces are snapshots: replace the object when a referenced project changes."""

    name: str
    modules: tuple[ReferencedProjectModule, ...] = field(default_factory=tuple)


@dataclass(frozen=True, slots=True)
class ReferencedProjectModuleSource:
    """A module of the referenced project as its file holds it, attribute header
    included."""

    name: str
    kind: ModuleSymbolKind
    source: str


_OPTION_PRIVATE_MODULE = re.compile(r"^[ \t]*Option[ \t]+Private[ \t]+Module\b", re.IGNORECASE | re.MULTILINE)
_EXPOSED = re.compile(r"^[ \t]*Attribute[ \t]+VB_Exposed[ \t]*=[ \t]*True\b", re.IGNORECASE | re.MULTILINE)
_PREDECLARED = re.compile(r"^[ \t]*Attribute[ \t]+VB_PredeclaredId[ \t]*=[ \t]*True\b", re.IGNORECASE | re.MULTILINE)
# No module of the referenced project: < is no identifier character.
_OUTSIDE = "<outside>"


def referenced_project_surface(
    name: str, modules: Sequence[ReferencedProjectModuleSource]
) -> ReferencedProjectSurface:
    """The surface of a referenced project, from its modules' sources."""
    index = ProjectIndex()
    for mod in modules:
        try:
            index.set_module(ModuleInput(mod.name, mod.kind, mod.source))
        except Exception:
            # A module that does not parse offers nothing it can be held to.
            continue
    # What a module of another project could see, were it in this one.
    visible_symbols = index.visible_identifier_symbols(_OUTSIDE)
    exposed_classes = {
        mod.name.lower(): mod.name
        for mod in modules
        if mod.kind is ModuleSymbolKind.CLASS and _EXPOSED.search(mod.source)
    }

    def children(module_name: str) -> list[VbaSymbol]:
        symbols = index.get_module(module_name)
        return list(symbols.root.children or []) if symbols is not None else []

    def surface_module(mod: ReferencedProjectModuleSource) -> ReferencedProjectModule:
        if mod.kind is ModuleSymbolKind.STANDARD:
            visible = _OPTION_PRIVATE_MODULE.search(mod.source) is None
            own = (
                [symbol for symbol in visible_symbols if symbol.module_name.lower() == mod.name.lower()]
                if visible
                else []
            )
            return ReferencedProjectModule(
                name=mod.name,
                kind=mod.kind,
                visible=visible,
                members=tuple(
                    symbol.name for symbol in own if symbol.kind not in (VbaSymbolKind.TYPE, VbaSymbolKind.ENUM)
                ),
                enums=_public_enums(own),
                types=tuple(symbol.name for symbol in own if symbol.kind in (VbaSymbolKind.TYPE, VbaSymbolKind.ENUM)),
                # What its Public procedures and variables give, so a chain through
                # one reaches the class it returns: `ProbeLibrary.NewWidget().Nope`.
                module_members=_object_members_of(children(mod.name), name, exposed_classes) if visible else None,
            )
        if _EXPOSED.search(mod.source) is None:
            return ReferencedProjectModule(name=mod.name, kind=mod.kind, visible=False)
        # A document module or form is a class module too: its name is a type and
        # its Public Enums are offered as an exposed class's are.
        symbols = children(mod.name)
        enums = _public_enums(symbols)
        if mod.kind is not ModuleSymbolKind.CLASS:
            # A document module or form's members are its host object's too: unjudged.
            return ReferencedProjectModule(
                name=mod.name, kind=mod.kind, visible=True,
                types=(mod.name, *(decl.name for decl in enums)), enums=enums,
            )
        return ReferencedProjectModule(
            name=mod.name,
            kind=mod.kind,
            visible=True,
            types=(mod.name, *(decl.name for decl in enums)),
            enums=enums,
            predeclared=_PREDECLARED.search(mod.source) is not None,
            object_members=_object_members_of(symbols, name, exposed_classes),
            default_member=any(is_default_project_object_member(symbol) for symbol in symbols),
            default_member_class=_default_member_class_of(symbols, name, exposed_classes),
        )

    return ReferencedProjectSurface(name=name, modules=tuple(surface_module(mod) for mod in modules))


@dataclass(frozen=True, slots=True)
class ReferencedProjectNames:
    """Name lookups shared by all rules analyzing the same surface."""

    # First module of each name; callers still check its visibility.
    modules: Mapping[str, ReferencedProjectModule]
    # Types offered by the surface, by lowercased name.
    types: Mapping[str, ReferencedProjectModule]
    # Members of visible standard modules and exposed classes' Enum members,
    # available after the project name.
    values: frozenset[str]
    # The Public Enums of visible modules, by lowercased name.
    enums: Mapping[str, ReferencedEnum]
    # Visible module names and their members, available unqualified.
    bare_names: frozenset[str]


# Upstream holds the index in a WeakMap keyed by the surface; surfaces are
# snapshots, so a replaced one builds a new index and the old is not retained.
_PROJECT_NAMES = IdentityLru(capacity=16)


def referenced_project_names(surface: ReferencedProjectSurface) -> ReferencedProjectNames:
    """Built once per surface object."""
    cached = _PROJECT_NAMES.get(surface)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    modules: dict[str, ReferencedProjectModule] = {}
    types: dict[str, ReferencedProjectModule] = {}
    values: set[str] = set()
    bare_names: set[str] = set()
    enums: dict[str, ReferencedEnum] = {}
    for module in surface.modules:
        key = module.name.lower()
        modules.setdefault(key, module)
        for type_name in module.types:
            types.setdefault(type_name.lower(), module)
        if not module.visible:
            continue
        # A class with no default instance is a type, never a value by name.
        if module.kind is not ModuleSymbolKind.CLASS or module.predeclared:
            bare_names.add(key)
        for member in module.members:
            bare_names.add(member.lower())
            if module.kind is ModuleSymbolKind.STANDARD:
                values.add(member.lower())
        for decl in module.enums or ():
            bare_names.add(decl.name.lower())
            enums.setdefault(decl.name.lower(), decl)
            # A standard module's Enum members are among its members already.
            for member in () if module.kind is ModuleSymbolKind.STANDARD else decl.members:
                bare_names.add(member.lower())
                values.add(member.lower())
    names = ReferencedProjectNames(
        modules=modules, types=types, values=frozenset(values), enums=enums, bare_names=frozenset(bare_names)
    )
    return _PROJECT_NAMES.put(names, surface)  # type: ignore[no-any-return]


@dataclass(frozen=True, slots=True)
class ReferencedLookupOrder:
    """What decides which library a bare name reaches."""

    # The project's references, in order; None where they are not known.
    libraries: Sequence[str] | None
    host_name: str


def referenced_project_reached(
    lower: str,
    surfaces: Sequence[ReferencedProjectSurface] | None,
    order: ReferencedLookupOrder,
    offers: Callable[[ReferencedProjectNames], bool],
) -> ReferencedProjectSurface | None:
    """The read project a bare name reaches. A name is looked up in VBA, the host's
    library and then the references in the project's order, and the first that has
    it wins. None where a library before the project might have it: one whose type
    names are known and hold it, or one XLIDE did not read, whose names are unknown.
    `offers` says whether a read project has the name."""
    if not surfaces or not order.libraries:
        return None
    for library in ["VBA", order.host_name, *order.libraries]:
        key = library.lower()
        surface = next((candidate for candidate in surfaces if candidate.name.lower() == key), None)
        if surface is not None:
            if offers(referenced_project_names(surface)):
                return surface
            continue
        known = library_type_names(library)
        if known is None or lower in known:
            return None
    return None


def _is_plain_name(part: str) -> bool:
    """Upstream's `/^[\\p{L}_][\\p{L}\\p{M}0-9_]*$/u`, by Unicode category."""
    if not part:
        return False
    first = part[0]
    if first != "_" and not unicodedata.category(first).startswith("L"):
        return False
    return all(
        ch == "_" or ch in "0123456789" or unicodedata.category(ch)[0] in ("L", "M") for ch in part[1:]
    )


def declared_type_is_referenced_enum(
    as_type: str,
    surfaces: Sequence[ReferencedProjectSurface] | None,
    order: ReferencedLookupOrder | None,
) -> bool:
    """Whether a declared type names a read referenced project's Public Enum:
    `ProbeLibrary.WidgetColor`, or a bare `WidgetColor` the lookup reaches. A
    variable of it is a Long, so a member on it is "Invalid qualifier" (XLIDE issue
    #1623). A name the project itself declares is the caller's to rule out."""
    parts = [part.strip().lower() for part in as_type.split(".")]
    if not surfaces or len(parts) > 2 or any(not _is_plain_name(part) for part in parts):
        return False
    name = parts[-1]
    project: ReferencedProjectSurface | None
    if len(parts) == 2:
        project = next((surface for surface in surfaces if surface.name.lower() == parts[0]), None)
    else:
        project = (
            referenced_project_reached(name, surfaces, order, lambda offered: name in offered.types)
            if order is not None
            else None
        )
    return project is not None and name in referenced_project_names(project).enums


def _default_member_class_of(
    symbols: Sequence[VbaSymbol], project_name: str, exposed_classes: Mapping[str, str]
) -> str | None:
    """The class a default member gives, when an index goes to it: a Public function
    or Property Get that declares parameters and returns an exposed class of the
    project. A default member taking none passes the index on to what it returns,
    which is not followed."""
    member = next(
        (
            symbol
            for symbol in symbols
            if is_default_project_object_member(symbol)
            and symbol.kind not in (VbaSymbolKind.PROPERTY_LET, VbaSymbolKind.PROPERTY_SET)
        ),
        None,
    )
    if member is None or _object_member_kind(member) is None or _param_count(member) == 0:
        return None
    return _exposed_class_named(member.as_type, project_name, exposed_classes)


def _param_count(symbol: VbaSymbol) -> int:
    signature = procedure_signature_from_symbol(symbol)
    return len(signature.params) if signature is not None else 0


def _object_members_of(
    symbols: Sequence[VbaSymbol], project_name: str, exposed_classes: Mapping[str, str]
) -> tuple[ReferencedObjectMember, ...]:
    """What an object of a class offers another project: Public procedures and
    variables. Friend, Private and events are none."""
    out: dict[str, ReferencedObjectMember] = {}
    for symbol in symbols:
        kind = _object_member_kind(symbol)
        if kind is None:
            continue
        key = symbol.name.lower()
        member = out.get(key)
        if member is None:
            member = ReferencedObjectMember(name=symbol.name, kind=kind)
            out[key] = member
        # A Property Let or Set takes the value; what the member gives is its Get's.
        if symbol.kind not in (VbaSymbolKind.PROPERTY_LET, VbaSymbolKind.PROPERTY_SET):
            member.returns_class = _exposed_class_named(symbol.as_type, project_name, exposed_classes)
            member.takes_arguments = _param_count(symbol) > 0
    return tuple(out.values())


def _object_member_kind(symbol: VbaSymbol) -> ReferencedObjectMemberKind | None:
    hidden = symbol.visibility in (SymbolVisibility.PRIVATE, SymbolVisibility.FRIEND)
    if symbol.kind in (VbaSymbolKind.SUB, VbaSymbolKind.FUNCTION):
        return None if hidden else "method"
    if symbol.kind in (VbaSymbolKind.PROPERTY_GET, VbaSymbolKind.PROPERTY_LET, VbaSymbolKind.PROPERTY_SET):
        return None if hidden else "property"
    if symbol.kind is VbaSymbolKind.MODULE_VARIABLE:
        return "property" if symbol.visibility in (SymbolVisibility.PUBLIC, SymbolVisibility.GLOBAL) else None
    return None


def _exposed_class_named(
    as_type: str | None, project_name: str, exposed_classes: Mapping[str, str]
) -> str | None:
    """The same project's exposed class a declared type names, bare or after the
    project's name."""
    names = declared_type_names(as_type) if as_type is not None else None
    if names is None or len(names) > 2 or (len(names) == 2 and names[0].lower() != project_name.lower()):
        return None
    return exposed_classes.get(names[-1].lower())


def declared_type_names(text: str) -> list[str] | None:
    """The names of a declared type as the lexer reads them, `ProbeLibrary.Widget`:
    across any line continuation it accepts, and with brackets off. None unless the
    text is names joined by dots."""
    tokens = [token for token in tokenize(text) if token.kind is not TokenKind.NEWLINE]
    names: list[str] = []
    for i in range(0, len(tokens), 2):
        name = token_name(tokens[i])
        if name is None or (i > 0 and tokens[i - 1].raw_text != "."):
            return None
        names.append(name)
    return names if names and len(tokens) % 2 == 1 else None


def _public_enums(symbols: Sequence[VbaSymbol]) -> tuple[ReferencedEnum, ...]:
    """The Public Enums among a module's symbols, each with its members."""
    return tuple(
        ReferencedEnum(
            name=symbol.name,
            members=tuple(member.name for member in symbol.children or () if member.kind is VbaSymbolKind.ENUM_MEMBER),
        )
        for symbol in symbols
        if symbol.kind is VbaSymbolKind.ENUM and symbol.visibility is not SymbolVisibility.PRIVATE
    )
