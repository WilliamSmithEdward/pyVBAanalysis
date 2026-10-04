"""Project-type-name resolution (reduced port of completion/typeCompletion.ts).

Ported from xlide_vscode/src/analyzer/completion/typeCompletion.ts. Only the
type-name *resolution* surface is ported - the seam the diagnostics rules call to
decide whether a bare or qualified ``As`` type-name resolves to a single project /
primitive / OLE / host type, to nothing, or to the ``'ambiguous'`` marker (the
ambiguity rule the deferred branches need).

The completion-UX paths (``detectTypePosition`` / ``readPartialTypeName`` /
``resolveTypeCompletions`` and the cursor-context plumbing) are intentionally NOT
ported, nor are the editor-only ``detail`` / ``documentation`` fields - no
diagnostic rule reads them. The project subset reuses the existing
``VbaProjectTypeName`` (with its ``VbaProjectTypeKind``) as input rather than a
parallel ``ProjectTypeName`` dataclass; ``ProjectTypeName`` names it. Upstream's
``TypeCompletionContext`` is passed as its two fields, ``project_types`` and
``model``.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from typing import Literal

from ..host.host_model import (
    HostObjectModel,
    get_excel_object_model,
    get_host_enums,
    host_display_name,
)
from ..identity_cache import IdentityLru
from ..symbols.symbol_model import VbaProjectTypeName

# A project-defined type the caller knows about (class/document/UserForm, Type, or Enum).
ProjectTypeName = VbaProjectTypeName

# Where a candidate type name comes from. The project subset reuses the
# VbaProjectTypeKind values (class/document/userform/enum/userType - identical
# strings); the remaining kinds are completion-source groupings.
TypeCompletionKind = Literal[
    "primitive",
    "external",
    "host",
    "module",
    "class",
    "document",
    "userform",
    "enum",
    "userType",
    "ambiguous",
]


@dataclass(frozen=True, slots=True)
class TypeCompletion:
    """A single resolved type name (editor-only detail/documentation dropped)."""

    name: str
    kind: TypeCompletionKind
    module_name: str | None = None


# Canonical VBA built-in data types valid in an `As` clause (MS-VBAL 5.2.3.1.4 /
# 2.1). `Decimal` is omitted: it is not directly declarable in VBA.
VBA_PRIMITIVE_TYPES: tuple[str, ...] = (
    "Boolean",
    "Byte",
    "Currency",
    "Date",
    "Double",
    "Integer",
    "Long",
    "LongLong",
    "LongPtr",
    "Object",
    "Single",
    "String",
    "Variant",
)

# External interface types available through VBA's default OLE Automation reference.
OLE_AUTOMATION_TYPES: tuple[TypeCompletion, ...] = (
    TypeCompletion(name="IUnknown", kind="external", module_name="stdole"),
)


@dataclass(slots=True)
class _ProjectTypeGroup:
    name: str
    kinds: set[str]
    count: int
    module_name: str | None


def project_type_candidates(
    project_types: Sequence[VbaProjectTypeName],
) -> list[TypeCompletion]:
    """Group project types by lowercased name; a count > 1 collapses to 'ambiguous'.

    This is THE ambiguity rule the deferred type-name branches resolve against: a
    name owned by more than one distinct project type yields the ambiguous marker.
    """
    grouped: dict[str, _ProjectTypeGroup] = {}
    for project_type in project_types:
        key = project_type.name.lower()
        group = grouped.get(key)
        if group is None:
            group = _ProjectTypeGroup(
                name=project_type.name, kinds=set(), count=0, module_name=project_type.module_name
            )
            grouped[key] = group
        group.kinds.add(project_type.kind.value)
        group.count += 1
        if not group.module_name and project_type.module_name:
            group.module_name = project_type.module_name

    out: list[TypeCompletion] = []
    for group in grouped.values():
        if group.count != 1:
            out.append(TypeCompletion(name=group.name, kind="ambiguous"))
            continue
        kind = next(iter(group.kinds))
        out.append(
            TypeCompletion(name=group.name, kind=kind, module_name=group.module_name)  # type: ignore[arg-type]
        )
    return out


def host_type_names(model: HostObjectModel) -> list[str]:
    """Short host type names (e.g. 'Workbook') derived from the host model types,
    the project's own host's first.

    `Range` is in both Word and Excel, and the bare name belongs to whichever
    library the project's host is, which is how VBA resolves it: by the reference
    list with the host at the top. A merged model lists a referenced library's
    keys first, so taking it as it comes handed the name to the wrong application.
    """
    host = f"{host_display_name(model)}.".lower()
    keys = list(model["types"])
    ordered = [k for k in keys if k.lower().startswith(host)] + [
        k for k in keys if not k.lower().startswith(host)
    ]
    out: list[str] = []
    for qualified in ordered:
        short = qualified.split(".")[-1]
        if short:
            out.append(short)
    return out


def type_completion_candidates(
    project_types: Sequence[VbaProjectTypeName] | None = None,
    model: HostObjectModel | None = None,
) -> list[TypeCompletion]:
    """De-duplicated type-name candidates, priority project > primitive > OLE > host."""
    resolved_model = model if model is not None else get_excel_object_model()
    seen: set[str] = set()
    out: list[TypeCompletion] = []

    def add(name: str, kind: TypeCompletionKind, module_name: str | None = None) -> None:
        key = name.lower()
        if key in seen:
            return
        seen.add(key)
        out.append(TypeCompletion(name=name, kind=kind, module_name=module_name))

    # 1. Project-defined types take precedence (can shadow a built-in name).
    for t in project_type_candidates(project_types or []):
        add(t.name, t.kind, t.module_name)
    # 2. VBA built-in data types.
    for name in VBA_PRIMITIVE_TYPES:
        add(name, "primitive")
    # 3. OLE Automation interface types from the default stdole reference.
    for t in OLE_AUTOMATION_TYPES:
        add(t.name, t.kind, t.module_name)
    # 4. Host object-model types, the project's own host's first.
    for name in host_type_names(resolved_model):
        add(name, "host")
    # 5. Host enumerations: `Dim k As XlAxisType` is ordinary VBA. They come last
    # so an object type of the same name wins.
    for entry in get_host_enums(resolved_model):
        add(entry["displayName"], "enum")
    return out


@dataclass(frozen=True, slots=True)
class _QualifiedTypeName:
    qualifier: str
    member: str


def qualified_type_name(name: str) -> _QualifiedTypeName | None:
    dot = name.find(".")
    if dot <= 0 or dot >= len(name) - 1:
        return None
    return _QualifiedTypeName(qualifier=name[:dot], member=name[dot + 1 :])


def project_type_candidates_in_module(
    module_name: str,
    project_types: Sequence[VbaProjectTypeName] | None,
) -> list[TypeCompletion]:
    lower_module = module_name.lower()
    filtered = [
        t
        for t in (project_types or [])
        if t.module_name is not None and t.module_name.lower() == lower_module
    ]
    return [
        TypeCompletion(name=c.name, kind=c.kind, module_name=module_name)
        for c in project_type_candidates(filtered)
    ]


def external_type_candidates_in_module(module_name: str) -> tuple[TypeCompletion, ...]:
    if module_name.lower() != "stdole":
        return ()
    return OLE_AUTOMATION_TYPES


# Host metadata is immutable for a model identity, as in the bare-name index:
# group its keys by library once per model.
_HOST_LIBRARY_INDEXES = IdentityLru()


def _host_library_index(model: HostObjectModel) -> dict[str, dict[str, str]]:
    """The libraries a model can answer for, by lowercased name, each mapping a
    lowercased bare type name to the first spelling of it there. A merged model
    carries one library per reference, which is what lets a Word document name
    `Excel.Range`."""
    cached = _HOST_LIBRARY_INDEXES.get(model)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    libraries: dict[str, dict[str, str]] = {}
    for qualified in model["types"]:
        dot = qualified.find(".")
        if dot <= 0:
            continue
        by_lower = libraries.setdefault(qualified[:dot].lower(), {})
        short = qualified[dot + 1 :]
        lower = short.lower()
        if not short or "." in short or lower in by_lower:
            continue
        by_lower[lower] = short
    return _HOST_LIBRARY_INDEXES.put(libraries, model)  # type: ignore[no-any-return]


def is_creatable_type_completion(candidate: TypeCompletion) -> bool:
    """True when the resolved type can be instantiated with ``New``: only project class and userform types qualify."""
    return candidate.kind == "class" or candidate.kind == "userform"


def resolve_type_name(
    name: str,
    project_types: Sequence[VbaProjectTypeName] | None = None,
    model: HostObjectModel | None = None,
) -> TypeCompletion | None:
    """Resolve a bare or qualified type name to its single candidate, the ambiguous
    marker, or None when nothing matches (the no-false-positive gate).

    A qualified ``Mod.Type`` name searches the module's project types, then the
    external ones, then the host library of that name (``Excel.Range``); a bare
    name searches the full de-duplicated candidate set and returns the single
    match (or the ``'ambiguous'`` marker for a project-type collision). With no
    project types the candidate list never collapses to ambiguous, so the marker
    is reachable only via a real cross-Enum collision.
    """
    qualified = qualified_type_name(name)
    if qualified is not None:
        member_lower = qualified.member.lower()
        project = next(
            (
                c
                for c in project_type_candidates_in_module(qualified.qualifier, project_types)
                if c.name.lower() == member_lower
            ),
            None,
        )
        if project is not None:
            return project
        external = next(
            (
                c
                for c in external_type_candidates_in_module(qualified.qualifier)
                if c.name.lower() == member_lower
            ),
            None,
        )
        if external is not None:
            return external
        resolved_model = model if model is not None else get_excel_object_model()
        library = _host_library_index(resolved_model).get(qualified.qualifier.lower())
        short = library.get(member_lower) if library is not None else None
        return TypeCompletion(name=short, kind="host") if short is not None else None
    return _bare_type_candidate_index(project_types, model).get(name.lower())


# Bare type names are resolved once per As-clause / New-expression check, and
# the candidate list is already de-duplicated first-wins by lowercased name, so
# it collapses losslessly into a dict - built once per (project types, model)
# identity instead of rebuilt and scanned linearly per lookup.
_BARE_TYPE_INDEX_CACHE = IdentityLru()
_NO_PROJECT_TYPES: tuple[VbaProjectTypeName, ...] = ()


def _bare_type_candidate_index(
    project_types: Sequence[VbaProjectTypeName] | None,
    model: HostObjectModel | None,
) -> dict[str, TypeCompletion]:
    types_key = project_types if project_types is not None else _NO_PROJECT_TYPES
    resolved_model = model if model is not None else get_excel_object_model()
    cached = _BARE_TYPE_INDEX_CACHE.get(types_key, resolved_model)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    # The candidate list is already unique per lowercased name (first-wins dedup
    # inside type_completion_candidates), so insertion order cannot matter here.
    index = {c.name.lower(): c for c in type_completion_candidates(project_types, model)}
    return _BARE_TYPE_INDEX_CACHE.put(index, types_key, resolved_model)  # type: ignore[no-any-return]
