"""Host object-model resolver, pure functions over a HostObjectModel.

Ported from xlide_vscode/src/analyzer/host/hostModel.ts. The Excel object model
itself is vendored as data/excel_host_model.json, mechanically extracted from the
generated XLIDE host modules (tools/extract_host_model.mjs), never hand-
transcribed, so the member surfaces stay exact (the no-false-positive contract
for member-not-found depends on the exhaustive set being complete). Defaults to
the Excel model but accepts any model dict for testing.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from collections.abc import Mapping
from functools import lru_cache
from pathlib import Path
from typing import TypedDict

from ..identity_cache import IdentityLru

_DATA_DIR = Path(__file__).resolve().parent.parent / "data"


class HostMember(TypedDict, total=False):
    name: str
    kind: str  # 'property' | 'method' | 'event' | ...
    returns: str
    returnsAnyOf: list[str]
    signature: str
    # Marked hidden in the type library: it resolves like any other member, but
    # the editor never offers it (XLIDE issue #56).
    hidden: bool


class HostConstant(TypedDict, total=False):
    name: str
    type: str
    value: str | int


class HostType(TypedDict, total=False):
    displayName: str
    members: list[HostMember]
    provenance: str
    exhaustive: bool


class _HostObjectModelCore(TypedDict):
    # Mapping rather than dict: a model is read-only data shared across every
    # analysis pass, and the empty model is a frozen singleton, so nothing may
    # mutate one in place.
    source: str
    aliases: Mapping[str, str]
    globals: Mapping[str, str]
    constants: Mapping[str, HostConstant]
    types: Mapping[str, HostType]
    memberSignatures: Mapping[str, Mapping[str, str]]


class HostObjectModel(_HostObjectModelCore, total=False):
    """A host object model. The optional keys arrived with XLIDE 10.x; inheriting
    from a total=False subclass keeps them optional on Python 3.10, which has no
    typing.NotRequired."""

    # The library's own name, which is the qualifier VBA writes (`Excel`, `Word`).
    # A merged model carries the project's own host here.
    hostName: str
    # The library's hidden global interface (`Excel.Global`), whose members VBA
    # binds bare for anyone who references the library. Access has none.
    globalType: str | None
    # Enumerations by name. In a merged model, a referenced library's entries
    # carry a `library` key naming where they came from.
    enums: Mapping[str, Mapping[str, object]]


@dataclass(frozen=True, slots=True)
class HostGlobal:
    name: str
    type: str


@lru_cache(maxsize=None)
def _load_host_model(file_name: str) -> HostObjectModel:
    """Load and cache one vendored host object model from the data directory."""
    raw = json.loads((_DATA_DIR / file_name).read_text(encoding="utf-8"))
    return raw  # type: ignore[no-any-return]


def get_excel_object_model() -> HostObjectModel:
    """The vendored Excel host object model (data/excel_host_model.json)."""
    return _load_host_model("excel_host_model.json")


def get_word_object_model() -> HostObjectModel:
    """The vendored Word host object model (data/word_host_model.json)."""
    return _load_host_model("word_host_model.json")


def get_powerpoint_object_model() -> HostObjectModel:
    """The vendored PowerPoint host object model (data/powerpoint_host_model.json)."""
    return _load_host_model("powerpoint_host_model.json")


def get_access_object_model() -> HostObjectModel:
    """The vendored Access host object model (data/access_host_model.json)."""
    return _load_host_model("access_host_model.json")


def get_vb6_object_model() -> HostObjectModel:
    """The vendored VB6 object model (data/vb6_host_model.json): the VB runtime's
    objects and constants, and the VB library of App, Screen, Printer, Form and the
    intrinsic controls. It offers and describes; it never proves a member absent."""
    return _load_host_model("vb6_host_model.json")


def _default(model: HostObjectModel | None) -> HostObjectModel:
    return model if model is not None else get_excel_object_model()


def _is_object_access_member(member: HostMember) -> bool:
    return member.get("kind") != "event"


@dataclass(slots=True)
class _HostTypeIndex:
    members: list[HostMember]
    by_lower_name: dict[str, HostMember]
    raw_by_lower_name: dict[str, HostMember]


@dataclass(slots=True)
class _HostModelIndex:
    members_by_type: dict[str, _HostTypeIndex]
    type_keys_by_lower: dict[str, str]
    globals_by_lower: dict[str, str]
    enums_by_lower: dict[str, Mapping[str, object]]
    constants_by_enum: dict[str, list[HostConstant]]


# Identity-keyed, and deliberately IdentityLru rather than a bare dict[int, ...]:
# these memos are keyed by the identity of a model the caller owns, so the cache
# must hold that model alive (a plain id() key is recyclable once the model is
# collected, which would serve one model's index for another) and must stay
# bounded (a caller that builds a model per call would otherwise grow it without
# limit). Capacity covers the four vendored hosts plus the empty model with room
# to spare.
_MODEL_INDEX_CACHE = IdentityLru(capacity=8)
_CONSTANT_INDEX_CACHE = IdentityLru(capacity=8)


def _host_model_index(model: HostObjectModel) -> _HostModelIndex:
    cached = _MODEL_INDEX_CACHE.get(model)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    members_by_type: dict[str, _HostTypeIndex] = {}
    type_keys_by_lower: dict[str, str] = {}
    for key, type_ in model["types"].items():
        key_lower = key.lower()
        if key_lower not in type_keys_by_lower:
            type_keys_by_lower[key_lower] = key
        members: list[HostMember] = []
        by_lower_name: dict[str, HostMember] = {}
        raw_by_lower_name: dict[str, HostMember] = {}
        for member in type_.get("members") or []:
            lower = member["name"].lower()
            if lower not in raw_by_lower_name:
                raw_by_lower_name[lower] = member
            if not _is_object_access_member(member):
                continue
            members.append(member)
            if lower not in by_lower_name:
                by_lower_name[lower] = member
        members_by_type[key] = _HostTypeIndex(members, by_lower_name, raw_by_lower_name)
    globals_by_lower: dict[str, str] = {}
    for key, global_type in (model.get("globals") or {}).items():
        key_lower = key.lower()
        if key_lower not in globals_by_lower:
            globals_by_lower[key_lower] = global_type
    enums_by_lower: dict[str, Mapping[str, object]] = {}
    for entry in (model.get("enums") or {}).values():
        lower = str(entry.get("displayName", "")).lower()
        if lower and lower not in enums_by_lower:
            enums_by_lower[lower] = entry
    # An enum's members are the constants that name it, so the two can never
    # disagree and the generated tables stay a single list.
    constants_by_enum: dict[str, list[HostConstant]] = {}
    for constant in (model.get("constants") or {}).values():
        enum_type = constant.get("type")
        if enum_type:
            constants_by_enum.setdefault(enum_type.lower(), []).append(constant)
    index = _HostModelIndex(
        members_by_type, type_keys_by_lower, globals_by_lower, enums_by_lower, constants_by_enum
    )
    return _MODEL_INDEX_CACHE.put(index, model)  # type: ignore[no-any-return]


def _host_constant_index(model: HostObjectModel) -> dict[str, HostConstant]:
    cached = _CONSTANT_INDEX_CACHE.get(model)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    index = {key.lower(): constant for key, constant in (model.get("constants") or {}).items()}
    return _CONSTANT_INDEX_CACHE.put(index, model)  # type: ignore[no-any-return]


def get_host_type(qualified: str, model: HostObjectModel | None = None) -> HostType | None:
    """The type metadata for a qualified type name (e.g. 'Excel.Range')."""
    return _default(model)["types"].get(qualified)


def get_host_members(qualified: str, model: HostObjectModel | None = None) -> list[HostMember]:
    """The object-access members of a qualified type, or an empty list if unknown."""
    type_index = _host_model_index(_default(model)).members_by_type.get(qualified)
    return type_index.members if type_index is not None else []


_HOST_MEMBER_NAMES_CACHE = IdentityLru(capacity=8)


def is_host_member_name(name: str, model: HostObjectModel | None = None) -> bool:
    """True when ``name`` is a member of ANY type in the host object model.

    Callers reasoning about a late-bound receiver cannot know its runtime type,
    so they need the weaker question "could this name legally dispatch somewhere
    in the host model at all?". Answering yes keeps them quiet; the set is
    deliberately broad for that reason. Case-insensitive."""
    resolved = _default(model)
    names = _HOST_MEMBER_NAMES_CACHE.get(resolved)
    if names is None:
        names = set()
        for type_index in _host_model_index(resolved).members_by_type.values():
            names.update(type_index.by_lower_name)
            names.update(type_index.raw_by_lower_name)
        _HOST_MEMBER_NAMES_CACHE.put(names, resolved)
    return name.lower() in names


def resolve_host_global(name: str, model: HostObjectModel | None = None) -> str | None:
    """A host-injected global identifier (ThisWorkbook, Application, ...) -> qualified type.

    O(1) via the cached, lowercase-keyed index (first-wins), instead of a linear
    scan on every identifier-token lookup."""
    resolved = _default(model)
    return _host_model_index(resolved).globals_by_lower.get(name.lower())


def resolve_host_constant(name: str, model: HostObjectModel | None = None) -> HostConstant | None:
    """A host enum constant such as xlUp or xlCalculationAutomatic (case-insensitive)."""
    return _host_constant_index(_default(model)).get(name.lower())


def resolve_host_enum(name: str, model: HostObjectModel | None = None) -> Mapping[str, object] | None:
    """The enumeration named `name`, case-insensitively. VBA accepts an enum name as
    a declared type (`Dim k As XlAxisType`) and as a qualifier
    (`XlAxisType.xlCategory`)."""
    if not name:
        return None
    return _host_model_index(_default(model)).enums_by_lower.get(name.strip().lower())


def get_host_enum_members(enum_name: str, model: HostObjectModel | None = None) -> list[HostConstant]:
    """The constants belonging to one enumeration, in declaration order."""
    return _host_model_index(_default(model)).constants_by_enum.get(enum_name.lower(), [])


def resolve_host_global_member(name: str, model: HostObjectModel | None = None) -> HostMember | None:
    """A bare identifier as a member of the host's hidden Global interface, the
    surface VBA calls unqualified: Word's InchesToPoints, Excel's Union (XLIDE issue
    #34). Object-access members only, never an event; None when the model carries
    no Global type or the name is not among its members."""
    resolved = _default(model)
    global_type = resolved.get("globalType")
    # A leading underscore marks a hidden dispatch name, never called by name.
    if not global_type or name.startswith("_"):
        return None
    type_index = _host_model_index(resolved).members_by_type.get(global_type)
    return type_index.by_lower_name.get(name.lower()) if type_index is not None else None


def resolve_host_member_signature(
    qualified: str, member: str, model: HostObjectModel | None = None
) -> str | None:
    """The verified call signature for a callable member of a host type, or None."""
    resolved = _default(model)
    lower = member.lower()
    type_index = _host_model_index(resolved).members_by_type.get(qualified)
    raw_member = type_index.raw_by_lower_name.get(lower) if type_index is not None else None
    if raw_member is not None and raw_member.get("kind") == "event":
        return None
    by_lower = type_index.by_lower_name.get(lower) if type_index is not None else None
    if by_lower is not None and by_lower.get("signature"):
        return by_lower["signature"]
    return (resolved.get("memberSignatures") or {}).get(qualified, {}).get(lower)


def get_host_globals(model: HostObjectModel | None = None) -> list[HostGlobal]:
    """All host-injected globals (canonical casing)."""
    return [HostGlobal(name=name, type=type_) for name, type_ in _default(model)["globals"].items()]


def get_host_constants(model: HostObjectModel | None = None) -> list[HostConstant]:
    """All host enum constants (canonical casing)."""
    return list((_default(model).get("constants") or {}).values())


def resolve_host_alias(type_name: str, model: HostObjectModel | None = None) -> str | None:
    """A declared type name (bare or qualified) -> qualified host type, or None."""
    if not type_name:
        return None
    resolved = _default(model)
    trimmed = type_name.strip()
    lower = trimmed.lower()
    if trimmed in resolved["types"]:
        return trimmed
    key = _host_model_index(resolved).type_keys_by_lower.get(lower)
    if key is not None:
        return key
    return resolved["aliases"].get(lower)


def resolve_member_return_type(
    qualified: str, member_name: str, model: HostObjectModel | None = None
) -> str | None:
    """The qualified type produced by accessing `member_name` on `qualified`, or None."""
    type_index = _host_model_index(_default(model)).members_by_type.get(qualified)
    member = type_index.by_lower_name.get(member_name.lower()) if type_index is not None else None
    return member.get("returns") if member is not None else None


_APPLICATION_MEMBER_NAMES = IdentityLru(capacity=8)


def application_member_names(model: HostObjectModel | None = None) -> frozenset[str]:
    """Lowercased member names of the host Application global (for the implicit-member
    negative lookup: `Calculate`, `Range`, ... are unqualified Application members).

    Keyed per model, so a Word caller gets Word's set and a host with no model
    injects nothing at all. Each vendored model is a cached singleton, so
    identity keying is stable across calls.

    Where the host has a Global interface (Excel, Word, PowerPoint), that is what
    VBA really calls bare, and resolve_host_global_member answers for it, hidden
    members and all. Application's documented members stand in for it here because
    they match it closely; its hidden ones do not, so they stay out. `Save` is a
    hidden method of Excel's `_Application` and no member of `_Global`, so a bare
    `Save` is "Sub or Function not defined". Access has no Global: its type library
    makes Application itself the object VBA binds bare, so there every member of
    it, hidden or not, is in scope.
    """
    resolved = _default(model)
    cached = _APPLICATION_MEMBER_NAMES.get(resolved)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    app_type = resolve_host_global("Application", resolved)
    members = get_host_members(app_type, resolved) if app_type is not None else []
    global_answers = resolved.get("globalType") is not None
    return _APPLICATION_MEMBER_NAMES.put(  # type: ignore[no-any-return]
        frozenset(
            member["name"].lower()
            for member in members
            if not (global_answers and member.get("hidden"))
        ),
        resolved,
    )
