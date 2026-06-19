"""Host object-model resolver — pure functions over a HostObjectModel.

Ported from xlide_vscode/src/analyzer/host/hostModel.ts. The Excel object model
itself is vendored as data/excel_host_model.json, mechanically extracted from the
generated XLIDE host modules (tools/extract_host_model.mjs) — never hand-
transcribed, so the member surfaces stay exact (the no-false-positive contract
for member-not-found depends on the exhaustive set being complete). Defaults to
the Excel model but accepts any model dict for testing.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import TypedDict

_DATA_DIR = Path(__file__).resolve().parent.parent / "data"


class HostMember(TypedDict, total=False):
    name: str
    kind: str  # 'property' | 'method' | 'event' | ...
    returns: str
    returnsAnyOf: list[str]
    signature: str


class HostConstant(TypedDict, total=False):
    name: str
    type: str
    value: str | int


class HostType(TypedDict, total=False):
    displayName: str
    members: list[HostMember]
    provenance: str
    exhaustive: bool


class HostObjectModel(TypedDict):
    source: str
    aliases: dict[str, str]
    globals: dict[str, str]
    constants: dict[str, HostConstant]
    types: dict[str, HostType]
    memberSignatures: dict[str, dict[str, str]]


@dataclass(frozen=True, slots=True)
class HostGlobal:
    name: str
    type: str


@lru_cache(maxsize=1)
def get_excel_object_model() -> HostObjectModel:
    raw = json.loads((_DATA_DIR / "excel_host_model.json").read_text(encoding="utf-8"))
    return raw  # type: ignore[no-any-return]


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


_MODEL_INDEX_CACHE: dict[int, _HostModelIndex] = {}
_CONSTANT_INDEX_CACHE: dict[int, dict[str, HostConstant]] = {}


def _host_model_index(model: HostObjectModel) -> _HostModelIndex:
    cached = _MODEL_INDEX_CACHE.get(id(model))
    if cached is not None:
        return cached
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
    index = _HostModelIndex(members_by_type, type_keys_by_lower)
    _MODEL_INDEX_CACHE[id(model)] = index
    return index


def _host_constant_index(model: HostObjectModel) -> dict[str, HostConstant]:
    cached = _CONSTANT_INDEX_CACHE.get(id(model))
    if cached is not None:
        return cached
    index = {key.lower(): constant for key, constant in (model.get("constants") or {}).items()}
    _CONSTANT_INDEX_CACHE[id(model)] = index
    return index


def get_host_type(qualified: str, model: HostObjectModel | None = None) -> HostType | None:
    """The type metadata for a qualified type name (e.g. 'Excel.Range')."""
    return _default(model)["types"].get(qualified)


def get_host_members(qualified: str, model: HostObjectModel | None = None) -> list[HostMember]:
    """The object-access members of a qualified type, or an empty list if unknown."""
    type_index = _host_model_index(_default(model)).members_by_type.get(qualified)
    return type_index.members if type_index is not None else []


def resolve_host_global(name: str, model: HostObjectModel | None = None) -> str | None:
    """A host-injected global identifier (ThisWorkbook, Application, ...) -> qualified type."""
    lower = name.lower()
    for key, type_ in _default(model)["globals"].items():
        if key.lower() == lower:
            return type_
    return None


def resolve_host_constant(name: str, model: HostObjectModel | None = None) -> HostConstant | None:
    """A host enum constant such as xlUp or xlCalculationAutomatic (case-insensitive)."""
    return _host_constant_index(_default(model)).get(name.lower())


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
