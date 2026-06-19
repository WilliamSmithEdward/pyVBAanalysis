"""Excel host object-model resolver (ported from src/analyzer/host)."""

from __future__ import annotations

from .host_model import (
    HostConstant,
    HostGlobal,
    HostMember,
    HostObjectModel,
    HostType,
    get_excel_object_model,
    get_host_constants,
    get_host_globals,
    get_host_members,
    get_host_type,
    resolve_host_alias,
    resolve_host_constant,
    resolve_host_global,
    resolve_host_member_signature,
    resolve_member_return_type,
)

__all__ = [
    "HostConstant",
    "HostGlobal",
    "HostMember",
    "HostObjectModel",
    "HostType",
    "get_excel_object_model",
    "get_host_constants",
    "get_host_globals",
    "get_host_members",
    "get_host_type",
    "resolve_host_alias",
    "resolve_host_constant",
    "resolve_host_global",
    "resolve_host_member_signature",
    "resolve_member_return_type",
]
