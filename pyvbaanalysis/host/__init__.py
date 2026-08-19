"""Host object-model resolvers and the host registry (ported from src/analyzer/host)."""

from __future__ import annotations

from .host_model import (
    HostConstant,
    HostGlobal,
    HostMember,
    HostObjectModel,
    HostType,
    application_member_names,
    get_access_object_model,
    get_excel_object_model,
    get_host_constants,
    get_host_globals,
    get_host_members,
    get_host_type,
    get_powerpoint_object_model,
    get_word_object_model,
    is_host_member_name,
    resolve_host_alias,
    resolve_host_constant,
    resolve_host_global,
    resolve_host_member_signature,
    resolve_member_return_type,
)
from .host_registry import (
    EMPTY_HOST_MODEL,
    VBA_HOST_TOKENS,
    host_object_model_for_token,
    host_token_for_file_name,
)

__all__ = [
    "EMPTY_HOST_MODEL",
    "HostConstant",
    "HostGlobal",
    "HostMember",
    "HostObjectModel",
    "HostType",
    "VBA_HOST_TOKENS",
    "application_member_names",
    "get_access_object_model",
    "get_excel_object_model",
    "get_host_constants",
    "get_host_globals",
    "get_host_members",
    "get_host_type",
    "get_powerpoint_object_model",
    "get_word_object_model",
    "host_object_model_for_token",
    "host_token_for_file_name",
    "is_host_member_name",
    "resolve_host_alias",
    "resolve_host_constant",
    "resolve_host_global",
    "resolve_host_member_signature",
    "resolve_member_return_type",
]
