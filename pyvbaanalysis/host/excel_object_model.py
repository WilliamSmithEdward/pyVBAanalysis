"""Ported from xlide_vscode/src/analyzer/host/excelObjectModel.ts, the parts that
are not the model itself.

The models are built upstream and vendored as data/*_host_model.json by
tools/extract_host_model.mjs, so the builders and the reference tables behind them
(officeReferenceTypes.ts, powerpointObjectModelData.ts, hostBoundOfficeTypes) have
no port: their output is already in the JSON. The model's types live in
host_model.py and are re-exported here under upstream's module name.
"""

from __future__ import annotations

from collections.abc import Mapping

from .host_model import (
    DispatchOnlyLibrary,
    HostConstant,
    HostEnum,
    HostMember,
    HostMemberKind,
    HostObjectModel,
    HostType,
    get_excel_object_model,
)

__all__ = [
    "DispatchOnlyLibrary",
    "HostConstant",
    "HostEnum",
    "HostMember",
    "HostMemberKind",
    "HostObjectModel",
    "HostType",
    "get_excel_object_model",
    "merge_host_constants",
]


def merge_host_constants(*sets: Mapping[str, HostConstant]) -> dict[str, HostConstant]:
    """Merges constant tables, later sets winning name collisions
    (case-insensitive), keyed by each winner's own name."""
    out: dict[str, HostConstant] = {}
    keys_by_lower_name: dict[str, str] = {}
    for one in sets:
        for constant in one.values():
            lower_name = constant["name"].lower()
            previous_key = keys_by_lower_name.get(lower_name)
            if previous_key:
                out.pop(previous_key, None)
            keys_by_lower_name[lower_name] = constant["name"]
            out[constant["name"]] = constant
    return out
