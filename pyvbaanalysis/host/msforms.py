"""Microsoft Forms members: what a UserForm and the controls on it carry.

Ported from msFormsControlMembers and resolveMsFormsTypeName in
xlide_vscode/src/analyzer/completion/memberAccess.ts, over data/msforms_members.json,
which tools/extract_host_model.mjs extracts from the generated
msformsReferenceMembers.ts and from userFormExtenderMembers.ts (the members VBA
wraps around a form, such as Show, Hide and Name, verified upstream on a live form).

A form whose control list is authoritative proves a member absent (XLIDE issue #26),
so the UserForm surface here must be exact: a real member missing from it would be a
false diagnostic on working code.
"""

from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path
from typing import Any

from .host_model import HostMember

_DATA = Path(__file__).resolve().parent.parent / "data" / "msforms_members.json"

# The forms type whose surface VBA extends.
VBA_USERFORM_TYPE = "MSForms.UserForm"

_QUALIFIED_TYPE_RE = re.compile(r"^MSForms\.([A-Za-z][A-Za-z0-9_]*)$")
_DECLARED_TYPE_RE = re.compile(r"^MSForms\s*\.\s*([A-Za-z][A-Za-z0-9_]*)$", re.IGNORECASE)


@lru_cache(maxsize=1)
def _msforms_data() -> dict[str, Any]:
    data: dict[str, Any] = json.loads(_DATA.read_text(encoding="utf-8"))
    return data


@lru_cache(maxsize=1)
def _control_class_names() -> frozenset[str]:
    return frozenset(_msforms_data()["controlClassNames"])


def msforms_control_members(type_name: str) -> list[HostMember] | None:
    """Members of `MSForms.ComboBox` and friends, for a form's controls, and of
    `MSForms.UserForm` for the form itself, where VBA's own additions join the type
    library's list. None for a type the forms metadata does not carry.

    A placed control also carries the `Control` base surface (Left, Top, Visible,
    Name, SetFocus, ...), which the library declares once on `MSForms.Control` rather
    than per type, so it is merged here; the per-type list wins a shared name."""
    data = _msforms_data()
    reference_members: dict[str, list[dict[str, Any]]] = data["referenceMembers"]
    match = _QUALIFIED_TYPE_RE.match(type_name)
    reference = reference_members.get(match.group(1)) if match is not None else None
    members = reference
    if type_name == VBA_USERFORM_TYPE:
        members = [*data["userFormExtenderMembers"], *(reference or [])]
    elif match is not None and match.group(1) in _control_class_names():
        own = {member["name"].lower() for member in reference or []}
        members = [
            *(reference or []),
            *(m for m in reference_members.get("Control", []) if m["name"].lower() not in own),
        ]
    if not members:
        return None
    out: list[HostMember] = []
    for member in members:
        host_member = HostMember(name=member["name"], kind=member["kind"])
        if member.get("returns") is not None:
            host_member["returns"] = member["returns"]
        if member.get("signature") is not None:
            host_member["signature"] = member["signature"]
        out.append(host_member)
    return out


def resolve_msforms_type_name(declared_type: str) -> str | None:
    """Canonical `MSForms.<Type>` for a declared type the forms metadata knows,
    case-insensitively: `Dim t As MSForms.TextBox` and a control member typed
    `MSForms.ComboBox` both chain through it. Qualified names only: a bare `TextBox`
    stays unresolved, since without the reference list it cannot be known that
    MSForms is what it means."""
    match = _DECLARED_TYPE_RE.match(declared_type.strip())
    if match is None:
        return None
    lower = match.group(1).lower()
    for key in _msforms_data()["referenceMembers"]:
        if key.lower() == lower:
            return f"MSForms.{key}"
    return None
