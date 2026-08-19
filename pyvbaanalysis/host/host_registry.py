"""Which Office host a module's VBA belongs to, and the object model that
answers for it.

Ported from xlide_vscode/src/analyzer/host/hostRegistry.ts. The resolvers have
always accepted any HostObjectModel and defaulted to Excel's; this is the seam
that lets a caller choose (XLIDE issue #24).

The semantics are deliberately asymmetric:

- ABSENT means Excel. Every existing caller keeps exactly the behavior it had.
- A NAMED host with no model means NO HOST KNOWLEDGE, an empty model rather
  than Excel's. Telling Word's ThisDocument it has Cells and Range was the bug
  that motivated the seam; silence is the honest answer until the host's own
  model exists.
"""

from __future__ import annotations

from types import MappingProxyType

from .host_model import (
    HostObjectModel,
    get_access_object_model,
    get_excel_object_model,
    get_powerpoint_object_model,
    get_word_object_model,
)

# The host tokens xlide_vbide normalizes from the process image, so an embedder
# passes the string it already has.
VBA_HOST_TOKENS = frozenset(
    {"excel", "word", "powerpoint", "access", "outlook", "visio", "project", "other"}
)

# A model that knows nothing: every lookup misses, so nothing is asserted. The
# maps are proxies because this singleton is handed to every caller that names an
# unmodelled host; a mutation would poison all of them.
EMPTY_HOST_MODEL: HostObjectModel = {
    "source": "none: a host whose object model is not yet available",
    "types": MappingProxyType({}),
    "aliases": MappingProxyType({}),
    "globals": MappingProxyType({}),
    "constants": MappingProxyType({}),
    "memberSignatures": MappingProxyType({}),
}

_MODELS_BY_TOKEN = {
    "excel": get_excel_object_model,
    "word": get_word_object_model,
    "powerpoint": get_powerpoint_object_model,
    "access": get_access_object_model,
}

# Container extension (without the dot) -> host token. XLIDE's own file
# surfaces are the caller, so the analyzer knows what kind of file a module
# came from without anyone having to say.
_HOST_BY_EXTENSION = {
    "xlsm": "excel",
    "xlsb": "excel",
    "xlam": "excel",
    "xltm": "excel",
    "xls": "excel",
    "xlt": "excel",
    "xla": "excel",
    "docm": "word",
    "dotm": "word",
    "doc": "word",
    "dot": "word",
    "pptm": "powerpoint",
    "potm": "powerpoint",
    "ppsm": "powerpoint",
    "ppam": "powerpoint",
    "ppt": "powerpoint",
    "ppa": "powerpoint",
    "accdb": "access",
    "accda": "access",
    "mdb": "access",
    "mda": "access",
}


def host_object_model_for_token(host: str | None) -> HostObjectModel | None:
    """The model a host token selects.

    Absent (or an unrecognised casing of ``excel``) answers None so downstream
    Excel defaults keep today's behavior; any other named host answers its own
    model, or the empty model when none is registered.
    """
    if host is None:
        return None
    token = host.strip().lower()
    if token == "" or token == "excel":
        return None
    loader = _MODELS_BY_TOKEN.get(token)
    return loader() if loader is not None else EMPTY_HOST_MODEL


def host_knowledge_is_absent(model: HostObjectModel | None) -> bool:
    """True when a model asserts nothing about its host.

    None is not absent knowledge: it means Excel, which is fully modelled. This
    is the empty model a NAMED but unmodelled host resolves to, and the rules
    that decide whether a bare name is legal cannot answer under it.
    """
    if model is None:
        return False
    return not model.get("types") and not model.get("globals")


def host_token_for_file_name(file_name: str) -> str | None:
    """The host a macro container implies, from its file name, or None."""
    _, dot, suffix = str(file_name).rpartition(".")
    if not dot:
        return None
    return _HOST_BY_EXTENSION.get(suffix.lower())
