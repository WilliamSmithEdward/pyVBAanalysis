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

from collections.abc import Mapping
from types import MappingProxyType
from typing import Any, cast

from .host_model import (
    HostConstant,
    HostObjectModel,
    HostType,
    get_access_object_model,
    get_excel_object_model,
    get_powerpoint_object_model,
    get_vb6_object_model,
    get_word_object_model,
)

# The host tokens xlide_vbide normalizes from the process image, so an embedder
# passes the string it already has, plus `vb6`: a VB6 project is not an Office host
# at all, but its code-behind needs the VB runtime's surface rather than Excel's,
# and the analyzer selects a model by this token.
VBA_HOST_TOKENS = frozenset(
    {"excel", "word", "powerpoint", "access", "outlook", "visio", "project", "vb6", "other"}
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
    "vb6": get_vb6_object_model,
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


# Merged models, keyed by the token list that produced them. Bounded by the handful
# of host combinations a project can name, and each model is a shared singleton.
_MERGED_BY_KEY: dict[str, HostObjectModel] = {}


def host_object_model_for_tokens(tokens: list[str]) -> HostObjectModel | None:
    """One model answering for a project's own host and every library it
    references, in the order VBA resolves them.

    A project that references another application's library can name its types and
    call its members, so a Word document with a reference to Excel has to be
    analyzed against both. The FIRST token wins every shared name, which is how VBA
    resolves an ambiguous one: by the reference list's order, the project's own
    host at the top.

    Every library's globals are merged, because a library marks its global object
    APPOBJECT in its type library and VBA binds that object's members bare for
    anyone who references it. The host's own still wins a collision.

    Returns None when the list adds nothing to what a single token would have
    given, so every existing caller keeps the model it had, including the bare
    `excel` that rides as the downstream Excel default.
    """
    known = [token for token in tokens if token in _MODELS_BY_TOKEN]
    if len(known) <= 1:
        return host_object_model_for_token(known[0] if known else (tokens[0] if tokens else None))
    key = "+".join(known)
    cached = _MERGED_BY_KEY.get(key)
    if cached is not None:
        return cached

    models = [_MODELS_BY_TOKEN[token]() for token in known]
    # Later models are applied first so the earlier ones overwrite them: the
    # project's own host wins every name it shares with a referenced library.
    layered = list(reversed(models))

    def merged(field: str) -> Mapping[str, Any]:
        out: dict[str, Any] = {}
        for one in layered:
            out.update(cast("Mapping[str, Any]", one.get(field) or {}))
        return MappingProxyType(out)

    # A type key already names its library; an enum key does not, so each
    # referenced library's enums carry theirs.
    enums: dict[str, Mapping[str, object]] = {}
    for one in layered:
        for name, entry in (one.get("enums") or {}).items():
            enums[name] = entry if one is models[0] else {**entry, "library": one.get("hostName")}

    result: HostObjectModel = {
        "source": " + ".join(one["source"] for one in models),
        "hostName": models[0].get("hostName", ""),
        "globalType": models[0].get("globalType"),
        # Each field merges the same-typed field of every layer, so the value
        # types carry over from the source models unchanged.
        "types": cast("Mapping[str, HostType]", merged("types")),
        "aliases": cast("Mapping[str, str]", merged("aliases")),
        "globals": cast("Mapping[str, str]", merged("globals")),
        "constants": cast("Mapping[str, HostConstant]", merged("constants")),
        "enums": MappingProxyType(enums),
        "memberSignatures": cast("Mapping[str, Mapping[str, str]]", merged("memberSignatures")),
    }
    _MERGED_BY_KEY[key] = result
    return result


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
