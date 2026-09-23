"""The Office type libraries a VBA project can reference, and the host model that
answers for each.

Ported from xlide_vscode/src/analyzer/host/hostLibraries.ts.

A project that references another application's library can name its types and
call its members: a Word document with a reference to the Excel library compiles
`Dim xl As Excel.Application`, and `xl.Calculate` is checked against Excel's object
model rather than Word's. Reading only the host a file's extension implies answers
nothing for that code.

A reference is stored as a libid, `*\\G{guid}#major.minor#lcid#path#name`. The GUID
is the identity: the path is a hint the host resolves through the registry, which
is why a file written on one machine loads on another whose libraries sit elsewhere.
So the GUID is what this maps, never the path or the description. Each GUID was
read upstream from the registered type library on a machine with Office 16,
together with the library's own name, which is the qualifier VBA writes.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Sequence
from types import MappingProxyType

# The library GUID a reference declares -> the host whose model answers.
_HOST_BY_LIBRARY_GUID: MappingProxyType[str, str] = MappingProxyType(
    {
        "{00020813-0000-0000-C000-000000000046}": "excel",
        "{00020905-0000-0000-C000-000000000046}": "word",
        "{91493440-5A91-11CF-8700-00AA0060263B}": "powerpoint",
        "{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}": "access",
    }
)

# The name each library gives itself, which is the qualifier VBA writes.
HOST_LIBRARY_NAMES: MappingProxyType[str, str] = MappingProxyType(
    {
        "excel": "Excel",
        "word": "Word",
        "powerpoint": "PowerPoint",
        "access": "Access",
        "outlook": "Outlook",
        "visio": "Visio",
        "project": "MSProject",
        "vb6": "VB",
        "other": "",
    }
)

_GUID_RE = re.compile(r"\{[0-9a-fA-F-]{36}\}")


def library_guid_of(libid: str) -> str | None:
    """The GUID inside a libid, upper-cased with its braces, or None."""
    found = _GUID_RE.search(libid)
    return found.group(0).upper() if found is not None else None


def host_token_for_libid(libid: str) -> str | None:
    """The host whose object model answers for a reference, or None for one this
    analyzer has no model for: stdole, the shared Office library, a third-party
    DLL. None means no knowledge, and silence is the honest answer for it; the
    same rule the registry applies to an unmodelled host."""
    guid = library_guid_of(libid)
    return _HOST_BY_LIBRARY_GUID.get(guid) if guid is not None else None


def host_tokens_for_project(host: str | None, libids: Iterable[str]) -> list[str]:
    """The hosts a project's references bring in, in the order VBA resolves them:
    the project's own host first, then each referenced library in the order the
    project declares it.

    VBA resolves an ambiguous name by that order, so a Word document referencing
    Excel keeps Word's Range for `Dim r As Range`. The host itself is included even
    when its library is not in the reference list, because a project's own host
    library is implicit.
    """
    out: list[str] = []
    for token in (host, *(host_token_for_libid(libid) for libid in libids)):
        if token is None or token in out:
            continue
        out.append(token)
    return out


def referenced_host_tokens(host: str | None, libids: Sequence[str]) -> list[str]:
    """The same list without the project's own host: the libraries its code can
    name because the project references them. The analyzer takes the host and these
    separately, because the host also decides what `Me` is and which host-specific
    rules apply, which a referenced library never does."""
    return [token for token in host_tokens_for_project(host, libids) if token != host]
