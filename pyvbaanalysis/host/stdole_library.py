"""stdole (OLE Automation, stdole2.tlb), which the default references of a new
Excel, Word, PowerPoint and Access file all hold.

Ported from xlide_vscode/src/analyzer/host/stdoleLibrary.ts (XLIDE #1665). What
it offers without a qualifier: the StdFunctions module's LoadPicture and
SavePicture, and the members of its OLE_TRISTATE and LoadPictureConstants enums.

Measured in Excel, Word, PowerPoint and Access 16.0 (build 20527) through the
VBE oracle: each name compiles bare under Option Explicit. The signatures are the
VBE's Quick Info for the two functions. In Access, Access's own
Application.LoadPicture(FileName) wins over stdole's.

A project without stdole could not be measured, since the oracle only adds
references, so these names are known whatever the reference list says.
"""

from __future__ import annotations

from collections.abc import Mapping

# stdole's two functions, by lowercased name, with the VBE's Quick Info signature.
STDOLE_FUNCTION_SIGNATURES: Mapping[str, str] = {
    "loadpicture": (
        "LoadPicture([filename], [widthDesired As Long], [heightDesired As Long], "
        "[flags As LoadPictureConstants = Default]) As IPictureDisp"
    ),
    "savepicture": "SavePicture(Picture As IPictureDisp, filename As String)",
}

# The members of stdole's OLE_TRISTATE and LoadPictureConstants enums, lowercased.
STDOLE_CONSTANT_NAMES: frozenset[str] = frozenset({
    "unchecked", "checked", "gray",
    "default", "monochrome", "vgacolor", "color",
})

# The hosts whose `hostName` the measurements above cover; a model with none is Excel's.
STDOLE_MEASURED_HOSTS: frozenset[str] = frozenset({"Excel", "Word", "PowerPoint", "Access"})
