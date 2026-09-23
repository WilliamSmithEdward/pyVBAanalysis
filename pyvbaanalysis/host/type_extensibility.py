"""Which host types VBA resolves a member against while compiling.

Ported from xlide_vscode/src/analyzer/host/typeExtensibility.ts.

A COM interface marked NONEXTENSIBLE can gain no members at run time, so a name
that is not on it can never resolve and VBA rejects it as a compile error.
Without that flag the object is extensible: VBA compiles the call and asks
IDispatch for the name when it runs. Excel leans on that heavily. `Application.Match`
is on no interface in the library at all, it is a worksheet function Excel resolves
dynamically, and it is ordinary VBA.

So a member list being complete is not enough to report an absent member: absence
is only provable where the interface is closed. Both facts are needed, and they
come from different places, the member list from the reference dump and this flag
from the type library's own TYPEFLAGS.

Upstream measured the flag from the registered libraries with LoadRegTypeLib
(Office 16) and confirmed it against the VBE for seven receivers. Of Excel's 747
interfaces only 27 are closed; the names are vendored as data/excel_closed_types.json
rather than transcribed here, because the no-false-positive contract for
member-not-found rests on the set being exact.
"""

from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path

_DATA_DIR = Path(__file__).resolve().parent.parent / "data"


@lru_cache(maxsize=1)
def excel_closed_type_names() -> frozenset[str]:
    """Excel types whose interface is NONEXTENSIBLE, so the VBE refuses a member
    that is not on it.

    `Worksheet` and `Chart` are the two a user meets: a typo on a worksheet
    variable is a compile error, the same typo on a Range is not.
    """
    raw = json.loads((_DATA_DIR / "excel_closed_types.json").read_text(encoding="utf-8"))
    return frozenset(raw["excelClosedTypes"])


# Model types that stand where the library returns a closed type. The Worksheets
# property of Application, Global and Workbook returns `Sheets` in the type
# library, and Sheets is closed, so the VBE refuses `Worksheets.Whatever` (XLIDE
# #79). The model returns `Worksheets` there instead, which keeps `Worksheets(1)` a
# Worksheet where Sheets would give a Worksheet or a Chart. The library's
# Worksheets interface is open but lists the same 29 members as Sheets, so the
# model's Worksheets is closed exactly as far as Sheets is. Upstream keeps this map
# private, so it is transcribed here; the closed set itself stays extracted.
_EXCEL_TYPES_STANDING_FOR: dict[str, str] = {"Worksheets": "Sheets"}


def host_type_resolves_when_compiling(qualified_name: str) -> bool:
    """Whether VBA resolves a member against this host type while compiling, so a
    name the model does not carry is genuinely absent rather than deferred.

    Takes the model's own key, qualified (`Excel.Range`) or bare, since the default
    host is Excel. A type the set does not name is extensible, which is the safe
    answer for one nobody has measured: nothing is reported for it.

    Only Excel has been measured against its type library. Another host's types keep
    the answer they had, which is what they have always been analyzed under, and
    Word and PowerPoint are closed almost throughout, so the flag would change
    little there anyway.
    """
    library, dot, display_name = qualified_name.partition(".")
    if not dot:
        library, display_name = "excel", qualified_name
    if library.lower() != "excel":
        return True
    return _EXCEL_TYPES_STANDING_FOR.get(display_name, display_name) in excel_closed_type_names()
