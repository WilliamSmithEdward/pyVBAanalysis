"""Built-in VBA runtime function/statement metadata (MS-VBAL Phase 9).

Ported from xlide_vscode/src/analyzer/runtime/vbaRuntime.ts: the verified intrinsic
functions and statements available in every VBA project (MsgBox, Left, CLng, Now,
Array, RGB, ...), the runtime constants (vbCrLf, vbObjectError, ...) and the
global objects (Err, Debug). All three tables are mechanically extracted to
data/vba_runtime_tables.json by tools/extract_runtime_tables.mjs, never
transcribed. They feed the undeclared-reference negative lookup: a name that
resolves to a runtime function, constant or object is never flagged undeclared.
"""

from __future__ import annotations

import json
from collections.abc import Mapping
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any, TypedDict

_DATA_DIR = Path(__file__).resolve().parent.parent / "data"


@dataclass(frozen=True, slots=True)
class VbaRuntimeParam:
    name: str
    type_: str | None = None
    optional: bool = False
    param_array: bool = False


@dataclass(frozen=True, slots=True)
class VbaRuntimeFunction:
    name: str
    signature: str
    kind: str  # 'function' | 'statement'
    returns: str | None = None
    params: tuple[VbaRuntimeParam, ...] | None = None
    explicit_call: str | None = None  # 'allowed' | 'forbidden'


@lru_cache(maxsize=1)
def _raw_runtime_tables() -> dict[str, Any]:
    raw: dict[str, Any] = json.loads((_DATA_DIR / "vba_runtime_tables.json").read_text(encoding="utf-8"))
    return raw


def _function_from(raw: Mapping[str, Any]) -> VbaRuntimeFunction:
    params = raw.get("params")
    return VbaRuntimeFunction(
        name=raw["name"],
        signature=raw["signature"],
        kind=raw["kind"],
        returns=raw.get("returns"),
        params=(
            None
            if params is None
            else tuple(
                VbaRuntimeParam(p["name"], p.get("type"), bool(p.get("optional")), bool(p.get("paramArray")))
                for p in params
            )
        ),
        explicit_call=raw.get("explicitCall"),
    )


# Extracted with the constant and object tables rather than transcribed by hand,
# so it cannot fall behind upstream's: the hand-kept copy lacked `Line` and
# thirteen others, and `Line Input #f, s` read as an undeclared variable.
VBA_RUNTIME_FUNCTIONS: tuple[VbaRuntimeFunction, ...] = tuple(
    _function_from(raw) for raw in _raw_runtime_tables()["functions"]
)


_BY_LOWER = {f.name.lower(): f for f in VBA_RUNTIME_FUNCTIONS}


def resolve_runtime_function(name: str) -> VbaRuntimeFunction | None:
    """Resolve a built-in VBA runtime function/statement by name (case-insensitive)."""
    return _BY_LOWER.get(name.lower())


def runtime_allows_explicit_call(fn: VbaRuntimeFunction) -> bool:
    """Whether this runtime entry may be the target of an explicit `Call` statement."""
    return fn.explicit_call != "forbidden"


# -- runtime constant + global-object tables -------------------------------


class VbaRuntimeConstant(TypedDict, total=False):
    name: str
    type: str
    value: str | int
    source: str
    # The VBA module of constants that holds it, as in `ColorConstants.vbRed`.
    module: str


class VbaRuntimeObjectMember(TypedDict, total=False):
    name: str
    kind: str
    signature: str
    returns: str
    writable: bool
    writeType: str


class VbaRuntimeObject(TypedDict, total=False):
    name: str
    type: str
    source: str
    exhaustive: bool
    members: list[VbaRuntimeObjectMember]


@lru_cache(maxsize=1)
def _runtime_tables() -> tuple[
    dict[str, VbaRuntimeConstant], dict[str, VbaRuntimeObject], dict[str, VbaRuntimeObject]
]:
    raw = _raw_runtime_tables()
    constants_by_lower = {c["name"].lower(): c for c in raw["constants"]}
    objects_by_lower = {o["name"].lower(): o for o in raw["objects"]}
    objects_by_type_lower = {o["type"].lower(): o for o in raw["objects"]}
    return constants_by_lower, objects_by_lower, objects_by_type_lower


def resolve_runtime_constant(name: str) -> VbaRuntimeConstant | None:
    """A built-in VBA runtime constant (vbCrLf, vbObjectError, ...) by name."""
    return _runtime_tables()[0].get(name.lower())


def resolve_runtime_object(name: str) -> VbaRuntimeObject | None:
    """A built-in VBA runtime global object (Err, Debug) by name."""
    return _runtime_tables()[1].get(name.lower())


def resolve_runtime_object_type(type_name: str) -> VbaRuntimeObject | None:
    """A built-in VBA runtime object by its qualified type (VBA.ErrObject)."""
    return _runtime_tables()[2].get(type_name.lower())


# -- names the VBA library uses as qualifiers ------------------------------

# The enumerations of the VBA library, as its type library names them.
_VBA_LIBRARY_ENUMS = (
    "VbVarType", "VbMsgBoxStyle", "VbMsgBoxResult", "VbFileAttribute", "VbStrConv", "VbDayOfWeek",
    "VbFirstWeekOfYear", "VbIMEStatus", "VbAppWinStyle", "VbCompareMethod", "VbCalendar",
    "VbDateTimeFormat", "VbTriState", "VbCallType", "VbQueryClose", "FormShowConstants",
)

# The VBA library's modules of constants, and those of functions.
_VBA_LIBRARY_CONSTANT_MODULES = ("Constants", "KeyCodeConstants", "ColorConstants", "SystemColorConstants")
_VBA_LIBRARY_FUNCTION_MODULES = (
    "Strings", "Conversion", "FileSystem", "DateTime", "Information", "Interaction", "Math", "Financial",
)


@dataclass(frozen=True, slots=True)
class VbaLibraryQualifier:
    """What a VBA library name qualifies: an enum's or a module's constants, or a
    module of functions."""

    name: str
    # The constants it holds; None for a module of functions.
    constants: tuple[VbaRuntimeConstant, ...] | None = None


@lru_cache(maxsize=1)
def _library_qualifiers_by_lower() -> dict[str, VbaLibraryQualifier]:
    constants = list(_runtime_tables()[0].values())
    out: dict[str, VbaLibraryQualifier] = {}
    for name in _VBA_LIBRARY_ENUMS:
        out[name.lower()] = VbaLibraryQualifier(name, tuple(c for c in constants if c.get("type") == name))
    for name in _VBA_LIBRARY_CONSTANT_MODULES:
        out[name.lower()] = VbaLibraryQualifier(name, tuple(c for c in constants if c.get("module") == name))
    for name in _VBA_LIBRARY_FUNCTION_MODULES:
        out[name.lower()] = VbaLibraryQualifier(name)
    return out


def resolve_vba_library_qualifier(name: str) -> VbaLibraryQualifier | None:
    """A name the VBA library defines that qualifies a member: an enum
    (`VbMsgBoxResult.vbYes`), a module of constants (`ColorConstants.vbRed`), or a
    module of functions (`Strings.Left`).

    VBA reads all of these, and is first in every project's references, so it has
    the name when a host library has one too, Excel's Constants enum among them.
    """
    return _library_qualifiers_by_lower().get(name.lower())
