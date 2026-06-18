"""VBA type-declaration suffix characters (MS-VBAL 5.6.1.1).

Ported from xlide_vscode/src/analyzer/parser/typeDeclarationSuffix.ts. A trailing
$ % & ! # @ ^ on a name fixes its declared type.
"""

from __future__ import annotations

_TYPE_DECLARATION_SUFFIX_TYPES: dict[str, str] = {
    "$": "String",
    "%": "Integer",
    "&": "Long",
    "!": "Single",
    "#": "Double",
    "@": "Currency",
    "^": "LongLong",
}


def is_type_declaration_suffix(value: str | None) -> bool:
    """True when value is one of the seven type-declaration suffix characters."""
    return value in _TYPE_DECLARATION_SUFFIX_TYPES


def type_name_for_declaration_suffix(suffix: str | None) -> str | None:
    """The fundamental type name for a suffix, or None when it is not a suffix."""
    if suffix is None:
        return None
    return _TYPE_DECLARATION_SUFFIX_TYPES.get(suffix)
