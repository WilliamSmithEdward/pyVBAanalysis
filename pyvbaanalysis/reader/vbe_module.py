"""Shared VBE export-format handling for the file and workbook readers.

A VBA module exported by the VBE, or read out of an Office file, is the code body
preceded by an export header. The header is:

* optional ``Attribute VB_*`` lines (valid module statements the parser tolerates), and
* for class, document, and UserForm modules, a non-VBA designer block: a leading
  ``VERSION ...`` line followed by a balanced ``Begin ... End`` block.

The parser handles Attribute lines but not the designer block, so this module
strips the designer block, derives the module name from the VB_Name attribute, and
classifies the module kind. Both reader.loose_file and reader.workbook build on it,
so the export-format knowledge lives in exactly one place.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from ..symbols import ModuleSymbolKind

_VB_NAME_RE = re.compile(
    r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"', re.IGNORECASE | re.MULTILINE
)
_VB_BASE_RE = re.compile(r"^\s*Attribute\s+VB_Base\s*=", re.IGNORECASE | re.MULTILINE)
# A UserForm designer block opens with ``Begin {GUID} Name`` (a 38-char braced GUID).
_DESIGNER_BEGIN_RE = re.compile(r"^\s*Begin\s*\{[0-9A-Fa-f-]{36}\}", re.MULTILINE)
_VERSION_FORM_RE = re.compile(r"^\s*VERSION\s+5\.", re.IGNORECASE)
_VERSION_CLASS_RE = re.compile(r"^\s*VERSION\s+\d+\.\d+\s+CLASS", re.IGNORECASE)


@dataclass(frozen=True, slots=True)
class LoadedModule:
    """One VBA module ready for analysis: its name, kind, and code-body source.

    ``source`` is the body with any non-VBA designer header removed, so it can be
    passed straight to analyze_module (or wrapped in a ModuleInput for a project pass).
    """

    name: str
    kind: ModuleSymbolKind
    source: str


def strip_export_header(text: str) -> str:
    """Remove the leading non-VBA designer block (a ``VERSION`` line plus a balanced
    ``Begin ... End`` block).

    Attribute lines are valid module statements the parser handles, so they are kept.
    A plain module with no ``VERSION`` line is returned unchanged. A truncated or
    corrupt header whose ``Begin`` block never closes is also returned unchanged
    rather than swallowing the whole body (a corrupt export stays visible).
    """
    lines = text.splitlines(keepends=True)
    if not lines or not lines[0].lstrip().upper().startswith("VERSION"):
        return text
    # Skip blank lines after VERSION, then consume a balanced Begin..End designer block.
    j = 1
    while j < len(lines) and lines[j].strip() == "":
        j += 1
    if not (j < len(lines) and lines[j].strip().lower().startswith("begin")):
        # A VERSION line with no designer block: strip just the VERSION line.
        return "".join(lines[1:])
    depth = 0
    for k in range(j, len(lines)):
        token = lines[k].strip().lower()
        # Designer blocks open with `Begin <...>` and close with a bare `End`. VBA
        # code uses `End Sub` / `End Function` / etc., which must not count as closers
        # (it only matters once a malformed block scans into the body).
        if token.startswith("begin"):
            depth += 1
        elif token == "end":
            depth -= 1
            if depth <= 0:
                return "".join(lines[k + 1 :])
    # The Begin block never balanced (truncated/corrupt export): do not strip the body.
    return text


def classify_module_kind(
    text: str,
    *,
    extension: str | None = None,
    pyopenvba_standard: bool | None = None,
) -> ModuleSymbolKind:
    """Classify a module as standard, class, document, or UserForm.

    Signals, strongest first: a UserForm designer block or ``.frm`` extension; an
    ``Attribute VB_Base`` line (the host-bound document modules, e.g. ThisWorkbook);
    then the file extension or the pyOpenVBA standard/other flag; falling back to the
    ``VERSION ... CLASS`` header. ``pyopenvba_standard`` is True for a pyOpenVBA
    standard module, False for its "other" bucket (class/document/designer), or None
    when no workbook reader supplied it.
    """
    ext = (extension or "").lower().lstrip(".")
    head = text[:8000]
    if ext == "frm" or _VERSION_FORM_RE.match(head) or _DESIGNER_BEGIN_RE.search(head):
        return ModuleSymbolKind.USERFORM
    if _VB_BASE_RE.search(head):
        return ModuleSymbolKind.DOCUMENT
    if ext == "bas":
        return ModuleSymbolKind.STANDARD
    if ext == "cls":
        return ModuleSymbolKind.CLASS
    if pyopenvba_standard is True:
        return ModuleSymbolKind.STANDARD
    if pyopenvba_standard is False:
        return ModuleSymbolKind.CLASS
    return ModuleSymbolKind.CLASS if _VERSION_CLASS_RE.match(head) else ModuleSymbolKind.STANDARD


def module_name_from_text(text: str, fallback: str) -> str:
    """The VB_Name attribute value, or ``fallback`` when the header has none."""
    match = _VB_NAME_RE.search(text)
    return match.group(1) if match is not None and match.group(1) else fallback


def loaded_module_from_text(
    text: str,
    *,
    name: str | None = None,
    extension: str | None = None,
    name_fallback: str | None = None,
    pyopenvba_standard: bool | None = None,
) -> LoadedModule:
    """Turn raw exported module text into a LoadedModule (name, kind, code body).

    Pass ``name`` when the source already knows it (a workbook reader has the VBE
    component name); otherwise the VB_Name attribute is used, then ``name_fallback``
    (typically a file stem), then ``"Module"``.
    """
    kind = classify_module_kind(
        text, extension=extension, pyopenvba_standard=pyopenvba_standard
    )
    resolved_name = name or module_name_from_text(text, name_fallback or "Module")
    return LoadedModule(name=resolved_name, kind=kind, source=strip_export_header(text))
