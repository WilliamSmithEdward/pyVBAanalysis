"""Controls declared by a UserForm's designer.

Ported from xlide_vscode/src/vbaUserFormControls.ts.

A form's controls are members of the form's class, declared by the designer rather
than by any line of code, so code-behind that says `RegionPick.AddItem` is correct
VBA. The declarations are within reach when a `.frm` header carries the control
tree in its `Begin ... End` blocks; this reads them out so the names (and their
types) can be handed to the analyzer as implicit members of the form.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from ..lexer.token_helpers import identifiers_in


@dataclass(frozen=True, slots=True)
class UserFormControl:
    """A control the designer declared on a form."""

    # Name the code-behind uses, e.g. `RegionPick`.
    name: str
    # Programmatic id exactly as the designer wrote it, e.g. `Forms.ComboBox.1`.
    prog_id: str
    # Type a member lookup should resolve against, e.g. `MSForms.ComboBox`.
    type: str


# `Begin <progId> <name>` opens a control block; the outermost one is the form
# itself and is not a control of itself. The class id is taken as an opaque token:
# a form's own is a `{...}` GUID and a control's is a prog id. The name is read
# separately, because it may use any locale's letters.
_BEGIN_RE = re.compile(r"^\s*Begin\s+(\{[^}]*\}|[A-Za-z0-9_.]+)\s+(.*)$")
_END_RE = re.compile(r"^\s*End\s*$", re.IGNORECASE)
_VERSION_RE = re.compile(r"^\s*VERSION\b", re.IGNORECASE)
_PROPERTY_LINE_RE = re.compile(r"^\s*[A-Za-z0-9_.()]+\s*=")
_PROPERTY_GROUP_RE = re.compile(r"^\s*(?:BeginProperty|EndProperty)\b", re.IGNORECASE)
_OLE_OBJECT_BLOB_RE = re.compile(r"^\s*OleObjectBlob\s*=", re.IGNORECASE)
_FORMS_PROG_ID_RE = re.compile(r"Forms\.([A-Za-z][A-Za-z0-9_]*)(?:\.[0-9]+)?")


def _begin(line: str) -> tuple[str, str] | None:
    """The prog id and control name of a `Begin` line, or None."""
    match = _BEGIN_RE.match(line)
    if match is None:
        return None
    rest = match.group(2)
    words = identifiers_in(rest)
    if not words or not rest.startswith(words[0]):
        return None
    return match.group(1), words[0]


def _version_line_index(lines: list[str]) -> int | None:
    """Index of the `VERSION` line a designer header opens with, past blank lines."""
    index = 0
    while index < len(lines) and lines[index].strip() == "":
        index += 1
    return index if index < len(lines) and _VERSION_RE.match(lines[index]) else None


def _is_designer_header_line(line: str) -> bool:
    """A blank, a property assignment or a property group: all a header holds
    besides its blocks."""
    return (
        line.strip() == ""
        or _PROPERTY_LINE_RE.match(line) is not None
        or _PROPERTY_GROUP_RE.match(line) is not None
    )


def has_authoritative_designer_header(source: str) -> bool:
    """True when the source carries a `.frm` designer header whose control list
    this parser can actually see, telling "this form declares no controls" from
    "nobody has read this form's designer".

    Two gates: the VERSION header must open the source, and the form block must NOT
    defer to an `OleObjectBlob` line. A real VBA export stores its controls in the
    binary `.frx` behind exactly that line, so such a header proves nothing about
    the control list; nested `Begin` control blocks appear only in sources that
    spell the controls out, and those are authoritative even when the list is
    empty.
    """
    lines = re.split(r"\r?\n", source)
    index = _version_line_index(lines)
    if index is None:
        return False
    depth = 0
    for line in lines[index + 1 :]:
        if _OLE_OBJECT_BLOB_RE.match(line):
            return False
        if _begin(line) is not None:
            depth += 1
            continue
        if _END_RE.match(line):
            depth -= 1
            if depth <= 0:
                return True
            continue
        if not _is_designer_header_line(line):
            return False
    return False


def parse_user_form_controls(source: str) -> list[UserFormControl]:
    """The controls a `.frm` designer header declares. Empty for source that is not
    a form header, so callers can pass any module."""
    lines = re.split(r"\r?\n", source)
    index = _version_line_index(lines)
    if index is None:
        return []
    out: list[UserFormControl] = []
    depth = 0
    for line in lines[index + 1 :]:
        begin = _begin(line)
        if begin is not None:
            depth += 1
            # Depth 1 is the form; only what it contains is a control.
            if depth > 1:
                prog_id, name = begin
                out.append(UserFormControl(name, prog_id, _control_type_for(prog_id)))
            continue
        if _END_RE.match(line):
            depth -= 1
            if depth <= 0:
                break
            continue
        # Property lines, property groups (a Font), and blanks are the only other
        # things a header holds; a line that is none of them means the block never
        # closed, so stop rather than run on into the code and invent controls out
        # of it.
        if not _is_designer_header_line(line):
            break
    return out


def _control_type_for(prog_id: str) -> str:
    """`Forms.ComboBox.1` is how the designer writes it; `MSForms.ComboBox` is what
    the type library calls it. An id in a shape not recognised here is passed
    through, since a wrong guess is worse than an unresolved type."""
    match = _FORMS_PROG_ID_RE.fullmatch(prog_id)
    return f"MSForms.{match.group(1)}" if match else prog_id
