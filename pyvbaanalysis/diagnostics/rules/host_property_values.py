"""Rule: host properties set to a literal the host refuses (XLIDE issue #204).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/hostPropertyValues.ts.

The type libraries give these properties a type, not a range, so the ranges
are a table. Every range below is one Excel, Word or PowerPoint 16.0 was seen
to refuse, through pyVBAharness on 2026-09-29, beside values it took: only
values inside a refused range are reported, never a value the table merely
does not know. `Font.Size = 409.5` runs in Excel and 409.6 raises; nothing
between is claimed.

The enum-typed ones (XLIDE issue #244) were swept from -100 to 999 in Excel
16.0: the refused ranges are the runs of that sweep that raised, so a constant
such as xlCenter (-4108) or xlPatternLinearGradient (4000), outside every run,
is never claimed. Calculation, CutCopyMode and ReferenceStyle took every value
of the sweep, and are not here.

Only a numeric literal is read, optionally signed. `ActiveWindow.Zoom = False`
runs where `Zoom = 0` raises, and a named constant such as xlVertical is its
own value, not a number the table can place. A String literal with no digit is
read too, against a table of its own (XLIDE issue #416).
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from typing import Protocol

from ...constants.integer_constant_expression import parse_vba_integer_literal
from ...js_compat import JS_WHITESPACE, js_number, js_number_to_string
from ...lexer.token_kinds import TokenKind, VbaToken


class _OwnedMember(Protocol):
    """What these checks read of a resolved member: upstream's MemberCompletion."""

    @property
    def name(self) -> str: ...

    @property
    def owner(self) -> str: ...


@dataclass(frozen=True, slots=True)
class _RaisedError:
    number: str
    text: str


@dataclass(frozen=True, slots=True)
class _Refused:
    """An inclusive range the host refuses; an open end is unbounded."""

    from_: float | None = None
    to: float | None = None


@dataclass(frozen=True, slots=True)
class _HostValueLimit:
    # The range that runs, as the message states it.
    runs: str
    # Inclusive ranges the host refuses; an open end is unbounded.
    refused: tuple[_Refused, ...]
    error: _RaisedError
    # Values inside a refused range that run: the host's named constants.
    allowed: tuple[float, ...] | None = None


def _excel_1004(prop: str, owner: str) -> _RaisedError:
    return _RaisedError("1004", f"Unable to set the {prop} property of the {owner} class")


_SUBSCRIPT = _RaisedError("9", "Subscript out of range")
_TYPE_MISMATCH = _RaisedError("13", "Type mismatch")
_WORD_RANGE = _RaisedError("5843", "One of the values passed to this method or property is out of range")


def _r(from_: float | None = None, to: float | None = None) -> _Refused:
    return _Refused(from_, to)


# By qualified owner type, then lower-cased property name.
_LIMITS: Mapping[str, Mapping[str, _HostValueLimit]] = {
    "Excel.Font": {
        "size": _HostValueLimit("1 to 409.5", (_r(to=0.5), _r(409.6)), _excel_1004("Size", "Font")),
        "underline": _HostValueLimit(
            "1 to 5, or an xlUnderlineStyle constant",
            (_r(-100, 0), _r(6, 999)),
            _excel_1004("Underline", "Font"),
        ),
    },
    "Excel.Interior": {
        "colorindex": _HostValueLimit("1 to 56, or an xlColorIndex constant", (_r(57),), _SUBSCRIPT),
        "pattern": _HostValueLimit(
            "-1 to 18, or an xlPattern constant", (_r(-100, -2), _r(19, 999)), _SUBSCRIPT
        ),
    },
    "Excel.Border": {
        "linestyle": _HostValueLimit(
            "0 to 13, or an xlLineStyle constant",
            (_r(-100, -1), _r(14, 999)),
            _excel_1004("LineStyle", "Border"),
        ),
        "weight": _HostValueLimit(
            "1 to 4, or an xlBorderWeight constant",
            (_r(-100, 0), _r(5, 999)),
            _excel_1004("Weight", "Border"),
        ),
    },
    "Excel.PageSetup": {
        "orientation": _HostValueLimit(
            "xlPortrait (1) or xlLandscape (2)",
            (_r(-100, 0), _r(3, 999)),
            _excel_1004("Orientation", "PageSetup"),
        ),
    },
    "Excel.Worksheet": {
        "visible": _HostValueLimit(
            "an xlSheetVisibility constant: -1, 0 or 2",
            (_r(-100, -2), _r(3, 999)),
            _excel_1004("Visible", "Worksheet"),
        ),
    },
    "Excel.Tab": {
        "colorindex": _HostValueLimit("1 to 56, or xlColorIndexNone", (_r(-1, 0), _r(57)), _SUBSCRIPT),
    },
    "Excel.Range": {
        "rowheight": _HostValueLimit("0 to 409.5", (_r(to=-0.5), _r(409.75)), _excel_1004("RowHeight", "Range")),
        "columnwidth": _HostValueLimit("0 to 255", (_r(to=-0.5), _r(255.5)), _excel_1004("ColumnWidth", "Range")),
        # xlHorizontal, xlVertical, xlUpward and xlDownward are the constants.
        "orientation": _HostValueLimit(
            "-90 to 90, or an xlOrientation constant",
            (_r(to=-91), _r(91)),
            _excel_1004("Orientation", "Range"),
            (-4128, -4166, -4171, -4170),
        ),
        "indentlevel": _HostValueLimit("up to 250", (_r(to=-16), _r(251)), _excel_1004("IndentLevel", "Range")),
        "horizontalalignment": _HostValueLimit(
            "1 to 8, or an xlHAlign constant",
            (_r(-100, 0), _r(9, 999)),
            _excel_1004("HorizontalAlignment", "Range"),
        ),
        "verticalalignment": _HostValueLimit(
            "1 to 5, or an xlVAlign constant",
            (_r(-100, 0), _r(6, 999)),
            _excel_1004("VerticalAlignment", "Range"),
        ),
    },
    "Excel.Window": {
        # -1 is True, which fits the selection.
        "zoom": _HostValueLimit("10 to 400, or True", (_r(to=9), _r(401)), _excel_1004("Zoom", "Window"), (-1,)),
        "windowstate": _HostValueLimit(
            "1 to 3, or an xlWindowState constant",
            (_r(-100, 0), _r(4, 999)),
            _excel_1004("WindowState", "Window"),
        ),
    },
    "Word.Font": {
        "size": _HostValueLimit("1 to 1638", (_r(to=0.5), _r(1638.5)), _WORD_RANGE),
        # Swept from -100 to 999 in Word 16.0 (XLIDE issue #245): the wdUnderline
        # values are scattered, and each gap between them is refused.
        "underline": _HostValueLimit(
            "a wdUnderline constant: -1 to 4, 6, 7, 9 to 11, 20, 23, 25 to 27, 39, 43 or 55",
            (
                _r(-100, -2),
                _r(5, 5),
                _r(8, 8),
                _r(12, 19),
                _r(21, 22),
                _r(24, 24),
                _r(28, 38),
                _r(40, 42),
                _r(44, 54),
                _r(56, 999),
            ),
            _WORD_RANGE,
        ),
    },
    "Word.Zoom": {
        "percentage": _HostValueLimit("10 to 500", (_r(to=9), _r(501)), _WORD_RANGE),
    },
    "Word.Paragraph": {
        "alignment": _HostValueLimit(
            "0 to 9", (_r(-100, -1), _r(10, 999)), _RaisedError("5148", "The number must be between 0 and 9")
        ),
        "linespacingrule": _HostValueLimit(
            "0 to 5", (_r(-100, -1), _r(6, 999)), _RaisedError("5148", "The number must be between 0 and 5")
        ),
        "leftindent": _HostValueLimit(
            "-1584 to 1584 points",
            (_r(to=-1585), _r(1585)),
            _RaisedError("5149", "The measurement must be between -1584 pt and 1584 pt"),
        ),
    },
    "PowerPoint.Font": {
        "size": _HostValueLimit(
            "1 to 4000",
            (_r(to=0), _r(4000.25)),
            _RaisedError("-2147024809", "The specified value is out of range"),
        ),
    },
}

_FLOAT_SUFFIX = re.compile(r"[!#@]\Z")


def signed_numeric_literal(tokens: Sequence[VbaToken]) -> int | float | None:
    """The number a value's tokens spell: a numeric literal, optionally signed."""
    toks = [tok for tok in tokens if tok.kind is not TokenKind.COMMENT]
    signed = len(toks) == 2 and (toks[0].raw_text == "-" or toks[0].raw_text == "+")
    literal = toks[1] if signed else toks[0] if len(toks) == 1 else None
    value: int | float | None = None
    if literal is not None and literal.kind is TokenKind.INTEGER_LITERAL:
        value = parse_vba_integer_literal(literal.raw_text)
    elif literal is not None and literal.kind is TokenKind.FLOAT_LITERAL:
        value = js_number(_FLOAT_SUFFIX.sub("", literal.raw_text))
    if value is None or not math.isfinite(value):
        return None
    return -value if signed and toks[0].raw_text == "-" else value


@dataclass(frozen=True, slots=True)
class _StringLimit:
    error: _RaisedError
    takes_boolean: bool


# Excel properties given a String that is no number (XLIDE issue #416, measured
# in Excel 16.0 on 2026-10-01, "abc" and "" each): the error each raises, and
# whether "True" and "False" convert for it. A String with a digit is not
# judged, since "12" runs for most of them and a locale decides the rest.
_STRING_LIMITS: Mapping[str, Mapping[str, _StringLimit]] = {
    "Excel.Font": {
        "bold": _StringLimit(_excel_1004("Bold", "Font"), True),
        "size": _StringLimit(_excel_1004("Size", "Font"), False),
    },
    # Through a typed Worksheet; ActiveSheet.Visible, late-bound, raises 1004.
    "Excel.Worksheet": {"visible": _StringLimit(_TYPE_MISMATCH, False)},
    "Excel.Range": {
        "columnwidth": _StringLimit(_excel_1004("ColumnWidth", "Range"), False),
        "rowheight": _StringLimit(_excel_1004("RowHeight", "Range"), False),
        "horizontalalignment": _StringLimit(_excel_1004("HorizontalAlignment", "Range"), False),
        "wraptext": _StringLimit(_excel_1004("WrapText", "Range"), True),
    },
    "Excel.Window": {"zoom": _StringLimit(_excel_1004("Zoom", "Window"), False)},
    "Excel.Application": {
        "screenupdating": _StringLimit(_TYPE_MISMATCH, True),
        "displayalerts": _StringLimit(_TYPE_MISMATCH, True),
        "calculation": _StringLimit(_TYPE_MISMATCH, False),
    },
    "Excel.Interior": {"color": _StringLimit(_TYPE_MISMATCH, False)},
    "Excel.Tab": {"color": _StringLimit(_TYPE_MISMATCH, False)},
}

# JavaScript's `\s` and `$`; its `/i` folds no letter outside ASCII onto these.
_BOOLEAN_TEXT = re.compile(
    "^[" + JS_WHITESPACE + "]*(true|false)[" + JS_WHITESPACE + r"]*\Z", re.IGNORECASE | re.ASCII
)
_DIGIT = re.compile("[0-9]")


def _literal_text(raw: str) -> str:
    """`rawText.slice(1, -1).replace(/""/g, '"')`."""
    return raw[1:-1].replace('""', '"')


def _host_property_string_problem(target: _OwnedMember, value_tokens: Sequence[VbaToken]) -> str | None:
    """The message for a String literal a host property refuses, or None."""
    owner_limits = _STRING_LIMITS.get(target.owner)
    limit = owner_limits.get(target.name.lower()) if owner_limits is not None else None
    toks = [tok for tok in value_tokens if tok.kind is not TokenKind.COMMENT]
    if limit is None or len(toks) != 1 or toks[0].kind is not TokenKind.STRING_LITERAL:
        return None
    text = _literal_text(toks[0].raw_text)
    boolean = _BOOLEAN_TEXT.search(text) is not None
    if _DIGIT.search(text) is not None or (boolean and limit.takes_boolean):
        return None
    bare = target.owner[target.owner.find(".") + 1 :]
    takes = (
        f"a number, True or False, and the String {toks[0].raw_text} is none of these"
        if limit.takes_boolean
        else f"a number, and the String {toks[0].raw_text} is not one"
    )
    return (
        f"{bare}.{target.name} takes {takes}. This will raise Run-time error "
        f"'{limit.error.number}': {limit.error.text}."
    )


# The Range properties that read a String starting "=" as a formula.
_FORMULA_PROPERTIES: frozenset[str] = frozenset(
    {"formula", "formular1c1", "formula2", "formula2r1c1", "formulalocal", "formular1c1local", "value", "value2"}
)

# JavaScript's `\s` and `$`.
_ENDS_WITH_OPERATOR = re.compile(r"[+\-*/^&=<>,][" + JS_WHITESPACE + r"]*\Z")


def formula_string_problem(target: _OwnedMember, value_tokens: Sequence[VbaToken]) -> str | None:
    """Why a formula String Excel cannot parse is one, or None: a parenthesis
    left open or closed twice, a string in it never closed, or an operator with
    nothing after it. `Range("A1").Formula = "=SUM(B1:B2"` raises 1004, except
    on a cell formatted as Text, which keeps it as text (XLIDE issue #276,
    measured in Excel 16.0)."""
    if target.owner != "Excel.Range" or target.name.lower() not in _FORMULA_PROPERTIES:
        return None
    toks = [tok for tok in value_tokens if tok.kind is not TokenKind.COMMENT]
    if len(toks) != 1 or toks[0].kind is not TokenKind.STRING_LITERAL:
        return None
    text = _literal_text(toks[0].raw_text)
    if not text.startswith("=") or len(text) < 2:
        return None
    depth = 0
    in_string = False
    why: str | None = None
    i = 1
    while i < len(text) and why is None:
        ch = text[i]
        if in_string:
            if ch == '"' and i + 1 < len(text) and text[i + 1] == '"':
                i += 1
            elif ch == '"':
                in_string = False
            i += 1
            continue
        if ch == '"':
            in_string = True
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth < 0:
                why = "closes a parenthesis it never opened"
        i += 1
    if why is None and in_string:
        why = "opens a string it never closes"
    elif why is None and depth > 0:
        why = "leaves a parenthesis open"
    elif why is None and _ENDS_WITH_OPERATOR.search(text) is not None:
        why = "ends with an operator"
    if why is None:
        return None
    return (
        f"The formula {toks[0].raw_text} {why}, so Excel cannot parse it. This will raise "
        "Run-time error '1004', unless the cell is formatted as Text, which keeps the String as text."
    )


def host_union_property_value_problem(
    owners: Sequence[str], name: str, value_tokens: Sequence[VbaToken]
) -> str | None:
    """The same, for a receiver that is one of several host types: `ActiveSheet`,
    a Worksheet or a Chart, and `Sheets(1)`. Judged only when each type refuses
    the value the same way (XLIDE issue #416), and told for the first."""
    # Reached late-bound, a sheet's Visible refuses any String but a number with
    # 1004, "True" too, where a typed Worksheet's raises 13 (measured in Excel
    # 16.0 on 2026-10-02).
    if (
        len(owners) == 0
        or not all(owner == "Excel.Worksheet" or owner == "Excel.Chart" for owner in owners)
        or name.lower() != "visible"
    ):
        return None
    toks = [tok for tok in value_tokens if tok.kind is not TokenKind.COMMENT]
    if len(toks) != 1 or toks[0].kind is not TokenKind.STRING_LITERAL or _DIGIT.search(toks[0].raw_text) is not None:
        return None
    error = _excel_1004("Visible", "Worksheet")
    return (
        f"Visible, on a sheet reached late-bound, takes an xlSheetVisibility constant, and the String "
        f"{toks[0].raw_text} is not one. This will raise Run-time error '{error.number}': {error.text}."
    )


def host_property_value_problem(
    target: _OwnedMember,
    value_tokens: Sequence[VbaToken],
    known: Callable[[Sequence[VbaToken]], int | float | None] | None = None,
) -> str | None:
    """The message for a host property Let the host refuses, or None."""
    string_problem = _host_property_string_problem(target, value_tokens)
    if string_problem:
        return string_problem
    owner_limits = _LIMITS.get(target.owner)
    limit = owner_limits.get(target.name.lower()) if owner_limits is not None else None
    # A local known to hold a number counts too: `n = 500` then `Font.Size = n`
    # (XLIDE issue #346).
    value: int | float | None = None
    if limit is not None:
        value = signed_numeric_literal(value_tokens)
        if value is None and known is not None:
            value = known(value_tokens)
    if limit is None or value is None or (limit.allowed is not None and value in limit.allowed):
        return None
    refused = any(
        (span.from_ is None or value >= span.from_) and (span.to is None or value <= span.to)
        for span in limit.refused
    )
    if not refused:
        return None
    bare = target.owner[target.owner.find(".") + 1 :]
    return (
        f"{bare}.{target.name} takes {limit.runs}; {js_number_to_string(value)} is outside that. "
        f"This will raise Run-time error '{limit.error.number}': {limit.error.text}."
    )
