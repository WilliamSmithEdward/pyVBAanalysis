"""Ported from xlide_vscode/src/analyzer/diagnostics/stringConversion.ts.

What VBA makes of a string it converts to a number, a Boolean or a Date
(XLIDE issue #188). Measured in Excel 16.0 (build 20326, en-US, 2026-09-29),
with `x = "..."` into a typed variable and through CLng, CDbl, CBool and
CDate, which agree.

A string converts through the locale: its decimal point, thousands separator
and currency symbol. "2.5", "1,000", "1 000" and "$5" read differently, or not
at all, under another locale, so a string is judged invalid only when no locale
can read it:

 - Numbers. Hex and octal strings convert like the literal ("&HFF" is 255,
   "&HFFFF" is -1, "&17" is 15). Otherwise a sign before or after, a
   parenthesized negative, a currency symbol at either end, digits with
   separators and an exponent run. Empty, no digits at all ("abc", ".", "$"),
   a sign before &H ("-&H10"), and any other character among the digits
   ("4x2", "5%", "1/2/2020", "12:30") raise 13.
 - Boolean. "True" and "False" in any case, exactly, and any number, which is
   True unless zero. " True " with blanks raises 13, as do words.
 - Date. Any number converts, "$5" included. No digits at all ("May",
   "Monday", "True", ".") raises 13, and so does a character no date or number
   uses ("5%", "-&H10").

The patterns use `[0-9]` where upstream's use `\\d`: a JavaScript regular
expression's digits are ASCII only.
"""

from __future__ import annotations

import math
import re
import unicodedata
from dataclasses import dataclass
from typing import Literal

from ..constants.integer_constant_expression import parse_vba_integer_literal
from ..js_compat import js_number


@dataclass(frozen=True, slots=True)
class NumericStringVerdict:
    """How a string converts to a number: never ('invalid'), or ('number') with
    the value when every locale reads it alike."""

    kind: Literal["invalid", "number"]
    value: float | None = None


_INVALID = NumericStringVerdict("invalid")
_NUMBER = NumericStringVerdict("number")

# The Date serials of 1/1/100 and 12/31/9999, the range a Date holds.
DATE_SERIAL_MIN = -657434
DATE_SERIAL_MAX = 2958465

_RADIX_STRING_RE = re.compile(r"&[Hh][0-9A-Fa-f]+|&[Oo]?[0-7]+")
_BARE_OCTAL_RE = re.compile(r"&[0-7]")
_DIGIT_RE = re.compile(r"[0-9]")
# The blanks a number's digits may group with include the no-break space.
_NUMBER_BODY_RE = re.compile("(?:[0-9]|[.,][0-9])[0-9.,' \t" + chr(0xA0) + "]*(?:[eEdD][+-]?[0-9]+)?")
_PLAIN_NUMBER_RE = re.compile(r"[0-9]+(?:[eEdD][+-]?[0-9]+)?")
_D_EXPONENT_RE = re.compile(r"[dD]")


def _strip_blank_edges(text: str) -> str:
    """`text.replace(/^[ \\t]+|[ \\t]+$/g, '')`."""
    return text.strip(" \t")


def _is_currency_or_letter(ch: str) -> bool:
    """`[\\p{Sc}\\p{L}]`."""
    category = unicodedata.category(ch)
    return category == "Sc" or category[0] == "L"


def _strip_currency(body: str) -> str:
    """`body.replace(/^[\\p{Sc}\\p{L}]+\\.?[ \\t]*/u, '').replace(/[ \\t]*[\\p{Sc}\\p{L}]+\\.?$/u, '')`."""
    end = 0
    while end < len(body) and _is_currency_or_letter(body[end]):
        end += 1
    if end > 0:
        if end < len(body) and body[end] == ".":
            end += 1
        while end < len(body) and body[end] in " \t":
            end += 1
        body = body[end:]
    stop = len(body)
    if stop > 0 and body[-1] == ".":
        stop -= 1
    start = stop
    while start > 0 and _is_currency_or_letter(body[start - 1]):
        start -= 1
    if start < stop:
        while start > 0 and body[start - 1] in " \t":
            start -= 1
        body = body[:start]
    return body


def numeric_string_verdict(text: str) -> NumericStringVerdict:
    trimmed = _strip_blank_edges(text)
    if len(trimmed) == 0:
        return _INVALID
    if _RADIX_STRING_RE.fullmatch(trimmed):
        # "&17" is octal like "&O17".
        value = parse_vba_integer_literal(f"&O{trimmed[1:]}" if _BARE_OCTAL_RE.match(trimmed) else trimmed)
        return _NUMBER if value is None else NumericStringVerdict("number", float(value))
    if not _DIGIT_RE.search(trimmed):
        return _INVALID
    body = trimmed
    negative = False
    exact = True
    for _pass in range(6):
        before = body
        if body.startswith("(") and body.endswith(")"):
            body = _strip_blank_edges(body[1:-1])
            negative = not negative
        elif body[:1] in ("+", "-"):
            negative = (not negative) if body[0] == "-" else negative
            body = _strip_blank_edges(body[1:])
        elif body[-1:] in ("+", "-"):
            negative = (not negative) if body[-1] == "-" else negative
            body = _strip_blank_edges(body[:-1])
        else:
            # A currency symbol, which some locale spells in letters ("kr").
            body = _strip_currency(body)
            if body != before:
                exact = False
        if body == before:
            break
    if not _NUMBER_BODY_RE.fullmatch(body):
        return _INVALID
    # Two "." and two ",": a second decimal point whichever of them the
    # locale reads as one. "1,2.3,4.5" raises 13 (issue #504, measured in
    # Excel 16.0); "1.5.5" alone is 155 where "." groups thousands.
    if body.count(".") >= 2 and body.count(",") >= 2:
        return _INVALID
    if not exact or not _PLAIN_NUMBER_RE.fullmatch(body):
        return _NUMBER
    number = js_number(_D_EXPONENT_RE.sub("e", body, count=1))
    if not math.isfinite(number):
        return _NUMBER
    return NumericStringVerdict("number", -number if negative else number)


_VAL_RADIX_RE = re.compile(r"&([Hh])([0-9A-Fa-f]{1,4})(?![0-9A-Fa-f])|&[Oo]?([0-7]{1,5})(?![0-7])")
_VAL_NUMBER_RE = re.compile(r"([-+]?)([0-9]+\.?[0-9]*|\.[0-9]+)?")
_VAL_EXPONENT_RE = re.compile(r"[eEdD][-+]?[0-9]+")


def val_prefix_value(text: str) -> float | None:
    """What Val reads from a string, the same in every locale (issue #703,
    measured in Excel 16.0): blanks, tabs and line feeds anywhere are dropped,
    then the number at the start is read with "." as the decimal point and an
    optional E or D exponent, up to the first character that cannot continue
    it. No number there is 0: Val("abc"), Val("0,5") is 0, Val("1,000") 1,
    Val("1 2 3") 123. None for a hex or octal string past what a plain positive
    literal holds, which this does not follow."""
    compact = re.sub(r"[ \t\n]", "", text)
    radix = _VAL_RADIX_RE.match(compact)
    if radix:
        whole = int(radix.group(2), 16) if radix.group(2) is not None else int(radix.group(3), 8)
        return float(whole) if whole < 0x8000 else None
    if compact.startswith("&"):
        return None
    number = _VAL_NUMBER_RE.match(compact)
    assert number is not None
    if number.group(2) is None:
        return 0.0
    exponent = _VAL_EXPONENT_RE.match(compact, number.end())
    value = js_number(
        f"{number.group(1)}{number.group(2)}{_D_EXPONENT_RE.sub('e', exponent.group(0), count=1) if exponent else ''}"
    )
    return 0.0 if value == 0 else value


_READINGS_RE = re.compile(r"([-+]?)([0-9][0-9.,]*)")


def numeric_string_readings(text: str) -> tuple[float, float] | None:
    """The values a string of digits and separators has where "." is the
    decimal point and "," groups thousands, and where it is the other way
    round: "3.5" is 3.5 and 35 (issue #703). None when either reading fails,
    or for any other spelling."""
    trimmed = _strip_blank_edges(text)
    match = _READINGS_RE.fullmatch(trimmed)
    if not match:
        return None
    sign, digits_text = match.group(1), match.group(2)

    def read(decimal: str, group: str) -> float | None:
        digits = digits_text.replace(group, "")
        if digits.count(decimal) > 1:
            return None
        value = js_number(digits.replace(decimal, ".", 1))
        return (-value if sign == "-" else value) if math.isfinite(value) else None

    dot = read(".", ",")
    comma = read(",", ".")
    return None if dot is None or comma is None else (dot, comma)


def is_invalid_numeric_string(text: str) -> bool:
    """Whether no locale converts the string to a number."""
    return numeric_string_verdict(text).kind == "invalid"


_BOOLEAN_WORD_RE = re.compile(r"true|false", re.IGNORECASE | re.ASCII)


def is_invalid_boolean_string(text: str) -> bool:
    """Whether no locale converts the string to a Boolean."""
    if _BOOLEAN_WORD_RE.fullmatch(text):
        return False
    return is_invalid_numeric_string(text)


def _is_calendar_day(year: int, month: int, day: int) -> bool:
    """Whether the Gregorian calendar has this day: 1900 is no leap year, 2000 is."""
    if month < 1 or month > 12 or day < 1:
        return False
    leap = (year % 4 == 0 and year % 100 != 0) or year % 400 == 0
    days = [31, 29 if leap else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1]
    return day <= days


_MANY_TIME_PARTS_RE = re.compile(r"[0-9]+(?::[0-9]+){3,}")
_TIME_RE = re.compile(r"([0-9]+):([0-9]+)(?::([0-9]+))?")


def is_invalid_time_string(text: str) -> bool:
    """Whether a string written as hours, minutes and seconds names no time:
    "25:00", "10:60", "10:00:60", "1:2:3:4" (issue #262, measured in Excel 16.0
    with CDate, DateValue and TimeValue). "13:00 PM" runs."""
    trimmed = _strip_blank_edges(text)
    if _MANY_TIME_PARTS_RE.fullmatch(trimmed):
        return True
    time = _TIME_RE.fullmatch(trimmed)
    return time is not None and (
        int(time.group(1)) > 23 or int(time.group(2)) > 59 or int(time.group(3) or "0") > 59
    )


_ISO_DATE_RE = re.compile(r"([0-9]{4})-([0-9]{1,2})-([0-9]{1,2})")
_DATE_SEPARATOR_RE = re.compile(r"[/.-]")
_DIGIT_RUN_RE = re.compile(r"[0-9]+")
_DATE_PUNCTUATION = " \t.,/:'-"


def is_invalid_date_string(text: str) -> bool:
    """Whether no locale converts the string to a Date."""
    trimmed = _strip_blank_edges(text)
    # A number every locale reads alike is that day's serial: "&HFF" is
    # 9/11/1900, and one past 12/31/9999 raises (issue #336, measured in
    # Excel 16.0).
    number = numeric_string_verdict(trimmed)
    if number.kind != "invalid" and number.value is not None:
        return number.value < DATE_SERIAL_MIN or number.value >= DATE_SERIAL_MAX + 1
    if not _DIGIT_RE.search(trimmed):
        return True
    if not is_invalid_numeric_string(trimmed):
        return False
    # A year first, then a month and a day in either order, that names no
    # day: "2020-02-30" (issue #239, measured in Excel 16.0).
    iso = _ISO_DATE_RE.fullmatch(trimmed)
    if iso:
        year, a, b = (int(part) for part in iso.groups())
        return year >= 100 and not _is_calendar_day(year, a, b) and not _is_calendar_day(year, b, a)
    # A year past 9999 no locale reads: "1/1/10000" (issue #444, measured in
    # Excel 16.0).
    if _DATE_SEPARATOR_RE.search(trimmed) and any(int(run) > 9999 for run in _DIGIT_RUN_RE.findall(trimmed)):
        return True
    # Dates use letters (month names, AM and PM), digits and these separators.
    return any(
        unicodedata.category(ch)[0] not in "LN" and ch not in _DATE_PUNCTUATION for ch in trimmed
    )
