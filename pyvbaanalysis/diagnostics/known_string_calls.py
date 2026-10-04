"""The whole numbers InStr, InStrRev, Len, Asc and AscW return for strings the code
states (XLIDE issue #201).

The "not found" bug reads `Left$(s, InStr(s, " ") - 1)`: with no space InStr
returns 0, and Left$ gets -1. The value rules evaluate an argument's arithmetic,
but a call inside it stopped them. Here each such call whose strings are
literals, or locals holding a literal, becomes its number, so the arithmetic
around it can be evaluated. Measured in Excel 16.0:

 - InStr([start,] s1, s2 [, compare]) is the first position of s2 in s1 at or
   after start, 0 when there is none, when s1 is "" or when start is past the
   end, and start itself when s2 is "", past the end too.
 - InStrRev(s1, s2 [, start [, compare]]) is the last position of s2 in s1 ending
   at or before start (-1, the default, is the end), 0 when there is none or
   start is past the end, and start (the length for -1) when s2 is "".
 - Comparison is binary unless the call passes vbTextCompare (1) or the module
   says Option Compare Text; Option Compare Database follows the database's
   locale, so a call is decided there only when both agree.
 - Asc of a character above 127 depends on the code page and is left alone.

Ported from xlide_vscode/src/analyzer/diagnostics/knownStringCalls.ts. Strings
are measured, searched and sliced in UTF-16 code units, as JavaScript's are, and
a date is a JavaScript-style UTC time value (UtcDate).
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from typing import Literal

from ..js_compat import JS_WHITESPACE, js_number_to_string, js_trim, utf16_length
from ..lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ..lexer.token_kinds import TokenKind, VbaToken
from .call_extraction import string_literal_value
from .walker import token_name, token_text

ModuleCompare = Literal["binary", "text", "database"]


# -- JavaScript Date, UTC only ------------------------------------------------

_MS_PER_DAY = 86400000
# TimeClip: a time value past 8.64e15 ms either side of the epoch is NaN.
_MAX_TIME = 8.64e15


def _days_from_civil(year: int, month: int, day: int) -> int:
    """Days since 1970-01-01 of a proleptic Gregorian date (month 1-12), any year."""
    y = year - 1 if month <= 2 else year
    era = y // 400
    year_of_era = y - era * 400
    day_of_year = (153 * ((month + 9) % 12) + 2) // 5 + day - 1
    day_of_era = year_of_era * 365 + year_of_era // 4 - year_of_era // 100 + day_of_year
    return era * 146097 + day_of_era - 719468


def _civil_from_days(days: int) -> tuple[int, int, int]:
    """The proleptic Gregorian (year, month 1-12, day) of a day count since 1970-01-01."""
    z = days + 719468
    era = z // 146097
    day_of_era = z - era * 146097
    year_of_era = (day_of_era - day_of_era // 1460 + day_of_era // 36524 - day_of_era // 146096) // 365
    day_of_year = day_of_era - (365 * year_of_era + year_of_era // 4 - year_of_era // 100)
    mp = (5 * day_of_year + 2) // 153
    day = day_of_year - (153 * mp + 2) // 5 + 1
    month = mp + 3 if mp < 10 else mp - 9
    return (year_of_era + era * 400 + (1 if month <= 2 else 0), month, day)


def _make_day(year: int, month: int, date: int) -> int:
    """ECMAScript MakeDay: a month outside 0-11 carries into the year, and a date
    outside the month carries into the next or previous one."""
    return _days_from_civil(year + month // 12, month % 12 + 1, 1) + date - 1


def _time_clip(time: float) -> float:
    if not math.isfinite(time) or abs(time) > _MAX_TIME:
        return math.nan
    return float(math.trunc(time))


@dataclass(frozen=True, slots=True)
class UtcDate:
    """A JavaScript Date read in UTC: `time` is milliseconds since the epoch, NaN
    for an invalid date."""

    time: float

    def get_time(self) -> float:
        return self.time

    def _civil(self) -> tuple[int, int, int] | None:
        if math.isnan(self.time):
            return None
        return _civil_from_days(math.floor(self.time / _MS_PER_DAY))

    def get_utc_full_year(self) -> float:
        civil = self._civil()
        return math.nan if civil is None else civil[0]

    def get_utc_month(self) -> float:
        """The month, 0 for January, as JavaScript counts it."""
        civil = self._civil()
        return math.nan if civil is None else civil[1] - 1

    def get_utc_date(self) -> float:
        civil = self._civil()
        return math.nan if civil is None else civil[2]


def date_utc(year: int, month: int, day: int = 1, hours: int = 0, minutes: int = 0, seconds: int = 0) -> float:
    """Date.UTC: a year from 0 to 99 is 1900 plus the year."""
    full_year = 1900 + year if 0 <= year <= 99 else year
    return _time_clip(
        _make_day(full_year, month, day) * _MS_PER_DAY + hours * 3600000 + minutes * 60000 + seconds * 1000
    )


def _set_utc_full_year(date: UtcDate, year: int, month: int, day: int) -> UtcDate:
    """`date.setUTCFullYear(year, month, day)`, keeping the time of day."""
    time = 0.0 if math.isnan(date.time) else date.time
    within_day = time - math.floor(time / _MS_PER_DAY) * _MS_PER_DAY
    return UtcDate(_time_clip(_make_day(year, month, day) * _MS_PER_DAY + within_day))


def _set_utc_hours(date: UtcDate, hours: int, minutes: int, seconds: int, ms: int) -> UtcDate:
    if math.isnan(date.time):
        return date
    day = math.floor(date.time / _MS_PER_DAY)
    return UtcDate(_time_clip(day * _MS_PER_DAY + hours * 3600000 + minutes * 60000 + seconds * 1000 + ms))


def _set_utc_date(date: UtcDate, day: int) -> UtcDate:
    if math.isnan(date.time):
        return date
    days = math.floor(date.time / _MS_PER_DAY)
    within_day = date.time - days * _MS_PER_DAY
    year, month, _ = _civil_from_days(days)
    return UtcDate(_time_clip(_make_day(year, month - 1, day) * _MS_PER_DAY + within_day))


# -- UTF-16 strings -------------------------------------------------------------


def _units(text: str) -> str:
    """The text with each astral character as its two surrogates, so indexing,
    slicing and searching count UTF-16 code units as JavaScript does."""
    if all(ord(ch) <= 0xFFFF for ch in text):
        return text
    out: list[str] = []
    for ch in text:
        code = ord(ch)
        if code > 0xFFFF:
            code -= 0x10000
            out.append(chr(0xD800 + (code >> 10)))
            out.append(chr(0xDC00 + (code & 0x3FF)))
        else:
            out.append(ch)
    return "".join(out)


def _from_units(units: str) -> str:
    """Rejoins each surrogate pair _units split; a lone surrogate stays as it is."""
    if all(not (0xD800 <= ord(ch) <= 0xDFFF) for ch in units):
        return units
    out: list[str] = []
    i = 0
    while i < len(units):
        high = ord(units[i])
        low = ord(units[i + 1]) if i + 1 < len(units) else 0
        if 0xD800 <= high <= 0xDBFF and 0xDC00 <= low <= 0xDFFF:
            out.append(chr(0x10000 + ((high - 0xD800) << 10) + (low - 0xDC00)))
            i += 2
        else:
            out.append(units[i])
            i += 1
    return "".join(out)


def _to_integer(value: float) -> float:
    """ECMAScript ToIntegerOrInfinity."""
    if math.isnan(value):
        return 0
    if math.isinf(value):
        return value
    return math.trunc(value)


def _slice(text: str, start: float, end: float | None = None) -> str:
    """String.prototype.slice over UTF-16 code units."""
    units = _units(text)
    length = len(units)

    def clamp(index: float) -> int:
        value = _to_integer(index)
        if value < 0:
            return int(max(length + value, 0))
        return int(min(value, length))

    begin = clamp(start)
    stop = length if end is None else clamp(end)
    return _from_units(units[begin:stop]) if begin < stop else ""


def _repeat(text: str, count: float) -> str:
    return text * int(_to_integer(count))


def _first_unit(text: str) -> str:
    """`text[0]`: the first UTF-16 code unit, a lone high surrogate for an astral character."""
    return _units(text)[0]


# -- Option Compare -------------------------------------------------------------

_last_compare: tuple[str, ModuleCompare] | None = None

# `^` with JavaScript's m flag starts a line after LF, CR, U+2028 or U+2029; `i`
# folds ASCII case only here, and `\b` is an ASCII word boundary.
_OPTION_COMPARE_RE = re.compile(
    "(?:^|(?<=[\n\râ€¨â€©]))[ \t]*Option[ \t]+Compare[ \t]+(Binary|Text|Database)\\b",
    re.IGNORECASE | re.ASCII,
)


def module_compare(source: str) -> ModuleCompare:
    """The module's Option Compare, from its source text. The array rules ask once
    per procedure, so the last module's answer is kept."""
    global _last_compare
    if _last_compare is not None and (_last_compare[0] is source or _last_compare[0] == source):
        return _last_compare[1]
    match = _OPTION_COMPARE_RE.search(source)
    word = match.group(1).lower() if match is not None else "binary"
    compare: ModuleCompare = "text" if word == "text" else "database" if word == "database" else "binary"
    _last_compare = (source, compare)
    return compare


@dataclass(frozen=True, slots=True)
class KnownStringCallContext:
    # Locals whose value is a known literal, by lower-cased name.
    known_strings: Mapping[str, str]
    # The whole number an argument's text evaluates to, if any.
    integer_value: Callable[[str], int | float | None]
    # Whether the project declares a procedure of this name, hiding VBA's.
    shadowed: Callable[[str], bool]
    compare: ModuleCompare
    # The date a local holds, by lower-cased name, where it is known (issue #559).
    date_of: Callable[[str], UtcDate | None] | None = None


_FOLDED: frozenset[str] = frozenset(
    {"instr", "instrrev", "len", "asc", "ascw", "year", "month", "day", "datediff"}
)

_WS = "[" + JS_WHITESPACE + "]"
# JavaScript's `\d` is ASCII, its `\s` the JavaScript whitespace, its `$` the end.
_DATE_LITERAL_RE = re.compile(
    "#" + _WS + "*([0-9]{1,2})/([0-9]{1,2})/([0-9]{3,4})" + _WS + "*"
    "(?:([0-9]{1,2}):([0-9]{2})(?::([0-9]{2}))?" + _WS + "*([Aa][Mm]|[Pp][Mm])?)?" + _WS + "*#"
)


def parse_date_literal(raw: str) -> UtcDate | None:
    """A `#m/d/yyyy#` date literal as a UTC date, or None for any other spelling."""
    # A year of three digits is that year: #1/1/100# (issue #262).
    match = _DATE_LITERAL_RE.fullmatch(raw)
    if match is None:
        return None
    month = int(match.group(1))
    day = int(match.group(2))
    year = int(match.group(3))
    if month < 1 or month > 12 or day < 1 or day > 31 or year < 100:
        return None
    hour = int(match.group(4) or 0)
    if match.group(7):
        hour = hour % 12 + (12 if match.group(7).upper() == "PM" else 0)
    date = _set_utc_full_year(UtcDate(0.0), year, month - 1, day)
    return _set_utc_hours(date, hour, int(match.group(5) or 0), int(match.group(6) or 0), 0)


def _strip_outer_parens(arg: Sequence[VbaToken]) -> list[VbaToken]:
    toks = [tok for tok in arg if tok.kind is not TokenKind.COMMENT]
    while len(toks) > 2 and toks[0].raw_text == "(" and match_paren_from(toks, 0) == len(toks) - 1:
        toks = toks[1:-1]
    return toks


def _raw(toks: Sequence[VbaToken], i: int) -> str | None:
    return toks[i].raw_text if 0 <= i < len(toks) else None


def _token_at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


_ISO_DATE_RE = re.compile(r"([0-9]{4})-([0-9]{1,2})-([0-9]{1,2})")
_SLASHED_DATE_RE = re.compile(r"([0-9]{1,2})[/-]([0-9]{1,2})[/-]([0-9]{4})")


def known_date(
    arg: Sequence[VbaToken],
    integer_value: Callable[[Sequence[VbaToken]], int | float | None],
    date_of: Callable[[str], UtcDate | None] | None = None,
) -> UtcDate | None:
    """The date a date expression names, where it is known (issue #510): a date
    literal, `DateSerial` of whole numbers from year 100 on, and `DateValue` or
    `CDate` of a string every locale reads alike, "12/31/9999" with a day past 12
    or "9999-12-31". None for anything else, and for a date past the range."""
    toks = _strip_outer_parens(arg)
    if len(toks) == 1 and toks[0].kind is TokenKind.DATE_LITERAL:
        return parse_date_literal(toks[0].raw_text)
    local_name = token_name(toks[0]) if len(toks) == 1 and date_of is not None else None
    local = local_name.lower() if local_name is not None else None
    if local and date_of is not None:
        return date_of(local)
    at = 2 if token_text(_token_at(toks, 0)) == "vba" and _raw(toks, 1) == "." else 0
    name = token_text(_token_at(toks, at))
    if _raw(toks, at + 1) != "(" or match_paren_from(toks, at + 1) != len(toks) - 1:
        return None
    args = split_top_level_token_groups(toks, at + 2, ",", len(toks) - 1)
    if name == "dateserial" and len(args) == 3:
        parts = [integer_value(group) for group in args]
        whole: list[int] = []
        for part in parts:
            if part is None or part < -32768 or part > 32767:
                return None
            whole.append(int(_to_integer(part)))
        if parts[0] is not None and parts[0] < 100:
            return None
        date = _set_utc_full_year(UtcDate(0.0), whole[0], whole[1] - 1, 1)
        date = _set_utc_date(date, whole[2])
        return date if _in_date_range(date) else None
    # CDate of a whole number counts days from December 30, 1899: CDate(-10000) is
    # August 13, 1872 (issue #559, measured in Excel 16.0).
    first = _token_at(args[0], 0) if len(args) > 0 else None
    days = (
        integer_value(args[0])
        if name == "cdate" and len(args) == 1 and (first is None or first.kind is not TokenKind.STRING_LITERAL)
        else None
    )
    if days is not None:
        date = UtcDate(_time_clip(date_utc(1899, 11, 30) + days * _MS_PER_DAY))
        return date if _in_date_range(date) else None
    if (
        (name == "datevalue" or name == "cdate")
        and len(args) == 1
        and len(args[0]) == 1
        and args[0][0].kind is TokenKind.STRING_LITERAL
    ):
        text = js_trim(string_literal_value(args[0][0].raw_text))
        iso = _ISO_DATE_RE.fullmatch(text)
        slashed = _SLASHED_DATE_RE.fullmatch(text)
        found: tuple[int, int, int] | None = None
        if iso is not None:
            found = (int(iso.group(1)), int(iso.group(2)), int(iso.group(3)))
        elif slashed is not None:
            a, b = int(slashed.group(1)), int(slashed.group(2))
            # Only an order every locale agrees on: one of the two is past 12.
            if a > 12 and b <= 12:
                found = (int(slashed.group(3)), b, a)
            elif b > 12 and a <= 12:
                found = (int(slashed.group(3)), a, b)
        if found is None or found[0] < 100 or found[1] < 1 or found[1] > 12 or found[2] < 1:
            return None
        date = _set_utc_full_year(UtcDate(0.0), found[0], found[1] - 1, found[2])
        return date if date.get_utc_date() == found[2] and _in_date_range(date) else None
    return None


def _in_date_range(date: UtcDate) -> bool:
    return date.get_time() >= date_utc(100, 0, 1) and date.get_utc_full_year() <= 9999


def _date_diff(interval: str, from_: UtcDate, to: UtcDate) -> int | float | None:
    """DateDiff of known dates, for the intervals that count whole units."""
    units = {"h": 3600000, "n": 60000, "s": 1000}

    def days(date: UtcDate) -> float:
        return math.floor(date.get_time() / _MS_PER_DAY)

    if interval in ("d", "y"):
        return days(to) - days(from_)
    if interval == "w":
        return math.trunc((days(to) - days(from_)) / 7)
    if interval == "m":
        return (
            (to.get_utc_full_year() - from_.get_utc_full_year()) * 12
            + to.get_utc_month()
            - from_.get_utc_month()
        )
    if interval == "q":
        return (
            (to.get_utc_full_year() - from_.get_utc_full_year()) * 4
            + math.floor(to.get_utc_month() / 3)
            - math.floor(from_.get_utc_month() / 3)
        )
    if interval == "yyyy":
        return to.get_utc_full_year() - from_.get_utc_full_year()
    unit = units.get(interval)
    if unit is None:
        return None
    return math.floor(to.get_time() / unit) - math.floor(from_.get_time() / unit)


def _number_text(value: int | float) -> str:
    return js_number_to_string(value)


def fold_known_string_calls(toks: Sequence[VbaToken], ctx: KnownStringCallContext) -> str | None:
    """The expression's text with each foldable call replaced by its number, or None
    when it has none."""
    from .rules.shared import is_bare_or_vba_qualified_intrinsic_call

    out: list[str] = []
    folded = False
    i = 0
    while i < len(toks):
        tok = toks[i]
        lower = token_text(tok)
        close = (
            match_paren_from(toks, i + 1)
            if lower in _FOLDED
            and _raw(toks, i + 1) == "("
            and is_bare_or_vba_qualified_intrinsic_call(toks, i)
            and not ctx.shadowed(lower)
            else -1
        )
        value = (
            _call_value(lower, split_top_level_token_groups(toks, i + 2, ",", close), ctx)
            if close > i + 1
            else None
        )
        if value is None:
            out.append(tok.raw_text)
            i += 1
            continue
        if _raw(toks, i - 1) == ".":
            del out[-2:]  # `VBA.InStr(...)`
        out.append(f"({_number_text(value)})" if value < 0 else _number_text(value))
        folded = True
        i = close + 1
    return " ".join(out) if folded else None


def _call_value(name: str, args: list[list[VbaToken]], ctx: KnownStringCallContext) -> int | float | None:
    if name == "len":
        text = _known_string(args[0], ctx) if len(args) == 1 else None
        return None if text is None else utf16_length(text)
    if name == "asc" or name == "ascw":
        text = _known_string(args[0], ctx) if len(args) == 1 else None
        code = ord(_first_unit(text)) if text else None
        return code if code is not None and (name == "ascw" or code < 128) else None
    if name == "instr":
        return _in_str(args, ctx)
    if name == "instrrev":
        return _in_str_rev(args, ctx)

    def whole(toks: Sequence[VbaToken]) -> int | float | None:
        return _whole_number(toks, ctx)

    # Parts of a known date, and DateDiff between two (issue #510).
    if name in ("year", "month", "day"):
        date = known_date(args[0], whole, ctx.date_of) if len(args) == 1 else None
        if date is None:
            return None
        part = (
            date.get_utc_full_year()
            if name == "year"
            else date.get_utc_month() + 1
            if name == "month"
            else date.get_utc_date()
        )
        return _js_value(part)
    if name == "datediff":
        interval_text = _known_string(args[0], ctx) if len(args) == 3 else None
        interval = interval_text.lower() if interval_text is not None else None
        from_ = known_date(args[1], whole, ctx.date_of) if len(args) == 3 else None
        to = known_date(args[2], whole, ctx.date_of) if len(args) == 3 else None
        if interval is None or from_ is None or to is None:
            return None
        diff = _date_diff(interval, from_, to)
        return None if diff is None else _js_value(diff)
    return None


def _js_value(number: int | float) -> int | float:
    """An int where JavaScript holds the number as a whole number exactly."""
    if isinstance(number, float) and number.is_integer() and abs(number) < 2**53:
        return int(number)
    return number


def _in_str(args: list[list[VbaToken]], ctx: KnownStringCallContext) -> int | float | None:
    # With a start, the strings move one place right; a compare needs a start.
    with_start = len(args) >= 3
    if len(args) < 2 or len(args) > 4:
        return None
    start = _whole_number(args[0], ctx) if with_start else 1
    s1 = _known_string(args[1 if with_start else 0], ctx)
    s2 = _known_string(args[2 if with_start else 1], ctx)
    text = _compare_mode(args[3] if len(args) > 3 else None, ctx)
    if start is None or start < 1 or s1 is None or s2 is None or text is None:
        return None
    first = start

    def search(a: str, b: str) -> int | float:
        # An empty s1 gives 0, whatever s2 is: InStr("", "") is 0 (issue #509,
        # measured in Excel 16.0).
        if len(a) == 0:
            return 0
        # An empty s2 is found at start, even past the end: InStr(5, "abc", "") is 5.
        if len(b) == 0:
            return first
        return _index_of(a, b, first - 1) + 1

    return _decided(s1, s2, text, search)


def _in_str_rev(args: list[list[VbaToken]], ctx: KnownStringCallContext) -> int | float | None:
    if len(args) < 2 or len(args) > 4:
        return None
    s1 = _known_string(args[0], ctx)
    s2 = _known_string(args[1], ctx)
    start = _whole_number(args[2], ctx) if len(args) > 2 else -1
    text = _compare_mode(args[3] if len(args) > 3 else None, ctx)
    if s1 is None or s2 is None or start is None or (start < 1 and start != -1) or text is None:
        return None
    last = start

    def search(a: str, b: str) -> int | float:
        a_units = _units(a)
        b_units = _units(b)
        end = len(a_units) if last == -1 else last
        if end > len(a_units):
            return 0
        if len(b_units) == 0:
            return end
        if end < len(b_units):
            return 0
        return _last_index_of(a_units, b_units, end - len(b_units)) + 1

    return _decided(s1, s2, text, search)


def _index_of(text: str, search: str, position: float) -> int:
    """String.prototype.indexOf over UTF-16 code units."""
    units = _units(text)
    start = int(min(max(_to_integer(position), 0), len(units)))
    return units.find(_units(search), start)


def _last_index_of(units: str, search: str, position: float) -> int:
    """String.prototype.lastIndexOf over code units already split by _units."""
    start = int(min(max(_to_integer(position), 0), len(units)))
    return units.rfind(search, 0, start + len(search))


_ASCII_RE = re.compile(r"[\x00-\x7f]*")


def _is_ascii(text: str) -> bool:
    return _ASCII_RE.fullmatch(text) is not None


def _decided(
    s1: str,
    s2: str,
    text: bool | Literal["either"],
    search: Callable[[str, str], int | float],
) -> int | float | None:
    """The search's result under the comparison that applies, or None where it
    cannot be told: a text comparison over characters beyond ASCII, whose case
    rules are the locale's, or a Database comparison the two disagree on."""
    binary = search(s1, s2)
    ascii_only = _is_ascii(s1 + s2)
    insensitive = search(s1.lower(), s2.lower()) if ascii_only else None
    if text == "either":
        return binary if insensitive == binary else None
    return insensitive if text else binary


def _compare_mode(arg: Sequence[VbaToken] | None, ctx: KnownStringCallContext) -> bool | Literal["either"] | None:
    """True for a text comparison, False for binary, 'either' under Option Compare Database."""
    if arg is None:
        return True if ctx.compare == "text" else "either" if ctx.compare == "database" else False
    toks = [tok for tok in arg if tok.kind is not TokenKind.COMMENT]
    word_name = token_name(toks[0]) if len(toks) == 1 else None
    word = word_name.lower() if word_name is not None else None
    value = 0 if word == "vbbinarycompare" else 1 if word == "vbtextcompare" else _whole_number(toks, ctx)
    return False if value == 0 else True if value == 1 else None


def _known_string(arg: Sequence[VbaToken], ctx: KnownStringCallContext) -> str | None:
    def name_value(tok: VbaToken) -> str | None:
        name = token_name(tok)
        return ctx.known_strings.get(name.lower() if name is not None else "")

    return fold_string_expression(
        arg,
        StringFoldContext(
            name_value=name_value,
            integer_value=lambda toks: _whole_number(toks, ctx),
            shadowed=ctx.shadowed,
        ),
    )


@dataclass(frozen=True, slots=True)
class StringFoldContext:
    """What a string expression is folded with."""

    # A String local's known value, by lowercased name.
    name_value: Callable[[VbaToken], str | None]
    # A whole-number argument's value.
    integer_value: Callable[[Sequence[VbaToken]], int | float | None]
    # Whether the project declares a procedure of this name, hiding VBA's.
    shadowed: Callable[[str], bool] | None = None


# The VBA string functions the folder runs, each with or without its `$`.
_STRING_FUNCTIONS: frozenset[str] = frozenset(
    {"left", "right", "mid", "lcase", "ucase", "strreverse", "trim", "ltrim", "rtrim", "space", "string", "replace"}
)

_LETTER_RE = re.compile(r"[A-Za-z]")


def fold_string_expression(arg: Sequence[VbaToken], ctx: StringFoldContext) -> str | None:
    """The text a string expression gives, where every part is known: literals,
    String locals known to hold one, `&` between them, and Left, Right, Mid, LCase,
    UCase, StrReverse, Trim, LTrim, RTrim, Space, String and Replace of known
    arguments (issue #509). None for anything else, and for a call that would
    raise."""
    toks = _strip_outer_parens(arg)
    if len(toks) == 0:
        return None
    parts = split_top_level_token_groups(toks, 0, "&", len(toks))
    if len(parts) > 1:
        texts: list[str] = []
        for part in parts:
            folded = fold_string_expression(part, ctx)
            if folded is None:
                return None
            texts.append(folded)
        return "".join(texts)
    if len(toks) == 1:
        if toks[0].kind is TokenKind.STRING_LITERAL:
            return string_literal_value(toks[0].raw_text)
        name = token_name(toks[0])
        if name is None or toks[0].kind is TokenKind.KEYWORD:
            return "" if token_text(toks[0]) == "vbnullstring" else None
        return ctx.name_value(toks[0])
    # `Mid$(s, 1)` lexes as Mid, a `$` and the arguments.
    at = 2 if token_text(toks[0]) == "vba" and _raw(toks, 1) == "." else 0
    function = token_text(_token_at(toks, at))
    at += 1 if _raw(toks, at + 1) == "$" else 0
    if (
        function not in _STRING_FUNCTIONS
        or _raw(toks, at + 1) != "("
        or match_paren_from(toks, at + 1) != len(toks) - 1
        or (ctx.shadowed is not None and ctx.shadowed(function))
    ):
        return None
    args = split_top_level_token_groups(toks, at + 2, ",", len(toks) - 1)

    def text(k: int) -> str | None:
        return fold_string_expression(args[k], ctx) if k < len(args) else None

    def whole(k: int) -> int | float | None:
        return ctx.integer_value(args[k]) if k < len(args) and len(args[k]) > 0 else None

    if function == "left" or function == "right":
        s = text(0)
        n = whole(1)
        if len(args) != 2 or s is None or n is None or n < 0:
            return None
        return _slice(s, 0, n) if function == "left" else _slice(s, max(0, utf16_length(s) - n))
    if function == "mid":
        s = text(0)
        start = whole(1)
        length = whole(2) if len(args) == 3 else (utf16_length(s) if s is not None else None)
        if (
            len(args) < 2
            or len(args) > 3
            or s is None
            or start is None
            or length is None
            or start < 1
            or length < 0
        ):
            return None
        return _slice(s, start - 1, start - 1 + length)
    if function in ("lcase", "ucase", "strreverse", "trim", "ltrim", "rtrim"):
        s = text(0) if len(args) == 1 else None
        if s is None or (not _is_ascii(s) and (function == "lcase" or function == "ucase")):
            return None
        if function == "lcase":
            return s.lower()
        if function == "ucase":
            return s.upper()
        if function == "strreverse":
            return "".join(reversed(s))
        if function == "trim":
            return s.strip(" ")
        if function == "ltrim":
            return s.lstrip(" ")
        return s.rstrip(" ")
    if function == "space":
        n = whole(0) if len(args) == 1 else None
        return None if n is None or n < 0 or n > 65535 else _repeat(" ", n)
    if function == "string":
        n = whole(0)
        c = text(1)
        if len(args) != 2 or n is None or n < 0 or n > 65535 or not c:
            return None
        return _from_units(_repeat(_first_unit(c), n))
    if function == "replace":
        s, find, by = text(0), text(1), text(2)
        if len(args) != 3 or s is None or find is None or by is None:
            return None
        # A letter in find matches by Option Compare, which this does not read.
        if _LETTER_RE.search(find):
            return None
        return s if len(find) == 0 else s.replace(find, by)
    return None


def _whole_number(arg: Sequence[VbaToken], ctx: KnownStringCallContext) -> int | float | None:
    """A whole-number argument, itself folded first: `InStr(InStr(s, "X") + 1, s, "X")`."""
    toks = [tok for tok in arg if tok.kind is not TokenKind.COMMENT]
    if len(toks) == 0:
        return None
    folded = fold_known_string_calls(toks, ctx)
    text = folded if folded is not None else " ".join(tok.raw_text for tok in toks)
    value = ctx.integer_value(text)
    return value if value is not None and math.isfinite(value) and float(value).is_integer() else None
