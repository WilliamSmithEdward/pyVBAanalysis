"""Ported from xlide_vscode/src/analyzer/constants/dateLiteral.ts."""

from __future__ import annotations

import datetime
import re

from ..js_compat import JS_WHITESPACE, js_trim

DAY_MS = 86400000
# Serial 0, December 30, 1899, as a UTC time.
DATE_EPOCH_MS = -2209161600000

_EPOCH_DATE = datetime.date(1899, 12, 30)

_WS = "[" + re.escape(JS_WHITESPACE) + "]"
_DATE_LITERAL_RE = re.compile(
    r"^(?:(\d{1,2})/(\d{1,2})/(\d{3,4})|(\d{3,4})-(\d{1,2})-(\d{1,2}))?"
    + _WS
    + r"*(?:(\d{1,2}):(\d{2})(?::(\d{2}))?"
    + _WS
    + r"*([AaPp][Mm])?)?$",
    re.ASCII,
)


def date_literal_serial(raw: str) -> float | None:
    """A date literal's serial: the days from December 30, 1899, with the time
    of day as the fraction. ``#12/31/9999#`` is 2958465, ``#1/1/100#`` is
    -657434 (XLIDE issue #203), ``#12:00:00 PM#`` is 0.5 and
    ``#1/1/2000 6:00 AM#`` is 36526.25 (issue #208). Read are ``#m/d/yyyy#``
    and ``#yyyy-mm-dd#``, a time of ``h:mm[:ss]`` with an optional AM or PM,
    and the two together; a month name or a two-digit year is left to the VBE."""
    text = raw
    if text.startswith("#"):
        text = text[1:]
    if text.endswith("#"):
        text = text[:-1]
    text = js_trim(text)
    # JS `$` matches only at the very end; Python's also before a final "\n",
    # which js_trim has already removed.
    match = _DATE_LITERAL_RE.match(text)
    if match is None or len(text) == 0:
        return None
    us_month, us_day, us_year, iso_year, iso_month, iso_day, hours, minutes, seconds, meridiem = match.groups()
    serial: float = 0
    if us_year is not None or iso_year is not None:
        year = int(us_year if us_year is not None else iso_year)
        month = int(us_month if us_month is not None else iso_month)
        day = int(us_day if us_day is not None else iso_day)
        if year < 100 or month < 1 or month > 12 or day < 1:
            return None
        try:
            at = datetime.date(year, month, day)
        except ValueError:
            return None  # #2/30/2020# is no date
        serial = (at - _EPOCH_DATE).days
    if hours is not None:
        hour = int(hours)
        minute = int(minutes)
        second = int(seconds) if seconds is not None else 0
        if meridiem is not None:
            if hour < 1 or hour > 12:
                return None
            hour = hour % 12 + (12 if meridiem[0] in "Pp" else 0)
        if hour > 23 or minute > 59 or second > 59:
            return None
        fraction = (hour * 3600 + minute * 60 + second) / 86400
        # Before 1899-12-30 the time runs the other way: -1.25 is 12/29/1899 6 AM.
        serial = serial - fraction if serial < 0 else serial + fraction
    return serial
