"""JavaScript text and number semantics the port reproduces.

Upstream's rules run on JavaScript strings and numbers, and a few of its built-ins
differ from their nearest Python counterparts in ways a diagnostic message or a
parse can show: String.prototype.trim strips U+FEFF where str.strip() keeps it and
keeps U+001C-U+001F and U+0085 where str.strip() strips them, and
Number.prototype.toString prints 1 and 1e+21 where Python prints 1.0 and 1e+21 but
also 1e-07 where JavaScript prints 1e-7. Rules use these helpers wherever the
difference could reach their output.
"""

from __future__ import annotations

import math
import re
from decimal import Decimal

# ECMAScript WhiteSpace (TAB, VT, FF, SP, NBSP, U+FEFF and the other Zs spaces) and
# LineTerminator (LF, CR, U+2028, U+2029): what String.prototype.trim strips and
# what `\s` matches in a JavaScript regular expression.
JS_WHITESPACE = "".join(
    chr(code)
    for code in (
        0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x20, 0xA0, 0x1680,
        0x2000, 0x2001, 0x2002, 0x2003, 0x2004, 0x2005, 0x2006, 0x2007, 0x2008, 0x2009, 0x200A,
        0x2028, 0x2029, 0x202F, 0x205F, 0x3000, 0xFEFF,
    )
)


def js_trim(text: str) -> str:
    """String.prototype.trim."""
    return text.strip(JS_WHITESPACE)


def utf16_length(text: str) -> int:
    """String.prototype.length: UTF-16 code units, two for an astral character."""
    return len(text) + sum(1 for ch in text if ord(ch) > 0xFFFF)


# `[0-9]`, not `\d`: JavaScript's digits are ASCII only.
_JS_DECIMAL_RE = re.compile(r"[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?")
_JS_RADIX_RE = re.compile(r"0(?:[xX]([0-9A-Fa-f]+)|[oO]([0-7]+)|[bB]([01]+))")


def js_number(text: str) -> float:
    """Number(string): a decimal literal, Infinity, or a 0x/0o/0b whole number,
    after trimming whitespace; an empty string is 0 and anything else NaN. Python's
    float() differs: it refuses an empty string and takes `inf`, `nan` and `1_0`."""
    trimmed = js_trim(text)
    if trimmed == "":
        return 0.0
    if _JS_DECIMAL_RE.fullmatch(trimmed) is not None:
        return float(trimmed)
    if trimmed in ("Infinity", "+Infinity"):
        return math.inf
    if trimmed == "-Infinity":
        return -math.inf
    radix = _JS_RADIX_RE.fullmatch(trimmed)
    if radix is None:
        return math.nan
    hex_digits, octal_digits, binary_digits = radix.groups()
    if hex_digits is not None:
        digits, base = hex_digits, 16
    elif octal_digits is not None:
        digits, base = octal_digits, 8
    else:
        digits, base = binary_digits, 2
    try:
        return float(int(digits, base))
    except OverflowError:
        return math.inf


# Integers below this are exact as a JavaScript number, so they print as Python
# prints them.
_EXACT_INTEGER_LIMIT = 2**53


def js_number_to_string(value: float) -> str:
    """Number.prototype.toString() for a JavaScript number: the shortest digits that
    read back as the same double, laid out the way ECMAScript's Number::toString
    lays them out (no `.0` on an integer, exponent form from 1e21 up and below
    1e-6, `e-7` rather than `e-07`)."""
    if isinstance(value, int) and not isinstance(value, bool) and -_EXACT_INTEGER_LIMIT < value < _EXACT_INTEGER_LIMIT:
        return str(value)
    number = float(value)
    if math.isnan(number):
        return "NaN"
    if math.isinf(number):
        return "Infinity" if number > 0 else "-Infinity"
    if number == 0:
        return "0"
    # repr() gives the shortest round-trip digits, the same digits ECMAScript
    # requires; only their layout differs.
    _, digit_tuple, exponent = Decimal(repr(abs(number))).normalize().as_tuple()
    assert isinstance(exponent, int)
    digits = "".join(str(digit) for digit in digit_tuple)
    k = len(digits)
    n = exponent + k
    if k <= n <= 21:
        body = digits + "0" * (n - k)
    elif 0 < n <= 21:
        body = f"{digits[:n]}.{digits[n:]}"
    elif -6 < n <= 0:
        body = "0." + "0" * (-n) + digits
    else:
        power = n - 1
        suffix = f"e{'+' if power >= 0 else '-'}{abs(power)}"
        body = digits + suffix if k == 1 else f"{digits[0]}.{digits[1:]}{suffix}"
    return ("-" if number < 0 else "") + body
