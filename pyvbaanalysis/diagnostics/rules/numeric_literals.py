"""Rule family: literal tokens the VBE refuses while compiling.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/numericLiterals.ts
(MS-VBAL 3.3.2 number tokens, 3.3.3 date tokens). Each form below was measured in
Excel 16.0 (build 20326): the refused ones are "Syntax error" (XLIDE issues #125
and #133), and the accepted ones at the edges compile.

Type-suffixed integers (suffix_integer_pct_* oracle cases, and #133):
  32768%, -32768%, &H10000%, 1E3%, 1.5%          refused   (% is Integer)
  32767%, &H8000%, &HFFFF%, &O100000%             accepted  (hex wraps to 16 bits)
  2147483648&, -2147483648&, &H100000000&         refused   (& is Long)
  2147483647&, &HFFFFFFFF&                        accepted  (hex wraps to 32 bits)
  9223372036854775808^                            refused   (^ is LongLong)
  9223372036854775807^                            accepted
  &H100000000, &O40000000000 (unsuffixed)         refused   (32 bits at most, issue #369)
  &HFF#, &HFF!                                    refused   (no float suffix on hex)
Type-suffixed floats:
  3.5E+38!                                        refused   (! is Single)
  3.402823E+38!                                   accepted
  922337203685477.5808@, 922337203685478@         refused   (@ is Currency)
  922337203685477.5807@                           accepted
  1.8E+308#, 1E+309 (unsuffixed)                  refused   (Double)
  1.79769313486231E+308#, 1E+308, 99999999999999999999   accepted
Radix prefix with no digits: &H                   refused
Date literals:
  #1/1/10000#, #2/30/2000#, #1/0/2000#            refused   (year, day)
  #25:00#, #1/1/2000 24:00:00#, ##                refused   (time, empty)
  #12/31/9999#, #1/1/100#, #13/1/2000#, #13:00 PM#  accepted (13/1 reads as 13 January)

The `&` suffix is ambiguous with concatenation: `s = 3000000000&"x"` is accepted as
`3000000000 & "x"` (oracle suffix_long_amp_glued_concat_accepted), so a
&-suffixed literal is judged only when nothing that could start an operand follows
it.
"""

from __future__ import annotations

import math
import re
from collections.abc import Sequence

from ...conditional import ConditionalActivityTracker
from ...js_compat import JS_WHITESPACE, js_number, js_number_to_string, js_trim
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import date_literal_month, tokenize_cached
from ...parser.nodes import Span
from ..context import PushFn

_INTEGER_MAX = 32767
_LONG_MAX = 2147483647
_LONGLONG_MAX = 9223372036854775807
_SINGLE_MAX = 3.402823e38
_CURRENCY_MAX = 922337203685477.5807
_OPERAND_STARTS = frozenset(
    {
        TokenKind.IDENTIFIER,
        TokenKind.KEYWORD,
        TokenKind.INTEGER_LITERAL,
        TokenKind.FLOAT_LITERAL,
        TokenKind.STRING_LITERAL,
        TokenKind.DATE_LITERAL,
        TokenKind.BRACKETED_IDENTIFIER,
    }
)
_SUFFIX_ENDS = frozenset(
    {TokenKind.NEWLINE, TokenKind.COLON, TokenKind.COMMENT, TokenKind.OPERATOR, TokenKind.PUNCTUATION}
)

# `[0-9]` rather than `\d`, and the JavaScript whitespace set rather than `\s`:
# upstream's patterns read ASCII digits and ECMAScript whitespace only.
_RADIX_LITERAL_RE = re.compile(r"&([hHoO]?)([0-9A-Fa-f]*)([%&^]?)")
_DECIMAL_LITERAL_RE = re.compile(r"([0-9]+)([%&^]?)")
_FLOAT_SUFFIX_RE = re.compile(r"[!#@]$")
_D_EXPONENT_RE = re.compile(r"[dD]")
_PLAIN_CURRENCY_RE = re.compile(r"([0-9]+)(?:\.([0-9]*))?@")
_JS_SPACE = "[" + JS_WHITESPACE + "]"
# The time comes last: h:m[:s] with ':' or '.' between, or h and AM/PM.
_DATE_TIME_RE = re.compile(
    r"(?:^|" + _JS_SPACE + r")([0-9]+)" + _JS_SPACE + r"*[:.]" + _JS_SPACE + r"*([0-9]+)"
    r"(?:" + _JS_SPACE + r"*[:.]" + _JS_SPACE + r"*([0-9]+))?" + _JS_SPACE + r"*(?:[apAP][mM]?)?\Z"
    r"|(?:^|" + _JS_SPACE + r")([0-9]+)" + _JS_SPACE + r"*[apAP][mM]?\Z"
)
_DATE_PART_SPLIT_RE = re.compile(_JS_SPACE + r"*[/,-]" + _JS_SPACE + r"*|" + _JS_SPACE + r"+")
_DIGITS_RE = re.compile(r"[0-9]+")


def check_suffixed_literal_overflow(
    source: str, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    tokens = tokenize_cached(source)
    for index, tok in enumerate(tokens):
        span = Span(tok.start, tok.end)
        if tok.kind is TokenKind.FLOAT_LITERAL:
            if activity is not None and activity.is_inactive(span):
                continue
            _check_float(tok, tokens[index + 1] if index + 1 < len(tokens) else None, push)
            continue
        if tok.kind is TokenKind.DATE_LITERAL:
            if activity is not None and activity.is_inactive(span):
                continue
            problem = _date_literal_problem(tok.raw_text)
            if problem is not None:
                push(
                    "dateLiteralInvalid",
                    f"The date literal {tok.raw_text} {problem}. VBE rejects this at compile time "
                    "as a Syntax error.",
                    span,
                )
            continue
        if (
            tok.raw_text == "#"
            and _starts_date_literal(tokens, index)
            and not (activity is not None and activity.is_inactive(span))
        ):
            # `#2000#` and `#1/1/-5#` are no date literal, so the lexer read the
            # '#' alone and a Double after it. After `=`, an operator or `(`,
            # a '#' starts no file number (issue #190).
            end = index + 1
            while (
                end < len(tokens)
                and tokens[end].kind is not TokenKind.NEWLINE
                and not tokens[end].raw_text.endswith("#")
            ):
                end += 1
            last = tokens[end] if end < len(tokens) and tokens[end].kind is not TokenKind.NEWLINE else tok
            text = source[tok.start : last.end]
            push(
                "dateLiteralInvalid",
                f"{text} is no date literal: a date needs a month and a day, or a time. VBE rejects "
                "this at compile time as a Syntax error.",
                Span(tok.start, last.end),
            )
            continue
        if tok.kind is not TokenKind.INTEGER_LITERAL or (activity is not None and activity.is_inactive(span)):
            continue
        _check_integer(tok, tokens[index + 1] if index + 1 < len(tokens) else None, push)


def _check_integer(tok: VbaToken, following: VbaToken | None, push: PushFn) -> None:
    raw = tok.raw_text
    span = Span(tok.start, tok.end)

    def reject(message: str) -> None:
        push("suffixedLiteralOverflow", f"{message} VBE rejects this at compile time as a Syntax error.", span)

    radix = _RADIX_LITERAL_RE.fullmatch(raw)
    if radix is not None:
        letter, digits, suffix = radix.groups()
        if len(digits) == 0:
            reject(f"'{raw}' names a radix with no digits after it.")
            return
        # `&HFF#` and `&HFF!`: only a decimal number takes a Double or Single
        # suffix (issue #369, measured in Excel 16.0).
        if following is not None and following.start == tok.end and following.raw_text in ("#", "!"):
            kind = "Double" if following.raw_text == "#" else "Single"
            push(
                "suffixedLiteralOverflow",
                f"The literal '{raw}{following.raw_text}' gives a hex or octal number a {kind} suffix, "
                "which only a decimal number takes. VBE rejects this at compile time as a Syntax error.",
                Span(tok.start, following.end),
            )
            return
        try:
            value = int(digits, 16 if letter.lower() == "h" else 8)
        except ValueError:
            return
        # A hex or octal literal wraps: 16 bits with %, 32 with &, 64 with ^, and
        # without a suffix 16 then 32 bits as the digits need.
        if suffix == "%" and value > 0xFFFF:
            reject(
                f"The literal '{raw}' does not fit the Integer its '%' suffix asks for: at most four "
                "hex digits (&HFFFF)."
            )
        elif suffix == "&" and value > 0xFFFFFFFF:
            reject(
                f"The literal '{raw}' does not fit the Long its '&' suffix asks for: at most eight "
                "hex digits (&HFFFFFFFF)."
            )
        elif suffix == "^" and value > 0xFFFFFFFFFFFFFFFF:
            reject(f"The literal '{raw}' does not fit the LongLong its '^' suffix asks for.")
        elif suffix == "" and value > 0xFFFFFFFF:
            # &H100000000 and &O40000000000 are refused, 64-bit Office too
            # (issue #369, measured in Excel 16.0).
            reject(
                f"The literal '{raw}' is wider than 32 bits, the most a hex or octal literal holds "
                "without the '^' suffix."
            )
        return
    decimal = _DECIMAL_LITERAL_RE.fullmatch(raw)
    if decimal is None:
        return
    digits, suffix = decimal.groups()
    # The sign belongs to the token's value: -32768% is refused as well as 32768%
    # (oracle suffix_integer_pct_neg_compile), so only the magnitude counts.
    value = int(digits)
    if suffix == "%" and value > _INTEGER_MAX:
        reject(f"The literal '{raw}' is outside the Integer range -32768 to 32767 of its '%' type suffix.")
    elif suffix == "&" and value > _LONG_MAX:
        # `3000000000&"x"` reads as concatenation; only a `&` nothing follows, or
        # an operator follows, is the Long suffix.
        if following is None or following.kind in _SUFFIX_ENDS:
            if not (following is not None and following.kind in _OPERAND_STARTS):
                reject(
                    f"The literal '{raw}' is outside the Long range -2147483648 to 2147483647 of its "
                    "'&' type suffix."
                )
    elif suffix == "^" and value > _LONGLONG_MAX:
        reject(
            f"The literal '{raw}' is outside the LongLong range of its '^' type suffix (at most "
            "9223372036854775807)."
        )


def _check_float(tok: VbaToken, following: VbaToken | None, push: PushFn) -> None:
    raw = tok.raw_text
    span = Span(tok.start, tok.end)
    suffix_match = _FLOAT_SUFFIX_RE.search(raw)
    suffix = suffix_match.group(0) if suffix_match is not None else ""
    value = js_number(_FLOAT_SUFFIX_RE.sub("", _D_EXPONENT_RE.sub("E", raw)))
    if not math.isfinite(value):
        push(
            "floatLiteralOverflow",
            f"The literal '{raw}' is outside the Double range (about 1.8E+308). VBE rejects this at "
            "compile time as a Syntax error.",
            span,
        )
        return
    if suffix == "!" and abs(value) > _SINGLE_MAX:
        push(
            "floatLiteralOverflow",
            f"The literal '{raw}' is outside the Single range (about 3.402823E+38) of its '!' type "
            "suffix. VBE rejects this at compile time as a Syntax error.",
            span,
        )
        return
    if suffix == "@" and _currency_overflows(raw):
        push(
            "floatLiteralOverflow",
            f"The literal '{raw}' is outside the Currency range (at most 922337203685477.5807) of its "
            "'@' type suffix. VBE rejects this at compile time as a Syntax error.",
            span,
        )
        return
    # `1.5%` and `1E3%`: the Integer suffix on a number that is not an integer token.
    if (
        following is not None
        and following.kind is TokenKind.UNKNOWN
        and following.raw_text == "%"
        and following.start == tok.end
    ):
        push(
            "suffixedLiteralOverflow",
            f"The literal '{raw}%' puts the '%' Integer type suffix on a fractional or exponent "
            "literal, which has no Integer form. VBE rejects this at compile time as a Syntax error.",
            Span(tok.start, following.end),
        )


def _currency_overflows(raw: str) -> bool:
    """Whether a Currency literal exceeds 922337203685477.5807 in magnitude, compared
    in scaled integers: at 9.2E+14 a Double's spacing is 0.125, so .5807 and .5808
    would read as the same number as floats."""
    plain = _PLAIN_CURRENCY_RE.fullmatch(raw)
    if plain is None:
        value = js_number(re.sub(r"@$", "", _D_EXPONENT_RE.sub("E", raw)))
        return math.isfinite(value) and abs(value) > _CURRENCY_MAX
    fraction = (plain.group(2) or "").ljust(4, "0")
    if len(fraction) > 4:
        return False  # more than four decimals: rounding the VBE applies is not modelled
    return int(plain.group(1)) * 10000 + int(fraction) > _LONGLONG_MAX


def _days_in_month(year: int, month: int) -> int:
    """What `new Date(Date.UTC(year, month, 0)).getUTCDate()` gives: the length of a
    1-based month, with Date.UTC's reading of years 0 to 99 as 1900 to 1999."""
    if 0 <= year <= 99:
        year += 1900
    if month == 2:
        leap = (year % 4 == 0 and year % 100 != 0) or year % 400 == 0
        return 29 if leap else 28
    return 30 if month in (4, 6, 9, 11) else 31


def _starts_date_literal(tokens: Sequence[VbaToken], index: int) -> bool:
    """Whether the '#' at `index` stands where only a value can: after `=`, an
    operator or `(`."""
    if index - 1 < 0:
        return False
    before = tokens[index - 1]
    return before.raw_text == "(" or (
        before.kind is TokenKind.OPERATOR and before.raw_text != "#" and before.raw_text != ":="
    )


def _num(value: float) -> str:
    return js_number_to_string(value)


def _date_literal_problem(raw: str) -> str | None:
    """What is wrong with a `#...#` date literal, or None when it is one the VBE
    accepts or one this check does not judge (named months and other regional
    forms are left alone)."""
    body = js_trim(raw[1:-1])
    if len(body) == 0:
        return "is empty"
    # The time comes last: h:m[:s] with ':' or '.' between, or h and AM/PM.
    # `#1.2.2000#` is such a time, and 2000 is no second.
    time = _DATE_TIME_RE.search(body)
    if time is not None:
        hour_text = time.group(1) if time.group(1) is not None else time.group(4)
        hour = js_number(hour_text)
        minute = 0.0 if time.group(2) is None else js_number(time.group(2))
        seconds = 0.0 if time.group(3) is None else js_number(time.group(3))
        if hour > 23:
            return f"names hour {_num(hour)}; hours run 0 to 23"
        if minute > 59:
            return f"names minute {_num(minute)}; minutes run 0 to 59"
        if seconds > 59:
            return f"names second {_num(seconds)}; seconds run 0 to 59"
    date_part = js_trim(body[: time.start()] if time is not None else body)
    return None if len(date_part) == 0 else _date_part_problem(date_part)


def _date_part_problem(text: str) -> str | None:
    """The date of a date literal: two or three parts, numbers or a month name,
    between '/', '-', ',' or blanks (issue #190, each measured in Excel 16.0).
    Numbers read month/day/year, and #13/1/2000# as 13 January. A first number
    of three digits or more is a year: #2000/12/1# is December 1 and
    #2000/13/1# is refused. With a month name the numbers are day and year:
    `#Jan 1, 2000#`, `#1 Sept 2000#`."""
    parts = [part for part in _DATE_PART_SPLIT_RE.split(text) if len(part) > 0]
    if len(parts) < 2 or len(parts) > 3:
        return None
    named = next((i for i, part in enumerate(parts) if _DIGITS_RE.fullmatch(part) is None), -1)
    year_text: str | None
    if named >= 0:
        from_name = date_literal_month(parts[named])
        numbers = [part for i, part in enumerate(parts) if i != named]
        if not isinstance(from_name, int) or any(_DIGITS_RE.fullmatch(part) is None for part in numbers):
            return None
        month = float(from_name)
        day = js_number(numbers[0])
        year_text = numbers[1] if len(numbers) > 1 else None
        if len(numbers) == 1 and day > 31:
            return None  # `#Jan 2000#` is a month and a year
    elif len(parts[0]) >= 3:
        if len(parts) != 3:
            return None
        year_text = parts[0]
        month = js_number(parts[1])
        day = js_number(parts[2])
        if month > 12:
            return f"names month {_num(month)}, which no calendar has"
    else:
        month = js_number(parts[0])
        day = js_number(parts[1])
        year_text = parts[2] if len(parts) > 2 else None
        if month > 12:
            # The VBE reads #13/1/2000# as 13 January when the first number
            # cannot be a month and the second can.
            if day <= 12:
                month, day = day, month
            else:
                return f"names month {_num(month)}, which no calendar has"
    year = None if year_text is None else js_number(year_text)
    if year is not None and year > 9999:
        return "names a year past 9999"
    if month < 1:
        return "names month 0"
    # With no year, February may have 29 days.
    if year is None or year_text is None:
        full_year = 2000
    elif len(year_text) <= 2:
        full_year = int(2000 + year if year < 30 else 1900 + year)
    else:
        full_year = int(year)
    days_in_month = _days_in_month(full_year, int(month))
    if day < 1 or day > days_in_month:
        return f"names day {_num(day)} in a month of {days_in_month} days"
    return None
