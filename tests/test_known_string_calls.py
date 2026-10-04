"""knownStringCalls.ts parity: InStr, InStrRev, Len, Asc, date parts and DateDiff
of known values folded to numbers (XLIDE issues #201, #509, #510, #559), and the
string folder. Expected values follow upstream's JavaScript semantics: UTF-16
lengths, Date arithmetic that carries out-of-range parts."""

from __future__ import annotations

from collections.abc import Sequence

from pyvbaanalysis.constants.integer_constant_expression import evaluate_integer_constant_expression
from pyvbaanalysis.diagnostics.known_string_calls import (
    KnownStringCallContext,
    StringFoldContext,
    UtcDate,
    date_utc,
    fold_known_string_calls,
    fold_string_expression,
    known_date,
    module_compare,
    parse_date_literal,
)
from pyvbaanalysis.diagnostics.walker import raw_expression_tokens
from pyvbaanalysis.lexer.token_kinds import VbaToken

_NO_CONSTANTS: dict[str, int | None] = {}


def _integer(text: str) -> int | None:
    return evaluate_integer_constant_expression(text, _NO_CONSTANTS)


def _integer_tokens(toks: Sequence[VbaToken]) -> int | None:
    return _integer(" ".join(tok.raw_text for tok in toks))


def _ctx(compare: str = "binary", known: dict[str, str] | None = None) -> KnownStringCallContext:
    return KnownStringCallContext(
        known_strings=known or {},
        integer_value=_integer,
        shadowed=lambda _name: False,
        compare=compare,  # type: ignore[arg-type]
    )


def _fold(text: str, ctx: KnownStringCallContext | None = None) -> str | None:
    return fold_known_string_calls(raw_expression_tokens(text), ctx or _ctx())


def _string(text: str, known: dict[str, str] | None = None) -> str | None:
    values = known or {}
    return fold_string_expression(
        raw_expression_tokens(text),
        StringFoldContext(
            name_value=lambda tok: values.get(tok.raw_text.lower()),
            integer_value=_integer_tokens,
        ),
    )


def _ymd(date: UtcDate | None) -> tuple[float, float, float] | None:
    if date is None:
        return None
    return (date.get_utc_full_year(), date.get_utc_month() + 1, date.get_utc_date())


def test_module_compare_reads_the_option_line() -> None:
    assert module_compare("Option Explicit\nSub A()\nEnd Sub") == "binary"
    assert module_compare("Option Explicit\r\n  option compare text\r\n") == "text"
    assert module_compare("Option Compare Database\n") == "database"
    # A carriage return alone starts a line too, as JavaScript's m flag reads it.
    assert module_compare("Option Explicit\rOption Compare Text") == "text"
    assert module_compare("' Option Compare Text\n") == "binary"


def test_instr_and_len_fold_into_the_arithmetic_around_them() -> None:
    assert _fold('InStr("abc", "c") - 1') == "3 - 1"
    assert _fold('Left$(s, InStr(s, " ") - 1)', _ctx(known={"s": "ab"})) == "Left $ ( s , 0 - 1 )"
    assert _fold('VBA.Len("abc") + 1') == "3 + 1"
    assert _fold('InStr(5, "abc", "")') == "5"
    assert _fold('InStr("", "")') == "0"
    assert _fold('InStrRev("abcabc", "b")') == "5"
    assert _fold('InStrRev("abcabc", "b", 4)') == "2"
    assert _fold('InStrRev("abc", "", -1)') == "3"
    assert _fold("x + 1") is None


def test_compare_modes() -> None:
    assert _fold('InStr(1, "ABC", "b", vbTextCompare)') == "2"
    assert _fold('InStr(1, "ABC", "b", 0)') == "0"
    assert _fold('InStr("ABC", "b")', _ctx("text")) == "2"
    # Under Option Compare Database a call is decided only where both agree.
    assert _fold('InStr("ABC", "b")', _ctx("database")) is None
    assert _fold('InStr("abc", "b")', _ctx("database")) == "2"


def test_len_and_asc_count_utf16_code_units() -> None:
    face = chr(0x1F600)
    assert _fold(f'Len("{face}")') == "2"
    assert _fold(f'AscW("{face}")') == str(0xD83D)
    assert _fold('Asc("A")') == "65"
    assert _fold(f'Asc("{chr(0xE9)}")') is None


def test_date_literals_carry_like_javascript_dates() -> None:
    assert _ymd(parse_date_literal("#2/30/2020#")) == (2020, 3, 1)
    assert _ymd(parse_date_literal("#1/1/100#")) == (100, 1, 1)
    assert parse_date_literal("#1/1/99#") is None
    assert parse_date_literal("#13/1/2020#") is None
    afternoon = parse_date_literal("#1/1/2000 1:30 PM#")
    assert afternoon is not None
    assert afternoon.get_time() == date_utc(2000, 0, 1, 13, 30)


def test_known_dates() -> None:
    def date(text: str) -> UtcDate | None:
        return known_date(raw_expression_tokens(text), _integer_tokens)

    assert _ymd(date("DateSerial(2020, 13, 1)")) == (2021, 1, 1)
    assert _ymd(date("DateSerial(2020, 3, 0)")) == (2020, 2, 29)
    assert date("DateSerial(99, 1, 1)") is None
    assert date("DateSerial(9999, 12, 32)") is None
    assert _ymd(date("CDate(-10000)")) == (1872, 8, 13)
    assert _ymd(date('DateValue("9999-12-31")')) == (9999, 12, 31)
    assert _ymd(date('CDate("31/12/2020")')) == (2020, 12, 31)
    assert date('CDate("1/2/2020")') is None
    assert date('DateValue("2020-02-30")') is None
    assert _ymd(date("((#5/6/2007#))")) == (2007, 5, 6)


def test_date_parts_and_date_diff() -> None:
    assert _fold("Year(#5/6/2007#) + Month(#5/6/2007#) + Day(#5/6/2007#)") == "2007 + 5 + 6"
    assert _fold('DateDiff("d", #1/2/2000#, #1/1/2000#)') == "(-1)"
    assert _fold('DateDiff("m", #12/31/1999#, #1/1/2000#)') == "1"
    assert _fold('DateDiff("yyyy", #12/31/1999#, #1/1/2000#)') == "1"
    assert _fold('DateDiff("q", #3/31/2000#, #4/1/2000#)') == "1"
    assert _fold('DateDiff("w", #1/1/2000#, #1/14/2000#)') == "1"
    assert _fold('DateDiff("h", #1/1/2000#, #1/2/2000#)') == "24"
    assert _fold('DateDiff("ww", #1/1/2000#, #1/14/2000#)') is None


def test_string_folding() -> None:
    assert _string('Left("abcdef", 3) & "x"') == "abcx"
    assert _string('Mid$("abcdef", 2, 2)') == "bc"
    assert _string('Mid("abcdef", 4)') == "def"
    assert _string('Right("abc", 5)') == "abc"
    assert _string('Replace("a-b", "-", "+")') == "a+b"
    assert _string('Replace("a-b", "a", "+")') is None
    assert _string('String(3, "xy")') == "xxx"
    assert _string('Space(2) & "|"') == "  |"
    assert _string('Trim("  a  ") & LTrim(" b") & RTrim("c ")') == "abc"
    assert _string('UCase(s) & StrReverse("ab")', {"s": "q"}) == "Qba"
    assert _string(f'UCase("{chr(0xE9)}")') is None
    assert _string('Left("abc", -1)') is None
    assert _string("s & t", {"s": "a"}) is None
