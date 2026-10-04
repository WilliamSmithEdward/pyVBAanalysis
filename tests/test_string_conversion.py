"""stringConversion (XLIDE issues #188, #239, #262, #336, #444, #504, #703).

Ported from upstream's tests/diagnostics/stringConversion.test.ts (the pure
function part); the value cases come from the verdicts its docstrings record as
measured in Excel 16.0.
"""

from __future__ import annotations

import pytest

from pyvbaanalysis.diagnostics.string_conversion import (
    NumericStringVerdict,
    is_invalid_boolean_string,
    is_invalid_date_string,
    is_invalid_time_string,
    numeric_string_readings,
    numeric_string_verdict,
    val_prefix_value,
)


@pytest.mark.parametrize(
    ("text", "value"),
    [
        ("&HFF", 255),
        ("&hff", 255),
        ("&HFFFF", -1),
        ("&H80000000", -2147483648),
        ("&17", 15),
        (" 5 ", 5),
        ("(5)", -5),
        ("5-", -5),
        ("1e3", 1000),
    ],
)
def test_numeric_string_verdict_reads_the_value_every_locale_agrees_on(text: str, value: float) -> None:
    assert numeric_string_verdict(text) == NumericStringVerdict("number", value)


@pytest.mark.parametrize("text", ["$5", "5 $", "1,000", "2.5", "0.0", "+5", "1 000", "5x", "5 kr."])
def test_numeric_string_verdict_reads_what_some_locale_converts(text: str) -> None:
    assert numeric_string_verdict(text).kind == "number"


def test_separators_and_currency_claim_no_value() -> None:
    assert numeric_string_verdict("1,000") == NumericStringVerdict("number")
    assert numeric_string_verdict("$5") == NumericStringVerdict("number")


@pytest.mark.parametrize(
    "text", ["", "   ", "abc", ".", "$", "5%", "-&H10", "4x2", "1/2/2020", "12:30", "1,2.3,4.5"]
)
def test_numeric_string_verdict_refuses_what_no_locale_converts(text: str) -> None:
    assert numeric_string_verdict(text).kind == "invalid"


def test_a_no_break_space_groups_digits() -> None:
    assert numeric_string_verdict("1" + chr(0xA0) + "000").kind == "number"


@pytest.mark.parametrize("text", ["True", "true", "TRUE", "False", "5", "0.0", "$5", "(5)", "&HFF", " 1 "])
def test_boolean_strings_that_convert(text: str) -> None:
    assert is_invalid_boolean_string(text) is False


@pytest.mark.parametrize("text", [" True ", "yes", "", "4x2", "5%", "fal" + chr(0x17F) + "e"])
def test_boolean_strings_that_raise(text: str) -> None:
    assert is_invalid_boolean_string(text) is True


@pytest.mark.parametrize("text", ["5", "1/2/2020", "Jan 1", "12:30", "$5", "2020-02-29", "2958465"])
def test_date_strings_that_convert(text: str) -> None:
    assert is_invalid_date_string(text) is False


@pytest.mark.parametrize(
    "text", ["5%", ".", "", "True", "-&H10", "May", "Monday", "abc", "2020-02-30", "1/1/10000", "2958466"]
)
def test_date_strings_that_raise(text: str) -> None:
    assert is_invalid_date_string(text) is True


@pytest.mark.parametrize(("text", "invalid"), [("25:00", True), ("10:60", True), ("10:00:60", True),
                                               ("1:2:3:4", True), ("23:59:59", False), ("13:00 PM", False)])
def test_time_strings(text: str, invalid: bool) -> None:
    assert is_invalid_time_string(text) is invalid


@pytest.mark.parametrize(
    ("text", "value"),
    [("abc", 0), ("0,5", 0), ("1,000", 1), ("1 2 3", 123), (" -1.5e2x", -150), ("2d1", 20), ("&H7FFF", 32767),
     ("&o17", 15), ("-0", 0)],
)
def test_val_prefix_value(text: str, value: float) -> None:
    assert val_prefix_value(text) == value


def test_val_prefix_value_leaves_a_wide_radix_string_unread() -> None:
    assert val_prefix_value("&HFFFF") is None
    assert val_prefix_value("&Z") is None


def test_numeric_string_readings_read_both_decimal_points() -> None:
    assert numeric_string_readings("3.5") == (3.5, 35)
    assert numeric_string_readings("-1,5") == (-15, -1.5)
    assert numeric_string_readings("1.2.3") is None
    assert numeric_string_readings("1x") is None
