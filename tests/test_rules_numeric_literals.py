"""M6: suffixed numeric-literal overflow (numericLiterals.ts parity)."""

from __future__ import annotations

from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

_CODE = "suffixed-literal-overflow"


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_integer_suffix_overflow() -> None:
    assert _CODE in _codes("x = 40000%")
    assert _CODE in _codes("x = 32768%")  # one past Integer max
    assert _CODE not in _codes("x = 32767%")  # the max is accepted
    # Only the % suffix; & is ambiguous with string concatenation.
    assert _CODE not in _codes('x = 3000000000&"y"')
    # Hex / octal / no-suffix literals never match.
    assert _CODE not in _codes("x = 40000")
    assert _CODE not in _codes("x = &HFFFF")


def _wrap(*lines: str) -> str:
    body = "\n".join(f"    {line}" for line in lines)
    return f"Option Explicit\nFunction Main() As Variant\n{body}\nEnd Function\n"


def _found(source: str, code: str) -> list[tuple[str, str]]:
    return [(source[d.span.start : d.span.end], d.message) for d in analyze_module(source) if d.code == code]


def test_unsuffixed_radix_literal_wider_than_32_bits_and_float_suffix_on_hex() -> None:
    # XLIDE issue #369 (upstream 2f49b93 sync), measured in Excel 16.0.
    source = _wrap("Dim v As Variant", "v = &H100000000", "v = &O40000000000", "v = &HFF#", "v = &HFF!", "Main = v")
    found = _found(source, _CODE)
    assert [text for text, _ in found] == ["&H100000000", "&O40000000000", "&HFF#", "&HFF!"]
    assert "wider than 32 bits" in found[0][1]
    assert "a hex or octal number a Double suffix" in found[2][1]
    assert "a hex or octal number a Single suffix" in found[3][1]
    assert _found(_wrap("Dim v As Variant", "v = &HFFFFFFFF", "Main = v"), _CODE) == []


def test_date_literal_month_names_and_forms() -> None:
    # XLIDE issue #190 (upstream literalForms.test.ts), measured in Excel 16.0.
    accepted = [
        "#Janu 1, 2000#", "#Januar 1, 2000#", "#Febr 1, 2000#", "#Sept 1, 2000#", "#Septembe 1, 2000#",
        "#Dece 1, 2000#", "#Jan. 1, 2000#", "#Sept. 1, 2000#", "#SEPT 1, 2000#", "#1 Sept 2000#",
        "#may 1, 2000#", "#2000/12/1#", "#2000-02-28#",
    ]
    source = _wrap("Dim v As Variant", *(f"v = {literal}" for literal in accepted), "Main = v")
    assert _found(source, "date-literal-invalid") == []
    refused = [
        ("#2000/13/1#", "month 13"),
        ("#2000-02-30#", "day 30"),
        ("#1.2.2000#", "second 2000"),
        ("#Febru 30, 2000#", "day 30"),
        ("#2000#", "is no date literal"),
        ("#1/1/-5#", "is no date literal"),
    ]
    for literal, message in refused:
        found = _found(_wrap("Dim v As Variant", f"v = {literal}", "Main = v"), "date-literal-invalid")
        assert [text for text, _ in found] == [literal], literal
        assert message in found[0][1], literal
    file_numbers = _wrap(
        "Dim s As String", 'Open "x.txt" For Input As #1', "Line Input #1, s", "s = Input(5, #1)", "Close #1", "Main = s"
    )
    assert _found(file_numbers, "date-literal-invalid") == []


def test_oracle_asserted_cases() -> None:
    if asserted_cases(_CODE):
        assert assert_oracle_behavior(_CODE) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        assert _CODE not in case_codes(case), f"{case.id}: {_CODE} false positive"
