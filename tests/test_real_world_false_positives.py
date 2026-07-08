"""Regression suite from analyzing the real-world ModernJsonInVBA library
(XLIDE v2.5.11 parity): five families of valid VBA that were wrongly flagged
as errors, each with negative controls proving the rules still catch the
neighboring genuine errors."""

from __future__ import annotations

import re

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module

_NAME_RE = re.compile(r"'([^']+)'")


def _codes(body: str) -> list[str]:
    src = f"Option Explicit\n{body}"
    # known_identifiers activates the undeclared-variable rule (it is inert
    # without a known-identifier universe, mirroring the project pipeline).
    diags = analyze_module(src, AnalyzeModuleOptions(known_identifiers=set()))
    out: list[str] = []
    for d in diags:
        m = _NAME_RE.search(d.message)
        out.append(f"{d.code}:{m.group(1) if m else ''}")
    return out


# -- Open-statement Access clause -------------------------------------------


def test_open_access_clause_not_flagged() -> None:
    for clause in ("Access Read", "Access Write", "Access Read Write"):
        body = (
            "Sub T()\n    Dim f As Long\n    Dim p As String\n    f = FreeFile\n"
            f"    Open p For Binary {clause} As #f\n    Close #f\nEnd Sub\n"
        )
        assert "undeclared-variable:Access" not in _codes(body), clause


def test_access_still_flagged_outside_clause_position() -> None:
    body = "Sub T()\n    Dim x As Long\n    x = Access\nEnd Sub\n"
    assert "undeclared-variable:Access" in _codes(body)


def test_undeclared_pathname_inside_open_still_flagged() -> None:
    body = (
        "Sub T()\n    Dim f As Long\n    f = FreeFile\n"
        "    Open missingPath For Binary Access Read As #f\nEnd Sub\n"
    )
    assert "undeclared-variable:missingPath" in _codes(body)


# -- Byte array assigned to a String scalar ----------------------------------


def test_byte_array_to_string_not_flagged() -> None:
    body = "Sub T()\n    Dim o() As Byte\n    Dim s As String\n    ReDim o(0 To 3)\n    s = o\nEnd Sub\n"
    assert not any(c.startswith("array-assignment-to-scalar") for c in _codes(body))


def test_byte_array_to_string_function_return_not_flagged() -> None:
    body = "Function F() As String\n    Dim o() As Byte\n    ReDim o(0 To 3)\n    F = o\nEnd Function\n"
    assert not any(c.startswith("array-assignment-to-scalar") for c in _codes(body))


def test_non_byte_string_pairs_still_flagged() -> None:
    long_to_string = (
        "Sub T()\n    Dim o() As Long\n    Dim s As String\n    ReDim o(0 To 3)\n    s = o\nEnd Sub\n"
    )
    byte_to_long = (
        "Sub T()\n    Dim o() As Byte\n    Dim n As Long\n    ReDim o(0 To 3)\n    n = o\nEnd Sub\n"
    )
    assert any(c.startswith("array-assignment-to-scalar") for c in _codes(long_to_string))
    assert any(c.startswith("array-assignment-to-scalar") for c in _codes(byte_to_long))


# -- ReDim inside a single-line If -------------------------------------------


def test_single_line_if_redim_target_not_flagged() -> None:
    body = (
        "Sub T()\n    Dim headers() As String\n    Dim colCount As Long\n    colCount = 3\n"
        "    If colCount > 0 Then ReDim headers(1 To colCount)\nEnd Sub\n"
    )
    assert not any(c.startswith("unallocated-dynamic-array-access") for c in _codes(body))


def test_single_line_if_redim_both_arms_not_flagged() -> None:
    body = (
        "Sub T()\n    Dim a() As String\n    Dim c As Long\n    c = 2\n"
        "    If c > 0 Then ReDim a(1 To c) Else ReDim a(1 To 5)\nEnd Sub\n"
    )
    assert not any(c.startswith("unallocated-dynamic-array-access") for c in _codes(body))


def test_genuine_unallocated_access_still_flagged() -> None:
    no_redim = "Sub T()\n    Dim a() As String\n    Dim x As String\n    x = a(1)\nEnd Sub\n"
    if_access = (
        "Sub T()\n    Dim a() As String\n    Dim x As String\n    Dim c As Long\n    c = 1\n"
        "    If c > 0 Then x = a(1)\nEnd Sub\n"
    )
    assert any(c.startswith("unallocated-dynamic-array-access") for c in _codes(no_redim))
    assert any(c.startswith("unallocated-dynamic-array-access") for c in _codes(if_access))


# -- ReDim ... As Type clause -------------------------------------------------


def test_redim_as_type_clause_not_flagged() -> None:
    for type_name in ("Collection", "Object", "String"):
        body = (
            f"Sub T()\n    Dim rowsByIdx() As {type_name}\n    Dim cap As Long\n    cap = 16\n"
            f"    ReDim rowsByIdx(1 To cap) As {type_name}\nEnd Sub\n"
        )
        assert f"undeclared-variable:{type_name}" not in _codes(body), type_name


def test_redim_preserve_as_type_not_flagged() -> None:
    body = (
        "Sub T()\n    Dim rowObjs() As Collection\n    ReDim rowObjs(1 To 8) As Collection\n"
        "    ReDim Preserve rowObjs(1 To 16) As Collection\nEnd Sub\n"
    )
    assert "undeclared-variable:Collection" not in _codes(body)


def test_redim_bounds_identifier_still_flagged() -> None:
    body = "Sub T()\n    Dim a() As Long\n    ReDim a(1 To undeclaredVar) As Long\nEnd Sub\n"
    assert "undeclared-variable:undeclaredVar" in _codes(body)


def test_bare_collection_use_still_flagged() -> None:
    body = "Sub T()\n    Dim x As Variant\n    x = Collection\nEnd Sub\n"
    assert "undeclared-variable:Collection" in _codes(body)


# -- parenless call with a parenthesized first argument -----------------------

_ASSERT = "Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)\nEnd Sub\n"


def _arity_codes(call: str) -> list[str]:
    body = f"{_ASSERT}Sub T()\n    Dim v As Long\n    v = 1\n    {call}\nEnd Sub\n"
    return [c for c in _codes(body) if c.startswith("argument-count")]


def test_parenthesized_first_argument_counts_both() -> None:
    for call in (
        'AssertTrue (v <> 0), "message"',
        'AssertTrue(v <> 0), "message"',
        "AssertTrue (v), (v = 1)",
        'AssertTrue (v) > 0, "message"',
        'AssertTrue v <> 0, "message"',
    ):
        assert _arity_codes(call) == [], call


def test_genuinely_one_argument_call_still_flagged() -> None:
    assert len(_arity_codes("AssertTrue (v <> 0)")) > 0


def test_three_argument_call_reports_correct_count() -> None:
    body = f'{_ASSERT}Sub T()\n    Dim v As Long\n    v = 1\n    AssertTrue (v <> 0), "m", 1\nEnd Sub\n'
    src = f"Option Explicit\n{body}"
    hits = [d for d in analyze_module(src) if d.code == "argument-count"]
    assert hits and "got 3" in hits[0].message
