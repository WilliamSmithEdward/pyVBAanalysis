"""M7: array-family rules (arrays.ts parity)."""

from __future__ import annotations

from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

# Every array-family code emitted so far (drives the no-false-positive sweep).
_ARRAY_CODES = (
    "redim-impossible-bounds",
    "array-declaration-impossible-bounds",
    "too-many-array-dimensions",
)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_redim_impossible_bounds() -> None:
    code = "redim-impossible-bounds"
    arr = "Dim values() As Long\n"
    # lower > upper on an explicit literal range fires.
    assert code in _codes(f"Sub S\n    {arr}    ReDim values(10 To 1)\nEnd Sub")
    # equal and increasing bounds are accepted.
    assert code not in _codes(f"Sub S\n    {arr}    ReDim values(1 To 1)\nEnd Sub")
    assert code not in _codes(f"Sub S\n    {arr}    ReDim values(1 To 10)\nEnd Sub")
    # Upper-only and variable bounds are not literal ranges -> quiet.
    assert code not in _codes(f"Sub S\n    {arr}    ReDim values(5)\nEnd Sub")
    assert code not in _codes(f"Sub S\n    {arr}    ReDim values(1 To n)\nEnd Sub")
    # Signed-literal folding still detects an impossible range.
    assert code in _codes(f"Sub S\n    {arr}    ReDim values(1 To -1)\nEnd Sub")
    # Multi-dimension: only the offending dimension fires.
    assert code in _codes("Sub S\n    Dim g() As Long\n    ReDim g(1 To 2, 10 To 1)\nEnd Sub")
    # ReDim Preserve is bound-checked too.
    assert code in _codes(f"Sub S\n    {arr}    ReDim Preserve values(10 To 1)\nEnd Sub")


def test_scalar_redim_target_suppresses_bound_check() -> None:
    # A scalar ReDim target is an invalidRedimTargets compile error, not this
    # runtime one; redim-impossible-bounds must not also fire on it.
    assert "redim-impossible-bounds" not in _codes(
        "Sub S\n    Dim x As Long\n    ReDim x(10 To 1)\nEnd Sub"
    )


def test_array_declaration_bounds() -> None:
    code = "array-declaration-impossible-bounds"
    assert code in _codes("Sub S\n    Dim a(10 To 1) As Long\nEnd Sub")
    assert code not in _codes("Sub S\n    Dim a(1 To 10) As Long\nEnd Sub")
    assert code not in _codes("Sub S\n    Dim a(5) As Long\nEnd Sub")
    assert code not in _codes("Sub S\n    Dim a(1 To n) As Long\nEnd Sub")
    # Module-level declarations are checked too.
    assert code in _codes("Dim m(3 To 2) As Long")


def test_too_many_array_dimensions() -> None:
    code = "too-many-array-dimensions"
    over = ", ".join("1" for _ in range(61))
    ok = ", ".join("1" for _ in range(60))
    assert code in _codes(f"Sub S\n    Dim a({over}) As Long\nEnd Sub")
    assert code not in _codes(f"Sub S\n    Dim a({ok}) As Long\nEnd Sub")


def test_oracle_asserted_cases() -> None:
    for code in ("redim-impossible-bounds", "too-many-array-dimensions"):
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_array_false_positives_on_accepted_cases() -> None:
    arr = set(_ARRAY_CODES)
    for case in accepted_cases():
        spurious = case_codes(case) & arr
        assert not spurious, f"{case.id}: array false positive {spurious}"
