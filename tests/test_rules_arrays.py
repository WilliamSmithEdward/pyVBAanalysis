"""M7: array-family rules (arrays.ts parity)."""

from __future__ import annotations

from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

# Every array-family code emitted so far (drives the no-false-positive sweep).
_ARRAY_CODES = (
    "redim-impossible-bounds",
    "array-declaration-impossible-bounds",
    "too-many-array-dimensions",
    "array-subscript-out-of-bounds",
    "redim-preserve-dimension-change",
    "scalar-redim",
    "fixed-array-redim",
    "invalid-erase-target",
    "erase-requires-array",
    "unallocated-dynamic-array-access",
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


def test_array_subscript_out_of_bounds() -> None:
    code = "array-subscript-out-of-bounds"
    fixed = "Dim a(1 To 10) As Long\n"
    # Above the upper bound and below an explicit lower bound both fire.
    assert code in _codes(f"Sub S\n    {fixed}    a(11) = 1\nEnd Sub")
    assert code in _codes(f"Sub S\n    {fixed}    a(0) = 1\nEnd Sub")
    # In-bounds subscripts stay silent.
    assert code not in _codes(f"Sub S\n    {fixed}    a(1) = 1\nEnd Sub")
    assert code not in _codes(f"Sub S\n    {fixed}    a(10) = 1\nEnd Sub")
    # Single-bound Dim b(5): the low region is Option-Base-dependent (silent);
    # only above-upper or negative fires.
    single = "Dim b(5) As Long\n"
    assert code not in _codes(f"Sub S\n    {single}    b(0) = 1\nEnd Sub")
    assert code in _codes(f"Sub S\n    {single}    b(6) = 1\nEnd Sub")
    assert code in _codes(f"Sub S\n    {single}    b(-1) = 1\nEnd Sub")
    # Variable subscripts are not statically provable; ReDim'd and member-access
    # arrays are excluded.
    assert code not in _codes(f"Sub S\n    {fixed}    a(i) = 1\nEnd Sub")
    assert code not in _codes("Sub S\n    Dim d() As Long\n    ReDim d(1 To 3)\n    d(99) = 1\nEnd Sub")
    assert code not in _codes(f"Sub S\n    {fixed}    obj.a(11) = 1\nEnd Sub")


def test_redim_preserve_dimensions() -> None:
    code = "redim-preserve-dimension-change"
    arr = "Dim grid() As Long\n"
    base = "    ReDim grid(1 To 2, 1 To 2)\n"
    # Resizing only the final dimension's upper bound is allowed.
    assert code not in _codes(f"Sub S\n    {arr}{base}    ReDim Preserve grid(1 To 2, 1 To 3)\nEnd Sub")
    # Changing a non-final dimension is flagged.
    assert code in _codes(f"Sub S\n    {arr}{base}    ReDim Preserve grid(1 To 3, 1 To 2)\nEnd Sub")
    # Changing the final dimension's lower bound is flagged.
    assert code in _codes(f"Sub S\n    {arr}{base}    ReDim Preserve grid(1 To 2, 0 To 3)\nEnd Sub")
    # Changing the dimension count is flagged.
    assert code in _codes(f"Sub S\n    {arr}{base}    ReDim Preserve grid(1 To 2)\nEnd Sub")
    # A plain ReDim (not Preserve) is never flagged.
    assert code not in _codes(f"Sub S\n    {arr}{base}    ReDim grid(1 To 5, 1 To 5)\nEnd Sub")
    # Non-literal bounds yield no comparable key -> quiet.
    assert code not in _codes(
        f"Sub S\n    {arr}    ReDim grid(1 To n)\n    ReDim Preserve grid(1 To m)\nEnd Sub"
    )


def test_invalid_redim_targets() -> None:
    # A scalar variable and a fixed-size array cannot be ReDim'd.
    assert "scalar-redim" in _codes("Sub S\n    Dim x As Long\n    ReDim x(1 To 10)\nEnd Sub")
    assert "fixed-array-redim" in _codes("Sub S\n    Dim a(1 To 3) As Long\n    ReDim a(1 To 9)\nEnd Sub")
    # A dynamic array is the legal ReDim target.
    dyn = "Sub S\n    Dim d() As Long\n    ReDim d(1 To 10)\nEnd Sub"
    assert "scalar-redim" not in _codes(dyn)
    assert "fixed-array-redim" not in _codes(dyn)
    # Variant (explicit and implicit) scalars are legal implicit ReDim targets.
    assert "scalar-redim" not in _codes("Sub S\n    Dim v As Variant\n    ReDim v(1 To 10)\nEnd Sub")
    assert "scalar-redim" not in _codes("Sub S\n    Dim v\n    ReDim v(1 To 10)\nEnd Sub")
    # An undeclared / unresolved name stays quiet.
    assert "scalar-redim" not in _codes("Sub S\n    ReDim unknown(1 To 10)\nEnd Sub")


def test_erase_targets() -> None:
    # Erasing a scalar (Object or Long) is a compile error; Variant is allowed.
    assert "erase-requires-array" in _codes("Sub S\n    Dim obj As Object\n    Erase obj\nEnd Sub")
    assert "erase-requires-array" in _codes("Sub S\n    Dim n As Long\n    Erase n\nEnd Sub")
    assert "erase-requires-array" not in _codes("Sub S\n    Dim v As Variant\n    Erase v\nEnd Sub")
    # Erasing an array (fixed or dynamic) is fine.
    assert "erase-requires-array" not in _codes("Sub S\n    Dim a(1 To 3) As Long\n    Erase a\nEnd Sub")
    assert "erase-requires-array" not in _codes("Sub S\n    Dim d() As Long\n    Erase d\nEnd Sub")
    # An expression target is not a variable/array name.
    assert "invalid-erase-target" in _codes("Sub S\n    Erase 1 + 2\nEnd Sub")
    assert "invalid-erase-target" not in _codes("Sub S\n    Dim a(1 To 3) As Long\n    Erase a\nEnd Sub")


def test_unallocated_dynamic_array_access() -> None:
    code = "unallocated-dynamic-array-access"
    arr = "Dim values() As Long\n"
    # Indexed access and LBound/UBound on an unallocated dynamic array fire.
    assert code in _codes(f"Sub S\n    {arr}    Debug.Print values(0)\nEnd Sub")
    assert code in _codes(f"Sub S\n    {arr}    Debug.Print LBound(values)\nEnd Sub")
    assert code in _codes(f"Sub S\n    {arr}    Debug.Print UBound(values)\nEnd Sub")
    # After ReDim it is allocated -> silent.
    assert code not in _codes(f"Sub S\n    {arr}    ReDim values(1 To 3)\n    Debug.Print values(0)\nEnd Sub")
    # Erase resets to unallocated -> fires again.
    assert code in _codes(
        f"Sub S\n    {arr}    ReDim values(1 To 3)\n    Erase values\n    Debug.Print values(0)\nEnd Sub"
    )
    # Passing to a helper makes the state unknown -> silent.
    assert code not in _codes(f"Sub S\n    {arr}    Init values\n    Debug.Print values(0)\nEnd Sub")
    # Branch merge: ReDim'd on every arm before access -> silent.
    assert code not in _codes(
        f"Sub S\n    {arr}    If c Then\n        ReDim values(1 To 2)\n"
        "    Else\n        ReDim values(1 To 3)\n    End If\n    Debug.Print values(0)\nEnd Sub"
    )
    # Static / fixed-size / member-access arrays are not tracked.
    assert code not in _codes("Sub S\n    Static s() As Long\n    Debug.Print s(0)\nEnd Sub")
    assert code not in _codes("Sub S\n    Dim f(1 To 3) As Long\n    Debug.Print f(0)\nEnd Sub")
    assert code not in _codes(f"Sub S\n    {arr}    Debug.Print obj.values(0)\nEnd Sub")


def test_oracle_asserted_cases() -> None:
    for code in (
        "redim-impossible-bounds",
        "too-many-array-dimensions",
        "array-subscript-out-of-bounds",
        "redim-preserve-dimension-change",
        "scalar-redim",
        "erase-requires-array",
        "unallocated-dynamic-array-access",
    ):
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_array_false_positives_on_accepted_cases() -> None:
    arr = set(_ARRAY_CODES)
    for case in accepted_cases():
        spurious = case_codes(case) & arr
        assert not spurious, f"{case.id}: array false positive {spurious}"
