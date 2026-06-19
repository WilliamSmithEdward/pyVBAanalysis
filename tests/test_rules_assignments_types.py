"""M8: assignment / Set / missing-return type rules (assignments.ts parity)."""

from __future__ import annotations

from oracle_support import (
    accepted_cases,
    assert_oracle_behavior,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import analyze_module

# Runtime-error-kind and compile-error-kind codes emitted by these rules.
_CODES = (
    "assignment-type-mismatch",
    "string-arithmetic-coercion",
    "array-assignment-to-scalar",
    "set-required",
    "set-requires-object",
    "missing-return-assignment",
)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


# Member-access assignment type checking (obj.Property = value) needs the host
# member-completion surface (projectClassMembers + resolveExactMemberCompletion),
# which is deferred to M9. The bare-assignment cases are fully covered.
_M9_MEMBER_ASSIGNMENT = frozenset(
    {
        "class_property_integer_nonnumeric_string_runtime",
        "class_public_field_integer_nonnumeric_string_runtime",
    }
)


def test_assignment_type_mismatch() -> None:
    assert "assignment-type-mismatch" in _codes('Sub S()\n    Dim n As Long\n    n = "blah"\nEnd Sub')


def test_string_arithmetic_in_assignment() -> None:
    assert "string-arithmetic-coercion" in _codes(
        'Sub S()\n    Dim n As Long\n    n = 1 + "abc"\nEnd Sub'
    )


def test_set_requires_object() -> None:
    assert "set-requires-object" in _codes("Sub S()\n    Dim n As Long\n    Set n = Nothing\nEnd Sub")


def test_object_assignment_requires_set() -> None:
    assert "set-required" in _codes("Sub S()\n    Dim o As Object\n    o = 5\nEnd Sub")


def test_array_assignment_to_scalar() -> None:
    src = "Sub S()\n    Dim a(3) As Long\n    Dim n As Long\n    n = a\nEnd Sub"
    assert "array-assignment-to-scalar" in _codes(src)


def test_missing_return_assignment() -> None:
    assert "missing-return-assignment" in _codes("Function F()\nEnd Function")
    # A function that assigns its return name is silent.
    assert "missing-return-assignment" not in _codes("Function F()\n    F = 1\nEnd Function")


def test_compatible_assignments_silent() -> None:
    src = "Sub S()\n    Dim n As Long\n    n = 5\nEnd Sub"
    assert not (_codes(src) & set(_CODES))


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code, skip_ids=_M9_MEMBER_ASSIGNMENT) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, _CODES)
        assert not spurious, f"{case.id}: assignment-type false positive {spurious}"
