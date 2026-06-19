"""M8: scalar-member-access (objectState.ts) + For Each loop types (controlFlow.ts)."""

from __future__ import annotations

from oracle_support import (
    accepted_cases,
    assert_oracle_behavior,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import analyze_module

_CODES = (
    "scalar-member-access",
    "for-each-control-variable-type",
    "for-each-source-type",
)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


# A user-defined-Type control variable is invalid, but detecting it needs the
# project/host type registry (resolveTypeName), which is deferred to M9. The
# known-scalar and array control-variable cases are fully covered.
_M9_TYPE_REGISTRY = frozenset({"for_each_udt_control_variable_compile"})


def test_scalar_member_access() -> None:
    assert "scalar-member-access" in _codes("Sub S()\n    Dim n As Long\n    n.Foo = 1\nEnd Sub")


def test_for_each_scalar_control_variable() -> None:
    src = "Sub S()\n    Dim i As Long\n    Dim c As Collection\n    For Each i In c\n    Next\nEnd Sub"
    assert "for-each-control-variable-type" in _codes(src)


def test_for_each_scalar_source() -> None:
    src = "Sub S()\n    Dim v As Variant\n    Dim n As Long\n    For Each v In n\n    Next\nEnd Sub"
    assert "for-each-source-type" in _codes(src)


def test_valid_for_each_silent() -> None:
    src = "Sub S()\n    Dim v As Variant\n    Dim c As Collection\n    For Each v In c\n    Next\nEnd Sub"
    assert not (_codes(src) & set(_CODES))


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code, skip_ids=_M9_TYPE_REGISTRY) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, _CODES)
        assert not spurious, f"{case.id}: scalar/for-each false positive {spurious}"
