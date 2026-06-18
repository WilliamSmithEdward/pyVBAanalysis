"""M6: duplicate-declaration rule family (duplicates.ts parity)."""

from __future__ import annotations

import pytest
from oracle_support import accepted_cases, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

_DUP_CODES = (
    "duplicate-procedure",
    "duplicate-declaration",
    "duplicate-module-variable",
    "duplicate-enum-member",
    "duplicate-type-field",
)


def _codes(source: str) -> list[str]:
    return [d.code for d in analyze_module(source)]


def test_duplicate_procedure() -> None:
    assert "duplicate-procedure" in _codes("Sub Foo\nEnd Sub\nSub Foo\nEnd Sub")
    # A Get/Let/Set accessor trio is valid; only repeated accessors collide.
    assert "duplicate-procedure" not in _codes(
        "Property Get X() As Long\nEnd Property\nProperty Let X(v As Long)\nEnd Property"
    )
    assert "duplicate-procedure" in _codes(
        "Property Get X() As Long\nEnd Property\nProperty Get X() As Long\nEnd Property"
    )
    # A Sub colliding with a Property accessor is ambiguous.
    assert "duplicate-procedure" in _codes("Sub X\nEnd Sub\nProperty Get X() As Long\nEnd Property")


def test_duplicate_declaration_in_procedure() -> None:
    assert "duplicate-declaration" in _codes("Sub S\n    Dim a As Long\n    Dim a As Long\nEnd Sub")
    # Parameter colliding with a local is also a duplicate (flat procedure scope).
    assert "duplicate-declaration" in _codes("Sub S(a As Long)\n    Dim a As Long\nEnd Sub")
    assert "duplicate-declaration" not in _codes("Sub S\n    Dim a As Long\n    Dim b As Long\nEnd Sub")


def test_duplicate_module_member() -> None:
    assert "duplicate-module-variable" in _codes("Public a As Long\nPublic a As Long")
    assert "duplicate-module-variable" in _codes("Private a As Long\nPublic Const a As Long = 1")
    assert "duplicate-module-variable" not in _codes("Public a As Long\nPublic b As Long")


def test_duplicate_enum_member() -> None:
    assert "duplicate-enum-member" in _codes("Enum Color\n    Red\n    Red\nEnd Enum")
    assert "duplicate-enum-member" not in _codes("Enum Color\n    Red\n    Green\nEnd Enum")


def test_duplicate_enum_member_across_inactive_branches_is_clean() -> None:
    # Same-named members in opposite #If arms never compile together.
    src = "Enum Color\n#If Win32 Then\n    Red\n#Else\n    Red\n#End If\nEnd Enum"
    assert "duplicate-enum-member" not in _codes(src)


def test_duplicate_type_field() -> None:
    assert "duplicate-type-field" in _codes("Type T\n    X As Long\n    X As Long\nEnd Type")
    assert "duplicate-type-field" not in _codes("Type T\n    X As Long\n    Y As Long\nEnd Type")


@pytest.mark.parametrize("code", _DUP_CODES)
def test_oracle_asserted_cases_fire(code: str) -> None:
    cases = asserted_cases(code)
    if not cases:
        pytest.skip(f"{code} has no asserted corpus cases")
    for case in cases:
        assert code in case_codes(case), f"{case.id}: expected {code}"


def test_no_duplicate_false_positives_on_accepted_cases() -> None:
    dup = set(_DUP_CODES)
    for case in accepted_cases():
        spurious = case_codes(case) & dup
        assert not spurious, f"{case.id}: duplicate false positive {spurious}"
