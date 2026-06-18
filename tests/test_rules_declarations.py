"""M6: declaration-site rule family (declarations.ts parity, self-contained slice)."""

from __future__ import annotations

import pytest
from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

# Every code emitted by the ported declaration rules.
_DECL_CODES = (
    "invalid-proc-header",
    "invalid-identifier-start",
    "invalid-identifier-character",
    "invalid-declaration-name",
    "dim-initializer",
    "unexpected-declaration-token",
    "type-declaration-character-as-clause",
    "option-after-declaration",
    "empty-type",
    "duplicate-option",
    "too-many-parameters",
    "identifier-too-long",
    "optional-udt-parameter",
    "byval-udt-parameter",
)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_procedure_header() -> None:
    assert "invalid-proc-header" in _codes("Sub My Sub\nEnd Sub")
    assert "invalid-proc-header" not in _codes("Sub MySub()\nEnd Sub")
    assert "invalid-proc-header" not in _codes("Function F() As Long\nEnd Function")
    # Sub/Property Let have no return type; an `As` right after the name is wrong.
    assert "invalid-proc-header" in _codes("Sub Foo As Long\nEnd Sub")


def test_invalid_identifier_starts() -> None:
    assert "invalid-identifier-start" in _codes("Dim 1bad As Long")
    assert "invalid-identifier-character" in _codes("Dim user-name As Long")
    assert "invalid-identifier-start" not in _codes("Dim good_name As Long")
    # Bracketed names may contain anything.
    assert "invalid-identifier-start" not in _codes("Dim [1weird] As Long")


def test_reserved_declaration_name() -> None:
    assert "invalid-declaration-name" in _codes("Dim Loop As Long")
    assert "invalid-declaration-name" not in _codes("Dim Total As Long")
    # Bracketed escapes a reserved word.
    assert "invalid-declaration-name" not in _codes("Dim [Loop] As Long")


def test_dim_initializer() -> None:
    assert "dim-initializer" in _codes("Sub S\n    Dim x As Long = 5\nEnd Sub")
    assert "dim-initializer" not in _codes("Sub S\n    Dim x As Long\nEnd Sub")
    # Const legitimately uses '='.
    assert "dim-initializer" not in _codes("Const C As Long = 5")


def test_unexpected_declaration_token() -> None:
    assert "unexpected-declaration-token" in _codes("Dim s As String junk")
    assert "unexpected-declaration-token" not in _codes("Dim s As String")
    assert "unexpected-declaration-token" not in _codes("Dim s As String, t As Long")


def test_type_declaration_character_with_as_clause() -> None:
    assert "type-declaration-character-as-clause" in _codes("Dim x$ As String")
    assert "type-declaration-character-as-clause" not in _codes("Dim x As String")
    assert "type-declaration-character-as-clause" not in _codes("Dim x$")


def test_option_placement_and_duplication() -> None:
    assert "option-after-declaration" in _codes("Public X As Long\nOption Explicit")
    assert "option-after-declaration" not in _codes("Option Explicit\nPublic X As Long")
    assert "duplicate-option" in _codes("Option Explicit\nOption Explicit")
    assert "duplicate-option" not in _codes("Option Explicit\nOption Base 1")


def test_empty_type() -> None:
    assert "empty-type" in _codes("Type T\nEnd Type")
    assert "empty-type" not in _codes("Type T\n    X As Long\nEnd Type")


def test_too_many_parameters() -> None:
    params = ", ".join(f"a{i} As Long" for i in range(61))
    assert "too-many-parameters" in _codes(f"Sub S({params})\nEnd Sub")
    ok_params = ", ".join(f"a{i} As Long" for i in range(60))
    assert "too-many-parameters" not in _codes(f"Sub S({ok_params})\nEnd Sub")


def test_identifier_too_long() -> None:
    assert "identifier-too-long" in _codes("Dim " + "a" * 256 + " As Long")
    assert "identifier-too-long" not in _codes("Dim " + "a" * 255 + " As Long")


def test_udt_parameter_constraints() -> None:
    udt = "Type T\n    X As Long\nEnd Type\n"
    assert "byval-udt-parameter" in _codes(udt + "Sub S(ByVal p As T)\nEnd Sub")
    assert "optional-udt-parameter" in _codes(udt + "Sub S(Optional p As T)\nEnd Sub")
    # ByRef UDT is fine.
    assert "byval-udt-parameter" not in _codes(udt + "Sub S(ByRef p As T)\nEnd Sub")


@pytest.mark.parametrize("code", _DECL_CODES)
def test_oracle_asserted_cases(code: str) -> None:
    if not asserted_cases(code):
        pytest.skip(f"{code} has no asserted corpus cases")
    assert assert_oracle_behavior(code) > 0


def test_no_declaration_false_positives_on_accepted_cases() -> None:
    decl = set(_DECL_CODES)
    for case in accepted_cases():
        spurious = case_codes(case) & decl
        assert not spurious, f"{case.id}: declaration false positive {spurious}"
