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
    "required-param-after-optional",
    "paramarray-non-variant",
    "paramarray-with-optional",
    "paramarray-not-last",
    "parameter-array-as-type-syntax",
    "property-accessor-signature-mismatch",
    "const-value-not-constant",
    "enum-member-not-constant",
    "module-declaration-in-procedure",
    "module-declaration-after-procedure",
    "statement-outside-procedure",
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


def test_parameter_order() -> None:
    assert "required-param-after-optional" in _codes("Sub S(Optional a As Long, b As Long)\nEnd Sub")
    assert "required-param-after-optional" not in _codes("Sub S(a As Long, Optional b As Long)\nEnd Sub")
    assert "paramarray-non-variant" in _codes("Sub S(ParamArray a() As Long)\nEnd Sub")
    assert "paramarray-non-variant" not in _codes("Sub S(ParamArray a() As Variant)\nEnd Sub")
    assert "paramarray-not-last" in _codes("Sub S(ParamArray a() As Variant, b As Long)\nEnd Sub")
    assert "paramarray-with-optional" in _codes("Sub S(Optional a As Long, ParamArray b() As Variant)\nEnd Sub")


def test_property_accessor_signatures() -> None:
    mismatch = (
        "Property Get X(i As Long) As Long\nEnd Property\n"
        "Property Let X(i As String, v As Long)\nEnd Property"
    )
    assert "property-accessor-signature-mismatch" in _codes(mismatch)
    ok = (
        "Property Get X(i As Long) As Long\nEnd Property\n"
        "Property Let X(i As Long, v As Long)\nEnd Property"
    )
    assert "property-accessor-signature-mismatch" not in _codes(ok)


def test_non_constant_values() -> None:
    assert "const-value-not-constant" in _codes("Const C As Long = Foo()")
    assert "const-value-not-constant" not in _codes("Const C As Long = 5 + 3")
    # Bare/qualified identifiers may be constants; stay quiet.
    assert "const-value-not-constant" not in _codes("Const C As Long = OTHER")
    assert "enum-member-not-constant" in _codes("Enum E\n    A = Bar()\nEnd Enum")
    assert "enum-member-not-constant" not in _codes("Enum E\n    A = 1\n    B\nEnd Enum")


def test_module_declaration_placement() -> None:
    assert "module-declaration-in-procedure" in _codes("Sub S\n    Public X As Long\nEnd Sub")
    assert "module-declaration-in-procedure" in _codes("Sub S\n    Option Explicit\nEnd Sub")
    assert "module-declaration-in-procedure" not in _codes("Sub S\n    Dim x As Long\nEnd Sub")
    assert "module-declaration-after-procedure" in _codes("Sub S\nEnd Sub\nPublic X As Long")
    assert "module-declaration-after-procedure" not in _codes("Public X As Long\nSub S\nEnd Sub")
    assert "statement-outside-procedure" in _codes("MsgBox 1")
    # Def* statements and Implements are legal module-level statement forms.
    assert "statement-outside-procedure" not in _codes("DefInt A-Z")


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
