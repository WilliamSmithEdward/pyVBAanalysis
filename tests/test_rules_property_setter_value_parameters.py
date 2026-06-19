"""M9: propertySetterValueParameters rule (declarations.ts parity, safe branches).

Ships the pure signature/structure branches:
  - property-setter-missing-value   (Let/Set with no parameters)
  - property-setter-return-type     (Let/Set declares an As return type)
  - property-set-scalar-value       (Property Set final value param is scalar)
DEFERS the propertyLetObjectValue branch (object-value resolution needs the host /
project class-assignment surface). The one asserted oracle case
(corpus_prop_005_compile -> property-setter-missing-value) is in a shipped branch,
so no skip_ids. The rule is wired in the real registry, so a plain analyze_module
exercises it.
"""

from __future__ import annotations

from oracle_support import (  # type: ignore[attr-defined]
    accepted_cases,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module

# Codes the shipped branches emit; the no-FP sweep checks all three.
_CODES = (
    "property-setter-missing-value",
    "property-setter-return-type",
    "property-set-scalar-value",
)
_ORACLE_CODE = "property-setter-missing-value"


def _codes(source: str, opts: AnalyzeModuleOptions | None = None) -> set[str]:
    return {d.code for d in analyze_module(source, opts)}


# -- direct positives ------------------------------------------------------


def test_let_missing_value_fires() -> None:
    assert "property-setter-missing-value" in _codes(
        "Public Property Let Name()\nEnd Property"
    )


def test_set_missing_value_fires() -> None:
    assert "property-setter-missing-value" in _codes(
        "Public Property Set Thing()\nEnd Property"
    )


def test_let_return_type_fires() -> None:
    assert "property-setter-return-type" in _codes(
        "Public Property Let P(ByVal v As Long) As Long\nEnd Property"
    )


def test_set_scalar_value_fires() -> None:
    assert "property-set-scalar-value" in _codes(
        "Public Property Set P(ByVal v As Long)\nEnd Property"
    )


# -- controls (must stay silent) -------------------------------------------


def test_well_formed_let_is_silent() -> None:
    src = "Public Property Let P(ByVal v As Long)\nEnd Property"
    assert not _codes(src) & set(_CODES)


def test_well_formed_set_object_value_is_silent() -> None:
    src = "Public Property Set P(ByVal v As Object)\nEnd Property"
    assert not _codes(src) & set(_CODES)


def test_property_get_is_unaffected() -> None:
    src = "Public Property Get P() As Long\nEnd Property"
    assert not _codes(src) & set(_CODES)


# -- oracle ----------------------------------------------------------------


def test_oracle_asserted_case_fires() -> None:
    cases = asserted_cases(_ORACLE_CODE)
    assert cases, "expected at least one asserted case"
    for case in cases:
        emitted: set[str] = set()
        for module in case.modules:
            emitted |= _codes(module.source)
        if case.expected == "rejected":
            assert _ORACLE_CODE in emitted, f"{case.id}: expected {_ORACLE_CODE}"


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        offenders = oracle_false_positives(case, _CODES)
        assert not offenders, f"{case.id}: false positive(s) {sorted(offenders)}"
