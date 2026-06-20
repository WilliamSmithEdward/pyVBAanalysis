"""M10: propertySetterValueParameters rule (declarations.ts full parity).

Ships every branch:
  - property-setter-missing-value   (Let/Set with no parameters)
  - property-setter-return-type     (Let/Set declares an As return type)
  - property-set-scalar-value       (Property Set final value param is scalar)
  - property-let-object-value       (Property Let final value param is an object)
The propertyLetObjectValue branch (M10 slice 3c) resolves the value-param type via
resolveKnownObjectAssignmentType over the host / project class-assignment surface,
reusing the same helper as typeOfIsAlwaysFalse. It is a compile-error code with no
asserted oracle case (the corpus has no positive), so it is covered by direct unit
positives plus the all-accepted no-FP sweep. The rule is wired in the real
registry, so a plain analyze_module exercises every branch.
"""

from __future__ import annotations

from oracle_support import (  # type: ignore[attr-defined]
    accepted_cases,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module

# Codes the rule emits; the no-FP sweep checks all four.
_CODES = (
    "property-setter-missing-value",
    "property-setter-return-type",
    "property-set-scalar-value",
    "property-let-object-value",
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


def test_let_generic_object_value_fires() -> None:
    # A Property Let whose final value param is As Object must use Property Set.
    assert "property-let-object-value" in _codes(
        "Public Property Let P(ByVal v As Object)\nEnd Property"
    )


def test_let_host_object_value_fires() -> None:
    # As Range resolves to a host object type (host model loaded by default).
    assert "property-let-object-value" in _codes(
        "Public Property Let P(ByVal v As Range)\nEnd Property"
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
