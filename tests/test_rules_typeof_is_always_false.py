"""M9: typeOfIsAlwaysFalse rule (typeOfIs.ts checkTypeOfIsCompatibility parity).

`TypeOf x Is T` is provably False when x's declared type is a concrete object class
incompatible with T. 0 oracle cases assert this code, so it is validated by a
no-false-positive sweep over the FULL accepted corpus (it must stay silent) plus
direct positives/controls. The rule is wired in the real registry, so a plain
analyze_module exercises it.
"""

from __future__ import annotations

from oracle_support import CASES, case_codes  # type: ignore[attr-defined]

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module
from pyvbaanalysis.symbols import ModuleSymbolKind

_CODE = "typeof-is-always-false"


def _codes(source: str, opts: AnalyzeModuleOptions | None = None) -> set[str]:
    return {d.code for d in analyze_module(source, opts)}


_STD = AnalyzeModuleOptions(module_name="Module1", module_kind=ModuleSymbolKind.STANDARD)


# -- direct positives ------------------------------------------------------


def test_two_concrete_host_classes_fires() -> None:
    src = (
        "Public Sub S()\n    Dim ws As Worksheet\n"
        "    If TypeOf ws Is Workbook Then\n    End If\nEnd Sub"
    )
    assert _CODE in _codes(src, _STD)


def test_range_vs_worksheet_fires() -> None:
    src = (
        "Public Sub S()\n    Dim r As Range\n"
        "    If TypeOf r Is Worksheet Then\n    End If\nEnd Sub"
    )
    assert _CODE in _codes(src, _STD)


# -- controls (must stay silent) -------------------------------------------


def test_same_type_is_silent() -> None:
    src = (
        "Public Sub S()\n    Dim ws As Worksheet\n"
        "    If TypeOf ws Is Worksheet Then\n    End If\nEnd Sub"
    )
    assert _CODE not in _codes(src, _STD)


def test_object_operand_is_silent() -> None:
    src = (
        "Public Sub S()\n    Dim o As Object\n"
        "    If TypeOf o Is Workbook Then\n    End If\nEnd Sub"
    )
    assert _CODE not in _codes(src, _STD)


def test_variant_operand_is_silent() -> None:
    src = (
        "Public Sub S()\n    Dim v As Variant\n"
        "    If TypeOf v Is Workbook Then\n    End If\nEnd Sub"
    )
    assert _CODE not in _codes(src, _STD)


def test_is_object_target_is_silent() -> None:
    src = (
        "Public Sub S()\n    Dim ws As Worksheet\n"
        "    If TypeOf ws Is Object Then\n    End If\nEnd Sub"
    )
    assert _CODE not in _codes(src, _STD)


def test_unknown_operand_type_is_silent() -> None:
    src = (
        "Public Sub S()\n    Dim x As Foo\n"
        "    If TypeOf x Is Workbook Then\n    End If\nEnd Sub"
    )
    assert _CODE not in _codes(src, _STD)


def test_undeclared_operand_is_silent() -> None:
    src = "Public Sub S()\n    If TypeOf x Is Workbook Then\n    End If\nEnd Sub"
    assert _CODE not in _codes(src, _STD)


# -- no-false-positive sweep over the full accepted corpus -----------------


def test_no_false_positives_on_accepted_cases() -> None:
    # 0 asserted oracle cases: validate by sweeping every accepted (compile-valid)
    # case. typeof-is-always-false is a runtime-risk diagnostic, but a provably-dead
    # TypeOf branch never appears in compile-accepted corpus code, so it must stay
    # silent across the whole accepted corpus.
    offenders = [
        case.id for case in CASES.values()
        if case.expected == "accepted" and _CODE in case_codes(case)
    ]
    assert offenders == [], f"{_CODE} false positive(s): {offenders}"
