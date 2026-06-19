"""M9: eventHandlerModuleScope rule (moduleKind.ts checkEventHandlerModuleScope parity).

Flags a Sub whose name matches an Excel event handler that the current module's
document type does not wire (or any non-document module). 0 oracle cases assert
this style-policy code, so it is validated by a no-false-positive sweep over the
FULL accepted corpus plus direct positives/controls. The rule reads the vendored
event catalogue (data/event_definitions.json, extracted by
tools/extract_event_definitions.mjs). The rule is wired in the real registry, so a
plain analyze_module exercises it.
"""

from __future__ import annotations

from oracle_support import CASES, case_codes  # type: ignore[attr-defined]

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module
from pyvbaanalysis.symbols import ModuleSymbolKind

_CODE = "event-handler-module-scope"


def _codes(source: str, opts: AnalyzeModuleOptions) -> set[str]:
    return {d.code for d in analyze_module(source, opts)}


def _doc(module_name: str) -> AnalyzeModuleOptions:
    return AnalyzeModuleOptions(module_name=module_name, module_kind=ModuleSymbolKind.DOCUMENT)


# -- positives -------------------------------------------------------------


def test_worksheet_event_in_thisworkbook_fires() -> None:
    src = "Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub"
    assert _CODE in _codes(src, _doc("ThisWorkbook"))


def test_workbook_event_in_worksheet_module_fires() -> None:
    src = "Private Sub Workbook_Open()\nEnd Sub"
    assert _CODE in _codes(src, _doc("Sheet1"))


def test_event_in_standard_module_fires() -> None:
    src = "Public Sub Workbook_Open()\nEnd Sub"
    opts = AnalyzeModuleOptions(module_name="Module1", module_kind=ModuleSymbolKind.STANDARD)
    assert _CODE in _codes(src, opts)


def test_chart_event_in_worksheet_module_fires() -> None:
    src = "Private Sub Chart_Activate()\nEnd Sub"
    assert _CODE in _codes(src, _doc("Sheet1"))


# -- controls (must stay silent) -------------------------------------------


def test_workbook_event_in_thisworkbook_is_silent() -> None:
    src = "Private Sub Workbook_Open()\nEnd Sub"
    assert _CODE not in _codes(src, _doc("ThisWorkbook"))


def test_worksheet_event_in_sheet_module_is_silent() -> None:
    src = "Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub"
    assert _CODE not in _codes(src, _doc("Sheet1"))


def test_chart_event_in_chart_module_is_silent() -> None:
    src = "Private Sub Chart_Activate()\nEnd Sub"
    assert _CODE not in _codes(src, _doc("Chart1"))


def test_non_event_sub_is_silent() -> None:
    src = "Private Sub DoStuff()\nEnd Sub"
    assert _CODE not in _codes(src, _doc("ThisWorkbook"))


def test_explicit_document_type_overrides_name() -> None:
    # documentType=worksheet wins over the ThisWorkbook name heuristic.
    opts = AnalyzeModuleOptions(
        module_name="ThisWorkbook",
        module_kind=ModuleSymbolKind.DOCUMENT,
        document_type="worksheet",
    )
    src = "Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub"
    assert _CODE not in _codes(src, opts)


# -- no-false-positive sweep over the full accepted corpus -----------------


def test_no_false_positives_on_accepted_cases() -> None:
    offenders = [
        case.id for case in CASES.values()
        if case.expected == "accepted" and _CODE in case_codes(case)
    ]
    assert offenders == [], f"{_CODE} false positive(s): {offenders}"
