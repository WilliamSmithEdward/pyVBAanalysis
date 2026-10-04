"""Overflow folding past the expression depth limit (XLIDE 0c39d113, #737).

Mirrors upstream's tests/diagnostics/overflowDepthRecovery.test.ts: an expression
nested past MAX_EXPRESSION_DEPTH is unknown, and the module's other overflow
findings are still reported.
"""

from __future__ import annotations

import pytest

from pyvbaanalysis.diagnostics.rules.overflow import check_overflow
from pyvbaanalysis.parser.nodes import Span
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.symbols.build_module_symbols import BuildModuleSymbolsOptions, build_module_symbols
from pyvbaanalysis.symbols.symbol_model import ModuleSymbolKind


def _pushed(source: str) -> list[tuple[str, str, Span]]:
    module = parse_module(source)
    symbols = build_module_symbols(
        "Module", ModuleSymbolKind.STANDARD, source, BuildModuleSymbolsOptions(parsed_module=module)
    )
    calls: list[tuple[str, str, Span]] = []

    def push(rule: str, message: str, span: Span, data: object = None) -> None:
        calls.append((rule, message, span))

    check_overflow(source, module, symbols, None, None, None, push)
    return calls


@pytest.mark.parametrize(
    "expression",
    [
        "(" * 3000 + "1" + ")" * 3000,
        "CInt(" * 1500 + "1" + ")" * 1500,
        "Sgn(" * 1500 + "1" + ")" * 1500,
        "Not " * 3000 + "1",
    ],
    ids=["parentheses", "conversions", "intrinsics", "Not"],
)
def test_keeps_later_overflow_findings_after_excessive_nesting(expression: str) -> None:
    source = (
        "Private Const Deep = " + expression + "\nPrivate Const Bad As Integer = 40000\n"
        "Sub Main()\nDim i As Integer\ni = 40000\nEnd Sub"
    )
    calls = _pushed(source)
    assert len([call for call in calls if call[0] == "constOverflow"]) == 1
    assert len([call for call in calls if call[0] == "arithmeticOverflow"]) == 1


def test_keeps_checking_after_deep_unknown_calls_in_a_procedure_statement() -> None:
    expression = "F(" * 1500 + "1" + ")" * 1500
    source = "Sub Main()\nDim i As Integer\nFoo " + expression + "\ni = 40000\nEnd Sub"
    assert len([call for call in _pushed(source) if call[0] == "arithmeticOverflow"]) == 1


def test_still_folds_ordinary_nested_expressions_and_reports_their_original_spans() -> None:
    source = "Private Const Deep = " + "CInt(" * 20 + "32767 + 1" + ")" * 20
    calls = _pushed(source)
    assert len(calls) == 1
    assert calls[0][0] == "constOverflow"
    span = calls[0][2]
    assert source[span.start : span.end] == "32767 + 1"
