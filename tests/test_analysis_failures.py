"""analyze_module reports what it could not check (XLIDE issue #178).

analyze_module never raises, and until upstream's #178 a failure of its own was
invisible: a rule that raised, or one walk visitor that raised and stopped the
walk for every rule after it, came back as fewer findings with nothing to say so.
Each failure is now reported to on_internal_error, and a visitor that raises
stops only its own rule. Mirrors upstream's tests/analysisInternalErrors.test.ts.
"""

from __future__ import annotations

import sys
from collections.abc import Callable
from typing import Any

import pytest

from pyvbaanalysis import AnalysisFailure, analyze_module, analyze_module_options_for, build_project_index
from pyvbaanalysis.diagnostics import DIAGNOSTIC_RULE_REGISTRY, AnalyzeModuleOptions
from pyvbaanalysis.diagnostics.registry import DiagnosticRuleEntry
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind

am_mod = sys.modules["pyvbaanalysis.diagnostics.analyze_module"]

LIB = "Option Explicit\r\nPublic Function Twice(ByVal n As Long) As Long\r\n    Twice = n * 2\r\nEnd Function\r\n"
CALLER = "\r\n".join(
    [
        "Option Explicit",
        "Sub First()",
        "    Dim r As Long",
        "    r = Twice(1, 2)",
        "End Sub",
        "Sub Second()",
        "    Dim q As Long",
        "    q = Twice(3, 4)",
        "End Sub",
        "",
    ]
)


def _options() -> AnalyzeModuleOptions:
    index = build_project_index(
        [
            ModuleInput("Lib", ModuleSymbolKind.STANDARD, LIB),
            ModuleInput("Caller", ModuleSymbolKind.STANDARD, CALLER),
        ]
    )
    return analyze_module_options_for(index, "Caller", ModuleSymbolKind.STANDARD)


def _analyze(opts: AnalyzeModuleOptions) -> tuple[list[str], list[tuple[AnalysisFailure, str]]]:
    failures: list[tuple[AnalysisFailure, str]] = []
    opts.on_internal_error = lambda error, where: failures.append((where, str(error)))
    codes = [f"{d.code}@{d.span.start}" for d in analyze_module(CALLER, opts)]
    return codes, failures


def _with_rules(monkeypatch: pytest.MonkeyPatch, first: DiagnosticRuleEntry | None, last: DiagnosticRuleEntry | None) -> None:
    registry = (*([first] if first else []), *DIAGNOSTIC_RULE_REGISTRY, *([last] if last else []))
    monkeypatch.setattr(am_mod, "DIAGNOSTIC_RULE_REGISTRY", registry)


def _raise(message: str) -> Callable[..., Any]:
    def fail(*args: object) -> Any:
        raise RuntimeError(message)

    return fail


def test_a_list_of_project_procedures_is_left_out_and_reported() -> None:
    baseline_codes, baseline_failures = _analyze(_options())
    assert baseline_failures == []
    assert len([code for code in baseline_codes if code.startswith("argument-count")]) == 2

    opts = _options()
    opts.project_procedures = [signature for group in (opts.project_procedures or {}).values() for signature in group]  # type: ignore[assignment]
    codes, failures = _analyze(opts)
    assert [(where.stage, where.rule) for where, _ in failures] == [("options", None)]
    assert "not a list" in failures[0][1]
    without = _options()
    without.project_procedures = None
    assert codes == _analyze(without)[0]


def test_a_rule_that_raises_is_reported_and_the_others_keep_their_findings(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    baseline, _ = _analyze(_options())
    _with_rules(monkeypatch, None, DiagnosticRuleEntry(name="probeRun", run=_raise("run failed")))
    codes, failures = _analyze(_options())
    assert [(where, message) for where, message in failures] == [(AnalysisFailure("rule", "probeRun"), "run failed")]
    assert codes == baseline


def test_a_statement_visitor_that_raises_stops_only_its_own_rule(monkeypatch: pytest.MonkeyPatch) -> None:
    baseline, _ = _analyze(_options())
    assert baseline
    # Registered first, it walks ahead of every real rule.
    probe = DiagnosticRuleEntry(
        name="probeStatements",
        procedure_statements=lambda ctx, push: _raise("statement visitor failed"),
    )
    _with_rules(monkeypatch, probe, None)
    codes, failures = _analyze(_options())
    assert failures == [(AnalysisFailure("statement-walk", "probeStatements"), "statement visitor failed")]
    assert codes == baseline


def test_an_expression_visitor_that_raises_stops_only_its_own_rule(monkeypatch: pytest.MonkeyPatch) -> None:
    baseline, _ = _analyze(_options())
    probe = DiagnosticRuleEntry(
        name="probeExpressions",
        procedure_expressions=lambda ctx, push: _raise("expression visitor failed"),
    )
    _with_rules(monkeypatch, probe, None)
    codes, failures = _analyze(_options())
    assert failures == [(AnalysisFailure("expression-walk", "probeExpressions"), "expression visitor failed")]
    assert codes == baseline


def test_the_whole_pass_failing_is_reported_and_returns_nothing() -> None:
    opts = _options()
    opts.parsed_module = object()  # type: ignore[assignment]
    codes, failures = _analyze(opts)
    assert codes == []
    assert [where.stage for where, _ in failures] == ["analysis"]


def test_a_callback_that_raises_is_ignored() -> None:
    opts = _options()
    opts.parsed_module = object()  # type: ignore[assignment]
    opts.on_internal_error = _raise("callback failed")
    assert analyze_module(CALLER, opts) == []
