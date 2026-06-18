"""M5b: the diagnostics engine skeleton (analyzeModule / walker / exprWalk parity).

The registry is empty, so analyze_module returns [] over the whole corpus. The
mechanism (per-rule buffering, registry-order flush, the three execution forms,
severity overrides) is exercised with synthetic rules patched into the registry.
"""

from __future__ import annotations

import sys

import pytest

from pyvbaanalysis.diagnostics import (
    DIAGNOSTIC_RULES,
    AnalyzeModuleOptions,
    DiagnosticSeverity,
    analyze_module,
)
from pyvbaanalysis.diagnostics.analyze_module import _severity_of
from pyvbaanalysis.diagnostics.registry import DiagnosticRuleEntry
from pyvbaanalysis.evidence import load_oracle_cases
from pyvbaanalysis.parser.nodes import Span
from pyvbaanalysis.symbols import ModuleSymbolKind

# The submodule object (not the re-exported analyze_module function): the engine
# reads DIAGNOSTIC_RULE_REGISTRY as a module global, so patch it on the submodule.
am_mod = sys.modules["pyvbaanalysis.diagnostics.analyze_module"]

_MODULE_KIND = {
    "standard": ModuleSymbolKind.STANDARD,
    "class": ModuleSymbolKind.CLASS,
    "document": ModuleSymbolKind.DOCUMENT,
    "userform": ModuleSymbolKind.USERFORM,
}
_SOURCES: list[tuple[str, str, str]] = [
    (f"{c.id}::{m.name}", m.module_type, m.source) for c in load_oracle_cases() for m in c.modules
]

# A real error-rule name and a real non-error rule name, used as synthetic codes.
_ERROR_RULE = next(n for n, m in DIAGNOSTIC_RULES.items() if m.default_severity is DiagnosticSeverity.ERROR)
_NON_ERROR_RULE = next(n for n, m in DIAGNOSTIC_RULES.items() if m.default_severity is not DiagnosticSeverity.ERROR)


@pytest.mark.parametrize(
    ("module_type", "source"),
    [(mt, s) for _, mt, s in _SOURCES],
    ids=[i for i, _, _ in _SOURCES],
)
def test_empty_registry_returns_empty_over_corpus(module_type: str, source: str) -> None:
    opts = AnalyzeModuleOptions(module_name="M", module_kind=_MODULE_KIND.get(module_type, ModuleSymbolKind.STANDARD))
    assert analyze_module(source, opts) == []


def test_analyze_module_never_throws_on_garbage() -> None:
    for src in (")(}{", "Sub", "If Then", "#If", "End Sub", "x = = )"):
        assert analyze_module(src) == []


def test_run_form_emits_and_stamps_metadata(monkeypatch: pytest.MonkeyPatch) -> None:
    def my_run(ctx: object, push: object) -> None:
        push(_ERROR_RULE, "boom", Span(2, 5))  # type: ignore[operator]

    monkeypatch.setattr(am_mod, "DIAGNOSTIC_RULE_REGISTRY", (DiagnosticRuleEntry(name=_ERROR_RULE, run=my_run),))
    diags = analyze_module("Sub S\nEnd Sub")
    assert len(diags) == 1
    meta = DIAGNOSTIC_RULES[_ERROR_RULE]
    assert diags[0].code == meta.code
    assert diags[0].message == "boom"
    assert diags[0].severity is DiagnosticSeverity.ERROR
    assert diags[0].span == Span(2, 5)
    assert diags[0].spec_reference == meta.spec_reference


def test_procedure_statements_form_visits_each_statement(monkeypatch: pytest.MonkeyPatch) -> None:
    def factory(ctx, push):  # type: ignore[no-untyped-def]
        def visitor(proc):  # type: ignore[no-untyped-def]
            def on_stmt(stmt):  # type: ignore[no-untyped-def]
                push(_ERROR_RULE, "stmt", stmt.span)
            return on_stmt
        return visitor

    monkeypatch.setattr(
        am_mod, "DIAGNOSTIC_RULE_REGISTRY", (DiagnosticRuleEntry(name=_ERROR_RULE, procedure_statements=factory),)
    )
    src = "Sub S\n    x = 1\n    y = 2\n    z = 3\nEnd Sub"
    diags = analyze_module(src)
    assert len(diags) == 3  # one per leaf statement
    assert all(d.message == "stmt" for d in diags)


def test_procedure_expressions_form_visits_each_expression(monkeypatch: pytest.MonkeyPatch) -> None:
    seen: list[str] = []

    def factory(ctx, push):  # type: ignore[no-untyped-def]
        def visitor(proc):  # type: ignore[no-untyped-def]
            def on_expr(expr):  # type: ignore[no-untyped-def]
                seen.append(type(expr).__name__)
            return on_expr
        return visitor

    monkeypatch.setattr(
        am_mod, "DIAGNOSTIC_RULE_REGISTRY", (DiagnosticRuleEntry(name=_ERROR_RULE, procedure_expressions=factory),)
    )
    analyze_module("Sub S\n    x = a + b\nEnd Sub")
    # x (lhs ident), a+b (binary), a, b => identifiers + binary visited.
    assert "BinaryExpr" in seen
    assert seen.count("IdentifierExpr") >= 3


def test_buffers_flush_in_registry_order(monkeypatch: pytest.MonkeyPatch) -> None:
    def make_run(message: str):  # type: ignore[no-untyped-def]
        def run(ctx, push):  # type: ignore[no-untyped-def]
            push(_ERROR_RULE, message, Span(0, 1))
        return run

    registry = (
        DiagnosticRuleEntry(name=_ERROR_RULE, run=make_run("first")),
        DiagnosticRuleEntry(name=_ERROR_RULE, run=make_run("second")),
    )
    monkeypatch.setattr(am_mod, "DIAGNOSTIC_RULE_REGISTRY", registry)
    diags = analyze_module("Sub S\nEnd Sub")
    assert [d.message for d in diags] == ["first", "second"]


def test_severity_override_off_disables_non_error_rule() -> None:
    code = DIAGNOSTIC_RULES[_NON_ERROR_RULE].code
    assert _severity_of(_NON_ERROR_RULE, {code: "off"}) is None
    assert _severity_of(_NON_ERROR_RULE, None) is DIAGNOSTIC_RULES[_NON_ERROR_RULE].default_severity


def test_severity_override_off_ignored_for_error_rule_without_downgrade() -> None:
    name = next(
        n for n, m in DIAGNOSTIC_RULES.items()
        if m.default_severity is DiagnosticSeverity.ERROR and not m.allow_severity_downgrade
    )
    code = DIAGNOSTIC_RULES[name].code
    assert _severity_of(name, {code: "off"}) is DiagnosticSeverity.ERROR


def test_severity_override_downgrade_when_allowed() -> None:
    candidates = [
        n for n, m in DIAGNOSTIC_RULES.items()
        if m.default_severity is DiagnosticSeverity.ERROR and m.allow_severity_downgrade
    ]
    if not candidates:
        pytest.skip("no downgradeable error rule in the catalogue")
    code = DIAGNOSTIC_RULES[candidates[0]].code
    assert _severity_of(candidates[0], {code: "warning"}) is DiagnosticSeverity.WARNING


def test_off_rule_emits_nothing(monkeypatch: pytest.MonkeyPatch) -> None:
    code = DIAGNOSTIC_RULES[_NON_ERROR_RULE].code

    def my_run(ctx, push):  # type: ignore[no-untyped-def]
        push(_NON_ERROR_RULE, "hidden", Span(0, 1))

    monkeypatch.setattr(am_mod, "DIAGNOSTIC_RULE_REGISTRY", (DiagnosticRuleEntry(name=_NON_ERROR_RULE, run=my_run),))
    assert analyze_module("Sub S\nEnd Sub", AnalyzeModuleOptions(severity_overrides={code: "off"})) == []


def test_inactive_branch_statements_not_visited(monkeypatch: pytest.MonkeyPatch) -> None:
    count = {"n": 0}

    def factory(ctx, push):  # type: ignore[no-untyped-def]
        def visitor(proc):  # type: ignore[no-untyped-def]
            def on_stmt(stmt):  # type: ignore[no-untyped-def]
                count["n"] += 1
            return on_stmt
        return visitor

    monkeypatch.setattr(
        am_mod, "DIAGNOSTIC_RULE_REGISTRY", (DiagnosticRuleEntry(name=_ERROR_RULE, procedure_statements=factory),)
    )
    src = "Sub S\n#If Win32 Then\n    a = 1\n#Else\n    b = 2\n#End If\nEnd Sub"
    analyze_module(src)
    assert count["n"] == 1  # only the active (#Else) branch statement is walked
