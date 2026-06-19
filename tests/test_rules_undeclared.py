"""M9: unresolved-name rule family (undeclared.ts parity)."""

from __future__ import annotations

from oracle_support import (
    accepted_cases,
    assert_oracle_behavior,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module

_CODES = (
    "option-explicit-missing",
    "undeclared-variable",
    "unknown-call",
    "non-callable-call",
)

# option-explicit-missing is a style-policy warning (vbeCompileEquivalent=false):
# firing on valid code that lacks Option Explicit is its intended behavior, not a
# false positive, so it is excluded from the no-false-positive sweep. The three
# compile-error codes must never fire on accepted code.
_NO_FALSE_POSITIVE_CODES = (
    "undeclared-variable",
    "unknown-call",
    "non-callable-call",
)


def _codes(source: str, opts: AnalyzeModuleOptions | None = None) -> set[str]:
    return {d.code for d in analyze_module(source, opts)}


# -- direct unit tests -----------------------------------------------------


def test_option_explicit_missing_fires_on_code_module() -> None:
    src = "Public Sub Foo()\n    Dim x As Long\nEnd Sub"
    assert "option-explicit-missing" in _codes(src)


def test_option_explicit_present_is_silent() -> None:
    src = "Option Explicit\nPublic Sub Foo()\n    Dim x As Long\nEnd Sub"
    assert "option-explicit-missing" not in _codes(src)


def test_option_explicit_skips_empty_module() -> None:
    assert "option-explicit-missing" not in _codes("")
    assert "option-explicit-missing" not in _codes('Attribute VB_Name = "Module1"\n')


def test_undeclared_variable_assignment_fires() -> None:
    src = "Option Explicit\nPublic Sub Foo()\n    notDeclared = 1\nEnd Sub"
    opts = AnalyzeModuleOptions(known_identifiers=frozenset({"foo"}))
    assert "undeclared-variable" in _codes(src, opts)


def test_undeclared_variable_self_gated_without_known_identifiers() -> None:
    # No known_identifiers supplied -> the rule no-ops (no false positives on a
    # single module analyzed without project context).
    src = "Option Explicit\nPublic Sub Foo()\n    notDeclared = 1\nEnd Sub"
    assert "undeclared-variable" not in _codes(src, None)


def test_undeclared_variable_self_gated_without_option_explicit() -> None:
    src = "Public Sub Foo()\n    notDeclared = 1\nEnd Sub"
    opts = AnalyzeModuleOptions(known_identifiers=frozenset({"foo"}))
    assert "undeclared-variable" not in _codes(src, opts)


def test_undeclared_variable_declared_local_is_silent() -> None:
    src = "Option Explicit\nPublic Sub Foo()\n    Dim x As Long\n    x = 1\nEnd Sub"
    opts = AnalyzeModuleOptions(known_identifiers=frozenset({"foo"}))
    assert "undeclared-variable" not in _codes(src, opts)


def test_unknown_call_statement_fires() -> None:
    src = "Public Sub Foo()\n    DoesNotExist\nEnd Sub"
    opts = AnalyzeModuleOptions(known_procedures=frozenset({"foo"}))
    assert "unknown-call" in _codes(src, opts)


def test_unknown_call_statement_self_gated() -> None:
    src = "Public Sub Foo()\n    DoesNotExist\nEnd Sub"
    assert "unknown-call" not in _codes(src, None)


def test_unknown_call_statement_known_procedure_is_silent() -> None:
    src = "Public Sub Foo()\n    Bar\nEnd Sub\nPublic Sub Bar()\nEnd Sub"
    opts = AnalyzeModuleOptions(known_procedures=frozenset({"foo", "bar"}))
    assert "unknown-call" not in _codes(src, opts)


def test_non_callable_call_statement_fires() -> None:
    src = 'Public Sub Foo()\n    Dim testStr As String\n    testStr = "hi"\n    testStr\nEnd Sub'
    assert "non-callable-call" in _codes(src)


def test_non_callable_call_with_argument_fires() -> None:
    src = 'Public Sub Foo()\n    Dim testStr As String\n    testStr = "hi"\n    testStr "x"\nEnd Sub'
    assert "non-callable-call" in _codes(src)


def test_callable_statement_is_silent() -> None:
    src = "Public Sub Foo()\n    Bar\nEnd Sub\nPublic Sub Bar()\nEnd Sub"
    assert "non-callable-call" not in _codes(src)


# -- oracle sweeps ---------------------------------------------------------

# Cases whose assertion needs infrastructure outside this slice. (None currently:
# the asserted undeclared-variable / non-callable-call cases resolve with only the
# host + project surfaces this slice already consumes.)
_SKIP_IDS: frozenset[str] = frozenset()


def test_oracle_asserted_cases() -> None:
    checked_any = False
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code, _SKIP_IDS) > 0
            checked_any = True
    assert checked_any


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, _NO_FALSE_POSITIVE_CODES)
        assert not spurious, f"{case.id}: undeclared-family false positive {spurious}"
