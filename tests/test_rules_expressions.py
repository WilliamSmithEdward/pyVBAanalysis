"""M6: expression-syntax rules (expressions.ts parity, self-contained slice)."""

from __future__ import annotations

from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

_CODE = "unbalanced-parens"


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_missing_close_paren() -> None:
    assert _CODE in _codes("Sub S\n    x = Foo(1, 2\nEnd Sub")


def test_unexpected_close_paren() -> None:
    assert _CODE in _codes("Sub S\n    x = 1)\nEnd Sub")


def test_balanced_is_clean() -> None:
    assert _CODE not in _codes("Sub S\n    x = Foo(1, (2 + 3))\nEnd Sub")
    # Parens reset at each statement boundary (newline / depth-0 colon).
    assert _CODE not in _codes("Sub S\n    a = (1) : b = (2)\nEnd Sub")
    # Parens inside strings/comments are distinct token kinds, never counted.
    assert _CODE not in _codes('Sub S\n    s = "a)b("\nEnd Sub')


def test_oracle_asserted_cases() -> None:
    if asserted_cases(_CODE):
        assert assert_oracle_behavior(_CODE) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        assert _CODE not in case_codes(case), f"{case.id}: {_CODE} false positive"
