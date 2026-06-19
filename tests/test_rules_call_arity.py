"""M8: call-argument arity (callArity.ts parity)."""

from __future__ import annotations

from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

_CODES = ("argument-count",)


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_too_few_arguments() -> None:
    code = "argument-count"
    src = "Sub Foo(a As Long, b As Long)\nEnd Sub\nSub S()\n    Foo 1\nEnd Sub"
    assert code in _codes(src)


def test_too_many_arguments() -> None:
    code = "argument-count"
    src = "Sub Foo(a As Long)\nEnd Sub\nSub S()\n    Foo 1, 2\nEnd Sub"
    assert code in _codes(src)


def test_correct_argument_count_silent() -> None:
    code = "argument-count"
    src = "Sub Foo(a As Long, Optional b As Long)\nEnd Sub\nSub S()\n    Foo 1\nEnd Sub"
    assert code not in _codes(src)


def test_named_argument_not_found() -> None:
    code = "argument-count"
    src = "Sub Foo(a As Long)\nEnd Sub\nSub S()\n    Foo zzz:=1\nEnd Sub"
    assert code in _codes(src)


def test_ambiguous_same_module_target_silent() -> None:
    # Two same-named procedures make the signature non-unique; stay silent.
    code = "argument-count"
    src = (
        "Sub Foo(a As Long)\nEnd Sub\nSub Foo(a As Long, b As Long)\nEnd Sub\n"
        "Sub S()\n    Foo 1, 2, 3\nEnd Sub"
    )
    assert code not in _codes(src)


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    codes = set(_CODES)
    for case in accepted_cases():
        spurious = case_codes(case) & codes
        assert not spurious, f"{case.id}: argument-count false positive {spurious}"
