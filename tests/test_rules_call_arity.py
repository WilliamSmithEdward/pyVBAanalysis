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


def test_same_module_alternatives_accepting_the_call_silent() -> None:
    # Two same-named procedures (one per `#If` arm, typically) leave the compiled
    # signature unknown; a call one of them accepts stays silent.
    code = "argument-count"
    src = (
        "Sub Foo(a As Long)\nEnd Sub\nSub Foo(a As Long, b As Long)\nEnd Sub\n"
        "Sub S()\n    Foo 1, 2\nEnd Sub"
    )
    assert code not in _codes(src)


def test_same_module_alternatives_all_rejecting_the_call() -> None:
    # A call that no alternative accepts is wrong under every build (XLIDE #58).
    src = (
        "Sub Foo(a As Long)\nEnd Sub\nSub Foo(a As Long, b As Long)\nEnd Sub\n"
        "Sub S()\n    Foo 1, 2, 3\nEnd Sub"
    )
    messages = [d.message for d in analyze_module(src) if d.code == "argument-count"]
    assert messages == ["Wrong number of arguments to 'Foo': expected 1 argument, but got 3."]


def test_single_line_if_call_statement() -> None:
    # The call a single-line If carries is a call statement too (XLIDE #46).
    src = "Sub Helper(ByVal a As Long)\nEnd Sub\nSub S()\n    If True Then Helper 1 Else Helper 1, 2\nEnd Sub"
    messages = [d.message for d in analyze_module(src) if d.code == "argument-count"]
    assert messages == ["Wrong number of arguments to 'Helper': expected 1 argument, but got 2."]


def test_host_member_call_arity() -> None:
    src = "Sub S()\n    Application.Calculate(1)\n    Err.Raise\n    Debug.Assert\nEnd Sub"
    messages = [d.message for d in analyze_module(src) if d.code == "argument-count"]
    assert messages == [
        "Wrong number of arguments to 'Calculate': expected 0 arguments, but got 1.",
        "Wrong number of arguments to 'Raise': expected between 1 and 5 arguments, but got 0.",
        "Wrong number of arguments to 'Assert': expected 1 argument, but got 0.",
    ]


def test_keyword_named_argument_is_named() -> None:
    # `Type` lexes as a keyword; before `:=` it is still a parameter name.
    src = (
        "Sub S()\n"
        "    ThisWorkbook.BreakLink Name:=\"test\", Type:=xlLinkTypeExcelLinks\n"
        "    ThisWorkbook.BreakLink Type:=1, \"test\"\n"
        "End Sub"
    )
    messages = [d.message for d in analyze_module(src) if d.code == "argument-count"]
    assert messages == [
        "A positional argument may not follow a named argument in the call to 'BreakLink'."
    ]


def test_oracle_asserted_cases() -> None:
    for code in _CODES:
        if asserted_cases(code):
            assert assert_oracle_behavior(code) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    codes = set(_CODES)
    for case in accepted_cases():
        spurious = case_codes(case) & codes
        assert not spurious, f"{case.id}: argument-count false positive {spurious}"
