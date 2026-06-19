"""M7: assignment-family rules (assignments.ts parity, Mid-statement slice)."""

from __future__ import annotations

from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

_CODE = "mid-statement-literal-target"


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_mid_statement_literal_target() -> None:
    # A string-literal Mid target is a compile error (with/without $ suffix, and MidB).
    assert _CODE in _codes('Sub S\n    Mid$("abcdef", 2, 3) = "XYZ"\nEnd Sub')
    assert _CODE in _codes('Sub S\n    Mid("abcdef", 2) = "zz"\nEnd Sub')
    assert _CODE in _codes('Sub S\n    MidB$("abcdef", 2) = "z"\nEnd Sub')
    # A writable String variable target is valid.
    assert _CODE not in _codes(
        'Sub S\n    Dim s As String\n    s = "abcdef"\n    Mid$(s, 2, 3) = "XYZ"\nEnd Sub'
    )
    # Mid used as a function on the RHS (not a replacement statement) is not flagged.
    assert _CODE not in _codes('Sub S\n    Dim r As String\n    r = Mid$("abcdef", 2, 3)\nEnd Sub')


def test_module_shadowing_mid_suppresses() -> None:
    # If the module declares its own Mid, the intrinsic rule does not apply.
    assert _CODE not in _codes('Sub S\n    Dim Mid As String\n    Mid$("abc", 1) = "z"\nEnd Sub')


def test_oracle_asserted_cases() -> None:
    if asserted_cases(_CODE):
        assert assert_oracle_behavior(_CODE) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        assert _CODE not in case_codes(case), f"{case.id}: {_CODE} false positive"
