"""Inline '@pyvba-ignore suppression directives."""

from __future__ import annotations

from pyvbaanalysis import AnalyzeModuleOptions, analyze_module

# A module body that fires array-declaration-impossible-bounds on the Dim line.
_BOUNDS = "array-declaration-impossible-bounds"


def _codes(source: str, opts: AnalyzeModuleOptions | None = None) -> set[str]:
    return {d.code for d in analyze_module(source, opts)}


def test_ignore_line_suppresses_its_own_line() -> None:
    src = (
        "Option Explicit\nSub S()\n"
        "    Dim a(10 To 1) As Long  '@pyvba-ignore: array-declaration-impossible-bounds\n"
        "End Sub\n"
    )
    assert _BOUNDS not in _codes(src)


def test_ignore_line_without_codes_suppresses_all() -> None:
    src = "Option Explicit\nSub S()\n    Dim a(10 To 1) As Long  '@pyvba-ignore\nEnd Sub\n"
    assert _BOUNDS not in _codes(src)


def test_ignore_next_line() -> None:
    src = (
        "Option Explicit\nSub S()\n"
        "    '@pyvba-ignore-next-line: array-declaration-impossible-bounds\n"
        "    Dim a(10 To 1) As Long\nEnd Sub\n"
    )
    assert _BOUNDS not in _codes(src)


def test_ignore_next_line_does_not_affect_other_lines() -> None:
    # The directive targets only the next line; a violation two lines down still fires.
    src = (
        "Option Explicit\nSub S()\n"
        "    '@pyvba-ignore-next-line\n"
        "    Dim ok As Long\n"
        "    Dim a(10 To 1) As Long\nEnd Sub\n"
    )
    assert _BOUNDS in _codes(src)


def test_ignore_file_before_first_source_line() -> None:
    src = "'@pyvba-ignore-file: option-explicit-missing\nSub S()\nEnd Sub\n"
    assert "option-explicit-missing" not in _codes(src)


def test_inline_suppression_can_be_disabled() -> None:
    src = "Option Explicit\nSub S()\n    Dim a(10 To 1) As Long  '@pyvba-ignore\nEnd Sub\n"
    assert _BOUNDS in _codes(src, AnalyzeModuleOptions(inline_suppression=False))


def test_directive_is_case_insensitive() -> None:
    src = (
        "Option Explicit\nSub S()\n"
        "    Dim a(10 To 1) As Long  '@PyVBA-Ignore: Array-Declaration-Impossible-Bounds\n"
        "End Sub\n"
    )
    assert _BOUNDS not in _codes(src)


def test_unknown_code_warns_and_suppresses_nothing() -> None:
    src = "Option Explicit\nSub S()  '@pyvba-ignore: not-a-real-code\nEnd Sub\n"
    assert "analysis-suppression-directive" in _codes(src)


def test_unknown_directive_verb_warns() -> None:
    src = "Option Explicit\nSub S()  '@pyvba-ignore-bogus\nEnd Sub\n"
    assert "analysis-suppression-directive" in _codes(src)


def test_late_ignore_file_warns_and_does_not_apply() -> None:
    src = "Sub S()\n    '@pyvba-ignore-file: option-explicit-missing\nEnd Sub\n"
    codes = _codes(src)
    assert "analysis-suppression-directive" in codes
    assert "option-explicit-missing" in codes  # the misplaced directive did not suppress


def test_doc_comment_is_not_a_directive() -> None:
    src = "Option Explicit\nSub S()\n    Dim a(10 To 1) As Long  '''@pyvba-ignore\nEnd Sub\n"
    assert _BOUNDS in _codes(src)


def test_rem_comment_is_not_a_directive() -> None:
    src = "Option Explicit\nSub S()\n    Dim a(10 To 1) As Long : Rem @pyvba-ignore\nEnd Sub\n"
    assert _BOUNDS in _codes(src)


def test_directive_text_in_string_is_not_a_directive() -> None:
    src = (
        "Option Explicit\nSub S()\n"
        "    Dim a(10 To 1) As Long\n"
        '    Dim s As String : s = "' + "'" + '@pyvba-ignore"\n'
        "End Sub\n"
    )
    assert _BOUNDS in _codes(src)


def test_directive_diagnostic_is_never_suppressible() -> None:
    # A whole-file catch-all must not hide the warning a malformed directive produces.
    src = "'@pyvba-ignore-file\nSub S()  '@pyvba-ignore: not-a-real-code\nEnd Sub\n"
    assert _codes(src) == {"analysis-suppression-directive"}
