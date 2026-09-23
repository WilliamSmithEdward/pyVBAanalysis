"""Dead-code rule family (deadCode.ts parity): variables nothing uses or nothing
reads, private procedures nothing calls, and statements no path reaches.

Every expectation below is what XLIDE 10.5.0 reports on the same source; the port
was diffed against it case by case. These are information-level findings, so the
compile-accepted corpus cases do not constrain them.
"""

from __future__ import annotations

import pytest

from pyvbaanalysis import analyze_module, analyze_project
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, VbaDiagnostic
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind

_DEAD_CODES = ("unused-variable", "variable-never-read", "unused-procedure", "unreachable-code")


def _findings(source: str, kind: ModuleSymbolKind = ModuleSymbolKind.STANDARD) -> list[str]:
    return [
        f"{d.code}: {d.message}"
        for d in analyze_module(source, AnalyzeModuleOptions(module_kind=kind))
        if d.code in _DEAD_CODES
    ]


# -- variables -------------------------------------------------------------


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        pytest.param(
            "Option Explicit\nSub S()\n    Dim a As Long\nEnd Sub\n",
            ["unused-variable: Local variable 'a' is declared but never used."],
            id="unused local",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    Dim a As Long\n    a = 5\nEnd Sub\n",
            ["variable-never-read: Variable 'a' is assigned but its value is never read."],
            id="assigned, never read",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    Dim a As Long, b As Long, c As Long\n    Debug.Print a + c\nEnd Sub\n",
            ["unused-variable: Local variable 'b' is declared but never used."],
            id="middle of a Dim list",
        ),
        pytest.param(
            "Option Explicit\nPrivate mCount As Long\nSub S()\nEnd Sub\n",
            ["unused-variable: Module-level variable 'mCount' is declared but never used."],
            id="private module variable",
        ),
        pytest.param(
            "Option Explicit\nConst LIMIT = 5\n",
            ["unused-variable: Constant 'LIMIT' is declared but never used."],
            id="bare module Const is private",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    Const K As Long = 2\nEnd Sub\n",
            ["unused-variable: Constant 'K' is declared but never used."],
            id="local Const",
        ),
        pytest.param(
            "Option Explicit\nPrivate mVal As Long\nSub S()\n    Dim mVal As Long\n    Debug.Print mVal\nEnd Sub\n",
            ["unused-variable: Module-level variable 'mVal' is declared but never used."],
            id="a local shadows the module variable",
        ),
        pytest.param(
            "Option Explicit\nPrivate Name As String\nSub S()\n    Debug.Print Application.Name\nEnd Sub\n",
            ["unused-variable: Module-level variable 'Name' is declared but never used."],
            id="a member name is not a mention",
        ),
        pytest.param(
            'Option Explicit\nPrivate Text As String\nSub S()\n    MsgBox Prompt:="x"\nEnd Sub\n'
            "Sub T(Text As String)\nEnd Sub\n",
            ["unused-variable: Module-level variable 'Text' is declared but never used."],
            id="a named argument is not a mention",
        ),
    ],
)
def test_variables_nothing_uses_or_reads(source: str, expected: list[str]) -> None:
    assert _findings(source) == expected


@pytest.mark.parametrize(
    "source",
    [
        pytest.param(
            "Option Explicit\nSub S()\n    Dim a As Long\n    a = 5\n    Debug.Print a\nEnd Sub\n",
            id="read",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    Dim x As Long\n    x = x + 1\nEnd Sub\n",
            id="x = x + 1 reads x",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    Dim i As Long\n    For i = 1 To 3\n    Next i\nEnd Sub\n",
            id="a For counter is read by the loop",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    Dim a(3) As Long\n    a(1) = 2\nEnd Sub\n",
            id="storing into an array reads the array",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    Dim a As Long\n    Fill a\nEnd Sub\nSub Fill(ByRef v As Long)\n    v = 1\nEnd Sub\n",
            id="passed to a procedure",
        ),
        pytest.param(
            "Option Explicit\nPublic gCount As Long\n", id="public variables are not tracked"
        ),
        pytest.param(
            "Option Explicit\nPrivate mHidden As Long\nAttribute mHidden.VB_VarUserMemId = 0\n",
            id="an attributed variable is not tracked",
        ),
    ],
)
def test_variables_in_use_are_silent(source: str) -> None:
    assert _findings(source) == []


# -- private procedures ----------------------------------------------------


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        pytest.param(
            "Option Explicit\nPrivate Sub Helper()\nEnd Sub\n",
            ["unused-procedure: Private Sub 'Helper' is never called."],
            id="never called",
        ),
        pytest.param(
            "Option Explicit\nPrivate Sub Helper()\n    Helper\nEnd Sub\n",
            ["unused-procedure: Private Sub 'Helper' is never called."],
            id="only calls itself",
        ),
        pytest.param(
            "Option Explicit\nPrivate Function Calc() As Long\n    Calc = 1\nEnd Function\n",
            ["unused-procedure: Private Function 'Calc' is never called."],
            id="function",
        ),
        pytest.param(
            "Option Explicit\nPrivate Property Get Size() As Long\n    Size = 1\nEnd Property\n",
            ["unused-procedure: Private Property 'Size' is never used."],
            id="property",
        ),
    ],
)
def test_private_procedures_nothing_calls(source: str, expected: list[str]) -> None:
    assert _findings(source) == expected


@pytest.mark.parametrize(
    ("source", "kind"),
    [
        pytest.param(
            "Option Explicit\nPrivate Sub Helper()\nEnd Sub\nPublic Sub Go()\n    Helper\nEnd Sub\n",
            ModuleSymbolKind.STANDARD,
            id="called",
        ),
        pytest.param(
            'Option Explicit\nPrivate Sub Poll()\nEnd Sub\nPublic Sub Go()\n    Application.OnTime Now, "Poll"\nEnd Sub\n',
            ModuleSymbolKind.STANDARD,
            id="named in a string",
        ),
        pytest.param(
            "Option Explicit\nPrivate Sub Class_Initialize()\nEnd Sub\n",
            ModuleSymbolKind.CLASS,
            id="class event handler",
        ),
        pytest.param(
            "Option Explicit\nPrivate Sub Auto_Open()\nEnd Sub\n",
            ModuleSymbolKind.STANDARD,
            id="Auto_Open",
        ),
        pytest.param(
            "Option Explicit\nPublic Sub Lonely()\nEnd Sub\n",
            ModuleSymbolKind.STANDARD,
            id="public procedures are not tracked",
        ),
    ],
)
def test_private_procedures_reached_some_other_way_are_silent(
    source: str, kind: ModuleSymbolKind
) -> None:
    assert _findings(source, kind) == []


def _project_findings(starter: str) -> dict[str, list[str]]:
    modules = [
        ModuleInput(
            "Timers", ModuleSymbolKind.STANDARD, "Option Explicit\nPrivate Sub Poll()\nEnd Sub\n"
        ),
        ModuleInput("Starter", ModuleSymbolKind.STANDARD, starter),
    ]
    return {
        name: [d.code for d in diagnostics if d.code in _DEAD_CODES]
        for name, diagnostics in analyze_project(modules).items()
    }


def test_a_string_in_another_module_can_name_a_private_procedure() -> None:
    """`Application.Run "Poll"` reaches a Private Sub in a standard module, so a
    name in any module's string literal counts as a use."""
    runs = 'Option Explicit\nPublic Sub Go()\n    Application.Run "Poll"\nEnd Sub\n'
    assert _project_findings(runs) == {"Timers": [], "Starter": []}
    qualified = (
        'Option Explicit\nPublic Sub Go()\n    Application.OnTime Now, "Timers.Poll"\nEnd Sub\n'
    )
    assert _project_findings(qualified) == {"Timers": [], "Starter": []}


def test_only_a_whole_word_in_a_string_counts() -> None:
    longer = 'Option Explicit\nPublic Sub Go()\n    Application.Run "Polling"\nEnd Sub\n'
    assert _project_findings(longer) == {"Timers": ["unused-procedure"], "Starter": []}


# -- unreachable code ------------------------------------------------------


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        pytest.param(
            "Option Explicit\nSub S()\n    Exit Sub\n    Debug.Print 1\n    Debug.Print 2\nEnd Sub\n",
            ["unreachable-code: Unreachable code after 'Exit Sub'."],
            id="after Exit Sub",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    GoTo Done\n    Debug.Print 1\nDone:\nEnd Sub\n",
            ["unreachable-code: Unreachable code after 'GoTo Done'."],
            id="between GoTo and its label",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    End\n    Debug.Print 1\nEnd Sub\n",
            ["unreachable-code: Unreachable code after 'End'."],
            id="after End",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    Dim i As Long\n    For i = 1 To 3\n        If i = 2 Then\n"
            "            Exit For\n            Debug.Print i\n        End If\n    Next i\nEnd Sub\n",
            ["unreachable-code: Unreachable code after 'Exit For'."],
            id="inside a nested block",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    On Error GoTo H\n    Exit Sub\nH:\n    Resume Next\n    Debug.Print 1\nEnd Sub\n",
            ["unreachable-code: Unreachable code after 'Resume Next'."],
            id="after Resume Next",
        ),
    ],
)
def test_statements_no_path_reaches(source: str, expected: list[str]) -> None:
    assert _findings(source) == expected


@pytest.mark.parametrize(
    "source",
    [
        pytest.param(
            "Option Explicit\nSub S()\n    Exit Sub\nHandler:\n    Debug.Print 1\nEnd Sub\n",
            id="a label is a landing",
        ),
        pytest.param(
            "Option Explicit\nSub S(ByVal v As Long)\n    Select Case v\n        Case 1\n            Exit Sub\n"
            "        Case 2\n            Debug.Print 2\n    End Select\nEnd Sub\n",
            id="a Case arm is a landing",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    Exit Sub\n    If True Then\nL1:\n        Debug.Print 1\n    End If\nEnd Sub\n",
            id="a block holding a label",
        ),
    ],
)
def test_reachable_statements_are_silent(source: str) -> None:
    assert _findings(source) == []


# -- the removal edits -----------------------------------------------------


def _apply_removal(source: str, diagnostic: VbaDiagnostic) -> str | None:
    data = diagnostic.data
    if data is None:
        return None
    holder = data.remove_declaration or data.remove_unreachable_code
    if holder is None:
        return None
    edit = holder.edit
    return source[: edit.span.start] + edit.new_text + source[edit.span.end :]


def _removed(source: str, code: str) -> str | None:
    return _apply_removal(source, next(d for d in analyze_module(source) if d.code == code))


@pytest.mark.parametrize(
    ("source", "after"),
    [
        pytest.param(
            "Option Explicit\nSub S()\n    Dim a As Long\n    Debug.Print 1\nEnd Sub\n",
            "Option Explicit\nSub S()\n    Debug.Print 1\nEnd Sub\n",
            id="the whole line",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    Dim a As Long, b As Long\n    Debug.Print b\nEnd Sub\n",
            "Option Explicit\nSub S()\n    Dim b As Long\n    Debug.Print b\nEnd Sub\n",
            id="first of a list",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    Dim a As Long, b As Long\n    Debug.Print a\nEnd Sub\n",
            "Option Explicit\nSub S()\n    Dim a As Long\n    Debug.Print a\nEnd Sub\n",
            id="last of a list",
        ),
        pytest.param(
            "Option Explicit\n''' The count.\nPrivate mCount As Long\nSub S()\nEnd Sub\n",
            "Option Explicit\nSub S()\nEnd Sub\n",
            id="with its doc comment",
        ),
        pytest.param(
            "Option Explicit\nSub S()\n    Dim a As Long ' spare\nEnd Sub\n",
            "Option Explicit\nSub S()\nEnd Sub\n",
            id="with a trailing comment",
        ),
        pytest.param(
            "Option Explicit\r\nSub S()\r\n    Dim a As Long\r\nEnd Sub\r\n",
            "Option Explicit\r\nSub S()\r\nEnd Sub\r\n",
            id="CRLF",
        ),
    ],
)
def test_removing_an_unused_declaration(source: str, after: str) -> None:
    assert _removed(source, "unused-variable") == after


def test_a_declaration_sharing_its_line_offers_no_removal() -> None:
    source = "Option Explicit\nSub S()\n    Dim a As Long: Debug.Print 1\nEnd Sub\n"
    assert _removed(source, "unused-variable") is None


def test_removing_unreachable_code_takes_the_whole_run() -> None:
    source = (
        "Option Explicit\nSub S()\n    Exit Sub\n    Debug.Print 1\n    Debug.Print 2\nEnd Sub\n"
    )
    assert (
        _removed(source, "unreachable-code") == "Option Explicit\nSub S()\n    Exit Sub\nEnd Sub\n"
    )
