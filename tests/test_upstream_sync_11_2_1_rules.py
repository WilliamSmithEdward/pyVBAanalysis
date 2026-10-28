"""Rule contracts from the XLIDE 11.1.0 -> 11.2.1 sync (rule fixes).

Each section names the upstream commit whose own tests it follows.
"""

import pytest

from pyvbaanalysis import analyze_module, analyze_project
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind


def _project(source: str, *others: tuple[str, ModuleSymbolKind, str], name: str = "Caller") -> list:
    modules = [ModuleInput(name, ModuleSymbolKind.STANDARD, source)]
    modules += [ModuleInput(other, kind, text) for other, kind, text in others]
    return analyze_project(modules)[name]


def _hits(source: str, code: str, options: AnalyzeModuleOptions | None = None) -> list:
    return [d for d in analyze_module(source, options) if d.code == code]


def _starts(source: str, code: str, options: AnalyzeModuleOptions | None = None) -> list[int]:
    return [d.span.start for d in _hits(source, code, options)]


# -- 53e6868d: the Else of a one-line If after a colon -----------------------

_ELSE_TAIL_QUIET = [
    "Sub T()\n    If a Then b = 1: Else c = 2\nEnd Sub\n",
    "Sub T()\n    If a Then b = 1: Else: c = 2\nEnd Sub\n",
    "Sub T()\n    If a Then b = 1: _\n        Else c = 2\nEnd Sub\n",
    "Sub T()\n    If a Then For i = 1 To 2: Next: Else x = 1\n"
    "    If a Then With c: End With: Else d = 1\nEnd Sub\n",
    "Sub T()\n    If a Then Select Case b: Case 1: c = 1: End Select: Else d = 1\nEnd Sub\n",
    "Sub T()\n    If a Then b = 1: Else If x Then d = 2: Else e = 3\nEnd Sub\n",
    "Sub T()\nL1: If a Then b = 1: Else c = 2\n10 If a Then b = 1: Else c = 2\nEnd Sub\n",
    "Sub T()\n#If True Then\n    If a Then b = 1: Else c = 2\n#Else\n    Else\n#End If\nEnd Sub\n",
    "Sub T()\n    If a Then If b Then c = 1 Else d = 2: Else e = 3\nEnd Sub\n",
    "Sub T()\n    If a Then If b Then c = 1: Else d = 2: Else e = 3\nEnd Sub\n",
    "Sub T()\n    If a Then b = 1: If c Then d = 2: Else e = 3\nEnd Sub\n",
    "Sub T()\n    If a Then obj.Else = 1: Else b\nEnd Sub\n",
]


@pytest.mark.parametrize("source", _ELSE_TAIL_QUIET)
def test_else_tail_after_colon_is_accepted(source: str) -> None:
    assert _hits(source, "else-without-if") == []


@pytest.mark.parametrize("source", [
    "Sub T()\n    If x Then\n        If a Then b = 1: Else c = 2\n    Else\n        y = 1\n    End If\nEnd Sub\n",
    "Sub T()\n    If x Then\n        If a Then If b Then c = 1 Else d = 2: Else e = 3\n    End If\nEnd Sub\n",
    "Sub T()\n    If x Then\n        If a Then b Else c: If d Then e: Else f\n    End If\nEnd Sub\n",
])
def test_else_tail_inside_block_if_is_not_a_clause(source: str) -> None:
    assert _hits(source, "else-without-if") == []
    assert _hits(source, "else-branch-order") == []


@pytest.mark.parametrize("source, marker", [
    ("Sub T()\n    If a Then b = 1: Else c = 2: Else d = 3\nEnd Sub\n", "Else d"),
    ("Sub T()\n    If a Then b = 1 Else c = 2: Else d = 3\nEnd Sub\n", "Else d"),
    ("Sub T()\n    If x Then\n        If a Then b = 1: Else c = 2: Else d = 3\n"
     "    Else\n        y = 1\n    End If\nEnd Sub\n", "Else d"),
    ("Sub T()\n    Else\nEnd Sub\n", "Else"),
    ("Sub T()\n    If a Then b = 1:\n    Else c = 2\nEnd Sub\n", "Else"),
    ("Sub T()\n    b = 1: Else c = 2\nEnd Sub\n", "Else"),
    ("Sub T()\n    If a Then b = 1: ElseIf c Then d = 2\nEnd Sub\n", "ElseIf"),
    ("Sub T()\n    If x Then\n        If a Then b = 1: ElseIf c Then d = 2\n    End If\nEnd Sub\n", "ElseIf"),
    ("Sub T()\n    If a Then b Else c: End If: Else d\nEnd Sub\n", "Else d"),
    ("Sub T()\n    If a Then obj.If = 1: Else b: Else c\nEnd Sub\n", "Else c"),
    ("Sub T()\n    If a Then b: Else c: Else d: If e Then f: Else g\nEnd Sub\n", "Else d"),
])
def test_stray_else_after_colon_is_still_reported(source: str, marker: str) -> None:
    assert _starts(source, "else-without-if") == [source.index(marker)]
    assert _hits(source, "else-branch-order") == []


def test_block_if_with_two_else_clauses_still_reported() -> None:
    source = "Sub T()\n    If x Then\n        a = 1\n    Else\n        b = 2\n    Else\n        c = 3\n    End If\nEnd Sub\n"
    assert len(_hits(source, "else-branch-order")) == 1


# -- a2321dcf: runtime readonly properties, host scalar array assignments ----

_FLAGS_FN = "Function MakeFlags() As Boolean()\nDim flags(1) As Boolean\nMakeFlags = flags\nEnd Function\n"


def _runtime_source(line: str) -> str:
    return (
        "Option Explicit\n" + _FLAGS_FN
        + "Function Main() As Variant\nDim values(1) As Boolean\n" + line + "\nMain = 1\nEnd Function"
    )


@pytest.mark.parametrize("target", ["Err.LastDllError", "VBA.Err.LastDllError"])
def test_runtime_readonly_property_is_reported(target: str) -> None:
    source = _runtime_source(target + " = 5")
    for findings in (analyze_module(source), _project(source)):
        found = [d for d in findings if d.code == "readonly-member-assignment"]
        assert len(found) == 1
        assert "Can't assign to read-only property" in found[0].message
        assert source[found[0].span.start:found[0].span.end] == "LastDllError"


@pytest.mark.parametrize("value", ["values", "[values]", "(values)", "MakeFlags()", "(MakeFlags())"])
def test_whole_array_into_host_scalar_setter(value: str) -> None:
    source = _runtime_source("ThisWorkbook.Worksheets(1).EnableCalculation = " + value)
    for findings in (analyze_module(source), _project(source)):
        errors = [d for d in findings if d.code == "assignment-type-mismatch"]
        assert len(errors) == 1
        assert "Run-time error '13'" in errors[0].message
        assert "compile error" not in errors[0].message


@pytest.mark.parametrize("line", [
    "Err.Number = 5", "VBA.Err.Number = 5", "Main = Err.LastDllError",
    "ThisWorkbook.Worksheets(1).EnableCalculation = True",
    "ThisWorkbook.Worksheets(1).EnableCalculation = values(0)",
    'ThisWorkbook.Worksheets(1).EnableCalculation = "True"',
])
def test_valid_runtime_and_host_scalar_use_is_quiet(line: str) -> None:
    assert [d for d in analyze_module(_runtime_source(line)) if d.severity == "error"] == []


@pytest.mark.parametrize("receiver", ["ThisWorkbook.Sheets(1)", "unknown"])
def test_ambiguous_receiver_assumes_no_scalar_setter(receiver: str) -> None:
    source = _runtime_source(receiver + ".EnableCalculation = values").replace(
        "Dim values(1) As Boolean", "Dim values(1) As Boolean\nDim unknown As Object"
    )
    assert _hits(source, "assignment-type-mismatch") == []


def test_shadowed_err_is_not_the_runtime_object() -> None:
    assert _hits("Sub T(ByVal Err As Object)\nErr.LastDllError = 5\nEnd Sub", "readonly-member-assignment") == []


@pytest.mark.parametrize("line", [
    "With ThisWorkbook.Worksheets(1)\n.EnableCalculation = values\nEnd With",
    "If True Then ThisWorkbook.Worksheets(1).EnableCalculation = MakeFlags()",
])
def test_host_scalar_setter_in_with_and_single_line_if(line: str) -> None:
    assert len(_hits(_runtime_source(line), "assignment-type-mismatch")) == 1


# -- e937c742, 31ff586f: reserved words as Type member names -----------------

_RESERVED_MEMBER_WORDS = [
    "Next", "Loop", "Wend", "For", "Do", "With", "Sub", "Function", "Select", "If",
    "End", "Else", "Case", "Set", "Dim", "Private", "Len", "True", "Date", "String",
    "Long", "Variant", "Local", "Attribute", "New",
]


def _reserved_member_type(words: list[str]) -> str:
    return (
        "Private Type Rec\n"
        + "".join(f"    {word} As Long\n" for word in words)
        + "    While _\n        As Long\n"
        + "    Loop As Long ' comment\n"
        + "    Next As [Long]\n"
        + "    Next As New Collection\n"
        + "    With As New [Collection]\n"
        + "#If Win64 Then\n    Next As LongLong\n#Else\n    Next As Long\n#End If\n"
        + "    Type As [Long]\n"
        + "    Type(1 To 2) As Long\n"
        + "End Type\n"
    )


def test_reserved_word_type_member_with_as_clause_is_accepted() -> None:
    findings = analyze_module(_reserved_member_type(_RESERVED_MEMBER_WORDS))
    for code in ("invalid-declaration-name", "unexpected-declaration-token", "type-member-without-type"):
        assert [d for d in findings if d.code == code] == []


def test_reserved_array_suffixed_and_incomplete_type_members_still_rejected() -> None:
    source = (
        "Private Type Rec\n    For(1 To 3) As Long\n    With As\n    Loop As ' comment\n"
        "    Next&\n    Next& As Long\n    Set As New\n    Set As New []\n    Set As []\n"
        "    Set As 123\n    Me As Long\n    Next\nEnd Type\n"
    )
    hits = _hits(source, "invalid-declaration-name")
    assert [source[d.span.start:d.span.end] for d in hits] == [
        "For", "With", "Loop", "Next", "Next", "Set", "Set", "Set", "Set", "Me", "Next",
    ]
    enum_source = "Private Enum Mode\n    Next\n    Loop = 2\nEnd Enum\n"
    hits = _hits(enum_source, "invalid-declaration-name")
    assert [enum_source[d.span.start:d.span.end] for d in hits] == ["Next", "Loop"]
