"""Rule contracts from the XLIDE 11.2.1 -> 11.2.11 sync (declarations and
statement shapes).

Each section names the upstream commit whose own tests it follows. The measured
cases are the oracle fixtures of the vendored corpus, analyzed as upstream's tests
analyze them: every module of the case in one project, in the case's host.
"""

from __future__ import annotations

import pytest

from pyvbaanalysis import analyze_module, analyze_module_options_for
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, diagnostic_metadata_for_code
from pyvbaanalysis.diagnostics.model import DiagnosticEvidenceKind, DiagnosticSeverity, VbaDiagnostic
from pyvbaanalysis.evidence import OracleCase
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex
from oracle_support import _RAW_MODULES, _HOSTS, CASES, oracle_module_input  # type: ignore[attr-defined]


def _measured(prefix: str) -> list[OracleCase]:
    return [case for case in CASES.values() if case.id.startswith(prefix)]


def _by_code(source: str, code: str, host: str = "excel") -> list[VbaDiagnostic]:
    return [d for d in analyze_module(source, AnalyzeModuleOptions(host=host)) if d.code == code]


def _compile_errors(source: str, host: str = "excel") -> list[str]:
    out = []
    for d in analyze_module(source, AnalyzeModuleOptions(host=host)):
        meta = diagnostic_metadata_for_code(d.code)
        if meta is not None and meta.diagnostic_kind == DiagnosticEvidenceKind.COMPILE_ERROR:
            out.append(f"{d.code}: {d.message}")
    return out


def _span_text(source: str, d: VbaDiagnostic) -> str:
    return source[d.span.start : d.span.end]


def _procedure(*lines: str) -> str:
    """Upstream's procedure(): a Main with the usual locals around the lines."""
    return "\r\n".join([
        "Option Explicit",
        "Sub Main()",
        "    Dim v As Variant, n As Long, o As Object, c As New Collection",
        *("    " + line for line in lines),
        "End Sub",
        "",
    ])


def _project_errors(inputs: list[ModuleInput], host: str | None = None) -> list[str]:
    """Upstream's errors(): every error of every module, as `Module: code: message`."""
    index = ProjectIndex()
    for module_input in inputs:
        index.set_module(module_input)
    out: list[str] = []
    for module_input in inputs:
        opts = analyze_module_options_for(
            index,
            module_input.module_name,
            module_input.module_kind,
            host=host,
            conditional_compilation=ConditionalCompilationEnvironment(project_constants={}),
        )
        out.extend(
            f"{module_input.module_name}: {diag.code}: {diag.message}"
            for diag in analyze_module(module_input.source, opts)
            if diag.severity == DiagnosticSeverity.ERROR
        )
    return out


def _case_errors(case: OracleCase) -> list[str]:
    inputs = [oracle_module_input(raw) for raw in _RAW_MODULES[case.id]]
    return _project_errors(inputs, _HOSTS.get(case.id))


def _module_errors(source: str) -> list[str]:
    return _project_errors([ModuleInput("Module1", ModuleSymbolKind.STANDARD, source)])


# -- 5cad8c3e: event sources and standard modules as types -------------------

_ISSUE_1530 = _measured("issue1530_")


def test_issue_1530_has_every_measured_case() -> None:
    assert len(_ISSUE_1530) == 108


@pytest.mark.parametrize("case", _ISSUE_1530, ids=lambda case: case.id)
def test_issue_1530_agrees_with_the_vbe(case: OracleCase) -> None:
    found = _case_errors(case)
    if case.expected == "accepted":
        assert found == []
    else:
        # One report, with the VBE's own reason.
        assert len(found) == 1
        reason = "A module is not a valid type" if "_module_type_" in case.id else "does not source automation events"
        assert reason in found[0]


def test_names_the_module_a_with_events_variable_or_a_dim_is_declared_as() -> None:
    found = _project_errors([
        ModuleInput("Helpers", ModuleSymbolKind.STANDARD, "Option Explicit\r\nSub P()\r\nEnd Sub\r\n"),
        ModuleInput("Holder", ModuleSymbolKind.CLASS, "Option Explicit\r\nPrivate WithEvents h As Helpers\r\nPrivate m As Helpers\r\n"),
    ])
    assert len(found) == 2
    assert "WithEvents variable 'h' is declared As Helpers, and Helpers is a standard module of this project." in found[0]
    assert "'Helpers' is a standard module of this project, and a module is not a type." in found[1]


def test_a_userform_whose_code_declares_an_event_is_a_valid_source() -> None:
    found = _project_errors([
        ModuleInput("Picker", ModuleSymbolKind.USERFORM, "Option Explicit\r\nPublic Event Picked(ByVal index As Long)\r\n"),
        ModuleInput(
            "Holder",
            ModuleSymbolKind.CLASS,
            "Option Explicit\r\nPrivate WithEvents p As Picker\r\nPrivate Sub p_Picked(ByVal index As Long)\r\n"
            "    Debug.Print index\r\nEnd Sub\r\n",
        ),
    ])
    assert found == []


# -- 25c82b6d: one error for a Const As a standard module -------------------

_ISSUE_1553 = _measured("issue1553_")


def test_issue_1553_has_every_measured_case() -> None:
    # Ten shapes refused and two accepted, in each of the four hosts.
    assert len(_ISSUE_1553) == 48


@pytest.mark.parametrize("case", _ISSUE_1553, ids=lambda case: case.id)
def test_issue_1553_agrees_with_the_vbe(case: OracleCase) -> None:
    found = [line.split(": ", 1)[1] for line in _case_errors(case)]
    if case.expected == "accepted":
        assert found == []
        return
    # One report, with the VBE's own reason.
    assert len(found) == 1
    assert found[0].startswith("const-invalid-type: ")
    if "object" in case.id:
        reason = "Invalid data type for constant"
    else:
        reason = "Syntax error" if "procedure_" in case.id else "Expected: type name"
    assert reason in found[0]


# -- 091fcd77: a block header without one of its parts -----------------------

_BLOCK_HEADERS = _measured("block_header_")


def test_block_headers_cover_every_host() -> None:
    assert {_HOSTS.get(case.id, "excel") for case in _BLOCK_HEADERS} == {"excel", "word", "powerpoint", "access"}


@pytest.mark.parametrize("case", _BLOCK_HEADERS, ids=lambda case: case.id)
def test_block_header_agrees_with_the_vbe(case: OracleCase) -> None:
    # `10 Case Is > o.To` is refused for the line number before the first Case,
    # which another check reports; its member o.To is no To of a range.
    source = case.modules[0].source
    host = _HOSTS.get(case.id, "excel")
    assert (len(_compile_errors(source, host)) > 0) == (case.expected == "rejected")
    own = 1 if case.expected == "rejected" and "_numbered_case_member_to_" not in case.id else 0
    assert len(_by_code(source, "malformed-statement", host)) == own


@pytest.mark.parametrize("lines, span, message", [
    (["Select Case n", "Case Is 5", "End Select"], "Is 5", "Case Is takes a comparison operator before its value"),
    (["Select Case n", "Case 1, Is", "End Select"], "Is", "Case Is takes a comparison operator"),
    (["Do Until", "Loop"], "Do Until", "Do Until needs a condition."),
    (["Do", "Loop While"], "Loop While", "Loop While needs a condition."),
    (["While", "Wend"], "While", "While needs a condition"),
    (["For Each v", "Next"], "For Each v", "needs In and what it walks"),
    (["For Each v In", "Next"], "For Each v In", "needs the collection or array it walks"),
    (["For Each In c", "Next"], "For Each In", "its loop variable before In"),
    (["For n = 1", "Next"], "For n = 1", "needs To and its end value"),
    (["For n = 1 To", "Next"], "For n = 1 To", "needs the loop's end value"),
    (["For n = 1 To 2 Step", "Next"], "For n = 1 To 2 Step", "Step needs a value"),
])
def test_block_header_reports(lines: list[str], span: str, message: str) -> None:
    source = _procedure(*lines)
    found = _by_code(source, "malformed-statement")
    assert len(found) == 1
    assert _span_text(source, found[0]) == span
    assert message in found[0].message


@pytest.mark.parametrize("lines", [
    ["10 For n = 1", "Next"], ["Select Case n", "10 Case Is 5", "End Select"], ["10 Do Until", "Loop"],
    ["10 While", "Wend"], ["L: For Each v", "Next"],
])
def test_block_header_after_a_line_number_or_a_label(lines: list[str]) -> None:
    assert len(_by_code(_procedure(*lines), "malformed-statement")) == 1


def test_block_header_member_named_like_a_keyword() -> None:
    assert len(_by_code(_procedure("For n = o.To To", "Next"), "malformed-statement")) == 1
    for lines in [
        ["10 Do", "20 Loop While o.Until"], ["For n = o.To To 3", "Next"], ["For Each v In o.In", "Next"],
        ["Do Until o!While", "Loop"], ["Select Case n", "10 Case Is > o.To", "End Select"],
        ["Select Case n", "Case Is > o!To", "End Select"],
    ]:
        assert _by_code(_procedure(*lines), "malformed-statement") == [], lines


@pytest.mark.parametrize("lines", [
    ["Select Case n", "Case Is = 1, Is <> 2, Is < 3, Is > 4, Is <= 5, Is >= 6", "End Select"],
    ["Select Case n", "Case 1 To 3, Is > 7: Beep", "End Select"],
    ["Select Case n", "Case Is _", "    > 1", "End Select"],
    ["Do Until n > 3", "    n = n + 1", "Loop"], ["Do", "    n = n + 1", "Loop Until (n > 3)"], ["Do: Loop While n < 0"],
    ["While (n < 3)", "    n = n + 1", "Wend"],
    ["For Each v In c", "Next"], ["For Each v In Array(1, 2)", "Next v"], ["For Each v _", "    In c", "Next"],
    ["For n = 1 To 3", "Next"], ["For n = LBound(Array(1)) To UBound(Array(1, 2)) Step -1", "Next n"],
    ["For n = 1 To _", "    3", "Next"],
    ["10 Do While n < 3", "    n = n + 1", "20 Loop"], ["10 For n = 1 To 2", "20 Next"],
])
def test_block_headers_that_compile_are_left_alone(lines: list[str]) -> None:
    assert _compile_errors(_procedure(*lines)) == []


def test_block_header_in_an_inactive_arm_is_left_alone() -> None:
    source = _procedure("#If False Then", "For Each v", "Next", "Do Until", "Loop", "#End If")
    assert _by_code(source, "malformed-statement") == []


# -- 1b98a0e3: End, Stop, Error and On statements in a shape the VBE refuses --

_STATEMENT_SHAPES = _measured("statement_shape_")


def _shape_procedure(*lines: str) -> str:
    """Upstream statementShapes.test.ts procedure(): the lines as given."""
    return "\r\n".join(["Option Explicit", "Sub Main()", "    Dim v As Variant, n As Long", *lines, "End Sub", ""])


def _shapes(source: str, host: str = "excel") -> list[VbaDiagnostic]:
    return _by_code(source, "malformed-statement", host)


def test_statement_shapes_cover_every_host() -> None:
    assert {_HOSTS.get(case.id, "excel") for case in _STATEMENT_SHAPES} == {"excel", "word", "powerpoint", "access"}


@pytest.mark.parametrize("case", _STATEMENT_SHAPES, ids=lambda case: case.id)
def test_statement_shape_agrees_with_the_vbe(case: OracleCase) -> None:
    expected = 1 if case.expected == "rejected" else 0
    assert len(_shapes(case.modules[0].source, _HOSTS.get(case.id, "excel"))) == expected


@pytest.mark.parametrize("line, span, message", [
    ("    End 1", "End 1", "'End' stops the program and takes nothing after it, so '1' cannot follow it."),
    ("    If v Then End Else Beep", "End Else", "so 'Else' cannot follow it"),
    ("    End While", "End While", "a While loop ends at Wend"),
    ("    Stop n", "Stop n", "'Stop' takes no argument."),
    ("    Error", "Error", "needs an error number"),
    ("    Error 5, 6", "Error 5, 6", "takes one error number"),
    ("    On Error GoTo -2", "On Error GoTo -2", "one line label, line number, 0 or -1"),
    ("    On Error Resume 0", "On Error Resume 0", "takes only Next"),
    ("    On Error Resume Next 1", "On Error Resume Next 1", "and nothing after it"),
    ("    On Error GoTo 0 1", "On Error GoTo 0 1", "and nothing after it"),
    ("    On n GoSub A, , B\r\nA:\r\nB:", "GoSub A, , B", "in each place of its list"),
])
def test_statement_shape_reports(line: str, span: str, message: str) -> None:
    source = _shape_procedure(line)
    found = _shapes(source)
    assert len(found) == 1
    assert _span_text(source, found[0]) == span
    assert message in found[0].message


@pytest.mark.parametrize("line", [
    "    Beep: Stop 1", "L: Error", "10 End 1", "    If v Then If n Then Stop 1",
    "    If v Then Beep Else On Error Resume", "    Beep: Error: Beep",
    "    If v Then Beep: Error Else Beep", "    If v Then Beep: Beep Else Stop 1",
    "    If v Then Beep: Beep Else End 1", "    If v Then Beep: End Else Beep", "    If v Then Beep Else Beep: Stop 1",
    "    If v Then\r\n        Beep\r\n    Else If n Then Stop 1\r\n    End If",
    "    If v Then\r\n        Beep\r\n    Else Stop 1\r\n    End If",
    "    If v Then\r\n        Beep\r\n    Else Error\r\n    End If",
    "    If v Then\r\n        Beep\r\n    Else On Error Resume Next 1\r\n    End If",
])
def test_statement_shape_after_a_colon_a_label_a_line_number_and_in_one_line_ifs(line: str) -> None:
    assert len(_shapes(_shape_procedure(line))) == 1


def test_statement_shape_on_an_else_if_line_a_procedure_line_and_before_a_block_if() -> None:
    for source in [
        _shape_procedure("    If False Then", "        Beep", "    ElseIf True Then Stop 1", "    End If"),
        _shape_procedure("    If False Then", "        Beep", "    ElseIf True Then Error", "    End If"),
        "Option Explicit\r\nSub Other(): Stop 1: End Sub\r\n",
        _shape_procedure("    If True Then Stop 1: If True Then", "        Beep", "    End If"),
    ]:
        assert len(_shapes(source)) == 1, source
    for source in [
        _shape_procedure("    If False Then", "        Beep", "    ElseIf True Then Stop", "    End If"),
        _shape_procedure("    If False Then", "        Beep", "    ElseIf True Then", "        Stop", "    End If"),
        "Option Explicit\r\nSub Other(): Stop: End Sub\r\n",
    ]:
        assert _compile_errors(source) == [], source


@pytest.mark.parametrize("line", [
    "    On Error GoTo Next 1", "    On Error Resume Stop", "    On Error Resume If",
    "    If v Then On Error Resume Else Beep", "    If v Then Beep: On Error Resume Else Beep",
    "    If v Then On Error GoTo Else Beep",
])
def test_statement_shape_leaves_a_reserved_jump_target_to_its_check(line: str) -> None:
    found = _shapes(_shape_procedure(line))
    assert len(found) == 1
    assert "cannot name a line label" in found[0].message


@pytest.mark.parametrize("line", [
    "    End", "    End ' note", "    If v Then End", "    If v Then Beep Else End", "    If v Then End: Beep",
    "    Stop", "    If v Then Stop Else Beep", "Stop: Beep",
    "    Error 5", "    Error (5)", "    If v Then Error 5 Else Beep", "Error: Beep", "10 Error: Beep",
    "    Error _\r\n        5", "    Error$ 5",
    "    v = Error", "    v = Error(5)", "    Debug.Print Error",
    "    On Error GoTo 0", "    On Error GoTo -1", "    On Error GoTo - 1", "    On Error GoTo L\r\nL:",
    "    On Error GoTo 10\r\n10 Beep",
    "    On Error Resume Next", "    On Local Error Resume Next", "    If v Then On Error Resume Next Else Beep",
    "    On n GoTo A, B\r\nA:\r\nB:", "    On n GoSub A\r\n    Exit Sub\r\nA:\r\n    Return",
    "    Dim r As Object\r\n    Set r = Nothing\r\n    If v Then r.End(1).Select",
    "    With Nothing\r\n        .End 1\r\n        .Stop 1\r\n    End With",
    "    If v Then Beep: Stop Else Beep", "    If v Then\r\n        End\r\n    End If",
    "    If v Then\r\n        Beep\r\n    Else If n Then Stop Else Beep\r\n    End If",
    "    If v Then\r\n        Beep\r\n    Else Stop\r\n    End If",
])
def test_statement_shapes_that_compile_are_left_alone(line: str) -> None:
    assert _compile_errors(_shape_procedure(line)) == []


def test_statement_shape_in_an_inactive_arm_is_left_alone() -> None:
    source = _shape_procedure("#If False Then", "    End 1", "    Stop 1", "    Error", "#End If")
    assert _shapes(source) == []
