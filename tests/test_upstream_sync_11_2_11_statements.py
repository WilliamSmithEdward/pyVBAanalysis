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


# -- 8a6f7172: file statements in a shape the VBE refuses --------------------

_FILE_STATEMENT_SHAPES = _measured("file_statement_shape_")

_MODULE_KINDS = {
    "class": ModuleSymbolKind.CLASS,
    "userform": ModuleSymbolKind.USERFORM,
    "document": ModuleSymbolKind.DOCUMENT,
}


def _file_procedure(*lines: str) -> str:
    """Upstream fileStatementShapes.test.ts procedure()."""
    return "\r\n".join([
        "Option Explicit",
        "Sub Main()",
        "    Dim f As Integer, s As String, n As Long, v As Variant, a(2) As Variant",
        "    f = FreeFile",
        *("    " + line for line in lines),
        "End Sub",
        "",
    ])


def _kind_shapes(source: str, kind: ModuleSymbolKind | None) -> list[VbaDiagnostic]:
    options = AnalyzeModuleOptions(host="excel", module_kind=kind)
    return [d for d in analyze_module(source, options) if d.code == "malformed-statement"]


def test_file_statement_shapes_cover_every_host() -> None:
    assert {_HOSTS.get(case.id, "excel") for case in _FILE_STATEMENT_SHAPES} == {
        "excel", "word", "powerpoint", "access"
    }


@pytest.mark.parametrize("case", _FILE_STATEMENT_SHAPES, ids=lambda case: case.id)
def test_file_statement_shape_agrees_with_the_vbe(case: OracleCase) -> None:
    # Each module as its own kind: a class's or a form's statements are read too.
    host = _HOSTS.get(case.id, "excel")
    found = [
        d
        for module in case.modules
        for d in analyze_module(module.source, AnalyzeModuleOptions(host=host, module_kind=_MODULE_KINDS.get(module.module_type)))
        if d.code == "malformed-statement"
    ]
    assert len(found) == (1 if case.expected == "rejected" else 0)


@pytest.mark.parametrize("line, span, message", [
    ('Open "x" For Append Access Read As #f', "Access Read", "For Append takes Access Write or Read Write."),
    ('Open "x" For Input Access Read Write As #f', "Access Read Write", "For Input takes Access Read."),
    ('Open "x" For Output Access Read As #f', "Access Read", "For Output takes Access Write."),
    ('Open "x" For Append Access As #f', "Access", "Access takes Read, Write or Read Write."),
    ("Width f, 80", "Width f, 80", "Width names its file number with #"),
    ("Width #f, 80, 1", "Width #f, 80, 1", "Width takes a file number and a width"),
    ("Print #f", "Print #f", "needs a comma after the file number"),
    ("Print #f; 1", "Print #f;", "not a semicolon"),
    ('Input #f, s & "x"', 'Input #f, s & "x"', "reads into one or more variables"),
    ('Line Input #f, "x"', 'Line Input #f, "x"', "reads into one variable"),
    ("Put #f, 1", "Put #f, 1", "a record and a value"),
    ("Get #f, 1", "Get #f, 1", "a record and a variable"),
    ("Get #f, , 1", "1", "Variable required - can't assign to this expression"),
    ('Name "a" As', 'Name "a" As', "Name renames with As"),
])
def test_file_statement_shape_reports(line: str, span: str, message: str) -> None:
    source = _file_procedure(line)
    found = _by_code(source, "malformed-statement")
    assert len(found) == 1
    assert _span_text(source, found[0]) == span
    assert message in found[0].message


def test_file_statement_shape_reads_variables_paths_and_parts() -> None:
    for line in [
        "Input #f, s And s", "Input #f, s s", 'Name ("a") "b"', "Name (s) As", 'Open "x" Access As #f', "Width #, 80",
        "Get #f, , v, v", "Put #f, , v, v", 'Name "a" = "b"', 'Name "a" = "b" As', "Get #f, , -1", "Get #f, , 1 + 2",
        'Get #f, , s & "x"',
    ]:
        assert len(_by_code(_file_procedure(line), "malformed-statement")) == 1, line
    for line in ["Input #f, s$", "Line Input #f, s$", "Input #f, a(1), v", 'Name ("a") As "b"', 'Open "x" Access Read As #f']:
        source = _file_procedure("Dim b(2) As String", "If False Then " + line, "If False Then Input #f, b$(1)")
        assert _by_code(source, "malformed-statement") == [], line


def test_file_statement_shape_follows_no_state_through_a_refused_statement() -> None:
    opened = _file_procedure('Open "x" For Input As #f', 'Name "x" As', "Close #f")
    codes = [d.code for d in analyze_module(opened, AnalyzeModuleOptions(host="excel"))]
    assert "file-already-open" not in codes
    assert "file-path-open" not in codes
    assert len(_by_code(opened, "malformed-statement")) == 1


@pytest.mark.parametrize("line", [
    "If False Then Print #f", "If False Then Beep Else Write #f", "Beep: Input #f", '10 Name "a" "b"',
])
def test_file_statement_shape_in_a_one_line_if_after_a_colon_and_a_line_number(line: str) -> None:
    assert len(_by_code(_file_procedure(line), "malformed-statement")) == 1


@pytest.mark.parametrize("line", [
    'Open "x" For Binary Access Read Write Lock Read As #f Len = 128', 'Open "x" For Append Access Read Write As #f',
    'Open "x" Access Read As #f', "Width #f, 80", "Print #f,", "Print #f, s; n", "Write #f, s, n",
    "Input #f, s, n, a(1)", "Line Input #f, s", "Get #f, , v", "Get f, 1, v", "Put #f, , v", "Put f, 1, a(1)",
    'Name "a" As "b"', 'Name s & "x" As s', "Debug.Print s",
])
def test_file_statement_shapes_that_compile_are_left_alone(line: str) -> None:
    assert _compile_errors(_file_procedure("If False Then", "    " + line, "End If")) == []


def test_file_statement_shape_leaves_names_that_only_look_like_the_statements() -> None:
    # A path that names a variable access has no Access clause.
    opens = [
        "Open access For Input As #f", "Open (access) For Input As #f", 'Open "x" & access For Input As #f',
        "Open CStr(Not access) For Input As #f", "Open CStr(Not access) As #f", 'Open "x" & Not access As #f',
        "Open CStr(1 Mod access) Access Read As #f",
        *(f'Open "x" {op} access As #f' for op in ["<>", "<=", ">=", "=<", "=>", "><", "=", "<", ">", "&", "+"]),
    ]
    for line in opens:
        assert _by_code(_file_procedure("Dim access As String", "If False Then " + line), "malformed-statement") == [], line
    # A member named access is part of the path.
    members = "\r\n".join([
        "Option Explicit", "Private Type R", "    access As String", "    s As String", "End Type", "Sub Main()",
        "    Dim t As R, f As Integer", "    f = FreeFile", "    If False Then Open t.access For Input As #f",
        "    If False Then Open t!access Access Read As #f", '    If False Then Open "x" & t.access For Input As #f',
        "End Sub", "",
    ])
    assert _by_code(members, "malformed-statement") == []
    # A member straight after Width or Name, a label, or an assignment to the
    # name itself is no such statement.
    for source in [
        "Option Explicit\r\nSub Main()\r\n    Dim Name As String\r\n    Name$ = \"x\"\r\n    Name = \"y\"\r\nEnd Sub\r\n",
        "Option Explicit\r\nSub Main(ByVal Width As Object, ByVal Name As Object)\r\n    Width.Add 1, 2\r\n"
        "    Width!k = 1\r\n    Width = 300\r\n    Name.Add 1\r\n    Name!k = 1\r\nWidth:\r\nName:\r\nEnd Sub\r\n",
        "Option Explicit\r\nSub Main()\r\n    Dim Width As Double\r\n    Width# = 300\r\nEnd Sub\r\n"
        "Sub Other()\r\n    Dim Width As Single\r\n    Width! = 1\r\nEnd Sub\r\n",
    ]:
        assert _by_code(source, "malformed-statement") == [], source
    # A With block's member is a variable to read into.
    record = "\r\n".join([
        "Option Explicit", "Private Type R", "    s As String", "End Type", "Sub Main()", "    Dim t As R, f As Integer",
        "    f = FreeFile", "    With t", "        If False Then Input #f, .s", "        If False Then Line Input #f, .s",
        "    End With", "End Sub", "",
    ])
    assert _by_code(record, "malformed-statement") == []
    # A form's Width and Name are its properties.
    form = 'Option Explicit\r\nPrivate Sub UserForm_Initialize()\r\n    Width = 300\r\n    Name = "x"\r\nEnd Sub\r\n'
    assert _kind_shapes(form, ModuleSymbolKind.USERFORM) == []


def test_file_statement_shape_past_an_element_list_or_a_procedure_of_the_name() -> None:
    for line in [
        "Width f, 80 = 80", "Width (f), 80", "Width(1) = 2", "Width(1).Add 1, 2", "Width(1)(2).Add 1, 2",
        "Width#(1) = 300", "Name(1).Add 1, 2", "Name(1)(2).Add 1", 'Name(1) = "x"', 'Name(1).s = "x"',
        'Name$(1) = "x"',
    ]:
        assert len(_by_code(_file_procedure(line), "malformed-statement")) == 1, line
    calls = (
        "Option Explicit\r\nSub Width(ByVal a As Long, ByVal b As Long)\r\nEnd Sub\r\n"
        "Sub Name(ByVal a As String, ByVal b As String)\r\nEnd Sub\r\nSub Main()\r\n    Width 1, 80\r\n"
        '    Name "a", "b"\r\nEnd Sub\r\n'
    )
    assert len(_by_code(calls, "malformed-statement")) == 2


def test_file_statement_shape_refuses_me_member_as_a_place_to_read_into() -> None:
    def source(line: str) -> str:
        return (
            "Option Explicit\r\nPublic s As String\r\nSub Load()\r\n    Dim f As Integer\r\n    f = FreeFile\r\n"
            f"    If False Then {line}\r\nEnd Sub\r\n"
        )

    for line in ["Input #f, Me.s", "Line Input #f, Me.s", "Get #f, , Me.s"]:
        found = _kind_shapes(source(line), ModuleSymbolKind.CLASS)
        assert len(found) == 1, line
        assert "Variable required - can't assign to this expression." in found[0].message
    # Another place that is no variable's shape makes it a Syntax error.
    mixed = _kind_shapes(source("Input #f, Me.s, 1"), ModuleSymbolKind.CLASS)
    assert len(mixed) == 1
    assert "Syntax error." in mixed[0].message


def test_file_statement_shape_reports_no_run_time_file_error() -> None:
    source = _file_procedure('Open "x" For Input As #f', "Print #f", "Close #f")
    codes = [d.code for d in analyze_module(source, AnalyzeModuleOptions(host="excel"))]
    assert [code for code in codes if code != "malformed-statement" and code.startswith("file-")] == []
    assert len(_by_code(source, "malformed-statement")) == 1


def test_file_statements_still_know_the_project_procedures() -> None:
    # OtherHelper may close and reopen the file, so nothing is known after it,
    # in the module's first procedure and in its second.
    first = _file_procedure('Open "x" For Output As #f', "n = OtherHelper", "Line Input #f, s", "Close #f")
    helper = "Option Explicit\r\nPublic Function OtherHelper() As Long\r\nEnd Function\r\n"
    for source in [first, first.replace("Sub Main()", "Sub Other()\r\nEnd Sub\r\nSub Main()")]:
        index = ProjectIndex()
        index.set_module(ModuleInput("Module1", ModuleSymbolKind.STANDARD, source))
        index.set_module(ModuleInput("Module2", ModuleSymbolKind.STANDARD, helper))
        opts = analyze_module_options_for(
            index,
            "Module1",
            ModuleSymbolKind.STANDARD,
            conditional_compilation=ConditionalCompilationEnvironment(project_constants={}),
        )
        assert "file-mode-mismatch" not in [d.code for d in analyze_module(source, opts)]


# -- 8893e3d0: DefDec, and a Declare without Lib after its name --------------

_DECLARATION_LINES = _measured("declaration_line_")


def _compile_diagnostics(source: str, host: str = "excel") -> list[VbaDiagnostic]:
    out = []
    for d in analyze_module(source, AnalyzeModuleOptions(host=host)):
        meta = diagnostic_metadata_for_code(d.code)
        if meta is not None and meta.diagnostic_kind == DiagnosticEvidenceKind.COMPILE_ERROR:
            out.append(d)
    return out


def _module(*lines: str) -> str:
    return "\r\n".join(["Option Explicit", *lines, ""])


def test_declaration_lines_cover_every_host() -> None:
    assert {_HOSTS.get(case.id, "excel") for case in _DECLARATION_LINES} == {"excel", "word", "powerpoint", "access"}


@pytest.mark.parametrize("case", _DECLARATION_LINES, ids=lambda case: case.id)
def test_declaration_line_agrees_with_the_vbe(case: OracleCase) -> None:
    # A Def statement in a one-line If also gets its letters' Variable not defined.
    found = [
        d
        for d in _compile_diagnostics(case.modules[0].source, _HOSTS.get(case.id, "excel"))
        if d.code != "undeclared-variable"
    ]
    assert len(found) == (1 if case.expected == "rejected" else 0)


@pytest.mark.parametrize("lines, span, message", [
    (["DefDec A-Z", "Sub Main()", "End Sub"], "DefDec", "Expected: line number or label or statement or end of statement"),
    (["Sub Main()", "End Sub", "DefDec A-Z"], "DefDec", "Syntax error"),
    (["Sub Main()", "    DefDec A-Z", "End Sub"], "DefDec", "Syntax error"),
    (["Private Declare PtrSafe Function F () As Long", "Sub Main()", "End Sub"], "(", "Expected: Lib"),
    (
        ['Private Declare PtrSafe Function F Alias "GetTickCount" Lib "kernel32" () As Long', "Sub Main()", "End Sub"],
        "Alias",
        "Expected: Lib",
    ),
    (["Private Declare PtrSafe Sub S", "Sub Main()", "End Sub"], "S", "Expected: Lib"),
])
def test_declaration_line_reports(lines: list[str], span: str, message: str) -> None:
    source = _module(*lines)
    found = _compile_diagnostics(source)
    assert len(found) == 1
    assert _span_text(source, found[0]) == span
    assert message in found[0].message


@pytest.mark.parametrize("lines", [
    ["DefInt I", "DefLngLng L", "DefCur C"],
    ['Private Declare PtrSafe Function F Lib "kernel32" Alias "GetTickCount" () As Long'],
    ['Private Declare PtrSafe Function F& Lib "kernel32" Alias "GetTickCount" ()'],
    ['Private Declare PtrSafe Function F CDecl Lib "kernel32" Alias "GetTickCount" () As Long'],
    ["Private Declare PtrSafe Function F _", '    Lib "kernel32" Alias "GetTickCount" () As Long'],
    ['Private Declare PtrSafe Sub S Lib "kernel32" Alias "Sleep" (ByVal ms As Long)'],
])
def test_declaration_lines_that_compile_are_left_alone(lines: list[str]) -> None:
    assert _compile_errors(_module(*lines, "Sub Main()", "End Sub")) == []


@pytest.mark.parametrize("name", ["9Sleep", *(f"9Sleep{suffix}" for suffix in ["$", "%", "&", "!", "#", "@", "^"])])
def test_declare_digit_start_name_is_left_to_the_invalid_name_check(name: str) -> None:
    source = _module(f'Private Declare PtrSafe Function {name} Lib "kernel32" Alias "GetTickCount" ()', "Sub Main()", "End Sub")
    assert _by_code(source, "invalid-proc-header") == []


def test_def_dec_claims_no_letters() -> None:
    messages = [d.message for d in _compile_diagnostics(_module("DefDec A", "DefInt A", "Sub Main()", "End Sub"))]
    assert len(messages) == 1
    assert "'DefDec' is a reserved word" in messages[0]


def _def_dec_messages(*lines: str) -> list[str]:
    return [
        d.message
        for d in _compile_diagnostics(_module(*lines))
        if "'DefDec' is a reserved word and no statement" in d.message
    ]


def test_def_dec_after_a_line_number_in_a_one_line_if_and_in_an_if_arm() -> None:
    above = "Expected: line number or label or statement or end of statement"
    for lines, error in [
        (["Sub Main()", "10 DefDec:", "End Sub"], "Syntax error"),
        (["Sub Main()", "    If False Then DefDec A-Z", "End Sub"], "Syntax error"),
        (["Sub Main()", "    If True Then Beep Else DefDec A", "End Sub"], "Syntax error"),
        (["#If False Then", "Sub Other()", "End Sub", "#Else", "DefDec A-Z", "#End If", "Sub Main()", "End Sub"], above),
        (["#If False Then", "Sub Other()", "End Sub", "#End If", "DefDec A-Z", "Sub Main()", "End Sub"], above),
    ]:
        found = _def_dec_messages(*lines)
        assert len(found) == 1, lines
        assert error in found[0], lines
    for line in ["    Beep: DefDec:", "foo: DefDec:", "    If True Then Beep: DefDec:", "    DefDec _\r\n:", " _\r\nDefDec:"]:
        found = _def_dec_messages("Sub Main()", line, "End Sub")
        assert len(found) == 1, line
        assert "Syntax error" in found[0], line


def test_def_dec_label_outside_the_procedures() -> None:
    found = [d.message for d in _compile_diagnostics(_module("DefDec:", "Sub Main()", "End Sub"))]
    assert len(found) == 1
    assert "'DefDec' is a reserved word and no statement" in found[0]
    found = [d.message for d in _compile_diagnostics(_module("Sub Main()", "End Sub", "DefDec:"))]
    assert len(found) == 1
    assert "Syntax error" in found[0]


@pytest.mark.parametrize("source", [
    "Option Explicit\rSub Main()\rDefDec:\rEnd Sub\r",
    "Option Explicit\rSub Main()\r\tDefDec:\rEnd Sub\r",
    "Option Explicit\nSub Main()\nDefDec:\nEnd Sub\n",
    "Option Explicit\rSub Main()\r _\rDefDec:\rEnd Sub\r",
    "Option Explicit\rSub Main()\r _  \rDefDec:\rEnd Sub\r",
    "Option Explicit\r\nSub Main()\r\n _\r\nDefDec:\r\nEnd Sub\r\n",
])
def test_def_dec_label_gets_one_report_with_any_line_end(source: str) -> None:
    assert len([d for d in _compile_diagnostics(source) if "DefDec" in d.message]) == 1


def test_def_dec_label_is_left_to_the_reserved_label_check() -> None:
    messages = [d.message for d in _compile_diagnostics(_module("Function Main() As Variant", "DefDec:", "    Main = 1", "End Function"))]
    assert messages == ["'DefDec' is a reserved word, so 'DefDec:' cannot be a line label. This is a VBE compile error: Syntax error."]


# -- 5b603c43: an Erase of one element of an array ---------------------------


def _spans(source: str, code: str) -> list[str]:
    return [source[d.span.start : d.span.end] for d in analyze_module(source) if d.code == code]


def _hits(source: str, code: str) -> list[VbaDiagnostic]:
    return [d for d in analyze_module(source) if d.code == code]


def test_erase_elements_of_measured_cases() -> None:
    cases = _measured("erase_element_")
    assert len(cases) == 280
    for case in cases:
        source = case.modules[0].source
        found = _by_code(source, "erase-requires-array", _HOSTS.get(case.id, "excel"))
        if case.expected == "accepted":
            assert found == [], case.id


def test_erase_of_one_element_of_an_array_whose_elements_are_no_variant() -> None:
    source = (
        "Private Type T\n    x() As Long\nEnd Type\nPrivate ModuleValues(2) As Long\nSub Main(p() As Long)\n"
        "    Dim a(2) As Long, d() As String, m(1, 1) As Object, u(2) As T, i As Long\n"
        "    Static s(2) As Collection\n    Dim t As T\n    Erase a(1), d(i), m(0, 1)\n"
        "    Erase u(1): Erase s(1) ' note\n10  Erase ModuleValues(1), p(1)\n    Erase a _\n        (1)\n"
        "    Erase t.x(1)\nEnd Sub\n"
    )
    hits = _hits(source, "erase-requires-array")
    # The Type member rule reports first.
    assert [_span_text(source, d) for d in hits] == [
        "t.x(1)", "a(1)", "d(i)", "m(0, 1)", "u(1)", "s(1)", "ModuleValues(1)", "p(1)", "a _\n        (1)",
    ]
    assert "Erase target 't.x(1)' is one element of the array member 't.x'" in hits[0].message
    assert "Erase target 'a(1)' is one element of the array 'a'" in hits[1].message
    assert hits[1].severity == DiagnosticSeverity.ERROR
    assert all("Expected array" in d.message for d in hits)
    assert "Erase target 'a (1)'" in hits[-1].message


def test_erase_of_a_scalar_given_a_subscript() -> None:
    source = (
        "Sub Main()\n    Dim n As Long, s As String, d As Date, b As Boolean, x As Double, c As Currency\n"
        "    Erase n(1), s(1), d(1), b(1), x(1), c(1)\nEnd Sub\n"
    )
    hits = _hits(source, "erase-requires-array")
    assert [_span_text(source, d) for d in hits] == ["n(1)", "s(1)", "d(1)", "b(1)", "x(1)", "c(1)"]
    assert "'n', which is declared As Long and is no array" in hits[0].message


def test_erase_of_an_element_of_another_modules_array() -> None:
    from pyvbaanalysis import analyze_project

    caller = "Sub Main()\n    Erase SharedValues(1)\nEnd Sub\n"
    diagnostics = analyze_project([
        ModuleInput("Caller", ModuleSymbolKind.STANDARD, caller),
        ModuleInput("Globals", ModuleSymbolKind.STANDARD, "Public SharedValues(2) As Long\n"),
    ])["Caller"]
    assert [caller[d.span.start : d.span.end] for d in diagnostics if d.code == "erase-requires-array"] == [
        "SharedValues(1)"
    ]


def test_erase_element_through_with_deflng_type_character_and_subscript_members() -> None:
    source = (
        "DefLng A-Z\nPrivate Type T\n    x() As Long\nEnd Type\nSub Main()\n"
        "    Dim a(2), n As Long, o As Object, t As T, b&(2)\n"
        "    Erase a(1), a(o.To), a(n + 1), a(-1), b&(1), b(1)\n    With t\n        Erase .x(1)\n    End With\n"
        "End Sub\n"
    )
    hits = _hits(source, "erase-requires-array")
    assert [_span_text(source, d) for d in hits] == [".x(1)", "a(1)", "a(o.To)", "a(n + 1)", "a(-1)", "b&(1)", "b(1)"]
    assert "Erase target 't.x(1)'" in hits[0].message
    # The element is reported once, not also as an arbitrary expression.
    assert _spans(source, "invalid-erase-target") == [".x(1)"]


def test_erase_element_keeps_the_dimension_count_report() -> None:
    source = "Sub Main()\n    Dim a(2) As Long, m(1, 1) As Long\n    Erase a(1, 2), m(1)\nEnd Sub\n"
    assert _spans(source, "erase-requires-array") == ["a(1, 2)", "m(1)"]
    assert len(_hits(source, "wrong-number-of-dimensions")) == 2


def test_erase_element_leaves_an_empty_subscript_and_a_named_argument_alone() -> None:
    source = (
        "Private Type T\n    x() As Long\nEnd Type\nSub Main()\n    Dim a(2) As Long, t As T, n As Long\n"
        "    Erase a(1,), a(,), a(,1), a(x:=1), n(1,)\n    Erase t.x(1,), t.x(,), t.x(i:=1)\nEnd Sub\n"
    )
    assert _hits(source, "erase-requires-array") == []


def test_erase_element_reads_a_named_argument_inside_a_call() -> None:
    source = (
        "Private Type T\n    x() As Long\nEnd Type\nSub Main()\n    Dim a(2) As Long, t As T\n"
        "    Erase a(F(x:=1)), t.x(F(x:=1))\nEnd Sub\nFunction F(ByVal x As Long) As Long\nEnd Function\n"
    )
    assert _spans(source, "erase-requires-array") == ["t.x(F(x:=1))", "a(F(x:=1))"]


def test_erase_element_leaves_a_member_of_another_rank_alone() -> None:
    source = "Private Type T\n    f(2) As Long\nEnd Type\nSub Main()\n    Dim t As T\n    Erase t.f(1, 2)\nEnd Sub\n"
    assert _hits(source, "erase-requires-array") == []


def test_erase_element_leaves_the_procedures_own_name_alone() -> None:
    source = (
        "Function G(ByVal i As Long) As Long()\n    Erase G(1)\nEnd Function\n"
        "Property Get P(ByVal i As Long) As String\n    Erase P(1)\nEnd Property\n"
    )
    assert _hits(source, "erase-requires-array") == []


def test_erase_element_quotes_a_member_subscript_as_written() -> None:
    source = (
        "Private Type T\n    x() As Long\nEnd Type\nSub Main()\n    Dim t As T, n As Long\n"
        "    Erase t.x(n Mod 2)\n    Erase t.x(n + 1)\nEnd Sub\n"
    )
    hits = _hits(source, "erase-requires-array")
    assert [_span_text(source, d) for d in hits] == ["t.x(n Mod 2)", "t.x(n + 1)"]
    assert "Erase target 't.x(n Mod 2)'" in hits[0].message
    assert "Erase target 't.x(n + 1)'" in hits[1].message
    # The arbitrary-expression check reads a member's subscript as before.
    assert _spans(source, "invalid-erase-target") == ["t.x(n + 1)"]


def test_erase_element_accepts_what_compiles() -> None:
    source = (
        "Private Type T\n    x() As Long\n    y() As Variant\nEnd Type\nPrivate a(2) As Long\nSub Main(q As Variant)\n"
        "    Dim v As Variant, w(2) As Variant, z(2), o As Object, c As New Collection, t As T, u(2) As T\n"
        "    Erase v(1), w(1), z(1), q(1), o(1), c(1), o.x(1), F(1), G(1)\n"
        "    Erase t.y(1), u(1).x, t.x, a(), a\n    Erase a(1 To 2), t.x(1 To 2)\nEnd Sub\n"
        "Sub Shadowed()\n    Dim a As Variant\n    Erase a(1)\nEnd Sub\n"
        "Function F(ByVal i As Long) As Variant\nEnd Function\nFunction G(ByVal i As Long) As Long\nEnd Function\n"
    )
    assert _hits(source, "erase-requires-array") == []


def test_erase_element_in_an_inactive_arm_is_left_alone() -> None:
    source = "Sub Main()\n    Dim a(2) As Long\n#If False Then\n    Erase a(1)\n#End If\nEnd Sub\n"
    assert _hits(source, "erase-requires-array") == []


# -- acf42528: a ReDim target without bounds ---------------------------------


def test_redim_bounds_measured_cases() -> None:
    cases = _measured("redim_bounds_")
    assert {_HOSTS.get(case.id, "excel") for case in cases} == {"excel", "word", "powerpoint", "access"}
    for case in cases:
        missing = [
            d
            for d in _by_code(case.modules[0].source, "malformed-statement", _HOSTS.get(case.id, "excel"))
            if "missing a bound" in d.message
        ]
        if case.expected == "accepted":
            assert missing == [], case.id


def test_redim_target_without_bounds() -> None:
    source = (
        "Private Type T\n    x() As Long\nEnd Type\nSub Main()\n"
        "    Dim a() As Long, b() As Long, m() As Long, v As Variant, t As T\n    ReDim a()\n"
        "    ReDim Preserve a() As Long\n    ReDim b(1), a, v\n    ReDim m(1, ), m(, 1)\n"
        "    ReDim a(1 To ), a(To 3)\n    ReDim t.x()\n    With t\n        ReDim .x\n    End With\nEnd Sub\n"
    )
    hits = _hits(source, "malformed-statement")
    assert [_span_text(source, d) for d in hits] == [
        "a()", "a()", "a", "v", "m(1, )", "m(, 1)", "a(1 To )", "a(To 3)", "t.x()", ".x",
    ]
    assert hits[0].severity == DiagnosticSeverity.ERROR
    for part in ["ReDim target 'a()' is missing a bound", "a(1 To 3)", "Syntax error"]:
        assert part in hits[0].message
    assert "t.x(1 To 3)" in hits[8].message


def test_redim_without_bounds_after_a_line_number_a_colon_a_comment_and_a_continuation() -> None:
    source = (
        "Sub Main()\n    Dim a() As Long, b() As Long\n10  ReDim a()\n    ReDim a(): Beep\n"
        "    ReDim a() ' note\n    ReDim b(1), _\n        a()\nEnd Sub\n"
    )
    assert _spans(source, "malformed-statement") == ["a()", "a()", "a()", "a()"]


def test_redim_without_bounds_of_a_scalar_or_constant() -> None:
    source = "Private Const c As Long = 1\nSub Main()\n    Dim n As Long\n    ReDim n\n    ReDim c()\nEnd Sub\n"
    assert _spans(source, "malformed-statement") == ["n", "c()"]
    assert _hits(source, "scalar-redim") == []


@pytest.mark.parametrize("line", ["ReDim a(1", "ReDim a(1))", "ReDim a(1) = 2", "ReDim b(1), a("])
def test_redim_unbalanced_or_continued_target_is_left_alone(line: str) -> None:
    source = "Sub Main()\n    Dim a() As Long, b() As Long\n    " + line + "\nEnd Sub\n"
    assert [d for d in _hits(source, "malformed-statement") if "missing a bound" in d.message] == []


def test_redim_without_bounds_keeps_the_element_type_report() -> None:
    source = "Sub Main()\n    Dim a() As Long\n    ReDim a() As String\nEnd Sub\n"
    assert _spans(source, "malformed-statement") == ["a()"]
    assert len(_hits(source, "redim-type-change")) == 1


def test_redim_target_that_is_no_name_is_left_alone() -> None:
    source = "Sub Main()\n    Dim a() As Long, b As Long\n    ReDim a b\nEnd Sub\n"
    assert [d for d in _hits(source, "malformed-statement") if "missing a bound" in d.message] == []


def test_redim_member_named_as_is_a_member() -> None:
    source = "Sub Main()\n    Dim o As Object\n    ReDim o.As(1)\nEnd Sub\n"
    assert _hits(source, "malformed-statement") == []


def test_redim_without_bounds_through_a_subscript() -> None:
    source = (
        "Private Type T\n    x() As Long\n    f(2) As Long\nEnd Type\nSub Main()\n    Dim u(2) As T\n"
        "    ReDim u(1).x(), u(1).f()\nEnd Sub\n"
    )
    assert _spans(source, "malformed-statement") == ["u(1).x()", "u(1).f()"]
    assert _hits(source, "fixed-array-redim") == []
    assert _hits(source, "scalar-redim") == []


def test_redim_type_member_without_bounds() -> None:
    source = (
        "Private Type T\n    n As Long\n    f(2) As Long\nEnd Type\nSub Main()\n    Dim t As T\n"
        "    ReDim t.n(), t.f(), t.f(1 To )\nEnd Sub\n"
    )
    assert _spans(source, "malformed-statement") == ["t.n()", "t.f()", "t.f(1 To )"]
    assert _hits(source, "scalar-redim") == []
    assert _hits(source, "fixed-array-redim") == []


def test_redim_scalar_with_empty_parentheses() -> None:
    source = "Sub Main()\n    Dim n As Long\n    ReDim n()\nEnd Sub\n"
    assert _spans(source, "malformed-statement") == ["n()"]
    assert _hits(source, "scalar-redim") == []


def test_redim_with_bounds_is_accepted() -> None:
    source = (
        "Private Type T\n    x() As Long\nEnd Type\nSub Main()\n"
        "    Dim a() As Long, m() As Long, o As Object, t As T, b()\n    ReDim a(1), m(1, 0 To 2)\n"
        "    ReDim Preserve a(0 To 1)\n    ReDim a(o.To), a(1 To o.To)\n    ReDim t.x(1)\n    With t\n"
        "        ReDim .x(1)\n    End With\n#If False Then\n    ReDim a()\n#End If\nEnd Sub\n"
    )
    assert _hits(source, "malformed-statement") == []
