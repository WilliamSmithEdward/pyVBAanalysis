"""Rule contracts from the XLIDE 11.2.1 -> 11.2.11 sync (calls, events,
conditional compilation and the rest).

Each section names the upstream commit whose own tests it follows.
"""

import json
import re

import pytest

from pyvbaanalysis import (
    AnalysisFailure,
    analyze_module,
    analyze_module_options_for,
    host_token_for_file_name,
)
from pyvbaanalysis.diagnostics.analyze_module import _run_rules, with_resolved_host_model
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment, parse_project_conditional_constants
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.diagnostics.rule_metadata import diagnostic_metadata_for_code
from pyvbaanalysis.diagnostics.rules.undeclared import designer_class_member_names
from pyvbaanalysis.evidence import DATA_DIR
from pyvbaanalysis.symbols import (
    ImplicitMember,
    ModuleInput,
    ModuleSymbolKind,
    ProjectIndex,
    ProjectIndexOptions,
)

_RAW_CASES = json.loads((DATA_DIR / "vbe_oracle_cases.json").read_text(encoding="utf-8"))["cases"]

_KINDS = {
    "standard": ModuleSymbolKind.STANDARD,
    "class": ModuleSymbolKind.CLASS,
    "document": ModuleSymbolKind.DOCUMENT,
    "userform": ModuleSymbolKind.USERFORM,
}


def _oracle_input(module: dict) -> ModuleInput:
    """An oracle fixture module as project input: its code name, kind and designer
    class, and a renamed document module's document type by the name it was made
    with (the subset of oracle_support.oracle_module_input these cases use)."""
    code_name = module.get("codeName")
    lower = module["name"].lower()
    document_type = (
        "workbook" if lower == "thisworkbook" else "document" if lower == "thisdocument" else "worksheet"
    )
    designer_class = module.get("designerClass")
    return ModuleInput(
        code_name if code_name is not None else module["name"],
        ModuleSymbolKind.USERFORM
        if designer_class is not None and designer_class.startswith("Access.")
        else _KINDS.get(module.get("type", "standard"), ModuleSymbolKind.STANDARD),
        module["source"],
        designer_class=module.get("designerClass"),
        document_type=document_type if code_name is not None else None,
        implicit_members=[ImplicitMember(m["name"], m["type"]) for m in module["implicitMembers"]]
        if module.get("implicitMembers") is not None
        else None,
    )


def _project_errors(modules: list[ModuleInput], host: str | None = None) -> list[str]:
    """Error findings of every module of a project, as `module: code: message`,
    analyzed as upstream's issue tests run them (known-empty project arguments)."""
    index = ProjectIndex()
    for module in modules:
        index.set_module(module)
    out = []
    for module in modules:
        opts = analyze_module_options_for(
            index, module.module_name, module.module_kind, host=host,
            conditional_compilation=ConditionalCompilationEnvironment(project_constants={}),
        )
        out += [
            f"{module.module_name}: {d.code}: {d.message}"
            for d in analyze_module(module.source, opts)
            if d.severity.value == "error"
        ]
    return out


def _hits(source: str, code: str, options: AnalyzeModuleOptions | None = None) -> list:
    return [d for d in analyze_module(source, options) if d.code == code]


def _compile_errors(*lines: str) -> list[str]:
    """Error findings that are not runtime errors, as `code: message`."""
    source = "\n".join(["Option Explicit", *lines, ""])
    out = []
    for d in analyze_module(source):
        meta = diagnostic_metadata_for_code(d.code)
        if d.severity.value != "error":
            continue
        if meta is not None and meta.diagnostic_kind.value == "deterministic-runtime-error":
            continue
        out.append(f"{d.code}: {d.message}")
    return out


# -- 4b579988: block structure the way the VBE reads it (issue #1496) ---------


@pytest.mark.parametrize("lines", [
    ("#If VBA7 Then", "Public Sub Foo(ByVal p As LongPtr)", "#Else",
     "Public Sub Foo(ByVal p As Long)", "#End If", "    Dim x As Long", "    x = 1", "End Sub"),
    ("#If VBA7 Then", "Public Sub Foo(ByVal p As LongPtr)", "#Else", "#End If",
     "    Dim x As Long", "    x = 1", "End Sub"),
    ("Public Sub Main()", "    Dim c As New Collection, d As New Collection, a As Boolean",
     "    With c", "        If a Then With d: .Add 1: End With", "        .Add 2", "    End With",
     "End Sub"),
    ("Public Sub Main()",
     "    Dim c As New Collection, e As New Collection, a As Boolean, outer As Boolean, b As Long",
     "    If outer Then", "        With c", "            If a Then b = 1: Else With e: .Add 3: End With",
     "            .Add 1", "        End With", "    End If", "End Sub"),
    ("Public Sub Main()", "    Dim x As Boolean, y As Long, z As Long", "10  If x Then: y = 1",
     "    z = 2", "End Sub"),
    ("Public Sub Main()", "    Dim c As New Collection, d As Object, a As Boolean", "    With c",
     "        If a Then With d.Else: End With", "        .Add 2", "    End With", "End Sub"),
    ("Public Sub Main()", "    Dim d As Object, x As Object, a As Boolean, y As Long",
     "    If a Then Set x = d.Then", "    y = 1", "End Sub"),
    ("Public Sub Main()", "    Dim c As New Collection, x!", "    If x! Then With c: .Add 1: End With",
     "End Sub"),
])
def test_1496_block_shapes_compile_clean(lines: tuple[str, ...]) -> None:
    assert _compile_errors(*lines) == []


def test_1496_member_named_then_is_not_a_reserved_keyword() -> None:
    for tail in ("d.Then", "d!Then"):
        source = f"Sub Main()\n    Dim d As Object, x As Object, a As Boolean\n    If a Then Set x = {tail}\nEnd Sub\n"
        assert _hits(source, "if-reserved-keyword-in-condition") == []


def test_1496_second_then_is_still_reported() -> None:
    errors = _compile_errors("Sub Main()", "    If True Then Then", "    End If", "End Sub")
    assert any(e.startswith("if-reserved-keyword-in-condition") for e in errors)


# -- b7cbc58c: a document module named bare is its module's object (#1494) ----

_ISSUE_1494 = [raw for raw in _RAW_CASES if raw["id"].startswith("issue1494_")]


def test_1494_has_every_measured_case() -> None:
    assert len(_ISSUE_1494) == 28


@pytest.mark.parametrize("raw", _ISSUE_1494, ids=[raw["id"] for raw in _ISSUE_1494])
def test_1494_agrees_with_the_vbe(raw: dict) -> None:
    found = _project_errors([_oracle_input(m) for m in raw["modules"]], raw.get("host"))
    if raw["expected"] == "accepted":
        assert found == []
    else:
        assert len(found) == 1
        assert re.search(
            r": (assignment-object-type-mismatch|argument-type-mismatch): .*Run-time error '13'", found[0]
        )


def _main_project(host: str, modules: list[ModuleInput], *body: str) -> list[str]:
    source = "\r\n".join(
        ["Option Explicit", "Public Sub Main()", *(f"    {line}" for line in body), "End Sub", ""]
    )
    return _project_errors([*modules, ModuleInput("Module1", ModuleSymbolKind.STANDARD, source)], host)


def _document(name: str, designer_class: str | None = None) -> ModuleInput:
    return ModuleInput(name, ModuleSymbolKind.DOCUMENT, "", designer_class=designer_class)


def test_1494_variable_of_the_document_type_is_judged_by_its_host_class() -> None:
    sheet = [_document("Sheet1")]
    assert _main_project("excel", sheet, "Dim s As Sheet1", "Dim w As Worksheet", "Set w = s") == []
    assert len(_main_project("excel", sheet, "Dim s As Sheet1", "Dim b As Workbook", "Set b = s")) == 1
    form = [_document("Form_Orders")]
    assert _main_project("access", form, "Dim f As Form_Orders", "Dim g As Form", "Set g = f") == []


def test_1494_access_form_listed_as_a_designer_module() -> None:
    form = [ModuleInput("Form_Orders", ModuleSymbolKind.USERFORM, "", designer_class="Access.Form")]
    assert len(_main_project("access", form, "Dim x As Report", "Set x = Form_Orders")) == 1
    assert _main_project("access", form, "Dim x As Form", "Set x = Form_Orders") == []


def test_1494_another_hosts_document_is_not_judged() -> None:
    assert _main_project("word", [_document("Sheet1")], "Dim r As Range", "Set r = Sheet1") == []


def test_1494_module_variable_of_the_same_name_takes_precedence() -> None:
    modules = [
        _document("Sheet1"),
        ModuleInput("Module2", ModuleSymbolKind.STANDARD, "Option Explicit\r\nPublic Sheet1 As Range\r\n"),
    ]
    assert _main_project("excel", modules, "Dim r As Range", "Set r = Sheet1") == []


# -- 8506a287: event parameters by DefType, WithEvents Collections (#1493) ----

_ISSUE_1493 = [raw for raw in _RAW_CASES if raw["id"].startswith("issue1493_")]


def test_1493_has_the_measured_cases() -> None:
    assert len(_ISSUE_1493) == 164


@pytest.mark.parametrize("raw", _ISSUE_1493, ids=[raw["id"] for raw in _ISSUE_1493])
def test_1493_agrees_with_the_vbe(raw: dict) -> None:
    found = _project_errors([_oracle_input(m) for m in raw["modules"]], raw.get("host"))
    if raw["expected"] == "accepted":
        assert found == []
    else:
        assert len(found) > 0


def _pulse(source_def: str, listener_def: str, handler: str) -> list[ModuleInput]:
    return [
        ModuleInput("Source", ModuleSymbolKind.CLASS, f"Option Explicit\r\n{source_def}\r\nPublic Event Pulse(ByVal x)\r\n"),
        ModuleInput(
            "Listener",
            ModuleSymbolKind.CLASS,
            f"Option Explicit\r\n{listener_def}\r\nPrivate WithEvents src As Source\r\n"
            f"Private Sub src_Pulse({handler})\r\n    Debug.Print 1\r\nEnd Sub\r\n",
        ),
    ]


def test_1493_names_the_def_type_that_typed_the_handler_parameter() -> None:
    found = _project_errors(_pulse("DefLng X", "DefStr X", "ByVal x"))
    assert len(found) == 1
    assert "parameter 1, 'x' is String (by the module's DefType), and the event's is Long" in found[0]


def test_1493_keeps_the_event_as_written() -> None:
    index = ProjectIndex()
    for module in _pulse("DefLng X", "", "ByVal x As Long"):
        index.set_module(module)
    source = next(t for t in index.project_member_surfaces("Listener") if t.name == "Source")
    event = next(m for m in source.members if m.name == "Pulse")
    assert event.signature == "Pulse(ByVal x)"
    assert [p.type_ for p in event.event_params or []] == ["Long"]


def _listener(as_type: str) -> ModuleInput:
    return ModuleInput("Listener", ModuleSymbolKind.CLASS, f"Option Explicit\r\nPrivate WithEvents c As {as_type}\r\n")


def test_1493_vba_collection_and_a_standard_module_named_collection() -> None:
    found = _project_errors([_listener("VBA.Collection")])
    assert len(found) == 1 and "VBA's Collection, which does not source automation events" in found[0]
    found = _project_errors([_listener("[VBA.Collection]")])
    assert len(found) == 1 and "User-defined type not defined" in found[0]
    found = _project_errors(
        [ModuleInput("Collection", ModuleSymbolKind.STANDARD, "Option Explicit\r\n"), _listener("Collection")]
    )
    assert len(found) == 1 and "A module is not a valid type" in found[0]


# -- ca38c7bf: .pot and .pps are PowerPoint files (#1538) ---------------------


@pytest.mark.parametrize("name", ["Deck.pot", "Show.PPS", "Deck.ppt"])
def test_1538_legacy_powerpoint_names_are_powerpoint(name: str) -> None:
    assert host_token_for_file_name(name) == "powerpoint"


def test_1538_legacy_powerpoint_file_reports_its_members() -> None:
    source = "Sub T()\n    Debug.Print ActivePresentation.Slides.Nope\nEnd Sub\n"
    opts = AnalyzeModuleOptions(host=host_token_for_file_name("Deck.pot"))
    assert [d.code for d in analyze_module(source, opts) if d.severity.value == "error"] == ["member-not-found"]


# -- e4b377d1: a project argument Win64 = 0 leaves the 64-bit rules on (#1551) -

_WIN64_PROBE = "#If Win64 Then\nPrivate m As Long\nPrivate m As Long\n#End If\n"
_WIN64_CASES = [
    ("declare_no_ptrsafe", "declare-missing-ptrsafe",
     'Private Declare Function GetTickCount Lib "kernel32" () As Long\n'),
    ("declare_ptrsafe", None, 'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long\n'),
    ("varptr_into_long", "longlong-narrowing",
     "Sub Main()\n    Dim x As Long\n    Dim n As Long\n    n = VarPtr(x)\nEnd Sub\n"),
    ("varptr_into_longptr", None,
     "Sub Main()\n    Dim x As Long\n    Dim n As LongPtr\n    n = VarPtr(x)\nEnd Sub\n"),
    ("addressof_byval_long", "addressof-misuse",
     "Sub Take(ByVal v As Long)\nEnd Sub\nSub Callback()\nEnd Sub\nSub Main()\n    Take AddressOf Callback\nEnd Sub\n"),
    ("addressof_byval_longptr", None,
     "Sub Take(ByVal v As LongPtr)\nEnd Sub\nSub Callback()\nEnd Sub\nSub Main()\n"
     "    Take AddressOf Callback\nEnd Sub\n"),
]


def _win64_error_codes(body: str, arguments_text: str, host: str) -> list[str]:
    """The project as the extension builds it: the arguments text parsed into the
    project's constants, given to the index and to the rules alike."""
    source = f"Option Explicit\n{body}"
    env = ConditionalCompilationEnvironment(project_constants=parse_project_conditional_constants(arguments_text))
    index = ProjectIndex(ProjectIndexOptions(conditional_compilation=env))
    index.set_module(ModuleInput("Module1", ModuleSymbolKind.STANDARD, source))
    opts = analyze_module_options_for(index, "Module1", ModuleSymbolKind.STANDARD, host=host, conditional_compilation=env)
    return [d.code for d in analyze_module(source, opts) if d.severity.value == "error"]


@pytest.mark.parametrize("host", ["excel", "word", "powerpoint", "access"])
def test_1551_project_win64_argument_turns_if_win64_off(host: str) -> None:
    assert _win64_error_codes(_WIN64_PROBE, "", host) == ["duplicate-module-variable"]
    assert _win64_error_codes(_WIN64_PROBE, "Win64 = 0", host) == []


@pytest.mark.parametrize("host", ["excel", "word", "powerpoint", "access"])
@pytest.mark.parametrize("name, rule, body", _WIN64_CASES, ids=[c[0] for c in _WIN64_CASES])
def test_1551_project_win64_argument_leaves_64_bit_rules_on(host: str, name: str, rule: str | None, body: str) -> None:
    assert _win64_error_codes(body, "Win64 = 0", host) == ([rule] if rule else [])


def test_1551_vb6_leaves_the_three_rules_off() -> None:
    rules = {"declare-missing-ptrsafe", "longlong-narrowing", "addressof-misuse"}
    for _name, _rule, body in _WIN64_CASES:
        assert [c for c in _win64_error_codes(body, "Win64 = 0", "vb6") if c in rules] == []


# -- c4d3b5c2: an empty argument no longer stops the runtime value checks -----
# XLIDE issue #1564, which the port's rules fuzzer found first.


def _runtime_value_lines(*procedure_a: str) -> list[str]:
    """The trimmed line of each runtime-argument-value finding."""
    source = "\r\n".join([
        "Option Explicit", "Public Sub A()", "    Dim s As String", "    Dim d As Date", '    s = "abc"',
        *(f"    {line}" for line in procedure_a),
        "End Sub", "Public Sub B()", "    Err.Raise 0", "End Sub", "",
    ])
    lines = source.split("\r\n")
    return [
        lines[source[: d.span.start].count("\r\n")].strip()
        for d in analyze_module(source)
        if d.code == "runtime-argument-value"
    ]


def test_1564_reports_err_raise_0_without_an_empty_argument() -> None:
    assert _runtime_value_lines('Mid(s, 2) = "z"') == ["Err.Raise 0"]


@pytest.mark.parametrize("line", [
    'Mid(s, ) = "z"',
    'Mid$(s, ) = "z"',
    'MidB(s, ) = "z"',
    'MidB$("abc", ) = "z"',
    'Mid(s, , 1) = "z"',
    'd = DateAdd("d", , d)',
    'd = DateAdd("d", 1, DateSerial(2026, , 1))',
    'd = DateAdd("d", 1, DateSerial(2026, 1, ))',
])
def test_1564_still_reports_err_raise_0_after_an_empty_argument(line: str) -> None:
    assert "Err.Raise 0" in _runtime_value_lines(line)


def test_1564_rules_do_not_raise_on_mid_with_an_empty_start() -> None:
    def raise_error(error: BaseException, where: AnalysisFailure) -> None:
        raise error

    source = 'Sub A()\n    Dim s As String\n    Mid(s, ) = "z"\nEnd Sub\n'
    _run_rules(source, with_resolved_host_model(AnalyzeModuleOptions()), raise_error)


# -- 647576aa: an .xlsm sheet's module takes its Worksheet members bare -------


def _undeclared_names(body: str, implicit: list[ImplicitMember]) -> list[str]:
    """Undeclared names in a sheet whose class the workbook supplied, with the
    host named and no model passed, as the extension analyzes an .xlsm."""
    src = f"Option Explicit\nPublic Function Probe() As Variant\n    {body}\nEnd Function\n"
    opts = AnalyzeModuleOptions(
        module_name="Sheet1", module_kind=ModuleSymbolKind.DOCUMENT, document_type="worksheet",
        designer_class="Excel.Worksheet", host="excel", implicit_members=implicit,
        known_identifiers=set(), known_procedures=set(),
    )
    return [
        src[d.span.start : d.span.end]
        for d in analyze_module(src, opts)
        if d.code in ("undeclared-variable", "unknown-call")
    ]


@pytest.mark.parametrize("body", [
    "Probe = UsedRange.Address", "Probe = Shapes.Count", "Probe = Index", "Activate", "Unprotect",
])
def test_1649_sheet_with_a_designer_class_takes_its_members_bare(body: str) -> None:
    assert _undeclared_names(body, []) == []
    assert _undeclared_names(body, [ImplicitMember("CommandButton1", "Object")]) == []


@pytest.mark.parametrize("body, name", [
    ("Probe = Nope", "Nope"), ("Probe = FullName", "FullName"), ("Nope", "Nope"),
])
def test_1649_sheet_with_a_designer_class_still_reports_what_it_lacks(body: str, name: str) -> None:
    assert _undeclared_names(body, []) == [name]


def test_1649_designer_class_members_default_to_excel() -> None:
    assert "usedrange" in designer_class_member_names("Excel.Worksheet", None)


# -- 3eca9bec: a statement that calls a control gets the VBE's message (#1555) -

_BUTTON = [ImplicitMember("CommandButton1", "Object")]
_CONTROL_HOSTS: dict[str, dict] = {
    "sheet": dict(module_name="Sheet1", module_kind=ModuleSymbolKind.DOCUMENT, document_type="worksheet",
                  designer_class="Excel.Worksheet", host="excel", implicit_members=_BUTTON),
    "document": dict(module_name="ThisDocument", module_kind=ModuleSymbolKind.DOCUMENT,
                     designer_class="Word.Document", host="word", implicit_members=_BUTTON),
    "slide": dict(module_name="Slide1", module_kind=ModuleSymbolKind.DOCUMENT,
                  designer_class="PowerPoint.Slide", host="powerpoint", implicit_members=_BUTTON),
    "form": dict(module_name="Form1", module_kind=ModuleSymbolKind.USERFORM, host="excel",
                 implicit_members=[ImplicitMember("CommandButton1", "MSForms.CommandButton")]),
    "access": dict(module_name="Form_Frm1", module_kind=ModuleSymbolKind.USERFORM, designer_class="Access.Form",
                   host="access", implicit_members=[
                       ImplicitMember("Detail", "Access.Section"),
                       ImplicitMember("CommandButton1", "Access.CommandButton"),
                   ]),
}
_CONTROL_CALLS = ["CommandButton1", "Call CommandButton1", "CommandButton1 1", "Call CommandButton1(1)"]
_AS_VARIABLE = (
    "non-callable-call CommandButton1: Cannot call 'CommandButton1' because it is a control, not a Sub or "
    "Function. This is a VBE compile error: Expected procedure, not variable."
)
_AS_PROPERTY = (
    "invalid-property-use CommandButton1: 'CommandButton1' is a control, and a statement cannot call one. "
    "This is a VBE compile error: Invalid use of property."
)


def _control_errors(body: str, known_procedures: frozenset[str] = frozenset(), **options: object) -> list[str]:
    src = f"Option Explicit\nPublic Sub Probe()\n    {body}\nEnd Sub\n"
    opts = AnalyzeModuleOptions(known_identifiers=set(), known_procedures=set(known_procedures), **options)  # type: ignore[arg-type]
    return [
        f"{d.code} {src[d.span.start : d.span.end]}: {d.message}"
        for d in analyze_module(src, opts)
        if d.severity.value == "error"
    ]


@pytest.mark.parametrize("where", ["sheet", "document", "slide"])
def test_1555_a_document_control_is_expected_procedure_not_variable(where: str) -> None:
    for body in _CONTROL_CALLS:
        assert _control_errors(body, **_CONTROL_HOSTS[where]) == [_AS_VARIABLE], body


@pytest.mark.parametrize("where", ["form", "access"])
def test_1555_a_form_control_is_invalid_use_of_property(where: str) -> None:
    for body in _CONTROL_CALLS:
        assert _control_errors(body, **_CONTROL_HOSTS[where]) == [_AS_PROPERTY], body


def test_1555_names_an_access_section_as_a_section() -> None:
    expected = [
        "invalid-property-use Detail: 'Detail' is a section, and a statement cannot call one. This is a VBE "
        "compile error: Invalid use of property."
    ]
    assert _control_errors("Detail", **_CONTROL_HOSTS["access"]) == expected
    report = {**_CONTROL_HOSTS["access"], "module_name": "Report_Rpt1", "designer_class": "Access.Report"}
    assert _control_errors("Detail", **report) == expected


@pytest.mark.parametrize("where", list(_CONTROL_HOSTS))
def test_1555_a_member_of_the_control_is_left_alone(where: str) -> None:
    assert _control_errors("CommandButton1.SetFocus", **_CONTROL_HOSTS[where]) == []
    assert _control_errors("Call CommandButton1.SetFocus", **_CONTROL_HOSTS[where]) == []


def test_1555_a_name_that_is_no_control_is_still_not_defined() -> None:
    for where in ("sheet", "form"):
        assert _control_errors("Nope", **_CONTROL_HOSTS[where]) == [
            "unknown-call Nope: Sub or Function not defined: 'Nope'."
        ]


def test_1555_what_the_procedure_or_module_declares_takes_the_name() -> None:
    sheet = _CONTROL_HOSTS["sheet"]
    found = _control_errors("Dim CommandButton1 As Long\n    CommandButton1", **sheet)
    assert [line.split(":")[0] for line in found] == ["non-callable-call CommandButton1"]
    src = "Option Explicit\nPublic Sub Probe()\n    CommandButton1\nEnd Sub\nPrivate Sub CommandButton1()\nEnd Sub\n"
    opts = AnalyzeModuleOptions(known_identifiers=set(), known_procedures={"probe", "commandbutton1"}, **sheet)  # type: ignore[arg-type]
    assert [d for d in analyze_module(src, opts) if d.code in ("non-callable-call", "unknown-call")] == []
    for where in ("sheet", "form"):
        with_enum = (
            "Option Explicit\nPrivate Enum Kinds\n    CommandButton1 = 1\nEnd Enum\nPublic Sub Probe()\n"
            "    Call CommandButton1\nEnd Sub\n"
        )
        opts = AnalyzeModuleOptions(known_identifiers=set(), known_procedures={"probe"}, **_CONTROL_HOSTS[where])  # type: ignore[arg-type]
        assert [d.message for d in analyze_module(with_enum, opts) if d.severity.value == "error"] == [
            "Cannot call 'CommandButton1' because it resolves to an enum member, not a Sub or Function."
        ]


def test_1555_the_control_wins_over_a_procedure_elsewhere() -> None:
    known = frozenset({"commandbutton1"})
    assert _control_errors("CommandButton1", known, **_CONTROL_HOSTS["sheet"]) == [_AS_VARIABLE]
    assert _control_errors("Call CommandButton1", known, **_CONTROL_HOSTS["form"]) == [_AS_PROPERTY]


@pytest.mark.parametrize("name", ["Beep", "Calculate"])
def test_1555_the_control_wins_over_vba_and_host_names(name: str) -> None:
    form = {**_CONTROL_HOSTS["form"], "implicit_members": [ImplicitMember(name, "MSForms.CommandButton")]}
    assert _control_errors(name, **form) == [
        f"invalid-property-use {name}: '{name}' is a control, and a statement cannot call one. This is a VBE "
        "compile error: Invalid use of property."
    ]


def test_1555_an_access_form_arriving_as_a_document_module() -> None:
    access = {**_CONTROL_HOSTS["access"], "module_kind": ModuleSymbolKind.DOCUMENT}
    assert _control_errors("CommandButton1", **access) == [_AS_PROPERTY]


def test_1555_vb6_forms_keep_the_old_message() -> None:
    vb6 = {**_CONTROL_HOSTS["form"], "host": "vb6", "designer_class": "VB.Form"}
    assert _control_errors("CommandButton1", **vb6) == [
        "unknown-call CommandButton1: Sub or Function not defined: 'CommandButton1'."
    ]


# -- a40a5501: accept the Date statement under Option Explicit (#1664) -------

_DEFAULT_REFERENCES = {
    "excel": ["stdole", "Office"], "word": ["stdole", "Normal", "Office"],
    "powerpoint": ["stdole", "Office"], "access": ["stdole", "DAO"],
}
_PROJECT_NAMES = {"excel": "VBAProject", "word": "Project", "powerpoint": "VBAProject", "access": "Database1"}
_DATE_FIXTURES = [raw for raw in _RAW_CASES if raw["id"].startswith("date_statement_")]


def _date_codes(source: str, host: str = "excel", references: list[str] | None = None) -> list[str]:
    """Each error code, with the project known as the extension knows it, which
    turns on the undeclared-name check."""
    index = ProjectIndex(ProjectIndexOptions(ConditionalCompilationEnvironment(project_constants={})))
    index.set_module(ModuleInput("Module1", ModuleSymbolKind.STANDARD, source))
    opts = analyze_module_options_for(
        index, "Module1", ModuleSymbolKind.STANDARD, host=host,
        referenced_libraries=references if references is not None else _DEFAULT_REFERENCES[host],
        project_name=_PROJECT_NAMES[host],
    )
    return [d.code for d in analyze_module(source, opts) if d.severity.value == "error"]


def _date_procedure(*lines: str) -> str:
    return "\r\n".join(
        ["Option Explicit", "Sub Main()", "    Dim d As Date", *(f"    {line}" for line in lines), "End Sub", ""]
    )


def test_1664_fixtures_cover_every_office_host() -> None:
    assert {raw["host"] for raw in _DATE_FIXTURES} == {"excel", "word", "powerpoint", "access"}


@pytest.mark.parametrize("raw", _DATE_FIXTURES, ids=[raw["id"] for raw in _DATE_FIXTURES])
def test_1664_date_statement_fixture(raw: dict) -> None:
    expected = [] if raw["expected"] == "accepted" else ["undeclared-variable"]
    assert _date_codes(raw["modules"][0]["source"], raw["host"]) == expected


@pytest.mark.parametrize("lines", [
    ("date = d",), ("d = Now: Date = d",), ("Here: Date = d",), ("10 Date = d",), ("Date _", "    = d"),
    ("Date = #1/1/2020#",), ("Date = Date",), ("Date() = d",), ('Date = "let me"',),
])
def test_1664_date_statement_is_accepted(lines: tuple[str, ...]) -> None:
    assert _date_codes(_date_procedure(*lines)) == []


@pytest.mark.parametrize("lines", [
    ("Date = nope",), ("Here: Let Date = d",), ("10 Let Date = d",), ("Dim o As Object", "Set Date = o"),
])
def test_1664_its_value_and_let_or_set_date_still_report(lines: tuple[str, ...]) -> None:
    assert _date_codes(_date_procedure(*lines)) == ["undeclared-variable"]


# -- fac65b6d: stdole's LoadPicture, SavePicture and constants (#1665) --------

_STDOLE_FIXTURES = [raw for raw in _RAW_CASES if raw["id"].startswith("stdole_name_")]
# What the analyzer reports for each refused fixture.
_STDOLE_REFUSED_AS = {
    "stdfunctions_bare": "undeclared-variable",
    "stdfunctions_nope": "undeclared-variable",
    "savepicture_one": "argument-count",
    "savepicture_three": "argument-count",
    "savepicture_none": "argument-count",
    "savepicture_value": "undeclared-variable",
    "checked_assign": "undeclared-variable",
    "color_assign": "undeclared-variable",
    "for_gray": "undeclared-variable",
    "input_color": "undeclared-variable",
    "const_loadpicture": "undeclared-variable",
}
# Refused by the VBE and not reported, as upstream has it: LoadPicture's
# arguments are checked in a statement call only, and Access's own
# Application.LoadPicture is not arity-checked.
_STDOLE_KNOWN_QUIET: dict[str, tuple[str, ...] | None] = {
    "loadpicture_five": None,
    "loadpicture_four": ("access",),
    "loadpicture_noarg": ("access",),
    "const_loadpicture": ("access",),
}


def _stdole_procedure(*lines: str) -> str:
    return "\r\n".join([
        "Option Explicit", "Sub Main()", "    Dim p As Object, v As Variant",
        *(f"    {line}" for line in lines), "End Sub", "",
    ])


def test_1665_fixtures_cover_every_office_host() -> None:
    assert {raw["host"] for raw in _STDOLE_FIXTURES} == {"excel", "word", "powerpoint", "access"}


@pytest.mark.parametrize("raw", _STDOLE_FIXTURES, ids=[raw["id"] for raw in _STDOLE_FIXTURES])
def test_1665_stdole_fixture(raw: dict) -> None:
    kind = re.sub(r"_compile$", "", re.sub(r"^stdole_name_(excel|word|powerpoint|access)_", "", raw["id"]))
    quiet_hosts = _STDOLE_KNOWN_QUIET.get(kind, ())
    quiet = kind in _STDOLE_KNOWN_QUIET and (quiet_hosts is None or raw["host"] in quiet_hosts)
    found = _date_codes(raw["modules"][0]["source"], raw["host"])
    assert found == ([] if raw["expected"] == "accepted" or quiet else [_STDOLE_REFUSED_AS[kind]])


@pytest.mark.parametrize("lines", [
    ('Set p = loadpicture("")',), ("Set p = LoadPicture",), ('Call SavePicture(p, "x")',),
    ("If v = GRAY Then v = vgacolor",),
    ("Set p = StdFunctions _", '    .LoadPicture("")'), ('Set p = StdFunctions.[LoadPicture]("")',),
    ('Call SavePicture(p, "x")',), ("If v Then _", '    SavePicture p, "x"'),
    ("If v = Color Then v = Gray",), ('If v Then Set p = LoadPicture("") Else LoadPicture ""',),
    ("Dim x As Single", 'If x! Then SavePicture p, "x"'),
    ("Get #1, Checked, v",), ("Input #Checked, v",), ("Put #1, , Checked",), ("Input #1, v(Checked)",),
])
def test_1665_stdole_names_accepted(lines: tuple[str, ...]) -> None:
    assert _date_codes(_stdole_procedure(*lines)) == []


@pytest.mark.parametrize("lines", [
    ("v = SavePicture",), ("v = StdFunctions.Nope",), ('StdFunctions.SavePicture p, "x"',),
    ("Input #1, _", "    Color"), ("For _", "    Gray = 1 To 3", "Next"),
    ("If v Then Color = 1",), ("If v Then Set Default = p",), ("If v Then v = 1 Else Checked = 2",),
    ("If v Then _", "    Color = 1"), ('LSet Color = "x"',), ('Mid(Color, 1, 1) = "x"',),
    ('Mid$(Color, 1) = "x"',), ('If v Then MidB$(Gray, 1) = "x"',), ("Color(1) = 1",),
    ("If v Then Color(1) = 1",), ("If v Then SavePicture = 1",), ("If v Then v = 1 Else SavePicture = 1",),
    ("Color(v(1)) = 1",), ("If v Then Checked",), ("If v Then Color 1, 2",), ("If v Then Call Checked",),
    ("Color(v(v(v(1)))) = 1",), ('Color(v("(")) = 1',),
    ("Dim x As Single", "If x! Then Color = 1"), ("Dim x As Single", "If True Then x = x! Else Color = 1"),
    ("Get #1, , Checked",), ("Input #1, v, Color",), ("Line Input #1, Gray",),
    ('Input #1, v(":"), Color',), ("Input #1, v.Then, Color",), ("Get #1, v!Else, Gray",),
    ('Set p = LoadPictures("")',), ("v = Grey",),
])
def test_1665_stdole_names_written_or_misused_still_report(lines: tuple[str, ...]) -> None:
    assert _date_codes(_stdole_procedure(*lines)) == ["undeclared-variable"]


def test_1665_stdole_names_stay_known_without_the_reference() -> None:
    assert _date_codes(_stdole_procedure('Set p = LoadPicture("")'), "excel", ["Office"]) == []


def test_1665_a_bare_stdfunctions_const_value_is_undeclared() -> None:
    assert _date_codes("Option Explicit\r\nConst K = StdFunctions\r\n") == ["undeclared-variable"]


def test_1665_a_hosts_own_loadpicture_keeps_its_arguments() -> None:
    opts = AnalyzeModuleOptions(host="vb6", module_name="Module1", module_kind=ModuleSymbolKind.STANDARD)
    found = analyze_module(_stdole_procedure('LoadPicture "", 0, 0, 1, 1'), opts)
    assert [d for d in found if d.code == "argument-count"] == []
    assert _date_codes(_stdole_procedure('Set p = LoadPicture("", 1, 1, 0)'), "access") == []


@pytest.mark.parametrize("lines, expected", [
    (("SavePicture",), ["argument-count"]),
    (("Call SavePicture()",), ["argument-count"]),
    (('Call SavePicture(p, "x")',), []),
])
def test_1665_savepicture_arguments_follow_stdole(lines: tuple[str, ...], expected: list[str]) -> None:
    assert _date_codes(_stdole_procedure(*lines)) == expected
