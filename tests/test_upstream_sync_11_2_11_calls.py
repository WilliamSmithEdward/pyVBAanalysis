"""Rule contracts from the XLIDE 11.2.1 -> 11.2.11 sync (calls, events,
conditional compilation and the rest).

Each section names the upstream commit whose own tests it follows.
"""

import json
import re

import pytest

from pyvbaanalysis import analyze_module, analyze_module_options_for, host_token_for_file_name
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment, parse_project_conditional_constants
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.diagnostics.rule_metadata import diagnostic_metadata_for_code
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

_ISSUE_1493 = [
    raw for raw in _RAW_CASES if raw["id"].startswith("issue1493_") and "longptr" not in raw["id"]
]


def test_1493_has_the_measured_cases() -> None:
    assert len(_ISSUE_1493) >= 118


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
