"""Rule contracts from the XLIDE 11.2.1 -> 11.2.11 sync (calls, events,
conditional compilation and the rest).

Each section names the upstream commit whose own tests it follows.
"""

import json
import re

import pytest

from pyvbaanalysis import analyze_module, analyze_module_options_for
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.diagnostics.rule_metadata import diagnostic_metadata_for_code
from pyvbaanalysis.evidence import DATA_DIR
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex

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
    return ModuleInput(
        code_name if code_name is not None else module["name"],
        _KINDS.get(module.get("type", "standard"), ModuleSymbolKind.STANDARD),
        module["source"],
        designer_class=module.get("designerClass"),
        document_type=document_type if code_name is not None else None,
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
