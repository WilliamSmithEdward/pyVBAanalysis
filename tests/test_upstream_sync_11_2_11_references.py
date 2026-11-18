"""Referenced libraries and referenced VBA projects, from the XLIDE 11.2.11 sync.

Each section names the upstream commit it ports; the cases are upstream's own
tests for that commit and the oracle fixtures they read.
"""

from __future__ import annotations

import json
import re
from typing import Any

import pytest
from oracle_support import oracle_module_input, oracle_project_references

from pyvbaanalysis import analyze_module, analyze_module_options_for
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.diagnostics.model import DiagnosticSeverity
from pyvbaanalysis.evidence import DATA_DIR
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex, ProjectIndexOptions
from pyvbaanalysis.symbols.referenced_projects import (
    ReferencedProjectModuleSource,
    ReferencedProjectSurface,
    referenced_project_names,
    referenced_project_surface,
)

_ORACLE: list[dict[str, Any]] = json.loads((DATA_DIR / "vbe_oracle_cases.json").read_text(encoding="utf-8"))["cases"]
_BY_ID = {case["id"]: case for case in _ORACLE}

_KIND = {
    "standard": ModuleSymbolKind.STANDARD,
    "class": ModuleSymbolKind.CLASS,
    "document": ModuleSymbolKind.DOCUMENT,
    "userform": ModuleSymbolKind.USERFORM,
}

# The names a new file of each host records, and its project's name (measured,
# upstream issue #1413).
_DEFAULTS = {
    "excel": ["stdole", "Office"],
    "word": ["stdole", "Normal", "Office"],
    "powerpoint": ["stdole", "Office"],
    "access": ["stdole", "DAO"],
}
_PROJECT_NAMES = {"excel": "VBAProject", "word": "Project", "powerpoint": "VBAProject", "access": "Database1"}
_LIBRARY_NAMES = {
    "{420B2830-E718-11CF-893D-00A0C9054228}": "Scripting",
    "{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}": "DAO",
}


def _module(*body: str) -> str:
    return "\r\n".join(
        ["Option Explicit", "Public Sub Main()", "    Dim x As Long, v As Variant", *(f"    {line}" for line in body), "End Sub", ""]
    )


def _errors(source: str, host: str = "excel", **known: Any) -> list[str]:
    """Each error, as `code: message`, of a lone standard module analyzed as the
    extension analyzes it with its project (upstream referenceQualifiers1495)."""
    index = ProjectIndex(ProjectIndexOptions(conditional_compilation=ConditionalCompilationEnvironment(project_constants={})))
    index.set_module(ModuleInput("Module1", ModuleSymbolKind.STANDARD, source))
    opts = analyze_module_options_for(
        index, "Module1", ModuleSymbolKind.STANDARD, host=host,
        conditional_compilation=ConditionalCompilationEnvironment(project_constants={}), **known,
    )
    return [f"{d.code}: {d.message}" for d in analyze_module(source, opts) if d.severity is DiagnosticSeverity.ERROR]


# 6a261810: a referenced library's name qualifies before a dot
# (referenceQualifiers1495.test.ts).

_ISSUE_1495 = [case for case in _ORACLE if case["id"].startswith("issue1495_")]


def test_issue_1495_has_every_measured_case() -> None:
    assert len(_ISSUE_1495) == 61


@pytest.mark.parametrize("case", _ISSUE_1495, ids=[case["id"] for case in _ISSUE_1495])
def test_issue_1495_agrees_with_the_vbe(case: dict[str, Any]) -> None:
    host = case.get("host") or "excel"
    found = _errors(
        case["modules"][0]["source"], host,
        referenced_libraries=[*_DEFAULTS[host], *(_LIBRARY_NAMES[ref["guid"]] for ref in case.get("references", []))],
        project_name=_PROJECT_NAMES[host],
    )
    if case["expected"] == "accepted":
        assert found == []
    else:
        assert len(found) == 1
        assert found[0].startswith("undeclared-variable: ")


def test_a_bare_library_name_is_a_library() -> None:
    assert _errors(_module("v = stdole"), referenced_libraries=["stdole"]) == [
        "undeclared-variable: 'stdole' names a library or project, which stands only before a dot. "
        "This is a VBE compile error: Expected variable or procedure, not project."
    ]


def test_a_variable_declared_with_a_library_name_is_kept() -> None:
    source = "\r\n".join(["Option Explicit", "Public Sub Main()", "    Dim office As Long", "    office = 1", "    Debug.Print office", "End Sub", ""])
    assert _errors(source) == []


def test_unread_references_leave_the_rule_as_it_was() -> None:
    assert len(_errors(_module("x = stdole.Checked"))) == 1
    assert _errors(_module("x = Excel.xlUp")) == []


def test_a_type_and_the_next_statements_dot_are_not_one_qualified_name() -> None:
    known: dict[str, Any] = {"referenced_libraries": ["stdole", "Office"], "project_name": "VBAProject"}
    with_new = ["Option Explicit", "Function Main() As Variant", "    With New Collection", "        .Add 1", "        Main = .Count", "    End With", "End Function", ""]
    assert _errors("\r\n".join(with_new), **known) == []
    with_dim = ["Option Explicit", "Sub Main()", "    With Application", "        Dim c As Collection", "        .StatusBar = False", "    End With", "End Sub", ""]
    assert _errors("\r\n".join(with_dim), **known) == []
    continued = ["Option Explicit", "Sub Main()", "    Dim c As Nowhere _", "        .Thing", "End Sub", ""]
    assert [e.split(":")[0] for e in _errors("\r\n".join(continued), **known)] == ["missing-library-reference"]


def test_another_vba_projects_name_qualifies_read_or_not() -> None:
    # e4905576 (issue #1621): the VBE compiles this (oracle
    # issue1503_qualified_compile). Unread, what follows the dot is left unjudged,
    # and a bare name too.
    known: dict[str, Any] = {
        "referenced_libraries": ["stdole", "Office", "ProbeLibrary"],
        "referenced_projects": ["ProbeLibrary"],
        "project_name": "VBAProject",
    }
    assert _errors(_module("x = ProbeLibrary.Service.Value()"), **known) == []
    assert _errors(_module("x = Nonexistent"), **known) == []
    # With every reference known, a name nothing supplies is still reported.
    type_libraries_only: dict[str, Any] = {"referenced_libraries": ["stdole", "Office"], "project_name": "VBAProject"}
    assert [e.split(":")[0] for e in _errors(_module("x = Nonexistent"), **type_libraries_only)] == ["undeclared-variable"]


def test_the_options_carry_referenced_projects() -> None:
    assert AnalyzeModuleOptions(referenced_projects=["ProbeLibrary"]).referenced_projects == ["ProbeLibrary"]


# The oracle's referenced-project fixtures, analyzed as upstream's tests analyze
# them: the case's modules as the project, its references read from the file the
# oracle compiled it in, and the referenced project read as the engine reads one.


def _case_errors(case: dict[str, Any], change: Any = None) -> list[str]:
    host = case.get("host") or "excel"
    inputs = [oracle_module_input(module) for module in case["modules"]]
    index = ProjectIndex(ProjectIndexOptions(conditional_compilation=ConditionalCompilationEnvironment(project_constants={})))
    for module_input in inputs:
        index.set_module(module_input)
    references = oracle_project_references(case["id"])
    if change is not None:
        references = change(references)
    out: list[str] = []
    for module_input in inputs:
        opts = analyze_module_options_for(
            index, module_input.module_name, module_input.module_kind, host=host,
            conditional_compilation=ConditionalCompilationEnvironment(project_constants={}), **references,
        )
        out.extend(
            f"{d.code}: {d.message}" for d in analyze_module(module_input.source, opts) if d.severity is DiagnosticSeverity.ERROR
        )
    return out


def _unread(references: dict[str, Any]) -> dict[str, Any]:
    return {**references, "referenced_project_surfaces": []}


# 81216992: read a referenced VBA project to judge names reached through it
# (referencedProjects1503.test.ts), with e4905576's change to its unread case.

_ISSUE_1503 = [case for case in _ORACLE if case["id"].startswith("issue1503_")]


def test_issue_1503_has_every_measured_case() -> None:
    assert len(_ISSUE_1503) == 76


@pytest.mark.parametrize("case", _ISSUE_1503, ids=[case["id"] for case in _ISSUE_1503])
def test_issue_1503_agrees_with_the_vbe(case: dict[str, Any]) -> None:
    found = _case_errors(case)
    if case["expected"] == "accepted":
        assert found == []
    else:
        assert len(found) > 0


def test_issue_1503_says_what_the_referenced_project_lacks() -> None:
    assert _case_errors(_BY_ID["issue1503_missing_member_compile"]) == [
        "member-not-found: Method or data member not found: 'ProbeLibrary.Service.Nope'."
    ]
    assert _case_errors(_BY_ID["issue1503_private_module_project_member_compile"]) == [
        "member-not-found: Method or data member not found: 'ProbeLibrary.Value'."
    ]
    assert _case_errors(_BY_ID["issue1503_private_class_type_compile"])[0].startswith(
        "missing-library-reference: 'ProbeLibrary' is a VBA project this project references, "
        "and it offers no type 'Thing'"
    )


def test_issue_1503_leaves_an_unread_project_and_its_names_unjudged() -> None:
    assert _case_errors(_BY_ID["issue1503_qualified_compile"], _unread) == []


def test_a_referenced_surface_hides_a_private_module() -> None:
    surface = referenced_project_surface(
        "ProbeLibrary",
        [
            ReferencedProjectModuleSource("Service", ModuleSymbolKind.STANDARD, "Option Explicit\r\nPublic Function Value() As Long\r\nEnd Function\r\n"),
            ReferencedProjectModuleSource("Secret", ModuleSymbolKind.STANDARD, "Option Private Module\r\nPublic Function Hidden() As Long\r\nEnd Function\r\n"),
        ],
    )
    service, secret = surface.modules
    assert (service.kind, service.visible, service.members) == (ModuleSymbolKind.STANDARD, True, ("Value",))
    assert (secret.visible, secret.members) == (False, ())


# fdeb8b76, ea5567ad, 0728873c, 826aba31, 84a193a3, a7332160: the classes of a
# referenced VBA project (referencedProjectClasses.test.ts).

_CLASSES = [case for case in _ORACLE if case["id"].startswith("referenced_class_")]
# Word's cases also reference Normal, before ProbeLibrary, as a project XLIDE has
# not read: a bare `Thing`, `Widget` or `Solo` could be Normal's.
_UNCHECKED = re.compile(
    r"^referenced_class_word_(private_bare_type|new_bare|predeclared_missing|enum_bare_missing|"
    r"module_enum_bare_missing|bare_missing_member)_compile$"
)


def test_referenced_classes_have_every_measured_case_in_all_four_hosts() -> None:
    assert len(_CLASSES) == 248
    for host in (None, "word", "powerpoint", "access"):
        assert len([case for case in _CLASSES if case.get("host") == host]) == 62


@pytest.mark.parametrize(
    "case", [c for c in _CLASSES if not _UNCHECKED.match(c["id"])], ids=[c["id"] for c in _CLASSES if not _UNCHECKED.match(c["id"])]
)
def test_referenced_class_agrees_with_the_vbe(case: dict[str, Any]) -> None:
    found = _case_errors(case)
    if case["expected"] == "accepted":
        assert found == []
        return
    assert len(found) == 1
    # Each says what the VBE says, but `WithEvents ... As New`, which the WithEvents
    # rule words as it does for any class, and an Implements, which says what the
    # VBE says and then what to add.
    match = re.search(r"Compile error: (.+?)\. Measured", case["description"])
    vbe = match.group(1) if match else None
    if case["id"].endswith("_withevents_new_compile"):
        assert found[0].startswith("withevents-declaration: WithEvents variable 'mW' cannot be declared As New.")
        return
    if case["id"].endswith("_implements_missing_compile"):
        assert found[0] == (
            f"implements-member-missing: {vbe}: add a Property 'Widget_Value', and 'Widget_Child', "
            "'Widget_Size', 'Widget_Twice', 'Widget_Poke'."
        )
        return
    assert (f"{vbe}: " if vbe in ("Variable not defined", "Method or data member not found") else f"VBE compile error: {vbe}.") in found[0]


@pytest.mark.parametrize("case", [c for c in _CLASSES if _UNCHECKED.match(c["id"])], ids=[c["id"] for c in _CLASSES if _UNCHECKED.match(c["id"])])
def test_referenced_class_is_left_unreported(case: dict[str, Any]) -> None:
    assert _case_errors(case) == []


_HOSTS = ["excel", "word", "powerpoint", "access"]
_EXPOSED = "Attribute VB_Exposed = True\r\nOption Explicit\r\n"
_PREDECLARED = f"Attribute VB_PredeclaredId = True\r\n{_EXPOSED}"


def _surface(name: str, *modules: tuple[str, ModuleSymbolKind, str]) -> ReferencedProjectSurface:
    return referenced_project_surface(name, [ReferencedProjectModuleSource(*module) for module in modules])


def _run(
    host: str,
    code: str,
    libraries: list[str],
    surfaces: list[ReferencedProjectSurface],
    params: str = "",
    referenced_projects: list[str] | None = None,
) -> list[str]:
    source = f"Option Explicit\r\nPublic Sub Main({params})\r\n    Dim x As Variant\r\n    {code}\r\nEnd Sub\r\n"
    opts = AnalyzeModuleOptions(
        module_name="Module1", module_kind=ModuleSymbolKind.STANDARD, host=host, project_name="VBAProject",
        referenced_libraries=["stdole", "Office", *libraries],
        referenced_projects=referenced_projects if referenced_projects is not None else [s.name for s in surfaces],
        referenced_project_surfaces=surfaces,
    )
    return [d.message for d in analyze_module(source, opts) if d.severity is DiagnosticSeverity.ERROR]


_FIRST = _surface("First", ("Palette", ModuleSymbolKind.CLASS, f"{_EXPOSED}Public Enum Solo\r\n    sOne = 1\r\nEnd Enum\r\n"))
_SECOND = _surface(
    "Second",
    ("Solo", ModuleSymbolKind.STANDARD, "Option Explicit\r\nPublic Function Good() As Long\r\nEnd Function\r\n"),
    ("Widget", ModuleSymbolKind.CLASS, _EXPOSED),
)


@pytest.mark.parametrize("host", _HOSTS)
def test_a_bare_name_reaches_the_earlier_projects_enum(host: str) -> None:
    assert _run(host, "x = Solo.sOne", ["First", "Second"], [_FIRST, _SECOND]) == []
    assert _run(host, "x = Solo.Nope", ["First", "Second"], [_FIRST, _SECOND]) == ["Method or data member not found: 'Solo.Nope'."]


@pytest.mark.parametrize("host", _HOSTS)
def test_a_bare_name_reaches_the_module_when_its_project_comes_first(host: str) -> None:
    assert _run(host, "x = Solo.Good()", ["Second", "First"], [_FIRST, _SECOND]) == []
    assert _run(host, "x = Solo.sOne", ["Second", "First"], [_FIRST, _SECOND]) == ["Method or data member not found: 'Solo.sOne'."]


@pytest.mark.parametrize("host", _HOSTS)
def test_a_bare_name_is_unjudged_behind_an_unread_library(host: str) -> None:
    assert _run(host, "x = Solo.Nope", ["Unread", "First"], [_FIRST]) == []
    assert _run(host, "Dim w As Object\r\n    Set w = New Widget", ["Unread", "Second"], [_SECOND]) == []


@pytest.mark.parametrize("host", _HOSTS)
def test_new_on_a_referenced_class_is_invalid(host: str) -> None:
    found = _run(host, "Dim w As Object\r\n    Set w = New Widget", ["Second"], [_SECOND])
    assert "Invalid use of New keyword" in ",".join(found)


_SOLO = ("Solo", ModuleSymbolKind.CLASS, f"{_PREDECLARED}Public Value As Long\r\n")


def _projects(host: str, code: str, projects: dict[str, list[tuple[str, ModuleSymbolKind, str]]], params: str = "") -> list[str]:
    surfaces = [_surface(name, *modules) for name, modules in projects.items()]
    return _run(host, code, list(projects), surfaces, params)


@pytest.mark.parametrize("host", _HOSTS)
def test_an_earlier_projects_enum_comes_before_a_later_default_instance(host: str) -> None:
    palette = ("Palette", ModuleSymbolKind.CLASS, f"{_EXPOSED}Public Enum Solo\r\n    sOne = 1\r\nEnd Enum\r\n")
    assert _projects(host, "x = Solo.sOne", {"First": [palette], "Second": [_SOLO]}) == []


@pytest.mark.parametrize("host", _HOSTS)
def test_an_earlier_hidden_module_does_not_stop_a_later_default_instance(host: str) -> None:
    hidden = ("Solo", ModuleSymbolKind.CLASS, "Option Explicit\r\n")
    secret = ("Solo", ModuleSymbolKind.STANDARD, "Option Private Module\r\nOption Explicit\r\n")
    for earlier in (hidden, secret):
        assert _projects(host, "x = Solo.Value", {"First": [earlier], "Second": [_SOLO]}) == []
        assert "Nope" in ",".join(_projects(host, "x = Solo.Nope", {"First": [earlier], "Second": [_SOLO]}))


@pytest.mark.parametrize("host", _HOSTS)
def test_a_run_time_function_comes_before_a_default_instance_of_its_name(host: str) -> None:
    shadow = ("CreateObject", ModuleSymbolKind.CLASS, f"{_PREDECLARED}Public Value As Long\r\n")
    assert _projects(host, 'x = CreateObject("Excel.Application").Visible', {"ProbeLibrary": [shadow]}) == []


@pytest.mark.parametrize("host", _HOSTS)
def test_a_parameterless_item_indexes_what_it_returns(host: str) -> None:
    modules = [
        ("Holder", ModuleSymbolKind.CLASS, f"{_EXPOSED}Public Function Item() As Widgets\r\nEnd Function\r\nPublic Function Make(ByVal n As Long) As Widgets\r\nEnd Function\r\n"),
        ("Widgets", ModuleSymbolKind.CLASS, f"{_EXPOSED}Public Own As Long\r\nPublic Function Get1(ByVal i As Long) As Object\r\nAttribute Get1.VB_UserMemId = 0\r\nEnd Function\r\n"),
    ]
    params = "ByVal h As ProbeLibrary.Holder"
    assert _projects(host, "x = h.Item(1).Anything", {"ProbeLibrary": modules}, params) == []
    assert _projects(host, "x = h.Make(1).Own", {"ProbeLibrary": modules}, params) == []
    assert "Nope" in ",".join(_projects(host, "x = h.Make(1).Nope", {"ProbeLibrary": modules}, params))


@pytest.mark.parametrize("host", _HOSTS)
@pytest.mark.parametrize("eol", ["\n", "\r\n", "\r"])
def test_a_continued_member_type_is_read(host: str, eol: str) -> None:
    modules = [("Widget", ModuleSymbolKind.CLASS, f"{_EXPOSED}Public Child As ProbeLibrary. _{eol}    Widget\r\nPublic Value As Long\r\n")]
    params = "ByVal w As ProbeLibrary.Widget"
    assert _projects(host, "x = w.Child.Value", {"ProbeLibrary": modules}, params) == []
    assert "Nope" in ",".join(_projects(host, "x = w.Child.Nope", {"ProbeLibrary": modules}, params))


_ENUMS = [
    ("Widget", ModuleSymbolKind.CLASS, f"{_EXPOSED}Public Enum WidgetColor\r\n    wcRed = 1\r\nEnd Enum\r\n"),
    ("Service", ModuleSymbolKind.STANDARD, "Option Explicit\r\nPublic Enum ServiceKind\r\n    skOne = 1\r\nEnd Enum\r\n"),
]


@pytest.mark.parametrize("host", _HOSTS)
def test_the_member_named_after_a_referenced_enum_is_checked(host: str) -> None:
    run = lambda code: _projects(host, code, {"ProbeLibrary": _ENUMS})  # noqa: E731
    assert run("x = ProbeLibrary.WidgetColor.wcRed + ProbeLibrary.ServiceKind.skOne") == []
    assert run("x = ProbeLibrary.WidgetColor.Nope") == ["Method or data member not found: 'ProbeLibrary.WidgetColor.Nope'."]
    assert run("x = ProbeLibrary.ServiceKind.Nope") == ["Method or data member not found: 'ProbeLibrary.ServiceKind.Nope'."]


@pytest.mark.parametrize("host", _HOSTS)
def test_an_exposed_document_module_is_a_type_with_its_members_unjudged(host: str) -> None:
    sheet = ("Sheet1", ModuleSymbolKind.DOCUMENT, f"{_EXPOSED}Public Sub Hello()\r\nEnd Sub\r\n")
    code = 'Dim s As ProbeLibrary.Sheet1\r\n    ProbeLibrary.Sheet1.Hello\r\n    x = ProbeLibrary.Sheet1.Range("A1")'
    assert _projects(host, code, {"ProbeLibrary": [sheet]}) == []
    module = _surface("ProbeLibrary", sheet).modules[0]
    assert (module.kind, module.visible, module.members, module.types, module.enums) == (
        ModuleSymbolKind.DOCUMENT, True, (), ("Sheet1",), ()
    )
    assert module.object_members is None


@pytest.mark.parametrize("host", _HOSTS)
def test_an_exposed_document_module_offers_its_enums(host: str) -> None:
    source = f"{_EXPOSED}Public Enum SheetKind\r\n    shOne = 1\r\nEnd Enum\r\n"
    sheet = ("Sheet1", ModuleSymbolKind.DOCUMENT, source)
    code = "Dim s As ProbeLibrary.SheetKind\r\n    x = ProbeLibrary.SheetKind.shOne + ProbeLibrary.shOne + shOne"
    assert _projects(host, code, {"ProbeLibrary": [sheet]}) == []
    assert _projects(host, "x = ProbeLibrary.SheetKind.Nope", {"ProbeLibrary": [sheet]}) == [
        "Method or data member not found: 'ProbeLibrary.SheetKind.Nope'."
    ]
    hidden = _surface("ProbeLibrary", ("Sheet1", ModuleSymbolKind.DOCUMENT, source.replace("Attribute VB_Exposed = True\r\n", ""))).modules[0]
    assert (hidden.visible, hidden.members, hidden.types, hidden.enums) == (False, (), (), None)


# 5a69aedc: the name index is built once per surface and shared
# (referencedProjectNameIndex.test.ts).


def test_the_name_index_is_shared_per_surface() -> None:
    surface = _surface("ProbeLibrary", *_ENUMS)
    assert referenced_project_names(surface) is referenced_project_names(surface)
    other = _surface("ProbeLibrary", *_ENUMS)
    assert referenced_project_names(other) is not referenced_project_names(surface)


# e7b2c218: a referenced type split by a line continuation
# (referencedTypeContinuations.test.ts).

_PROVIDER = _surface(
    "ProbeLibrary",
    ("Widget", ModuleSymbolKind.CLASS, _EXPOSED),
    ("Thing", ModuleSymbolKind.CLASS, "Option Explicit\r\n"),
)


def _type_errors(host: str, type_text: str, eol: str) -> list[Any]:
    source = f"Option Explicit\nPublic Sub Main()\n    Dim w As {type_text}\nEnd Sub\n".replace("\n", eol)
    opts = AnalyzeModuleOptions(
        module_name="Module1", module_kind=ModuleSymbolKind.STANDARD, host=host,
        referenced_libraries=[*_DEFAULTS[host], "ProbeLibrary"], project_name=_PROJECT_NAMES[host],
        referenced_projects=["ProbeLibrary"], referenced_project_surfaces=[_PROVIDER],
    )
    return [d for d in analyze_module(source, opts) if d.severity is DiagnosticSeverity.ERROR]


@pytest.mark.parametrize("host", _HOSTS)
@pytest.mark.parametrize("eol", ["\n", "\r\n", "\r"])
def test_a_continued_exposed_type_is_accepted(host: str, eol: str) -> None:
    for type_text in (
        "ProbeLibrary.Widget",
        "ProbeLibrary. _\n        Widget",
        "ProbeLibrary. _\u3000\n        Widget",
        "ProbeLibrary. _\u0019\n        Widget",
        "[ProbeLibrary]. _\n        [Widget]",
        "ProbeLibrary _\n        . _\n        Widget",
    ):
        assert _type_errors(host, type_text, eol) == [], type_text


@pytest.mark.parametrize("host", _HOSTS)
@pytest.mark.parametrize("type_name", ["Thing", "Missing"])
def test_a_continued_inaccessible_type_is_rejected(host: str, type_name: str) -> None:
    for eol in ("\n", "\r\n", "\r"):
        errors = _type_errors(host, f"ProbeLibrary. _\n        {type_name}", eol)
        assert len(errors) == 1
        assert errors[0].code == "missing-library-reference"
        assert f"offers no type '{type_name}'" in errors[0].message


# e4905576: a bare name behind an unread referenced project
# (referencedProjectsUnread1621.test.ts).

_ACCEPTED_WITH_PROJECT = [case for case in _ORACLE if case.get("projectReference") and case["expected"] == "accepted"]


def test_unread_cases_cover_every_host() -> None:
    assert sorted({case.get("host") or "excel" for case in _ACCEPTED_WITH_PROJECT}) == _HOSTS_SORTED


_HOSTS_SORTED = ["access", "excel", "powerpoint", "word"]


@pytest.mark.parametrize("case", _ACCEPTED_WITH_PROJECT, ids=[case["id"] for case in _ACCEPTED_WITH_PROJECT])
def test_an_accepted_referenced_project_case_is_clean_read_or_unread(case: dict[str, Any]) -> None:
    assert _case_errors(case) == []
    assert _case_errors(case, _unread) == []


def _module_run(code: str, **references: Any) -> list[str]:
    source = f"Option Explicit\r\nPublic Sub Main()\r\n    Dim x As Variant\r\n    {code}\r\nEnd Sub\r\n"
    return [e.split(": ", 1)[1] for e in _errors(source, "excel", project_name="VBAProject", **references)]


def test_bare_names_and_calls_are_unjudged_behind_an_unread_project() -> None:
    unread: dict[str, Any] = {
        "referenced_libraries": ["stdole", "Office", "Unread"],
        "referenced_projects": ["Unread"],
        "referenced_project_surfaces": [],
    }
    assert _module_run("x = Widget", **unread) == []
    assert _module_run("Call Helper(1)", **unread) == []
    assert _module_run("Helper 1", **unread) == []
    assert _module_run("x = Unread.Service.Value()", **unread) == []


def test_a_name_nothing_supplies_is_reported_where_every_reference_is_known() -> None:
    known: dict[str, Any] = {"referenced_libraries": ["stdole", "Office"]}
    assert _module_run("x = Nonexistent", **known)[0].startswith("Variable not defined: 'Nonexistent'")
    assert _module_run("Helper 1", **known)[0].startswith("Sub or Function not defined")


def test_a_bare_class_name_stays_reported_when_only_a_read_project_has_it() -> None:
    case = next(c for c in _ACCEPTED_WITH_PROJECT if c["projectReference"]["name"] == "ProbeLibrary")
    references = oracle_project_references(case["id"])
    references.pop("project_name", None)
    assert len(_module_run("x = Widget", **references)) > 0
