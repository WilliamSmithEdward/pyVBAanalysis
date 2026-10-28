"""Host, library and document-module contracts from the XLIDE 11.2.1 sync.

Each section names the upstream commit it ports; the cases are upstream's own
tests for that commit.
"""

from __future__ import annotations

import json
from typing import Any

import pytest

from pyvbaanalysis import analyze_module, analyze_module_options_for
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.diagnostics.model import DiagnosticSeverity
from pyvbaanalysis.evidence import DATA_DIR
from pyvbaanalysis.host.host_libraries import host_token_for_libid
from pyvbaanalysis.host.host_model import get_host_global_members, resolve_host_global_member
from pyvbaanalysis.host.host_registry import host_object_model_for_tokens
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex, ProjectIndexOptions

_ORACLE: list[dict[str, Any]] = json.loads((DATA_DIR / "vbe_oracle_cases.json").read_text(encoding="utf-8"))["cases"]

_KIND = {
    "standard": ModuleSymbolKind.STANDARD,
    "class": ModuleSymbolKind.CLASS,
    "document": ModuleSymbolKind.DOCUMENT,
    "userform": ModuleSymbolKind.USERFORM,
}


def _codes(source: str, code: str, **options: Any) -> list[Any]:
    return [d for d in analyze_module(source, AnalyzeModuleOptions(**options)) if d.code == code]


# 3f5fc9aa: member-exists-in-base-class (baseClassMembers.test.ts).

_CLASH = "member-exists-in-base-class"
_BASE_MEMBER_CASES = [case for case in _ORACLE if case["id"].startswith("base_member_")]


def _module_options(host: str, module: dict[str, Any]) -> dict[str, Any]:
    """What the extension tells the analyzer about each kind of module."""
    if host == "access" and module["type"] == "document":
        report = module["name"].startswith("Report_")
        return {"module_kind": ModuleSymbolKind.USERFORM, "designer_class": "Access.Report" if report else "Access.Form"}
    if module["type"] == "document" and host == "excel" and module["name"] != "ThisWorkbook":
        return {"module_kind": ModuleSymbolKind.DOCUMENT, "document_type": "worksheet"}
    return {"module_kind": _KIND.get(module["type"], ModuleSymbolKind.STANDARD)}


def test_base_member_fixtures_cover_every_office_host() -> None:
    assert {case["host"] for case in _BASE_MEMBER_CASES} == {"excel", "word", "powerpoint", "access"}


@pytest.mark.parametrize("case", _BASE_MEMBER_CASES, ids=[case["id"] for case in _BASE_MEMBER_CASES])
def test_base_member_fixture(case: dict[str, Any]) -> None:
    clash = case["expected"] == "rejected" and "Member already exists" in case["description"]
    found = sum(
        len(_codes(module["source"], _CLASH, module_name=module["name"], host=case["host"], **_module_options(case["host"], module)))
        for module in case["modules"]
    )
    assert found == (1 if clash else 0)


_WORDS = "Public Function Words() As Long\n    Words = 1\nEnd Function\n"


def test_base_member_names_the_class_and_the_clash() -> None:
    [diagnostic] = _codes(_WORDS, _CLASH, host="word", module_name="ThisDocument", module_kind=ModuleSymbolKind.DOCUMENT)
    assert "'Words' is already a member of Word.Document" in diagnostic.message
    assert "Member already exists in an object module from which this object module derives" in diagnostic.message


def test_base_member_sheet_of_unknown_type_takes_what_both_kinds_share() -> None:
    def sheet(src: str) -> list[Any]:
        return _codes(src, _CLASH, module_name="Sheet1", module_kind=ModuleSymbolKind.DOCUMENT)

    # Name is on both; Range is a Worksheet's only, and the sheet may be a chart.
    assert len(sheet("Public Function Name() As String\nEnd Function\n")) == 1
    assert sheet("Public Sub Range()\nEnd Sub\n") == []


def test_base_member_leaves_standard_class_and_vb6_modules_alone() -> None:
    assert _codes(_WORDS, _CLASH, host="word", module_name="Module1", module_kind=ModuleSymbolKind.STANDARD) == []
    assert _codes(_WORDS, _CLASH, host="word", module_name="Class1", module_kind=ModuleSymbolKind.CLASS) == []
    caption = "Public Function Caption() As String\nEnd Function\n"
    assert _codes(caption, _CLASH, host="vb6", module_name="Form1", module_kind=ModuleSymbolKind.USERFORM, designer_class="VB.Form") == []


def test_base_member_skips_public_const_declare_and_array() -> None:
    src = (
        "Public Const Words As Long = 1\n"
        'Public Declare PtrSafe Function Words Lib "kernel32" Alias "GetTickCount" () As Long\n'
        "Public Words() As Long\n"
    )
    assert _codes(src, _CLASH, host="word", module_name="ThisDocument", module_kind=ModuleSymbolKind.DOCUMENT) == []


# 3f5fc9aa: invalid-as-type-name for a name none of a project's references holds,
# per Office host (hostReferenceTypes.test.ts).

_REFERENCES: dict[str, list[str]] = {
    "excel": ["stdole", "Office"],
    "word": ["stdole", "Normal", "Office"],
    "powerpoint": ["stdole", "Office"],
    "access": ["stdole", "DAO"],
}
_REFTYPE_CASES = [
    case
    for case in _ORACLE
    if case["id"].startswith("reftype_")
    or case["id"] == "host_excel_worksheet_type_compile"
    or case["id"] in {f"host_{h}_{k}_type_compile" for h in ("word", "powerpoint", "access") for k in ("excel_worksheet", "own_library")}
]
# Rejected by the VBE, not reported, and why: the Access model carries the Office
# library's types whether or not the database references Office.
_KNOWN_SILENT = {"reftype_access_office_commandbar_compile"}


def _unknown_types(case: dict[str, Any]) -> int:
    host = case.get("host") or "excel"
    source = "".join(module["source"] for module in case["modules"]) if "source" not in case else case["source"]
    found = _codes(source, "invalid-as-type-name", host=host, referenced_libraries=_REFERENCES[host])
    return sum("none of the libraries it references" in d.message for d in found)


def test_reftype_fixtures_cover_every_office_host() -> None:
    assert {case.get("host") or "excel" for case in _REFTYPE_CASES} == {"excel", "word", "powerpoint", "access"}


@pytest.mark.parametrize("case", _REFTYPE_CASES, ids=[case["id"] for case in _REFTYPE_CASES])
def test_reftype_fixture(case: dict[str, Any]) -> None:
    # A Word document's Normal template could hold the name, so Word stays silent.
    reported = case["expected"] == "rejected" and case.get("host") != "word" and case["id"] not in _KNOWN_SILENT
    assert _unknown_types(case) == (1 if reported else 0)


def test_reftype_silent_where_the_references_are_not_known() -> None:
    src = "Option Explicit\nSub Main()\n    Dim w As Worksheet\nEnd Sub\n"
    assert _codes(src, "invalid-as-type-name", host="access") == []


# dc4cab57: the Outlook object library as a project reference (outlookReference.test.ts,
# hostInterop.test.ts).

_CASES_BY_ID = {case["id"]: case for case in _ORACLE}
_HOST_PREFIXES = ("", "word_", "powerpoint_", "access_")


def _each_host(name: str, prefixes: tuple[str, ...] = _HOST_PREFIXES) -> list[str]:
    return [f"outlook_{prefix}{name}_compile" for prefix in prefixes]


def _error_codes(case: dict[str, Any], referenced_libraries: list[str] | None = None, project_name: str | None = None) -> list[str]:
    """Each error code, analyzed as the extension analyzes the case's host and
    references. The port reports a missing reference only where the reference list
    is known, so a case without references is given a known-empty list."""
    index = ProjectIndex(ProjectIndexOptions(conditional_compilation=ConditionalCompilationEnvironment(project_constants={})))
    for module in case["modules"]:
        index.set_module(ModuleInput(module["name"], _KIND.get(module["type"], ModuleSymbolKind.STANDARD), module["source"]))
    tokens = [token for ref in case.get("references", []) if (token := host_token_for_libid(ref["guid"])) is not None]
    out: list[str] = []
    for module in case["modules"]:
        opts = analyze_module_options_for(
            index, module["name"], _KIND.get(module["type"], ModuleSymbolKind.STANDARD),
            host=case.get("host") or "excel", referenced_hosts=tokens,
            referenced_libraries=referenced_libraries, project_name=project_name,
            conditional_compilation=ConditionalCompilationEnvironment(project_constants={}),
        )
        out.extend(d.code for d in analyze_module(module["source"], opts) if d.severity is DiagnosticSeverity.ERROR)
    return out


_OUTLOOK_ACCEPTED = [
    case_id
    for name in (
        "ref_dim_app", "ref_dim_mailitem", "ref_dim_bare_mailitem", "ref_new_app", "ref_constant",
        "ref_qualified_constant", "ref_enum_constant", "ref_member_ok", "ref_member_missing",
        "ref_item_member_ok", "ref_item_member_missing", "ref_bare_session", "ref_bare_createitem",
        "ref_bare_getnamespace", "ref_module_level", "ref_withevents",
    )
    for case_id in _each_host(name)
]


@pytest.mark.parametrize("case_id", _OUTLOOK_ACCEPTED)
def test_outlook_referenced_case_compiles_and_nothing_reports_it(case_id: str) -> None:
    case = _CASES_BY_ID[case_id]
    assert case["expected"] == "accepted"
    assert len(case["references"]) == 1
    assert _error_codes(case) == []


_OUTLOOK_REJECTED = [
    (case_id, code)
    for name, code in (
        ("noref_dim_app", "missing-library-reference"),
        ("noref_constant", "undeclared-variable"),
        ("ref_bare_unknown", "undeclared-variable"),
        ("ref_wrong_arguments", "argument-count"),
    )
    for case_id in _each_host(name)
]


@pytest.mark.parametrize("case_id, code", _OUTLOOK_REJECTED)
def test_outlook_case_is_refused_as(case_id: str, code: str) -> None:
    case = _CASES_BY_ID[case_id]
    assert case["expected"] == "rejected"
    assert _error_codes(case) == [code]


@pytest.mark.parametrize("case_id", _each_host("noref_dim_bare_mailitem"))
def test_outlook_bare_type_waits_for_the_reference_list(case_id: str) -> None:
    case = _CASES_BY_ID[case_id]
    assert case["expected"] == "rejected"
    assert _error_codes(case) == []
    assert _error_codes(case, ["stdole", "Office"]) == ["invalid-as-type-name"]


@pytest.mark.parametrize("case_id", _each_host("ref_dim_bare_mailitem"))
def test_outlook_bare_type_resolves_once_the_list_holds_outlook(case_id: str) -> None:
    assert _error_codes(_CASES_BY_ID[case_id], ["stdole", "Office", "Outlook"]) == []


@pytest.mark.parametrize("case_id", _each_host("xlref_bare_activeworkbook", _HOST_PREFIXES[1:]))
def test_referenced_excel_global_binds_bare(case_id: str) -> None:
    case = _CASES_BY_ID[case_id]
    assert case["expected"] == "accepted"
    assert _error_codes(case) == []


@pytest.mark.parametrize("case_id", _each_host("noxlref_bare_activeworkbook", _HOST_PREFIXES[1:]))
def test_unreferenced_excel_global_is_not_defined(case_id: str) -> None:
    case = _CASES_BY_ID[case_id]
    assert case["expected"] == "rejected"
    assert _error_codes(case) == ["undeclared-variable"]


def test_outlook_unreferenced_carries_the_fix_and_resolves_referenced() -> None:
    def found(hosts: list[str]) -> list[Any]:
        source = "Dim o As Outlook.Application\n"
        opts = AnalyzeModuleOptions(host=hosts[0], referenced_hosts=hosts[1:])
        return [d for d in analyze_module(source, opts) if d.code == "missing-library-reference"]

    [missing] = found(["word"])
    assert missing.data is not None and missing.data.add_library_reference is not None
    assert missing.data.add_library_reference.library == "outlook"
    assert found(["word", "outlook"]) == []


def test_outlook_library_and_globals_in_the_merged_model() -> None:
    assert host_token_for_libid("{00062FFF-0000-0000-C000-000000000046}#9.6#0#C:\\x.olb#Outlook") == "outlook"
    model = host_object_model_for_tokens(["excel", "outlook"])
    assert model is not None
    assert list(model.get("referencedGlobalTypes") or []) == ["Outlook.Application"]
    # The host's own Global wins a name both have; Outlook's Session binds after it.
    application = resolve_host_global_member("Application", model)
    assert application is not None and "Outlook" not in json.dumps(application)
    assert resolve_host_global_member("Session", model) is not None
    names = {member["name"] for member in get_host_global_members(model)}
    assert {"Session", "CreateItem", "GetNamespace", "ActiveWorkbook"} <= names
    excel_only = host_object_model_for_tokens(["excel", "word"])
    assert excel_only is not None and resolve_host_global_member("Session", excel_only) is None


# 728291c6: a type qualified by a library the project does not reference
# (unreferencedQualifiers1413.test.ts).

# The names the cases' references go by, as the dir stream records them (measured).
_LIBRARY_NAMES = {
    "{420B2830-E718-11CF-893D-00A0C9054228}": "Scripting",
    "{3F4DACA7-160D-11D2-A8E9-00104B365C9F}": "VBScript_RegExp_55",
    "{B691E011-1797-432E-907A-4D8C69339129}": "ADODB",
    "{F5078F18-C551-11D3-89B9-0000F81FE221}": "MSXML2",
    "{0002E157-0000-0000-C000-000000000046}": "VBIDE",
    "{662901FC-6951-4854-9EB2-D9A2570F2B2E}": "WinHttp",
    "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}": "Shell32",
    "{00062FFF-0000-0000-C000-000000000046}": "Outlook",
}
# What a new file of each host records before any reference is added, and its
# project's name (measured).
_DEFAULTS = {
    "excel": ["stdole", "Office"],
    "word": ["stdole", "Normal", "Office"],
    "powerpoint": ["stdole", "Office"],
    "access": ["stdole", "DAO"],
}
_PROJECT_NAMES = {"excel": "VBAProject", "word": "Project", "powerpoint": "VBAProject", "access": "Database1"}


def _codes_1413(case: dict[str, Any], known: bool = True) -> list[str]:
    """Each error code, with the project's references and name known, as the
    extension has them."""
    host = case.get("host") or "excel"
    if not known:
        return _error_codes(case)
    libraries = [*_DEFAULTS[host], *(_LIBRARY_NAMES[ref["guid"]] for ref in case.get("references", []))]
    return _error_codes(case, libraries, _PROJECT_NAMES[host])


def _each_host_1413(names: list[str]) -> list[str]:
    return [f"issue1413_{prefix}{name}_compile" for name in names for prefix in _HOST_PREFIXES]


_LIBRARIES_1413 = ["scripting", "regexp", "adodb", "msxml", "vbide", "winhttp", "shell", "outlook"]


@pytest.mark.parametrize(
    "case_id",
    _each_host_1413([
        *(f"noref_{name}" for name in _LIBRARIES_1413), "unknown_qualifier", "unknown_new",
        "unknown_module_level", "unknown_parameter", "unknown_implements", "unknown_declare",
    ]),
)
def test_unreferenced_qualifier_is_user_defined_type_not_defined(case_id: str) -> None:
    case = _CASES_BY_ID[case_id]
    assert case["expected"] == "rejected"
    assert _codes_1413(case) == ["missing-library-reference"]


@pytest.mark.parametrize(
    "case_id",
    _each_host_1413([
        *(f"ref_{name}" for name in _LIBRARIES_1413), "ref_lowercase_qualifier", "ref_declare_outlook",
        "default_vba", "default_stdole", "module_type", "other_module_class", "string_holds_qualifier",
    ]),
)
def test_referenced_qualifier_compiles_and_nothing_reports_it(case_id: str) -> None:
    case = _CASES_BY_ID[case_id]
    assert case["expected"] == "accepted"
    assert _codes_1413(case) == []


@pytest.mark.parametrize("case_id", _each_host_1413(["noref_adodb", "unknown_qualifier", "unknown_implements"]))
def test_unreferenced_qualifier_waits_for_the_references(case_id: str) -> None:
    assert _codes_1413(_CASES_BY_ID[case_id], known=False) == []


@pytest.mark.parametrize(
    "case_id, expected",
    [
        ("issue1413_default_office_compile", "accepted"),
        ("issue1413_word_default_office_compile", "accepted"),
        ("issue1413_powerpoint_default_office_compile", "accepted"),
        ("issue1413_access_default_office_compile", "rejected"),
        # The project's own name qualifies its types, and only it.
        ("issue1413_project_vbaproject_class_compile", "accepted"),
        ("issue1413_word_project_vbaproject_class_compile", "rejected"),
        ("issue1413_powerpoint_project_vbaproject_class_compile", "accepted"),
        ("issue1413_access_project_vbaproject_class_compile", "rejected"),
        ("issue1413_project_project_class_compile", "rejected"),
        ("issue1413_word_project_project_class_compile", "accepted"),
        ("issue1413_powerpoint_project_project_class_compile", "rejected"),
        ("issue1413_access_project_project_class_compile", "rejected"),
    ],
)
def test_default_references_and_project_name_qualify(case_id: str, expected: str) -> None:
    case = _CASES_BY_ID[case_id]
    assert case["expected"] == expected
    assert _codes_1413(case) == ([] if expected == "accepted" else ["missing-library-reference"])


def test_unreferenced_qualifier_message_span_and_once_per_library() -> None:
    source = "Option Explicit\r\nPrivate a As Nowhere.Thing\r\nPrivate b As Nowhere.Other\r\nPrivate c As Elsewhere.Thing\r\n"
    index = ProjectIndex(ProjectIndexOptions(conditional_compilation=ConditionalCompilationEnvironment(project_constants={})))
    index.set_module(ModuleInput("Module1", ModuleSymbolKind.STANDARD, source))
    opts = analyze_module_options_for(
        index, "Module1", ModuleSymbolKind.STANDARD,
        referenced_libraries=["stdole", "Office"], project_name="VBAProject",
    )
    found = [d for d in analyze_module(source, opts) if d.code == "missing-library-reference"]
    assert [source[d.span.start:d.span.end] for d in found] == ["Nowhere.Thing", "Elsewhere.Thing"]
    assert "'Nowhere' is no library this project references and none of its modules" in found[0].message
    assert "User-defined type not defined" in found[0].message
    assert all(d.data is None or d.data.add_library_reference is None for d in found)


def test_missing_scripting_reference_is_reported_once() -> None:
    assert _codes_1413(_CASES_BY_ID["issue1413_noref_scripting_compile"]) == ["missing-library-reference"]


def test_qualified_runtime_function_name_is_another_librarys_type() -> None:
    src = "Option Explicit\nSub Main()\n    Dim s As Shell32.Shell\n    Dim v As VBA.Shell\n    Dim w As Shell\nEnd Sub\n"
    found = _codes(src, "invalid-as-type-name")
    # VBA.Shell and the bare Shell are the runtime function; Shell32.Shell is not.
    assert [d.span.start for d in found] == [src.index("VBA.Shell") + 4, src.rindex("Shell")]


# f6575c3b: a renamed workbook or document module keeps its host type
# (ownObjectMembers228.test.ts, vbaProjectIndexService.test.ts).


def _project_module_findings(
    source: str,
    modules: list[tuple[str, ModuleSymbolKind, str, str | None]],
    current: str,
    code: str,
    host: str = "excel",
    document_type: str | None = None,
) -> list[str]:
    """The text of each finding of `code` in `current`, analyzed with the project's
    modules (name, kind, source, document type) indexed."""
    index = ProjectIndex()
    for name, kind, text, doc_type in modules:
        index.set_module(ModuleInput(name, kind, text, document_type=doc_type))
    kind = next(kind for name, kind, _, _ in modules if name == current)
    opts = analyze_module_options_for(index, current, kind, host=host)
    opts.document_type = document_type
    return [source[d.span.start:d.span.end] for d in analyze_module(source, opts) if d.code == code]


_ME = "Option Explicit\nPublic Sub T()\n    Me.UsedRange.Select\n    Debug.Print Me.FullName\nEnd Sub\n"


@pytest.mark.parametrize(
    "module_name, document_type, host",
    [
        ("ThisWorkbook", None, "excel"),
        ("ThisWorkbook", "workbook", "excel"),
        ("BookCode", "workbook", "excel"),
        ("ThisDocument", None, "word"),
        ("DocCode", "document", "word"),
    ],
)
def test_me_in_a_document_module_is_its_host_object(module_name: str, document_type: str | None, host: str) -> None:
    modules = [(module_name, ModuleSymbolKind.DOCUMENT, _ME, document_type)]
    found = _project_module_findings(_ME, modules, module_name, "member-not-found", host, document_type)
    assert found == ["UsedRange"]


@pytest.mark.parametrize(
    "module_name, document_type",
    [("BookCode", None), ("Sheet1", "worksheet"), ("Chart1", "chart")],
)
def test_me_in_a_sheet_or_unnamed_document_keeps_to_its_own_members(module_name: str, document_type: str | None) -> None:
    # With no document type a name proves nothing, and a sheet's Me waits for its
    # designer class: only with it are the sheet's ActiveX controls known.
    modules = [(module_name, ModuleSymbolKind.DOCUMENT, _ME, document_type)]
    assert _project_module_findings(_ME, modules, module_name, "member-not-found", "excel", document_type) == []


def test_document_type_reaches_the_project_type_across_edits() -> None:
    index = ProjectIndex()
    index.set_module(ModuleInput("BookCode", ModuleSymbolKind.DOCUMENT, "Public Sub Alpha()\nEnd Sub\n", document_type="workbook"))
    index.set_module(ModuleInput("Sheet1", ModuleSymbolKind.DOCUMENT, ""))

    def document_types() -> list[tuple[str, str | None]]:
        return [
            (surface.name, surface.document_type)
            for surface in index.project_member_surfaces("Sheet1")
            if surface.kind == "document"
        ]

    assert document_types() == [("BookCode", "workbook"), ("Sheet1", None)]
    index.set_module(ModuleInput("Sheet1", ModuleSymbolKind.DOCUMENT, "Public Sub Delta()\nEnd Sub\n"))
    assert document_types() == [("BookCode", "workbook"), ("Sheet1", None)]
    index.set_module(ModuleInput("BookCode", ModuleSymbolKind.DOCUMENT, "Public Sub Gamma()\nEnd Sub\n"))
    assert document_types() == [("BookCode", None), ("Sheet1", None)]
    index.remove_module("BookCode")
    assert document_types() == [("Sheet1", None)]
