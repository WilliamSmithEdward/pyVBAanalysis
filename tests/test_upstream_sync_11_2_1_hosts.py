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


def _error_codes(case: dict[str, Any], referenced_libraries: list[str] | None = None) -> list[str]:
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
            referenced_libraries=referenced_libraries,
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
