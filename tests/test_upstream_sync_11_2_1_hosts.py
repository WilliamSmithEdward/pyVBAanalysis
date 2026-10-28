"""Host, library and document-module contracts from the XLIDE 11.2.1 sync.

Each section names the upstream commit it ports; the cases are upstream's own
tests for that commit.
"""

from __future__ import annotations

import json
from typing import Any

import pytest

from pyvbaanalysis import analyze_module
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.evidence import DATA_DIR
from pyvbaanalysis.symbols import ModuleSymbolKind

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
