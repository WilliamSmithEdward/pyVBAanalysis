"""Referenced Office libraries (hostLibraries.ts and hostRegistry.ts parity): which
host answers for each reference a project declares, the one model that answers for
all of them, and the reader that carries a container's reference list through.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyvbaanalysis.host.host_libraries import (
    host_token_for_libid,
    host_tokens_for_project,
    library_guid_of,
    referenced_host_tokens,
)
from pyvbaanalysis.host.host_registry import (
    host_object_model_for_token,
    host_object_model_for_tokens,
)
from pyvbaanalysis.reader import analyze_office_file, read_office_project
from pyvbaanalysis.reader import workbook as workbook_mod
from pyvbaanalysis.reader.vbe_module import loaded_module_from_text

# Libids as a project stores them. The GUID is the identity; the path and the
# description are hints the host resolves through the registry.
_EXCEL = (
    r"*\G{00020813-0000-0000-C000-000000000046}#1.9#0#"
    r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE#Microsoft Excel 16.0 Object Library"
)
_WORD = (
    r"*\G{00020905-0000-0000-C000-000000000046}#8.7#0#"
    r"C:\Program Files\Microsoft Office\root\Office16\MSWORD.OLB#Microsoft Word 16.0 Object Library"
)
_STDOLE = r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation"
_OFFICE = (
    r"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#"
    r"C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library"
)
# A reference to another VBA project names a project, not a registered library.
_PROJECT = r"*\CNormal"


# -- libids ----------------------------------------------------------------


def test_the_guid_is_read_out_of_a_libid() -> None:
    assert library_guid_of(_WORD) == "{00020905-0000-0000-C000-000000000046}"
    assert library_guid_of(_WORD.lower()) == "{00020905-0000-0000-C000-000000000046}"
    assert library_guid_of(_PROJECT) is None


@pytest.mark.parametrize(
    ("libid", "host"),
    [
        (_EXCEL, "excel"),
        (_WORD, "word"),
        (
            r"*\G{91493440-5A91-11CF-8700-00AA0060263B}#2.c#0#C:\Office16\MSPPT.OLB#Microsoft PowerPoint 16.0 Object Library",
            "powerpoint",
        ),
        (
            r"*\G{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}#9.0#0#C:\Office16\MSACC.OLB#Microsoft Access 16.0 Object Library",
            "access",
        ),
        # The same library installed somewhere else is the same library.
        (r"*\G{00020905-0000-0000-C000-000000000046}#8.7#0#D:\Office\MSWORD.OLB#Word", "word"),
        (_STDOLE, None),
        (_OFFICE, None),
        (_PROJECT, None),
    ],
)
def test_each_office_library_maps_to_the_host_that_answers_for_it(
    libid: str, host: str | None
) -> None:
    assert host_token_for_libid(libid) == host


def test_the_projects_own_host_leads_then_declaration_order() -> None:
    assert host_tokens_for_project("word", [_STDOLE, _EXCEL, _OFFICE]) == ["word", "excel"]
    # The host's own library in the list neither repeats nor moves it.
    assert host_tokens_for_project("excel", [_WORD, _EXCEL]) == ["excel", "word"]
    assert host_tokens_for_project(None, [_WORD]) == ["word"]


def test_referenced_hosts_leave_out_the_projects_own() -> None:
    assert referenced_host_tokens("excel", [_EXCEL, _WORD]) == ["word"]
    assert referenced_host_tokens("word", [_STDOLE, _PROJECT, _OFFICE]) == []


# -- one model for a project and its references ----------------------------


def test_a_single_known_library_is_the_plain_model() -> None:
    word = host_object_model_for_token("word")
    assert host_object_model_for_tokens(["word"]) is word
    # Outlook has no model, so it adds nothing.
    assert host_object_model_for_tokens(["word", "outlook"]) is word
    # Excel is the default model, which rides as no model at all.
    assert host_object_model_for_tokens(["excel"]) is None


def test_the_first_library_wins_a_shared_name() -> None:
    """VBA resolves an ambiguous name by the reference list's order, with the
    project's own host at the top."""
    word_first = host_object_model_for_tokens(["word", "excel"])
    excel_first = host_object_model_for_tokens(["excel", "word"])
    assert word_first is not None and excel_first is not None
    assert word_first["globals"]["Selection"] == "Word.Selection"
    assert excel_first["globals"]["Selection"] == "Excel.Range"
    assert word_first.get("hostName") == "Word"


def test_every_library_contributes_its_types() -> None:
    merged = host_object_model_for_tokens(["word", "excel"])
    assert merged is not None
    assert "Word.Document" in merged["types"] and "Excel.Workbook" in merged["types"]


def test_a_referenced_librarys_enums_carry_its_name() -> None:
    word_first = host_object_model_for_tokens(["word", "excel"])
    excel_first = host_object_model_for_tokens(["excel", "word"])
    assert word_first is not None and excel_first is not None
    assert word_first["enums"]["XlAboveBelow"].get("library") == "Excel"
    assert excel_first["enums"]["WdAlertLevel"].get("library") == "Word"
    # The host's own enums need no label.
    assert "library" not in word_first["enums"]["WdAlertLevel"]


def test_a_merged_model_is_built_once() -> None:
    assert host_object_model_for_tokens(["word", "excel"]) is host_object_model_for_tokens(
        ["word", "excel"]
    )


# -- the reader carries the reference list ---------------------------------

_USES_WORD = "Option Explicit\nPublic Sub S()\n    Dim doc As Word.Document\n    Set doc = Nothing\nEnd Sub\n"


def _analyze_with_libids(monkeypatch: pytest.MonkeyPatch, libids: list[str] | None) -> list[str]:
    module = loaded_module_from_text(_USES_WORD, name="Module1", pyopenvba_standard=True)
    monkeypatch.setattr(workbook_mod, "_read_office_project", lambda _path: ([module], libids))
    return [d.code for d in analyze_office_file("Book.xlsm")["Module1"]]


def test_a_workbook_that_references_word_can_name_word_types(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    assert "missing-library-reference" not in _analyze_with_libids(
        monkeypatch, [_STDOLE, _EXCEL, _WORD]
    )


def test_a_workbook_without_the_reference_reports_it(monkeypatch: pytest.MonkeyPatch) -> None:
    assert "missing-library-reference" in _analyze_with_libids(
        monkeypatch, [_STDOLE, _EXCEL, _OFFICE]
    )


def test_an_unreadable_reference_list_proves_nothing(monkeypatch: pytest.MonkeyPatch) -> None:
    assert "missing-library-reference" not in _analyze_with_libids(monkeypatch, None)


def test_an_authored_container_reads_a_known_reference_list(tmp_path: Path) -> None:
    pyopenvba = pytest.importorskip("pyopenvba")
    path = tmp_path / "Fixture.docm"
    with pyopenvba.WordFile.create_new(path) as document:
        document.set_module("Module1", "Option Explicit\n")
        document.save()
    project = read_office_project(path)
    assert project.host == "word"
    # The template references stdole, Normal and Office: no other application.
    assert project.referenced_hosts == []
    assert {module.name for module in project.modules} >= {"ThisDocument", "Module1"}
