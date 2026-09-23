"""missingLibraryReference (missingReference.ts parity): early binding to another
Office application's library the project does not reference.

Every expectation with a known reference list is what XLIDE 10.5.0 reports on the
same source. The unknown-list gate is the port's own: upstream reads an absent list
as "nothing referenced", while here only a list read from the container is taken as
proof that a library is missing.
"""

from __future__ import annotations

import pytest
from oracle_support import assert_oracle_behavior

from pyvbaanalysis import analyze_module, analyze_project
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind

_CODE = "missing-library-reference"


def _message(library: str) -> str:
    return (
        f"'{library}' is not referenced by this project, so {library}.* cannot be resolved. "
        f"Add a reference to the {library} object library, or use late binding: "
        f'Dim x As Object: Set x = CreateObject("{library}.Application").'
    )


def _missing(
    source: str, *, host: str | None = None, referenced: list[str] | None = None
) -> list[str]:
    opts = AnalyzeModuleOptions(host=host, referenced_hosts=referenced)
    return [d.message for d in analyze_module(source, opts) if d.code == _CODE]


@pytest.mark.parametrize(
    ("source", "library"),
    [
        pytest.param(
            "Public Sub S()\n    Dim doc As Word.Document\nEnd Sub\n", "Word", id="As clause"
        ),
        pytest.param(
            "Public Sub S()\n    Dim p As Object\n    Set p = New PowerPoint.Application\nEnd Sub\n",
            "PowerPoint",
            id="New",
        ),
        pytest.param(
            "Public Sub S()\n    Dim v As Long\n    v = Word.wdMainTextStory\nEnd Sub\n",
            "Word",
            id="qualified constant",
        ),
        pytest.param(
            "Private mApp As Access.Application\n", "Access", id="module-level declaration"
        ),
    ],
)
def test_naming_an_unreferenced_library_reports(source: str, library: str) -> None:
    assert _missing(source, referenced=[]) == [_message(library)]


def test_each_library_is_reported_once_per_module() -> None:
    source = (
        "Public Sub S()\n    Dim a As Word.Document\n    Dim b As Access.Application\n"
        "    Dim c As Word.Range\nEnd Sub\n"
    )
    assert _missing(source, referenced=[]) == [_message("Word"), _message("Access")]


def test_the_finding_carries_the_library_to_add() -> None:
    source = "Public Sub S()\n    Dim doc As Word.Document\nEnd Sub\n"
    found = next(
        d
        for d in analyze_module(source, AnalyzeModuleOptions(referenced_hosts=[]))
        if d.code == _CODE
    )
    assert source[found.span.start : found.span.end] == "Word.Document"
    assert found.data is not None and found.data.add_library_reference is not None
    assert found.data.add_library_reference.library == "word"


def test_the_projects_own_host_needs_no_reference() -> None:
    source = "Public Sub S()\n    Dim xl As Excel.Application\nEnd Sub\n"
    assert _missing(source, referenced=[]) == []
    assert _missing(source, host="word", referenced=[]) == [_message("Excel")]
    assert _missing(source, host="word", referenced=["excel"]) == []


@pytest.mark.parametrize(
    "source",
    [
        pytest.param(
            'Public Sub S()\n    Dim xl As Object\n    Set xl = CreateObject("Word.Application")\nEnd Sub\n',
            id="late binding names nothing",
        ),
        pytest.param(
            "Public Sub S()\n    Dim f As Office.FileDialog\nEnd Sub\n",
            id="the shared Office library",
        ),
        pytest.param(
            "Public Sub S()\n    Dim o As Outlook.Application\nEnd Sub\n",
            id="a library the analyzer cannot add",
        ),
    ],
)
def test_silent_where_there_is_no_reference_to_add(source: str) -> None:
    assert _missing(source, referenced=[]) == []


def test_a_referenced_library_is_silent() -> None:
    assert (
        _missing("Public Sub S()\n    Dim doc As Word.Document\nEnd Sub\n", referenced=["word"])
        == []
    )


def test_an_unknown_reference_list_proves_nothing() -> None:
    """A loose .bas file carries no reference list, so a missing reference cannot be
    proven there; the same source read from a container that lists none reports."""
    source = "Public Sub S()\n    Dim doc As Word.Document\nEnd Sub\n"
    assert _missing(source) == []
    assert _missing(source, referenced=[]) == [_message("Word")]


def test_a_referenced_library_answers_for_its_names() -> None:
    """With Word referenced, its constants resolve bare, as VBA binds them."""
    modules = [
        ModuleInput(
            "Mod1",
            ModuleSymbolKind.STANDARD,
            "Option Explicit\nPublic Sub S()\n    Debug.Print wdStory\nEnd Sub\n",
        )
    ]
    assert analyze_project(modules, referenced_hosts=["word"]) == {"Mod1": []}
    assert [d.code for d in analyze_project(modules, referenced_hosts=[])["Mod1"]] == [
        "undeclared-variable"
    ]


def test_a_referenced_library_is_checked_against_its_own_model() -> None:
    source = (
        "Option Explicit\nPublic Sub S(ByVal ws As Excel.Worksheet)\n    ws.NoSuchMember\nEnd Sub\n"
    )
    modules = [ModuleInput("Mod1", ModuleSymbolKind.STANDARD, source)]
    found = analyze_project(modules, host="word", referenced_hosts=["excel"])["Mod1"]
    assert [(d.code, d.message) for d in found] == [
        ("member-not-found", "Method or data member not found: 'Excel.Worksheet.NoSuchMember'.")
    ]


def test_oracle_asserted_cases() -> None:
    assert assert_oracle_behavior(_CODE) > 0
