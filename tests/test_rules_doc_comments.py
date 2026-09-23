"""Doc-comment rule family (docComments.ts parity): the XML tags of a ''' block
checked against the declaration they document, with the fixes a code action offers.

Every finding, span and fix below is what XLIDE 10.5.0 produces on the same
source; the port was diffed against it case by case.
"""

from __future__ import annotations

import pytest

from pyvbaanalysis import analyze_module
from pyvbaanalysis.diagnostics import VbaDiagnostic

Fix = tuple[str, bool, list[tuple[int, int, str]]]


def _doc_findings(source: str) -> list[VbaDiagnostic]:
    return [d for d in analyze_module(source) if d.code.startswith("doc-")]


def _located(source: str) -> list[tuple[str, int, int]]:
    return [(d.code, d.span.start, d.span.end) for d in _doc_findings(source)]


def _fixes(diagnostic: VbaDiagnostic) -> list[Fix]:
    data = diagnostic.data
    fixes = data.doc_comment_fixes if data is not None and data.doc_comment_fixes else ()
    return [
        (fix.title, fix.is_preferred, [(e.span.start, e.span.end, e.new_text) for e in fix.edits])
        for fix in fixes
    ]


@pytest.mark.parametrize(
    "source",
    [
        pytest.param(
            "''' <summary>Adds.</summary>\n''' <param name=\"a\">First.</param>\n''' <returns>Sum.</returns>\n"
            "Public Function Add(a As Long) As Long\n    Add = a\nEnd Function\n",
            id="complete",
        ),
        pytest.param(
            "''' Adds two numbers.\nPublic Function Add(a As Long) As Long\n    Add = a\nEnd Function\n",
            id="plain prose is not XML",
        ),
        pytest.param(
            "''''''''''''''''''''\nPublic Sub S(a As Long)\nEnd Sub\n", id="a row of apostrophes"
        ),
        pytest.param(
            "''' <summary>S.</summary>\n''' <param name=\"a\" type=\"Long\"></param>\nPublic Sub S(a As Long)\nEnd Sub\n",
            id="an empty param with a type attribute",
        ),
        pytest.param(
            "''' <summary>The size.</summary>\nPublic Property Let Size(ByVal v As Long)\nEnd Property\n",
            id="a Property Let value parameter",
        ),
        pytest.param(
            "''' <summary>The size.</summary>\nPublic Property Get Size() As Long\n    Size = 1\nEnd Property\n",
            id="returns is optional on Property Get",
        ),
    ],
)
def test_blocks_in_line_with_their_declaration_are_silent(source: str) -> None:
    assert _doc_findings(source) == []


def test_missing_params_and_returns_each_offer_an_insertion() -> None:
    source = (
        "''' <summary>Adds.</summary>\n"
        "Public Function Add(a As Long, b As Long) As Long\n    Add = a + b\nEnd Function\n"
    )
    found = _doc_findings(source)
    assert [(d.code, d.span.start, d.span.end) for d in found] == [
        ("doc-param-missing", 49, 50),
        ("doc-param-missing", 60, 61),
        ("doc-returns-missing", 45, 48),
    ]
    both = (
        "Add the 2 missing <param> tags",
        False,
        [(29, 29, "''' <param name=\"a\"></param>\n''' <param name=\"b\"></param>\n")],
    )
    assert _fixes(found[0]) == [
        ("Add a <param> for 'a'", True, [(29, 29, "''' <param name=\"a\"></param>\n")]),
        both,
    ]
    assert _fixes(found[1]) == [
        ("Add a <param> for 'b'", True, [(29, 29, "''' <param name=\"b\"></param>\n")]),
        both,
    ]
    assert _fixes(found[2]) == [("Add a <returns>", True, [(29, 29, "''' <returns></returns>\n")])]


def test_an_insertion_keeps_the_blocks_line_endings() -> None:
    source = (
        "''' <summary>Adds.</summary>\r\n"
        "Public Function Add(a As Long, b As Long) As Long\r\n    Add = a + b\r\nEnd Function\r\n"
    )
    first = _doc_findings(source)[0]
    assert _fixes(first)[0] == (
        "Add a <param> for 'a'",
        True,
        [(30, 30, "''' <param name=\"a\"></param>\r\n")],
    )


def test_a_param_naming_no_parameter_offers_a_rename_and_a_removal() -> None:
    source = "''' <summary>S.</summary>\n''' <param name=\"zz\">?</param>\nPublic Sub S(a As Long)\nEnd Sub\n"
    unknown, missing = _doc_findings(source)
    assert (unknown.code, unknown.span.start, unknown.span.end) == ("doc-param-unknown", 43, 45)
    assert _fixes(unknown) == [
        ("Rename the <param> to 'a'", True, [(43, 45, "a")]),
        ("Remove the <param> for 'zz'", False, [(26, 57, "")]),
    ]
    assert missing.code == "doc-param-missing"


def test_a_param_with_no_name_offers_to_name_it() -> None:
    source = "''' <summary>S.</summary>\n''' <param>?</param>\nPublic Sub S(a As Long)\nEnd Sub\n"
    unknown = _doc_findings(source)[0]
    assert _fixes(unknown) == [
        ("Name the <param> 'a'", True, [(36, 36, ' name="a"')]),
        ("Remove the <param>", False, [(26, 47, "")]),
    ]


def test_a_param_name_is_read_with_its_entities_decoded() -> None:
    source = "''' <summary>S.</summary>\n''' <param name=\"a&amp;b\">x</param>\nPublic Sub S(a As Long)\nEnd Sub\n"
    unknown = _doc_findings(source)[0]
    assert _fixes(unknown)[1] == ("Remove the <param> for 'a&b'", False, [(26, 62, "")])


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        pytest.param(
            "''' <summary>S.</summary>\n''' <param name=\"a\"></param>\nPublic Sub S(a As Long)\nEnd Sub\n",
            [("doc-param-missing", 30, 46)],
            id="an empty param documents nothing",
        ),
        pytest.param(
            "''' <summary>S.</summary>\n''' <param name=\"a\"/>\nPublic Sub S(a As Long)\nEnd Sub\n",
            [("doc-param-missing", 30, 47)],
            id="a self-closing param documents nothing",
        ),
        pytest.param(
            "''' <summary>F.</summary>\n''' <returns></returns>\nPublic Function F() As Long\n    F = 1\nEnd Function\n",
            [("doc-returns-missing", 30, 39)],
            id="an empty returns",
        ),
        pytest.param(
            "''' <summary>S.</summary>\n''' <returns>Nothing.</returns>\nPublic Sub S()\nEnd Sub\n",
            [("doc-returns-unexpected", 30, 39)],
            id="returns on a Sub",
        ),
        pytest.param(
            "''' <summary>F.\nPublic Sub S()\nEnd Sub\n",
            [("doc-tag-unclosed", 4, 13)],
            id="unclosed",
        ),
        pytest.param(
            "''' <summary>One.</summary>\n''' <summary>Two.</summary>\nPublic Sub S()\nEnd Sub\n",
            [("doc-tag-duplicate", 32, 41)],
            id="a second summary",
        ),
        pytest.param(
            "''' <summary>Sleeps.</summary>\n"
            'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long\n',
            [("doc-returns-missing", 64, 76)],
            id="a Declare Function",
        ),
        pytest.param(
            "''' <summary>Changed.</summary>\nPublic Event Changed(ByVal value As Long)\n",
            [("doc-param-missing", 59, 64)],
            id="an Event's parameters",
        ),
        pytest.param(
            "''' <summary>S.</summary>\n'@xlide-analysis-disable-next-line some-rule\nPublic Sub S(a As Long)\nEnd Sub\n",
            [("doc-param-missing", 84, 85)],
            id="a directive between the block and the member",
        ),
    ],
)
def test_block_findings(source: str, expected: list[tuple[str, int, int]]) -> None:
    assert _located(source) == expected


def test_removing_a_misplaced_or_repeated_tag() -> None:
    on_a_sub = (
        "''' <summary>S.</summary>\n''' <returns>Nothing.</returns>\nPublic Sub S()\nEnd Sub\n"
    )
    assert _fixes(_doc_findings(on_a_sub)[0]) == [("Remove the <returns>", False, [(26, 58, "")])]
    repeated = "''' <summary>One.</summary>\n''' <summary>Two.</summary>\nPublic Sub S()\nEnd Sub\n"
    assert _fixes(_doc_findings(repeated)[0]) == [
        ("Remove the repeated <summary>", False, [(28, 56, "")])
    ]
