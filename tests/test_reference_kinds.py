"""classify_reference_kinds (referenceKinds.ts parity): whether each mention of a
name reads it, writes it, or modifies it in place.

Every expectation is XLIDE 10.6.0's classification of the same source.
"""

from __future__ import annotations

import pytest

from pyvbaanalysis.lexer.token_kinds import TokenKind
from pyvbaanalysis.lexer.tokenize import tokenize
from pyvbaanalysis.references import classify_reference_kinds


def _classified(source: str) -> list[tuple[str, str]]:
    names = [
        t
        for t in tokenize(source)
        if t.kind in (TokenKind.IDENTIFIER, TokenKind.BRACKETED_IDENTIFIER)
    ]
    kinds = classify_reference_kinds(source, [t.start for t in names])
    return [(t.raw_text, kinds[t.start]) for t in names]


@pytest.mark.parametrize(
    ("body", "expected"),
    [
        pytest.param("x = x + 1", [("x", "write"), ("x", "read")], id="assignment"),
        pytest.param(
            "Set o = New Collection\n    Let n = n * 2",
            [("o", "write"), ("Collection", "read"), ("n", "write"), ("n", "read")],
            id="Set and Let",
        ),
        pytest.param(
            "a(i) = 1\n    a.b.c = 2\n    a(i).b = 3",
            [
                ("a", "write"),
                ("i", "read"),
                ("a", "read"),
                ("b", "read"),
                ("c", "write"),
                ("a", "read"),
                ("i", "read"),
                ("b", "write"),
            ],
            id="element and member targets",
        ),
        pytest.param(
            "With obj\n        .y = 1\n        .z(2) = 3\n    End With",
            [("obj", "read"), ("y", "write"), ("z", "write")],
            id="With members",
        ),
        pytest.param(
            "For i = 1 To n\n    Next i\n    For Each item In items\n    Next item",
            [
                ("i", "write"),
                ("n", "read"),
                ("i", "read"),
                ("item", "write"),
                ("items", "read"),
                ("item", "read"),
            ],
            id="loop variables",
        ),
        pytest.param(
            "ReDim arr(n)\n    ReDim Preserve arr(n + 1), other(2) As Long",
            [
                ("arr", "write"),
                ("n", "read"),
                ("arr", "readwrite"),
                ("n", "read"),
                ("other", "readwrite"),
            ],
            id="ReDim Preserve keeps the contents",
        ),
        pytest.param("Erase arr, other", [("arr", "write"), ("other", "write")], id="Erase"),
        pytest.param(
            "Mid(s, 1, 2) = t",
            [("Mid", "read"), ("s", "readwrite"), ("t", "read")],
            id="the Mid statement",
        ),
        # The lexer gives a type character a token of its own, so `Mid$(` is
        # `Mid`, `$`, `(`, and the statement read its target as only read (XLIDE
        # issue #80).
        pytest.param(
            'Mid$(s, 2, 3) = "ab"', [("Mid", "read"), ("s", "readwrite")], id="the Mid$ statement"
        ),
        pytest.param(
            'MidB(s, 2, 3) = "ab"', [("MidB", "read"), ("s", "readwrite")], id="the MidB statement"
        ),
        pytest.param(
            'MidB$(s, 2, 3) = "ab"', [("MidB", "read"), ("s", "readwrite")], id="the MidB$ statement"
        ),
        pytest.param(
            "x = Mid$(s, 2, 3)",
            [("x", "write"), ("Mid", "read"), ("s", "read")],
            id="the Mid$ function",
        ),
        pytest.param(
            "If x = 1 Then y = 2 Else z = 3\n    If a Then b = c",
            [
                ("x", "read"),
                ("y", "write"),
                ("z", "write"),
                ("a", "read"),
                ("b", "write"),
                ("c", "read"),
            ],
            id="single-line If",
        ),
        pytest.param(
            "x = 1: y = x: z = y ' note",
            [("x", "write"), ("y", "write"), ("x", "read"), ("z", "write"), ("y", "read")],
            id="colon-separated statements",
        ),
        pytest.param(
            "If a(b = c) Then d = e\n    q = (r = s)",
            [
                ("a", "read"),
                ("b", "read"),
                ("c", "read"),
                ("d", "write"),
                ("e", "read"),
                ("q", "write"),
                ("r", "read"),
                ("s", "read"),
            ],
            id="a comparison is not an assignment",
        ),
        pytest.param(
            "Me.Value = 3\n    [bracketed name] = 4",
            [("Value", "write"), ("[bracketed name]", "write")],
            id="Me member and bracketed name",
        ),
    ],
)
def test_statement_references(body: str, expected: list[tuple[str, str]]) -> None:
    # The procedure's own name is a declaration, so it classifies as a write.
    assert _classified(f"Sub S()\n    {body}\nEnd Sub\n") == [("S", "write"), *expected]


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        pytest.param(
            "Sub S()\n    Dim x As Long, y(10) As String, z\n    Static k As Integer\nEnd Sub\n",
            [("S", "write"), ("x", "write"), ("y", "write"), ("z", "write"), ("k", "write")],
            id="variable declarations",
        ),
        pytest.param(
            "Const A = 1, B = A + 1\nPrivate Const C As Long = 5\n",
            [("A", "write"), ("B", "write"), ("A", "read"), ("C", "write")],
            id="constants read what they are built from",
        ),
        pytest.param(
            "Public Function F(ByVal p As Long, Optional q As Long = 3, ParamArray r() As Variant) As Long\nEnd Function\n",
            [("F", "write"), ("p", "write"), ("q", "write"), ("r", "write")],
            id="parameters",
        ),
        pytest.param(
            "Private WithEvents app As Excel.Application\nPublic gx As Long\n",
            [("app", "write"), ("Excel", "read"), ("Application", "read"), ("gx", "write")],
            id="a type name is read",
        ),
    ],
)
def test_declarations(source: str, expected: list[tuple[str, str]]) -> None:
    assert _classified(source) == expected


def test_only_the_requested_offsets_are_classified() -> None:
    source = "Sub S()\n    x = x + 1\nEnd Sub\n"
    second_x = source.rindex("x")
    assert classify_reference_kinds(source, [second_x]) == {second_x: "read"}
