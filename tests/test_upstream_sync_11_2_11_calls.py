"""Rule contracts from the XLIDE 11.2.1 -> 11.2.11 sync (calls, events,
conditional compilation and the rest).

Each section names the upstream commit whose own tests it follows.
"""

import pytest

from pyvbaanalysis import analyze_module
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.diagnostics.rule_metadata import diagnostic_metadata_for_code


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
