"""Compiler-directive contracts from the pinned XLIDE 11.2.1 release.

Ported from upstream's remDirective1321, colonDirective1323,
misplacedDirective1331, stringDirectiveCondition,
conditionalConstUnknownValue and nullNothingDirectiveCondition tests. The
parts that check analyzeVbaStructure, XLIDE's block-balance pass the port
leaves out, are not carried over.
"""

from __future__ import annotations

import re

import pytest

from pyvbaanalysis import analyze_module
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.diagnostics.model import VbaDiagnostic as Diagnostic


def _analyze(source: str) -> list[Diagnostic]:
    return list(
        analyze_module(
            source,
            AnalyzeModuleOptions(
                known_identifiers=set(),
                conditional_compilation=ConditionalCompilationEnvironment(project_constants={}),
            ),
        )
    )


def _by_code(source: str, code: str) -> list[Diagnostic]:
    return [d for d in _analyze(source) if d.code == code]


def _malformed(source: str) -> list[Diagnostic]:
    return _by_code(source, "malformed-statement")


def _undeclared(source: str) -> list[str]:
    names = []
    for d in _by_code(source, "undeclared-variable"):
        found = re.search(r"'(\w+)'", d.message)
        names.append(found.group(1) if found else "")
    return names


def _span_text(source: str, d: Diagnostic) -> str:
    return source[d.span.start : d.span.end]


def _fn(*lines: str) -> str:
    return "\n".join(
        ["Option Explicit", "Function Main() As Variant", *lines, "    Main = 1", "End Function", ""]
    )


def _mod(line: str) -> str:
    return "\n".join([line, "Function Main() As Variant", "    Main = 1", "End Function", ""])


# -- Rem on a directive line (XLIDE issue #1321, commit 316756ad) -------------


@pytest.mark.parametrize(
    "line",
    [
        "#If Win64 Then Rem note",
        "#If True Rem note Then",
        "#If 1Rem note Then",
        "#If True _\n    Rem note Then",
        "#If True Then REM note",
    ],
)
def test_rem_if_line_opens_no_block(line: str) -> None:
    src = _fn(line, "    a = 1", "#Else", "    b = 2", "#End If")
    found = _malformed(src)
    assert len(found) == 1
    assert "Rem does not start a comment" in found[0].message
    assert "opens no #If block" in found[0].message
    assert "with ' instead" in found[0].message


@pytest.mark.parametrize(
    "line", ["#Else Rem note", "#ElseIf False Rem note Then", "#ElseIf False Then Rem note"]
)
def test_rem_else_in_compiled_arm_runs_on(line: str) -> None:
    src = _fn("#If True Then", "    a = 1", line, "    b = 2", "#End If")
    found = _malformed(src)
    assert len(found) == 1
    assert _span_text(src, found[0]) == re.sub(r"(Rem).*", r"\1", line)
    assert "starts no new arm" in found[0].message
    assert _undeclared(src) == ["a", "b"]


@pytest.mark.parametrize("line", ["#End If Rem note", "#EndIf Rem note"])
def test_rem_end_if_closes_nothing(line: str) -> None:
    src = _fn("#If True Then", "    a = 1", line)
    found = _malformed(src)
    assert len(found) == 1
    assert "closes no #If block" in found[0].message


@pytest.mark.parametrize("line", ["#Else Rem note", "#ElseIf True Rem note Then"])
def test_rem_line_in_skipped_arm_is_quiet(line: str) -> None:
    src = _fn("#If False Then", "    a = 1", line, "    b = 2", "#End If")
    assert _malformed(src) == []
    assert _undeclared(src) == []


def test_rem_line_at_module_level() -> None:
    src = "\n".join(
        ["#If Win64 Then Rem note", "Private x As Long", "#End If", "Function Main() As Variant",
         "    Main = 1", "End Function", ""]
    )  # fmt: skip
    found = _malformed(src)
    assert len(found) == 1
    assert "opens no #If block" in found[0].message


@pytest.mark.parametrize(
    "src",
    [
        _fn("#If Remark Then", "#End If"),
        _fn('#If "Rem" = "Rem" Then', "#End If"),
        _fn("#If True Then ' Rem note", "#End If"),
        _fn("#If True Then", "    Rem note", "#End If"),
        _fn("#If True Then", "#Else", "#End If ' Rem"),
        _fn("#If x.Rem Then", "#End If"),
        _fn("#If x!Rem Then", "#End If"),
    ],
)
def test_rem_lookalikes_are_quiet(src: str) -> None:
    assert _malformed(src) == []


@pytest.mark.parametrize(
    "line, error",
    [
        ("#Const A = True Rem note", "Expected: end of statement"),
        ("#Const A = True _\n    Rem note", "Expected: end of statement"),
        ("#Const A = Rem note", "Expected: expression"),
        ("#Const Rem = 1", "Expected: identifier"),
    ],
)
def test_rem_on_const_line(line: str, error: str) -> None:
    found = _malformed(_mod(line))
    assert len(found) == 1
    assert "#Const line" in found[0].message
    assert error in found[0].message


@pytest.mark.parametrize(
    "line",
    [
        "#If False Then\n#Const A = True Rem note\n#End If",
        "#Const A = True ' Rem note",
        '#Const A = "Rem"',
        "#Const Remark = 1",
    ],
)
def test_rem_on_const_line_quiet(line: str) -> None:
    assert _malformed(_mod(line)) == []


@pytest.mark.parametrize("name", ["Empty", "True", "If", "Not", "String", "Debug", "Me"])
def test_reserved_const_name(name: str) -> None:
    src = _mod(f"#Const {name} = 1")
    found = _malformed(src)
    assert len(found) == 1
    assert _span_text(src, found[0]) == name
    assert "reserved word" in found[0].message
    assert "Expected: identifier" in found[0].message


def test_reserved_const_name_reported_once_with_rem() -> None:
    src = _mod("#Const Empty = 1 Rem note")
    found = _malformed(src)
    assert len(found) == 1
    assert _span_text(src, found[0]) == "Empty"


@pytest.mark.parametrize(
    "line",
    ["#Const Error = 1", "#Const Win64 = 1", "#Const Object = 1", "#Const [Rem] = 1",
     "#If False Then\n#Const Empty = 1\n#End If"],
)  # fmt: skip
def test_allowed_const_names(line: str) -> None:
    assert _malformed(_mod(line)) == []


# -- Reserved names the lexer reads as identifiers (commit 475b0215) ---------


@pytest.mark.parametrize("name", ["vb_name", "VB_Description", "lineinput", "attribute", "local"])
def test_reserved_const_name_read_as_identifier(name: str) -> None:
    src = _mod(f"#Const {name} = 1")
    found = _malformed(src)
    assert len(found) == 1
    assert _span_text(src, found[0]) == name
    assert "Expected: identifier" in found[0].message


@pytest.mark.parametrize("name", ["[vb_name]", "vb_foo", "lineinputx", "Local2"])
def test_near_reserved_const_names_are_quiet(name: str) -> None:
    assert _malformed(_mod(f"#Const {name} = 1")) == []
