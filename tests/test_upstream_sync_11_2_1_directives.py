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
from pyvbaanalysis.diagnostics.model import DiagnosticSeverity
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


# -- A colon after a compiler directive (XLIDE issue #1323, commit b730a081) --


def _errors(source: str) -> list[str]:
    return [d.code for d in _analyze(source) if d.severity is DiagnosticSeverity.ERROR]


def _plain_fn(*lines: str) -> str:
    return "\n".join(["Function Main() As Variant", *lines, "End Function", ""])


_ZERO = "    Main = 1 / 0"


@pytest.mark.parametrize(
    "source",
    [
        _plain_fn("#If True Then:", "    Main = 1", "#End If"),
        _plain_fn("#If True Then :", "    Main = 1", "#End If"),
        _plain_fn("#If True Then: ' note", "    Main = 1", "#End If"),
        _plain_fn("#If True Then", "    Main = 1", "#End If:"),
        _plain_fn("#If True Then", "    Main = 1", "#End If: ' note"),
        _plain_fn("#If True Then", "    Main = 1", "#Else:", "    Main = 2", "#End If"),
        _plain_fn("#If True Then", "    Main = 1", "#Else: Main = 2", "#End If"),
        "#Const A = 1:\n" + _plain_fn("    Main = 1"),
        "#Const A = 1: ' note\n" + _plain_fn("    Main = 1"),
        "#Const A = 1: #Const B = 2\n" + _plain_fn("    Main = 1"),
    ],
)
def test_colon_directive_refused_where_compiled(source: str) -> None:
    assert "directive-trailing-statement" in _errors(source)


@pytest.mark.parametrize(
    "source",
    [
        _plain_fn("#If False Then", "#If True Then", "#Else: Main = 2", "#End If", "#End If",
                  "    Main = 1"),
        "#If False Then\n#Const A = 1: x = 1\n#End If\n" + _plain_fn("    Main = 1"),
        "#If False Then\n#Const A = 1: Rem note\n#End If\n" + _plain_fn("    Main = 1"),
        _plain_fn("    Main = 1", "#If #1/1/2000 12:00:00 AM# < 1 Then", "    Main = 2", "#End If"),
    ],
)  # fmt: skip
def test_colon_directive_quiet_where_compiled_by_excel(source: str) -> None:
    assert _errors(source) == []


@pytest.mark.parametrize(
    "line",
    ["#Else:", "#Else: ' note", "#Else: Rem note", "#ElseIf True Then: ' note",
     "#ElseIf True Then:"],
)  # fmt: skip
def test_colon_else_takes_no_new_arm(line: str) -> None:
    source = _plain_fn("    Main = 1", "#If False Then", "    Main = 2", line, _ZERO, "#End If")
    assert _errors(source) == []


def test_plain_else_takes_its_arm() -> None:
    source = _plain_fn("    Main = 1", "#If False Then", "    Main = 2", "#Else", _ZERO, "#End If")
    assert "division-by-zero" in _errors(source)


def test_colon_const_defines_nothing_even_in_skipped_arm() -> None:
    later = _plain_fn("    Main = 1", "#If A = 1 Then", _ZERO, "#End If")
    assert _errors("#If False Then\n#Const A = 1: Rem note\n#End If\n" + later) == []
    assert _errors("#If False Then\n#Const A = 1: x = 1\n#End If\n" + later) == []
    # A plain #Const in a skipped arm still defines A (XLIDE issue #192).
    assert "division-by-zero" in _errors("#If False Then\n#Const A = 1\n#End If\n" + later)


def test_colon_directive_spans() -> None:
    source = _plain_fn("#If True Then", "    Main = 1", "#Else: Main = 2", "#End If:")
    found = _by_code(source, "directive-trailing-statement")
    assert [_span_text(source, d) for d in found] == ["Main = 2", ":"]
    assert "starts no new arm" in found[0].message
    assert "closes no #If block" in found[1].message
    const = "#Const A = 1: ' note\n" + _plain_fn("    Main = 1")
    found = _by_code(const, "directive-trailing-statement")
    assert [_span_text(const, d) for d in found] == ["' note"]
    assert "defines no constant" in found[0].message


# -- A #Const line holding a Rem defines nothing (commit e3eca840) -----------


@pytest.mark.parametrize(
    "line, taken",
    [
        ("#Const A = 1", "a"),
        ("#Const A = 1 ' note", "a"),
        ("#Const A = 1 Rem note", "b"),
        ("#Const A = 1 _\n    Rem note", "b"),
    ],
)
def test_rem_const_defines_nothing_even_in_skipped_arm(line: str, taken: str) -> None:
    source = (
        f"Option Explicit\n#If False Then\n{line}\n#End If\nSub S()\n#If A = 1 Then\n"
        "    a = 1\n#Else\n    b = 2\n#End If\nEnd Sub\n"
    )
    assert _undeclared(source) == [taken]


def test_rem_const_after_colon_is_directive_trailing_statements() -> None:
    src = _mod("#Const A = 1: #Const B = 1 Rem n")
    assert len(_by_code(src, "directive-trailing-statement")) == 1
    assert _malformed(src) == []
    skipped = _mod("#If False Then\n#Const A = 1: #Const B = 1 Rem n\n#End If")
    assert _by_code(skipped, "directive-trailing-statement") == []
    assert _malformed(skipped) == []


@pytest.mark.parametrize("eol", ["\n", "\r\n", "\r"])
def test_colon_directive_findings_keep_their_lines(eol: str) -> None:
    source = eol.join(
        ["#Const A = 1: Debug.Print 1", "#Const B = 2: 'comment", "#Const C = 3",
         "#Const D = 4:Debug.Print 4"]
    )  # fmt: skip
    found = _by_code(source, "directive-trailing-statement")
    assert [_span_text(source, d) for d in found] == ["Debug.Print 1", "'comment", "Debug.Print 4"]
