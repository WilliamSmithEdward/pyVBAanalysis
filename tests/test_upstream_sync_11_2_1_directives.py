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
from pyvbaanalysis.conditional import (
    ConditionalCompilationEnvironment,
    conditional_activity_at_offset,
    create_conditional_activity_tracker,
)
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.diagnostics.model import DiagnosticSeverity
from pyvbaanalysis.diagnostics.model import VbaDiagnostic as Diagnostic
from pyvbaanalysis.parser import parse_module
from pyvbaanalysis.parser.nodes import Span


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


# -- A directive that does not start its logical line (XLIDE issue #1331,
#    commit bd5eccec) ---------------------------------------------------------


@pytest.mark.parametrize(
    "source",
    [
        _plain_fn("    x = 1: #If True Then", "    Main = 1", "#End If"),
        _plain_fn("    x = 1: _", "#If True Then", "    Main = 1", "#End If"),
        _plain_fn("L1: #If True Then", "    Main = 1", "#End If"),
        _plain_fn("10 #If True Then", "    Main = 1", "#End If"),
        _plain_fn("#If True Then", "    Main = 1", "    x = 1: #End If"),
        _plain_fn("#If True Then", "    Main = 1", "10 #End If"),
        _plain_fn("#If True Then", "    Main = 1", "    x = 1: #Else", "    Main = 2", "#End If"),
        _plain_fn("#If True Then", "    Main = 1", "10 #Else", "    Main = 2", "#End If"),
        _plain_fn("    x = 1: #Const A = 1", "    Main = 1"),
        _plain_fn("10 #Const A = 1", "    Main = 1"),
        "#Const B = 1: #If False Then\n#End If\n" + _plain_fn("    Main = 1"),
    ],
)
def test_misplaced_directive_refused_where_compiled(source: str) -> None:
    assert _errors(source) != []


@pytest.mark.parametrize(
    "source",
    [
        _plain_fn("    Main = 1", "    #If True Then", "    Main = 2", "    #End If"),
        _plain_fn("    Main = 1", "\t#If True Then", "    Main = 2", "\t#End If"),
        _plain_fn("    Main = 1", " _", "#If True Then", "    Main = 2", "#End If"),
        _plain_fn("    Main = 1", "#If False Then", "    x = 1: #Else", _ZERO, "#End If"),
    ],
)
def test_misplaced_directive_quiet_where_excel_compiles(source: str) -> None:
    assert _errors(source) == []


@pytest.mark.parametrize(
    "line",
    [
        "Private y As Long: #Const A = 1",
        "Private y As Long: _\n#Const A = 1",
        "L1: #Const A = 1",
        "10 #Const A = 1",
        "#Const B = 1: #Const A = 1",
        "#Else: #Const A = 1",
    ],
)
def test_misplaced_const_defines_nothing_even_in_skipped_arm(line: str) -> None:
    later = _plain_fn("    Main = 1", "#If A = 1 Then", _ZERO, "#End If")
    assert _errors(f"#If False Then\n{line}\n#End If\n{later}") == []


def test_misplaced_directive_reported_once() -> None:
    def once(source: str) -> list[Diagnostic]:
        return [
            d
            for d in analyze_module(source)
            if d.code in ("malformed-statement", "directive-trailing-statement")
        ]

    assert len(once(_plain_fn("    x = 1: #Const A = 1", "    Main = 1"))) == 1
    assert len(once("#Const B = 1: #Const A = 1\n" + _plain_fn("    Main = 1"))) == 1


def test_misplaced_directive_messages() -> None:
    source = _plain_fn("10 #If True Then", "    Main = 1", "#End If")
    found = _malformed(source)
    assert [_span_text(source, d) for d in found] == ["#If"]
    assert "after a line number" in found[0].message
    assert found[0].message.endswith("Syntax error.")
    top = "Private y As Long: #Const A = 1\n" + _plain_fn("    Main = 1")
    found = _malformed(top)
    assert [_span_text(top, d) for d in found] == ["#Const"]
    assert "after a ':'" in found[0].message
    assert "Expected: line number or label or statement or end of statement" in found[0].message


@pytest.mark.parametrize("eol", ["\n", "\r\n", "\r"])
def test_colon_directive_findings_keep_their_lines(eol: str) -> None:
    source = eol.join(
        ["#Const A = 1: Debug.Print 1", "#Const B = 2: 'comment", "#Const C = 3",
         "#Const D = 4:Debug.Print 4"]
    )  # fmt: skip
    found = _by_code(source, "directive-trailing-statement")
    assert [_span_text(source, d) for d in found] == ["Debug.Print 1", "'comment", "Debug.Print 4"]


# -- String directive conditions (XLIDE issue #1339, commit 09dabe4f) --------

_PROJECT_ENV = ConditionalCompilationEnvironment(project_constants={})


def _arms(condition: str, setup: str = "") -> str:
    return (
        f"{setup}Function Main() As Variant\n#If {condition} Then\n    Main = 1\n#Else\n"
        "    Main = 2\n#End If\nEnd Function\n"
    )


def _activities(src: str, env: ConditionalCompilationEnvironment | None = _PROJECT_ENV) -> list[str]:
    tracker = create_conditional_activity_tracker(parse_module(src), env)
    assert tracker is not None
    out = []
    for text in ("Main = 1", "Main = 2"):
        at = src.index(text)
        out.append(tracker.activity_for_span(Span(at, at + 1)).value)
    return out


def _mismatches(src: str) -> list[Diagnostic]:
    return _by_code(src, "directive-type-mismatch")


def _sfn(*lines: str) -> str:
    return "\n".join(["Function Main() As Variant", *lines, "    Main = 1", "End Function", ""])


@pytest.mark.parametrize(
    "condition, setup",
    [('"0"', ""), ('"00"', ""), ("0 & 0", ""), ('"False"', ""), ('"0.0"', ""),
     ("S", '#Const S = "0"\n'), ('" 1 " - 1', ""), ('"&H10" And 1', "")],
)  # fmt: skip
def test_string_condition_takes_else_arm(condition: str, setup: str) -> None:
    assert _activities(_arms(condition, setup)) == ["inactive", "active"]


@pytest.mark.parametrize(
    "condition",
    ['"1"', '"2"', '"-1"', '"True"', '"true"', '"&H10"', '"1e3"', '"2.5"', '"1,000"', 'Not "0"',
     '"1" And 1'],
)  # fmt: skip
def test_string_condition_takes_if_arm(condition: str) -> None:
    assert _activities(_arms(condition)) == ["active", "inactive"]


def test_string_condition_in_elseif_and_nested_if() -> None:
    else_if = (
        'Function Main() As Variant\n#If False Then\n    Main = 0\n#ElseIf "0" Then\n'
        "    Main = 1\n#Else\n    Main = 2\n#End If\nEnd Function\n"
    )
    assert _activities(else_if) == ["inactive", "active"]
    nested = (
        'Function Main() As Variant\n#If True Then\n#If "0" Then\n    Main = 1\n#Else\n'
        "    Main = 2\n#End If\n#End If\nEnd Function\n"
    )
    assert _activities(nested) == ["inactive", "active"]


def test_string_condition_analyzes_the_compiled_arm() -> None:
    src = (
        'Option Explicit\nSub S()\n#If "0" Then\n    skipped = 1\n#Else\n    compiled = 1\n'
        "#End If\nEnd Sub\n"
    )
    assert _undeclared(src) == ["compiled"]


@pytest.mark.parametrize(
    "condition",
    ['""', '"abc"', '" True "', '"$5"', '"1e400"', '"1.5.5"', '"&H100000000"',
     '"&H100000000" And 1'],
)  # fmt: skip
def test_string_condition_left_unknown(condition: str) -> None:
    assert _activities(_arms(condition)) == ["unknown", "unknown"]


@pytest.mark.parametrize(
    "directive, shown",
    [
        ('#If "abc" Then', '"abc"'),
        ('#If "" Then', '""'),
        ('#If " True " Then', '" True "'),
        ('#If Not "abc" Then', '"abc"'),
        ('#If "abc" = 1 Then', '"abc"'),
        ('#If "a" & "b" Then', '"ab"'),
        ('#If UNDEFINED_X - "abc" Then', '"abc"'),
        ('#If "abc" = #1/1/2020# Then', '"abc"'),
        ('#If #1/1/2020# + "abc" Then', '"abc"'),
        ('#If Null - "abc" Then', '"abc"'),
        ('#If "abc" * Null Then', '"abc"'),
    ],
)
def test_directive_type_mismatch(directive: str, shown: str) -> None:
    src = _sfn(directive, "#End If")
    found = _mismatches(src)
    assert len(found) == 1
    assert _span_text(src, found[0]) == directive
    assert found[0].severity is DiagnosticSeverity.ERROR
    assert shown in found[0].message
    assert "Type mismatch" in found[0].message


def test_type_mismatch_after_string_const() -> None:
    read = '#Const S = "x"\n' + _sfn("#If S Then", "#End If")
    found = _mismatches(read)
    assert [_span_text(read, d) for d in found] == ["#If S Then"]
    assert _mismatches('#Const S = "x"\n' + _sfn()) == []


def test_type_mismatch_in_const_value() -> None:
    unread = '#Const A = "abc" + 1\n' + _sfn()
    found = _mismatches(unread)
    assert [_span_text(unread, d) for d in found] == ['#Const A = "abc" + 1']
    assert "#Const value" in found[0].message
    assert len(_mismatches('#If False Then\n#Const A = "abc" + 1\n#End If\n' + _sfn())) == 1


@pytest.mark.parametrize("condition", ['"abc"', 'Not "abc"', '"a" & "b"', 'UNDEFINED_X - "abc"'])
def test_type_mismatch_in_skipped_arm(condition: str) -> None:
    src = _sfn("#If False Then", f"#If {condition} Then", "#End If", "#End If")
    assert [_span_text(src, d) for d in _mismatches(src)] == [f"#If {condition} Then"]


def test_type_mismatch_after_taken_arm() -> None:
    skipped = '#Const S = "x"\n' + _sfn("#If False Then", "#If S Then", "#End If", "#End If")
    assert len(_mismatches(skipped)) == 1
    assert len(_mismatches(_sfn("#If True Then", '#ElseIf "abc" Then', "#End If"))) == 1


@pytest.mark.parametrize(
    "condition",
    ['"0"', '"1"', '"True"', '"false"', '"2.5"', '"1,000"', '"&H10" And 1', '" 1 " - 1',
     'Not "0"', '"1" And 1', '"2" - 2'],
)  # fmt: skip
def test_converted_strings_are_quiet(condition: str) -> None:
    assert _mismatches(_sfn(f"#If {condition} Then", "#End If")) == []


def test_string_beside_date_and_null() -> None:
    date = _arms('"1/1/2020" = #1/1/2020#')
    assert _mismatches(date) == []
    assert _activities(date) == ["unknown", "unknown"]
    for condition in ('"abc" = Null', 'Null + "abc"'):
        src = _sfn(f"#If {condition} Then", "#End If")
        assert _mismatches(src) == []
        assert len(_by_code(src, "null-directive-condition")) == 1


@pytest.mark.parametrize(
    "condition, conversion",
    [('1 = "True"', "as a number"), ('"abc" = True', "as True or False"),
     ('"abc" = -1', "as a number"), ('"abc" = #1/1/2020#', "as a date"),
     ('(1) = "abc"', "as a number"), ('"abc" + (1)', "as a number")],
)  # fmt: skip
def test_conversion_beside_a_literal(condition: str, conversion: str) -> None:
    found = _mismatches(_arms(condition))
    assert len(found) == 1
    assert conversion in found[0].message


@pytest.mark.parametrize(
    "condition", ['"True" = True', 'UNDEFINED_X = "abc"', 'UNDEFINED_X = ""', 'UNDEFINED_X = "0"']
)
def test_no_conversion_beside_a_name(condition: str) -> None:
    assert _mismatches(_arms(condition)) == []
    assert _mismatches(_sfn("#If False Then", f"#If {condition} Then", "#End If", "#End If")) == []


def test_variant_comparison_arms() -> None:
    assert _mismatches(_arms('A = "abc"', "#Const A = 1\n")) == []
    assert _activities(_arms('UNDEFINED_X = "0"')) == ["active", "inactive"]
    assert _activities(_arms('"True" = True')) == ["unknown", "unknown"]


def test_host_strings_report_nothing_until_redefined() -> None:
    def mismatches(src: str, env: ConditionalCompilationEnvironment) -> list[Diagnostic]:
        options = AnalyzeModuleOptions(known_identifiers=set(), conditional_compilation=env)
        return [d for d in analyze_module(src, options) if d.code == "directive-type-mismatch"]

    host = ConditionalCompilationEnvironment(project_constants={"FLAG": "x"})
    read = _sfn("#If FLAG Then", "#End If")
    assert mismatches(read, host) == []
    assert len(mismatches('#Const FLAG = "y"\n' + read, host)) == 1
    failed = "#Const FLAG = 1 / 0\n" + read
    assert mismatches(failed, host) == []
    compiler = ConditionalCompilationEnvironment(
        compiler_constants={"FLAG": "x"}, project_constants={}
    )
    assert mismatches(read, compiler) == []
    assert mismatches(failed, compiler) == []


def test_skipped_if_after_string_const_in_skipped_arm() -> None:
    src = _sfn("#If False Then", '#Const S = "x"', "#If S Then", "#End If", "#End If")
    found = _mismatches(src)
    assert [_span_text(src, d) for d in found] == ["#If S Then"]
    assert "as True or False" in found[0].message


def test_unknown_constant_concatenation_is_quiet() -> None:
    src = _sfn('#If (MISSING & "x") - 1 Then', "#End If")
    assert [d for d in analyze_module(src) if d.code == "directive-type-mismatch"] == []


@pytest.mark.parametrize(
    "condition", ['VBA7 And "x" = "x"', '"x" = Empty', 'Empty = "x"', '"x" + "y" = "xy"', 'S & "y" = "xy"']
)
def test_string_operations_are_quiet(condition: str) -> None:
    assert _mismatches(_arms(condition, '#Const S = "x"\n')) == []


@pytest.mark.parametrize(
    "condition",
    ['"abc" Mod 2', '"abc" \\ 2', '"abc" * 2', '"abc" / 2', '-"abc"', '+"abc"', '"abc" ^ 2',
     '2 ^ -"abc"', '"abc" And 1', '"abc" Or 1', '"abc" Xor 1', '"abc" Eqv 1', '"abc" Imp 1',
     '1 + "abc"', '"abc" < 1'],
)  # fmt: skip
def test_every_operator_converts(condition: str) -> None:
    assert len(_mismatches(_arms(condition))) == 1


@pytest.mark.parametrize(
    "condition", ['"6" Mod 4', '"6" \\ 4', '"6" * 1', '-"1"', '"True" Or 0', '"1" + 1 = 2', '"2" < 3']
)
def test_every_operator_converts_good_strings(condition: str) -> None:
    assert _mismatches(_arms(condition)) == []


@pytest.mark.parametrize(
    "condition", ['("abc" - 1', '"abc" -', '"abc" 1', '"abc" - Not', '"abc" - "oops']
)
def test_unreadable_expression_reports_nothing(condition: str) -> None:
    assert _mismatches(_arms(condition)) == []


def test_mismatch_message_quotes_the_string() -> None:
    long = "x" * 50
    found = _mismatches(_arms(f'"a""b" & "{long}"'))
    assert '"a""b' + "x" * 37 + '..."' in found[0].message
    assert "#ElseIf condition" in _mismatches(_sfn("#If False Then", '#ElseIf "abc" Then', "#End If"))[0].message


# -- A #Const whose value reads a name with no value (XLIDE issue #1347,
#    commit 518dd817) ---------------------------------------------------------


def _win64_module(lines: str) -> str:
    return (
        f"{lines}\nSub Main()\n#If Win64 Then\n    Debug.Print 1\n#Else\n    Debug.Print 2\n"
        "#End If\nEnd Sub\n"
    )


def _win64_arms(source: str, env: ConditionalCompilationEnvironment | None = None) -> list[str]:
    parsed = parse_module(source)
    tracker = create_conditional_activity_tracker(parsed, env)
    assert tracker is not None
    out = []
    for text in ("Debug.Print 1", "Debug.Print 2"):
        at = source.index(text)
        activity = tracker.activity_for_span(Span(at, at + 1))
        # The offset replay shares the change.
        assert conditional_activity_at_offset(parsed, at, env) is activity
        out.append(activity.value)
    return out


@pytest.mark.parametrize("value", ["UNDEFINED_X", "MISSING + 1"])
def test_const_reading_unknown_name_shadows_builtin(value: str) -> None:
    assert _win64_arms(_win64_module(f"#Const Win64 = {value}")) == ["unknown", "unknown"]


def test_const_reading_undefined_name_with_project_arguments() -> None:
    assert _win64_arms(_win64_module("#Const Win64 = UNDEFINED_X"), _PROJECT_ENV) == [
        "inactive",
        "active",
    ]
    assert _win64_arms(_win64_module("#Const Win64 = MISSING + 1"), _PROJECT_ENV) == [
        "active",
        "inactive",
    ]
    env = ConditionalCompilationEnvironment(project_constants={"MISSING": -1})
    assert _win64_arms(_win64_module("#Const Win64 = MISSING + 1"), env) == ["inactive", "active"]


def test_const_builtin_value_and_default() -> None:
    assert _win64_arms(_win64_module("#Const Win64 = 0")) == ["inactive", "active"]
    assert _win64_arms(_win64_module("' no #Const")) == ["active", "inactive"]


def test_const_reading_unknown_name_shadows_twinbasic() -> None:
    source = "#Const TWINBASIC = UNDEFINED_X\n#If TWINBASIC Then\nDim a\n#Else\nDim b\n#End If\n"
    tracker = create_conditional_activity_tracker(parse_module(source))
    assert tracker is not None
    at = source.index("Dim b")
    assert tracker.activity_for_span(Span(at, at + 1)).value == "unknown"


def test_const_reading_unknown_const_is_unknown() -> None:
    source = _win64_module("#Const VBA7 = UNDEFINED_X\n#Const Win64 = VBA7 + 1")
    assert _win64_arms(source) == ["unknown", "unknown"]


@pytest.mark.parametrize(
    "value",
    ["1 / 0", "UNDEFINED_X +", "(UNDEFINED_X", "UNDEFINED_X 1", "UNDEFINED_X + ,",
     "UNDEFINED_X + Then", 'UNDEFINED_X + "bad', "[UNDEFINED_X]"],
)  # fmt: skip
def test_const_failing_otherwise_keeps_earlier_value(value: str) -> None:
    assert _win64_arms(_win64_module(f"#Const Win64 = {value}")) == ["active", "inactive"]


@pytest.mark.parametrize("condition", ["True", "False"])
def test_duplicate_const_across_arms(condition: str) -> None:
    src = (
        f"#If {condition} Then\n#Const A = 1\n#Else\n#Const A = 2\n#End If\n"
        "Function Main() As Variant\n    Main = 1\nEnd Function\n"
    )
    assert len(_by_code(src, "duplicate-const-directive")) == 1
