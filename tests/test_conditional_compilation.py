"""M3: conditional-compilation activity (conditionalCompilation.ts parity)."""

from __future__ import annotations

from pyvbaanalysis.conditional import (
    ConditionalActivity,
    ConditionalCompilationEnvironment,
    conditional_activity_for_span,
    create_conditional_activity_tracker,
    evaluate_conditional_expression,
    index_conditional_compilation,
    module_has_conditional_directives,
    parse_project_conditional_constants,
)
from pyvbaanalysis.parser import parse_module
from pyvbaanalysis.parser.nodes import Span

_ACTIVE = ConditionalActivity.ACTIVE
_INACTIVE = ConditionalActivity.INACTIVE
_UNKNOWN = ConditionalActivity.UNKNOWN


def _activity_at(src: str, needle: str) -> ConditionalActivity:
    module = parse_module(src)
    off = src.index(needle)
    return conditional_activity_for_span(module, Span(off, off))


def test_default_compiler_constants() -> None:
    # 64-bit Office on Windows: Win32 means Windows and is on there too (XLIDE
    # issue #192); an on constant is 1, not True (XLIDE issue #214).
    assert _activity_at("#If VBA7 Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#If VBA6 Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#If Win64 Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#If Win32 Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#If Win16 Then\nDim a\n#End If", "Dim a") is _INACTIVE
    assert _activity_at("#If Mac Then\nDim a\n#End If", "Dim a") is _INACTIVE
    # Not 1 is -2, which is True.
    assert _activity_at("#If Not Win64 Then\nDim a\n#End If", "Dim a") is _ACTIVE


def test_unknown_constant_stays_unknown() -> None:
    assert _activity_at("#If NotDefined Then\nDim a\n#End If", "Dim a") is _UNKNOWN


def test_project_const_definition() -> None:
    on = "#Const Custom = 1\n#If Custom Then\nDim a\n#End If"
    off = "#Const Custom = 0\n#If Custom Then\nDim a\n#End If"
    assert _activity_at(on, "Dim a") is _ACTIVE
    assert _activity_at(off, "Dim a") is _INACTIVE


def test_elseif_chain() -> None:
    src = (
        "#If Mac Then\n"
        "Dim a\n"
        "#ElseIf VBA7 Then\n"
        "Dim b\n"
        "#Else\n"
        "Dim c\n"
        "#End If"
    )
    assert _activity_at(src, "Dim a") is _INACTIVE
    assert _activity_at(src, "Dim b") is _ACTIVE
    assert _activity_at(src, "Dim c") is _INACTIVE  # an earlier arm was true


def test_else_after_false_is_active() -> None:
    src = "#If Mac Then\nDim a\n#Else\nDim b\n#End If"
    assert _activity_at(src, "Dim a") is _INACTIVE
    assert _activity_at(src, "Dim b") is _ACTIVE


def test_else_after_unknown_is_unknown() -> None:
    src = "#If Mystery Then\nDim a\n#Else\nDim b\n#End If"
    assert _activity_at(src, "Dim a") is _UNKNOWN
    assert _activity_at(src, "Dim b") is _UNKNOWN


def test_comparison_operators() -> None:
    # VBA7 is 1, not True (-1): `#If VBA7 = True` is false (XLIDE issue #214).
    assert _activity_at("#If VBA7 = True Then\nDim a\n#End If", "Dim a") is _INACTIVE
    assert _activity_at("#If VBA7 = 1 Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#If Mac <> True Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#If VBA7 <> True Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#If VBA7 <> 1 Then\nDim a\n#End If", "Dim a") is _INACTIVE


def test_not_and_or() -> None:
    assert _activity_at("#If Not Mac Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#If VBA7 And Win64 Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#If VBA7 And Mac Then\nDim a\n#End If", "Dim a") is _INACTIVE
    assert _activity_at("#If Mac Or VBA7 Then\nDim a\n#End If", "Dim a") is _ACTIVE


def test_numeric_equality_parity() -> None:
    # #Const N = 5 stored as integer-valued; N = 5 compares equal (no "5.0" drift).
    src = "#Const N = 5\n#If N = 5 Then\nDim a\n#End If"
    assert _activity_at(src, "Dim a") is _ACTIVE


def test_hex_and_octal_literals_evaluate() -> None:
    # XLIDE v2.5.8: &H/&O literals parse with the correct radix, so a hex
    # condition evaluates instead of falling to the unknown branch.
    assert _activity_at("#If &HFF Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#If &H0 Then\nDim a\n#End If", "Dim a") is _INACTIVE
    assert _activity_at("#If &O17 = 15 Then\nDim a\n#End If", "Dim a") is _ACTIVE


def test_relational_operators() -> None:
    # XLIDE v2.5.8: <, >, <=, >= join = and <> (booleans coerce to -1/0). Win64
    # is 1 since XLIDE issue #214.
    assert _activity_at("#If Win64 >= 1 Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#If Win64 < 0 Then\nDim a\n#End If", "Dim a") is _INACTIVE
    assert _activity_at("#If True < 0 Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#Const N = 5\n#If N > 4 Then\nDim a\n#End If", "Dim a") is _ACTIVE
    assert _activity_at("#Const N = 5\n#If N < 4 Then\nDim a\n#End If", "Dim a") is _INACTIVE
    # A non-numeric operand stays unmodeled -> unknown.
    assert _activity_at('#Const S = "x"\n#If S > 1 Then\nDim a\n#End If', "Dim a") is _UNKNOWN


def test_string_constant_comparison() -> None:
    src = '#Const S = "x"\n#If S = "x" Then\nDim a\n#End If'
    assert _activity_at(src, "Dim a") is _ACTIVE


def test_nested_inactive_outer_keeps_inner_inactive() -> None:
    src = (
        "#If Mac Then\n"
        "#If VBA7 Then\n"
        "Dim a\n"
        "#End If\n"
        "#End If"
    )
    assert _activity_at(src, "Dim a") is _INACTIVE


def test_const_in_inactive_branch_is_applied() -> None:
    # The VBE reads every #Const line, one inside a #If False included (XLIDE
    # issue #192, measured in Excel 16.0).
    src = (
        "#If Mac Then\n"
        "#Const Flag = 1\n"
        "#End If\n"
        "#If Flag Then\n"
        "Dim a\n"
        "#End If"
    )
    assert _activity_at(src, "Dim a") is _ACTIVE


def test_directives_inside_procedure_body() -> None:
    src = (
        "Sub S\n"
        "#If Mac Then\n"
        "    Dim a As Long\n"
        "#End If\n"
        "End Sub"
    )
    assert _activity_at(src, "Dim a") is _INACTIVE


def test_tracker_none_without_directives() -> None:
    module = parse_module("Sub S\n    x = 1\nEnd Sub")
    assert create_conditional_activity_tracker(module) is None
    assert not module_has_conditional_directives(module)


def test_tracker_binary_search_matches_offsets() -> None:
    src = "#If Mac Then\nDim a\n#Else\nDim b\n#End If"
    module = parse_module(src)
    tracker = create_conditional_activity_tracker(module)
    assert tracker is not None
    a_off = src.index("Dim a")
    b_off = src.index("Dim b")
    assert tracker.is_inactive(Span(a_off, a_off))
    assert not tracker.is_inactive(Span(b_off, b_off))


def test_evaluate_conditional_expression_direct() -> None:
    assert evaluate_conditional_expression("1 = 1") is True
    assert evaluate_conditional_expression("1 = 2") is False
    assert evaluate_conditional_expression("Not False") is True
    assert evaluate_conditional_expression("") is None
    assert evaluate_conditional_expression(None) is None
    env = ConditionalCompilationEnvironment(compiler_constants={"Flag": True})
    assert evaluate_conditional_expression("Flag And True", env) is True
    # Bare unknown name resolves to undefined.
    assert evaluate_conditional_expression("Missing") is None


def test_index_conditional_compilation_constants() -> None:
    src = "#Const A = 1\n#Const B = 2\n#If A Then\nDim x\n#End If"
    module = parse_module(src)
    index = index_conditional_compilation(module)
    consts = {c.name: c.value for c in index.constants}
    assert consts == {"A": 1, "B": 2}
    assert len(index.directives) >= 4  # 2 Const, 1 If, 1 EndIf


def test_parse_project_conditional_constants() -> None:
    assert parse_project_conditional_constants("DEBUG_MODE = 1 : Level=-2 : Tag = beta") == {
        "DEBUG_MODE": 1,
        "Level": -2,
        "Tag": "beta",
    }
    assert parse_project_conditional_constants(None) == {}
    assert parse_project_conditional_constants("") == {}
    # Not an identifier, or no `=`: skipped, as upstream skips it.
    assert parse_project_conditional_constants("2nd = 1 : Bad Name = 1 : Loose") == {}
    # The value pattern is ASCII: an Arabic-Indic digit is text. A name is a VBA
    # identifier, any letter opening it (XLIDE issue #207), and no underscore.
    assert parse_project_conditional_constants("A = ٣ : État = 1") == {"A": "٣", "État": 1}
    assert parse_project_conditional_constants("_X = 1") == {}


def test_directive_values_beyond_numbers() -> None:
    # XLIDE issue #208: Empty, Null, Nothing and dates; #192: VBA's operators.
    assert evaluate_conditional_expression("#1/2/2000# > #1/1/2000#") is True
    assert evaluate_conditional_expression("#12:00:00 AM#") is not None
    assert evaluate_conditional_expression("Null Or True") is True
    assert evaluate_conditional_expression("Null And False") is False
    assert evaluate_conditional_expression("Nothing Is Nothing") is True
    assert evaluate_conditional_expression('"ABC" Like "a*"') is True
    assert evaluate_conditional_expression('"A" = "a"') is True
    assert evaluate_conditional_expression("1 And 2") == 0
    assert evaluate_conditional_expression("Not 1") == -2
    assert evaluate_conditional_expression("-2 ^ 2") == -4
    assert evaluate_conditional_expression("7 \\ 2 + 7 Mod 4") == 6
    assert evaluate_conditional_expression("&HFFFF") == -1
    assert evaluate_conditional_expression('Empty & "x" = "x"') is True


def test_null_condition_directives() -> None:
    from pyvbaanalysis.conditional.conditional_compilation import null_condition_directives

    src = "#Const N = Null\n#If N Then\nDim a\n#ElseIf Null Then\nDim b\n#End If\n#If Mac Then\n#If Null Then\n#End If\n#End If"
    module = parse_module(src)
    found = null_condition_directives(module)
    assert [d.span.start for d in found] == [src.index("#If N")]
    assert null_condition_directives(module, None) == found
