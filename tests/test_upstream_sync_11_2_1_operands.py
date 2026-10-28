"""Rule contracts from the XLIDE 11.1.0 -> 11.2.1 sync (operands, bounds, Is).

Each section names the upstream commit whose own tests it follows. The oracle
cases are read from the vendored vbe_oracle_cases.json and analyzed as the
extension analyzes the case's host.
"""

from __future__ import annotations

import json
import re
from typing import Any

import pytest

from pyvbaanalysis import analyze_module, analyze_module_options_for
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics.model import DiagnosticSeverity
from pyvbaanalysis.evidence import DATA_DIR
from pyvbaanalysis.runtime.vba_value_results import (
    ARGUMENT_TYPED_FUNCTIONS,
    BARE_VALUE_RESULT_FUNCTIONS,
    TYPEOF_KEYWORD_OPERANDS,
    VALUE_RESULT_FUNCTIONS,
)
from pyvbaanalysis.symbols import ModuleSymbolKind, ProjectIndex
from oracle_support import oracle_module_input  # type: ignore[attr-defined]

_ORACLE: dict[str, dict[str, Any]] = {
    case["id"]: case
    for case in json.loads((DATA_DIR / "vbe_oracle_cases.json").read_text(encoding="utf-8"))["cases"]
}

_KIND = {
    "standard": ModuleSymbolKind.STANDARD,
    "class": ModuleSymbolKind.CLASS,
    "document": ModuleSymbolKind.DOCUMENT,
    "userform": ModuleSymbolKind.USERFORM,
}


def _case(case_id: str) -> dict[str, Any]:
    return _ORACLE[case_id]


def _modules(case: dict[str, Any]) -> list[dict[str, Any]]:
    if "modules" in case:
        return list(case["modules"])
    return [{"name": "Module1", "type": "standard", "source": case["source"]}]


def _errors_for(case: dict[str, Any], ignore: frozenset[str] = frozenset()) -> list[str]:
    """Each error as `code: spanned text`, analyzed in the case's host."""
    host = case.get("host", "excel")
    modules = _modules(case)
    index = ProjectIndex()
    for module in modules:
        # What the designer and the attributes declare too (XLIDE issue #1419).
        index.set_module(oracle_module_input(module))
    out: list[str] = []
    for module in modules:
        kind = _KIND.get(module.get("type", "standard"), ModuleSymbolKind.STANDARD)
        opts = analyze_module_options_for(
            index, module["name"], kind, host=host,
            conditional_compilation=ConditionalCompilationEnvironment(project_constants={}),
        )
        source = module["source"]
        out += [
            f"{d.code}: {source[d.span.start:d.span.end]}"
            for d in analyze_module(source, opts)
            if d.severity is DiagnosticSeverity.ERROR and d.code not in ignore
        ]
    return out


def _codes_of(errors: list[str]) -> list[str]:
    return [finding.split(":")[0] for finding in errors]


_HOSTS = ("", "word_", "powerpoint_", "access_")


def _each_host(name: str, phase: str) -> list[str]:
    return [f"issue1412_{host}{name}_{phase}" for host in _HOSTS]


# -- 43c8c17b: ReDim element types by identity, bounds a Long cannot hold ----

_CHANGES = [
    *(
        case_id
        for name in ("r_project_collection_to_vba", "r_vba_collection_to_project", "r_long_to_integer")
        for case_id in _each_host(name, "compile")
    ),
    "issue1412_r_project_range_to_excel_compile",
    "issue1412_word_r_project_range_to_word_compile",
]


@pytest.mark.parametrize("case_id", _CHANGES)
def test_redim_changing_the_named_type_is_reported(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "rejected"
    assert len([f for f in _errors_for(case) if f.startswith("redim-type-change:")]) == 1


_SAME = [
    *(
        case_id
        for name in ("r_project_collection_same", "r_collection_to_vba_no_class")
        for case_id in _each_host(name, "compile")
    ),
    "issue1412_r_range_to_excel_no_class_compile",
    "issue1412_word_r_range_to_word_no_class_compile",
]


@pytest.mark.parametrize("case_id", _SAME)
def test_redim_naming_the_same_type_is_quiet(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "accepted"
    assert _errors_for(case) == []


_MALFORMED_LITERALS = [
    "l_dim_explicit_lower_long_overflow", "l_dim_upper_long_overflow", "l_redim_long_overflow",
    "l_redim_explicit_lower_long_overflow", "l_dim_integer_suffix_overflow",
]


@pytest.mark.parametrize("case_id", [c for name in _MALFORMED_LITERALS for c in _each_host(name, "compile")])
def test_suffixed_bound_out_of_range_is_only_the_literal(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "rejected"
    assert _codes_of(_errors_for(case)) == ["suffixed-literal-overflow"]


_OVERFLOWS = [
    "l_redim_double_overflow", "l_redim_unsuffixed_overflow",
    "l_redim_explicit_upper_double_overflow", "l_redim_negative_lower_double_overflow",
]


@pytest.mark.parametrize("case_id", [c for name in _OVERFLOWS for c in _each_host(name, "runtime")])
def test_redim_bound_a_long_cannot_hold_overflows(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "rejected"
    assert _codes_of(_errors_for(case)) == ["arithmetic-overflow"]


@pytest.mark.parametrize(
    "case_id",
    [
        *_each_host("l_dim_long_max", "compile"),
        *(c for name in ("l_redim_double_in_range", "b_element_bound", "b_long_bound") for c in _each_host(name, "runtime")),
    ],
)
def test_bounds_a_long_holds_are_quiet(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "accepted"
    assert _errors_for(case) == []


# -- b7fcb4d6: expression shapes judged the same in every statement ----------

# The run-time 91 of an unset Object is no syntax finding.
_UNSET = frozenset({"object-variable-not-set"})


def _ids1415(*names: str) -> list[str]:
    return [f"issue1415_{host}{name}_compile" for name in names for host in _HOSTS]


def _expect_only(case_id: str, expected: str, codes: list[str]) -> None:
    case = _case(case_id)
    assert case["expected"] == expected
    assert _codes_of(_errors_for(case, _UNSET)) == codes


_OPERATORS = [
    "plus", "minus", "times", "divide", "intdiv", "mod", "concat", "eq", "ne", "lt", "gt", "le", "ge",
    "and", "or", "xor", "eqv", "imp", "like",
]


@pytest.mark.parametrize(
    "case_id",
    _ids1415(
        *(f"g1_paren_array_{op}" for op in _OPERATORS),
        *(f"g1_double_paren_array_{op}" for op in ("plus", "concat", "eq", "and")),
        "l_bare_array_like", "l_like_pattern_array",
    ),
)
def test_parenthesised_array_is_still_an_array(case_id: str) -> None:
    _expect_only(case_id, "rejected", ["non-scalar-binary-operand"])


@pytest.mark.parametrize(
    "case_id",
    _ids1415(
        "g1_paren_array_pow", "g1_paren_element_plus_ok", "g1_paren_variant_plus_ok", "l_element_like_ok",
        "g2_is_paren_object_ok", "g2_is_collection_element_ok", "g3_typeof_ok", "g3_typeof_call_ok",
        "g4_udt_scalar_field_and_ok",
        "g5_case_element_ok", "g5_print_element_ok", "g5_seek_scalar_ok", "g5_lock_range_ok", "g5_close_scalar_ok",
        "g6_inline_else_ok", "g6_with_bang_ok", "h_call_commas_ok", "h_index_commas_ok", "h_unary_minus_ok",
    ),
)
def test_expression_shapes_that_compile_are_quiet(case_id: str) -> None:
    _expect_only(case_id, "accepted", [])


@pytest.mark.parametrize(
    "case_id",
    _ids1415(
        "g2_is_paren_left", "g2_is_paren_right", "g2_is_double_paren", "g2_is_element_left", "g2_is_element_right",
        "g5_print_number_is", "g5_write_number_is",
    ),
)
def test_scalar_is_operand_in_parens_or_element(case_id: str) -> None:
    _expect_only(case_id, "rejected", ["is-operator-non-object"])


@pytest.mark.parametrize(
    "case_id",
    _ids1415(
        "g3_typeof_paren", "g3_typeof_is_long", "g3_typeof_is_string", "g3_typeof_is_variant",
        "g3_typeof_call_is_long",
        "g6_with_bang_comma", "g6_with_dot_comma", "h_bare_comma", "h_member_comma", "h_let_comma",
        "h_set_comma", "h_inline_else_comma",
    ),
)
def test_malformed_typeof_and_assignment_commas_are_syntax_errors(case_id: str) -> None:
    _expect_only(case_id, "rejected", ["invalid-expression-syntax"])


@pytest.mark.parametrize(
    "case_id",
    _ids1415(
        *(f"g4_udt_field_{op}" for op in ("and", "or", "xor", "eqv", "imp", "plus", "like")),
        "g5_case_array", "g5_case_is_array", "g5_case_is_array_plus", "g5_print_array_plus",
        "g5_print_tab_array", "g5_print_spc_array", "g5_print_number_array", "g5_write_number_array",
        "g5_seek_array", "g5_put_record_array", "g5_get_record_array", "g5_line_input_number_array",
        "g5_open_number_array", "g5_close_number_array", "g5_lock_record_array", "g5_unlock_record_array",
        "g5_lock_number_array", "g5_unlock_number_array", "g5_lock_range_end_array",
    ),
)
def test_statement_operand_positions_refuse_an_array(case_id: str) -> None:
    _expect_only(case_id, "rejected", ["non-scalar-binary-operand"])


@pytest.mark.parametrize(
    "case_id", _ids1415("g6_inline_else_dangling_plus", "h_trailing_plus", "h_inline_trailing_plus")
)
def test_value_ending_at_an_operator_stays_quiet_while_typed(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "rejected"
    assert "invalid-expression-syntax" not in _codes_of(_errors_for(case, _UNSET))


def _non_scalar(source: str) -> int:
    body = f"Sub T()\n{source}\nEnd Sub\n"
    return len([d for d in analyze_module(body) if d.code == "non-scalar-binary-operand"])


def test_like_and_parentheses_judge_the_array() -> None:
    assert _non_scalar('    Dim a(3) As Long\n    Dim b As Boolean\n    b = (a Like "x")') == 1
    assert _non_scalar('    Dim a(3) As String\n    Dim b As Boolean\n    b = "x" Like a') == 1
    assert _non_scalar('    Dim a(3) As Long\n    Dim s As String\n    s = "x" & (a)') == 1
    assert _non_scalar('    Dim a(3) As Long\n    Dim s As String\n    s = "x" & ((a))') == 1
    assert _non_scalar("    Dim a(3) As Long\n    Dim v As Variant\n    v = (a(0)) + 1") == 0
    assert _non_scalar("    Dim a(3) As Long\n    Dim v As Variant\n    v = (a) ^ 1") == 0
    assert _non_scalar("    Dim a(3) As Long\n    If (a Is Nothing) Then\n    End If") == 0


# -- 0abb4c64: For bounds, subscripted Next counters and nested inline Ifs ----


def _ids1417(*names: str) -> list[str]:
    return [f"issue1417_{host}{name}_compile" for name in names for host in _HOSTS]


def _expect_1417(case_id: str, expected: str, codes: list[str]) -> None:
    case = _case(case_id)
    assert case["expected"] == expected
    assert _codes_of(_errors_for(case)) == codes


@pytest.mark.parametrize(
    "case_id",
    _ids1417(
        "n_next_index_long", "n_next_index_integer", "n_next_index_byte", "n_next_index_date",
        "n_next_index_single", "n_next_index_double", "n_next_index_currency", "n_next_index_longlong",
        "n_next_index_longptr", "n_next_index_long_suffix", "n_next_index_integer_suffix",
    ),
)
def test_subscripted_next_counter_is_expected_array(case_id: str) -> None:
    _expect_1417(case_id, "rejected", ["scalar-indexed"])


@pytest.mark.parametrize("case_id", _ids1417("n_next_index_variant"))
def test_subscripted_variant_next_counter_is_invalid_next(case_id: str) -> None:
    _expect_1417(case_id, "rejected", ["next-variable-mismatch"])


@pytest.mark.parametrize("case_id", _ids1417("n_next_index_boolean"))
def test_subscripted_boolean_next_counter_is_the_counter_type_first(case_id: str) -> None:
    _expect_1417(case_id, "rejected", ["for-counter-type"])


@pytest.mark.parametrize(
    "case_id",
    _ids1417(
        "n_next_ok", "n_next_suffix_ok",
        "f_lineno_ok", "f2_variant_holding_ok", "f2_element_ok", "f2_ubound_ok",
        "e_else_if_ok", "e2_nested_else_if_ok",
    ),
)
def test_for_and_next_shapes_that_compile_are_quiet(case_id: str) -> None:
    _expect_1417(case_id, "accepted", [])


def test_subscripted_next_inside_if_block_reports_once() -> None:
    source = "\r\n".join(
        ["Option Explicit", "Sub Main()", "    Dim i As Long", "    If True Then", "        For i = 1 To 2",
         "        Next i(1)", "    End If", "End Sub", ""]
    )
    case = {"modules": [{"name": "Module1", "type": "standard", "source": source}]}
    assert _codes_of(_errors_for(case)) == ["scalar-indexed"]


@pytest.mark.parametrize(
    "case_id",
    _ids1417(
        "f_lineno_array_start", "f_lineno_array_end", "f_lineno_array_step", "f_array_start",
        "f2_dynamic_array_end", "f2_variant_array_start", "f2_variant_counter_end", "f2_paren_array_end",
        "f2_string_array_end",
    ),
)
def test_array_as_for_bound_is_type_mismatch(case_id: str) -> None:
    _expect_1417(case_id, "rejected", ["for-bound-array"])


@pytest.mark.parametrize("case_id", _ids1417("e_else_if_is", "e_if_is", "e2_then_if_is"))
def test_is_in_nested_single_line_if_is_judged(case_id: str) -> None:
    _expect_1417(case_id, "rejected", ["is-operator-non-object"])


# -- 8373ca74: Is and TypeOf judge VBA's own function results ------------------

_CASES_1416 = [case for case_id, case in _ORACLE.items() if case_id.startswith("issue1416_")]

# The analyzer's code for each compile error the VBE gave.
_CODE_FOR_MESSAGE = {
    "Type mismatch": "is-operator-non-object",
    "Syntax error": "invalid-expression-syntax",
    "Method or data member not found": "member-not-found",
}

# Refused by the VBE and not reported: `DoEvents Is Nothing`, since the
# expression parser does not read a keyword-named function without parentheses
# as a call, and an Enum member, `ShadeOne Is Nothing`, since the rule does not
# resolve constants.
_KNOWN_QUIET_1416 = re.compile(r"^issue1416_(word_|powerpoint_|access_)?(is_doevents|enum3_member_is)_compile$")
_MESSAGE_1416 = re.compile(r" is rejected: Compile error: (.+?)\. Measured in ")


def test_function_result_cases_cover_every_host() -> None:
    assert len(_CASES_1416) == 3088
    for host in (None, "word", "powerpoint", "access"):
        assert len([case for case in _CASES_1416 if case.get("host") == host]) == 772


@pytest.mark.parametrize(
    "case_id", [case["id"] for case in _CASES_1416 if not _KNOWN_QUIET_1416.match(case["id"])]
)
def test_function_result_case_gets_the_vbe_error(case_id: str) -> None:
    case = _case(case_id)
    found = _MESSAGE_1416.search(case["description"])
    expected = [] if case["expected"] == "accepted" else [_CODE_FOR_MESSAGE.get(found.group(1) if found else "")]
    assert None not in expected
    assert _codes_of(_errors_for(case)) == expected


def test_function_result_tables_stay_apart() -> None:
    assert not ARGUMENT_TYPED_FUNCTIONS & VALUE_RESULT_FUNCTIONS
    assert not BARE_VALUE_RESULT_FUNCTIONS & VALUE_RESULT_FUNCTIONS
    assert "cvar" in TYPEOF_KEYWORD_OPERANDS
    assert "cvar" not in VALUE_RESULT_FUNCTIONS


def _inline_errors(*body: str) -> list[str]:
    source = "\r\n".join(["Option Explicit", *body, ""])
    return _codes_of(_errors_for({"modules": [{"name": "Module1", "type": "standard", "source": source}]}))


@pytest.mark.parametrize(
    "body",
    [
        ("Function Sqr(x) As Object", "End Function", "Sub Main()", "    Debug.Print Sqr(1) Is Nothing", "End Sub"),
        ("Sub Main()", "    Dim VBA As Variant", '    Debug.Print VBA.Mid$("ab", 1) Is Nothing', "End Sub"),
        ("Sub Main()", "    Dim v As Variant", "    Debug.Print Abs(v) Is Nothing", "End Sub"),
        ("Sub Main()", '    Debug.Print Right("ab", 1) Is Nothing', "End Sub"),
    ],
)
def test_shadows_and_variants_leave_function_results_alone(body: tuple[str, ...]) -> None:
    assert _inline_errors(*body) == []


def test_deeply_nested_argument_typed_calls_do_not_overflow() -> None:
    # The argument walk runs on an explicit stack; nested calls stay bounded by
    # the expression parser's depth cap.
    call = "1"
    for _ in range(200):
        call = f"Abs({call})"
    assert isinstance(_inline_errors("Sub Main()", f"    Debug.Print {call} Is Nothing", "End Sub"), list)


# -- a58e427f: designer controls and attributes reach the analyzer ----------

_CASES_1419 = [case for case_id, case in _ORACLE.items() if case_id.startswith("issue1419_")]

_CODE_FOR_MESSAGE_1419 = {
    "Type mismatch": "is-operator-non-object",
    "Syntax error": "reserved-keyword-in-expression",
    "Method or data member not found": "member-not-found",
}

# `Len() Is Nothing` is the VBE's Syntax error, reported as argument-count.
# `Me.Len` on a form whose fixture does not state its control list stays
# quiet: without the list nothing proves Len absent. The same form with an
# empty designer, u_empty_designer_me_len, is reported.
_OTHER_CODES_1419 = {
    "u_len_control_4": ["argument-count"],
    "u_no_control_me_len": [],
}
_FAMILY_1419 = re.compile(r"^issue1419_(word_|powerpoint_|access_)?")


def test_designer_input_cases_cover_every_host() -> None:
    assert len(_CASES_1419) == 233
    assert {case.get("host", "excel") for case in _CASES_1419} == {"excel", "word", "powerpoint", "access"}


@pytest.mark.parametrize("case_id", [case["id"] for case in _CASES_1419])
def test_designer_input_case_gets_the_vbe_error(case_id: str) -> None:
    case = _case(case_id)
    family = _FAMILY_1419.sub("", case_id).removesuffix("_compile")
    found = _MESSAGE_1416.search(case["description"])
    expected = _OTHER_CODES_1419.get(family)
    if expected is None:
        expected = [] if case["expected"] == "accepted" else [_CODE_FOR_MESSAGE_1419.get(found.group(1) if found else "")]
    assert None not in expected
    assert _codes_of(_errors_for(case)) == expected


def test_fixture_adapter_types_controls_by_prog_id() -> None:
    picker = oracle_module_input(
        {"name": "Form1", "type": "userform", "source": "",
         "designerControls": [{"name": "Picker", "progId": "Forms.ComboBox.1"}]}
    )
    assert [(member.name, member.type) for member in picker.implicit_members] == [("Picker", "MSForms.ComboBox")]
    assert oracle_module_input({"name": "Form1", "type": "userform", "source": "", "designerControls": []}).implicit_members == []
    assert oracle_module_input({"name": "Form1", "type": "userform", "source": ""}).implicit_members is None
    sheet = oracle_module_input(
        {"name": "Sheet1", "type": "document", "source": "", "designerClass": "Excel.Worksheet", "implicitMembers": []}
    )
    assert (sheet.designer_class, sheet.implicit_members) == ("Excel.Worksheet", [])
    assert oracle_module_input({"name": "Timer", "type": "class", "source": "", "predeclaredId": True}).predeclared_id is True


# -- 7ec04b3f: a ReDim bound is an expression --------------------------------


def _ids1412(*names: str) -> list[str]:
    return [f"issue1412_{host}{name}_compile" for name in names for host in _HOSTS]


@pytest.mark.parametrize("case_id", _ids1412("rb_is_nothing", "rb_cdbl_is", "rb_typeof_n"))
def test_redim_bound_is_operand_is_judged(case_id: str) -> None:
    _expect_1417(case_id, "rejected", ["is-operator-non-object"])


@pytest.mark.parametrize(
    "case_id",
    _ids1412("rb_array_upper", "rb_array_lower", "rb_array_paren", "rb_array_preserve", "rb_array_second_dim"),
)
def test_array_as_redim_bound_is_type_mismatch(case_id: str) -> None:
    _expect_1417(case_id, "rejected", ["redim-bound-array"])


@pytest.mark.parametrize("case_id", _ids1412("rb_array_plus"))
def test_array_under_an_operator_in_redim_bound(case_id: str) -> None:
    _expect_1417(case_id, "rejected", ["non-scalar-binary-operand"])


@pytest.mark.parametrize(
    "case_id", _ids1412("rb_element_ok", "rb_ubound_ok", "rb_variant_ok", "rb_object_is_ok", "rb_two_dims_ok")
)
def test_redim_bounds_that_compile_are_quiet(case_id: str) -> None:
    _expect_1417(case_id, "accepted", [])


def _redim_main(*body: str) -> list[str]:
    source = "\r\n".join(
        ["Option Explicit", "Sub Main()", "    Dim a(1) As Long, d() As Long, n As Long", *body, "End Sub", ""]
    )
    return _codes_of(_errors_for({"modules": [{"name": "Module1", "type": "standard", "source": source}]}))


def test_redim_reads_each_variable_and_leaves_as_type_alone() -> None:
    assert _redim_main("    Dim e() As Long", "    ReDim d(1) As Long, e(a)") == ["redim-bound-array"]


def test_redim_in_if_arm_reports_through_operand_rules_only() -> None:
    assert _redim_main("    If n = 0 Then ReDim d(n Is Nothing)") == ["is-operator-non-object"]


# -- b931a850: an array bound a Long cannot hold, read past whole numbers ----

# A bound is made whole half to even, then must fit a Long: 2147483647.5 is
# 2147483648 and overflows, 2147483647.4 and -2147483648.5 do not. A Double
# Const, and an expression of Doubles, is folded with VBA's typing.

_DIM_OVERFLOWS = [
    "l_dim_double_overflow", "d_half_up", "d_low_past", "d_low_double", "d_const_double", "d_local_const_double",
    "d_const_untyped", "d_const_plus", "d_double_plus", "d_divide", "d_single", "d_currency", "d_second_dim",
    "d_module_level", "d_type_field", "d_long_sum",
]


@pytest.mark.parametrize("case_id", [c for name in _DIM_OVERFLOWS for c in _each_host(name, "compile")])
def test_dim_bound_a_long_cannot_hold_is_array_bound_overflow(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "rejected"
    assert _codes_of(_errors_for(case)) == ["array-bound-overflow"]


_REDIM_OVERFLOWS = [
    "l_redim_const_overflow", "r_half_up", "r_low_past", "r_local_const", "r_const_plus", "r_double_plus",
    "r_divide", "r_single", "r_currency",
]


@pytest.mark.parametrize("case_id", [c for name in _REDIM_OVERFLOWS for c in _each_host(name, "runtime")])
def test_folded_redim_bound_a_long_cannot_hold_overflows(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "rejected"
    assert _codes_of(_errors_for(case)) == ["arithmetic-overflow"]


@pytest.mark.parametrize(
    "case_id",
    [
        *(c for name in ("d_under_half_ok", "d_low_half_ok", "d_const_ok", "d_divide_ok") for c in _each_host(name, "compile")),
        *(c for name in ("r_under_half_ok", "r_low_half_ok", "r_const_ok") for c in _each_host(name, "runtime")),
    ],
)
def test_folded_bounds_a_long_holds_are_quiet(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "accepted"
    assert _errors_for(case) == []


def _module_of(*lines: str) -> dict[str, Any]:
    source = "\r\n".join(["Option Explicit", *lines, ""])
    return {"modules": [{"name": "Module1", "type": "standard", "source": source}]}


def test_array_bound_overflow_names_the_bound_and_underlines_the_dimension() -> None:
    case = _module_of("Const N As Double = 2147483647#", "Sub Main()", "    Dim a(1, 0 To N + 1) As Long", "End Sub")
    assert _errors_for(case) == ["array-bound-overflow: 0 To N + 1"]
    source = case["modules"][0]["source"]
    [message] = [d.message for d in analyze_module(source) if d.code == "array-bound-overflow"]
    assert message == (
        "Array 'a' bound 2147483648 for dimension 2 is outside the range of a Long; "
        "this is a VBE compile error: Overflow."
    )


def test_bound_past_the_range_is_not_also_reversed_bounds() -> None:
    case = _module_of("Sub Main()", "    Dim a(2147483648# To 1) As Long", "End Sub")
    assert _errors_for(case) == ["array-bound-overflow: 2147483648# To 1"]


def test_suffixed_dim_literal_out_of_range_is_the_literal_only() -> None:
    case = _module_of("Sub Main()", "    Dim a(2147483648&) As Long", "End Sub")
    assert _codes_of(_errors_for(case)) == ["suffixed-literal-overflow"]


def test_bound_reads_another_modules_const_and_a_local_hides_it() -> None:
    shared = {"name": "Limits", "type": "standard", "source": "Option Explicit\r\nPublic Const BIG As Double = 3E+09\r\n"}
    user = {
        "name": "Module1", "type": "standard",
        "source": "Option Explicit\r\nSub Main()\r\n    Dim a(BIG) As Long\r\nEnd Sub\r\n",
    }
    assert _errors_for({"modules": [shared, user]}) == ["array-bound-overflow: BIG"]
    hidden = _module_of("Const BIG As Double = 3E+09", "Sub Main()", "    Dim BIG As Long", "    Dim a(BIG) As Long", "End Sub")
    assert "array-bound-overflow" not in _codes_of(_errors_for(hidden))
