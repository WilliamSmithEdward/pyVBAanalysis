"""Rule contracts from the XLIDE 11.1.0 -> 11.2.1 sync (operands, bounds, Is).

Each section names the upstream commit whose own tests it follows. The oracle
cases are read from the vendored vbe_oracle_cases.json and analyzed as the
extension analyzes the case's host.
"""

from __future__ import annotations

import json
from typing import Any

import pytest

from pyvbaanalysis import analyze_module, analyze_module_options_for
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics.model import DiagnosticSeverity
from pyvbaanalysis.evidence import DATA_DIR
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex

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
        index.set_module(ModuleInput(module["name"], _KIND.get(module.get("type", "standard"), ModuleSymbolKind.STANDARD), module["source"]))
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
