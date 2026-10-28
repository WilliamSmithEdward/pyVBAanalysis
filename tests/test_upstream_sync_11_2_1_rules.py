"""Rule contracts from the XLIDE 11.1.0 -> 11.2.1 sync (rule fixes).

Each section names the upstream commit whose own tests it follows.
"""

import pytest

from pyvbaanalysis import analyze_module, analyze_module_options_for, analyze_project
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex
from oracle_support import CASES, case_codes  # type: ignore[attr-defined]


def _project(source: str, *others: tuple[str, ModuleSymbolKind, str], name: str = "Caller") -> list:
    modules = [ModuleInput(name, ModuleSymbolKind.STANDARD, source)]
    modules += [ModuleInput(other, kind, text) for other, kind, text in others]
    return analyze_project(modules)[name]


def _hits(source: str, code: str, options: AnalyzeModuleOptions | None = None) -> list:
    return [d for d in analyze_module(source, options) if d.code == code]


def _starts(source: str, code: str, options: AnalyzeModuleOptions | None = None) -> list[int]:
    return [d.span.start for d in _hits(source, code, options)]


# -- 53e6868d: the Else of a one-line If after a colon -----------------------

_ELSE_TAIL_QUIET = [
    "Sub T()\n    If a Then b = 1: Else c = 2\nEnd Sub\n",
    "Sub T()\n    If a Then b = 1: Else: c = 2\nEnd Sub\n",
    "Sub T()\n    If a Then b = 1: _\n        Else c = 2\nEnd Sub\n",
    "Sub T()\n    If a Then For i = 1 To 2: Next: Else x = 1\n"
    "    If a Then With c: End With: Else d = 1\nEnd Sub\n",
    "Sub T()\n    If a Then Select Case b: Case 1: c = 1: End Select: Else d = 1\nEnd Sub\n",
    "Sub T()\n    If a Then b = 1: Else If x Then d = 2: Else e = 3\nEnd Sub\n",
    "Sub T()\nL1: If a Then b = 1: Else c = 2\n10 If a Then b = 1: Else c = 2\nEnd Sub\n",
    "Sub T()\n#If True Then\n    If a Then b = 1: Else c = 2\n#Else\n    Else\n#End If\nEnd Sub\n",
    "Sub T()\n    If a Then If b Then c = 1 Else d = 2: Else e = 3\nEnd Sub\n",
    "Sub T()\n    If a Then If b Then c = 1: Else d = 2: Else e = 3\nEnd Sub\n",
    "Sub T()\n    If a Then b = 1: If c Then d = 2: Else e = 3\nEnd Sub\n",
    "Sub T()\n    If a Then obj.Else = 1: Else b\nEnd Sub\n",
]


@pytest.mark.parametrize("source", _ELSE_TAIL_QUIET)
def test_else_tail_after_colon_is_accepted(source: str) -> None:
    assert _hits(source, "else-without-if") == []


@pytest.mark.parametrize("source", [
    "Sub T()\n    If x Then\n        If a Then b = 1: Else c = 2\n    Else\n        y = 1\n    End If\nEnd Sub\n",
    "Sub T()\n    If x Then\n        If a Then If b Then c = 1 Else d = 2: Else e = 3\n    End If\nEnd Sub\n",
    "Sub T()\n    If x Then\n        If a Then b Else c: If d Then e: Else f\n    End If\nEnd Sub\n",
])
def test_else_tail_inside_block_if_is_not_a_clause(source: str) -> None:
    assert _hits(source, "else-without-if") == []
    assert _hits(source, "else-branch-order") == []


@pytest.mark.parametrize("source, marker", [
    ("Sub T()\n    If a Then b = 1: Else c = 2: Else d = 3\nEnd Sub\n", "Else d"),
    ("Sub T()\n    If a Then b = 1 Else c = 2: Else d = 3\nEnd Sub\n", "Else d"),
    ("Sub T()\n    If x Then\n        If a Then b = 1: Else c = 2: Else d = 3\n"
     "    Else\n        y = 1\n    End If\nEnd Sub\n", "Else d"),
    ("Sub T()\n    Else\nEnd Sub\n", "Else"),
    ("Sub T()\n    If a Then b = 1:\n    Else c = 2\nEnd Sub\n", "Else"),
    ("Sub T()\n    b = 1: Else c = 2\nEnd Sub\n", "Else"),
    ("Sub T()\n    If a Then b = 1: ElseIf c Then d = 2\nEnd Sub\n", "ElseIf"),
    ("Sub T()\n    If x Then\n        If a Then b = 1: ElseIf c Then d = 2\n    End If\nEnd Sub\n", "ElseIf"),
    ("Sub T()\n    If a Then b Else c: End If: Else d\nEnd Sub\n", "Else d"),
    ("Sub T()\n    If a Then obj.If = 1: Else b: Else c\nEnd Sub\n", "Else c"),
    ("Sub T()\n    If a Then b: Else c: Else d: If e Then f: Else g\nEnd Sub\n", "Else d"),
])
def test_stray_else_after_colon_is_still_reported(source: str, marker: str) -> None:
    assert _starts(source, "else-without-if") == [source.index(marker)]
    assert _hits(source, "else-branch-order") == []


def test_block_if_with_two_else_clauses_still_reported() -> None:
    source = "Sub T()\n    If x Then\n        a = 1\n    Else\n        b = 2\n    Else\n        c = 3\n    End If\nEnd Sub\n"
    assert len(_hits(source, "else-branch-order")) == 1


# -- a2321dcf: runtime readonly properties, host scalar array assignments ----

_FLAGS_FN = "Function MakeFlags() As Boolean()\nDim flags(1) As Boolean\nMakeFlags = flags\nEnd Function\n"


def _runtime_source(line: str) -> str:
    return (
        "Option Explicit\n" + _FLAGS_FN
        + "Function Main() As Variant\nDim values(1) As Boolean\n" + line + "\nMain = 1\nEnd Function"
    )


@pytest.mark.parametrize("target", ["Err.LastDllError", "VBA.Err.LastDllError"])
def test_runtime_readonly_property_is_reported(target: str) -> None:
    source = _runtime_source(target + " = 5")
    for findings in (analyze_module(source), _project(source)):
        found = [d for d in findings if d.code == "readonly-member-assignment"]
        assert len(found) == 1
        assert "Can't assign to read-only property" in found[0].message
        assert source[found[0].span.start:found[0].span.end] == "LastDllError"


@pytest.mark.parametrize("value", ["values", "[values]", "(values)", "MakeFlags()", "(MakeFlags())"])
def test_whole_array_into_host_scalar_setter(value: str) -> None:
    source = _runtime_source("ThisWorkbook.Worksheets(1).EnableCalculation = " + value)
    for findings in (analyze_module(source), _project(source)):
        errors = [d for d in findings if d.code == "assignment-type-mismatch"]
        assert len(errors) == 1
        assert "Run-time error '13'" in errors[0].message
        assert "compile error" not in errors[0].message


@pytest.mark.parametrize("line", [
    "Err.Number = 5", "VBA.Err.Number = 5", "Main = Err.LastDllError",
    "ThisWorkbook.Worksheets(1).EnableCalculation = True",
    "ThisWorkbook.Worksheets(1).EnableCalculation = values(0)",
    'ThisWorkbook.Worksheets(1).EnableCalculation = "True"',
])
def test_valid_runtime_and_host_scalar_use_is_quiet(line: str) -> None:
    assert [d for d in analyze_module(_runtime_source(line)) if d.severity == "error"] == []


@pytest.mark.parametrize("receiver", ["ThisWorkbook.Sheets(1)", "unknown"])
def test_ambiguous_receiver_assumes_no_scalar_setter(receiver: str) -> None:
    source = _runtime_source(receiver + ".EnableCalculation = values").replace(
        "Dim values(1) As Boolean", "Dim values(1) As Boolean\nDim unknown As Object"
    )
    assert _hits(source, "assignment-type-mismatch") == []


def test_shadowed_err_is_not_the_runtime_object() -> None:
    assert _hits("Sub T(ByVal Err As Object)\nErr.LastDllError = 5\nEnd Sub", "readonly-member-assignment") == []


@pytest.mark.parametrize("line", [
    "With ThisWorkbook.Worksheets(1)\n.EnableCalculation = values\nEnd With",
    "If True Then ThisWorkbook.Worksheets(1).EnableCalculation = MakeFlags()",
])
def test_host_scalar_setter_in_with_and_single_line_if(line: str) -> None:
    assert len(_hits(_runtime_source(line), "assignment-type-mismatch")) == 1


# -- e937c742, 31ff586f: reserved words as Type member names -----------------

_RESERVED_MEMBER_WORDS = [
    "Next", "Loop", "Wend", "For", "Do", "With", "Sub", "Function", "Select", "If",
    "End", "Else", "Case", "Set", "Dim", "Private", "Len", "True", "Date", "String",
    "Long", "Variant", "Local", "Attribute", "New",
]


def _reserved_member_type(words: list[str]) -> str:
    return (
        "Private Type Rec\n"
        + "".join(f"    {word} As Long\n" for word in words)
        + "    While _\n        As Long\n"
        + "    Loop As Long ' comment\n"
        + "    Next As [Long]\n"
        + "    Next As New Collection\n"
        + "    With As New [Collection]\n"
        + "#If Win64 Then\n    Next As LongLong\n#Else\n    Next As Long\n#End If\n"
        + "    Type As [Long]\n"
        + "    Type(1 To 2) As Long\n"
        + "End Type\n"
    )


def test_reserved_word_type_member_with_as_clause_is_accepted() -> None:
    findings = analyze_module(_reserved_member_type(_RESERVED_MEMBER_WORDS + ["As", "Implements"]))
    for code in (
        "invalid-declaration-name", "unexpected-declaration-token", "invalid-as-type-name",
        "type-member-without-type",
    ):
        assert [d for d in findings if d.code == code] == []


def test_reserved_array_suffixed_and_incomplete_type_members_still_rejected() -> None:
    source = (
        "Private Type Rec\n    For(1 To 3) As Long\n    With As\n    Loop As ' comment\n"
        "    Next&\n    Next& As Long\n    Set As New\n    Set As New []\n    Set As []\n"
        "    Set As 123\n    Me As Long\n    Next\nEnd Type\n"
    )
    hits = _hits(source, "invalid-declaration-name")
    assert [source[d.span.start:d.span.end] for d in hits] == [
        "For", "With", "Loop", "Next", "Next", "Set", "Set", "Set", "Set", "Me", "Next",
    ]
    enum_source = "Private Enum Mode\n    Next\n    Loop = 2\nEnd Enum\n"
    hits = _hits(enum_source, "invalid-declaration-name")
    assert [enum_source[d.span.start:d.span.end] for d in hits] == ["Next", "Loop"]


# -- 31ff586f, 90a0641f: Type members named As or Implements -----------------

_AS_E = "As" + chr(0xE9)
_AS_BETA = "As" + chr(0x3B2) + "eta"
_IMPLEMENTS_SOURCE = (
    "Implements As_Foo\nImplements Asx\n"
    f"Implements {_AS_E}\nImplements {_AS_BETA}\n"
    "Private Type Rec\n    As As Long\n    Implements As Long\nEnd Type\n"
)


@pytest.mark.parametrize("line, name", [
    ("Implements As_Foo", "As_Foo"),
    ("  implements Asx", "Asx"),
    (f"Implements {_AS_E}", _AS_E),
    (f"Implements {_AS_BETA}", _AS_BETA),
    ("Implements Lib.IShape", "Lib.IShape"),
    ("Implements Lib.", "Lib"),
    ("Implements As Long", None),
    ("Implements As", None),
    ("Implements AS", None),
    ("Implements 1x", None),
])
def test_implements_statement_name(line: str, name: str | None) -> None:
    from pyvbaanalysis.lexer.token_helpers import implements_statement_name

    assert implements_statement_name(line) == name


def test_type_members_named_as_and_implements_implement_nothing() -> None:
    from pyvbaanalysis.symbols import ProjectIndex

    index = ProjectIndex()
    index.set_module(ModuleInput("Widget", ModuleSymbolKind.CLASS, _IMPLEMENTS_SOURCE))
    assert index.implemented_interface_names() == frozenset(
        name.lower() for name in ("As_Foo", "Asx", _AS_E, _AS_BETA)
    )


# -- b7c597c3: a procedure closer in each arm of an #If ----------------------

@pytest.mark.parametrize("source", [
    "Option Explicit\n#Const FLAG = 0\nSub T()\n#If FLAG Then\nEnd Sub\n#Else\nEnd Sub\n#End If\n",
    "Option Explicit\n#Const FLAG = 0\nFunction Main() As Variant\n    Main = 5\n"
    "#If FLAG Then\nEnd Function\n#Else\nEnd Function\n#End If\n",
    "Option Explicit\n#Const FLAG = 1\nFunction Main() As Variant\n    Main = 5\n"
    "#If FLAG Then\nEnd Function\n#Else\nEnd Function\n#End If\n",
])
def test_procedure_closer_in_each_arm_of_an_if(source: str) -> None:
    assert [d for d in analyze_module(source) if d.severity == "error"] == []


def test_other_statement_in_the_other_arm_is_still_outside_a_procedure() -> None:
    source = "Option Explicit\n#Const FLAG = 0\nSub T()\n#If FLAG Then\nEnd Sub\n#Else\nx = 1\n#End If\n"
    assert len(_hits(source, "statement-outside-procedure")) == 1


# -- a102b93f: false positives Excel 16.0 compiles or runs -------------------
# Each case is in the vendored vbe_oracle_cases.json (upstream's
# excelMeasuredFalsePositives.test.ts). The structural block-balance codes
# upstream also names are not the port's (it has no structural pass).

_A102_COMPILED_OR_RAN = [
    ("print_udt_len_argument_runtime", ("udt-value-mismatch",)),
    ("print_udt_lenb_argument_runtime", ("udt-value-mismatch",)),
    ("print_udt_len_second_argument_runtime", ("udt-value-mismatch",)),
    ("print_udt_user_function_argument_runtime", ("udt-value-mismatch",)),
    ("print_udt_varptr_argument_runtime", ("udt-value-mismatch",)),
    ("print_udt_len_plus_runtime", ("udt-value-mismatch",)),
    ("print_udt_grouped_len_runtime", ("udt-value-mismatch",)),
    ("print_udt_len_of_grouped_runtime", ("udt-value-mismatch",)),
    ("optional_collection_default_number_runtime", ("parameter-default-type-mismatch",)),
    ("optional_object_default_number_runtime", ("parameter-default-type-mismatch",)),
    ("optional_object_default_string_runtime", ("parameter-default-type-mismatch",)),
    ("optional_range_default_zero_runtime", ("parameter-default-type-mismatch",)),
    ("optional_collection_default_true_runtime", ("parameter-default-type-mismatch",)),
    ("optional_collection_default_date_runtime", ("parameter-default-type-mismatch",)),
    ("optional_collection_default_null_runtime", ("parameter-default-type-mismatch",)),
    ("optional_class_default_number_runtime", ("parameter-default-type-mismatch",)),
    ("type_named_range_shadows_excel_runtime", ("object-variable-not-set",)),
    ("type_named_worksheet_shadows_excel_runtime", ("object-variable-not-set",)),
    ("public_type_named_range_other_module_runtime", ("object-variable-not-set",)),
    ("enum_named_range_shadows_excel_runtime", ("object-variable-not-set",)),
    ("thisworkbook_type_set_activeworkbook_runtime", ("assignment-object-type-mismatch",)),
    ("thisworkbook_type_member_runtime", ("assignment-object-type-mismatch",)),
    ("thisworkbook_type_parameter_runtime", ("argument-type-mismatch",)),
    ("dim_shared_local_runtime", ("invalid-declaration-name", "undeclared-variable", "unexpected-declaration-token",)),
    ("dim_shared_two_names_runtime", ("invalid-declaration-name", "unexpected-declaration-token",)),
    ("dim_shared_module_runtime", ("invalid-declaration-name", "undeclared-variable", "unexpected-declaration-token",)),
    ("private_shared_module_runtime", ("invalid-declaration-name", "unexpected-declaration-token",)),
    ("dim_shared_array_runtime", ("invalid-declaration-name", "undeclared-variable", "unexpected-declaration-token",)),
    ("dim_shared_new_runtime", ("invalid-declaration-name", "undeclared-variable", "unexpected-declaration-token",)),
    ("public_shared_module_runtime", ("invalid-declaration-name", "undeclared-variable", "unexpected-declaration-token",)),
    ("global_shared_module_runtime", ("invalid-declaration-name", "undeclared-variable", "unexpected-declaration-token",)),
    ("dim_shared_dynamic_array_runtime", ("invalid-declaration-name", "unexpected-declaration-token",)),
    ("erase_two_arrays_empty_parentheses_runtime", ("unallocated-dynamic-array-access",)),
    ("label_dollar_goto_runtime", ("undeclared-variable", "undefined-label",)),
    ("label_dollar_goto_dollar_runtime", ("undeclared-variable",)),
    ("label_percent_runtime", ("undeclared-variable",)),
    ("label_hash_runtime", ("undeclared-variable",)),
    ("label_dollar_then_long_variable_runtime", ("type-suffix-mismatch", "undeclared-variable",)),
    ("label_ampersand_runtime", ("undeclared-variable", "undefined-label",)),
    ("label_bang_runtime", ("undeclared-variable", "undefined-label",)),
    ("label_at_runtime", ("undeclared-variable", "undefined-label",)),
    ("label_caret_runtime", ("invalid-expression-syntax", "undeclared-variable", "undefined-label",)),
    ("label_dollar_then_statement_runtime", ("undeclared-variable",)),
    ("label_dollar_resume_runtime", ("undeclared-variable", "undefined-label",)),
    ("resume_double_zero_runtime", ("undefined-label",)),
    ("resume_zero_handler_runtime", ("undefined-label",)),
    ("err_unknown_member_assign_compile", ("member-not-found",)),
    ("err_unknown_member_read_compile", ("member-not-found",)),
    ("err_unknown_member_call_compile", ("member-not-found",)),
    ("me_let_document_module_compile", ("undeclared-variable",)),
    ("me_let_class_module_compile", ("undeclared-variable",)),
    ("me_let_keyword_class_compile", ("undeclared-variable",)),
    ("now_indexed_result_compile", ("argument-count",)),
    ("time_indexed_result_compile", ("argument-count",)),
    ("command_indexed_result_compile", ("argument-count",)),
    ("now_indexed_two_compile", ("argument-count",)),
    ("now_indexed_print_compile", ("argument-count",)),
    ("str_null_runtime", ("argument-type-mismatch",)),
    ("static_public_sub_runtime", ("unmatched-block-closer",)),
    ("static_friend_sub_runtime", ("unmatched-block-closer",)),
    ("byref_qualified_parenthesized_copy_runtime", ("byref-argument-type-mismatch",)),
    ("byref_qualified_parenthesized_byte_runtime", ("byref-argument-type-mismatch",)),
    ("byref_qualified_parenthesized_single_runtime", ("byref-argument-type-mismatch",)),
    ("byref_qualified_parenthesized_double_runtime", ("byref-argument-type-mismatch",)),
    ("byref_qualified_parenthesized_currency_runtime", ("byref-argument-type-mismatch",)),
    ("byref_qualified_parenthesized_string_runtime", ("byref-argument-type-mismatch",)),
    ("byref_qualified_parenthesized_date_runtime", ("byref-argument-type-mismatch",)),
    ("byref_qualified_parenthesized_boolean_runtime", ("byref-argument-type-mismatch",)),
    ("lset_string_number_runtime", ("string-arithmetic-coercion",)),
    ("rset_string_number_runtime", ("string-arithmetic-coercion",)),
    ("lset_string_long_variable_runtime", ("string-arithmetic-coercion",)),
    ("udt_power_compile", ("non-scalar-binary-operand", "udt-value-mismatch",)),
    ("udt_power_right_compile", ("non-scalar-binary-operand", "udt-value-mismatch",)),
    ("udt_not_runtime", ("udt-value-mismatch",)),
    ("array_power_runtime", ("non-scalar-binary-operand",)),
    ("udt_field_power_compile", ("non-scalar-binary-operand",)),
    ("label_line_number_and_dollar_runtime", ("undeclared-variable", "undefined-label",)),
    ("document_type_sheet_worksheet_runtime", ("assignment-object-type-mismatch",)),
    ("str_null_into_variant_runtime", ("assignment-type-mismatch",)),
    ("optional_object_default_string_expression_runtime", ("parameter-default-not-constant", "parameter-default-type-mismatch",)),
    ("line_number_then_label_caret_compile", ("invalid-expression-syntax", "undefined-label",)),
    ("line_number_then_label_hash_compile", ("invalid-expression-syntax", "undefined-label",)),
    ("label_caret_before_statement_runtime", ("invalid-expression-syntax", "undefined-label",)),
    ("label_at_before_statement_runtime", ("invalid-expression-syntax", "undefined-label",)),
    ("let_me_document_module_runtime", ("invalid-assignment-target",)),
    ("me_assign_document_module_runtime", ("invalid-assignment-target",)),
    ("project_time_variant_function_index_runtime", ("argument-count",)),
    ("static_private_function_compile", ("unmatched-block-closer",)),
    ("static_public_property_compile", ("unmatched-block-closer",)),
]

_A102_STILL_REPORTED = [
    ("print_udt_grouped_compile", ("udt-value-mismatch",)),
    ("print_udt_bare_compile", ("udt-value-mismatch",)),
    ("print_udt_as_array_index_compile", ("udt-value-mismatch",)),
    ("type_named_range_set_compile", ("set-requires-object",)),
    ("static_shared_compile", ("invalid-declaration-name", "unexpected-declaration-token",)),
    ("dim_shared_as_type_compile", ("invalid-declaration-name",)),
    ("goto_zero_compile", ("undefined-label",)),
    ("resume_one_undefined_compile", ("undefined-label",)),
    ("resume_zero_without_error_runtime", ("resume-without-error",)),
    ("err_raise_too_many_arguments_compile", ("argument-count",)),
    ("set_me_class_compile", ("invalid-assignment-target",)),
    ("set_me_document_module_compile", ("invalid-assignment-target",)),
    ("timer_with_argument_compile", ("argument-count",)),
    ("erl_with_argument_compile", ("argument-count",)),
    ("doevents_with_argument_compile", ("argument-count",)),
    ("str_dollar_null_runtime", ("argument-type-mismatch",)),
    ("oct_dollar_null_runtime", ("argument-type-mismatch",)),
    ("byref_qualified_no_parentheses_compile", ("byref-argument-type-mismatch",)),
    ("byref_qualified_call_compile", ("byref-argument-type-mismatch",)),
    ("byref_qualified_function_compile", ("byref-argument-type-mismatch",)),
    ("udt_plus_compile", ("non-scalar-binary-operand", "udt-value-mismatch",)),
    ("collection_power_compile", ("collection-operand",)),
    ("longlong_parenthesized_byref_long_compile", ("longlong-narrowing",)),
    ("longlong_qualified_parenthesized_byref_long_compile", ("longlong-narrowing",)),
    ("longlong_call_double_parentheses_compile", ("longlong-narrowing",)),
    ("longlong_named_parenthesized_compile", ("longlong-narrowing",)),
    ("longlong_qualified_byval_long_compile", ("longlong-narrowing",)),
    ("longlong_qualified_named_byval_long_compile", ("longlong-narrowing",)),
    ("longlong_named_byval_long_compile", ("longlong-narrowing",)),
    ("erase_empty_parentheses_long_compile", ("erase-requires-array",)),
    ("array_unary_minus_compile", ("non-scalar-binary-operand",)),
    ("document_type_sheet_workbook_runtime", ("assignment-object-type-mismatch",)),
    ("document_type_thisworkbook_application_runtime", ("assignment-object-type-mismatch",)),
    ("print_udt_array_index_in_call_compile", ("udt-value-mismatch",)),
    ("print_udt_array_index_single_line_if_compile", ("udt-value-mismatch",)),
    ("print_udt_field_array_index_compile", ("udt-value-mismatch",)),
    ("print_udt_collection_item_argument_compile", ("udt-value-mismatch",)),
    ("project_now_function_index_compile", ("argument-count",)),
    ("now_statement_parenthesized_compile", ("argument-count",)),
    ("time_dollar_argument_compile", ("argument-count",)),
    ("command_dollar_argument_compile", ("argument-count",)),
    ("str_null_into_string_runtime", ("assignment-type-mismatch",)),
    ("optional_object_default_array_call_compile", ("parameter-default-not-constant",)),
    ("set_me_standard_module_compile", ("me-outside-object-module",)),
    ("static_global_sub_header_compile", ("invalid-proc-header",)),
    ("print_udt_cstr_compile", ("udt-value-mismatch",)),
    ("byref_class_method_no_parentheses_compile", ("byref-argument-type-mismatch",)),
    ("private_shared_withevents_compile", ("unexpected-declaration-token",)),
]


@pytest.mark.parametrize("case_id, codes", _A102_COMPILED_OR_RAN)
def test_excel_measured_false_positive_is_quiet(case_id: str, codes: tuple[str, ...]) -> None:
    reported = case_codes(CASES[case_id])
    assert [code for code in codes if code in reported] == []


@pytest.mark.parametrize("case_id, codes", _A102_STILL_REPORTED)
def test_excel_measured_control_is_still_reported(case_id: str, codes: tuple[str, ...]) -> None:
    reported = case_codes(CASES[case_id])
    assert [code for code in codes if code not in reported] == []


def _single_line_if_errors(line: str) -> list[str]:
    source = (
        "Option Explicit\nType T\n    a As Long\nEnd Type\nSub Main()\n"
        f"    Dim t As T, v As Variant, a(1) As Long\n    {line}\nEnd Sub\n"
    )
    return [f"{d.code}@{d.span.start}" for d in analyze_module(source) if d.severity == "error"]


@pytest.mark.parametrize("line, count", [
    ("If True Then v = t + 2", 1),
    ("If True Then v = t + 2 Else v = t * 2", 2),
    ("If True Then Debug.Print t", 1),
    ("If True Then Debug.Print Len(a(t))", 1),
    ("If t + 1 Then v = 1", 1),
])
def test_single_line_if_reports_each_type_once(line: str, count: int) -> None:
    codes = _single_line_if_errors(line)
    assert len(codes) == count
    assert len(set(codes)) == len(codes)


# -- 58fd2c90: a default member that returns an object, read into a Variant ---

_DEFAULT_ITEMS_CLASS = (
    "Option Explicit\nPrivate mItems As New Collection\nPublic Sub Add(ByVal v As Variant)\n"
    "    mItems.Add v\nEnd Sub\nPublic Property Get Items() As Collection\n"
    "Attribute Items.VB_UserMemId = 0\n    Set Items = mItems\nEnd Property\n"
)


@pytest.mark.parametrize("host", [None, "word", "powerpoint", "access"])
@pytest.mark.parametrize("lines", [
    ["Main = c"],
    ["Dim v As Variant", "v = c"],
    ["Dim v", "v = c"],
    ["Dim a(1) As Variant", "a(0) = c"],
    ["c.Add 1", "Dim v As Variant", "v = c", "If v.Count <> 1 Then Err.Raise 5"],
    ["Take (c)"],
])
def test_collection_returning_default_member_is_kept(lines: list[str], host: str | None) -> None:
    source = (
        "Option Explicit\nFunction Main() As Variant\n    Dim c As New C1\n" + "\n".join(lines)
        + "\nEnd Function\nSub Take(ByVal v As Variant)\nEnd Sub"
    )
    modules = [
        ModuleInput("Module1", ModuleSymbolKind.STANDARD, source),
        ModuleInput("C1", ModuleSymbolKind.CLASS, _DEFAULT_ITEMS_CLASS),
    ]
    findings = analyze_project(modules, host=host)["Module1"]
    assert [d for d in findings if d.severity == "error"] == []


# -- 28c839b1: module names a variable takes, and qualified Nothing reads ------
# The oracle cases issue1411_* (measured in Excel, Word, PowerPoint and Access).

_HOST_PREFIXES = ("", "word_", "powerpoint_", "access_")


def _case_errors(case_id: str) -> list[str]:
    """Each error finding as `code: spanned text`, analyzed in the case's host."""
    case = CASES[case_id]
    host = next((h.rstrip("_") for h in _HOST_PREFIXES[1:] if case_id.startswith(f"issue1411_{h}")), "excel")
    index = ProjectIndex()
    for module in case.modules:
        index.set_module(ModuleInput(module.name, _ORACLE_KIND.get(module.module_type, ModuleSymbolKind.STANDARD), module.source))
    out: list[str] = []
    for module in case.modules:
        opts = analyze_module_options_for(
            index, module.name, _ORACLE_KIND.get(module.module_type, ModuleSymbolKind.STANDARD), host=host,
            conditional_compilation=ConditionalCompilationEnvironment(project_constants={}),
        )
        out += [
            f"{d.code}: {module.source[d.span.start:d.span.end]}"
            for d in analyze_module(module.source, opts) if d.severity == "error"
        ]
    return out


_ORACLE_KIND = {
    "standard": ModuleSymbolKind.STANDARD,
    "class": ModuleSymbolKind.CLASS,
    "document": ModuleSymbolKind.DOCUMENT,
    "userform": ModuleSymbolKind.USERFORM,
}


def _each_host(names: list[str], phase: str) -> list[str]:
    return [f"issue1411_{host}{name}_{phase}" for name in names for host in _HOST_PREFIXES]


_SHADOWED_1411 = [
    f"b_{receiver}_{by}"
    for receiver in ("strings", "math", "other")
    for by in ("local", "parameter", "module", "variant_local", "current_module")
    if by != "current_module" or receiver != "other"
] + ["b_other_collection_local"] + [
    f"d_{shape}_shadowed"
    for shape in ("statement", "call_statement", "assignment", "nested_argument", "set_assignment")
]
_UNSHADOWED_1411 = ["b_strings_unshadowed", "b_math_unshadowed", "b_other_unshadowed"] + [
    f"d_{shape}_unshadowed"
    for shape in ("statement", "call_statement", "assignment", "nested_argument", "set_assignment")
] + ["d_public_variable_named_like_module"]


@pytest.mark.parametrize("case_id", _each_host(_SHADOWED_1411, "compile"))
def test_module_name_a_variable_takes_is_that_variable(case_id: str) -> None:
    assert CASES[case_id].expected == "accepted"
    assert _case_errors(case_id) == []


@pytest.mark.parametrize("case_id", _each_host(_UNSHADOWED_1411, "compile"))
def test_unshadowed_module_call_is_still_checked(case_id: str) -> None:
    assert CASES[case_id].expected == "rejected"
    assert len([f for f in _case_errors(case_id) if f.startswith("argument-count:")]) == 1


@pytest.mark.parametrize("case_id", _each_host(["b_other_long_local"], "compile"))
def test_long_named_like_module_is_a_member_of_a_long(case_id: str) -> None:
    findings = _case_errors(case_id)
    assert any(f.startswith("scalar-member-access:") for f in findings)
    assert not any(f.startswith("argument-count:") for f in findings)


@pytest.mark.parametrize("case_id", _each_host(
    [f"a_nothing_read_{form}" for form in ("print", "call", "label", "if", "assign", "with")]
    + ["a_reset_unqualified_read_qualified"],
    "runtime",
))
def test_module_variable_reset_through_its_module(case_id: str) -> None:
    assert CASES[case_id].expected == "rejected"
    assert _case_errors(case_id) == ["object-variable-not-set: Module2.g"]


@pytest.mark.parametrize("case_id", _each_host([
    "a_reset_then_new", "a_reset_then_call_init", "a_shadowed_own_module_receiver",
    "e_block_guard", "e_line_guard", "e_exit_between", "e_as_new",
], "runtime"))
def test_module_variable_reset_controls_run(case_id: str) -> None:
    assert CASES[case_id].expected == "accepted"
    assert _case_errors(case_id) == []


# -- 7283146d: a procedure's types come from its own module's DefType ----------

@pytest.mark.parametrize("case_id, expected", [
    *[(case_id, ["byref-argument-type-mismatch: s"]) for case_id in _each_host(["f_byref_string_variable"], "compile")],
    *[(case_id, ['argument-type-mismatch: "abc"']) for case_id in _each_host(["f_text_string_literal"], "runtime")],
    *[(case_id, []) for case_id in _each_host(["f_numeric_string_literal"], "runtime")],
    *[(case_id, []) for case_id in _each_host(["f_caller_deftype_only"], "compile")],
])
def test_untyped_parameter_takes_declaring_module_deftype(case_id: str, expected: list[str]) -> None:
    assert CASES[case_id].expected == ("accepted" if not expected else "rejected")
    assert _case_errors(case_id) == expected


_IS_REFUSED_1411 = [
    ("g_is_long_function", "Bar()"),
    ("g_is_long_function_with_argument", "Bar(1)"),
    ("g_is_string_function", "Bar()"),
    ("g_is_long_array_function", "Bar()"),
    ("g_is_deflng_bare", "Foo()"),
    ("g_is_deflng_qualified", "Other.Foo()"),
    ("g_is_deflng_caller_defobj", "Foo()"),
    ("g_typeof_long_identifier", "n"),
    ("g_typeof_string_identifier", "s"),
    ("g_typeof_long_function", "Bar()"),
    ("g_typeof_deflng_qualified", "Other.Foo()"),
]


@pytest.mark.parametrize("case_id, operand", [
    (case_id, operand) for name, operand in _IS_REFUSED_1411 for case_id in _each_host([name], "compile")
])
def test_is_and_typeof_refuse_a_scalar_operand(case_id: str, operand: str) -> None:
    assert CASES[case_id].expected == "rejected"
    assert _case_errors(case_id) == [f"is-operator-non-object: {operand}"]


@pytest.mark.parametrize("case_id", _each_host([
    "g_is_object_function", "g_is_variant_function", "g_is_defobj_qualified",
    "g_typeof_variant_identifier", "g_typeof_object_identifier", "g_typeof_defobj_bare",
], "compile"))
def test_is_and_typeof_accept_object_operands(case_id: str) -> None:
    assert CASES[case_id].expected == "accepted"
    assert _case_errors(case_id) == []


@pytest.mark.parametrize("case_id", _each_host(["g_is_deflng_shadowed_qualifier"], "compile"))
def test_is_leaves_a_late_bound_call_alone(case_id: str) -> None:
    assert [f for f in _case_errors(case_id) if f.startswith("is-operator-non-object:")] == []
