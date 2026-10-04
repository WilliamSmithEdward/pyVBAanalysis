"""Integration regressions for the XLIDE 2f49b93 sync."""

from pyvbaanalysis import analyze_module
from pyvbaanalysis.conditional import evaluate_conditional_expression
from pyvbaanalysis.constants import evaluate_integer_constant_expression
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.diagnostics.condition_value import ConditionFacts, condition_value, number_value
from pyvbaanalysis.diagnostics.const_expr import collect_module_literal_integer_constants
from pyvbaanalysis.diagnostics.rules.long_long_narrowing import _wide_expression
from pyvbaanalysis.diagnostics.rules.overflow import _NameLookup, _Typed, _fold
from pyvbaanalysis.lexer.tokenize import tokenize
from pyvbaanalysis.diagnostics.dataflow import BlockEnteringState, walk_entering_blocks
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.parser.expression_stack import run_expression
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex


def test_condition_folders_reach_upstream_depth() -> None:
    expression = "Abs(" * 200 + "1" + ")" * 200
    facts = ConditionFacts(value=lambda name: None)
    assert number_value(tokenize(expression), facts) == 1
    assert condition_value(tokenize("(" * 200 + "True" + ")" * 200), facts) is True


def test_expression_stack_preserves_parent_error_handling() -> None:
    def child():
        yield from ()
        raise ValueError("unknown value")

    def parent():
        try:
            return (yield child())
        except ValueError:
            return None

    assert run_expression(parent()) is None


def test_conditional_expression_reaches_upstream_depth() -> None:
    assert evaluate_conditional_expression("(" * 200 + "1" + ")" * 200) == 1


def test_overflow_folder_reaches_upstream_depth() -> None:
    value = _fold(tokenize("Sgn(" * 200 + "1" + ")" * 200), 0, _NameLookup(lambda name: None))
    assert isinstance(value, _Typed)
    assert value.value == 1


def test_longlong_reader_reaches_deep_parentheses() -> None:
    assert _wide_expression(tokenize("(" * 1100 + "1^" + ")" * 1100), {}, lambda name: False)


def test_fractional_val_constant_is_preserved() -> None:
    assert evaluate_integer_constant_expression('Val("1.5")', {}) == 1.5
    source = 'Public Const K = Val("1.5")'
    assert collect_module_literal_integer_constants(parse_module(source), None)["k"] == 1.5
    index = ProjectIndex()
    index.set_module(ModuleInput("Constants", ModuleSymbolKind.STANDARD, source))
    assert index.visible_external_integer_constant_expressions("Caller")["k"] == "1.5"


def test_integer_constant_folder_reaches_its_depth_budget() -> None:
    assert evaluate_integer_constant_expression("(" * 120 + "1" + ")" * 120, {}) == 1


def test_bare_builtin_in_function_result_is_checked() -> None:
    source = "Option Explicit\nFunction Main() As Variant\nMain = Name\nEnd Function"
    found = analyze_module(source, AnalyzeModuleOptions(known_identifiers=set(), raw_rule_output=True))
    assert any(d.code == "undeclared-variable" and "opens the Name statement" in d.message for d in found)


def test_subtree_touch_scan_is_linear() -> None:
    source = "Sub Main()\n" + "If x Then\n" * 100 + "x = 1\n" + "End If\n" * 100 + "End Sub"
    procedure = parse_module(source).members[0]
    scans = 0

    def touches(leaf):
        nonlocal scans
        scans += 1
        return {"x"}

    walk_entering_blocks(source, procedure.body, lambda node: False, lambda node: None,
                         BlockEnteringState(snapshot=lambda: None, restore=lambda state: None,
                                            forget=lambda names: None, touches=touches))
    assert scans <= 301
