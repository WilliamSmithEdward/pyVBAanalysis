"""conditionValue, conditionOperands and nullOperators: the pure readers of a
condition. The depth cases are ported from upstream's
tests/conditionValueDepth.test.ts."""

from __future__ import annotations

import pytest

from pyvbaanalysis.diagnostics.condition_operands import ConditionOperand, condition_operands
from pyvbaanalysis.diagnostics.condition_value import (
    ConditionFacts,
    condition_value,
    if_condition_tokens,
    number_value,
)
from pyvbaanalysis.diagnostics.null_operators import operator_yields_null
from pyvbaanalysis.diagnostics.walker import raw_expression_tokens

NO_FACTS = ConditionFacts(value=lambda _lower: None)


def _numeric(text: str, facts: ConditionFacts = NO_FACTS) -> float | None:
    return number_value(raw_expression_tokens(text), facts)


def _condition(text: str, facts: ConditionFacts = NO_FACTS) -> bool | None:
    return condition_value(raw_expression_tokens(text), facts)


@pytest.mark.parametrize("fn", ["Abs", "Len", "IIf"])
def test_thousands_of_nested_calls_are_not_known(fn: str) -> None:
    text = "IIf(" * 5000 + "1" + ", 1, 0)" * 5000 if fn == "IIf" else f"{fn}(" * 5000 + "1" + ")" * 5000
    assert _numeric(text) is None


def test_parentheses_and_prefix_operators_are_bounded() -> None:
    assert _condition("(" * 5000 + "True" + ")" * 5000) is None
    assert _condition("Not " * 5000 + "True") is None
    assert _numeric("-" * 5000 + "1") is None
    assert _numeric("Abs(" * 128 + "(" * 128 + "1" + ")" * 256) is None


def test_independent_operands_do_not_share_a_budget() -> None:
    assert _numeric(" + ".join(["Abs(1)"] * 500)) == 500
    assert _condition("Not Not (Abs(-2) = 2)") is True


def test_shallow_nesting_is_read() -> None:
    assert _numeric("Abs(" * 20 + "-1" + ")" * 20) == 1


def test_known_names_and_operators() -> None:
    known: dict[str, float | str] = {"n": 0, "s": "abc"}
    facts = ConditionFacts(value=known.get)
    assert _condition("n > 0", facts) is False
    assert _condition("n = 0 And s = \"abc\"", facts) is True
    assert _condition("n = 1 Or x", facts) is None
    assert _condition("Len(s) = 3", facts) is True
    assert _condition("IsNumeric(s)", facts) is False
    assert _condition("n Mod 2 = 0", facts) is True
    assert _numeric("7 \\ 2") == 3
    assert _numeric("2.5 \\ 1") == 2
    assert _numeric("-2 ^ 2") == -4
    assert _numeric("2147483647 + 1") is None
    assert _numeric("0 ^ -1") is None


def test_strings_compare_by_the_module_option_compare() -> None:
    assert _condition('"a" = "A"') is None
    assert _condition('"a" = "b"') is False
    binary = ConditionFacts(value=lambda _lower: None, compare="binary")
    text = ConditionFacts(value=lambda _lower: None, compare="text")
    assert _condition('"A" < "a"', binary) is True
    assert _condition('"a" = "A"', text) is True
    assert _condition('"abc" Like "a*"', binary) is True
    assert _condition('"abc" Like "A*"', binary) is False
    assert _condition('"abc" Like "A*"', text) is True
    assert _condition('"a1" Like "[a-c]#"', binary) is True
    assert _condition('"d1" Like "[!a-c]#"', binary) is True
    assert _condition('"b" Like "[c-a]"', binary) is None
    assert _condition('"x" Like "[]"', binary) is False


def test_string_functions_of_known_strings() -> None:
    binary = ConditionFacts(value=lambda _lower: None, compare="binary")
    assert _numeric('InStr("abcabc", "c")', binary) == 3
    assert _numeric('InStr(4, "abcabc", "c")', binary) == 6
    assert _numeric('InStr(1, "abc", "B", vbTextCompare)') == 2
    assert _numeric('StrComp("a", "b", vbBinaryCompare)') == -1
    assert _condition('Replace("aXa", "x", "b", 1, -1, vbTextCompare) = "aba"') is True
    assert _condition('UCase("ab") = "AB"') is True
    assert _numeric("InStr(1, \"\U0001F600x\", \"x\")", binary) == 3


def test_ranges_and_builtins() -> None:
    facts = ConditionFacts(value=lambda _lower: None, range=lambda lower: (1000, 1059) if lower == "b" else None)
    assert _condition("b > 5000", facts) is False
    assert _condition("b + 10 >= 1010", facts) is True
    assert _condition("b = 1030", facts) is None
    typed = ConditionFacts(value=lambda _lower: None, type_of=lambda lower: "long()" if lower == "a" else None)
    assert _numeric("VarType(a)", typed) == 8195
    assert _condition('TypeName(a) = "Long()"', typed) is True
    assert _condition("IsArray(a)", typed) is True
    assert _numeric("vbArray + vbLong") == 8195
    assert _condition('"n" & 5 & True = "n5True"') is True
    assert _numeric('Val("12abc")') == 12
    assert _numeric("Sgn(-3) + Int(-1.5) + Fix(-1.5)") == -4


def test_if_condition_tokens() -> None:
    toks = raw_expression_tokens("If (a Then) Then b")
    condition = if_condition_tokens(toks)
    assert condition is not None
    assert [tok.raw_text for tok in condition] == ["(", "a", "Then", ")"]
    assert if_condition_tokens(raw_expression_tokens("x = 1")) is None


def test_condition_operands() -> None:
    assert condition_operands(raw_expression_tokens("If x Then")) == [ConditionOperand(1, "condition")]
    assert condition_operands(raw_expression_tokens("Select Case c")) == [ConditionOperand(2, "select")]
    assert condition_operands(raw_expression_tokens("y = Not a And b")) == [
        ConditionOperand(3, "not"),
        ConditionOperand(5, "logical"),
    ]
    assert condition_operands(raw_expression_tokens("y = IIf(c, 1, 2)")) == [ConditionOperand(4, "iif")]
    assert condition_operands(raw_expression_tokens("y = o.IIf(c, 1, 2)")) == []


def test_operator_yields_null() -> None:
    def holds(text: str) -> bool:
        return operator_yields_null(raw_expression_tokens(text), lambda tok: tok.raw_text.lower() == "null")

    assert holds("Null + 1")
    assert holds("-(Null)")
    assert holds("Abs(Null)")
    assert not holds("Null & 1")
    assert not holds("Null And 0")
    assert holds("Null And 1")
    assert not holds("40000 Or Null")
    assert holds("0 Or Null")
    assert not holds("False Imp Null")
    assert holds("Null Imp False")
    assert not holds("Not " * 3000 + "1")
