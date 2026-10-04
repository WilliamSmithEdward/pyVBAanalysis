"""Rule family: overflow the analyzer can prove (XLIDE issue #116).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/overflow.ts.

VBA types a whole-number literal as Integer when it fits and does Integer
arithmetic on two Integers, so `60 * 60 * 24` overflows before the Long it is
assigned to ever sees it. Every case here was measured in Excel 16.0 (build
20326, 2026-09-26): each compiles and raises error 6 every time it runs, or - for
a Const - is refused with "Overflow" while compiling.

 - arithmetic-overflow: `secs = 60 * 60 * 24`, `32767 + 1` into a Long,
   `50000 * 50000`, `2147483647 + 1`, `10 ^ 309`, `Exp(1000)`, Integer Consts
   multiplied, `CInt(40000)`, `CByte(-1)`, `CLng(2147483647.5)`,
   `CSng(1E+39)`, `Hex(1E+20)`, `Abs(CInt(-32768))`, `-i` with i = -32768,
   and an assignment whose folded value the target type cannot hold after
   rounding: `Byte = 255.5`, `Integer = 32767.5`, `Date = 3000000`.
 - const-overflow: the same folding on a Const's value, a compile error.
   Inside any argument or operand too (issue #232): `CStr(CInt(40000))`,
   `IIf(True, 0, CInt(40000))`, `"x" & CInt(40000)`, `z(CInt(40000))`.
   LongLong: `CLngLng(1E+19)`, `9223372036854775807^ + 1`; a LongPtr is
   judged against LongLong's range, which it never exceeds.
 - for-counter-overflow: `For i = 1 To 32767` with i an Integer, and
   `For b = 0 To 255` with b a Byte: the increment after the last pass
   overflows the counter. `To 32766` runs.

The folder follows MS-VBAL 5.6.9.3: two Bytes make a Byte, Byte and Integer
make Integer, Long makes Long, Single and Double make Double, Currency makes
Currency, and a Date plus or minus a number is a Date; `/` and `^` make Double.
A value the folder cannot type stays unknown and nothing is reported for it.

Every value is a Python float, as upstream's are JavaScript numbers, so the
arithmetic rounds and overflows to infinity the way upstream's does; exact
LongLong, Currency and Decimal values are Python ints, as upstream's are
bigints. The helpers at the end of the module print numbers the way JavaScript
does.
"""

from __future__ import annotations

import math
import re
import time
from collections.abc import Callable, Generator, Iterator, Mapping, Sequence
from dataclasses import dataclass, field, replace
from decimal import ROUND_HALF_UP, Context, Decimal
from functools import lru_cache
from typing import Any, Union, cast

from ...conditional import ConditionalActivityTracker
from ...constants.date_literal import DATE_EPOCH_MS, DAY_MS, date_literal_serial
from ...constants.integer_constant_expression import bankers_round, parse_vba_integer_literal
from ...flow.procedure_labels import jump_target_label_declaration
from ...host.host_model import HostObjectModel
from ...js_compat import JS_WHITESPACE, js_number, js_number_to_string, js_trim, utf16_length
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.expression_limits import MAX_EXPRESSION_DEPTH
from ...parser.expression_stack import run_expression
from ...parser.nodes import (
    BodyNode,
    DoBlockNode,
    ForBlockNode,
    IfBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    SelectBlockNode,
    Span,
    StatementNode,
    VariableGroupNode,
    WhileBlockNode,
    WithBlockNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ...symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbol, VbaSymbolKind
from ...types.type_inference import (
    def_type_of,
    known_local_literal_values_at,
    procedure_symbol_for,
    type_environment_for,
)
from ...types.type_names import normalize_type
from ..block_headers import block_header_leaves, block_header_statements, is_loop_block, select_arms
from ..call_extraction import string_literal_value
from ..callable_signatures import build_module_type_signatures
from ..context import PushFn, statement_tokens
from ..dataflow import tracked_locals_named_whole
from ..function_results import FunctionResult, function_result_named, known_function_results
from ..known_locals import KnownLocalValue
from ..loop_counters import check_each_counter_pass, loop_counters_at
from ..string_conversion import numeric_string_verdict, val_prefix_value
from ..type_fields import ModuleTypes, field_chain, module_types, variable_root, variable_symbol_in
from ..walker import (
    active_module_members,
    bare_assignment_target,
    block_footer_line_span,
    block_header_line_span,
    first_executable_token_index,
    for_each_statement,
    for_each_variable_group,
    raw_expression_tokens,
    set_assignment_target,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .shared import body_may_leave_loop, is_bare_or_vba_qualified_intrinsic_call, names_in


@dataclass(frozen=True, slots=True)
class _Typed:
    value: float
    # 'byte' | 'integer' | 'long' | 'longlong' | 'single' | 'double' | 'currency' | 'date'
    type: str
    # A LongLong's exact value: a double cannot tell 2^63 - 1 from 2^63.
    exact: int | None = None
    # Made of literals and Consts only, so the VBE folds it while compiling:
    # negating the Long minimum there wraps to itself (issue #235).
    constant: bool = False
    # A Boolean, which folds as the Integer -1 or 0 but goes into a Byte as 255
    # or 0: `b = True` and `CByte(True)` store 255 (issue #326).
    boolean: bool = False
    # A Date past the Date range that `number + Date` or `number - Date` made
    # without raising (issue #330): the expression, as the message names it.
    # Most uses of it raise (issue #405).
    past_date: str | None = None
    # Held in a Variant, whose arithmetic widens the type instead of
    # overflowing: 32767 + 1 is the Long 32768 (issue #480, measured in Excel
    # 16.0).
    variant: bool = False
    # A Currency value in ten-thousandths, exactly (issue #494).
    scaled: int | None = None
    # A whole Decimal's exact value, from CDec (issue #502): the type says Double.
    decimal: int | None = None


@dataclass(frozen=True, slots=True)
class _Overflow:
    span: Span
    detail: str


_Folded = Union[_Typed, _Overflow, None]


@dataclass(frozen=True, slots=True)
class _Range:
    min: float
    max: float
    label: str


_RANGES: dict[str, _Range] = {
    "byte": _Range(0.0, 255.0, "Byte"),
    "integer": _Range(-32768.0, 32767.0, "Integer"),
    "long": _Range(-2147483648.0, 2147483647.0, "Long"),
    # As doubles these are -2^63 and 2^63; _in_range compares a LongLong exactly.
    "longlong": _Range(-9223372036854775808.0, 9223372036854775807.0, "LongLong"),
    "single": _Range(-3.402823e38, 3.402823e38, "Single"),
    "double": _Range(-1.7976931348623157e308, 1.7976931348623157e308, "Double"),
    "currency": _Range(-922337203685477.5807, 922337203685477.5807, "Currency"),
    "date": _Range(-657434.0, 2958465.0, "Date"),
}

_RANK: dict[str, int] = {
    "byte": 0, "integer": 1, "long": 2, "longlong": 3, "single": 4, "double": 5, "currency": 6, "date": 7,
}

# The types a value is stored in without rounding to a whole number.
_UNROUNDED_TYPES = frozenset({"single", "double", "currency", "date"})


def _label(type_name: str) -> str:
    return _RANGES[type_name].label


def _arithmetic_result_type(a: str, b: str, op: str) -> str:
    """The result type of `a op b` for + - * \\ Mod (MS-VBAL 5.6.9.3), as Excel 16.0
    computes it (issue #203): two Bytes make a Byte, so 200 + 100 overflows; a Date
    plus or minus a number, or two Dates added, make a Date, so #12/31/9999# + 1
    overflows; two Dates subtracted make a Double."""
    if (a == "longlong" or b == "longlong") and (a == "single" or b == "single"):
        return "double"
    # A Date with a Currency is a Date for + and -, and a Double for *: c + t is
    # VarType 7, t * c VarType 5 (issue #409, measured in Excel 16.0).
    if (a == "date" or b == "date") and op in ("+", "-", "*"):
        return "double" if op == "*" or (a == "date" and b == "date" and op == "-") else "date"
    if a == "currency" or b == "currency":
        return "double" if a in ("double", "single") or b in ("double", "single") else "currency"
    if a == "date" or b == "date":
        return "date" if op == "+" or (op == "-" and not (a == "date" and b == "date")) else "double"
    return a if _RANK[a] >= _RANK[b] else b


# 2^63: one past the largest LongLong.
_LONGLONG_LIMIT = 2**63


def _in_range(value: float, type_name: str, exact: int | None = None) -> bool:
    if type_name == "longlong":
        if exact is not None:
            return -_LONGLONG_LIMIT <= exact < _LONGLONG_LIMIT
        return math.isfinite(value) and -(2.0**63) <= value < 2.0**63
    bounds = _RANGES[type_name]
    # A Date's range is of days: any time of 12/31/9999 is in it.
    checked = _js_trunc(value) if type_name == "date" else value
    return math.isfinite(value) and bounds.min <= checked <= bounds.max


def _round(value: float) -> float:
    """bankersRound, VBA's rounding to a whole number."""
    return float(bankers_round(value))


_INTEGER_SUFFIX_RE = re.compile(r"[%&^]$")
_FLOAT_SUFFIX_RE = re.compile(r"[!#@]$")
_EXPONENT_LETTER_RE = re.compile(r"[dD]")
_VOWEL_START_RE = re.compile(r"^[AEIOU]")
_DIGITS_RE = re.compile(r"[0-9]+")
_RADIX_LONGLONG_RE = re.compile(r"&([Hh])([0-9A-Fa-f]+)|&[Oo]?([0-7]+)")
_JS_SPACE = "[" + JS_WHITESPACE + "]"
_DECIMAL_STRING_RE = re.compile(r"[-+]?[0-9]+(?:[eE][-+]?[0-9]+)?")
_RADIX_STRING_RE = re.compile(r"([-+]?)(&[Hh][0-9A-Fa-f]+|&[Oo]?[0-7]+)")
_SPELLED_WHOLE_RE = re.compile('"?' + _JS_SPACE + "*([-+]?)([0-9]+)" + _JS_SPACE + '*"?')
_CURRENCY_TEXT_RE = re.compile(r"([0-9]*)(?:\.([0-9]{0,4}))?")
_ANY_DIGIT_RE = re.compile(r"[0-9]")
_BOOLEAN_TEXT_RE = re.compile(_JS_SPACE + r"*(true|false)" + _JS_SPACE + r"*", re.IGNORECASE | re.ASCII)
# JavaScript's line terminators besides "\n", where its `^` with the m flag also
# matches; and its `.`, which matches none of them.
_OTHER_LINE_ENDS = "\r" + chr(0x2028) + chr(0x2029)
_LINE_START = "(?:^|(?<=[" + _OTHER_LINE_ENDS + "]))"
_JS_DOT = "[^\n" + _OTHER_LINE_ENDS + "]"
_DEFTYPE_RE = re.compile(
    _LINE_START + r"[ \t]*Def(Bool|Byte|Int|Lng|LngLng|LngPtr|Cur|Sng|Dbl|Dec|Date|Str|Obj|Var)[ \t]+[A-Za-z]",
    re.IGNORECASE | re.MULTILINE | re.ASCII,
)
_REDIM_LINE_RE = re.compile(
    _LINE_START + r"[ \t]*(?:[0-9]+[ \t]+)?ReDim[ \t]+(?:Preserve[ \t]+)?([^'\r\n]*)",
    re.IGNORECASE | re.MULTILINE | re.ASCII,
)
_REDIM_ITEM_RE = re.compile(
    "^" + _JS_SPACE + r"*([A-Za-z][A-Za-z0-9_]*)" + _JS_SPACE + r"*\(" + _JS_DOT + r"*\)" + _JS_SPACE
    + r"*(?:As" + _JS_SPACE + r"+([A-Za-z][A-Za-z0-9_]*))?" + _JS_SPACE + r"*\Z",
    re.IGNORECASE | re.ASCII,
)
_EMPTY_ARRAY_SUFFIX_RE = re.compile(r"\(" + _JS_SPACE + r"*\)" + _JS_SPACE + r"*$")
_CELL_RANGE_RE = re.compile(r"([A-Z]{1,3})([0-9]+)(?::([A-Z]{1,3})([0-9]+))?")
_WHOLE_COLUMNS_RE = re.compile(r"([A-Z]{1,3}):([A-Z]{1,3})")
_WHOLE_ROWS_RE = re.compile(r"([0-9]+):([0-9]+)")


def _literal_typed(tok: VbaToken | None) -> _Typed | None:
    """A literal's natural type and value: 3 is Integer, 40000 is Long, 3000000000
    is Double."""
    typed = _literal_value(tok)
    return replace(typed, constant=True) if typed is not None else None


def _literal_value(tok: VbaToken | None) -> _Typed | None:
    if tok is None:
        return None
    # A Boolean in arithmetic is an Integer: True is -1 and False 0, so `1 /
    # False` divides by zero (issue #235).
    if tok.kind is TokenKind.KEYWORD:
        word = token_text(tok)
        if word == "true":
            return _Typed(-1.0, "integer", boolean=True)
        if word == "false":
            return _Typed(0.0, "integer", boolean=True)
        return None
    if tok.kind is TokenKind.INTEGER_LITERAL:
        raw = tok.raw_text
        suffix_match = _INTEGER_SUFFIX_RE.search(raw)
        suffix = suffix_match.group(0) if suffix_match is not None else None
        if suffix == "^":
            exact = _long_long_literal(raw)
            return None if exact is None else _Typed(float(exact), "longlong", exact=exact)
        parsed = parse_vba_integer_literal(raw)
        if parsed is None:
            return None
        value = float(parsed)
        if suffix == "%":
            return _Typed(value, "integer") if _in_range(value, "integer") else None
        if suffix == "&":
            return _Typed(value, "long")
        # A hex or octal literal arrives already signed by its width
        # (parse_vba_integer_literal, XLIDE issue #141): &H8000 is -32768 and an
        # Integer, &H80000000 is -2147483648 and a Long.
        if _in_range(value, "integer"):
            return _Typed(value, "integer")
        if _in_range(value, "long"):
            return _Typed(value, "long")
        return _Typed(value, "double")
    if tok.kind is TokenKind.FLOAT_LITERAL:
        raw = _EXPONENT_LETTER_RE.sub("E", tok.raw_text)
        suffix_match = _FLOAT_SUFFIX_RE.search(raw)
        suffix = suffix_match.group(0) if suffix_match is not None else None
        value = js_number(_FLOAT_SUFFIX_RE.sub("", raw, count=1))
        if not math.isfinite(value):
            return None
        scaled = _currency_scaled(raw[:-1] if raw.endswith("@") else raw) if suffix == "@" else None
        type_name = "single" if suffix == "!" else "currency" if suffix == "@" else "double"
        return _Typed(value, type_name, scaled=scaled)
    if tok.kind is TokenKind.DATE_LITERAL:
        serial = date_literal_serial(tok.raw_text)
        return None if serial is None else _Typed(float(serial), "date")
    return None


def _long_long_literal(raw: str) -> int | None:
    """A `^` literal's exact value: `9223372036854775807^`, `&H7FFFFFFFFFFFFFFF^` (a
    hex or octal literal keeps the sign of its width). A decimal past the range is
    a syntax error that suffixed-literal-overflow reports."""
    body = raw[:-1]
    if _DIGITS_RE.fullmatch(body) is not None:
        value = int(body)
        return value if value < _LONGLONG_LIMIT else None
    radix = _RADIX_LONGLONG_RE.fullmatch(body)
    if radix is None:
        return None
    value = int(radix.group(2), 16) if radix.group(2) is not None else int(radix.group(3), 8)
    if value >= 2**64:
        return None
    return value - 2**64 if value >= 2**63 else value


def _number_in_string(text: str) -> _Typed | str | None:
    """The number a conversion reads from a string (issue #184), for the spellings
    every locale reads alike: whole digits with an optional exponent, and &H and &O
    literals, which keep the sign of their width as a literal does ("&H8000" is
    -32768). A decimal point or a thousands separator is read by the locale and is
    not judged. "overflow" for a number past the Double range."""
    trimmed = js_trim(text)
    if _DECIMAL_STRING_RE.fullmatch(trimmed) is not None:
        value = js_number(trimmed)
        return _Typed(value, "double") if math.isfinite(value) else "overflow"
    radix = _RADIX_STRING_RE.fullmatch(trimmed)
    parsed = parse_vba_integer_literal(radix.group(2)) if radix is not None else None
    if radix is None or parsed is None:
        # A parenthesized or trailing sign, `"(5)"` and `"5-"`, is -5 in every
        # locale (issue #703, measured in Excel 16.0).
        verdict = numeric_string_verdict(trimmed)
        if verdict.kind == "number" and verdict.value is not None:
            return _Typed(float(verdict.value), "double")
        return None
    return _Typed(float(-parsed if radix.group(1) == "-" else parsed), "double")


def _val_of_string(text: str) -> _Typed | str | None:
    """What `Val` reads from a string, the same in every locale (issue #703)."""
    value = val_prefix_value(text)
    if value is None:
        return None
    return _Typed(float(value), "double") if math.isfinite(value) else "overflow"


@dataclass(frozen=True, slots=True)
class _ChainSegment:
    name: str
    args: list[list[VbaToken]] | None = None


class _NameLookup:
    """What a name means to the folder: a typed value, or nothing. `declares`, where
    given, says the procedure or module declares the name, so a host global of that
    spelling is hidden."""

    __slots__ = ("_lookup", "declares", "whole_sheet_cells", "with_subject")

    def __init__(
        self,
        lookup: Callable[[str], _Typed | None],
        declares: Callable[[str], bool] | None = None,
        whole_sheet_cells: Callable[[str], bool] | None = None,
        with_subject: Callable[[], Sequence[_ChainSegment] | None] | None = None,
    ) -> None:
        self._lookup = lookup
        self.declares = declares
        # A local that holds a whole sheet's Cells: `Set r = Cells` (issue #278).
        self.whole_sheet_cells = whole_sheet_cells
        # The innermost With's subject as a chain: `With ActiveSheet.Range("A1:A5")`
        # (issue #685).
        self.with_subject = with_subject

    def __call__(self, lower: str) -> _Typed | None:
        return self._lookup(lower)


# The operators that read both sides as numbers: arithmetic and comparison. The
# logical operators convert a String to a number too: `"" And 255` is a Type
# mismatch (issue #494, measured in Excel 16.0).
_BINARY_ON_NUMBERS = frozenset(
    {"+", "-", "*", "/", "\\", "^", "mod", "=", "<>", "<", ">", "<=", ">=", "and", "or", "xor", "eqv", "imp"}
)

# The operators a String beside a number converts under in a Const (issue #494).
_ARITHMETIC_BESIDE = frozenset({"+", "-", "*", "/"})

# Logical precedence; equal precedence splits at the last operator.
_LOGICAL_PRECEDENCE: dict[str, int] = {"imp": 0, "eqv": 1, "xor": 2, "or": 3, "and": 4}

# A `Null` operand of a logical operator, told apart by identity (issue #685).
_NULL_OPERAND = _Typed(0.0, "long")


class _NotLogical:
    """What _TypedFolder._logical answers for an expression with no logical operator."""


_NOT_LOGICAL = _NotLogical()


@dataclass(slots=True)
class _LogicalOperator:
    at: int
    rank: int
    word: str


@dataclass(slots=True)
class _LogicalValue:
    value: _Typed
    start: int
    to: int


class _TypedFolder:
    """Folds an arithmetic expression over literals, Consts and known locals with
    VBA's result typing, stopping at the first operation whose result its type
    cannot hold.

    It recurses only over one expression's tokens: parentheses, unary signs and
    call arguments.
    """

    __slots__ = ("_toks", "_base", "_names", "_division_by_zero", "_nesting", "_index")

    def __init__(
        self,
        toks: Sequence[VbaToken],
        base: int,
        names: _NameLookup,
        division_by_zero: Callable[[Span], None] | None = None,
        nesting: int = 0,
    ) -> None:
        self._toks = toks
        self._base = base
        self._names = names
        # Told of a division by zero, which a Const cannot hold: "Division by
        # zero" while compiling.
        self._division_by_zero = division_by_zero
        self._nesting = nesting
        self._index = 0

    def _child(self, toks: Sequence[VbaToken], division_by_zero: Callable[[Span], None] | None, nesting: int) -> Generator[Any, Any, _Folded]:
        return (cast("_Typed | None", (yield _TypedFolder(toks, self._base, self._names, division_by_zero, nesting).fold())))

    def fold(self) -> Generator[Any, Any, _Folded]:
        # Parsing already treats deeper expressions as recovery input. Keep an
        # untypable expression from aborting the rest of the overflow rule.
        if self._nesting >= MAX_EXPRESSION_DEPTH:
            return None
        if len(self._toks) == 0:
            return None
        logical = (cast("_Folded | _NotLogical", (yield self._logical())))
        if not isinstance(logical, _NotLogical):
            return logical
        # `Not` binds below every arithmetic operator: `Not 255 + 256` is Not 511
        # (issue #235, measured in Excel 16.0).
        if self._toks[0].kind is TokenKind.KEYWORD and token_text(self._toks[0]) == "not":
            operand = (cast("_Folded", (yield self._child(self._toks[1:], self._division_by_zero, self._nesting + 1))))
            if operand is None or isinstance(operand, _Overflow):
                return operand
            return _not_of(operand, self._span(0, len(self._toks) - 1))
        result = (cast("_Folded", (yield self._additive())))
        if isinstance(result, _Overflow):
            return result
        return result if self._index == len(self._toks) else None

    def _at(self, i: int) -> VbaToken | None:
        return self._toks[i] if 0 <= i < len(self._toks) else None

    def _raw_at(self, i: int) -> str | None:
        tok = self._at(i)
        return tok.raw_text if tok is not None else None

    def _logical(self) -> Generator[Any, Any, _Folded | _NotLogical]:
        """`a And b`, Or, Xor, Eqv and Imp follow VBA precedence and left
        associativity. Each operand is converted to a Long first, so one outside
        the Long range overflows: `1E10 And 1` is "Overflow" in a Const (issue
        #367, measured in Excel 16.0) and error 6 at run time (#323)."""
        depth = 0
        operators: list[_LogicalOperator] | None = None
        for i, tok in enumerate(self._toks):
            depth += 1 if tok.raw_text == "(" else -1 if tok.raw_text == ")" else 0
            if depth != 0 or i == 0 or tok.kind is not TokenKind.KEYWORD:
                continue
            word = token_text(tok)
            rank = _LOGICAL_PRECEDENCE.get(word)
            if rank is not None:
                if operators is None:
                    operators = []
                operators.append(_LogicalOperator(i, rank, word))
        if operators is None:
            return _NOT_LOGICAL
        # Consume disjoint operands from left to right. Reducing higher or equal
        # precedence before the next operator preserves the recursive evaluator's
        # left associativity and its first unknown/overflow result.
        values: list[_LogicalValue] = []
        pending: list[_LogicalOperator] = []
        start = 0
        for i in range(len(operators) + 1):
            incoming = operators[i] if i < len(operators) else None
            to = incoming.at if incoming is not None else len(self._toks)
            operand = self._toks[start:to]
            # A Null operand makes the result Null, but the other is still
            # converted to a Long: `Null And 1E10` overflows (issue #685).
            is_null = len(operand) == 1 and token_text(operand[0]) == "null"
            folded: _Folded
            if is_null:
                folded = _NULL_OPERAND
            else:
                folded = self._string_operand(operand)
                if folded is None:
                    folded = (cast("_Folded", (yield self._child(operand, self._division_by_zero, self._nesting))))
            if folded is None or isinstance(folded, _Overflow):
                return folded
            values.append(_LogicalValue(folded, start, to - 1))
            while pending and (incoming is None or pending[-1].rank >= incoming.rank):
                op = pending.pop()
                right = values.pop()
                left = values.pop()
                if left.value is _NULL_OPERAND or right.value is _NULL_OPERAND:
                    combined = self._null_logical(left.value, right.value, op.at, left.start, right.to)
                else:
                    combined = self._logical_value(left.value, right.value, op.word, op.at, left.start, right.to)
                if combined is None or isinstance(combined, _Overflow):
                    return combined
                values.append(_LogicalValue(combined, left.start, right.to))
            if incoming is not None:
                pending.append(incoming)
                start = incoming.at + 1
        return values[0].value

    def _null_logical(self, left: _Typed, right: _Typed, at: int, start: int, to: int) -> _Folded:
        """A logical operator with a Null side: the other side past the Long range
        overflows; otherwise the result is not followed."""
        other = right if left is _NULL_OPERAND else left
        if other is _NULL_OPERAND or other.type == "longlong" or _in_range(_round(other.value), "long"):
            return None
        return _Overflow(
            self._span(start, to),
            f"{_describe(other)} is outside the Long range that {self._toks[at].raw_text} converts its operands to",
        )

    def _logical_value(self, left: _Typed, right: _Typed, word: str, at: int, start: int, to: int) -> _Folded:
        if left.type == "longlong" or right.type == "longlong":
            return None
        span = self._span(start, to)
        operands = [_round(left.value), _round(right.value)]
        outside = next((typed for k, typed in enumerate((left, right)) if not _in_range(operands[k], "long")), None)
        if outside is not None:
            return _Overflow(
                span,
                f"{_describe(outside)} is outside the Long range that {self._toks[at].raw_text} converts its operands to",
            )
        # Both are whole numbers in the Long range, where Python's bitwise
        # operators agree with JavaScript's 32-bit ones.
        a, b = int(operands[0]), int(operands[1])
        if word == "and":
            value = a & b
        elif word == "or":
            value = a | b
        elif word == "xor":
            value = a ^ b
        elif word == "eqv":
            value = ~(a ^ b)
        else:
            value = ~a | b

        def small(operand: _Typed) -> bool:
            return operand.type in ("byte", "integer")

        type_name = (
            "byte" if left.type == "byte" and right.type == "byte" else "integer" if small(left) and small(right) else "long"
        )
        kept = value & 0xFF if type_name == "byte" else value
        return _Typed(
            float(kept),
            type_name,
            constant=left.constant and right.constant,
            boolean=left.boolean and right.boolean,
        )

    def _string_operand(self, toks: Sequence[VbaToken]) -> _Folded:
        """A string literal operand of an operator that converts to a number first, a
        logical one, `\\` or Mod, read as the number it spells: `"3E9" Or 0`
        overflows the Long and `"1" \\ False` divides by zero in a Const (issue #458,
        measured in Excel 16.0). None for anything else, and for a string that
        spells no number."""
        if len(toks) != 1 or toks[0].kind is not TokenKind.STRING_LITERAL:
            return None
        tok = toks[0]
        read = _number_in_string(string_literal_value(tok.raw_text))
        span = Span(self._base + tok.start, self._base + tok.end)
        if read == "overflow":
            return _Overflow(span, f"{tok.raw_text} spells a number past the Double range")
        return replace(read, constant=True) if isinstance(read, _Typed) else None

    def _span(self, start: int, to: int) -> Span:
        return Span(self._base + self._toks[start].start, self._base + self._toks[to].end)

    def _additive(self) -> Generator[Any, Any, _Folded]:
        start = self._index
        left = (cast("_Folded", (yield self._multiplicative())))
        while isinstance(left, _Typed):
            op = self._at(self._index)
            if op is None or op.kind is not TokenKind.OPERATOR or op.raw_text not in ("+", "-"):
                break
            self._index += 1
            right = (cast("_Folded", (yield self._multiplicative())))
            if not isinstance(right, _Typed):
                return right
            left = self._combine(left, right, op.raw_text, start, self._index - 1)
        return left

    # MS-VBAL 5.6.9 arithmetic precedence, highest first: ^, unary minus, * and /,
    # \, Mod, + and -. Folding *, /, \, Mod and ^ at one level read `32000 \ 2 * 4`
    # as 16000 * 4 and reported an overflow on code that runs, and missed
    # `1 Mod 200 * 200`, which does overflow (XLIDE issue #145).
    def _left_associative(
        self,
        operators: tuple[str, ...],
        left: Callable[[], Generator[Any, Any, _Folded]],
        right: Callable[[], Generator[Any, Any, _Folded]] | None = None,
    ) -> Generator[Any, Any, _Folded]:
        operand_of = right if right is not None else left
        start = self._index
        value = (cast("_Folded", (yield left())))
        while isinstance(value, _Typed):
            op = self._at(self._index)
            if op is None:
                break
            word = op.raw_text if op.kind is TokenKind.OPERATOR else token_text(op)
            if word not in operators:
                break
            self._index += 1
            operand = (cast("_Folded", (yield operand_of())))
            if not isinstance(operand, _Typed):
                return operand
            value = self._combine(value, operand, word, start, self._index - 1)
        return value

    def _multiplicative(self) -> Generator[Any, Any, _Folded]:
        return (cast("_Folded", (yield self._left_associative(("mod",), self._integer_division))))

    def _integer_division(self) -> Generator[Any, Any, _Folded]:
        return (cast("_Folded", (yield self._left_associative(("\\",), self._product))))

    def _product(self) -> Generator[Any, Any, _Folded]:
        return (cast("_Folded", (yield self._left_associative(("*", "/"), self._unary))))

    def _power(self) -> Generator[Any, Any, _Folded]:
        """`a ^ b` binds above unary minus (`-2 ^ 2` is -4); the exponent may carry
        its own sign."""
        return (cast("_Folded", (yield self._left_associative(("^",), self._primary, self._unary))))

    def _unary(self, depth: int = 0) -> Generator[Any, Any, _Folded]:
        tok = self._at(self._index)
        if tok is not None and tok.kind is TokenKind.OPERATOR and tok.raw_text in ("-", "+"):
            if depth + self._nesting >= MAX_EXPRESSION_DEPTH:
                return None
            start = self._index
            self._index += 1
            operand = (cast("_Folded", (yield self._unary(depth + 1))))
            if not isinstance(operand, _Typed):
                return operand
            if tok.raw_text == "+":
                return operand
            # `-&H80000000` and `-(-2147483647 - 1)` are folded while compiling and
            # give the Long minimum back; `l = -l` overflows (issue #235).
            if operand.constant and operand.type == "long" and operand.value == _RANGES["long"].min:
                return operand
            value = -operand.value
            exact = -operand.exact if operand.exact is not None else None
            # Negating a Byte gives an Integer: `-b` with b = 200 is -200 (issue
            # #362, measured in Excel 16.0).
            type_name = "integer" if operand.type == "byte" else operand.type
            if not _in_range(value, type_name, exact):
                negated = str(operand.exact) if exact is not None else _show_number(operand.value)
                gives = str(exact) if exact is not None else _show_number(-operand.value)
                return _Overflow(
                    self._span(start, self._index - 1),
                    f"Negating {negated} gives {gives}, which does not fit {_label(type_name)}",
                )
            return _Typed(value, type_name, exact=exact, constant=operand.constant)
        return (cast("_Folded", (yield self._power())))

    def _word(self, at: int) -> str:
        tok = self._at(at)
        return tok.raw_text if tok is not None and tok.kind is TokenKind.OPERATOR else token_text(tok)

    def _primary(self) -> Generator[Any, Any, _Folded]:
        tok = self._at(self._index)
        if tok is None:
            return None
        if tok.raw_text == "(":
            close = match_paren_from(self._toks, self._index)
            if close < 0:
                return None
            inner_start, inner_end = self._index + 1, close
            # `((x))` folds as `(x)`: peel the layers that only wrap another
            # parenthesis here, so redundant nesting costs no recursion. Each
            # layer still counts toward the depth limit, as a fold of it would.
            peeled = 0
            matches: list[int] | None = None
            while (
                inner_end - inner_start >= 2
                and self._toks[inner_start].raw_text == "("
                and self._nesting + 1 + peeled < MAX_EXPRESSION_DEPTH
            ):
                if matches is None:
                    matches = _paren_matches(self._toks)
                if matches[inner_start] != inner_end - 1:
                    break
                inner_start += 1
                inner_end -= 1
                peeled += 1
            value = (cast("_Folded", (yield self._child(self._toks[inner_start:inner_end], self._division_by_zero, self._nesting + 1 + peeled))))
            self._index = close + 1
            return value
        literal = _literal_typed(tok)
        if literal is not None:
            self._index += 1
            return literal
        date = self._current_date()
        if date is None:
            date = self._date_serial()
        if date is not None:
            return date
        # `"1" \ False`: the string is an operand of `\` or Mod itself, not of a
        # `*`, `/` or `^` that binds tighter.
        before = self._word(self._index - 1)
        after = self._word(self._index + 1)
        is_string = tok.kind is TokenKind.STRING_LITERAL
        if (
            is_string
            and (before in ("\\", "mod") or after in ("\\", "mod"))
            and after not in ("*", "/", "^")
            and before not in ("*", "/", "^", "-", "+")
        ):
            self._index += 1
            return self._string_operand([tok])
        # `^` takes both operands as Doubles, a string literal too: `"12" ^ 32767`
        # overflows (issue #331, measured in Excel 16.0).
        if is_string and (before == "^" or after == "^"):
            read = _number_in_string(string_literal_value(tok.raw_text))
            if read == "overflow":
                return _Overflow(
                    self._span(self._index, self._index), f"{tok.raw_text} spells a number past the Double range"
                )
            if isinstance(read, _Typed):
                self._index += 1
                return _Typed(read.value, "double")
        # In a Const, a string beside a number in `+`, `-`, `*` or `/` is the
        # number it spells, a Double: `1 - "1E3"` is -999, and `&H7FFFFFFF * "2"`
        # is 4294967294 (issue #556). Beside a Currency it is a Currency:
        # `922337203685477.5807@ + "2"` overflows (issue #494, measured in Excel
        # 16.0). Two strings under `+` join instead.
        beside = (
            self._at(self._index - 2)
            if before in _ARITHMETIC_BESIDE
            else self._at(self._index + 2)
            if after in _ARITHMETIC_BESIDE
            else None
        )
        beside_value = (
            _literal_value(beside) if beside is not None and beside.kind is not TokenKind.STRING_LITERAL else None
        )
        beside_type = beside_value.type if beside_value is not None else None
        if is_string and self._division_by_zero is not None and beside_type:
            read = _number_in_string(string_literal_value(tok.raw_text))
            if not isinstance(read, _Typed):
                if read is None:
                    return None
                return _Overflow(
                    self._span(self._index, self._index), f"{tok.raw_text} spells a number past the Double range"
                )
            self._index += 1
            type_name = "currency" if beside_type == "currency" else "double"
            if not _in_range(read.value, type_name):
                return _Overflow(
                    self._span(self._index - 1, self._index - 1),
                    f"{tok.raw_text} is outside the {_label(type_name)} range",
                )
            scaled = _currency_scaled(js_number_to_string(abs(read.value))) if type_name == "currency" else None
            if scaled is not None and read.value < 0:
                scaled = -scaled
            return _Typed(read.value, type_name, constant=True, scaled=scaled)
        # A With member at an operand's start: `.Rows.Count` (issue #411).
        previous = self._at(self._index - 1)
        leading = (
            self._index == 0
            or (previous is not None and previous.raw_text in ("(", ",", "="))
            or (previous is not None and previous.kind in (TokenKind.OPERATOR, TokenKind.KEYWORD))
        )
        if (
            tok.raw_text == "."
            and leading
            and (
                self._names(_WITH_SHEET) is not None
                or (self._names.with_subject is not None and self._names.with_subject() is not None)
            )
        ):
            return (cast("_Folded", (yield self._sheet_size(True))))
        name = token_name(tok)
        if not name:
            return None
        size = (cast("_Folded", (yield self._sheet_size())))
        if size is not None:
            return size

        # A String local known to hold a number, beside -, *, /, \, ^ or Mod, is
        # that number as a Double: `s ^ 32767` with s = "12" overflows (issue
        # #331, measured in Excel 16.0). Beside + two Strings join.
        def numeric_beside(word: str) -> bool:
            return word in ("-", "*", "/", "\\", "^", "mod")

        lone = (
            self._raw_at(self._index + 1) != "("
            and self._raw_at(self._index + 1) != "."
            and self._raw_at(self._index - 1) != "."
        )
        # A `-` at the start is a sign, not an operator.
        if lone and (numeric_beside(after) or (numeric_beside(before) and (before != "-" or self._index >= 2))):
            spelled = self._names(f'"{name.lower()}')
            if spelled is not None:
                self._index += 1
                return _Typed(spelled.value, "double")
        # `VBA.CInt(...)` and `CInt(...)`.
        callee_index = self._index
        if name.lower() == "vba" and self._raw_at(self._index + 1) == "." and token_name(self._at(self._index + 2)):
            callee_index = self._index + 2
        callee = (token_name(self._at(callee_index)) or "").lower()
        if self._raw_at(callee_index + 1) == "(" and (callee in _CONVERSIONS or callee == "val"):
            close = match_paren_from(self._toks, callee_index + 1)
            if close < 0:
                return None
            start = self._index
            argument = self._toks[callee_index + 2 : close]
            self._index = close + 1
            # A string literal is read as the number it spells (issue #184):
            # `CInt("&H10000")` is CInt(65536), `Val("1e400")` is past a Double.
            if len(argument) == 1 and argument[0].kind is TokenKind.STRING_LITERAL:
                text = string_literal_value(argument[0].raw_text)
                read = _val_of_string(text) if callee == "val" else _number_in_string(text)
                if read == "overflow":
                    return _Overflow(
                        self._span(start, close), f"{argument[0].raw_text} spells a number past the Double range"
                    )
                if not isinstance(read, _Typed) or callee == "val":
                    return read if isinstance(read, _Typed) else None
                return self._convert(callee, read, self._span(start, close), argument[0].raw_text)
            # `CInt(s)` with s a String local known to hold "40000" (issue #407):
            # the lookup answers a `"` key with the number it spells.
            argument_name = token_name(argument[0]) if len(argument) == 1 else None
            held = self._names(f'"{argument_name.lower()}') if argument_name else None
            if held is not None and callee != "val":
                return self._convert(callee, held, self._span(start, close), argument[0].raw_text)
            if callee == "val":
                return None
            inner = (cast("_Folded", (yield self._child(argument, None, self._nesting + 1))))
            if not isinstance(inner, _Typed):
                return inner
            return self._convert(callee, inner, self._span(start, close))
        # `Round(c)` of a Currency comes back a Currency, rounded half to even:
        # Round(922337203685477.5807@) is past the Currency range and raises 6
        # (issue #332, measured in Excel 16.0).
        if callee == "round" and self._raw_at(callee_index + 1) == "(":
            close = match_paren_from(self._toks, callee_index + 1)
            args = [] if close < 0 else split_top_level_token_groups(self._toks, callee_index + 2, ",", close)
            inner = (cast("_Folded", (yield self._child(args[0], None, self._nesting + 1)))) if len(args) == 1 else None
            if isinstance(inner, _Typed) and inner.type == "currency":
                rounded = _round(inner.value)
                if rounded < _RANGES["currency"].min or rounded > _RANGES["currency"].max:
                    return _Overflow(
                        self._span(callee_index, close),
                        f"Round of {_show_number(inner.value)} gives {_show_number(rounded)}, past the Currency range",
                    )
        if self._raw_at(callee_index + 1) == "(" and callee in _RESULT_FUNCTIONS:
            close = match_paren_from(self._toks, callee_index + 1)
            result = (
                None
                if close < 0
                else (cast("_Folded", (yield self._function_result(
                    callee, split_top_level_token_groups(self._toks, callee_index + 2, ",", close)
                ))))
            )
            if result is not None:
                self._index = close + 1
                return result
            return None
        if self._raw_at(self._index + 1) == ".":
            # `Rows.Count`: a two-part member the lookup may know as a constant.
            member = token_name(self._at(self._index + 2))
            following = self._raw_at(self._index + 3)
            if member and following != "." and following != "(":
                known = self._names(f"{name.lower()}.{member.lower()}")
                if known is not None:
                    self._index += 3
                    return known
            return None
        if self._raw_at(self._index + 1) == "(":
            # `F()`, a Function of the module whose result is known (issue #448).
            result = (
                self._names(f"{name.lower()}()")
                if self._raw_at(self._index + 2) == ")" and self._raw_at(self._index - 1) != "."
                else None
            )
            if result is not None:
                self._index += 3
            return result  # None for a call the folder does not know
        known = self._names(name.lower())
        if known is None:
            return None
        self._index += 1
        return known

    def _sheet_size(self, with_subject: bool = False) -> Generator[Any, Any, _Folded]:
        """A size Excel fixes, read from a member chain at the current token (issue
        #411, measured in Excel 16.0): a worksheet's `Rows.Count` (1048576) or
        `Columns.Count` (16384) through ActiveSheet, Application, Worksheets(n) or a
        Worksheet local; `Cells(r, c).Row` and `.Column`; and the Row, Column,
        Count, Rows.Count and Columns.Count of `Range("A1:B2")`, a literal address.
        A Long each. None, and nothing consumed, for any other chain."""
        # `.Rows.Count` inside `With ActiveSheet` reads the With's sheet (issue
        # #411), and inside `With ActiveSheet.Range("A1:A40000")` its range
        # (issue #685).
        subject = (
            self._names.with_subject()
            if with_subject and self._names(_WITH_SHEET) is None and self._names.with_subject is not None
            else None
        )
        segments: list[_ChainSegment] = (
            [] if not with_subject else list(subject) if subject is not None else [_ChainSegment(_WITH_SHEET)]
        )
        i = self._index + 1 if with_subject else self._index
        while True:
            name = token_name(self._at(i))
            if not name:
                return None
            end = i
            args: list[list[VbaToken]] | None = None
            if self._raw_at(i + 1) == "(":
                close = match_paren_from(self._toks, i + 1)
                if close < 0:
                    return None
                args = split_top_level_token_groups(self._toks, i + 2, ",", close)
                end = close
            segments.append(_ChainSegment(name.lower(), args))
            if self._raw_at(end + 1) != ".":
                i = end + 1
                break
            i = end + 2
        if len(segments) < 2 or self._raw_at(i) == "(":
            return None

        def fold(expr: list[VbaToken]) -> Generator[Any, Any, float | None]:
            folded = (cast("_Folded", (yield self._child(expr, self._division_by_zero, self._nesting + 1))))
            return folded.value if isinstance(folded, _Typed) else None

        value = cast("float | None", (yield _sheet_size_of(segments, fold, self._names)))
        if value is None:
            return None
        start = self._index
        self._index = i
        # Range.Count is a Long, and a sheet has 17,179,869,184 cells: `Cells.Count`
        # raises 6 (issue #278, measured in Excel 16.0).
        if value > 2147483647:
            text = "".join(tok.raw_text for tok in self._toks[start:i])
            return _Overflow(
                self._span(start, i - 1),
                f"{text} counts {js_number_to_string(value)} cells, past the Long range Range.Count returns; "
                "CountLarge counts them",
            )
        return _Typed(float(value), "long")

    def _function_result(self, callee: str, args: list[list[VbaToken]]) -> Generator[Any, Any, _Folded]:
        """What a VBA function returns for arguments the folder can read (issue #407,
        measured in Excel 16.0): Sgn is an Integer, -1, 0 or 1; Choose and IIf give
        the argument they pick; Len of `String(n, c)` or of a literal is a Long; Asc
        or AscW of a literal, and AscW of `ChrW(n)`, an Integer. None for anything
        else."""
        # Every argument is evaluated, the ones Choose and IIf do not pick too: an
        # overflow in any of them raises (issue #258).
        values: dict[int, _Typed | None] = {}
        for arg in args:
            folded = (cast("_Folded", (yield self._child(arg, self._division_by_zero, self._nesting + 1)))) if len(arg) > 0 else None
            if isinstance(folded, _Overflow):
                return folded
            values[id(arg)] = folded

        def fold(toks: list[VbaToken] | None) -> Generator[Any, Any, _Typed | None]:
            # Original arguments were already checked above, including unknowns.
            if toks is not None and id(toks) in values:
                return values[id(toks)]
            folded = (
                (cast("_Folded", (yield self._child(toks, self._division_by_zero, self._nesting + 1))))
                if toks is not None and len(toks) > 0
                else None
            )
            return folded if isinstance(folded, _Typed) else None

        def call(toks: list[VbaToken] | None, name: str) -> list[list[VbaToken]] | None:
            head = token_text(toks[0]) if toks is not None and len(toks) >= 3 else ""
            open_at = 2 if toks is not None and len(toks) > 1 and toks[1].raw_text == "$" else 1
            if (
                toks is not None
                and head == name
                and open_at < len(toks)
                and toks[open_at].raw_text == "("
                and match_paren_from(toks, open_at) == len(toks) - 1
            ):
                return split_top_level_token_groups(toks, open_at + 1, ",", len(toks) - 1)
            return None

        if callee == "sgn":
            value = (cast("_Typed | None", (yield fold(args[0])))) if len(args) == 1 else None
            return _Typed(_js_sign(value.value), "integer") if value is not None else None
        if callee == "choose":
            index = (cast("_Typed | None", (yield fold(args[0])))) if args else None
            k = _js_trunc(index.value) if index is not None else None
            if k is not None and k >= 1 and k < len(args):
                return (cast("_Typed | None", (yield fold(args[int(k)]))))
            return None
        if callee == "iif":
            word = token_text(args[0][0]) if len(args) == 3 and len(args[0]) == 1 else ""
            return (cast("_Typed | None", (yield fold(args[1])))) if word == "true" else (cast("_Typed | None", (yield fold(args[2])))) if word == "false" else None
        if callee == "len":
            only = args[0] if len(args) == 1 else None
            if only is not None and len(only) == 1 and only[0].kind is TokenKind.STRING_LITERAL:
                return _Typed(float(utf16_length(string_literal_value(only[0].raw_text))), "long")
            made = call(only, "string")
            count = (cast("_Typed | None", (yield fold(made[0])))) if made is not None and len(made) == 2 else None
            if count is not None and _is_integer(count.value) and count.value >= 0:
                return _Typed(count.value, "long")
            return None
        if callee in ("asc", "ascw"):
            # `Asc("a")` is 97, an Integer (issue #351, measured in Excel 16.0).
            only = args[0] if len(args) == 1 else None
            if only is not None and len(only) == 1 and only[0].kind is TokenKind.STRING_LITERAL:
                text = string_literal_value(only[0].raw_text)
                return _Typed(float(ord(text[0])), "integer") if len(text) > 0 and ord(text[0]) < 128 else None
            if callee == "asc":
                return None
            made = call(only, "chrw")
            code = (cast("_Typed | None", (yield fold(made[0])))) if made is not None and len(made) == 1 else None
            if code is None or not _is_integer(code.value) or code.value < -32768 or code.value > 65535:
                return None
            return _Typed(code.value - 65536 if code.value > 32767 else code.value, "integer")
        return None

    def _current_date(self) -> _Typed | None:
        """`Date` and `Now`, a Date whose serial is today's: past an Integer's 32767
        on any clock set after May 1989, so `i = Now` with i As Integer overflows
        (issue #327, measured in Excel 16.0). Not when a name of the procedure's or
        a call hides them."""
        word = token_text(self._at(self._index))
        following = self._raw_at(self._index + 1)
        if (
            (word != "date" and word != "now")
            or following in ("(", ".", "$")
            or self._raw_at(self._index - 1) == "."
            or self._names(word) is not None
        ):
            return None
        self._index += 1
        today = math.floor((time.time() * 1000 - DATE_EPOCH_MS) / DAY_MS)
        return _Typed(today + 0.5 if word == "now" else float(today), "date")

    def _date_serial(self) -> _Typed | None:
        """`DateSerial(2020, 1, 1)` with whole-number literal arguments: the Date it
        names, with month and day rolling over as VBA rolls them (issue #327,
        measured in Excel 16.0)."""
        if (
            token_text(self._at(self._index)) != "dateserial"
            or self._raw_at(self._index + 1) != "("
            or self._raw_at(self._index - 1) == "."
        ):
            return None
        close = match_paren_from(self._toks, self._index + 1)
        args = [] if close < 0 else split_top_level_token_groups(self._toks, self._index + 2, ",", close)
        parts = [
            parse_vba_integer_literal(arg[0].raw_text)
            if len(arg) == 1 and arg[0].kind is TokenKind.INTEGER_LITERAL
            else None
            for arg in args
        ]
        if len(parts) != 3 or any(part is None for part in parts):
            return None
        year, month, day = (int(part) for part in parts if part is not None)
        if year < 100 or year > 9999:
            return None
        ms = _date_utc(year, month - 1, day)
        self._index = close + 1
        return _Typed(_js_round((ms - DATE_EPOCH_MS) / DAY_MS) if math.isfinite(ms) else math.nan, "date")

    def _convert(self, callee: str, inner: _Typed, span: Span, shown: str | None = None) -> _Folded:
        if shown is None:
            shown = _show_number(inner.value)
        target = _CONVERSIONS[callee]
        if target == "abs":
            # Abs of the smallest Long hands it back unchanged: Abs(CLng(-2147483647
            # - 1)) is -2147483648, where Abs(CInt(-32768)) overflows (issue #218,
            # measured in Excel 16.0).
            if inner.type == "long" and inner.value == _RANGES["long"].min:
                return inner
            magnitude = abs(inner.value)
            if _in_range(magnitude, inner.type):
                return _Typed(magnitude, inner.type)
            return _Overflow(span, f"Abs({_show_number(inner.value)}) does not fit {_label(inner.type)}")
        # Int and Fix of that Date are a Date still past the range, and raise 6;
        # CDate hands it back unchanged (issue #405, measured in Excel 16.0).
        if inner.past_date is not None and target in ("int", "fix"):
            return _Overflow(
                span, f"{'Int' if target == 'int' else 'Fix'} of {inner.past_date} is a Date outside the Date range"
            )
        if inner.past_date is not None and target == "date":
            return inner
        if target in ("int", "fix"):
            whole = _js_floor(inner.value) if target == "int" else _js_trunc(inner.value)
            # A LongLong stays one: `Fix(q) Mod 7` works in LongLong (issue #480).
            return _Typed(whole, inner.type if inner.type in ("byte", "integer", "long", "longlong") else "double")
        if target == "exp":
            power = _js_exp(inner.value)
            if _in_range(power, "double"):
                return _Typed(power, "double")
            return _Overflow(span, f"Exp({_show_number(inner.value)}) exceeds the Double range")
        if target == "decimal":
            # A Decimal holds up to 2^96 - 1: CDec("1E28") runs, CDec("1E30") and
            # CDec(1E+30) overflow (issue #218). A whole one is followed exactly, as
            # a Double too wide to tell its neighbours apart (issue #502).
            if not _decimal_fits(inner.value, shown):
                return _Overflow(span, f"CDec({shown}) does not fit Decimal")
            decimal = _spelled_whole(shown)
            if decimal is None and _is_safe_integer(inner.value):
                decimal = int(inner.value)
            return None if decimal is None else _Typed(float(decimal), "double", decimal=decimal)
        if target in ("hex", "oct"):
            # Hex and Oct take a value that fits a Long, or a LongLong on 64-bit once
            # rounded: Hex(3000000000.5) runs there (issue #332, measured in Excel
            # 16.0); 1E+20 fits neither.
            if _in_range(inner.value, "long") or abs(_round(inner.value)) < 9.2e18:
                return None
            return _Overflow(
                span,
                f"{'Hex' if callee == 'hex' else 'Oct'}({js_number_to_string(inner.value)}) "
                "takes a value outside the Long range",
            )
        if target in ("longlong", "longptr"):
            value = _round(inner.value)
            exact = inner.exact
            if exact is None:
                exact = _spelled_whole(shown)
            if exact is None and _is_safe_integer(value):
                exact = int(value)
            if _in_range(value, "longlong", exact):
                return _Typed(value, "longlong", exact=exact)
            fits = "a LongPtr, whose range is at most a LongLong's" if target == "longptr" else "LongLong"
            return _Overflow(span, f"{_CONVERSION_NAMES[callee]}({shown}) does not fit {fits}")
        stored = inner.value if target in _UNROUNDED_TYPES else _stored_value(inner, target)[0]
        # A conversion to the type the constant already has is folded away:
        # `-CLng(&H80000000)` wraps, `-CLng(-2147483648#)` overflows (issue #235).
        constant = inner.constant and inner.type == target
        if _in_range(stored, target):
            return _Typed(stored, target, constant=constant)
        return _Overflow(span, f"{_CONVERSION_NAMES[callee]}({shown}) does not fit {_label(target)}")

    def _combine(self, left: _Typed, right: _Typed, op: str, start: int, end: int) -> _Folded:
        folded = self._combine_values(left, right, op, start, end)
        # Both halves folded while compiling: so is the result.
        if isinstance(folded, _Typed) and left.constant and right.constant:
            return replace(folded, constant=True)
        return folded

    def _combine_values(self, left: _Typed, right: _Typed, op: str, start: int, end: int) -> _Folded:
        span = self._span(start, end)
        op_shown = "Mod" if op == "mod" else op
        # `\` and Mod convert both operands to a Long before they divide, so a
        # Decimal past the Long range overflows first: `m Mod 0` raises 6, not 11
        # (issue #502, measured in Excel 16.0).
        if op in ("\\", "mod") and left.type != "longlong" and not _in_range(_round(left.value), "long"):
            return _Overflow(
                span, f"{_describe(left)} is outside the Long range that {op_shown} converts its operands to"
            )
        # `1 / 0`, `1 \ 0.4` and `1 Mod False` divide by zero; outside a Const that
        # is division-by-zero's to report.
        divisor = right.value if op == "/" else _round(right.value) if op in ("\\", "mod") else None
        if divisor is not None and divisor == 0:
            if self._division_by_zero is not None:
                self._division_by_zero(span)
            return None
        # A whole Decimal with a whole number: `+`, `-` and `*` stay exact and
        # overflow past 2^96 - 1 (issue #502, measured in Excel 16.0).
        if (left.decimal is not None or right.decimal is not None) and op in ("+", "-", "*"):
            a = left.decimal if left.decimal is not None else _exact_of(left)
            b = right.decimal if right.decimal is not None else _exact_of(right)
            if a is None or b is None:
                return None
            result = a + b if op == "+" else a - b if op == "-" else a * b
            if -_DECIMAL_LIMIT < result < _DECIMAL_LIMIT:
                return _Typed(float(result), "double", decimal=result)
            return _Overflow(span, f"{_describe(left)} {op} {_describe(right)} is past the Decimal range")
        if (left.type == "longlong" or right.type == "longlong") and op not in ("/", "^"):
            return _combine_long_long(left, right, op, span)
        # A Currency sum in ten-thousandths, exactly: a double cannot tell
        # 922337203685477.5807 from one past it (issue #494, measured in Excel 16.0).
        if op in ("+", "-") and _arithmetic_result_type(left.type, right.type, op) == "currency":
            sa = _scaled_of(left)
            sb = _scaled_of(right)
            if sa is not None and sb is not None:
                total = sa + sb if op == "+" else sa - sb
                if -_LONGLONG_LIMIT <= total < _LONGLONG_LIMIT:
                    return _Typed(float(total) / 10000, "currency", scaled=total)
                return _Overflow(span, f"{_describe(left)} {op} {_describe(right)} is past the Currency range")
        if op == "/":
            if right.value == 0:
                return None  # division by zero is another rule's
            type_name = "currency" if left.type == "currency" or right.type == "currency" else "double"
            value = left.value / right.value
        elif op == "^":
            type_name = "double"
            value = _js_pow(left.value, right.value)
            if math.isnan(value) or (left.value == 0 and right.value < 0):
                # 0 ^ -1 raises 5, Invalid procedure call: runtime-value-out-of-range reports it.
                return None
        elif op in ("\\", "mod"):
            type_name = _arithmetic_result_type(left.type, right.type, op)
            a_value = _round(left.value)
            b_value = _round(right.value)
            if b_value == 0:
                return None
            # Both operands are converted to a Long first: `1E10 \ 2` and `7 Mod
            # 922337203685477@` overflow (issue #458, measured in Excel 16.0).
            outside = next(
                (
                    typed
                    for typed, rounded in ((left, a_value), (right, b_value))
                    if not _in_range(rounded, "long")
                ),
                None,
            )
            if outside is not None:
                return _Overflow(
                    span, f"{_describe(outside)} is outside the Long range that {op_shown} converts its operands to"
                )
            # JavaScript's % keeps the dividend's sign, as math.fmod does.
            value = math.fmod(a_value, b_value) if op == "mod" else _js_trunc(a_value / b_value)
        else:
            type_name = _arithmetic_result_type(left.type, right.type, op)
            if op == "+":
                value = left.value + right.value
            elif op == "-":
                value = left.value - right.value
            else:
                value = left.value * right.value
        # `number + Date` and `number - Date` hold a serial past the Date range
        # without raising; only a later use fails (issue #330, measured in Excel
        # 16.0). `Date + number` raises 6.
        if not _in_range(value, type_name) and type_name == "date" and left.type != "date":
            return _Typed(value, type_name, past_date=f"{_describe(left)} {op} {_describe(right)}")
        if (left.variant or right.variant) and op in ("+", "-", "*"):
            widened: str | None = type_name
            while widened is not None and not _in_range(value, widened):
                widened = _VARIANT_WIDENING.get(widened)
            return _Typed(value, widened, variant=True) if widened is not None else None
        if not _in_range(value, type_name):
            if type_name == "date":
                outcome = f"falls {'after 12/31/9999' if value > 0 else 'before 1/1/100'}, outside the Date range"
            else:
                outcome = f"is {_show_number(value)}, outside the {_label(type_name)} range"
            return _Overflow(span, f"{_describe(left)} {op_shown} {_describe(right)} {outcome}")
        return _Typed(value, type_name)


def _spelled_whole(shown: str) -> int | None:
    """A whole number spelled in digits, as a conversion's string or literal argument
    shows it."""
    digits = _SPELLED_WHOLE_RE.fullmatch(shown)
    if digits is None:
        return None
    return -int(digits.group(2)) if digits.group(1) == "-" else int(digits.group(2))


_WHOLE_TYPES = frozenset({"byte", "integer", "long", "longlong"})


def _scaled_of(typed: _Typed) -> int | None:
    """A Currency or whole-number operand in ten-thousandths, exactly, when it has
    that."""
    if typed.scaled is not None:
        return typed.scaled
    return int(typed.value) * 10000 if typed.type in _WHOLE_TYPES and _is_safe_integer(typed.value) else None


def _currency_scaled(text: str) -> int | None:
    """A Currency literal's digits in ten-thousandths: `922337203685477.5807@`."""
    match = _CURRENCY_TEXT_RE.fullmatch(text)
    if match is None or (match.group(1) == "" and not match.group(2)):
        return None
    return int(match.group(1) or "0") * 10000 + int((match.group(2) or "").ljust(4, "0"))


def _exact_of(typed: _Typed) -> int | None:
    """A whole-number operand's exact value, when it has one."""
    if typed.exact is not None:
        return typed.exact
    return int(typed.value) if typed.type in _WHOLE_TYPES and _is_safe_integer(typed.value) else None


def _combine_long_long(left: _Typed, right: _Typed, op: str, span: Span) -> _Folded:
    """`+ - * \\ Mod` with a LongLong operand. With another whole number the result
    is a LongLong, folded exactly; with a Single or Double it is a Double. With a
    Currency or a Date the result type is not modelled, and nothing is judged."""
    other = right.type if left.type == "longlong" else left.type
    if other in ("single", "double"):
        if op in ("\\", "mod"):
            return None
        value = (
            left.value + right.value if op == "+" else left.value - right.value if op == "-" else left.value * right.value
        )
        if _in_range(value, "double"):
            return _Typed(value, "double")
        return _Overflow(span, f"{_describe(left)} {op} {_describe(right)} is outside the Double range")
    a = _exact_of(left)
    b = _exact_of(right)
    if a is None or b is None:
        return None
    if op in ("\\", "mod"):
        if b == 0:
            return None  # division by zero is another rule's
        exact = _trunc_mod(a, b) if op == "mod" else _trunc_div(a, b)
    else:
        exact = a + b if op == "+" else a - b if op == "-" else a * b
    if not _in_range(0, "longlong", exact):
        return _Overflow(
            span,
            f"{_describe(left)} {'Mod' if op == 'mod' else op} {_describe(right)} is {exact}, outside the LongLong range",
        )
    return _Typed(float(exact), "longlong", exact=exact)


def _not_of(operand: _Typed, span: Span) -> _Folded:
    """`Not x`, the bitwise complement in x's type (issue #235, measured in Excel
    16.0): an Integer or a Long stays itself, `Not 32767` is -32768 and `Not 0` is
    -1; a Single, Double, Currency or Date is rounded to a Long first, so `Not
    32768!` is -32769. A Byte stays a Byte."""
    # Not of a Boolean is a Boolean: `b = Not False` stores 255 in a Byte.
    constant = operand.constant
    boolean = operand.boolean
    if operand.type == "byte":
        return _Typed(255 - operand.value, "byte", constant=constant, boolean=boolean)
    if operand.type in ("integer", "long"):
        return _Typed(-operand.value - 1, operand.type, constant=constant, boolean=boolean)
    if operand.type == "longlong":
        exact = _exact_of(operand)
        if exact is None:
            return None
        return _Typed(float(-exact - 1), "longlong", exact=-exact - 1, constant=constant, boolean=boolean)
    whole = _round(operand.value)
    if not _in_range(whole, "long"):
        return _Overflow(
            span, f"Not {_show_number(operand.value)} rounds to {_show_number(whole)}, outside the Long range"
        )
    return _Typed(-whole - 1, "long", constant=constant, boolean=boolean)


def _describe(typed: _Typed) -> str:
    if typed.scaled is not None:
        sign = "-" if typed.scaled < 0 else ""
        digits = str(abs(typed.scaled)).rjust(5, "0")
        fraction = digits[-4:].rstrip("0")
        return f"{sign}{digits[:-4]}{'.' + fraction if fraction else ''} (Currency)"
    if typed.decimal is not None:
        return f"{typed.decimal} (Decimal)"
    if typed.type == "date":
        return f"{_date_text(typed.value)} (Date)"
    shown = str(typed.exact) if typed.exact is not None else _show_number(typed.value)
    return f"{shown} ({_label(typed.type)})"


def _date_text(value: float) -> str:
    """`#m/d/yyyy#` or `#m/d/yyyy h:mm:ss AM#`, as upstream builds it from a JavaScript
    Date. Before serial 0 the fraction counts forward from the day's start too:
    -1.25 is 12/29/1899 6:00:00 AM."""
    day = _js_trunc(value)
    seconds = _js_round(abs(value - day) * 86400)
    civil = _civil_from_ms(DATE_EPOCH_MS + day * DAY_MS)
    date = "NaN/NaN/NaN" if civil is None else f"{civil[1]}/{civil[2]}/{civil[0]}"
    if seconds == 0:
        return f"#{date}#"
    if not math.isfinite(seconds):
        return f"#{date} NaN:NaN:NaN PM#"
    whole = int(seconds)
    hour = whole // 3600
    clock = f"{12 if hour % 12 == 0 else hour % 12}:{str((whole // 60) % 60).rjust(2, '0')}:{str(whole % 60).rjust(2, '0')}"
    return f"#{date} {clock} {'AM' if hour < 12 else 'PM'}#"


def _range_text(type_name: str) -> str:
    """A type's range as the message shows it; a LongLong's ends print exactly."""
    if type_name == "longlong":
        return "-9223372036854775808 to 9223372036854775807"
    bounds = _RANGES[type_name]
    return f"{js_number_to_string(bounds.min)} to {js_number_to_string(bounds.max)}"


def _stored_value(folded: _Typed, target: str) -> tuple[float, int | None]:
    """A folded value converted to a target type, rounded as VBA stores it, with a
    LongLong's exact value."""
    if target in _UNROUNDED_TYPES:
        return folded.value, None
    if folded.boolean and target == "byte":
        return (0.0 if folded.value == 0 else 255.0), None
    value = _round(folded.value)
    # A Date goes into a Byte through an Integer, keeping the low byte: -1 stores
    # 255 and 1000 stores 232; past the Integer range it overflows. CByte of one
    # does not wrap (issue #624, measured in Excel 16.0).
    if folded.type == "date" and target == "byte" and -32768 <= value <= 32767:
        return math.fmod(math.fmod(value, 256) + 256, 256), None
    if target == "longlong" and folded.exact is not None:
        return value, folded.exact
    return value, None


def _show_number(value: float) -> str:
    """A value as VBA would print it: whole numbers plain, huge or fractional ones in
    E notation."""
    if not math.isfinite(value):
        return "a value past the Double maximum" if value > 0 else "a value past the Double minimum"
    magnitude = abs(value)
    if _is_integer(value) and magnitude < 1e15:
        return js_number_to_string(value)
    if magnitude >= 1e15 or (magnitude < 1e-4 and value != 0):
        text = re.sub(r"\.?0+e", "E", _to_exponential(value, 4), count=1)
        text = re.sub(r"e\+?", "E+", text, count=1)
        return text.replace("E+-", "E-", 1)
    return js_number_to_string(value)


def _article(label: str) -> str:
    return "an" if _VOWEL_START_RE.search(label) is not None else "a"


# The VBA functions whose result the folder works out (issue #407).
_RESULT_FUNCTIONS = frozenset({"sgn", "choose", "iif", "len", "asc", "ascw"})

# The type a Variant's arithmetic widens to when a result does not fit (issue #480).
_VARIANT_WIDENING: dict[str, str] = {"byte": "integer", "integer": "long", "long": "double", "single": "double"}

_CONVERSIONS: dict[str, str] = {
    "cbyte": "byte", "cint": "integer", "clng": "long", "csng": "single", "cdbl": "double",
    "ccur": "currency", "cdate": "date", "abs": "abs", "int": "int", "fix": "fix", "exp": "exp",
    "hex": "hex", "oct": "oct", "cdec": "decimal", "clnglng": "longlong", "clngptr": "longptr",
}

# 2^96, one past the largest Decimal.
_DECIMAL_LIMIT = 79228162514264337593543950336


def _decimal_fits(value: float, shown: str) -> bool:
    """Whether a Decimal holds the value. A double cannot tell 2^96 - 1 from 2^96, so
    a whole number spelled in digits is compared exactly."""
    digits = _SPELLED_WHOLE_RE.fullmatch(shown)
    if digits is not None:
        return int(digits.group(2)) < _DECIMAL_LIMIT
    return abs(value) < 7.9228162514264337e28


_CONVERSION_NAMES: dict[str, str] = {
    "cbyte": "CByte", "cint": "CInt", "clng": "CLng", "csng": "CSng", "cdbl": "CDbl",
    "ccur": "CCur", "cdate": "CDate", "clnglng": "CLngLng", "clngptr": "CLngPtr",
}


def _numeric_type_of(as_type: str | None) -> str | None:
    # Upstream tests `normalized in RANGES`, which on a JavaScript object also
    # holds for an inherited name such as `constructor`. That case never reaches
    # this rule upstream: its lexer gives a token spelled Constructor a canonical
    # text that is no string, and parsing the module throws first.
    normalized = normalize_type(as_type)
    if not normalized:
        return None
    if normalized in _RANGES:
        return normalized
    # A LongPtr is a Long or a LongLong by platform: past LongLong's range it
    # overflows on both, and inside it is not judged.
    return "longlong" if normalized == "longptr" else None


@dataclass(slots=True)
class _ResolveFrame:
    # The Const being resolved, or None for a value folded outside one.
    name: str | None
    toks: Sequence[VbaToken]
    # The Consts whose resolution failed while this frame waited on them.
    failed: set[str] = field(default_factory=set)
    # The first pending Const this frame's last fold met unresolved.
    need: str | None = None


class _ConstantResolver:
    """Upstream's recursive `resolve`, run on an explicit stack so a long chain of
    Consts naming each other cannot exhaust Python's recursion limit.

    A frame whose fold meets a pending Const not yet resolved resolves that one
    first, with the same names in progress upstream's `resolving` set would hold,
    then folds again from the start: every lookup before it came from `folded`, so
    the second fold reaches the same point and goes on with the answer. Upstream
    caches no failure, so a failure is remembered only by the frame that asked,
    whose fold stops there, and is worked out afresh anywhere else.
    """

    __slots__ = ("_pending", "_base", "_enum_values", "folded", "_resolving", "_frames", "_lookup")

    def __init__(self, pending: Mapping[str, VbaSymbol], base: Mapping[str, _Typed]) -> None:
        self._pending = pending
        self._base = base
        self._enum_values: dict[str, _Typed] = {}
        self.folded: dict[str, _Typed] = {}
        self._resolving: set[str] = set()
        self._frames: list[_ResolveFrame] = []
        self._lookup = _NameLookup(self._look)

    @property
    def enum_values(self) -> dict[str, _Typed]:
        return self._enum_values

    def _look(self, lower: str) -> _Typed | None:
        if lower not in self._pending:
            enum_value = self._enum_values.get(lower)
            return enum_value if enum_value is not None else self._base.get(lower)
        cached = self.folded.get(lower)
        if cached is not None:
            return cached
        frame = self._frames[-1]
        if lower in self._resolving or lower in frame.failed:
            return None
        if frame.need is None:
            frame.need = lower
        return None

    def evaluate(self, toks: Sequence[VbaToken]) -> _Folded:
        """A value outside any Const, folded over the Consts: an Enum member's."""
        return self._run(_ResolveFrame(None, toks))

    def resolve(self, lower: str) -> _Typed | None:
        """A pending Const's folded value, or None when it does not fold."""
        if lower not in self._pending:
            enum_value = self._enum_values.get(lower)
            return enum_value if enum_value is not None else self._base.get(lower)
        cached = self.folded.get(lower)
        if cached is not None:
            return cached
        if lower in self._resolving:
            return None
        self._resolving.add(lower)
        result = self._run(_ResolveFrame(lower, _const_value_tokens(self._pending[lower].default_raw or "")))
        return result if isinstance(result, _Typed) else None

    def _run(self, root: _ResolveFrame) -> _Folded:
        frames = self._frames
        frames.append(root)
        while True:
            frame = frames[-1]
            frame.need = None
            value = _fold(frame.toks, 0, self._lookup)
            if frame.need is not None:
                need = frame.need
                self._resolving.add(need)
                frames.append(_ResolveFrame(need, _const_value_tokens(self._pending[need].default_raw or "")))
                continue
            frames.pop()
            if frame.name is None:
                return value
            self._resolving.discard(frame.name)
            typed = self._finish(frame.name, value)
            if frame is root:
                return typed
            if typed is None:
                frames[-1].failed.add(frame.name)

    def _finish(self, lower: str, value: _Folded) -> _Typed | None:
        if not isinstance(value, _Typed):
            return None
        symbol = self._pending[lower]
        declared = _numeric_type_of(symbol.as_type)
        kept = _stored_value(value, declared) if declared else None
        # `Const T As Boolean = 1` is True.
        if normalize_type(symbol.as_type) == "boolean":
            truth = _Typed(0.0 if value.value == 0 else -1.0, "integer", constant=True, boolean=True)
            self.folded[lower] = truth
            return truth
        if declared and kept is not None:
            typed = _Typed(kept[0], declared, exact=kept[1], constant=True)
            if not _in_range(kept[0], declared, kept[1]):
                return None
        else:
            typed = value
        self.folded[lower] = typed
        return typed


def _constant_lookup(base: Mapping[str, _Typed], candidates: Sequence[VbaSymbol]) -> Mapping[str, _Typed]:
    """Folded values of the Consts in `candidates`, layered over `base`: a name
    declared in both takes the candidate's value (a procedure's own Const wins over
    the module's), and a candidate's value may refer to a base constant. The module
    and project layer is folded once per pass and each procedure adds only its own
    Consts on top; folding the whole project's constants again for every procedure
    was 15% of a large module's analysis (XLIDE issue #139). The Consts a procedure
    can name are folded with their declared or natural type: `Private Const HOURS As
    Integer = 24` is an Integer 24, and `Const K = 40000` a Long. A Const the folder
    cannot fold is left out."""
    # Later entries shadow earlier ones.
    pending: dict[str, VbaSymbol] = {}
    enums: list[VbaSymbol] = []
    for symbol in candidates:
        if symbol.kind is VbaSymbolKind.CONSTANT and symbol.default_raw is not None:
            pending[symbol.name.lower()] = symbol
        elif symbol.kind is VbaSymbolKind.ENUM:
            enums.append(symbol)
    if not pending and not enums:
        return base
    resolver = _ConstantResolver(pending, base)
    # An Enum member is a Long: its own value, or one more than the member before
    # it, the first 0 (issue #255). `E.eBig` names it too.
    for symbol in enums:
        following: float | None = 0.0
        for member in symbol.children or []:
            if member.kind is not VbaSymbolKind.ENUM_MEMBER:
                continue
            value: _Folded
            if member.default_raw is None:
                value = None if following is None else _Typed(following, "long")
            else:
                value = resolver.evaluate(_const_value_tokens(member.default_raw))
            if not isinstance(value, _Typed) or not _is_integer(value.value) or not _in_range(value.value, "long"):
                following = None
                continue
            typed = _Typed(value.value, "long", constant=True)
            resolver.enum_values[member.name.lower()] = typed
            resolver.enum_values[f"{symbol.name.lower()}.{member.name.lower()}"] = typed
            following = value.value + 1
    out = {**base, **resolver.enum_values}
    for lower in pending:
        typed_value = resolver.resolve(lower)
        # A candidate that did not fold still shadows the base name.
        if typed_value is not None:
            out[lower] = typed_value
        else:
            out.pop(lower, None)
    return out


@lru_cache(maxsize=4096)
def _const_value_tokens(raw: str) -> tuple[VbaToken, ...]:
    # Upstream's rawExpressionTokens for the Const's text, comments dropped.
    # Memoized: the project's Consts are folded again in every module's pass.
    return tuple(_without_comments(raw_expression_tokens(raw)))


_SHEET_ROWS = 1048576
_SHEET_COLUMNS = 16384

# The lookup's name for the subject of the With a statement sits in, when that
# names a sheet.
_WITH_SHEET = "#with"


def _chain_segments(toks: Sequence[VbaToken]) -> list[_ChainSegment] | None:
    """The segments of a member chain written whole: `ActiveSheet`, `Worksheets(1)`,
    `ThisWorkbook.Worksheets(2)`, `a.b(1).c`. None for anything else (upstream's
    chainSegments and chainOf, which are the same walk)."""
    segments: list[_ChainSegment] = []
    i = 0
    while i < len(toks):
        name = token_name(toks[i])
        if not name:
            return None
        end = i
        args: list[list[VbaToken]] | None = None
        if _raw_at(toks, i + 1) == "(":
            close = match_paren_from(toks, i + 1)
            if close < 0:
                return None
            args = split_top_level_token_groups(toks, i + 2, ",", close)
            end = close
        segments.append(_ChainSegment(name.lower(), args))
        if end + 1 == len(toks):
            return segments
        if toks[end + 1].raw_text != ".":
            return None
        i = end + 2
    return None


def _names_sheet(receiver: Sequence[_ChainSegment], names: _NameLookup) -> bool:
    """Whether a member chain names a whole worksheet (issue #411): nothing (the
    active sheet), ActiveSheet, Application, Worksheets(n) from Application,
    ThisWorkbook, ActiveWorkbook or Workbooks(n), or a local As Worksheet."""
    if len(receiver) == 0:
        return True
    first = receiver[0]
    if len(receiver) == 1:
        return (
            (
                first.args is None
                and (
                    first.name in ("activesheet", "application")
                    or (first.name == _WITH_SHEET and names(_WITH_SHEET) is not None)
                )
            )
            or (first.name == "worksheets" and first.args is not None and len(first.args) == 1)
            or (first.args is None and names(f"{first.name}.rows.count") is not None)
        )
    second = receiver[1]
    if len(receiver) == 2 and second.name == "worksheets" and second.args is not None and len(second.args) == 1:
        return (first.args is None and first.name in ("thisworkbook", "activeworkbook", "application")) or (
            first.name == "workbooks" and first.args is not None and len(first.args) == 1
        )
    return (
        len(receiver) == 2
        and first.args is None
        and first.name == "application"
        and second.args is None
        and second.name == "activesheet"
    )


def _whole_sheet_cell_locals(
    source: str,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    names: _NameLookup,
    activity: ConditionalActivityTracker | None,
) -> set[str]:
    """The procedure's Range, Object or Variant locals that one Set gives a whole
    sheet's Cells, `Set r = Cells` or `Set r = ActiveSheet.Cells`, and that nothing
    else assigns or passes whole (issue #278, measured in Excel 16.0: `r.Count`
    raises 6)."""
    proc_sym = procedure_symbol_for(symbols, member)
    candidates = {
        child.name.lower()
        for child in ((proc_sym.children if proc_sym is not None else None) or [])
        if child.kind is VbaSymbolKind.LOCAL_VARIABLE
        and not child.is_array
        and child.visibility is not SymbolVisibility.STATIC
        and (normalize_type(child.as_type) or "variant") in ("range", "object", "variant")
    }
    if not candidates:
        return candidates
    sets: dict[str, list[list[VbaToken]]] = {}
    ruled: set[str] = set()

    def visit(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            toks = statement_tokens(source, span)
            set_target = set_assignment_target(source, span)
            target = set_target[0].lower() if set_target is not None else None
            if target and target in candidates:
                eq = next((k for k, tok in enumerate(toks) if tok.raw_text == "="), -1)
                sets.setdefault(target, []).append(list(toks[eq + 1 :]))
                continue
            bare = bare_assignment_target(source, span)
            if bare is not None and bare[0]:
                ruled.add(bare[0].lower())
            for lower in tracked_locals_named_whole(toks, 0, lambda name: name in candidates, set()):
                ruled.add(lower)

    for_each_statement(member.body, visit, activity)
    out: set[str] = set()
    for lower, values in sets.items():
        chain = _chain_segments(values[0]) if len(values) == 1 and lower not in ruled else None
        last = chain[-1] if chain else None
        if chain is not None and last is not None and last.name == "cells" and last.args is None and _names_sheet(
            chain[:-1], names
        ):
            out.add(lower)
    return out


@dataclass(frozen=True, slots=True)
class _Block:
    row: int
    column: int
    rows: int
    columns: int


def _column_number(letters: str) -> int:
    number = 0
    for ch in letters:
        number = number * 26 + ord(ch) - 64
    return number


def _literal_range(segment: _ChainSegment | None) -> _Block | None:
    """The cells `Range("A1:B2")` names, from a literal A1 address: a cell, a block,
    whole columns or whole rows."""
    if segment is None or segment.name != "range" or segment.args is None or len(segment.args) != 1:
        return None
    arg = _without_comments(segment.args[0])
    if len(arg) != 1 or arg[0].kind is not TokenKind.STRING_LITERAL:
        return None
    text = string_literal_value(arg[0].raw_text).replace("$", "").upper()
    cells = _CELL_RANGE_RE.fullmatch(text)
    if cells is not None:
        r1, c1 = int(cells.group(2)), _column_number(cells.group(1))
        r2, c2 = (int(cells.group(4)), _column_number(cells.group(3))) if cells.group(3) else (r1, c1)
        valid = all(1 <= r <= _SHEET_ROWS for r in (r1, r2)) and all(1 <= c <= _SHEET_COLUMNS for c in (c1, c2))
        return _Block(min(r1, r2), min(c1, c2), abs(r2 - r1) + 1, abs(c2 - c1) + 1) if valid else None
    whole_columns = _WHOLE_COLUMNS_RE.fullmatch(text)
    if whole_columns is not None:
        c1, c2 = _column_number(whole_columns.group(1)), _column_number(whole_columns.group(2))
        if c1 <= _SHEET_COLUMNS and c2 <= _SHEET_COLUMNS:
            return _Block(1, min(c1, c2), _SHEET_ROWS, abs(c2 - c1) + 1)
        return None
    whole_rows = _WHOLE_ROWS_RE.fullmatch(text)
    if whole_rows is not None:
        r1, r2 = int(whole_rows.group(1)), int(whole_rows.group(2))
        if all(1 <= r <= _SHEET_ROWS for r in (r1, r2)):
            return _Block(min(r1, r2), 1, abs(r2 - r1) + 1, _SHEET_COLUMNS)
        return None
    return None


def _sheet_size_of(
    chain: Sequence[_ChainSegment], fold: Callable[[list[VbaToken]], Generator[Any, Any, float | None]], names: _NameLookup
) -> Generator[Any, Any, float | None]:
    """The size a member chain reads, where Excel fixes it (issue #411):
    `ws.Rows.Count`, `Cells(Rows.Count, 1).Row`, `Range("A1:A40000").Rows.Count`."""
    if names("rows.count") is None:
        return None  # not Excel, or a name of the procedure's hides it
    # `cells.Count` with a variable of the code's own named cells.
    if chain[0].name in ("cells", "range", "rows", "columns") and names.declares is not None and names.declares(
        chain[0].name
    ):
        return None
    # `Set r = Cells`, then `r.Count` reads the sheet's Cells (issue #278).
    segments: list[_ChainSegment] = (
        [_ChainSegment("cells"), *chain[1:]]
        if chain[0].args is None and names.whole_sheet_cells is not None and names.whole_sheet_cells(chain[0].name)
        else list(chain)
    )
    n = len(segments)
    last = segments[n - 1]
    before = segments[n - 2]
    if last.args is not None:
        return None
    if last.name == "count" and before.name in ("rows", "columns") and before.args is None:
        receiver = segments[: n - 2]
        block = _literal_range(receiver[-1]) if len(receiver) >= 1 else None
        if block is not None and _names_sheet(receiver[:-1], names):
            return float(block.rows if before.name == "rows" else block.columns)
        if _names_sheet(receiver, names):
            return float(_SHEET_ROWS if before.name == "rows" else _SHEET_COLUMNS)
        return None
    receiver = segments[: n - 2]
    # `Cells.Count` on the whole sheet, `Cells.Cells.Count` too (issue #278).
    if last.name == "count" and before.name == "cells" and before.args is None:
        sheet = receiver[:-1] if receiver and receiver[-1].name == "cells" and receiver[-1].args is None else receiver
        return float(_SHEET_ROWS * _SHEET_COLUMNS) if _names_sheet(sheet, names) else None
    if not _names_sheet(receiver, names):
        return None
    block = _literal_range(before)
    if block is not None:
        if last.name == "count":
            return float(block.rows * block.columns)
        if last.name == "row":
            return float(block.row)
        if last.name == "column":
            return float(block.column)
        return None
    if before.name == "cells" and before.args is not None and len(before.args) == 2 and last.name in ("row", "column"):
        return cast("float | None", (yield fold(before.args[0 if last.name == "row" else 1])))
    return None


_EXCEL_HOST_CONSTANTS: dict[str, _Typed] = {
    "rows.count": _Typed(1048576.0, "long"),
    "columns.count": _Typed(16384.0, "long"),
}


def _host_constant_values(host_model: HostObjectModel | None) -> Mapping[str, _Typed]:
    """Host members whose value is fixed: Excel's `Rows.Count` is 1048576 and
    `Columns.Count` 16384 on every worksheet since Excel 2007 (a Long each). Absent
    model means Excel (XLIDE issue #28)."""
    host_name = host_model.get("hostName") if host_model is not None else None
    if host_name is not None and host_name != "Excel":
        return {}
    return _EXCEL_HOST_CONSTANTS


def check_overflow(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    host_model: HostObjectModel | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    # Module-level Consts: a folded overflow is the compile error.
    module_symbols = [*(project_visible_symbols or []), *(symbols.root.children or [])]
    module_constants = _constant_lookup({}, module_symbols)
    _check_const_declarations(
        source,
        [member for member in mod.members if isinstance(member, VariableGroupNode)],
        module_constants,
        activity,
        push,
    )
    host_values = _host_constant_values(host_model)
    types = module_types(source, mod, activity)
    results = known_function_results(source, mod, activity)
    deftypes = _DEFTYPE_RE.search(source) is not None
    # Names the project declares, which hide a host global of that spelling.
    module_names = {symbol.name.lower() for symbol in module_symbols}
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_procedure(
                source, member, symbols, module_constants, host_values, types, results, deftypes,
                module_names, activity, push,
            )


def _check_procedure(
    source: str,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    module_constants: Mapping[str, _Typed],
    host_values: Mapping[str, _Typed],
    types: ModuleTypes,
    results: Mapping[str, FunctionResult],
    deftypes: bool,
    module_names: set[str],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    proc_sym = procedure_symbol_for(symbols, member)
    children = (proc_sym.children if proc_sym is not None else None) or []
    # A local or parameter hides a module Const or Enum member of its name.
    constants = dict(_constant_lookup(module_constants, children))
    for child in children:
        if child.kind in (VbaSymbolKind.LOCAL_VARIABLE, VbaSymbolKind.PARAMETER):
            constants.pop(child.name.lower(), None)
    # A local declared with no type is a Variant, which holds a Double as a
    # Double: `Dim v: v = 3E9` then `v Mod 2` raises 6 (issue #323, measured in
    # Excel 16.0). A DefType statement types it otherwise.
    untyped = (
        []
        if deftypes
        else [
            child
            for child in children
            if child.kind is VbaSymbolKind.LOCAL_VARIABLE
            and not child.as_type
            and not child.is_array
            and child.name[-1:] != ""
            and (child.name[-1].isascii() and (child.name[-1].isalnum() or child.name[-1] == "_"))
        ]
    )
    base_env = type_environment_for(symbols, member)
    env: Mapping[str, str] = (
        base_env if not untyped else {**base_env, **{child.name.lower(): "Variant" for child in untyped}}
    )
    # What each local holds as the statement being checked is reached: a value
    # stored only in a branch that does not run is not there (issue #565).
    known_at = known_local_literal_values_at(source, member, symbols, activity)
    at: list[BodyNode | None] = [None]
    # Values a straight run of top-level statements has just stored: `i = 32767`
    # followed by `i = i + 1`.
    just_assigned: dict[str, _Typed] = {}

    def lookup(lower: str) -> _Typed | None:
        known = known_at(at[0])
        # `"s`: the number a String local known to hold one spells, for a
        # conversion to read (issue #407).
        if lower.startswith('"'):
            held = known.get(lower[1:])
            if (
                held is not None
                and held.kind == "string"
                and not held.content_mutated
                and normalize_type(env.get(lower[1:])) == "string"
            ):
                read = _number_in_string(str(held.value))
                return read if isinstance(read, _Typed) else None
            return None
        constant = constants.get(lower)
        if constant is not None:
            return constant
        recent = just_assigned.get(lower)
        if recent is not None:
            return recent
        local = known.get(lower)
        type_name = _numeric_type_of(env.get(lower))
        if local is not None and local.kind == "number" and type_name:
            return _Typed(float(local.value), type_name)
        # A Boolean is an Integer in arithmetic, True -1: `n - b` with n the largest
        # Long and b True overflows (issue #331, measured in Excel 16.0). Kept a
        # Boolean, as True is, so a Byte takes it as 255 (issue #624).
        if local is not None and local.kind == "number" and normalize_type(env.get(lower)) == "boolean":
            return _Typed(float(local.value), "integer", boolean=True)
        if "." not in lower:
            # `b = F()` with F a Function of the module returning 300 (issue #448).
            result = function_result_named(lower[:-2] if lower.endswith("()") else lower, results, member, symbols)
            result_type = _numeric_type_of(result.type) if result is not None and result.kind == "number" else None
            if result is not None and result.kind == "number" and result_type:
                return _Typed(float(result.value), result_type)
            return None
        # `ws.Rows.Count` with ws As Worksheet is the sheet's (issue #411).
        head = lower[: lower.index(".")]
        if head in env:
            return host_values.get(lower[len(head) + 1 :]) if normalize_type(env.get(head)) == "worksheet" else None
        return host_values.get(lower)

    names = _NameLookup(lookup, lambda lower: lower in env or lower in module_names)
    sheet_cells = _whole_sheet_cell_locals(source, member, symbols, names, activity)
    names.whole_sheet_cells = lambda lower: lower in sheet_cells
    groups: list[VariableGroupNode] = []
    for_each_variable_group(member.body, groups.append, activity)
    _check_const_declarations(source, groups, constants, activity, push)

    # `t.i = t.i + 1`: a numeric member of a Type value as the target, and `v(2) =
    # 40000` an element of an array in any module (issues #253, #685).
    def member_target(span: Span) -> _AssignmentTarget | None:
        return _member_assignment_target(source, span, symbols, member, types)

    def reached(node: BodyNode | None) -> None:
        at[0] = node

    _check_procedure_body(source, member, env, names, just_assigned, activity, push, member_target, reached)
    at[0] = None
    _check_by_val_arguments(source, member, symbols, names, activity, push)
    _check_accumulating_loops(source, member, env, names, known_at, activity, push)


@dataclass(frozen=True, slots=True)
class _WithState:
    sheet: bool
    subject: list[_ChainSegment] | None


def _check_by_val_arguments(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    names: _NameLookup,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """An argument the folder can read, passed to a ByVal number parameter of a
    procedure of the module that cannot hold it: `TakeI(Rows.Count)` with `ByVal i
    As Integer` converts 1048576 and raises 6 (issue #411, measured in Excel
    16.0). A lone literal is argument-type-mismatch's."""
    signatures = build_module_type_signatures(symbols)
    if not signatures:
        return
    proc_sym = procedure_symbol_for(symbols, proc)
    own = {
        name.lower()
        for name in [
            proc.name,
            *(param.name for param in proc.params),
            *(child.name for child in ((proc_sym.children if proc_sym is not None else None) or [])),
        ]
    }
    # The With each statement sits in, innermost: `TakeI(.Rows.Count)` inside
    # `With ActiveSheet` (issue #685). Keyed by id(node), as the port keys node maps.
    with_of: dict[int, _WithState] = {}
    walk: list[tuple[Iterator[BodyNode], _WithState | None]] = [(iter(proc.body), None)]
    while walk:
        nodes, within = walk.pop()
        for node in nodes:
            if within is not None:
                with_of[id(node)] = within
            inner = within
            if isinstance(node, WithBlockNode):
                header = block_header_statements(source, node)[0]
                toks = statement_tokens(source, header.span) if header is not None else []
                segments = (
                    _chain_segments(toks[1:])
                    if token_text(_at(toks, 0)) == "with" and not (_raw_at(toks, 1) or "").startswith(".")
                    else None
                )
                inner = _WithState(
                    segments is not None and len(segments) > 0 and _names_sheet(segments, names), segments
                )
            if isinstance(node, IfBlockNode):
                for branch in node.branches:
                    walk.append((iter(branch.body), inner))
            else:
                body = getattr(node, "body", None)
                if isinstance(body, list):
                    walk.append((iter(body), inner))

    def names_at(stmt: BodyNode) -> _NameLookup:
        within = with_of.get(id(stmt))
        if within is None:
            return names

        def lookup(lower: str) -> _Typed | None:
            if lower == _WITH_SHEET:
                return _Typed(0.0, "long") if within.sheet else None
            return names(lower)

        return _NameLookup(lookup, names.declares, names.whole_sheet_cells, lambda: within.subject)

    def visit(stmt: LeafStatementNode) -> None:
        stmt_names = names_at(stmt)
        for span in statement_and_branch_spans(stmt):
            toks = statement_tokens(source, span)
            for i, tok in enumerate(toks):
                name = token_name(tok)
                lower = name.lower() if name else None
                signature = (
                    signatures.get(lower) if lower and lower not in own and _raw_at(toks, i - 1) != "." else None
                )
                if signature is None:
                    continue
                # `F(a, b)` in a value, `Call F(a, b)`, or the statement `F a, b`.
                parenthesized = _raw_at(toks, i + 1) == "("
                statement_call = i == 0 and not parenthesized and len(toks) > 1 and toks[1].raw_text != "="
                if not parenthesized and not statement_call:
                    continue
                close = match_paren_from(toks, i + 1) if parenthesized else len(toks)
                args = (
                    [] if close < 0 else split_top_level_token_groups(toks, i + 2 if parenthesized else i + 1, ",", close)
                )
                for k, arg in enumerate(args):
                    param = signature.params[k] if k < len(signature.params) else None
                    value = _without_comments(arg)
                    # ByRef takes an expression, not a variable, as a temporary of
                    # the parameter's type: `TakeIR(Rows.Count)` overflows too
                    # (issue #685, measured in Excel 16.0). A variable is
                    # ByRef-mismatch's.
                    expression = not (len(value) == 1 and token_name(value[0]) is not None)
                    type_name = (
                        _numeric_type_of(param.type_)
                        if param is not None
                        and (param.by_ref is False or expression)
                        and not param.is_array
                        and not param.param_array
                        else None
                    )
                    if (
                        not type_name
                        or param is None
                        or len(value) == 0
                        or any(part.raw_text == ":=" for part in value)
                        or (_literal_typed(value[-1]) is not None and len(value) <= 2)
                    ):
                        continue
                    folded = _fold(value, span.start, stmt_names)
                    if not isinstance(folded, _Typed):
                        continue
                    kept, kept_exact = _stored_value(folded, type_name)
                    if not _in_range(kept, type_name, kept_exact):
                        passing = "ByVal" if param.by_ref is False else "ByRef"
                        push(
                            "arithmeticOverflow",
                            f"Argument '{param.name}' of '{signature.name}' is {passing} {_label(type_name)}, and "
                            f"{''.join(part.raw_text for part in value)} is {_show_number(kept)}, outside its range "
                            f"{_range_text(type_name)}. This will raise Run-time error '6': Overflow.",
                            Span(span.start + value[0].start, span.start + value[-1].end),
                        )

    for_each_statement(proc.body, visit, activity)


# Statement heads after which a loop's pass may not run on.
_LOOP_LEAVING_HEADS = frozenset({"exit", "goto", "gosub", "resume", "return", "on", "stop"})

# The most passes a loop is run for to find its overflow.
_MAX_ACCUMULATED_PASSES = 100000


def _check_accumulating_loops(
    source: str,
    proc: ProcedureNode,
    env: Mapping[str, str],
    names: _NameLookup,
    start_values: Callable[[BodyNode | None], Mapping[str, KnownLocalValue]],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """A whole-number local a loop changes by the same statement every pass, until it
    no longer fits its type (issue #263, measured in Excel 16.0):

     - `For i = 1 To 300: t = t + i` on an Integer t, and `p = p * i` from p = 1 to
       20 on a Long, are run pass by pass from the value the local holds as the
       loop starts.
     - `Do: i = i + 1: Loop Until i > 32767` on an Integer, and `While b <= 255: b
       = b + 1: Wend` on a Byte, cannot end any other way: no value of the type
       passes the exit test.

    The step is the loop's own top-level statement `x = x + k`, `x = x - k` or `x =
    x * k`, k a whole number or the counter; nothing else in the body names x, and
    nothing may leave the pass."""

    def skip(node: BodyNode) -> bool:
        return (activity is not None and activity.is_inactive(node.span)) or not isinstance(
            getattr(node, "body", None), list
        )

    for node in iter_body_nodes(proc.body, skip):
        if isinstance(node, ForBlockNode):
            _accumulate_for(source, node, env, names, start_values(node), activity, push)
        elif isinstance(node, (DoBlockNode, WhileBlockNode)):
            _endless_step(source, node, env, activity, push, start_values(node))


@dataclass(frozen=True, slots=True)
class _LoopStep:
    name: str
    type: str
    op: str  # '+' | '-' | '*'
    # The other operand: a whole number, or the counter (None).
    by: float | None
    span: Span
    text: str


def _loop_step_in(
    source: str,
    body: Sequence[BodyNode],
    env: Mapping[str, str],
    counter: str | None,
    activity: ConditionalActivityTracker | None,
    only: str | None = None,
) -> _LoopStep | None:
    """The body's one step of a whole-number local, when the body is plain
    statements that run every pass."""
    statements: list[tuple[list[VbaToken], Span]] = []
    for node in body:
        if activity is not None and activity.is_inactive(node.span):
            continue
        if not is_leaf_statement(node) or (isinstance(node, StatementNode) and node.single_line_if_branches):
            return None
        toks = statement_tokens(source, node.span)
        head = token_text(_at(toks, 0))
        if (
            head in _LOOP_LEAVING_HEADS
            or (head == "end" and len(toks) == 1)
            or jump_target_label_declaration(source, node.span) is not None
        ):
            return None
        statements.append((toks, node.span))
    step: _LoopStep | None = None
    for toks, span in statements:
        target_name = token_name(_at(toks, 0))
        target = target_name.lower() if target_name else None
        type_name = _numeric_type_of(env.get(target)) if target else None
        if (
            not target
            or _raw_at(toks, 1) != "="
            or not type_name
            or type_name not in _WHOLE_TYPES
            or type_name == "longlong"
            or (only and target != only)
        ):
            continue
        value = toks[2:]
        if len(value) != 3 or value[1].raw_text not in ("+", "-", "*"):
            continue
        op = value[1].raw_text
        if _names_local(value[0], target):
            found, by = _step_operand(value[2], counter)
        elif op != "-" and _names_local(value[2], target):
            found, by = _step_operand(value[0], counter)
        else:
            found, by = False, None
        if not found or step is not None:
            return None  # one step only
        step = _LoopStep(target, type_name, op, by, span, js_trim(source[span.start : span.end]))
    if step is None:
        return None
    # Nothing else names the local.
    mentions = [
        toks
        for toks, _span in statements
        if any(
            (token_name(tok) or "").lower() == step.name and token_name(tok) is not None and _raw_at(toks, k - 1) != "."
            for k, tok in enumerate(toks)
        )
    ]
    return step if len(mentions) == 1 else None


def _names_local(tok: VbaToken | None, lower: str) -> bool:
    name = token_name(tok)
    return name is not None and name.lower() == lower


def _step_operand(tok: VbaToken | None, counter: str | None) -> tuple[bool, float | None]:
    """(found, value) for the other operand of a step: a whole number, or the
    counter as None."""
    if tok is not None and tok.kind is TokenKind.INTEGER_LITERAL:
        parsed = parse_vba_integer_literal(tok.raw_text)
        return (parsed is not None, float(parsed) if parsed is not None else None)
    return (bool(counter) and counter is not None and _names_local(tok, counter), None)


def _step_span(source: str, step: _LoopStep) -> Span:
    return Span(step.span.start, step.span.start + len(_js_trim_end(source[step.span.start : step.span.end])))


def _accumulate_for(
    source: str,
    node: ForBlockNode,
    env: Mapping[str, str],
    names: _NameLookup,
    start: Mapping[str, KnownLocalValue],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    if node.each or not node.control_variable:
        return
    counter = node.control_variable.lower()
    header = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
    eq = next((k for k, tok in enumerate(header) if tok.raw_text == "="), -1)
    to = next((k for k, tok in enumerate(header) if token_text(tok) == "to"), -1)
    step_at = next((k for k, tok in enumerate(header) if token_text(tok) == "step"), -1)

    def fold(toks: Sequence[VbaToken]) -> float | None:
        folded = None if len(toks) == 0 else _fold(toks, node.span.start, names)
        return folded.value if isinstance(folded, _Typed) and _is_integer(folded.value) else None

    first = fold(header[eq + 1 : to]) if eq > 0 and to > eq else None
    limit = fold(header[to + 1 : step_at if step_at > 0 else len(header)]) if to > 0 else None
    increment = fold(header[step_at + 1 :]) if step_at > 0 else 1.0
    if first is None or limit is None or not increment:
        return
    step = _loop_step_in(source, node.body, env, counter, activity)
    initial = start.get(step.name) if step is not None else None
    if step is None or step.name == counter or initial is None or initial.kind != "number" or not isinstance(
        initial.value, (int, float)
    ) or not _is_integer(initial.value):
        return
    value = float(initial.value)
    passes = 0
    c = first
    while (c <= limit) if increment > 0 else (c >= limit):
        passes += 1
        if passes > _MAX_ACCUMULATED_PASSES:
            return
        by = c if step.by is None else step.by
        value = value + by if step.op == "+" else value - by if step.op == "-" else value * by
        if not _in_range(value, step.type):
            if passes == 1:
                return  # the walk into the loop reports its first pass
            label = _label(step.type)
            push(
                "arithmeticOverflow",
                f"On the pass of the For loop where '{node.control_variable}' is {js_number_to_string(c)}, "
                f"'{step.text}' makes '{step.name}' {js_number_to_string(value)}, which does not fit "
                f"{_article(label)} {label}. This will raise Run-time error '6': Overflow.",
                _step_span(source, step),
            )
            return
        c += increment


def _endless_step(
    source: str,
    node: BodyNode,
    env: Mapping[str, str],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    start_values: Mapping[str, KnownLocalValue],
) -> None:
    """`Do ... Loop Until i > 32767` stepping an Integer: no value ends the loop, so
    the step overflows."""
    # The test: `Do While x`, `Do Until x`, `Loop While x`, `Loop Until x`, `While x`.
    header = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
    footer = statement_tokens_after_leading_label(source, block_footer_line_span(source, node.span))
    tests: list[tuple[str, list[VbaToken]]] = []
    for line in (header, footer):
        head = token_text(_at(line, 0))
        word = "while" if head == "while" else token_text(_at(line, 1)) if head in ("do", "loop") else ""
        if word in ("while", "until"):
            tests.append((word, list(line[1 if head == "while" else 2 :])))
    if len(tests) != 1:
        return
    keyword, condition = tests[0]
    if len(condition) not in (3, 4):
        return
    # `x op c`, the constant signed or not.
    cond_name = token_name(condition[0])
    name = cond_name.lower() if cond_name else None
    op = condition[1].raw_text
    negative = len(condition) == 4 and condition[2].raw_text == "-"
    literal = condition[3 if negative else 2]
    raw = parse_vba_integer_literal(literal.raw_text) if literal.kind is TokenKind.INTEGER_LITERAL else None
    if not name or raw is None or op not in ("<", "<=", ">", ">=", "=", "<>"):
        return
    limit = float(-raw if negative else raw)
    start = start_values.get(name)
    # The body is plain statements with no Exit, GoTo or End: _loop_step_in finds
    # the step only there.
    body = getattr(node, "body", None)
    step = _loop_step_in(source, body if isinstance(body, list) else [], env, None, activity, name)
    if step is None or step.op == "*" or step.by is None or step.by <= 0:
        return
    bounds = _RANGES[step.type]

    def holds(w: float) -> bool:
        if op == "<":
            return w < limit
        if op == "<=":
            return w <= limit
        if op == ">":
            return w > limit
        if op == ">=":
            return w >= limit
        if op == "=":
            return w == limit
        return w != limit

    def ends(w: float) -> bool:
        return not holds(w) if keyword == "while" else holds(w)

    # The test changes at the constant, so the range's ends and the values around
    # the constant show whether any value ends the loop.
    probes = [w for w in (bounds.min, bounds.max, limit - 1, limit, limit + 1) if bounds.min <= w <= bounds.max]
    cond_text = " ".join(tok.raw_text for tok in condition)
    article = _article(bounds.label)
    if keyword == "while":
        reason = f"the loop runs while {cond_text}, which {article} {bounds.label} always is"
    else:
        reason = f"the loop ends only when {cond_text}, which {article} {bounds.label} never is"
    if any(ends(w) for w in probes):
        # From a known start the loop reaches only start, start + step, ...: `i =
        # 0: Do Until i < 0: i = i + 1000` never ends before 33000 overflows an
        # Integer (issue #479, measured in Excel 16.0).
        held: float | None
        if start is not None and start.kind == "number" and isinstance(start.value, (int, float)) and _is_integer(
            start.value
        ):
            held = float(start.value)
        elif start is not None and start.kind == "empty":
            held = 0.0
        else:
            held = None
        if held is None or held < bounds.min or held > bounds.max:
            return
        delta = step.by if step.op == "+" else -step.by
        last = math.floor(((bounds.max if delta > 0 else bounds.min) - held) / delta)

        def reached(k: float) -> float | None:
            return held + k * delta if held is not None and 0 <= k <= last else None

        nearest: list[float] = []
        for w in (limit - 1, limit, limit + 1):
            nearest.extend((math.floor((w - held) / delta), math.ceil((w - held) / delta)))
        candidates = [w for w in (reached(k) for k in [0, last, *nearest]) if w is not None]
        # A test after the step reads the stepped value; one before reads the start too.
        foot_test = token_text(_at(footer, 0)) == "loop" and token_text(_at(footer, 1)) in ("while", "until")
        if any(ends(w) and not (foot_test and w == held) for w in candidates):
            return
        held_text = js_number_to_string(held)
        if keyword == "while":
            reason = f"from {held_text}, every value it reaches keeps {cond_text} true"
        else:
            reason = f"from {held_text}, no value it reaches makes {cond_text} true"
    push(
        "arithmeticOverflow",
        f"'{step.name}' is {article} {bounds.label}, and {reason}, so '{step.text}' runs until it does not fit. "
        "This will raise Run-time error '6': Overflow.",
        _step_span(source, step),
    )


def _check_const_declarations(
    source: str,
    groups: Sequence[VariableGroupNode],
    constants: Mapping[str, _Typed],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    lookup = _NameLookup(constants.get)
    for group in groups:
        if not group.is_const or (activity is not None and activity.is_inactive(group.span)):
            continue
        for decl in group.declarations:
            if decl.default_raw is None:
                continue
            toks = statement_tokens(source, decl.span)
            eq = next((i for i, tok in enumerate(toks) if tok.raw_text == "="), -1)
            if eq < 0:
                continue
            value = _without_comments(toks[eq + 1 :])
            divided: list[Span] = []

            def on_division(span: Span, divided: list[Span] = divided) -> None:
                if not divided:
                    divided.append(span)

            folded = _fold(value, decl.span.start, lookup, on_division)
            if divided:
                push(
                    "constEvaluationError",
                    f"Const '{decl.name}' divides by zero while it is evaluated. "
                    "This is a VBE compile error: Division by zero.",
                    divided[0],
                )
                continue
            refused = _string_const_refusal(value, decl.as_type, decl.span.start)
            if refused is not None:
                detail, refused_span, overflow = refused
                push(
                    "constOverflow" if overflow else "constEvaluationError",
                    f"Const '{decl.name}': {detail}. This is a VBE compile error: "
                    f"{'Overflow' if overflow else 'Type mismatch'}.",
                    refused_span,
                )
                continue
            if isinstance(folded, _Overflow):
                push(
                    "constOverflow",
                    f"Const '{decl.name}' overflows while it is evaluated: {folded.detail}. "
                    "This is a VBE compile error: Overflow.",
                    folded.span,
                )
                continue
            declared = _numeric_type_of(decl.as_type)
            if isinstance(folded, _Typed) and declared:
                kept, kept_exact = _stored_value(folded, declared)
                if not _in_range(kept, declared, kept_exact):
                    label = "LongPtr" if normalize_type(decl.as_type) == "longptr" else _label(declared)
                    shown = str(folded.exact) if folded.exact is not None else js_number_to_string(folded.value)
                    outside = "even a LongLong's range" if label == "LongPtr" else "that range"
                    push(
                        "constOverflow",
                        f"Const '{decl.name}' is declared As {label} but its value {shown} is outside {outside}. "
                        "This is a VBE compile error: Overflow.",
                        Span(decl.span.start + value[0].start, decl.span.start + value[-1].end),
                    )


def _string_const_refusal(
    value: Sequence[VbaToken], as_type: str | None, base: int
) -> tuple[str, Span, bool] | None:
    """A Const whose value is a string the declared type cannot take (issue #235,
    measured in Excel 16.0): `As Long = "abc"`, `As Long = ""`, `As Boolean =
    "abc"`, `= -"abc"` and `Not ""` are "Type mismatch", and `As Integer = "40000"`
    is "Overflow". `As Long = "12"`, `As Boolean = "True"` and `-"12"` compile. A
    string with a digit in it is read by the locale, so only one with none is judged
    a mismatch. (detail, span, overflow)."""
    # `"" + 1`, `"abc" < 1`, `1 / ""`: a string with no digit beside a number or a
    # Date in arithmetic or a comparison (issue #367, measured in Excel 16.0). `"2"
    # * 2`, `"a" & 1` and `"a" = "b"` compile.
    if len(value) == 3 and (token_text(value[1]) or value[1].raw_text) in _BINARY_ON_NUMBERS:
        a, b = value[0], value[2]
        text = a if a.kind is TokenKind.STRING_LITERAL else b if b.kind is TokenKind.STRING_LITERAL else None
        other = b if text is a else a
        number_or_date = other.kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL, TokenKind.DATE_LITERAL)
        if text is not None and number_or_date and _ANY_DIGIT_RE.search(string_literal_value(text.raw_text)) is None:
            return (
                f"{text.raw_text} is no number, so {value[1].raw_text} cannot take it beside {other.raw_text}",
                Span(base + value[0].start, base + value[2].end),
                False,
            )
    operator = (
        value[0]
        if len(value) == 2 and (value[0].raw_text == "-" or token_text(value[0]) == "not")
        else None
    )
    literal = _at(value, 1 if operator is not None else 0)
    if len(value) != (2 if operator is not None else 1) or literal is None or literal.kind is not TokenKind.STRING_LITERAL:
        return None
    span = Span(base + value[0].start, base + literal.end)
    text_value = string_literal_value(literal.raw_text)
    has_digit = _ANY_DIGIT_RE.search(text_value) is not None
    if operator is not None:
        if has_digit:
            return None
        return (f"'{operator.raw_text}' cannot work on {literal.raw_text}, which is no number", span, False)
    declared = normalize_type(as_type)
    if declared == "boolean":
        if has_digit or _BOOLEAN_TEXT_RE.fullmatch(text_value) is not None:
            return None
        return (f"{literal.raw_text} is no Boolean", span, False)
    numeric = _numeric_type_of(as_type)
    if not numeric or numeric == "date":
        return None
    if not has_digit:
        label = _label(numeric)
        return (f"{literal.raw_text} is no number, so it cannot be {_article(label)} {label}", span, False)
    read = _number_in_string(text_value)
    if read == "overflow":
        return (f"{literal.raw_text} spells a number past the Double range", span, True)
    if isinstance(read, _Typed) and numeric in _WHOLE_TYPES and not _in_range(_round(read.value), numeric):
        return (f"{literal.raw_text} is outside the {_label(numeric)} range", span, True)
    return None


@dataclass(slots=True)
class _BodyFrame:
    nodes: Iterator[BodyNode]
    loop_touched: frozenset[str]


@dataclass(slots=True)
class _BlockFrame:
    node: BodyNode
    # The bodies to walk, in order, each with the names its enclosing loops change.
    bodies: list[tuple[Sequence[BodyNode], frozenset[str]]]
    entry: dict[str, _Typed]
    touched: set[str]
    after: StatementNode | None
    # Restore what was known on entry before each body: an If's arms, a Select's Cases.
    restore_each: bool
    is_with: bool
    next_body: int = 0


def _check_procedure_body(
    source: str,
    proc: ProcedureNode,
    env: Mapping[str, str],
    outer_names: _NameLookup,
    just_assigned: dict[str, _Typed],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    member_target: Callable[[Span], _AssignmentTarget | None] | None = None,
    reached: Callable[[BodyNode | None], None] = lambda _node: None,
) -> None:
    # Whether each enclosing With names a sheet, innermost last (issue #411), and
    # its subject's chain (issue #685).
    with_sheets: list[bool] = []
    with_subjects: list[list[_ChainSegment] | None] = []

    def lookup(lower: str) -> _Typed | None:
        if lower == _WITH_SHEET:
            return _Typed(0.0, "long") if with_sheets and with_sheets[-1] else None
        return outer_names(lower)

    names = _NameLookup(
        lookup,
        outer_names.declares,
        outer_names.whole_sheet_cells,
        lambda: with_subjects[-1] if with_subjects else None,
    )

    def with_segments(node: BodyNode) -> tuple[list[VbaToken], list[_ChainSegment] | None]:
        header = block_header_statements(source, node)[0]
        toks = statement_tokens(source, header.span) if header is not None else []
        segments = _chain_segments(toks[1:]) if token_text(_at(toks, 0)) == "with" else None
        return toks, segments

    def with_subject_of(node: BodyNode) -> list[_ChainSegment] | None:
        toks, segments = with_segments(node)
        # Only a chain from the sheet or Application: a leading `.` would read an outer With.
        if segments and not (_raw_at(toks, 1) or "").startswith("."):
            return segments
        return None

    def with_names_sheet(node: BodyNode) -> bool:
        _toks, segments = with_segments(node)
        return segments is not None and len(segments) > 0 and _names_sheet(segments, names)

    counters = loop_counters_at(source, proc.body, activity)

    def forget(touched: set[str] | frozenset[str]) -> None:
        for lower in touched:
            _forget_name(just_assigned, lower)

    def restore(entry: Mapping[str, _Typed]) -> None:
        just_assigned.clear()
        just_assigned.update(entry)

    # Upstream's recursive visit, on an explicit stack. `loop_touched`: names an
    # enclosing loop changes, which a block nested in it forgets, since it may run
    # on a later pass.
    stack: list[_BodyFrame | _BlockFrame] = [_BodyFrame(iter(proc.body), frozenset())]
    while stack:
        frame = stack[-1]
        if isinstance(frame, _BlockFrame):
            if frame.next_body < len(frame.bodies):
                body, loop_touched = frame.bodies[frame.next_body]
                frame.next_body += 1
                if frame.restore_each:
                    restore(frame.entry)
                stack.append(_BodyFrame(iter(body), loop_touched))
                continue
            stack.pop()
            if frame.is_with:
                with_sheets.pop()
                with_subjects.pop()
            restore(frame.entry)
            forget(frame.touched)
            if frame.after is not None:
                # A Loop line runs after the body, which may have changed it.
                reached(None)
                _check_statement(source, frame.after.span, env, names, push, member_target)
            continue
        node = next(frame.nodes, None)
        if node is None:
            stack.pop()
            continue
        if activity is not None and activity.is_inactive(node.span):
            continue
        if isinstance(node, ForBlockNode):
            _check_for_counter(source, node, env, names, push)
        node_body = getattr(node, "body", None)
        if isinstance(node_body, list):
            # Its header line is evaluated as it is entered: `For i = 1 To
            # CInt(40000)`, `Select Case CInt(40000)` (issue #233). Its body is
            # entered with what is known then: each If arm and each Case from
            # there, a loop's body as its first pass runs it. After it, only what
            # it never names is still known (issue #237).
            before, after = block_header_statements(source, node)
            if before is not None:
                reached(node)
                _check_statement(source, before.span, env, names, push, member_target)
            # Each If and ElseIf condition, from the state the block is entered
            # with: `If d And 1 Then` (issue #407).
            if isinstance(node, IfBlockNode):
                for header in block_header_leaves(source, node):
                    _check_statement(source, header.span, env, names, push, member_target)
            # Every name a block mentions, its own lines included: `For i = ...`
            # and `If Store(k, n) Then` change what they name as well (issue #237).
            touched: set[str] = set(names_in(source, node.span))
            forget(frame.loop_touched)
            # Its own lines run first: `If Store(k, n) Then` changes n.
            for header in block_header_leaves(source, node):
                forget(set(names_in(source, header.span)))
            entry = dict(just_assigned)
            bodies: list[tuple[Sequence[BodyNode], frozenset[str]]]
            is_with = False
            restore_each = False
            if isinstance(node, IfBlockNode):
                bodies = [(branch.body, frame.loop_touched) for branch in node.branches]
                restore_each = True
            elif isinstance(node, SelectBlockNode):
                bodies = [(arm, frame.loop_touched) for arm in select_arms(source, node.body)]
                restore_each = True
            elif isinstance(node, WithBlockNode):
                with_sheets.append(with_names_sheet(node))
                with_subjects.append(with_subject_of(node))
                bodies = [(node_body, frame.loop_touched)]
                is_with = True
            else:
                bodies = [(node_body, frame.loop_touched | touched if is_loop_block(node) else frame.loop_touched)]
            stack.append(_BlockFrame(node, bodies, entry, touched, after, restore_each, is_with))
            continue
        if not is_leaf_statement(node):
            continue
        reached(node)
        spans = statement_and_branch_spans(node)
        straight_line = len(spans) == 1 and not (
            isinstance(node, StatementNode) and node.single_line_if_branches
        )
        if not straight_line:
            just_assigned.clear()
        for span in spans:
            # A loop counter on its first and last passes as well: `For i = 32760 To
            # 32770` then `CInt(i)` overflows on the last (issue #263).
            stored: list[tuple[str, _Typed] | None] = [None]

            def check(
                values: Mapping[str, float], report: PushFn, span: Span = span, stored: list[tuple[str, _Typed] | None] = stored
            ) -> None:
                if len(values) == 0:
                    stored[0] = _check_statement(source, span, env, names, report, member_target)
                    return

                def pass_lookup(lower: str) -> _Typed | None:
                    value = values.get(lower)
                    type_name = None if value is None else _numeric_type_of(env.get(lower))
                    return _Typed(float(value), type_name) if value is not None and type_name else names(lower)

                _check_statement(source, span, env, _NameLookup(pass_lookup), report, member_target)

            check_each_counter_pass(source, span, counters.get(node), lambda _atom, _counter: None, check, push)
            if not straight_line:
                continue
            # Any other mention of a tracked name (a ByRef pass, a label a GoTo
            # could reach) ends what is known about it.
            toks = statement_tokens(source, span)
            if (
                jump_target_label_declaration(source, span) is not None
                or token_text(_at(toks, first_executable_token_index(toks))) == "gosub"
            ):
                just_assigned.clear()
                continue
            for tok in toks:
                name = token_name(tok)
                if name:
                    _forget_name(just_assigned, name.lower())
            if stored[0] is not None:
                just_assigned[stored[0][0]] = stored[0][1]


def _forget_name(just_assigned: dict[str, _Typed], lower: str) -> None:
    """A name, and every member path under it: `t` forgets `t.i`."""
    just_assigned.pop(lower, None)
    prefix = f"{lower}."
    for key in [key for key in just_assigned if key.startswith(prefix)]:
        del just_assigned[key]


@dataclass(frozen=True, slots=True)
class _AssignmentTarget:
    """What a statement assigns: a name, or a member path or array element with its
    declared type."""

    name: str
    value_tokens: list[VbaToken]
    as_type: str | None = None
    # An array element, whose value is not the name's.
    element: bool = False


def _member_assignment_target(
    source: str,
    span: Span,
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    types: ModuleTypes,
) -> _AssignmentTarget | None:
    """`t.i = value`, `t.v(2) = value` and `v(2) = value` (issue #253): a numeric
    member of a Type value, or an element of an array variable or member."""
    toks = statement_tokens(source, span)
    first = first_executable_token_index(toks)
    at = first + 1 if token_text(_at(toks, first)) == "let" else first
    name = token_name(_at(toks, at))
    lower = name.lower() if name else None
    variable = variable_symbol_in(symbols, proc, lower) if lower else None
    # `ReDim a(2)` with nothing declaring a declares it (issue #685).
    redimmed = (
        _implicit_redim_type(source, proc, symbols, lower)
        if variable is None and lower and _raw_at(toks, at + 1) == "("
        else None
    )
    if ((variable is not None and variable.is_array) or redimmed is not None) and _raw_at(toks, at + 1) == "(":
        close = match_paren_from(toks, at + 1)
        if variable is not None:
            as_type = (
                _EMPTY_ARRAY_SUFFIX_RE.sub("", variable.as_type, count=1) if variable.as_type is not None else None
            )
        else:
            as_type = redimmed[0] if redimmed is not None else None
        if close > 0 and _raw_at(toks, close + 1) == "=":
            return _AssignmentTarget(toks[at].raw_text, list(toks[close + 2 :]), as_type, True)
        return None
    root = variable_root(toks, at, variable, types)
    chain = field_chain(toks, root, types) if root is not None else []
    step = chain[-1] if chain else None
    end = (step.close if step.close is not None else step.at) if step is not None else -1
    if (
        step is None
        or _raw_at(toks, end + 1) != "="
        or (step.field.is_array and step.open is None)
        or (not step.field.is_array and not step.path)
    ):
        return None
    return _AssignmentTarget(
        step.path if step.path is not None else step.display,
        list(toks[end + 2 :]),
        step.field.type_name,
        bool(step.field.is_array),
    )


def _implicit_redim_type(
    source: str, proc: ProcedureNode, symbols: ModuleSymbols, lower: str
) -> tuple[str | None] | None:
    """The element type of an array a ReDim in the procedure declares, nothing else
    declaring it: `ReDim a(2) As Integer`, or `ReDim a(2)` under `DefInt A-Z` (issue
    #685, measured in Excel 16.0). None when no ReDim declares the name; else a
    one-tuple of the type, None for no type."""
    text = source[proc.span.start : proc.span.end]
    for match in _REDIM_LINE_RE.finditer(text):
        clause = match.group(1)
        depth = 0
        start = 0
        items: list[str] = []
        for i in range(len(clause) + 1):
            ch = clause[i] if i < len(clause) else None
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif (ch == "," and depth == 0) or i == len(clause):
                items.append(clause[start:i])
                start = i + 1
        for item in items:
            parts = _REDIM_ITEM_RE.search(item)
            if parts is not None and parts.group(1).lower() == lower:
                return (parts.group(2) if parts.group(2) is not None else def_type_of(symbols, parts.group(1)),)
    return None


def _check_statement(
    source: str,
    span: Span,
    env: Mapping[str, str],
    names: _NameLookup,
    push: PushFn,
    member_target: Callable[[Span], _AssignmentTarget | None] | None = None,
) -> tuple[str, _Typed] | None:
    """The value a bare assignment provably stores, when the rule can tell, as
    (lowercased name, value)."""
    toks = statement_tokens(source, span)
    first = first_executable_token_index(toks)
    head = token_text(_at(toks, first))
    if head in ("const", "dim", "static", "redim"):
        return None
    reported: set[tuple[int, int]] = set()

    def report(folded: _Overflow) -> None:
        key = (folded.span.start, folded.span.end)
        if key not in reported:
            reported.add(key)
            push("arithmeticOverflow", f"{folded.detail}. This will raise Run-time error '6': Overflow.", folded.span)

    stored: tuple[str, _Typed] | None = None
    plain = bare_assignment_target(source, span)
    bare = (
        _AssignmentTarget(plain[0], list(plain[2]))
        if plain is not None
        else member_target(span)
        if member_target is not None
        else None
    )
    if bare is not None:
        value = _without_comments(bare.value_tokens)
        folded = _fold(value, span.start, names)
        lower = bare.name.lower()
        declared = bare.as_type if bare.as_type is not None else env.get(lower)
        target = _numeric_type_of(declared)
        if isinstance(folded, _Overflow):
            report(folded)
        elif isinstance(folded, _Typed) and folded.past_date is not None and target == "date":
            # A Date local holds it without raising (issue #405).
            if not bare.element:
                stored = (lower, folded)
        elif (
            isinstance(folded, _Typed)
            and not target
            and not bare.element
            and lower in env
            and (normalize_type(declared) or "variant") == "variant"
        ):
            # A Variant holds the value with its own type: `v = 2147483648#` then `v
            # Mod 7` converts a Double to Long and raises 6 (issue #480).
            stored = (lower, replace(folded, variant=True))
        elif isinstance(folded, _Typed) and target:
            kept, kept_exact = _stored_value(folded, target)
            if not _in_range(kept, target, kept_exact):
                shown = str(kept_exact) if kept_exact is not None else _show_number(kept)
                rounded = f" ({_show_number(folded.value)} rounds to {_show_number(kept)})" if kept != folded.value else ""
                label = (
                    "LongPtr, which holds no more than a LongLong"
                    if normalize_type(declared) == "longptr"
                    else _label(target)
                )
                into = f"an element of '{bare.name}'" if bare.element else f"'{bare.name}'"
                # A Variant's number past the Date range is a Type mismatch, where a
                # typed one's is an Overflow (issue #329, measured in Excel 16.0).
                variant_date = target == "date" and folded.variant
                error = "'13': Type mismatch" if variant_date else "'6': Overflow"
                push(
                    "assignmentTypeMismatch" if variant_date else "arithmeticOverflow",
                    f"Assignment to {into} stores {shown}{rounded} in {_article(label)} {label}, whose range is "
                    f"{_range_text(target)}. This will raise Run-time error {error}.",
                    Span(span.start + value[0].start, span.start + value[-1].end),
                )
                return None
            if not bare.element:
                stored = (lower, _Typed(kept, target, exact=kept_exact))
    # A condition is evaluated whole: `If d And 1 Then` with d past the Long range
    # converts it for And and raises 6 (issue #407, measured in Excel 16.0), as `x
    # = d And 1` does.
    do_condition = head in ("do", "loop") and token_text(_at(toks, first + 1)) in ("while", "until")
    condition_from = first + 1 if head in ("if", "elseif", "while") else first + 2 if do_condition else -1
    if condition_from > 0:
        then = next((k for k in range(condition_from, len(toks)) if token_text(toks[k]) == "then"), -1)
        condition = _without_comments(toks[condition_from : len(toks) if then < 0 else then])
        condition_folded = _fold(condition, span.start, names) if condition else None
        if isinstance(condition_folded, _Overflow):
            report(condition_folded)

    # Every other part the statement evaluates on its own: a call's arguments, an
    # operand of & or a comparison, an array index, a conversion anywhere (issue
    # #232). `Main = CStr(CInt(40000))` and `IIf(True, 0, CInt(40000))` overflow as
    # `Main = CInt(40000)` does.
    def report_past_date(folded: _Typed, at: Span) -> None:
        key = (at.start, at.end)
        if key not in reported:
            reported.add(key)
            push(
                "runtimeArgumentValue",
                f"{folded.past_date} is {_show_number(folded.value)}, a Date outside the Date range that VBA holds "
                "without raising; reading it here as text or as a date raises Run-time error '5': Invalid procedure "
                "call or argument.",
                at,
            )

    _check_parts(toks, span.start, names, report, report_past_date)
    return stored


# Keywords that stand between separately evaluated parts: the operators the folder
# does not fold, and the statement words around an expression. A keyword that is
# an operand or a function, `Date` or `CInt`, is not one: `Date - 32767% - 2%` is
# Date arithmetic, and its tail is no Integer sum.
_PART_KEYWORDS = frozenset(
    {
        "and", "or", "xor", "eqv", "imp", "not", "like", "is", "if", "then", "else", "elseif",
        "to", "step", "print", "call", "set", "let", "return", "while", "until", "case", "with",
        "select", "each", "in", "goto", "gosub", "on",
    }
)

_FOLDED_OPERATORS = frozenset({"+", "-", "*", "/", "\\", "^", "(", ")", ".", "!"})


def _ends_part(tok: VbaToken) -> bool:
    """Tokens that end one separately evaluated part of an expression: a comma,
    `:=`, every operator the folder does not fold (&, comparisons), and the keywords
    above, so `Debug.Print 200 * 200` leaves `200 * 200` to fold."""
    if tok.raw_text in (",", ";", ":="):
        return True
    if tok.kind is TokenKind.OPERATOR:
        return tok.raw_text not in _FOLDED_OPERATORS
    return tok.kind is TokenKind.KEYWORD and token_text(tok) in _PART_KEYWORDS


def _check_parts(
    toks: Sequence[VbaToken],
    base: int,
    names: _NameLookup,
    report: Callable[[_Overflow], None],
    report_past_date: Callable[[_Typed, Span], None] | None = None,
) -> None:
    """Folds each part of `toks` that is evaluated on its own; a part that does not
    fold is searched for parenthesized parts that do, so an argument nested at any
    depth is reached. Upstream recurses per parenthesis; each level here is a
    generator on an explicit stack, run in the same order."""
    stack: list[Generator[tuple[list[VbaToken], str, int], None, None]] = [
        _parts_level(toks, base, names, report, report_past_date, None, 0)
    ]
    while stack:
        nested = next(stack[-1], None)
        if nested is None:
            stack.pop()
            continue
        stack.append(_parts_level(nested[0], base, names, report, report_past_date, nested[1], nested[2]))


def _parts_level(
    toks: Sequence[VbaToken],
    base: int,
    names: _NameLookup,
    report: Callable[[_Overflow], None],
    report_past_date: Callable[[_Typed, Span], None] | None,
    callee: str | None,
    nesting: int,
) -> Generator[tuple[list[VbaToken], str, int], None, None]:
    """One level of upstream's checkParts; yields each parenthesized part to check
    before it goes on."""
    if nesting >= MAX_EXPRESSION_DEPTH:
        return
    ends: list[int] = []
    depth = 0
    for i, tok in enumerate(toks):
        raw = tok.raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
        elif depth == 0 and _ends_part(tok):
            ends.append(i)
    ends.append(len(toks))
    start = 0
    for to in ends:
        part_start = start
        piece = _without_comments(toks[start:to])
        start = to + 1
        if not piece:
            continue
        folded = _fold(piece, base, names)
        if isinstance(folded, _Overflow):
            report(folded)
            continue
        if folded is not None:
            if folded.past_date is not None and _reads_past_date(toks, part_start, to, callee):
                if report_past_date is not None:
                    report_past_date(folded, Span(base + piece[0].start, base + piece[-1].end))
            continue
        i = 0
        while i < len(piece):
            if piece[i].raw_text != "(":
                i += 1
                continue
            close = match_paren_from(piece, i)
            if close < 0:
                break
            # A conversion folds with its call: `CInt(40000)` overflows though 40000 does not.
            inner_callee = token_text(_at(piece, i - 1))
            if inner_callee in _CONVERSIONS and is_bare_or_vba_qualified_intrinsic_call(piece, i - 1):
                call_start = i - 3 if _raw_at(piece, i - 2) == "." else i - 1
                call = _fold(piece[call_start : close + 1], base, names)
                if isinstance(call, _Overflow):
                    report(call)
                    i = close + 1
                    continue
            yield (piece[i + 1 : close], inner_callee, nesting + 1)
            i = close + 1


# The built-ins that raise 5 when given a Date past the Date range (issue #405,
# measured in Excel 16.0). CDate, CVar, CLng, CDbl, CCur, IsDate, IsNumeric,
# TypeName, VarType and a comparison take it and run.
_PAST_DATE_READERS = frozenset(
    {
        "cstr", "str", "format", "formatdatetime", "year", "month", "day", "weekday", "hour", "minute",
        "second", "dateadd", "datepart", "datediff", "datevalue", "ucase", "mid", "left", "val", "instr",
        "replace",
    }
)


def _reads_past_date(toks: Sequence[VbaToken], start: int, to: int, callee: str | None) -> bool:
    """Whether the part `toks[start..to)` is read as text or as a date: an argument
    of one of the built-ins above, an operand of `&`, or what `Debug.Print` prints."""
    if callee is not None and callee in _PAST_DATE_READERS:
        return True
    if _raw_at(toks, start - 1) == "&" or _raw_at(toks, to) == "&":
        return True
    return (
        token_text(_at(toks, start - 1)) == "print"
        and _raw_at(toks, start - 2) == "."
        and token_text(_at(toks, start - 3)) == "debug"
    )


def _check_for_counter(
    source: str,
    node: ForBlockNode,
    env: Mapping[str, str],
    names: _NameLookup,
    push: PushFn,
) -> None:
    """`For i = 1 To 32767` with i an Integer: the counter is incremented past its
    last value before the exit test, and the increment overflows (measured: `To
    32767` raises, `To 32766` runs; `For b = 0 To 255` raises for a Byte)."""
    if node.each or not node.control_variable:
        return
    type_name = _numeric_type_of(env.get(node.control_variable.lower()))
    if not type_name or type_name in _UNROUNDED_TYPES:
        return
    # The header with any lines a ` _` continues it onto (issue #289).
    header = block_header_line_span(source, node.span)
    toks = statement_tokens_after_leading_label(source, header)
    to = next((i for i, tok in enumerate(toks) if token_text(tok) == "to"), -1)
    if to < 0:
        return
    step = next((i for i, tok in enumerate(toks) if token_text(tok) == "step"), -1)
    limit_toks = _without_comments(toks[to + 1 : step if step > 0 else len(toks)])
    bounds = _RANGES[type_name]
    # The For line converts its start, limit and step to the counter's type as it
    # runs, before the first pass (issue #263, measured in Excel 16.0): `For b = 5
    # To 3 Step -1` on a Byte raises 6 there, and so does a limit past the type,
    # whatever Exit For the body holds.
    if type_name != "longlong":
        eq_at = next((i for i, tok in enumerate(toks) if tok.raw_text == "="), -1)
        parts: list[tuple[str, Sequence[VbaToken]]] = [
            ("start", toks[eq_at + 1 : to] if eq_at > 0 else []),
            ("limit", limit_toks),
            ("step", toks[step + 1 :] if step > 0 else []),
        ]
        for which, part in parts:
            value = _without_comments(part)
            folded = None if not value else _fold(value, header.start, names)
            if not isinstance(folded, _Typed) or _in_range(_round(folded.value), type_name):
                continue
            push(
                "forCounterOverflow",
                f"Counter '{node.control_variable}' is {bounds.label}, and the For line converts its {which} "
                f"{js_number_to_string(folded.value)} to {bounds.label} as it starts, which does not fit. "
                "This will raise Run-time error '6': Overflow.",
                Span(header.start + value[0].start, header.start + value[-1].end),
            )
            return
    limit = _fold(limit_toks, header.start, names)
    step_value = (
        _fold(_without_comments(toks[step + 1 :]), header.start, names) if step > 0 else _Typed(1.0, "integer")
    )
    if (
        not isinstance(limit, _Typed)
        or not isinstance(step_value, _Typed)
        or step_value.value == 0
        or not _is_integer(step_value.value)
    ):
        return
    # The counter's last value is the start plus whole steps: `For i = 1 To 32766
    # Step 2` ends at 32765 and runs (measured). A step of 1 or -1 ends at the
    # limit whatever the start.
    eq = next((i for i, tok in enumerate(toks) if tok.raw_text == "="), -1)
    start = _fold(_without_comments(toks[eq + 1 : to]), header.start, names) if eq > 0 else None
    if abs(step_value.value) == 1:
        last = limit.value
    elif isinstance(start, _Typed) and _is_integer(start.value):
        passes = _js_floor((limit.value - start.value) / step_value.value)
        if passes < 0:
            return  # the loop body never runs and the counter stays at the start
        last = start.value + passes * step_value.value
    else:
        return
    last_shown = js_number_to_string(last)
    if step_value.value > 0:
        overflows = last + step_value.value > bounds.max
    else:
        overflows = last + step_value.value < bounds.min
    if type_name == "longlong":
        # Exactly: a double cannot tell 2^63 - 1 from 2^63 (issue #232).
        exact_last = _exact_of(limit) if abs(step_value.value) == 1 else None
        exact_step = _exact_of(step_value)
        if exact_last is None or exact_step is None:
            return
        overflows = not _in_range(0, "longlong", exact_last + exact_step)
        last_shown = str(exact_last)
    if not overflows or body_may_leave_loop(source, node.body):
        return
    push(
        "forCounterOverflow",
        f"Counter '{node.control_variable}' is {bounds.label}; after its last pass at {last_shown} the loop adds "
        f"{js_number_to_string(step_value.value)}, which does not fit. This will raise Run-time error '6': Overflow.",
        Span(header.start + limit_toks[0].start, header.start + limit_toks[-1].end),
    )


def _fold(
    toks: Sequence[VbaToken],
    base: int,
    names: _NameLookup,
    division_by_zero: Callable[[Span], None] | None = None,
) -> _Folded:
    """A fresh fold of `toks`."""
    try:
        return run_expression(_TypedFolder(toks, base, names, division_by_zero).fold())
    except RecursionError:
        # Name lookups supplied by callers can recurse independently of the
        # folder. Keep an unavailable name from dropping the module's findings.
        return None


def _without_comments(toks: Sequence[VbaToken]) -> list[VbaToken]:
    return [tok for tok in toks if tok.kind is not TokenKind.COMMENT]


def _paren_matches(toks: Sequence[VbaToken]) -> list[int]:
    """For each '(' the index match_paren_from gives it, -1 elsewhere, in one pass."""
    matches = [-1] * len(toks)
    opened: list[int] = []
    for i, tok in enumerate(toks):
        if tok.raw_text == "(":
            opened.append(i)
        elif tok.raw_text == ")" and opened:
            matches[opened.pop()] = i
    return matches


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


# -- JavaScript number semantics ---------------------------------------------------


def _is_integer(value: float) -> bool:
    """Number.isInteger: finite and whole."""
    return float(value).is_integer()


def _is_safe_integer(value: float) -> bool:
    """Number.isSafeInteger."""
    return _is_integer(value) and abs(value) <= 2**53 - 1


def _js_floor(value: float) -> float:
    return float(math.floor(value)) if math.isfinite(value) else value


def _js_trunc(value: float) -> float:
    return float(math.trunc(value)) if math.isfinite(value) else value


def _js_round(value: float) -> float:
    """Math.round: half up, toward +Infinity."""
    return float(math.floor(value + 0.5)) if math.isfinite(value) else value


def _js_sign(value: float) -> float:
    if math.isnan(value):
        return value
    return 1.0 if value > 0 else -1.0 if value < 0 else value


def _js_exp(value: float) -> float:
    try:
        return math.exp(value)
    except OverflowError:
        return math.inf


def _js_pow(base: float, exponent: float) -> float:
    """Math.pow for the finite operands the folder has: NaN where JavaScript gives
    NaN and an infinity where it overflows, where math.pow raises instead."""
    if math.isnan(exponent):
        return math.nan
    if exponent == 0:
        return 1.0
    if math.isnan(base):
        return math.nan
    if base == 0:
        return 0.0 if exponent > 0 else math.inf
    if base < 0 and not _is_integer(exponent):
        return math.nan
    try:
        return math.pow(base, exponent)
    except OverflowError:
        odd = _is_integer(exponent) and math.fmod(exponent, 2) != 0
        return -math.inf if base < 0 and odd else math.inf


def _js_trim_end(text: str) -> str:
    """String.prototype.trimEnd."""
    return text.rstrip(JS_WHITESPACE)


def _trunc_div(a: int, b: int) -> int:
    """A bigint's `/`: the quotient truncated toward zero."""
    quotient = abs(a) // abs(b)
    return quotient if (a < 0) == (b < 0) else -quotient


def _trunc_mod(a: int, b: int) -> int:
    """A bigint's `%`: the remainder with the dividend's sign."""
    return a - b * _trunc_div(a, b)


# The most a JavaScript Date holds either side of 1970, in milliseconds.
_JS_TIME_LIMIT = 8.64e15


def _days_from_civil(year: int, month: int, day: int) -> int:
    """Days from 1970-01-01 to a proleptic Gregorian date (month 1 to 12)."""
    year -= 1 if month <= 2 else 0
    era = (year if year >= 0 else year - 399) // 400
    yoe = year - era * 400
    doy = (153 * (month + (-3 if month > 2 else 9)) + 2) // 5 + day - 1
    doe = yoe * 365 + yoe // 4 - yoe // 100 + doy
    return era * 146097 + doe - 719468


def _date_utc(year: int, month_index: int, day: int) -> float:
    """Date.UTC(year, monthIndex, day) for a year of 100 or more: months and days
    past their ranges roll over, and a time past a Date's limits is NaN."""
    whole_year = year + month_index // 12
    month = month_index % 12 + 1
    ms = (_days_from_civil(whole_year, month, 1) + day - 1) * 86400000
    return float(ms) if abs(ms) <= _JS_TIME_LIMIT else math.nan


def _civil_from_ms(ms: float) -> tuple[int, int, int] | None:
    """The UTC (year, month, day) of a time, as a JavaScript Date reads it; None
    where the Date would be invalid."""
    if not math.isfinite(ms) or abs(ms) > _JS_TIME_LIMIT:
        return None
    days = math.floor(ms / 86400000) + 719468
    era = (days if days >= 0 else days - 146096) // 146097
    doe = days - era * 146097
    yoe = (doe - doe // 1460 + doe // 36524 - doe // 146096) // 365
    doy = doe - (365 * yoe + yoe // 4 - yoe // 100)
    mp = (5 * doy + 2) // 153
    day = doy - (153 * mp + 2) // 5 + 1
    month = mp + 3 if mp < 10 else mp - 9
    year = yoe + era * 400 + (1 if month <= 2 else 0)
    return year, month, day


_EXACT = Context(prec=64, rounding=ROUND_HALF_UP)


def _to_exponential(value: float, fraction_digits: int) -> str:
    """Number.prototype.toExponential for a finite value: the exact binary value
    rounded to fraction_digits + 1 significant digits, a tie going up as the
    specification asks (Python's own e-format sends a tie to even)."""
    sign = "-" if value < 0 else ""
    magnitude = Decimal(abs(float(value)))
    if magnitude == 0:
        coefficient = "0" * (fraction_digits + 1)
        exponent = 0
    else:
        exponent = magnitude.adjusted()
        rounded = magnitude.quantize(Decimal(1).scaleb(exponent - fraction_digits, _EXACT), context=_EXACT)
        if rounded.adjusted() > exponent:
            exponent += 1
            rounded = magnitude.quantize(Decimal(1).scaleb(exponent - fraction_digits, _EXACT), context=_EXACT)
        coefficient = "".join(str(digit) for digit in rounded.as_tuple().digits)
    body = coefficient[0] + ("." + coefficient[1:] if fraction_digits else "")
    return f"{sign}{body}e{'+' if exponent >= 0 else '-'}{abs(exponent)}"
