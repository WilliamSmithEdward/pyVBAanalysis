"""Ported from xlide_vscode/src/analyzer/diagnostics/conditionValue.ts.

What an If condition evaluates to when the code makes it plain (XLIDE issue
#273).

A guard whose outcome is known decides which arm runs: with c never Set,
`If c Is Nothing Then Exit Function` always leaves, so nothing after it runs;
with n = 0, `If n > 0 Then` never runs its arm. The rules that track values and
object state ask this module, and skip what cannot run.

Only what is certain is answered: a comparison of two known numbers, two strings
that differ or match under any Option Compare, `Is Nothing` on an object whose
state is known, IsNumeric and Len of a known string, and Not, And and Or over
those. Anything else is None.

Numbers are JavaScript numbers here, Python floats: a whole number from a
literal is held as a float, so arithmetic and its overflow follow upstream's.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING, TypeGuard, Union

from ..constants.date_literal import date_literal_serial
from ..constants.integer_constant_expression import bankers_round, parse_vba_integer_literal
from ..js_compat import js_number, js_number_to_string, utf16_length
from ..lexer.token_helpers import match_paren_from, split_top_level_token_groups, token_name
from ..lexer.token_helpers import token_word as token_text
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.expression_limits import MAX_EXPRESSION_DEPTH
from .string_conversion import is_invalid_date_string, numeric_string_verdict, val_prefix_value

if TYPE_CHECKING:
    from .known_string_calls import ModuleCompare


def _string_literal_value(raw: str) -> str:
    """A string literal's text, its doubled quotes undone."""
    if raw.startswith('"'):
        raw = raw[1:]
    if raw.endswith('"'):
        raw = raw[:-1]
    return raw.replace('""', '"')


@dataclass(frozen=True, slots=True)
class ConditionFacts:
    """What a condition may read of the names it mentions. Only `value` is
    required; each other reader is optional, as upstream's interface has it."""

    # The number or string a name is known to hold, by lowercased name.
    value: Callable[[str], float | str | None]
    # Whether an object name is known Nothing (True) or known set (False).
    is_nothing: Callable[[str], bool | None] | None = None
    # The whole numbers a name is known to lie between, both included:
    # `Second(Now) + 1000` (issue #565).
    range: Callable[[str], tuple[float, float] | None] | None = None
    # Whether a name is known to hold Null (issue #664).
    is_null: Callable[[str], bool | None] | None = None
    # The module's Option Compare, which decides strings that differ only in
    # case and how strings order (issue #686). Without it only what holds under
    # either is decided.
    compare: ModuleCompare | None = None
    # A local's declared type, lowercased: "long", "variant", "long()" (issue #691).
    type_of: Callable[[str], str | None] | None = None
    # A fixed one-dimension array's bounds (issue #691).
    bounds: Callable[[str], tuple[float, float] | None] | None = None
    # Whether a Variant is known Empty (True) or known to hold a value (False) (issue #691).
    is_empty: Callable[[str], bool | None] | None = None
    # The Count of a Collection known to hold this many (issue #691).
    count: Callable[[str], float | None] | None = None


# VarType's answer for each declared type (issue #691, measured in Excel 16.0).
_VAR_TYPES: dict[str, int] = {
    "integer": 2, "long": 3, "single": 4, "double": 5, "currency": 6, "date": 7, "string": 8,
    "boolean": 11, "byte": 17, "longlong": 20,
}

# TypeName's answer for each declared type.
_TYPE_NAMES: dict[str, str] = {
    "integer": "Integer", "long": "Long", "single": "Single", "double": "Double",
    "currency": "Currency", "date": "Date", "string": "String", "boolean": "Boolean",
    "byte": "Byte", "longlong": "LongLong",
}

# The VbVarType constants a guard compares with.
_VB_CONSTANTS: dict[str, int] = {
    "vbempty": 0, "vbnull": 1, "vbinteger": 2, "vblong": 3, "vbsingle": 4, "vbdouble": 5,
    "vbcurrency": 6, "vbdate": 7, "vbstring": 8, "vbobject": 9, "vberror": 10, "vbboolean": 11,
    "vbvariant": 12, "vbdecimal": 14, "vbbyte": 17, "vblonglong": 20, "vbarray": 8192,
}


@dataclass(frozen=True, slots=True)
class _Nothing:
    """An object name, for `Is Nothing`: known Nothing, known set, or not known."""

    nothing: bool | None


@dataclass(frozen=True, slots=True)
class _Range:
    """What a name holds when only its range is known."""

    range: tuple[float, float]


_Value = Union[float, str, bool, None]
_Operand = Union[float, str, bool, None, _Nothing, _Range]

_RELATIONAL = frozenset({"=", "<>", "<", ">", "<=", ">="})


def _is_number(value: object) -> TypeGuard[float]:
    """`typeof value === 'number'`: a bool is not one."""
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _num(value: object) -> _Value:
    """A value as the parser holds it: a number as a float."""
    if isinstance(value, int) and not isinstance(value, bool):
        return float(value)
    if isinstance(value, (float, str, bool)):
        return value
    return None


def _is_integer(value: object) -> TypeGuard[float]:
    """Number.isInteger."""
    if not _is_number(value):
        return False
    number = float(value)
    return math.isfinite(number) and number == math.floor(number)


def _js_floor(value: float) -> float:
    return float(math.floor(value)) if math.isfinite(value) else value


def _js_trunc(value: float) -> float:
    return float(math.trunc(value)) if math.isfinite(value) else value


def _js_round(value: float) -> float:
    """bankersRound, which leaves Infinity and NaN as they are."""
    return float(bankers_round(value)) if math.isfinite(value) else value


def _js_sign(value: float) -> float:
    if math.isnan(value):
        return value
    return 1.0 if value > 0 else -1.0 if value < 0 else 0.0


def _js_pow(base: float, exponent: float) -> float:
    """`base ** exponent` as JavaScript computes it: NaN, not an error, where
    there is no real result, and Infinity past the double range."""
    if math.isnan(exponent):
        return math.nan
    if math.isinf(exponent) and abs(base) == 1:
        return math.nan
    try:
        return math.pow(base, exponent)
    except OverflowError:
        return math.inf
    except ValueError:
        return math.nan


def condition_value(toks: Sequence[VbaToken], facts: ConditionFacts) -> bool | None:
    """The value of `toks` as a condition: True, False, or None when it is not certain."""
    parser = _ConditionParser([tok for tok in toks if tok.kind is not TokenKind.COMMENT], facts)
    try:
        value = parser.or_expr()
    except RecursionError:
        # Port-only: each nesting level costs a dozen Python frames, so a
        # condition nested past Python's limit is not known here, where
        # upstream reads it up to MAX_EXPRESSION_DEPTH.
        return None
    if not parser.done():
        return None
    return _truth(value)


def number_value(toks: Sequence[VbaToken], facts: ConditionFacts) -> float | None:
    """The number `toks` evaluates to where the facts make it plain: `Len(s)` with
    s known (issue #685)."""
    parser = _ConditionParser([tok for tok in toks if tok.kind is not TokenKind.COMMENT], facts)
    try:
        value = parser.or_expr()
    except RecursionError:
        return None
    return value if parser.done() and _is_number(value) else None


def if_condition_tokens(toks: Sequence[VbaToken]) -> list[VbaToken] | None:
    """The tokens between `If` (or `ElseIf`) and `Then` in a statement or header."""
    head = token_text(toks[0] if toks else None)
    if head != "if" and head != "elseif":
        return None
    depth = 0
    for i in range(1, len(toks)):
        raw = toks[i].raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
        elif depth == 0 and token_text(toks[i]) == "then":
            return list(toks[1:i])
    return None


def _truth(value: _Operand) -> bool | None:
    if isinstance(value, bool):
        return value
    if _is_number(value):
        return value != 0
    return None


class _NotACall:
    """NOT_A_CALL: the word is not called here."""


_NOT_A_CALL = _NotACall()

# The built-ins a condition reads of known values (issue #691).
_BUILTIN_CALLS: frozenset[str] = frozenset(
    {
        "abs", "sgn", "int", "fix", "round", "val", "cstr", "len", "isempty", "isarray", "isdate",
        "typename", "vartype", "iif", "ubound", "lbound",
    }
)

# The string functions a condition reads of known strings (issue #686).
_STRING_CALLS: frozenset[str] = frozenset({"lcase", "ucase", "instr", "strcomp", "replace"})

_FLOAT_SUFFIX_RE = re.compile(r"[!#@]$")
_D_EXPONENT_RE = re.compile(r"[dD]")


def _float_literal_value(raw: str) -> float:
    """`Number(raw.replace(/[!#@]$/, '').replace(/[dD]/, 'e'))`."""
    if raw and raw[-1] in "!#@":
        raw = raw[:-1]
    return js_number(_D_EXPONENT_RE.sub("e", raw, count=1))


class _ConditionParser:
    __slots__ = ("_toks", "_facts", "_nesting", "_index")

    def __init__(self, toks: Sequence[VbaToken], facts: ConditionFacts, nesting: int = 0) -> None:
        self._toks = toks
        self._facts = facts
        self._nesting = nesting
        self._index = 0

    def _at(self, i: int) -> VbaToken | None:
        return self._toks[i] if 0 <= i < len(self._toks) else None

    def _raw(self, i: int) -> str | None:
        tok = self._at(i)
        return tok.raw_text if tok is not None else None

    def _descend(self, read: Callable[[], _Operand]) -> _Operand:
        """Share the recovery budget across calls, parentheses and prefix operators."""
        if self._nesting + 1 >= MAX_EXPRESSION_DEPTH:
            self._index = len(self._toks) + 1
            return None
        self._nesting += 1
        try:
            return read()
        finally:
            self._nesting -= 1

    def done(self) -> bool:
        return self._index >= len(self._toks)

    def _word(self) -> str:
        return token_text(self._at(self._index))

    def or_expr(self) -> _Value:
        if self._nesting >= MAX_EXPRESSION_DEPTH:
            return None
        value = self._and_expr()
        while self._word() == "or":
            self._index += 1
            right = self._and_expr()
            a = _truth(value)
            b = _truth(right)
            value = True if a is True or b is True else False if a is False and b is False else None
        return value

    def _and_expr(self) -> _Value:
        value = self._not_expr()
        while self._word() == "and":
            self._index += 1
            right = self._not_expr()
            a = _truth(value)
            b = _truth(right)
            value = False if a is False or b is False else True if a is True and b is True else None
        return value

    def _not_expr(self) -> _Value:
        if self._word() == "not":
            self._index += 1
            value = _truth(self._descend(self._not_expr))
            return None if value is None else not value
        return self._comparison()

    def _comparison(self) -> _Value:
        left = self.concat()
        op = self._raw(self._index)
        if self._word() == "is" and token_text(self._at(self._index + 1)) == "nothing":
            self._index += 2
            return left.nothing if isinstance(left, _Nothing) else None
        # `s Like "b*"`, by the module's compare mode (issue #686).
        if self._word() == "like":
            self._index += 1
            pattern = self.concat()
            if isinstance(left, str) and isinstance(pattern, str):
                return _like_match(left, pattern, _text_compare(self._facts.compare, left + pattern))
            return None
        if op is None or op not in _RELATIONAL:
            return None if isinstance(left, (_Nothing, _Range)) else left
        self._index += 1
        right = self.concat()
        # A range against a number: decided where every value in it agrees.
        left_range = _range_of(left)
        right_range = _range_of(right)
        if isinstance(left, _Range) or isinstance(right, _Range):
            if left_range is None or right_range is None:
                return None
            lo, hi = left_range
            rlo, rhi = right_range
            corners = [_compare(lo, rlo, op), _compare(lo, rhi, op), _compare(hi, rlo, op), _compare(hi, rhi, op)]
            # `=` and `<>` hold at an inner value the corners miss.
            if (op == "=" or op == "<>") and not (hi < rlo or rhi < lo) and not (lo == hi and rlo == rhi):
                return None
            return True if all(corners) else False if not any(corners) else None
        if isinstance(left, (_Nothing, _Range)) or isinstance(right, (_Nothing, _Range)) or left is None or right is None:
            return None
        if _is_number(left) and _is_number(right):
            return _compare(float(left), float(right), op)
        if isinstance(left, str) and isinstance(right, str):
            order = _string_order(left, right, self._facts.compare)
            if order is not None:
                return _compare(order, 0, op)
            if op != "=" and op != "<>":
                return None
            # Equal both ways or different both ways, whatever Option Compare says.
            if left == right:
                return op == "="
            if left.lower() != right.lower():
                return op == "<>"
        return None

    def _sum(self) -> _Operand:
        """Whole numbers joined by +, -, * or Mod: `b Mod 2 = 0` with b known
        (issue #565). A result past the Long range, which may overflow, and any
        operand that is not a known whole number make it None."""
        value = self._modulo()
        op = self._raw(self._index)
        while op == "+" or op == "-":
            self._index += 1
            value = _whole_arithmetic(value, self._modulo(), op)
            op = self._raw(self._index)
        return value

    def concat(self) -> _Operand:
        """`&` joins strings, and whole numbers and Booleans as VBA spells them:
        `"a" & k` (issue #691)."""
        if self._nesting >= MAX_EXPRESSION_DEPTH:
            return None
        value = self._sum()
        while self._raw(self._index) == "&":
            self._index += 1
            right = self._sum()
            a = _spelled(value)
            b = _spelled(right)
            value = a + b if a is not None and b is not None else None
        return value

    def _modulo(self) -> _Operand:
        value = self._integer_division()
        while self._word() == "mod":
            self._index += 1
            value = _whole_arithmetic(value, self._integer_division(), "mod")
        return value

    def _integer_division(self) -> _Operand:
        """`k \\ 2`: each side rounded half to even, then divided toward zero (issue #691)."""
        value = self._product()
        while self._raw(self._index) == "\\":
            self._index += 1
            right = self._product()
            if _is_number(value) and _is_number(right) and _js_round(float(right)) != 0:
                value = _js_trunc(_js_round(float(value)) / _js_round(float(right))) + 0
            else:
                value = None
        return value

    def _product(self) -> _Operand:
        value = self._unary()
        op = self._raw(self._index)
        while op == "*" or op == "/":
            self._index += 1
            right = self._unary()
            if op == "*":
                value = _whole_arithmetic(value, right, "*")
            elif _is_number(value) and _is_number(right) and right != 0:
                value = float(value) / float(right)
            else:
                value = None
            op = self._raw(self._index)
        return value

    def _unary(self) -> _Operand:
        """`-d`, binding looser than `^`: -2 ^ 2 is -4."""
        if self._raw(self._index) == "-":
            self._index += 1
            value = self._descend(self._unary)
            return -float(value) + 0 if _is_number(value) else None
        return self._power()

    def _power(self) -> _Operand:
        """`k ^ 2` (issue #691)."""
        value = self._operand()
        while self._raw(self._index) == "^":
            self._index += 1
            right = self._operand()
            result = (
                _js_pow(float(value), float(right))
                if _is_number(value) and _is_number(right)
                else math.nan
            )
            value = result if math.isfinite(result) else None
        return value

    def _operand(self) -> _Operand:
        """A literal, a known name, an object name (for Is Nothing), IsNumeric(...),
        or a parenthesized condition."""
        tok = self._at(self._index)
        if tok is None:
            return None
        facts = self._facts
        if tok.raw_text == "(":
            self._index += 1
            inner = self._descend(self.or_expr)
            if self._raw(self._index) != ")":
                self._index = len(self._toks) + 1
                return None
            self._index += 1
            return inner
        following = self._at(self._index + 1)
        if tok.raw_text == "-" and following is not None and following.kind is TokenKind.INTEGER_LITERAL:
            self._index += 2
            literal = parse_vba_integer_literal(following.raw_text)
            return None if literal is None else -float(literal)
        self._index += 1
        if tok.kind is TokenKind.INTEGER_LITERAL:
            literal = parse_vba_integer_literal(tok.raw_text)
            return None if literal is None else float(literal)
        if tok.kind is TokenKind.STRING_LITERAL:
            return _string_literal_value(tok.raw_text)
        # `x > #1/1/2010#`: a date as its serial number (issue #691).
        if tok.kind is TokenKind.DATE_LITERAL:
            return _num(date_literal_serial(tok.raw_text))
        if tok.kind is TokenKind.FLOAT_LITERAL:
            number = _float_literal_value(tok.raw_text)
            return number if math.isfinite(number) else None
        word = token_text(tok)
        if word == "true" or word == "false":
            return word == "true"
        next_raw = self._raw(self._index)
        if facts.value(word) is None and word in _VB_CONSTANTS and next_raw != "(" and next_raw != ".":
            return float(_VB_CONSTANTS[word])
        # Built-ins of known values (issue #691).
        if word in _BUILTIN_CALLS and next_raw == "(":
            called = self._builtin_call(word)
            if not isinstance(called, _NotACall):
                return called
            next_raw = self._raw(self._index)
        # `c.Count` of a Collection known to hold so many (issue #691).
        if (
            next_raw == "."
            and token_text(self._at(self._index + 1)) == "count"
            and self._raw(self._index + 2) != "("
            and self._raw(self._index + 2) != "."
        ):
            count = facts.count((token_name(tok) or "").lower()) if facts.count is not None else None
            if count is not None:
                self._index += 2
                return float(count)
        # IsNull of a local a straight line set to Null, or to a number or a
        # string (issue #664).
        if word == "isnull" and next_raw == "(" and self._raw(self._index + 2) == ")":
            arg = self._at(self._index + 1)
            self._index += 3
            if token_text(arg) == "null":
                return True
            lower = (token_name(arg) or "").lower()
            held = facts.is_null(lower) if facts.is_null is not None else None
            if held is not None:
                return held
            return False if facts.value(lower) is not None else None
        if word == "isnumeric" and next_raw == "(" and self._raw(self._index + 2) == ")":
            arg = self._at(self._index + 1)
            assert arg is not None
            self._index += 3
            value = (
                _string_literal_value(arg.raw_text)
                if arg.kind is TokenKind.STRING_LITERAL
                else _num(facts.value((token_name(arg) or "").lower()))
            )
            if _is_number(value):
                return True
            if not isinstance(value, str):
                return None
            verdict = numeric_string_verdict(value)
            return False if verdict.kind == "invalid" else True if verdict.value is not None else None
        # Len or LenB of a known String: `If Len(s) > 1 Then` with s empty (issue #577).
        if (word == "len" or word == "lenb") and next_raw == "(" and self._raw(self._index + 2) == ")":
            arg = self._at(self._index + 1)
            assert arg is not None
            self._index += 3
            lower = (token_name(arg) or "").lower()
            # Len of a Long is its size, 4, whatever string it was given (issue #685).
            is_string = arg.kind is TokenKind.STRING_LITERAL
            declared = None if is_string or facts.type_of is None else facts.type_of(lower)
            value = _string_literal_value(arg.raw_text) if is_string else _num(facts.value(lower))
            if isinstance(value, str) and (declared is None or declared == "string" or declared == "variant"):
                return float(utf16_length(value) * (2 if word == "lenb" else 1))
            return None
        # LCase, UCase, InStr, StrComp and Replace of known strings (issue #686).
        if word in _STRING_CALLS:
            called = self._string_call(word)
            if not isinstance(called, _NotACall):
                return called
        found_name = token_name(tok)
        name = found_name.lower() if found_name is not None else None
        next_raw = self._raw(self._index)
        if not name or next_raw == "(" or next_raw == ".":
            self._index = len(self._toks) + 1
            return None
        if self._word() == "is":
            return _Nothing(facts.is_nothing(name) if facts.is_nothing is not None else None)
        known = _num(facts.value(name))
        held_range = facts.range(name) if known is None and facts.range is not None else None
        return _Range(held_range) if held_range is not None else known

    def _string_call(self, word: str) -> _Value | _NotACall:
        """A call of a string function at the token before the index, its arguments
        each a condition operand. NOT_A_CALL where no `(` follows; None where an
        argument or the result is not known."""
        open_at = self._index
        if self._raw(open_at) == "$":
            open_at += 1
        if self._raw(open_at) != "(" or self._raw(self._index - 2) == ".":
            return _NOT_A_CALL
        close = match_paren_from(self._toks, open_at)
        if close < 0:
            return _NOT_A_CALL
        args = [self._argument_value(arg) for arg in split_top_level_token_groups(self._toks, open_at + 1, ",", close)]
        self._index = close + 1
        mode = self._facts.compare

        def compare_arg(value: _Value) -> bool | None:
            if value is None or not _is_number(value):
                return None
            return True if value == 1 else False if value == 0 else None

        def arg(k: int) -> _Value:
            return args[k] if k < len(args) else None

        if word == "lcase" or word == "ucase":
            s = arg(0)
            if len(args) == 1 and isinstance(s, str) and s.isascii():
                return s.lower() if word == "lcase" else s.upper()
            return None
        if word == "instr":
            padded: list[_Value] = args if len(args) >= 3 else [1.0, *args]
            start = padded[0] if len(padded) > 0 else None
            s1 = padded[1] if len(padded) > 1 else None
            s2 = padded[2] if len(padded) > 2 else None
            how = padded[3] if len(padded) > 3 else None
            if (
                not _is_integer(start)
                or float(start) < 1
                or not isinstance(s1, str)
                or not isinstance(s2, str)
                or len(args) > 4
            ):
                return None
            text = compare_arg(how) if len(args) == 4 else _text_compare(mode, s1 + s2)
            if text is None or (text and not (s1 + s2).isascii()):
                return None
            haystack = _utf16_units(s1.lower() if text else s1)
            needle = _utf16_units(s2.lower() if text else s2)
            begin = int(float(start))
            if len(haystack) == 0:
                return 0.0
            if begin > len(haystack):
                return float(begin) if len(needle) == 0 else 0.0
            return float(haystack.find(needle, begin - 1) + 1)
        if word == "strcomp":
            a, b, how = arg(0), arg(1), arg(2)
            if not isinstance(a, str) or not isinstance(b, str) or len(args) > 3:
                return None
            if len(args) == 3:
                text = compare_arg(how)
                order = _string_order(a, b, "database" if text is None else "text" if text else "binary")
            else:
                order = _string_order(a, b, mode)
            return None if order is None else float(order)
        if word == "replace":
            expression, find, replacement, start, count, how = (arg(k) for k in range(6))
            if (
                not isinstance(expression, str)
                or not isinstance(find, str)
                or not isinstance(replacement, str)
                or (start is not None and not (_is_number(start) and start == 1))
                or (count is not None and not (_is_number(count) and count == -1))
                or len(args) > 6
            ):
                return None
            if len(find) == 0:
                return expression
            text = compare_arg(how) if len(args) == 6 else _text_compare(mode, expression + find)
            if text is None or (text and not (expression + find).isascii()):
                return None
            if text:
                return re.sub(re.escape(find), lambda _match: replacement, expression, flags=re.IGNORECASE | re.ASCII)
            return expression.replace(find, replacement)
        return None

    def _builtin_call(self, word: str) -> _Value | _NotACall:
        """A built-in whose arguments are known (issue #691): Abs, Sgn, Int, Fix,
        Round, Val, CStr and Len of values; IsEmpty, IsArray, TypeName, VarType,
        UBound and LBound of a local whose declaration or value says; IsDate of a
        string no locale reads as a date; IIf of a known condition."""
        open_at = self._index
        close = match_paren_from(self._toks, open_at)
        if close < 0 or self._raw(open_at - 2) == ".":
            return _NOT_A_CALL
        groups = [
            [tok for tok in group if tok.kind is not TokenKind.COMMENT]
            for group in split_top_level_token_groups(self._toks, open_at + 1, ",", close)
        ]
        # A one-token Len is read below, as it was.
        if word == "len" and len(groups[0]) == 1:
            return _NOT_A_CALL
        self._index = close + 1
        facts = self._facts

        def name(k: int) -> str | None:
            if k < len(groups) and len(groups[k]) == 1 and groups[k][0].kind is TokenKind.IDENTIFIER:
                found = token_name(groups[k][0])
                return found.lower() if found is not None else None
            return None

        def value(k: int) -> _Value:
            return self._argument_value(groups[k]) if k < len(groups) else None

        def number(k: int) -> float | None:
            v = value(k)
            return float(v) if _is_number(v) else None

        def typed(k: int) -> str | None:
            lower = name(k)
            return facts.type_of(lower) if lower and facts.type_of is not None else None

        def empty(k: int) -> bool | None:
            lower = name(k)
            held = facts.is_empty(lower) if lower and facts.is_empty is not None else None
            if held is not None:
                return held
            declared = typed(k)
            return False if declared is not None and declared != "variant" else None

        one = len(groups) == 1
        if word == "abs":
            n = number(0)
            return abs(n) if one and n is not None else None
        if word == "sgn":
            n = number(0)
            return _js_sign(n) + 0 if one and n is not None else None
        if word == "int" or word == "fix":
            n = number(0)
            if not one or n is None:
                return None
            return (_js_floor(n) if word == "int" else _js_trunc(n)) + 0
        if word == "round":
            n = number(0)
            return _js_round(n) + 0 if one and n is not None else None
        if word == "val":
            s = value(0)
            return val_prefix_value(s) if one and isinstance(s, str) else None
        if word == "cstr":
            return _spelled(value(0)) if one else None
        if word == "len":
            s = value(0)
            return float(utf16_length(s)) if one and isinstance(s, str) else None
        if word == "isempty":
            return empty(0) if one else None
        if word == "isarray":
            declared = typed(0)
            return None if not one or declared is None or declared == "variant" else declared.endswith("()")
        if word == "isdate":
            s = value(0)
            return False if one and isinstance(s, str) and is_invalid_date_string(s) else None
        if word == "typename" or word == "vartype":
            declared = typed(0)
            if not one or declared is None:
                return None
            base = declared[:-2] if declared.endswith("()") else declared
            array = declared.endswith("()")
            if base == "variant":
                if not array and empty(0) is True:
                    return "Empty" if word == "typename" else 0.0
                return 8204.0 if array and word == "vartype" else None
            if word == "typename":
                type_name = _TYPE_NAMES.get(base)
                return None if type_name is None else f"{type_name}{'()' if array else ''}"
            var_type = _VAR_TYPES.get(base)
            return None if var_type is None else float(var_type + (8192 if array else 0))
        if word == "iif":
            if len(groups) != 3:
                return None
            parser = _ConditionParser(groups[0], facts, self._nesting + 1)
            condition = _truth(parser.or_expr())
            return None if condition is None or not parser.done() else value(1 if condition else 2)
        if word == "ubound" or word == "lbound":
            lower = name(0)
            bounds = facts.bounds(lower) if lower and facts.bounds is not None else None
            dimension = number(1) if len(groups) == 2 else 1.0
            if bounds is not None and dimension == 1 and len(groups) <= 2:
                return float(bounds[1 if word == "ubound" else 0])
            return None
        return None

    def _argument_value(self, arg: Sequence[VbaToken]) -> _Value:
        """An argument's value: a condition operand, and vbBinaryCompare or
        vbTextCompare as 0 or 1."""
        toks = [tok for tok in arg if tok.kind is not TokenKind.COMMENT]
        word = token_text(toks[0]) if len(toks) == 1 else ""
        if word == "vbbinarycompare" or word == "vbtextcompare":
            return 1.0 if word == "vbtextcompare" else 0.0
        parser = _ConditionParser(toks, self._facts, self._nesting + 1)
        value = parser.concat()
        return value if parser.done() and not isinstance(value, (_Nothing, _Range)) else None


def _spelled(value: _Operand) -> str | None:
    """A value as `&` and CStr spell it: whole numbers, Booleans and strings; None
    for a fraction, which the locale spells."""
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "True" if value else "False"
    return js_number_to_string(float(value)) if _is_integer(value) else None


def _utf16_units(text: str) -> str:
    """The text as JavaScript indexes it: an astral character as its two surrogates."""
    if text.isascii() or all(ord(ch) <= 0xFFFF for ch in text):
        return text
    out: list[str] = []
    for ch in text:
        code = ord(ch)
        if code > 0xFFFF:
            code -= 0x10000
            out.append(chr(0xD800 + (code >> 10)))
            out.append(chr(0xDC00 + (code & 0x3FF)))
        else:
            out.append(ch)
    return "".join(out)


def _text_compare(mode: ModuleCompare | None, sample: str) -> bool | None:
    """Whether the module compares these strings as text (True) or binary (False),
    or None where that is not known: no Option Compare given to the rule, Option
    Compare Database, or text with a character past ASCII, which the locale folds."""
    if mode == "binary":
        return False
    return True if mode == "text" and sample.isascii() else None


_ALPHANUMERIC_RE = re.compile(r"[A-Za-z0-9]*")


def _string_order(a: str, b: str, mode: ModuleCompare | None) -> int | None:
    """How two strings order, -1, 0 or 1, where the module's compare mode makes it
    certain (issue #686, measured in Excel 16.0): Binary orders ASCII by character
    code, so "A" < "a"; Text ignores case, and its locale sort is followed only for
    letters and digits. Equal strings are 0 in any mode."""
    if a == b:
        return 0
    if mode == "binary" and (a + b).isascii():
        return -1 if a < b else 1
    if mode == "text" and _ALPHANUMERIC_RE.fullmatch(a + b):
        x = a.lower()
        y = b.lower()
        return 0 if x == y else -1 if x < y else 1
    return None


_LETTER_RE = re.compile(r"[A-Za-z]")
# `(.)-(.)` in JavaScript: `.` stops at every line terminator.
_RANGE_RE = re.compile(r"([^\n\r\u2028\u2029])-([^\n\r\u2028\u2029])")


def _class_atom(ch: str) -> str:
    return f"\\u{ord(ch):04x}"


def _like_match(subject: str, pattern: str, text: bool | None) -> bool | None:
    """Whether `subject Like pattern` holds: `*`, `?`, `#` and `[list]` with `!` and
    ranges. None where the compare mode decides a letter and is not known, or the
    pattern is not one this reads."""
    if text is None and _LETTER_RE.search(subject + pattern):
        return None
    # Upstream builds a JavaScript regular expression, which reads UTF-16 code
    # units; the same units are matched here.
    pattern = _utf16_units(pattern)
    source: list[str] = []
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "*":
            source.append("[\\s\\S]*")
        elif ch == "?":
            source.append("[\\s\\S]")
        elif ch == "#":
            source.append("[0-9]")
        elif ch == "[":
            end = pattern.find("]", i + 1)
            if end < 0:
                return None
            members = pattern[i + 1 : end]
            negate = members.startswith("!")
            if negate:
                members = members[1:]
            if any(found.group(1) > found.group(2) for found in _RANGE_RE.finditer(members)):
                return None
            source.append(_character_class(members, negate))
            i = end
        else:
            source.append(re.escape(ch))
        i += 1
    try:
        compiled = re.compile("".join(source), re.IGNORECASE | re.ASCII if text else 0)
    except re.error:
        return None
    return compiled.fullmatch(_utf16_units(subject)) is not None


def _character_class(members: str, negate: bool) -> str:
    """The class upstream writes as `[${negate ? '^' : ''}${escaped members}]`,
    read the way JavaScript reads it: each member one character, and `a-b` a
    range. Written with every character escaped, so Python reads no set
    operation into it."""
    if not members:
        # `[]` matches nothing in JavaScript, and `[^]` any character.
        return "[\\s\\S]" if negate else "(?!)"
    parts: list[str] = []
    k = 0
    while k < len(members):
        if k + 2 < len(members) and members[k + 1] == "-":
            parts.append(f"{_class_atom(members[k])}-{_class_atom(members[k + 2])}")
            k += 3
        else:
            parts.append(_class_atom(members[k]))
            k += 1
    return f"[{'^' if negate else ''}{''.join(parts)}]"


def _range_of(operand: _Operand) -> tuple[float, float] | None:
    """The range an operand lies in: a number is a range of one."""
    if _is_integer(operand):
        number = float(operand)
        return (number, number)
    if isinstance(operand, _Range):
        return operand.range
    return None


def _whole_arithmetic(left: _Operand, right: _Operand, op: str) -> _Value | _Range:
    # A range moves under + and -: `Second(Now) + 1000` lies in 1000 to 1059.
    if (op == "+" or op == "-") and (isinstance(left, _Range) or isinstance(right, _Range)):
        a = _range_of(left)
        b = _range_of(right)
        if a is None or b is None:
            return None
        return _Range((a[0] + b[0], a[1] + b[1]) if op == "+" else (a[0] - b[1], a[1] - b[0]))
    if not _is_integer(left) or not _is_integer(right):
        return None
    x = float(left)
    y = float(right)
    if op == "mod" and y == 0:
        return None
    value = x + y if op == "+" else x - y if op == "-" else x * y if op == "*" else math.fmod(x, y)
    return value + 0 if abs(value) <= 2147483647 else None


def _compare(a: float, b: float, op: str) -> bool:
    if op == "=":
        return a == b
    if op == "<>":
        return a != b
    if op == "<":
        return a < b
    if op == ">":
        return a > b
    if op == "<=":
        return a <= b
    return a >= b
