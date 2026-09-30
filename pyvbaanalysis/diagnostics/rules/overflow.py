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
 - for-counter-overflow: `For i = 1 To 32767` with i an Integer, and
   `For b = 0 To 255` with b a Byte: the increment after the last pass
   overflows the counter. `To 32766` runs.

The folder follows MS-VBAL 5.6.9.3: Byte and Integer operands make Integer
results, Long makes Long, Single and Double make Double, Currency makes
Currency; `/` and `^` make Double. A value the folder cannot type stays unknown
and nothing is reported for it.

Every value is a Python float, as upstream's are JavaScript numbers, so the
arithmetic rounds and overflows to infinity the way upstream's does; the helpers
at the end of the module print numbers the way JavaScript does.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Context, Decimal
from functools import lru_cache
from typing import Union

from ...conditional import ConditionalActivityTracker, inactive_node_skip
from ...constants.integer_constant_expression import parse_vba_integer_literal
from ...flow.procedure_labels import statement_label_declaration
from ...host.host_model import HostObjectModel
from ...js_compat import js_number, js_number_to_string
from ...lexer.token_helpers import match_paren_from
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ForBlockNode,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    VariableGroupNode,
    is_leaf_statement,
    iter_body_nodes_in_context,
)
from ...symbols.symbol_model import ModuleSymbols, VbaSymbol, VbaSymbolKind
from ...types.type_inference import procedure_symbol_for, type_environment_for
from ...types.type_names import normalize_type
from ..context import PushFn, statement_tokens
from ..known_locals import KnownLocalValue, known_local_literal_values
from ..walker import (
    active_module_members,
    bare_assignment_target,
    first_executable_token_index,
    for_each_variable_group,
    raw_expression_tokens,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .shared import is_bare_or_vba_qualified_intrinsic_call


@dataclass(frozen=True, slots=True)
class _Typed:
    value: float
    type: str


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
    "single": _Range(-3.402823e38, 3.402823e38, "Single"),
    "double": _Range(-1.7976931348623157e308, 1.7976931348623157e308, "Double"),
    "currency": _Range(-922337203685477.5807, 922337203685477.5807, "Currency"),
    "date": _Range(-657434.0, 2958465.0, "Date"),
}

_RANK: dict[str, int] = {
    "byte": 0, "integer": 1, "long": 2, "single": 3, "double": 4, "currency": 5, "date": 6,
}

# The types a value is stored in without rounding to a whole number.
_UNROUNDED_TYPES = frozenset({"single", "double", "currency", "date"})

_CONVERSIONS: dict[str, str] = {
    "cbyte": "byte", "cint": "integer", "clng": "long", "csng": "single", "cdbl": "double",
    "ccur": "currency", "cdate": "date", "abs": "abs", "int": "int", "fix": "fix", "exp": "exp",
    "hex": "hex", "oct": "oct",
}

_CONVERSION_NAMES: dict[str, str] = {
    "cbyte": "CByte", "cint": "CInt", "clng": "CLng", "csng": "CSng", "cdbl": "CDbl",
    "ccur": "CCur", "cdate": "CDate",
}

_INTEGER_SUFFIX_RE = re.compile(r"[%&^]$")
_FLOAT_SUFFIX_RE = re.compile(r"[!#@]$")
_EXPONENT_LETTER_RE = re.compile(r"[dD]")
_VOWEL_START_RE = re.compile(r"^[AEIOU]")


def _label(type_name: str) -> str:
    return _RANGES[type_name].label


def _arithmetic_result_type(a: str, b: str) -> str:
    """The result type of `a op b` for + - * \\ Mod (MS-VBAL 5.6.9.3)."""
    if a == "currency" or b == "currency":
        floating = a in ("double", "single") or b in ("double", "single")
        return "double" if floating else "currency"
    if a == "date" or b == "date":
        return "double"
    wider = a if _RANK[a] >= _RANK[b] else b
    return "integer" if wider == "byte" else wider


def _in_range(value: float, type_name: str) -> bool:
    bounds = _RANGES[type_name]
    return math.isfinite(value) and bounds.min <= value <= bounds.max


def _bankers_round(value: float) -> float:
    """VBA's rounding to a whole number: banker's rounding at .5."""
    floor = _js_floor(value)
    fraction = value - floor
    if fraction > 0.5:
        return floor + 1
    if fraction < 0.5:
        return floor
    return floor if math.fmod(floor, 2) == 0 else floor + 1


def _literal_typed(tok: VbaToken) -> _Typed | None:
    """A literal's natural type and value: 3 is Integer, 40000 is Long, 3000000000
    is Double."""
    if tok.kind is TokenKind.INTEGER_LITERAL:
        raw = tok.raw_text
        suffix_match = _INTEGER_SUFFIX_RE.search(raw)
        suffix = suffix_match.group(0) if suffix_match is not None else None
        parsed = parse_vba_integer_literal(raw)
        if parsed is None:
            return None
        value = float(parsed)
        if suffix == "%":
            return _Typed(value, "integer") if _in_range(value, "integer") else None
        if suffix == "&":
            return _Typed(value, "long")
        if suffix == "^":
            return None  # LongLong is not modelled here
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
        type_name = "single" if suffix == "!" else "currency" if suffix == "@" else "double"
        return _Typed(value, type_name)
    return None


# What a name means to the folder: a typed value, or nothing.
_NameLookup = Callable[[str], Union[_Typed, None]]


class _TypedFolder:
    """Folds an arithmetic expression over literals, Consts and known locals with
    VBA's result typing, stopping at the first operation whose result its type
    cannot hold.

    It recurses only over one expression's tokens: parentheses, unary signs and
    conversion arguments.
    """

    __slots__ = ("_toks", "_base", "_names", "_index")

    def __init__(self, toks: Sequence[VbaToken], base: int, names: _NameLookup) -> None:
        self._toks = toks
        self._base = base
        self._names = names
        self._index = 0

    def fold(self) -> _Folded:
        if len(self._toks) == 0:
            return None
        try:
            result = self._additive()
        except RecursionError:
            # Port-only: each parenthesis that holds an operation costs one fold
            # per precedence level, so an expression nested past the interpreter
            # limit is unknown here, where XLIDE folds it. Letting the error out
            # would drop every overflow finding in the module.
            return None
        if isinstance(result, _Overflow):
            return result
        return result if self._index == len(self._toks) else None

    def _at(self, i: int) -> VbaToken | None:
        return self._toks[i] if 0 <= i < len(self._toks) else None

    def _span(self, start: int, end: int) -> Span:
        return Span(self._base + self._toks[start].start, self._base + self._toks[end].end)

    def _additive(self) -> _Folded:
        start = self._index
        left = self._multiplicative()
        while isinstance(left, _Typed):
            op = self._at(self._index)
            if op is None or op.kind is not TokenKind.OPERATOR or op.raw_text not in ("+", "-"):
                break
            self._index += 1
            right = self._multiplicative()
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
        left: Callable[[], _Folded],
        right: Callable[[], _Folded] | None = None,
    ) -> _Folded:
        operand_of = right if right is not None else left
        start = self._index
        value = left()
        while isinstance(value, _Typed):
            op = self._at(self._index)
            if op is None:
                break
            word = op.raw_text if op.kind is TokenKind.OPERATOR else token_text(op)
            if word not in operators:
                break
            self._index += 1
            operand = operand_of()
            if not isinstance(operand, _Typed):
                return operand
            value = self._combine(value, operand, word, start, self._index - 1)
        return value

    def _multiplicative(self) -> _Folded:
        return self._left_associative(("mod",), self._integer_division)

    def _integer_division(self) -> _Folded:
        return self._left_associative(("\\",), self._product)

    def _product(self) -> _Folded:
        return self._left_associative(("*", "/"), self._unary)

    def _power(self) -> _Folded:
        """`a ^ b` binds above unary minus (`-2 ^ 2` is -4); the exponent may carry
        its own sign."""
        return self._left_associative(("^",), self._primary, self._unary)

    def _unary(self) -> _Folded:
        tok = self._at(self._index)
        if tok is not None and tok.kind is TokenKind.OPERATOR and tok.raw_text in ("-", "+"):
            start = self._index
            self._index += 1
            operand = self._unary()
            if not isinstance(operand, _Typed):
                return operand
            if tok.raw_text == "+":
                return operand
            value = -operand.value
            if not _in_range(value, operand.type):
                return _Overflow(
                    self._span(start, self._index - 1),
                    f"Negating {_show_number(operand.value)} gives {_show_number(-operand.value)}, "
                    f"which does not fit {_label(operand.type)}",
                )
            return _Typed(value, operand.type)
        return self._power()

    def _primary(self) -> _Folded:
        tok = self._at(self._index)
        if tok is None:
            return None
        if tok.raw_text == "(":
            close = match_paren_from(self._toks, self._index)
            if close < 0:
                return None
            inner_start, inner_end = self._index + 1, close
            # `((x))` folds as `(x)`: peel the layers that only wrap another
            # parenthesis here, so redundant nesting costs no recursion.
            while (
                inner_end - inner_start >= 2
                and self._toks[inner_start].raw_text == "("
                and match_paren_from(self._toks, inner_start) == inner_end - 1
            ):
                inner_start += 1
                inner_end -= 1
            nested = _TypedFolder(self._toks[inner_start:inner_end], self._base, self._names)
            value = nested.fold()
            self._index = close + 1
            return value
        literal = _literal_typed(tok)
        if literal is not None:
            self._index += 1
            return literal
        name = token_name(tok)
        if not name:
            return None
        # `VBA.CInt(...)` and `CInt(...)`.
        callee_index = self._index
        dot = self._at(self._index + 1)
        if (
            name.lower() == "vba"
            and dot is not None
            and dot.raw_text == "."
            and token_name(self._at(self._index + 2))
        ):
            callee_index = self._index + 2
        callee = (token_name(self._at(callee_index)) or "").lower()
        opener = self._at(callee_index + 1)
        if opener is not None and opener.raw_text == "(" and callee in _CONVERSIONS:
            close = match_paren_from(self._toks, callee_index + 1)
            if close < 0:
                return None
            argument = self._toks[callee_index + 2 : close]
            inner = _TypedFolder(argument, self._base, self._names).fold()
            start = self._index
            self._index = close + 1
            if not isinstance(inner, _Typed):
                return inner
            return self._convert(callee, inner, self._span(start, close))
        following = self._at(self._index + 1)
        if following is not None and following.raw_text == ".":
            # `Rows.Count`: a two-part member the lookup may know as a constant.
            member = token_name(self._at(self._index + 2))
            after_tok = self._at(self._index + 3)
            after = after_tok.raw_text if after_tok is not None else None
            if member and after != "." and after != "(":
                known = self._names(f"{name.lower()}.{member.lower()}")
                if known is not None:
                    self._index += 3
                    return known
            return None
        if following is not None and following.raw_text == "(":
            return None  # a call the folder does not know
        known = self._names(name.lower())
        if known is None:
            return None
        self._index += 1
        return known

    def _convert(self, callee: str, inner: _Typed, span: Span) -> _Folded:
        target = _CONVERSIONS[callee]
        if target == "abs":
            magnitude = abs(inner.value)
            if _in_range(magnitude, inner.type):
                return _Typed(magnitude, inner.type)
            return _Overflow(
                span, f"Abs({_show_number(inner.value)}) does not fit {_label(inner.type)}"
            )
        if target in ("int", "fix"):
            whole = _js_floor(inner.value) if target == "int" else _js_trunc(inner.value)
            return _Typed(
                whole, inner.type if inner.type in ("byte", "integer", "long") else "double"
            )
        if target == "exp":
            power = _js_exp(inner.value)
            if _in_range(power, "double"):
                return _Typed(power, "double")
            return _Overflow(span, f"Exp({_show_number(inner.value)}) exceeds the Double range")
        if target in ("hex", "oct"):
            # Hex and Oct take a value that fits a Long (or a LongLong on 64-bit
            # for whole numbers; 1E+20 fits neither).
            if _in_range(inner.value, "long") or (
                _is_integer(inner.value) and abs(inner.value) < 9.2e18
            ):
                return None
            return _Overflow(
                span,
                f"{'Hex' if callee == 'hex' else 'Oct'}({js_number_to_string(inner.value)}) "
                "takes a value outside the Long range",
            )
        value = inner.value if target in _UNROUNDED_TYPES else _bankers_round(inner.value)
        if _in_range(value, target):
            return _Typed(value, target)
        return _Overflow(
            span,
            f"{_CONVERSION_NAMES[callee]}({_show_number(inner.value)}) "
            f"does not fit {_label(target)}",
        )

    def _combine(self, left: _Typed, right: _Typed, op: str, start: int, end: int) -> _Folded:
        span = self._span(start, end)
        if op == "/":
            if right.value == 0:
                return None  # division by zero is another rule's
            currency = left.type == "currency" or right.type == "currency"
            type_name = "currency" if currency else "double"
            value = left.value / right.value
        elif op == "^":
            type_name = "double"
            value = _js_pow(left.value, right.value)
            if math.isnan(value) or (left.value == 0 and right.value < 0):
                # 0 ^ -1 raises 5, Invalid procedure call: runtime-value-out-of-range reports it.
                return None
        elif op in ("\\", "mod"):
            type_name = _arithmetic_result_type(left.type, right.type)
            a = _bankers_round(left.value)
            b = _bankers_round(right.value)
            if b == 0:
                return None
            # JavaScript's % keeps the dividend's sign, as math.fmod does.
            value = math.fmod(a, b) if op == "mod" else _js_trunc(a / b)
        else:
            type_name = _arithmetic_result_type(left.type, right.type)
            if op == "+":
                value = left.value + right.value
            elif op == "-":
                value = left.value - right.value
            else:
                value = left.value * right.value
        if not _in_range(value, type_name):
            shown = _show_number(value)
            return _Overflow(
                span,
                f"{_describe(left)} {'Mod' if op == 'mod' else op} {_describe(right)} is {shown}, "
                f"outside the {_label(type_name)} range",
            )
        return _Typed(value, type_name)


def _describe(typed: _Typed) -> str:
    return f"{_show_number(typed.value)} ({_label(typed.type)})"


def _show_number(value: float) -> str:
    """A value as VBA would print it: whole numbers plain, huge or fractional ones
    in E notation."""
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
    return None


def _constant_lookup(base: Mapping[str, _Typed], candidates: Sequence[VbaSymbol]) -> Mapping[str, _Typed]:
    """The Consts a procedure can name, folded with their declared or natural type:
    `Private Const HOURS As Integer = 24` is an Integer 24, and `Const K = 40000` a
    Long. A Const the folder cannot fold is left out.

    The values of the Consts in `candidates` are layered over `base`: a name
    declared in both takes the candidate's value (a procedure's own Const wins over
    the module's), and a candidate's value may refer to a base constant. The module
    and project layer is folded once per pass and each procedure adds only its own
    Consts on top; folding the whole project's constants again for every procedure
    was 15% of a large module's analysis (XLIDE issue #139)."""
    # Later entries shadow earlier ones.
    pending: dict[str, VbaSymbol] = {}
    for symbol in candidates:
        if symbol.kind is VbaSymbolKind.CONSTANT and symbol.default_raw is not None:
            pending[symbol.name.lower()] = symbol
    if not pending:
        return base
    folded: dict[str, _Typed] = {}
    for lower in pending:
        if lower not in folded:
            _resolve_constant(lower, pending, base, folded)
    out = dict(base)
    for lower in pending:
        typed = folded.get(lower)
        # A candidate that did not fold still shadows the base name.
        if typed is not None:
            out[lower] = typed
        else:
            out.pop(lower, None)
    return out


@dataclass(slots=True)
class _ResolveFrame:
    name: str
    # The Consts whose resolution failed while this frame waited on them.
    failed: set[str] = field(default_factory=set)
    # The Const this frame's last fold met unresolved.
    need: str | None = None


def _resolve_constant(
    root: str, pending: Mapping[str, VbaSymbol], base: Mapping[str, _Typed], out: dict[str, _Typed]
) -> None:
    """Upstream's recursive `resolve` of one name, run on an explicit stack so a
    long chain of Consts naming each other cannot exhaust Python's recursion
    limit. A name that is not pending reads the `base` layer.

    A frame whose fold meets a pending Const not yet resolved resolves that one
    first, with the same names in progress upstream's `resolving` set would hold,
    then folds again from the start: every lookup before it came from `out`, so
    the second fold reaches the same point and goes on with the answer. Upstream
    caches no failure, so a failure is remembered only by the frame that asked,
    whose parse stops there, and is worked out afresh anywhere else.
    """
    frames = [_ResolveFrame(root)]
    resolving = {root}

    def lookup(lower: str) -> _Typed | None:
        if lower not in pending:
            return base.get(lower)
        cached = out.get(lower)
        if cached is not None:
            return cached
        frame = frames[-1]
        if lower in resolving or lower in frame.failed:
            return None
        frame.need = lower
        return None

    while frames:
        frame = frames[-1]
        frame.need = None
        symbol = pending[frame.name]
        folded = _TypedFolder(_const_value_tokens(symbol.default_raw or ""), 0, lookup).fold()
        if frame.need is not None:
            frames.append(_ResolveFrame(frame.need))
            resolving.add(frame.need)
            continue
        frames.pop()
        resolving.discard(frame.name)
        typed = _typed_constant(symbol, folded)
        if typed is not None:
            out[frame.name] = typed
        elif frames:
            frames[-1].failed.add(frame.name)


def _typed_constant(symbol: VbaSymbol, folded: _Folded) -> _Typed | None:
    if not isinstance(folded, _Typed):
        return None
    declared = _numeric_type_of(symbol.as_type)
    typed = _Typed(folded.value, declared) if declared else folded
    if declared and not _in_range(_bankers_round(folded.value), declared):
        return None
    return typed


@lru_cache(maxsize=4096)
def _const_value_tokens(raw: str) -> tuple[VbaToken, ...]:
    # Upstream's rawExpressionTokens for the Const's text. Memoized: the project's
    # Consts are folded again in every module's pass.
    return tuple(_without_comments(raw_expression_tokens(raw)))


_EXCEL_HOST_CONSTANTS: dict[str, _Typed] = {
    "rows.count": _Typed(1048576.0, "long"),
    "columns.count": _Typed(16384.0, "long"),
}


def _host_constant_values(host_model: HostObjectModel | None) -> Mapping[str, _Typed]:
    """Host members whose value is fixed: Excel's `Rows.Count` is 1048576 and
    `Columns.Count` 16384 on every worksheet since Excel 2007 (a Long each).
    Absent model means Excel (XLIDE issue #28)."""
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
    module_constants = _constant_lookup({}, [*(project_visible_symbols or []), *(symbols.root.children or [])])
    _check_const_declarations(
        source,
        [member for member in mod.members if isinstance(member, VariableGroupNode)],
        module_constants,
        activity,
        push,
    )
    host_values = _host_constant_values(host_model)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        proc_sym = procedure_symbol_for(symbols, member)
        constants = _constant_lookup(module_constants, (proc_sym.children if proc_sym is not None else None) or [])
        env = type_environment_for(symbols, member)
        known = known_local_literal_values(source, member, symbols, activity)
        # Values a straight run of top-level statements has just stored:
        # `i = 32767` followed by `i = i + 1`.
        just_assigned: dict[str, _Typed] = {}
        names = _procedure_name_lookup(constants, just_assigned, known, env, host_values)
        groups: list[VariableGroupNode] = []
        for_each_variable_group(member.body, groups.append, activity)
        _check_const_declarations(source, groups, constants, activity, push)
        _check_procedure_body(source, member, env, names, just_assigned, activity, push)


def _procedure_name_lookup(
    constants: Mapping[str, _Typed],
    just_assigned: Mapping[str, _Typed],
    known: Mapping[str, KnownLocalValue],
    env: Mapping[str, str],
    host_values: Mapping[str, _Typed],
) -> _NameLookup:
    def names(lower: str) -> _Typed | None:
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
        if "." in lower and lower[: lower.index(".")] not in env:
            return host_values.get(lower)
        return None

    return names


def _check_const_declarations(
    source: str,
    groups: Sequence[VariableGroupNode],
    constants: Mapping[str, _Typed],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
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
            folded = _TypedFolder(value, decl.span.start, constants.get).fold()
            if isinstance(folded, _Overflow):
                push(
                    "constOverflow",
                    f"Const '{decl.name}' overflows while it is evaluated: {folded.detail}. "
                    "This is a VBE compile error: Overflow.",
                    folded.span,
                )
                continue
            declared = _numeric_type_of(decl.as_type)
            if (
                folded is not None
                and declared
                and not _in_range(
                    folded.value if declared in _UNROUNDED_TYPES else _bankers_round(folded.value),
                    declared,
                )
            ):
                push(
                    "constOverflow",
                    f"Const '{decl.name}' is declared As {_label(declared)} but its value "
                    f"{js_number_to_string(folded.value)} is outside that range. "
                    "This is a VBE compile error: Overflow.",
                    Span(decl.span.start + value[0].start, decl.span.start + value[-1].end),
                )


def _check_procedure_body(
    source: str,
    proc: ProcedureNode,
    env: Mapping[str, str],
    names: _NameLookup,
    just_assigned: dict[str, _Typed],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    # Upstream clears the tracked values on entering a block and again on leaving
    # it. Nothing inside a block adds one, so the walk that pairs each node with
    # whether it sits at the top level of the body needs only the first clear.
    for node, top_level in iter_body_nodes_in_context(
        proc.body, True, _nested_body, inactive_node_skip(activity)
    ):
        if isinstance(node, ForBlockNode):
            _check_for_counter(source, node, env, names, push)
        if isinstance(getattr(node, "body", None), list):
            # A block may run any number of times: nothing stored before it
            # is known after it, and nothing inside it is straight-line.
            just_assigned.clear()
            continue
        if not is_leaf_statement(node):
            continue
        spans = statement_and_branch_spans(node)
        straight_line = (
            top_level
            and len(spans) == 1
            and not (isinstance(node, StatementNode) and node.single_line_if_branches is not None)
        )
        if not straight_line:
            just_assigned.clear()
        for span in spans:
            stored = _check_statement(source, span, env, names, push)
            if not straight_line:
                continue
            # Any other mention of a tracked name (a ByRef pass, a label a
            # GoTo could reach) ends what is known about it.
            toks = statement_tokens(source, span)
            if (
                statement_label_declaration(source, span) is not None
                or token_text(_token_at(toks, first_executable_token_index(toks))) == "gosub"
            ):
                just_assigned.clear()
                continue
            for tok in toks:
                name = token_name(tok)
                lower = name.lower() if name is not None else None
                if lower and lower in just_assigned:
                    del just_assigned[lower]
            if stored is not None:
                just_assigned[stored[0]] = stored[1]


def _nested_body(_block: BodyNode, _top_level: bool) -> bool:
    return False


def _check_statement(
    source: str,
    span: Span,
    env: Mapping[str, str],
    names: _NameLookup,
    push: PushFn,
) -> tuple[str, _Typed] | None:
    """The value a bare assignment provably stores, when the rule can tell, as
    (lowercased name, value)."""
    toks = statement_tokens(source, span)
    first = first_executable_token_index(toks)
    head = token_text(_token_at(toks, first))
    if head in ("const", "dim", "static", "redim"):
        return None
    bare = bare_assignment_target(source, span)
    if bare is not None:
        bare_name, _name_span, value_tokens = bare
        value = _without_comments(value_tokens)
        folded = _TypedFolder(value, span.start, names).fold()
        if isinstance(folded, _Overflow):
            push(
                "arithmeticOverflow",
                f"{folded.detail}. This will raise Run-time error '6': Overflow.",
                folded.span,
            )
            return None
        target = _numeric_type_of(env.get(bare_name.lower()))
        if folded is not None and target:
            stored = folded.value if target in _UNROUNDED_TYPES else _bankers_round(folded.value)
            if not _in_range(stored, target):
                rounded = (
                    f" ({_show_number(folded.value)} rounds to {_show_number(stored)})"
                    if stored != folded.value
                    else ""
                )
                bounds = _RANGES[target]
                push(
                    "arithmeticOverflow",
                    f"Assignment to '{bare_name}' stores {_show_number(stored)}{rounded} in "
                    f"{_article(bounds.label)} {bounds.label}, whose range is "
                    f"{js_number_to_string(bounds.min)} to {js_number_to_string(bounds.max)}. "
                    "This will raise Run-time error '6': Overflow.",
                    Span(span.start + value[0].start, span.start + value[-1].end),
                )
                return None
            return (bare_name.lower(), _Typed(stored, target))
        return None
    # Conversion calls anywhere else in the statement: `Debug.Print CInt(40000)`.
    for i in range(len(toks) - 1):
        callee = token_text(toks[i])
        if (
            callee not in _CONVERSIONS
            or toks[i + 1].raw_text != "("
            or not is_bare_or_vba_qualified_intrinsic_call(toks, i)
        ):
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0:
            continue
        start = i - 2 if i >= 1 and toks[i - 1].raw_text == "." else i
        folded = _TypedFolder(toks[start : close + 1], span.start, names).fold()
        if isinstance(folded, _Overflow):
            push(
                "arithmeticOverflow",
                f"{folded.detail}. This will raise Run-time error '6': Overflow.",
                folded.span,
            )
    return None


_LOOP_LEAVING_EXITS = frozenset({"for", "sub", "function", "property"})


def _body_may_leave_loop(source: str, body: Sequence[BodyNode]) -> bool:
    """True when a statement in the loop's body can leave the loop before the
    counter passes its type: `Exit For` (not one belonging to a nested For),
    `Exit Sub`/`Function`/`Property`, `GoTo`, or `End` (XLIDE issue #145). Such a
    loop's overflow is not proved, so it is not reported.

    Upstream recurses into each nested body; the explicit-stack walk visits the
    same statements, each with whether a nested For holds it.
    """
    for node, inside_nested_for in iter_body_nodes_in_context(
        body, False, lambda block, outer: outer or isinstance(block, ForBlockNode)
    ):
        if not is_leaf_statement(node):
            continue
        toks = statement_tokens_after_leading_label(source, node.span)
        for i, tok in enumerate(toks):
            word = token_text(tok)
            if word == "goto" or (word == "end" and len(toks) == 1):
                return True
            if word == "exit":
                target = token_text(_token_at(toks, i + 1))
                if target in _LOOP_LEAVING_EXITS and (target != "for" or not inside_nested_for):
                    return True
    return False


def _check_for_counter(
    source: str,
    node: ForBlockNode,
    env: Mapping[str, str],
    names: _NameLookup,
    push: PushFn,
) -> None:
    """`For i = 1 To 32767` with i an Integer: the counter is incremented past its
    last value before the exit test, and the increment overflows (measured:
    `To 32767` raises, `To 32766` runs; `For b = 0 To 255` raises for a Byte)."""
    if node.each or not node.control_variable:
        return
    type_name = _numeric_type_of(env.get(node.control_variable.lower()))
    if not type_name or type_name in _UNROUNDED_TYPES:
        return
    header_end = source.find("\n", node.span.start)
    header = Span(
        node.span.start, node.span.end if header_end < 0 else min(header_end, node.span.end)
    )
    toks = statement_tokens_after_leading_label(source, header)
    to = next((i for i, tok in enumerate(toks) if token_text(tok) == "to"), -1)
    if to < 0:
        return
    step = next((i for i, tok in enumerate(toks) if token_text(tok) == "step"), -1)
    limit_toks = _without_comments(toks[to + 1 : step if step > 0 else len(toks)])
    limit = _TypedFolder(limit_toks, header.start, names).fold()
    step_value = (
        _TypedFolder(_without_comments(toks[step + 1 :]), header.start, names).fold()
        if step > 0
        else _Typed(1.0, "integer")
    )
    if (
        not isinstance(limit, _Typed)
        or not isinstance(step_value, _Typed)
        or step_value.value == 0
        or not _is_integer(step_value.value)
    ):
        return
    # The counter's last value is the start plus whole steps: `For i = 1 To
    # 32766 Step 2` ends at 32765 and runs (measured). A step of 1 or -1 ends
    # at the limit whatever the start.
    eq = next((i for i, tok in enumerate(toks) if tok.raw_text == "="), -1)
    start = (
        _TypedFolder(_without_comments(toks[eq + 1 : to]), header.start, names).fold()
        if eq > 0
        else None
    )
    if abs(step_value.value) == 1:
        last = limit.value
    elif isinstance(start, _Typed) and _is_integer(start.value):
        passes = _js_floor((limit.value - start.value) / step_value.value)
        if passes < 0:
            return  # the loop body never runs and the counter stays at the start
        last = start.value + passes * step_value.value
    else:
        return
    bounds = _RANGES[type_name]
    if step_value.value > 0:
        overflows = last + step_value.value > bounds.max
    else:
        overflows = last + step_value.value < bounds.min
    if not overflows or _body_may_leave_loop(source, node.body):
        return
    push(
        "forCounterOverflow",
        f"Counter '{node.control_variable}' is {bounds.label}; after its last pass at "
        f"{js_number_to_string(last)} the loop adds {js_number_to_string(step_value.value)}, which does not fit. "
        "This will raise Run-time error '6': Overflow.",
        Span(header.start + limit_toks[0].start, header.start + limit_toks[-1].end),
    )


def _token_at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _without_comments(toks: Sequence[VbaToken]) -> list[VbaToken]:
    return [tok for tok in toks if tok.kind is not TokenKind.COMMENT]


# -- JavaScript number semantics ---------------------------------------------------


def _is_integer(value: float) -> bool:
    """Number.isInteger: finite and whole."""
    return float(value).is_integer()


def _js_floor(value: float) -> float:
    return float(math.floor(value)) if math.isfinite(value) else value


def _js_trunc(value: float) -> float:
    return float(math.trunc(value)) if math.isfinite(value) else value


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
        rounded = magnitude.quantize(
            Decimal(1).scaleb(exponent - fraction_digits, _EXACT), context=_EXACT
        )
        if rounded.adjusted() > exponent:
            exponent += 1
            rounded = magnitude.quantize(
                Decimal(1).scaleb(exponent - fraction_digits, _EXACT), context=_EXACT
            )
        coefficient = "".join(str(digit) for digit in rounded.as_tuple().digits)
    body = coefficient[0] + ("." + coefficient[1:] if fraction_digits else "")
    return f"{sign}{body}e{'+' if exponent >= 0 else '-'}{abs(exponent)}"
