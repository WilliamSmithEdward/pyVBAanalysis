"""What a Function of the module returns, where its text fixes it (XLIDE issue
#448). A Function that never assigns its name returns its type's default: 0, or ""
for a String. One whose every assignment is the same literal, and that assigns it
before anything that may leave, returns that literal as its type holds it. So `10
/ F()` divides by 0 and `n = F()` stores "abc" in an Integer, each measured in
Excel 16.0. Anything else that names the result (a read, a ByRef pass, a
non-literal value), an Err.Raise or Error that may end the Function first, or a
Variant that is never assigned (Empty) leaves the result unknown.

Ported from xlide_vscode/src/analyzer/diagnostics/functionResults.ts.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from typing import Literal, Union

from ..conditional import ConditionalActivityTracker
from ..identity_cache import IdentityLru
from ..js_compat import js_number, js_number_to_string
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import (
    BodyNode,
    ConditionalDirectiveNode,
    ModuleNode,
    ProcedureNode,
    ProcKind,
    Span,
    StatementNode,
    VariableGroupNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ..symbols.symbol_model import ModuleSymbols
from ..types.type_names import normalize_type
from .call_extraction import string_literal_value
from .walker import (
    active_module_members,
    block_footer_line_span,
    block_header_line_span,
    raw_expression_tokens,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)


@dataclass(frozen=True, slots=True)
class FunctionNumberResult:
    """A number carries the Function's own type, unless it is a Variant."""

    value: int | float
    type: str | None = None
    kind: Literal["number"] = "number"


@dataclass(frozen=True, slots=True)
class FunctionStringResult:
    value: str
    kind: Literal["string"] = "string"


@dataclass(frozen=True, slots=True)
class FunctionNullResult:
    kind: Literal["null"] = "null"


FunctionResult = Union[FunctionNumberResult, FunctionStringResult, FunctionNullResult]

_INTEGER_TYPES: dict[str, tuple[int, int]] = {
    "byte": (0, 255),
    "integer": (-32768, 32767),
    "long": (-2147483648, 2147483647),
}
_FRACTIONAL_TYPES: frozenset[str] = frozenset({"single", "double", "currency"})
_SUFFIX_TYPES: dict[str, str] = {
    "%": "integer", "&": "long", "!": "single", "#": "double", "@": "currency", "$": "string",
}
# Statement heads after which a Function may return before a later assignment.
_LEAVING_HEADS: frozenset[str] = frozenset(
    {"exit", "goto", "gosub", "return", "end", "resume", "on", "stop", "error"}
)

_EXACT_INTEGER_LIMIT = 2**53


def _js_value(number: float) -> int | float:
    """A JavaScript number as the port holds it: an int where it is a whole number
    JavaScript holds exactly, so it prints as JavaScript prints it."""
    return int(number) if number.is_integer() and abs(number) < _EXACT_INTEGER_LIMIT else number


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) out of range."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return None if tok is None else tok.raw_text


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name is not None else None


class _KnownFunctionResults(dict[str, FunctionResult]):
    """A results map that carries what it needs to run one of its Functions for a
    call's arguments: upstream keys that in a WeakMap by the map itself."""

    __slots__ = ("source", "procedures", "activity")

    def __init__(
        self,
        source: str,
        procedures: Mapping[str, ProcedureNode | None],
        activity: ConditionalActivityTracker | None,
    ) -> None:
        super().__init__()
        self.source = source
        self.procedures = procedures
        self.activity = activity


_RESULTS = IdentityLru()


def known_function_results(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None
) -> Mapping[str, FunctionResult]:
    """Each Function of the module whose result is known, by lowercased name."""
    cached = _RESULTS.get(mod)
    if isinstance(cached, _KnownFunctionResults) and cached.source == source and cached.activity is activity:
        return cached
    procedures: dict[str, ProcedureNode | None] = {}
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            lower = member.name.lower()
            procedures[lower] = None if lower in procedures else member
    out = _KnownFunctionResults(source, procedures, activity)
    for lower, proc in procedures.items():
        result = _result_of(source, proc, activity) if proc is not None and proc.proc_kind is ProcKind.FUNCTION else None
        if result is not None:
            out[lower] = result
    _RESULTS.put(out, mod)
    return out


@dataclass(frozen=True, slots=True)
class FunctionResultAt:
    result: FunctionResult
    end: int


def function_result_at(
    toks: Sequence[VbaToken],
    start: int,
    results: Mapping[str, FunctionResult],
    caller: ProcedureNode,
    symbols: ModuleSymbols,
) -> FunctionResultAt | None:
    """The known result a call stands for at `toks[start]`: `F()`, or a bare `F` not
    followed by an argument list, with F a Function of the module the calling
    procedure does not shadow. None otherwise."""
    name = _lower_name(_at(toks, start))
    result = results.get(name) if name else None
    if (
        result is None
        or name is None
        or _raw(toks, start - 1) == "."
        or _raw(toks, start - 1) == "!"
        or _shadowed(name, caller, symbols)
    ):
        return None
    if _raw(toks, start + 1) == "(":
        return FunctionResultAt(result=result, end=start + 2) if _raw(toks, start + 2) == ")" else None
    if _raw(toks, start + 1) == "." or _raw(toks, start + 1) == "!":
        return None
    return FunctionResultAt(result=result, end=start)


def function_result_named(
    lower: str,
    results: Mapping[str, FunctionResult],
    caller: ProcedureNode,
    symbols: ModuleSymbols,
) -> FunctionResult | None:
    """The known result of the Function a name calls from `caller`, unless the caller shadows it."""
    result = results.get(lower)
    return result if result is not None and not _shadowed(lower, caller, symbols) else None


# JavaScript's `\w` and `\d` are ASCII; its `$` (no m flag) is the end of the input.
_CALL_WITH_ARGS_RE = re.compile(r"^([A-Za-z0-9_]+)\((-?[0-9]+(?:,-?[0-9]+)*)?\)\Z")
_BARE_NAME_RE = re.compile(r"^([A-Za-z0-9_]+)\Z")


def function_integer_result(
    name: str,
    results: Mapping[str, FunctionResult],
    caller: ProcedureNode,
    symbols: ModuleSymbols,
) -> int | float | None:
    """A whole-number result for an integer lookup's name, `f` or `f()`."""
    lowered = name.lower()
    call = _CALL_WITH_ARGS_RE.match(lowered) or _BARE_NAME_RE.match(lowered)
    stripped = lowered[:-2] if lowered.endswith("()") else lowered
    result = function_result_named(stripped, results, caller, symbols)
    if result is None and call is not None:
        args_text = call.group(2) if call.re is _CALL_WITH_ARGS_RE else None
        args = [_js_value(js_number(arg)) for arg in args_text.split(",")] if args_text else []
        result = _call_result(call.group(1), args, results, caller, symbols)
    if isinstance(result, FunctionNumberResult) and float(result.value).is_integer():
        return result.value
    return None


def _call_result(
    lower: str,
    args: Sequence[int | float],
    results: Mapping[str, FunctionResult],
    caller: ProcedureNode,
    symbols: ModuleSymbols,
) -> FunctionResult | None:
    """`Sign1(-1)`: a Function of the module run for whole-number arguments, its
    result held as its type holds it (issue #562)."""
    from ..types.type_inference import function_result_for

    context = results if isinstance(results, _KnownFunctionResults) else None
    proc = context.procedures.get(lower) if context is not None else None
    if context is None or proc is None or _shadowed(lower, caller, symbols):
        return None
    value = function_result_for(
        context.source,
        proc,
        symbols,
        context.activity,
        [raw_expression_tokens(js_number_to_string(arg)) for arg in args],
    )
    literal = _literal_of([tok for tok in value if tok.kind is not TokenKind.COMMENT]) if value else None
    type_ = _SUFFIX_TYPES.get(proc.type_suffix) if proc.type_suffix else normalize_type(proc.return_type or "Variant")
    return _held_as(literal, type_) if literal is not None and type_ else None


def _shadowed(lower: str, caller: ProcedureNode, symbols: ModuleSymbols) -> bool:
    from ..types.type_inference import procedure_symbol_for

    if caller.name.lower() == lower or any(param.name.lower() == lower for param in caller.params):
        return True
    proc_symbol = procedure_symbol_for(symbols, caller)
    return any(child.name.lower() == lower for child in (proc_symbol.children if proc_symbol is not None else None) or [])


def _result_of(
    source: str, proc: ProcedureNode, activity: ConditionalActivityTracker | None
) -> FunctionResult | None:
    if "(" in (proc.return_type or ""):
        return None
    type_ = _SUFFIX_TYPES.get(proc.type_suffix) if proc.type_suffix else normalize_type(proc.return_type or "Variant")
    if not type_:
        return None
    lower = proc.name.lower()
    literals: set[tuple[str, int | float | str | None]] = set()
    assigned: FunctionResult | None = None
    other = False

    def visit(toks: Sequence[VbaToken]) -> None:
        nonlocal assigned, other
        head = 1 if token_text(_at(toks, 0)) == "let" else 0
        first = token_text(_at(toks, 0))
        if first == "error" or (first == "err" and token_text(_at(toks, 2)) == "raise"):
            other = True
        for i, tok in enumerate(toks):
            if _lower_name(tok) != lower or _raw(toks, i - 1) == "." or _raw(toks, i - 1) == "!":
                continue
            value = _literal_of(toks[i + 2 :]) if i == head and _raw(toks, i + 1) == "=" else None
            if value is None:
                other = True
                continue
            literals.add(_literal_key(value))
            assigned = value

    _walk(source, proc.body, activity, visit)
    if other or len(literals) > 1:
        return None
    if assigned is None:
        if type_ == "string":
            return FunctionStringResult(value="")
        if type_ in _INTEGER_TYPES or type_ in _FRACTIONAL_TYPES or type_ == "boolean":
            return FunctionNumberResult(value=0, type=type_)
        return None
    converted = _held_as(assigned, type_)
    held = (
        FunctionNumberResult(value=converted.value, type=type_)
        if isinstance(converted, FunctionNumberResult) and type_ != "variant"
        else converted
    )
    if held is None:
        return None
    if isinstance(held, FunctionNumberResult):
        is_default = held.value == 0
    elif isinstance(held, FunctionStringResult):
        is_default = held.value == ""
    else:
        is_default = False
    return held if is_default or _assigns_first(source, proc.body, activity, lower) else None


def _literal_key(value: FunctionResult) -> tuple[str, int | float | str | None]:
    """Upstream compares literals by JSON.stringify, which drops no field here."""
    if isinstance(value, FunctionNullResult):
        return ("null", None)
    return (value.kind, value.value)


def _skip(activity: ConditionalActivityTracker | None) -> Callable[[BodyNode], bool]:
    def skip(node: BodyNode) -> bool:
        return (
            (activity is not None and activity.is_inactive(node.span))
            or isinstance(node, (ConditionalDirectiveNode, VariableGroupNode))
        )

    return skip


def _walk(
    source: str,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    visit: Callable[[Sequence[VbaToken]], None],
) -> None:
    """Each statement, and each block's header and footer line, on an explicit
    stack in upstream's recursive order."""
    for node in iter_body_nodes(body, _skip(activity)):
        if is_leaf_statement(node):
            for span in statement_and_branch_spans(node):
                visit(_tokens(source, span))
            continue
        visit(_tokens(source, block_header_line_span(source, node.span)))
        visit(_tokens(source, block_footer_line_span(source, node.span)))


def _assigns_first(
    source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None, lower: str
) -> bool:
    """Whether a top-level statement assigns the result before anything that may leave."""
    skip = _skip(activity)
    for node in body:
        if skip(node):
            continue
        if is_leaf_statement(node):
            toks = _tokens(source, node.span)
            head = 1 if token_text(_at(toks, 0)) == "let" else 0
            single_line_if = isinstance(node, StatementNode) and node.single_line_if_branches is not None
            if not single_line_if and _lower_name(_at(toks, head)) == lower and _raw(toks, head + 1) == "=":
                return True
            if _may_leave(source, [node], activity):
                return False
            continue
        if _may_leave(source, [node], activity):
            return False
    return False


def _may_leave(source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None) -> bool:
    leaves = False

    def visit(toks: Sequence[VbaToken]) -> None:
        nonlocal leaves
        first = token_text(_at(toks, 0))
        if first in _LEAVING_HEADS or (first == "err" and token_text(_at(toks, 2)) == "raise"):
            leaves = True

    _walk(source, body, activity, visit)
    return leaves


def _tokens(source: str, span: Span) -> list[VbaToken]:
    return statement_tokens_after_leading_label(source, span)


_PLAIN_NUMBER_RE = re.compile(r"^[0-9.]+(?:[Ee][+-]?[0-9]+)?[%&!#@]?\Z")
_NUMBER_SUFFIX_RE = re.compile(r"[%&!#@]\Z")


def _literal_of(toks: Sequence[VbaToken]) -> FunctionResult | None:
    if len(toks) == 1 and toks[0].kind is TokenKind.STRING_LITERAL:
        return FunctionStringResult(value=string_literal_value(toks[0].raw_text))
    word = token_text(toks[0]) if len(toks) == 1 else ""
    if word == "null":
        return FunctionNullResult()
    if word == "true" or word == "false":
        return FunctionNumberResult(value=-1 if word == "true" else 0)
    negative = len(toks) == 2 and toks[0].raw_text == "-"
    number = _at(toks, 1 if negative else 0)
    if (
        len(toks) != (2 if negative else 1)
        or number is None
        or (number.kind is not TokenKind.INTEGER_LITERAL and number.kind is not TokenKind.FLOAT_LITERAL)
        or _PLAIN_NUMBER_RE.match(number.raw_text) is None
    ):
        return None
    value = js_number(_NUMBER_SUFFIX_RE.sub("", number.raw_text, count=1))
    if not math.isfinite(value):
        return None
    return FunctionNumberResult(value=_js_value(-value if negative else value))


def _held_as(value: FunctionResult, type_: str) -> FunctionResult | None:
    """A literal as a Function of this type returns it, or None where the
    assignment raises or the value is not plain."""
    if type_ == "variant":
        return value
    if not isinstance(value, FunctionNumberResult):
        return value if type_ == "string" and isinstance(value, FunctionStringResult) else None
    integer_range = _INTEGER_TYPES.get(type_)
    if integer_range is not None:
        rounded = _round_half_even(value.value)
        return FunctionNumberResult(value=rounded) if integer_range[0] <= rounded <= integer_range[1] else None
    if type_ in _FRACTIONAL_TYPES:
        return value
    return FunctionNumberResult(value=0 if value.value == 0 else -1) if type_ == "boolean" else None


def _js_round(value: float) -> int:
    """Math.round: the nearest integer, a half rounding up toward +Infinity."""
    floor = math.floor(value)
    return floor + 1 if value - floor >= 0.5 else floor


def _round_half_even(value: int | float) -> int:
    floor = math.floor(value)
    if value - floor != 0.5:
        return _js_round(value)
    return floor if floor % 2 == 0 else floor + 1
