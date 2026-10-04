"""Rule family: an argument a call leaves out, read by the procedure it calls
(XLIDE issue #260).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/omittedArguments.ts.

An omitted Optional Variant is Missing in the callee, a Variant holding error
448, and an omitted typed Optional holds its default: 0, "", or the value the
declaration gives. A ParamArray holds what the call passed, 0-based whatever
Option Base says. Each is known only with the call and the procedure together,
so this reads them at the call: `Opt()` against `Opt = x + 1` raises 13,
`Total()` against `Total = args(0)` raises 9. Every case was measured in Excel
16.0.

The callee is read from its first line up to the first use of the parameter.
Blocks that neither use it nor leave are passed over; anything that may jump
(Exit, GoTo, Return, End, Resume, On Error, Stop, Error) ends the read, and so
does a single-line If arm or a block that uses it, since that use runs only on
some paths. Only a use measured to raise is reported: IsMissing, IsError, CStr,
CLng and a Variant parameter read a Missing value cleanly. Calls into the same
module only, where the body is known.
"""

from __future__ import annotations

import json
import math
import re
from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass, field
from typing import Any

from ...conditional import ConditionalActivityTracker
from ...js_compat import js_number, js_number_to_string
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ConditionalDirectiveNode,
    LeafStatementNode,
    ModuleNode,
    ParameterNode,
    ProcedureNode,
    ProcKind,
    Span,
    StatementNode,
    VariableGroupNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ...symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbol, VbaSymbolKind
from ...types.type_inference import (
    known_local_literal_values_at,
    procedure_symbol_for,
    type_environment_for,
)
from ...types.type_names import is_known_scalar_type, is_numeric_type, normalize_type
from ..call_extraction import CallArguments, extract_call, is_named_slot, string_literal_value
from ..callable_signatures import (
    bare_callable_source_shadowed,
    callable_type_signatures_for,
    expression_calls,
    source_name_scope_for,
)
from ..context import PushFn, statement_tokens
from ..string_conversion import is_invalid_numeric_string
from ..walker import (
    ProcedureStatementVisitor,
    active_module_members,
    block_footer_line_span,
    block_header_line_span,
    for_each_statement_with_headers,
    match_paren_from,
    raw_expression_tokens,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .arrays import FixedArrayBound, local_fixed_arrays, module_option_base
from .shared import is_bare_or_vba_qualified_intrinsic_call


@dataclass(frozen=True, slots=True)
class _OmittedValue:
    """What an omitted or passed parameter holds in the callee."""

    kind: str  # "missing" | "nothing" | "unallocated" | "number" | "string"
    value: int | float | str | None = None

    def key(self) -> str:
        """Upstream's JSON.stringify of the value: its cache key."""
        if self.kind == "number":
            return f'{{"kind":"number","value":{js_number_to_string(float(self.value or 0))}}}'
        if self.kind == "string":
            return json.dumps({"kind": "string", "value": self.value}, ensure_ascii=False, separators=(",", ":"))
        return f'{{"kind":"{self.kind}"}}'


@dataclass(frozen=True, slots=True)
class _RaisingUse:
    """The callee's first use of an omitted parameter, when that use raises."""

    rule: str
    # What the use does, completing "and 'Proc' ...".
    does: str
    error: str
    # Absolute span of the use in the callee.
    span: Span


@dataclass(frozen=True, slots=True)
class _ParamArrayRead:
    """A ParamArray read past what the call passes."""

    index: int | float
    span: Span


# Intrinsics that raise 13 on a Missing first argument (each measured).
_MISSING_RAISING_INTRINSICS: frozenset[str] = frozenset({
    "len", "lenb", "left", "left$", "right", "right$", "mid", "mid$", "trim", "trim$", "ucase", "ucase$",
    "lcase", "lcase$", "instr", "int", "abs", "fix", "sgn", "val", "format", "format$", "hex", "hex$",
    "chr", "chr$", "asc", "space", "space$", "str", "str$", "cdate", "lbound",
})

# Intrinsics that raise 13 on a Missing second argument too.
_MISSING_RAISING_SECOND: frozenset[str] = frozenset({"instr"})

# Binary operators that raise 13 on a Missing operand (each measured).
_MISSING_RAISING_OPERATORS: frozenset[str] = frozenset({
    "+", "-", "*", "/", "\\", "^", "&", "=", "<", ">", "<=", ">=", "<>", "mod", "and", "or", "like",
})

# Conversions that raise 13 on a string that is not a number (CLng("") and
# CDbl("") measured).
_NUMBER_CONVERSIONS: frozenset[str] = frozenset(
    {"clng", "cint", "cdbl", "csng", "ccur", "cbyte", "clnglng", "clngptr", "cdec"}
)

# Statement heads after which the read cannot assume it runs on.
_LEAVING_HEADS: frozenset[str] = frozenset(
    {"exit", "goto", "gosub", "return", "end", "resume", "on", "stop", "error"}
)

_SUFFIX_TYPES: dict[str, str] = {
    "%": "integer", "&": "long", "^": "longlong", "!": "single", "#": "double", "@": "currency",
    "$": "string",
}


def _at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """toks[index], or None past either end (JavaScript's `toks[i]?.`)."""
    return toks[index] if 0 <= index < len(toks) else None


def _raw(toks: Sequence[VbaToken], index: int) -> str | None:
    tok = _at(toks, index)
    return tok.raw_text if tok is not None else None


def check_omitted_argument_reads(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    procedures: dict[str, ProcedureNode | None] = {}
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode) or member.proc_kind not in (ProcKind.SUB, ProcKind.FUNCTION):
            continue
        lower = member.name.lower()
        procedures[lower] = None if lower in procedures else member
    if not any(proc is not None and len(proc.params) > 0 for proc in procedures.values()):
        return lambda _member: None
    module_signatures = callable_type_signatures_for(symbols, None)
    reads: dict[tuple[str, str, str], _RaisingUse | _ParamArrayRead | None] = {}
    callee = _CalleeReader(source, symbols, activity, procedures, module_option_base(mod, activity))

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        values_at: Any = None
        # A dynamic array local is unallocated where a statement first names it
        # (XLIDE issue #449).
        proc_sym = procedure_symbol_for(symbols, member)
        declared: dict[str, VbaSymbol] = {
            child.name.lower(): child
            for child in (proc_sym.children if proc_sym is not None else None) or []
            if child.kind == VbaSymbolKind.LOCAL_VARIABLE
            and child.visibility != SymbolVisibility.STATIC
            and not child.is_auto_instantiated
        }
        first_named: dict[str, int] | None = None
        static_procedure = any(word.lower() == "static" for word in member.modifiers)

        def named_first_here(lower: str, offset: int) -> bool:
            nonlocal first_named
            if first_named is None:
                found: dict[str, int] = {}
                first_named = found

                def record(node: LeafStatementNode) -> None:
                    for tok in statement_tokens(source, node.span):
                        name = token_name(tok)
                        lowered = name.lower() if name else None
                        if lowered and lowered not in found:
                            found[lowered] = node.span.start

                for_each_statement_with_headers(source, member.body, record, activity)
            return first_named.get(lower) == offset

        def visitor(stmt: LeafStatementNode) -> None:
            # A literal the call passes, or a local whose value is known here.
            def argument_value(slot: Sequence[VbaToken]) -> _OmittedValue | None:
                nonlocal values_at
                toks = [tok for tok in slot if tok.kind is not TokenKind.COMMENT]
                if len(toks) == 1 and token_text(toks[0]) == "nothing":
                    return _OmittedValue("nothing")
                literal = _literal_value(toks)
                if literal is not None or len(toks) != 1 or toks[0].kind is not TokenKind.IDENTIFIER:
                    return literal
                lower = toks[0].raw_text.lower()
                local = declared.get(lower)
                if local is not None and named_first_here(lower, stmt.span.start) and not static_procedure:
                    # A never-set object passed is nothingPassedToMemberRead's (#343).
                    if local.is_array and local.array_bounds is None:
                        return _OmittedValue("unallocated")
                if values_at is None:
                    values_at = known_local_literal_values_at(source, member, symbols, activity)
                known = values_at(stmt).get(lower)
                if known is not None and known.kind == "number":
                    value = known.value
                    return _OmittedValue("number", value if not isinstance(value, str) else js_number(value))
                if known is not None and known.kind == "string":
                    return _OmittedValue("string", str(known.value))
                return None

            calls: list[CallArguments] = []
            statement_call = extract_call(source, stmt.span)
            if statement_call is not None:
                calls.append(statement_call)
            for call in expression_calls(source, stmt.span, module_signatures, source_names):
                if not any(other.name_span.start == call.name_span.start for other in calls):
                    calls.append(call)
            for branch in statement_and_branch_spans(stmt)[1:]:
                branch_call = extract_call(source, branch)
                if branch_call is not None and not any(
                    other.name_span.start == branch_call.name_span.start for other in calls
                ):
                    calls.append(branch_call)
            for call in calls:
                proc = None if call.qualifier else procedures.get(call.name.lower())
                if proc is None or bare_callable_source_shadowed(call.name, source_names):
                    continue
                for omitted in (
                    *_omitted_parameters(proc, call),
                    *_supplied_parameters(proc, call, argument_value),
                ):
                    key = (proc.name, omitted.param.name, omitted.key)
                    if key not in reads:
                        reads[key] = callee.first_raising_use(
                            proc, omitted.param, omitted.value, omitted.passed, omitted.skipped
                        )
                    use = reads[key]
                    if use is not None:
                        _report(source, proc, omitted, use, push)

        return visitor

    return factory


@dataclass(slots=True)
class _Omitted:
    param: ParameterNode
    # Where to report: the skipped slot, or the call's name.
    span: Span
    key: str
    # What the parameter holds; None for a ParamArray.
    value: _OmittedValue | None = None
    # For a ParamArray: how many values the call passes, and which it skips.
    passed: int = 0
    skipped: AbstractSet[int] = field(default_factory=frozenset)
    # The call passes the value rather than leaving it out (XLIDE issue #449).
    supplied: bool = False
    # What the call passes, where the parameter's type rounds it: 0.4 into a Long.
    rounded: int | float | None = None


_BRACKETS_RE = re.compile(r"^\[|\]$")


def _unbracketed_lower(text: str) -> str:
    return _BRACKETS_RE.sub("", text).lower()


def _supplied_parameters(
    proc: ProcedureNode,
    call: CallArguments,
    value_of: Callable[[Sequence[VbaToken]], _OmittedValue | None],
) -> list[_Omitted]:
    """The parameters a call passes a known value to, with what each holds in the
    callee: the value converted to the parameter's type as a ByVal copy is (XLIDE
    issue #449, measured in Excel 16.0: `F(0.4)` into a Long divides by 0)."""
    params = proc.params
    out: list[_Omitted] = []
    first_named = next((k for k, slot in enumerate(call.slots) if is_named_slot(slot)), -1)
    for k, slot in enumerate(call.slots):
        named = is_named_slot(slot)
        if not named and first_named >= 0 and k > first_named:
            continue
        param: ParameterNode | None
        if named:
            wanted = _unbracketed_lower(slot[0].raw_text)
            param = next((candidate for candidate in params if candidate.name.lower() == wanted), None)
        else:
            param = params[k] if k < len(params) else None
        if param is None or param.param_array or len(slot) == 0:
            continue
        passed = value_of(slot[2:] if named else slot)
        # An array parameter takes only an array: one never allocated is followed
        # in (XLIDE issue #449).
        if param.is_array != (passed is not None and passed.kind == "unallocated"):
            continue
        value = _held_as(passed, _parameter_type(param)) if passed is not None else None
        if passed is not None and value is not None:
            rounded = (
                passed.value
                if passed.kind == "number" and value.kind == "number" and passed.value != value.value
                else None
            )
            slot_span = call.slot_spans[k] if call.slot_spans is not None and k < len(call.slot_spans) else None
            out.append(
                _Omitted(
                    param=param,
                    value=value,
                    span=slot_span if slot_span is not None else call.name_span,
                    key=value.key(),
                    supplied=True,
                    rounded=rounded if not isinstance(rounded, str) else None,
                )
            )
    return out


_INTEGER_RANGES: dict[str, tuple[int, int]] = {
    "byte": (0, 255),
    "integer": (-32768, 32767),
    "long": (-2147483648, 2147483647),
}


def _held_as(value: _OmittedValue, type_name: str) -> _OmittedValue | None:
    """A passed value as a parameter of this type holds it, or None where the call
    itself would fail or it is not known."""
    if value.kind == "unallocated":
        return value
    if value.kind == "nothing":
        return value if type_name == "variant" or not is_known_scalar_type(type_name) else None
    if type_name == "variant":
        return value
    if value.kind == "string":
        return value if type_name == "string" else None
    if value.kind != "number" or isinstance(value.value, str) or value.value is None:
        return None
    number = value.value
    integer_range = _INTEGER_RANGES.get(type_name)
    if integer_range is not None:
        rounded = _round_half_even(number)
        return (
            _OmittedValue("number", rounded)
            if integer_range[0] <= rounded <= integer_range[1]
            else None
        )
    if type_name in ("double", "single", "currency"):
        return value
    return _OmittedValue("number", 0 if number == 0 else -1) if type_name == "boolean" else None


def _js_round(value: float) -> float:
    """Math.round: halves round up."""
    return math.floor(value + 0.5)


def _round_half_even(value: int | float) -> int | float:
    if not math.isfinite(value):
        return value
    floor = math.floor(value)
    diff = value - floor
    if diff != 0.5:
        return _js_round(value)
    return floor if floor % 2 == 0 else floor + 1


def _omitted_parameters(proc: ProcedureNode, call: CallArguments) -> list[_Omitted]:
    """The parameters a call leaves to their omitted value, and the ParamArray it fills."""
    params = proc.params
    named = [slot for slot in call.slots if is_named_slot(slot)]
    first_named = next((k for k, slot in enumerate(call.slots) if is_named_slot(slot)), -1)
    positional = call.slots if first_named < 0 else call.slots[:first_named]
    array_at = next((k for k, param in enumerate(params) if param.param_array), -1)
    fixed = len(params) if array_at < 0 else array_at
    # A call argument-count refuses is that rule's.
    if (array_at < 0 and len(positional) > len(params)) or (
        len(named) > 0 and array_at >= 0 and len(positional) > fixed
    ):
        return []
    named_lower = {_unbracketed_lower(slot[0].raw_text) for slot in named}
    out: list[_Omitted] = []
    for k in range(fixed):
        param = params[k]
        supplied = (k < len(positional) and len(positional[k]) > 0) or param.name.lower() in named_lower
        if supplied:
            continue
        if not param.optional:
            return []
        value = _omitted_value(param)
        if value is not None:
            skipped_slot = k < len(positional)
            slot_span = (
                call.slot_spans[k]
                if skipped_slot and call.slot_spans is not None and k < len(call.slot_spans)
                else None
            )
            out.append(
                _Omitted(
                    param=param,
                    value=value,
                    span=slot_span if slot_span is not None else call.name_span,
                    key=value.key(),
                )
            )
    if array_at >= 0 and len(named) == 0:
        rest = positional[fixed:]
        skipped = {j for j, slot in enumerate(rest) if len(slot) == 0}
        out.append(
            _Omitted(
                param=params[array_at],
                passed=len(rest),
                skipped=skipped,
                span=call.name_span,
                key=f"{len(rest)}:{','.join(str(j) for j in sorted(skipped))}",
            )
        )
    return out


def _omitted_value(param: ParameterNode) -> _OmittedValue | None:
    """What an omitted Optional holds: Missing, its type's default, or the value it declares."""
    type_name = _parameter_type(param)
    if param.default_raw is not None:
        toks = [tok for tok in raw_expression_tokens(param.default_raw) if tok.kind is not TokenKind.COMMENT]
        literal = _literal_value(toks)
        if literal is None:
            return None
        if type_name == "variant":
            return literal
        if type_name == "string":
            return literal if literal.kind == "string" else None
        return (
            literal
            if (is_numeric_type(type_name) or type_name == "boolean") and literal.kind == "number"
            else None
        )
    if type_name == "variant":
        return _OmittedValue("missing")
    if type_name == "string":
        return _OmittedValue("string", "")
    return _OmittedValue("number", 0) if is_numeric_type(type_name) or type_name == "boolean" else None


_NUMBER_SUFFIX_RE = re.compile(r"[%&^!#@]$")


def _literal_value(toks: Sequence[VbaToken]) -> _OmittedValue | None:
    if len(toks) == 1 and toks[0].kind is TokenKind.STRING_LITERAL:
        return _OmittedValue("string", string_literal_value(toks[0].raw_text))
    word = token_text(toks[0]) if len(toks) == 1 else ""
    if word in ("true", "false"):
        return _OmittedValue("number", -1 if word == "true" else 0)
    negative = len(toks) == 2 and toks[0].raw_text == "-"
    if len(toks) != (2 if negative else 1):
        return None
    number = toks[1 if negative else 0]
    if number.kind not in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL):
        return None
    value = js_number(_NUMBER_SUFFIX_RE.sub("", number.raw_text, count=1))
    if not math.isfinite(value):
        return None
    value = -value if negative else value
    return _OmittedValue("number", int(value) if float(value).is_integer() else value)


def _parameter_type(param: ParameterNode) -> str:
    if param.type_suffix:
        return _SUFFIX_TYPES.get(param.type_suffix, "")
    normalized = normalize_type(param.as_type)
    return normalized if normalized is not None else "variant"


def _number_text(value: int | float | str | None) -> str:
    """JavaScript's String(number)."""
    if isinstance(value, str) or value is None:
        return str(value)
    return js_number_to_string(float(value))


def _report(
    source: str,
    proc: ProcedureNode,
    omitted: _Omitted,
    use: _RaisingUse | _ParamArrayRead,
    push: PushFn,
) -> None:
    name = omitted.param.name
    where = f"line {_line_of(source, use.span.start)}"
    if isinstance(use, _ParamArrayRead):
        bound = omitted.passed - 1
        passes = (
            "no values"
            if omitted.passed == 0
            else "1 value"
            if omitted.passed == 1
            else f"{omitted.passed} values"
        )
        push(
            "arraySubscriptOutOfBounds",
            f"'{proc.name}' reads {name}({_number_text(use.index)}) ({where}), and this call passes "
            f"{passes} to that ParamArray, so its upper bound is {bound}. This will raise Run-time "
            "error '9': Subscript out of range.",
            omitted.span,
        )
        return
    value = omitted.value
    if value is None or value.kind == "missing":
        holds = "Missing"
    elif value.kind == "nothing":
        holds = "Nothing"
    elif value.kind == "unallocated":
        holds = "an array never allocated"
    elif value.kind == "string":
        holds = json.dumps(value.value, ensure_ascii=False)
    else:
        holds = _number_text(value.value)
    if omitted.supplied:
        passes = (
            f"{holds} to '{name}'"
            if omitted.rounded is None
            else f"{_number_text(omitted.rounded)} to '{name}', which holds {holds}"
        )
        push(
            use.rule,
            f"This call passes {passes}, and '{proc.name}' {use.does} ({where}). This will raise "
            f"Run-time error {use.error}.",
            omitted.span,
        )
        return
    subject = f"skips an element of '{name}'" if omitted.param.param_array else f"omits '{name}'"
    push(
        use.rule,
        f"This call {subject}, so it is {holds} in '{proc.name}', which {use.does} ({where}). "
        f"This will raise Run-time error {use.error}.",
        omitted.span,
    )


def _line_of(source: str, offset: int) -> int:
    return source.count("\n", 0, max(0, offset)) + 1


@dataclass(frozen=True, slots=True)
class _Use:
    toks: Sequence[VbaToken]
    index: int
    span_start: int


class _CalleeReader:
    """Reads a callee from its first line to the first use of a parameter."""

    def __init__(
        self,
        source: str,
        symbols: ModuleSymbols,
        activity: ConditionalActivityTracker | None,
        procedures: Mapping[str, ProcedureNode | None],
        option_base: int,
    ) -> None:
        self._source = source
        self._symbols = symbols
        self._activity = activity
        self._procedures = procedures
        self._option_base = option_base
        # By procedure start: one pass, one module.
        self._arrays: dict[int, Mapping[str, FixedArrayBound]] = {}

    def first_raising_use(
        self,
        proc: ProcedureNode,
        param: ParameterNode,
        value: _OmittedValue | None,
        passed: int,
        skipped: AbstractSet[int],
    ) -> _RaisingUse | _ParamArrayRead | None:
        lower = param.name.lower()
        use = self._first_use(proc.body, lower)
        if use is None:
            return None
        if param.param_array:
            return self._param_array_read(proc, use.toks, use.index, use.span_start, passed, skipped)
        if value is None:
            return None
        return self._classify(proc, use.toks, use.index, use.index, use.span_start, value, _parameter_type(param))

    def _is_inactive(self, node: BodyNode) -> bool:
        return self._activity is not None and self._activity.is_inactive(node.span)

    def _first_use(self, body: Sequence[BodyNode], lower: str) -> _Use | None:
        """The first use of `lower` on the path every call runs, or None when
        something before it may leave or the first use is on only some paths."""
        for node in body:
            if self._is_inactive(node) or isinstance(node, ConditionalDirectiveNode):
                continue
            if isinstance(node, VariableGroupNode):
                if self._mention_in(self._tokens(node.span), lower) >= 0:
                    return None
                continue
            if not is_leaf_statement(node):
                header = block_header_line_span(self._source, node.span)
                head_toks = self._tokens(header)
                at = self._mention_in(head_toks, lower)
                if at >= 0:
                    return _Use(head_toks, at, header.start)
                child = getattr(node, "body", None)
                if (
                    not isinstance(child, list)
                    or self._block_stops(child, lower)
                    or self._mention_in(
                        self._tokens(block_footer_line_span(self._source, node.span)), lower
                    )
                    >= 0
                ):
                    return None
                continue
            toks = self._tokens(node.span)
            if token_text(_at(toks, 0)) in _LEAVING_HEADS or _is_err_raise(toks):
                return None
            if isinstance(node, StatementNode) and node.single_line_if_branches:
                # The condition runs every time; the arms do not.
                then = next((k for k, tok in enumerate(toks) if token_text(tok) == "then"), -1)
                at = self._mention_in(toks if then < 0 else toks[:then], lower)
                if at >= 0:
                    return _Use(toks, at, node.span.start)
                if self._leaf_stops(node, lower):
                    return None
                continue
            at = self._mention_in(toks, lower)
            if at >= 0:
                return _Use(toks, at, node.span.start)
        return None

    def _block_stops(self, body: Sequence[BodyNode], lower: str) -> bool:
        """Whether a block's body uses the name or may leave."""
        for node in iter_body_nodes(body, self._is_inactive):
            if is_leaf_statement(node):
                if self._leaf_stops(node, lower):
                    return True
                continue
            if (
                self._mention_in(self._tokens(block_header_line_span(self._source, node.span)), lower) >= 0
                or self._mention_in(self._tokens(block_footer_line_span(self._source, node.span)), lower) >= 0
            ):
                return True
        return False

    def _leaf_stops(self, node: LeafStatementNode, lower: str) -> bool:
        for span in statement_and_branch_spans(node):
            toks = self._tokens(span)
            if (
                token_text(_at(toks, 0)) in _LEAVING_HEADS
                or _is_err_raise(toks)
                or self._mention_in(toks, lower) >= 0
            ):
                return True
        return False

    def _tokens(self, span: Span) -> list[VbaToken]:
        return statement_tokens_after_leading_label(self._source, span)

    @staticmethod
    def _mention_in(toks: Sequence[VbaToken], lower: str) -> int:
        """Where the statement names the parameter itself, not a member or a named argument."""
        for i, tok in enumerate(toks):
            name = token_name(tok)
            if (
                name is not None
                and name.lower() == lower
                and _raw(toks, i - 1) not in (".", "!")
                and _raw(toks, i + 1) != ":="
            ):
                return i
        return -1

    def _param_array_read(
        self,
        proc: ProcedureNode,
        toks: Sequence[VbaToken],
        at: int,
        span_start: int,
        passed: int,
        skipped: AbstractSet[int],
    ) -> _RaisingUse | _ParamArrayRead | None:
        # `ReDim p(3)` reads nothing: it is a compile error of its own, "Invalid
        # ParamArray use" (XLIDE issue #445).
        if _raw(toks, at + 1) != "(" or token_text(_at(toks, 0)) == "redim":
            return None
        close = match_paren_from(toks, at + 1)
        if close < 0:
            return None
        inner = toks[at + 2 : close]
        negative = len(inner) == 2 and inner[0].raw_text == "-"
        if (
            len(inner) != (2 if negative else 1)
            or inner[1 if negative else 0].kind is not TokenKind.INTEGER_LITERAL
            or _raw(toks, close + 1) == "("
        ):
            return None
        literal = inner[1 if negative else 0]
        number = js_number(re.sub(r"[%&^]$", "", literal.raw_text, count=1))
        index: int | float = (-1 if negative else 1) * number
        if not math.isnan(index) and float(index).is_integer():
            index = int(index)
        span = Span(span_start + toks[at].start, span_start + toks[close].end)
        if index < 0 or index >= passed:
            return _ParamArrayRead(index, span)
        if index in skipped:
            return self._classify(proc, toks, at, close, span_start, _OmittedValue("missing"))
        return None

    def _classify(
        self,
        proc: ProcedureNode,
        toks: Sequence[VbaToken],
        first: int,
        last: int,
        span_start: int,
        value: _OmittedValue,
        type_name: str = "variant",
    ) -> _RaisingUse | None:
        """Whether the use of the value at tokens first..last raises, and how."""
        prev = _at(toks, first - 1)
        nxt = _at(toks, last + 1)
        next_text = token_text(nxt)
        span = Span(span_start + toks[first].start, span_start + toks[last].end)
        if value.kind == "unallocated":
            # `a(1)`, `UBound(a)` or `LBound(a)` on an array never allocated (XLIDE
            # issue #449, measured in Excel 16.0). ReDim allocates it, and Erase of
            # an unallocated array runs.
            if token_text(_at(toks, 0)) in ("redim", "erase"):
                return None
            bound = (
                prev is not None
                and prev.raw_text == "("
                and token_text(_at(toks, first - 2)) in ("ubound", "lbound")
                and _raw(toks, first - 3) != "."
            )
            if next_text == "(" or bound:
                return _RaisingUse(
                    "unallocatedDynamicArrayAccess",
                    f"uses it in {_quote(self._operation_text(toks, _assignment_index(toks), span_start))}",
                    "'9': Subscript out of range",
                    span,
                )
            return None
        if value.kind == "nothing":
            # `c.Count` or `c(1)` on Nothing (XLIDE issue #449, measured in Excel 16.0).
            if next_text in ("(", ".", "!"):
                return _RaisingUse(
                    "objectVariableNotSet",
                    f"uses it in {_quote(self._operation_text(toks, _assignment_index(toks), span_start))}",
                    "'91': Object variable or With block variable not set",
                    span,
                )
            return None
        if next_text in ("(", ".", "!"):
            return None
        assign_at = _assignment_index(toks)
        if last + 1 == assign_at:
            return None  # the statement assigns it
        prev_operator = None if first - 1 == assign_at else _operator_word(prev, _at(toks, first - 2))
        next_operator = _operator_word(nxt, _at(toks, last))
        operation = _quote(self._operation_text(toks, assign_at, span_start))
        if value.kind == "missing":
            if prev_operator and (
                prev_operator in _MISSING_RAISING_OPERATORS or prev_operator in ("not", "unary")
            ):
                return _missing_use(span, f"uses it in {operation}")
            if next_operator and next_operator in _MISSING_RAISING_OPERATORS:
                return _missing_use(span, f"uses it in {operation}")
            if (
                first == assign_at + 1
                and last == len(toks) - 1
                and assign_at == (2 if token_text(_at(toks, 0)) == "let" else 1)
            ):
                target_tok = toks[assign_at - 1]
                target = (token_name(target_tok) or "").lower()
                declared_type = (
                    (_function_type(proc) if proc.proc_kind is ProcKind.FUNCTION else None)
                    if target == proc.name.lower()
                    else type_environment_for(self._symbols, proc).get(target)
                )
                normalized = normalize_type(declared_type)
                if normalized and is_known_scalar_type(normalized):
                    return _missing_use(span, f"assigns it to '{target_tok.raw_text}', a {_capitalize(normalized)}")
                return None
            if self._is_whole_condition(toks, first, last):
                return _missing_use(span, f"tests it in {_quote(self._head_text(toks))}")
            argument = _argument_of(toks, first, last)
            if argument is not None:
                callee_lower = argument.name
                procedure = self._procedures.get(callee_lower)
                if procedure is not None:
                    target_param = procedure.params[argument.position] if argument.position < len(procedure.params) else None
                    param_type = _parameter_type(target_param) if target_param is not None else "variant"
                    if (
                        target_param is not None
                        and not target_param.param_array
                        and target_param.by_val
                        and param_type != "variant"
                        and is_known_scalar_type(param_type)
                    ):
                        return _missing_use(
                            span,
                            f"passes it ByVal to the {_capitalize(param_type)} '{target_param.name}' of "
                            f"'{procedure.name}'",
                        )
                    return None
                if is_bare_or_vba_qualified_intrinsic_call(toks, argument.callee) and (
                    (argument.position == 0 and callee_lower in _MISSING_RAISING_INTRINSICS)
                    or (argument.position == 1 and callee_lower in _MISSING_RAISING_SECOND)
                ):
                    return _missing_use(span, f"passes it to {argument.display}")
            return None
        if value.kind == "number":
            number = value.value
            if (
                number == 0
                and prev_operator
                and prev_operator in ("/", "\\", "mod")
                and next_text != "^"
            ):
                return _RaisingUse("divisionByZero", f"divides by it in {operation}", "'11': Division by zero", span)
            if isinstance(number, str) or number is None:
                return None
            return self._number_use(proc, toks, first, last, assign_at, span_start, number, type_name)
        if not isinstance(value.value, str) or not is_invalid_numeric_string(value.value):
            return None

        def coerces(word: str | None) -> bool:
            return word is not None and word in ("-", "*", "/", "\\", "^", "mod", "unary", "not")

        def number_literal(tok: VbaToken | None) -> bool:
            return tok is not None and tok.kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL)

        if (
            coerces(prev_operator)
            or coerces(next_operator)
            or (prev_operator == "+" and number_literal(_at(toks, first - 2)))
            or (next_operator == "+" and number_literal(_at(toks, last + 2)))
        ):
            return _RaisingUse(
                "stringArithmeticCoercion", f"uses it as a number in {operation}", "'13': Type mismatch", span
            )
        argument = _argument_of(toks, first, last)
        if (
            argument is not None
            and argument.position == 0
            and argument.count == 1
            and argument.name in _NUMBER_CONVERSIONS
            and is_bare_or_vba_qualified_intrinsic_call(toks, argument.callee)
        ):
            return _RaisingUse(
                "runtimeConversionValue", f"converts it with {argument.display}", "'13': Type mismatch", span
            )
        return None

    def _number_use(
        self,
        proc: ProcedureNode,
        toks: Sequence[VbaToken],
        first: int,
        last: int,
        assign_at: int,
        span_start: int,
        value: int | float,
        type_name: str,
    ) -> _RaisingUse | None:
        """A number that raises where the callee uses it (XLIDE issue #449, each
        measured in Excel 16.0): Mid's start below 1 or Left's and Right's length
        below 0 (5), the index of a local fixed array outside its bounds (9), and
        `i * 2`, `2 * i` or `i + 1` past the range of the type they work in (6)."""
        span = Span(span_start + toks[first].start, span_start + toks[last].end)
        argument = _argument_of(toks, first, last)
        if (
            argument is not None
            and argument.position == 1
            and is_bare_or_vba_qualified_intrinsic_call(toks, argument.callee)
        ):
            start = argument.name in ("mid", "mid$")
            length = argument.name in ("left", "left$", "right", "right$")
            if (start and value < 1) or (length and value < 0):
                return _RaisingUse(
                    "runtimeArgumentValue",
                    f"passes it to {argument.display} as its {'start' if start else 'length'}",
                    "'5': Invalid procedure call or argument",
                    span,
                )
        if (
            argument is not None
            and argument.count == 1
            and math.isfinite(value)
            and float(value).is_integer()
            and _raw(toks, argument.callee + 1) == "("
            and _raw(toks, argument.callee - 1) != "."
        ):
            array = self._fixed_arrays(proc).get(token_text(toks[argument.callee]))
            dim = array.dims[0] if array is not None and len(array.dims) == 1 else None
            if dim is not None and (value < dim.lower or value > dim.upper):
                return _RaisingUse(
                    "arraySubscriptOutOfBounds",
                    f"reads {argument.display}({_number_text(value)}), whose bounds are "
                    f"{_number_text(dim.lower)} To {_number_text(dim.upper)}",
                    "'9': Subscript out of range",
                    span,
                )
        # The whole value of an assignment: `x op literal` or `literal op x`.
        integer_range = _INTEGER_RANGES.get(type_name)
        if (
            integer_range is None
            or assign_at < 0
            or len(toks) - assign_at != 4
            or (first != assign_at + 1 and first != len(toks) - 1)
        ):
            return None
        operator = toks[assign_at + 2].raw_text
        literal = toks[assign_at + 3 if first == assign_at + 1 else assign_at + 1]
        if (
            operator not in ("+", "-", "*")
            or literal.kind is not TokenKind.INTEGER_LITERAL
            or re.fullmatch(r"[0-9]+", literal.raw_text) is None
        ):
            return None
        other = int(literal.raw_text)
        # The literal is an Integer or a Long, and the wider type does the sum.
        other_range = (
            _INTEGER_RANGES["integer"]
            if other <= 32767
            else _INTEGER_RANGES["long"]
            if other <= 2147483647
            else None
        )
        if other_range is None:
            return None
        works = (
            other_range
            if other_range[1] > integer_range[1]
            else _INTEGER_RANGES["integer"]
            if type_name == "byte"
            else integer_range
        )
        a, b = (value, other) if first == assign_at + 1 else (other, value)
        result = a + b if operator == "+" else a - b if operator == "-" else a * b
        if works[0] <= result <= works[1]:
            return None
        type_label = "Long" if works is _INTEGER_RANGES["long"] else "Integer"
        return _RaisingUse(
            "arithmeticOverflow",
            f"computes {_number_text(result)} in "
            f"{_quote(self._operation_text(toks, assign_at, span_start))}, past the {type_label} range",
            "'6': Overflow",
            span,
        )

    def _fixed_arrays(self, proc: ProcedureNode) -> Mapping[str, FixedArrayBound]:
        found = self._arrays.get(proc.span.start)
        if found is None:
            found = local_fixed_arrays(self._source, proc, self._activity, self._option_base)
            self._arrays[proc.span.start] = found
        return found

    @staticmethod
    def _is_whole_condition(toks: Sequence[VbaToken], first: int, last: int) -> bool:
        """`If x Then`, `ElseIf x Then`, `Do While x`, `Loop Until x`, `While x`, `Select Case x`."""
        before = " ".join(token_text(tok) for tok in toks[:first])
        after = token_text(_at(toks, last + 1))
        if before in ("if", "elseif") and after == "then":
            return True
        return last == len(toks) - 1 and before in (
            "do while", "do until", "loop while", "loop until", "while", "select case",
        )

    @staticmethod
    def _head_text(toks: Sequence[VbaToken]) -> str:
        return " ".join(tok.raw_text for tok in toks)

    def _operation_text(self, toks: Sequence[VbaToken], assign_at: int, span_start: int) -> str:
        """The statement's value: what follows an assignment's `=`, or the whole statement."""
        then = next((k for k, tok in enumerate(toks) if token_text(tok) == "then"), -1)
        end = then - 1 if then > 0 else len(toks) - 1
        from_ = assign_at + 1 if 0 <= assign_at < end else 0
        return self._source[span_start + toks[from_].start : span_start + toks[end].end]


def _missing_use(span: Span, does: str) -> _RaisingUse:
    return _RaisingUse("variantValueMisuse", does, "'13': Type mismatch", span)


def _quote(text: str) -> str:
    """The text in quotes, cut to 60 UTF-16 units as upstream's `slice` cuts it."""
    units = text.encode("utf-16-le")
    if len(units) // 2 > 60:
        text = units[: 57 * 2].decode("utf-16-le", errors="surrogatepass") + "..."
    return f"'{text}'"


def _capitalize(type_name: str) -> str:
    names = {"longlong": "LongLong", "longptr": "LongPtr"}
    return names.get(type_name, type_name[:1].upper() + type_name[1:])


def _function_type(proc: ProcedureNode) -> str | None:
    if proc.type_suffix:
        return _SUFFIX_TYPES.get(proc.type_suffix)
    return proc.return_type if proc.return_type is not None else "variant"


def _operator_word(tok: VbaToken | None, before: VbaToken | None) -> str | None:
    """The operator a neighbouring token is, lowercased: a binary operator, `not`,
    or "unary" for a sign with no operand before it."""
    if tok is None:
        return None
    word = token_text(tok)
    if word == "not":
        return "not"
    if word in ("mod", "and", "or", "like"):
        return word if tok.kind in (TokenKind.KEYWORD, TokenKind.OPERATOR) else None
    if tok.kind is not TokenKind.OPERATOR:
        return None
    if word in ("-", "+") and before is not None and not _ends_operand(before):
        return "unary"
    if word in ("-", "+") and before is None:
        return "unary"
    return word


def _ends_operand(tok: VbaToken) -> bool:
    return (
        tok.kind
        in (
            TokenKind.IDENTIFIER,
            TokenKind.INTEGER_LITERAL,
            TokenKind.FLOAT_LITERAL,
            TokenKind.STRING_LITERAL,
            TokenKind.DATE_LITERAL,
        )
        or tok.raw_text == ")"
        or (
            tok.kind is TokenKind.KEYWORD
            and token_text(tok) in ("true", "false", "nothing", "empty", "null", "me")
        )
    )


@dataclass(frozen=True, slots=True)
class _ArgumentOf:
    callee: int
    name: str
    display: str
    position: int
    count: int


def _argument_of(toks: Sequence[VbaToken], first: int, last: int) -> _ArgumentOf | None:
    """The call whose argument the tokens first..last are, standing alone in their slot."""
    prev = _raw(toks, first - 1)
    nxt = _raw(toks, last + 1)
    if prev not in ("(", ",") or nxt not in (")", ","):
        return None
    depth = 0
    position = 0
    for i in range(first - 1, -1, -1):
        raw = toks[i].raw_text
        if raw == ")":
            depth += 1
        elif raw == "(":
            if depth == 0:
                # `Left$(` lexes as `Left` and `$`.
                callee = i - 2 if _raw(toks, i - 1) == "$" else i - 1
                callee_tok = _at(toks, callee)
                if token_name(callee_tok) is None or callee_tok is None:
                    return None
                suffix = "$" if callee == i - 2 else ""
                close = match_paren_from(toks, i)
                count = 1
                inner = 0
                for k in range(i + 1, close):
                    if toks[k].raw_text == "(":
                        inner += 1
                    elif toks[k].raw_text == ")":
                        inner -= 1
                    elif toks[k].raw_text == "," and inner == 0:
                        count += 1
                return _ArgumentOf(
                    callee, token_text(callee_tok) + suffix, callee_tok.raw_text + suffix, position, count
                )
            depth -= 1
        elif raw == "," and depth == 0:
            position += 1
    return None


# Heads whose `=` compares; any other statement's first top-level `=` assigns.
_COMPARING_HEADS: frozenset[str] = frozenset({"if", "elseif", "do", "loop", "while", "select", "case"})


def _assignment_index(toks: Sequence[VbaToken]) -> int:
    if token_text(_at(toks, 0)) in _COMPARING_HEADS:
        return -1
    depth = 0
    for i, tok in enumerate(toks):
        raw = tok.raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
        elif raw == "=" and depth == 0 and tok.kind is TokenKind.OPERATOR:
            return i
    return -1


def _is_err_raise(toks: Sequence[VbaToken]) -> bool:
    return token_text(_at(toks, 0)) == "err" and _raw(toks, 1) == "." and token_text(_at(toks, 2)) == "raise"
