"""Rule family: deterministic runtime argument / conversion values.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/runtimeValues.ts:
constant arguments that are provably outside a runtime function's accepted
range and conversions of provably invalid literals.
"""

from __future__ import annotations

import json
import math
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, replace
from decimal import Decimal

from ...conditional import ConditionalActivityTracker
from ...constants.date_literal import date_literal_serial
from ...constants.integer_constant_expression import (
    IntegerConstantLookup,
    bankers_round,
    evaluate_integer_constant_expression,
    parse_vba_integer_literal,
    resolve_raw_integer_constants,
)
from ...host.host_model import HostObjectModel
from ...js_compat import js_number_to_string, utf16_length
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import BodyNode, LeafStatementNode, ModuleNode, ProcedureNode, Span, is_leaf_statement
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureSignature, VbaSymbol, VbaSymbolKind
from ...types.type_inference import (
    known_local_literal_values_at,
    picked_values,
    procedure_symbol_for,
    type_environment_for,
)
from ...types.type_names import is_known_scalar_type, normalize_type
from ..argument_inference import infer_expression_type
from ..call_extraction import (
    CallableTypeSignature,
    empty_arg_split,
    named_argument_slot,
    split_arg_slots,
    string_literal_value,
    unwrap_outer_parens,
)
from ..callable_signatures import (
    SourceNameScope,
    callable_type_signatures_for,
    procedure_integer_constant_lookup,
    runtime_callable_source_shadowed,
    source_name_scope_for,
)
from ..const_expr import collect_module_literal_integer_constants
from ..context import PushFn, statement_tokens
from ..function_results import FunctionResult, function_integer_result, known_function_results
from ..known_locals import KnownLocalValue
from ..known_string_calls import (
    KnownStringCallContext,
    UtcDate,
    fold_known_string_calls,
    known_date,
    module_compare,
    parse_date_literal,
)
from ..loop_counters import CounterAtom, LoopCounter, check_each_counter_pass, loop_counters_at
from ..straight_line_values import ReachingAssignments, straight_line_assignments
from ..string_conversion import (
    is_invalid_boolean_string,
    is_invalid_date_string,
    is_invalid_numeric_string,
    is_invalid_time_string,
)
from ..type_fields import fixed_string_length, module_types
from ..walker import (
    ProcedureStatementVisitor,
    raw_expression_tokens,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .arrays import FixedArrayBound, StatementShapes, known_array_shapes_at, module_option_base, redim_shapes_at
from .shared import is_bare_or_vba_qualified_intrinsic_call

_NO_COUNTER_VALUES: Mapping[str, int | float] = {}

_RAISES_5 = "This will raise Run-time error '5': Invalid procedure call or argument."


@dataclass(frozen=True, slots=True)
class _LocalDeclaration:
    """A local's declared type, and its dimensions: 0 for a scalar, the count for a fixed array."""

    as_type: str
    dimensions: int


@dataclass(frozen=True, slots=True)
class _RuntimeArgumentValueSpec:
    canonical_name: str
    parameter_name: str
    argument_index: int
    minimum: int | None = None
    maximum: int | None = None
    # The value must be strictly above this: Log(0) raises, Log(0.5) runs.
    exclusive_minimum: int | None = None
    # Single values inside the range that still raise: InStrRev's Start of 0.
    disallowed: tuple[int, ...] | None = None
    # A Double parameter, compared as passed: Sqr(-0.4) raises. Others round first.
    fractional: bool = False
    # Whether a whole number is accepted, where no range says it: StrConv's Conversion.
    accepts: Callable[[int | float], bool] | None = None
    # An empty string literal raises: Asc(""), String(3, "").
    empty_string_raises: bool = False
    # A string literal must be one of these (case-insensitive): DateAdd's interval.
    allowed_strings: tuple[str, ...] | None = None
    minimum_slot_count: int | None = None
    allow_named: bool = True
    # Whether the function also has a `$`-suffixed spelling.
    string_suffix: bool = False
    # The parameter's type, where no signature gives it: a value outside it
    # raises error 6, Overflow, before any bound is read (XLIDE issue #218):
    # TimeSerial(32768, 0, 0), ChrB(256), String(1E+10, "a"). Without it, a
    # value past the Long range is left to argument-type-mismatch.
    overflow_type: str | None = None  # 'Byte' | 'Integer' | 'Long'
    # The bounds raise error 6 rather than 5: Error(65536) (issue #218).
    bounds_overflow: bool = False
    # InStr returns before it reads Start or Compare when either string is
    # empty: `InStr(0, "abc", "")` is 0 (issue #481, measured in Excel 16.0).
    # A Start past the Long range still overflows.
    skipped_by_empty_string: bool = False


@dataclass(frozen=True, slots=True)
class _RuntimeArgumentValueHit:
    display_name: str
    parameter_name: str
    value: int | float | str
    span: Span
    # 6 for Overflow; 5 otherwise (None).
    error: int | None = None
    # The whole message, for a check that is not one argument's bound.
    message: str | None = None


@dataclass(frozen=True, slots=True)
class _RuntimeArgumentValueCall:
    display_name: str
    specs: tuple[_RuntimeArgumentValueSpec, ...]
    slots: list[list[VbaToken]]


@dataclass(frozen=True, slots=True)
class _OutsideBounds:
    value: int | float | str
    span: Span
    error: int | None = None


@dataclass(frozen=True, slots=True)
class _MessageHit:
    message: str
    span: Span


_OVERFLOW_RANGES: dict[str, tuple[int, int]] = {
    "Byte": (0, 255),
    "Integer": (-32768, 32767),
    "Long": (-2147483648, 2147483647),
}


class _RuntimeValueLookup:
    """The IntegerConstantLookup upstream's checkRuntimeArgumentValues builds per
    procedure: a loop counter bound to one pass's value, a constant, a local with
    one known whole-number value, a Date local a straight line has just set to a
    date literal, then a Function of the module whose result is known."""

    __slots__ = ("_get",)

    def __init__(self, get: Callable[[str], int | float | None]) -> None:
        self._get = get

    def get(self, name: str, /) -> int | None:
        # Upstream's lookup hands back any number (a date serial may carry a
        # time fraction); the evaluator reads it as it is.
        return self._get(name)  # type: ignore[return-value]


def check_runtime_argument_values(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_integer_constants: Mapping[str, str | None] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    host_model: HostObjectModel | None = None,
) -> ProcedureStatementVisitor:
    """Some runtime-library arguments have deterministic value bounds even when the
    argument type itself is valid. This slice is VBE-oracle-backed for integer bounds
    on selected string runtime functions, which compile but raise Run-time error 5
    when the value is outside the proven range."""
    module_signatures = callable_type_signatures_for(symbols, project_procedures)
    project_constants = resolve_raw_integer_constants(project_integer_constants or {}, {})
    module_constants = collect_module_literal_integer_constants(mod, activity, project_constants)
    host_name = host_model.get("hostName") if host_model is not None else None
    host = host_name.lower() if host_name is not None else None
    compare = module_compare(source)
    types = module_types(source, mod, activity)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        constants = procedure_integer_constant_lookup(
            member, module_constants, symbols, project_visible_symbols, activity, host_model
        )
        # A local the procedure never assigns holds its default, and one whose
        # every assignment is one literal holds that (XLIDE issue #118): `Asc(s)`
        # with s never assigned is Asc(""), and `Mid(s, 5, 1) = "x"` after s =
        # "abc" starts past the end. Or the literal the last assignment before
        # the statement stores (issue #180).
        values_at = known_local_literal_values_at(source, member, symbols, activity)
        known: Mapping[str, KnownLocalValue] = {}
        # A local declared as a scalar, which `Join(n)` refuses (issue #239), or
        # as a fixed array, whose element type and dimensions Join reads.
        proc_symbol = procedure_symbol_for(symbols, member)
        locals_ = (proc_symbol.children if proc_symbol is not None else None) or []
        # Built on first use, as upstream's `??=` builds them.
        lazy: dict[str, object] = {}
        current_stmt: BodyNode | None = None
        counter_values: Mapping[str, int | float] = _NO_COUNTER_VALUES

        def reaching_assignments() -> Mapping[int, ReachingAssignments]:
            # The locals a straight line has just set (issue #364), by id(statement).
            if "reaching" not in lazy:
                lazy["reaching"] = straight_line_assignments(source, member.body, activity)
            return lazy["reaching"]  # type: ignore[return-value]

        def held_at(stmt: BodyNode, lower: str) -> list[VbaToken] | None:
            at = reaching_assignments().get(id(stmt))
            held = at.get(lower) if at is not None else None
            return None if held is None else [tok for tok in held if tok.kind is not TokenKind.COMMENT]

        def redim_shape(stmt: LeafStatementNode, lower: str) -> FixedArrayBound | None:
            if "shapes" not in lazy:
                lazy["shapes"] = redim_shapes_at(source, symbols, member, activity, module_option_base(mod, activity))
            shapes: StatementShapes = lazy["shapes"]  # type: ignore[assignment]
            at = shapes.get(stmt)
            return at.get(lower) if at is not None else None

        def known_shape(stmt: LeafStatementNode, lower: str) -> FixedArrayBound | None:
            if "known_shapes" not in lazy:
                lazy["known_shapes"] = known_array_shapes_at(
                    source, symbols, member, activity, module_option_base(mod, activity)
                )
            shapes_at: Callable[[LeafStatementNode], Mapping[str, FixedArrayBound]] = lazy["known_shapes"]  # type: ignore[assignment]
            return shapes_at(stmt).get(lower)

        def declaration_of(lower: str) -> _LocalDeclaration | None:
            local = next((child for child in locals_ if child.name.lower() == lower), None)
            if local is None or local.kind is not VbaSymbolKind.LOCAL_VARIABLE:
                return None
            type_ = normalize_type(local.as_type)
            if local.is_array and local.array_bounds is None:
                # A dynamic array may still be unallocated, which Join takes; one
                # a straight line has ReDim'd is judged as a fixed one (issue #342).
                stmt = current_stmt
                shape = redim_shape(stmt, lower) if stmt is not None and is_leaf_statement(stmt) else None
                return (
                    _LocalDeclaration(local.as_type if local.as_type is not None else "Variant", len(shape.dims))
                    if shape is not None
                    else None
                )
            if local.is_array:
                return _LocalDeclaration(
                    local.as_type if local.as_type is not None else "Variant",
                    len(split_top_level_token_groups(raw_expression_tokens(local.array_bounds or ""), 0, ",")),
                )
            if type_ is None or type_ == "variant":
                # A Variant holding a block's `.Value` has two dimensions (issue #492).
                stmt = current_stmt
                shape = known_shape(stmt, lower) if stmt is not None and is_leaf_statement(stmt) else None
                return (
                    _LocalDeclaration("Variant", len(shape.dims))
                    if shape is not None and len(shape.dims) > 1
                    else None
                )
            return _LocalDeclaration(local.as_type or "", 0) if is_known_scalar_type(type_) else None

        strings_for: dict[int, tuple[Mapping[str, KnownLocalValue], Mapping[str, str], Mapping[str, int]]] = {}

        def strings_at(values: Mapping[str, KnownLocalValue]) -> tuple[Mapping[str, str], Mapping[str, int]]:
            # Keyed by the map's identity, as upstream's Map is; the entry holds
            # the map so its id cannot be reused while cached.
            cached = strings_for.get(id(values))
            if cached is not None and cached[0] is values:
                return cached[1], cached[2]
            strings = picked_values(values, _plain_string)
            lengths = picked_values(values, _string_length)
            strings_for[id(values)] = (values, strings, lengths)
            return strings, lengths

        # A loop counter bound to one pass's value (issue #200).
        counters = loop_counters_at(source, member.body, activity)

        def lookup_get(name: str) -> int | float | None:
            counter = None if len(counter_values) == 0 else counter_values.get(name.lower())
            if counter is not None:
                return counter
            constant = constants.get(name)
            if constant is not None:
                return constant
            local = known.get(name.lower())
            if local is not None:
                value = local.value
                if local.kind == "number" and not isinstance(value, str) and _is_integer_value(value):
                    # Past 2**53 a JavaScript number is a float; keep it one.
                    return int(value) if abs(value) < 2**53 else value
                return None
            # A Date local a straight line has just set to a Date literal passes
            # its serial: Chr(d) with d = #1/2/2000# (issue #332).
            if current_stmt is not None and normalize_type(env.get(name.lower())) == "date":
                held = held_at(current_stmt, name.lower())
                if held is not None and len(held) == 1 and held[0].kind is TokenKind.DATE_LITERAL:
                    return date_literal_serial(held[0].raw_text)
            # `Mid(s, F())` with F a Function of the module returning 0 (issue #448).
            if "results" not in lazy:
                lazy["results"] = known_function_results(source, mod, activity)
            results: Mapping[str, FunctionResult] = lazy["results"]  # type: ignore[assignment]
            return function_integer_result(name, results, member, symbols)

        lookup = _RuntimeValueLookup(lookup_get)

        def visitor(stmt: LeafStatementNode) -> None:
            nonlocal known, current_stmt, counter_values
            known = values_at(stmt)
            current_stmt = stmt
            known_strings, known_string_lengths = strings_at(known)

            # The argument's type; for a Variant local, the type of what a
            # straight line has just put in it.
            def value_type(slot: Sequence[VbaToken] | None) -> str | None:
                # What a Variant holds gives its type; a typed local keeps its own:
                # `Dim a As Long: a = -0.5` holds 0, a Long (issue #664).
                def held(lower: str) -> list[VbaToken] | None:
                    if (normalize_type(env.get(lower)) or "variant") != "variant":
                        return None
                    return held_at(stmt, lower)

                return _static_value_type(
                    [tok for tok in (slot or []) if tok.kind is not TokenKind.COMMENT],
                    env,
                    module_signatures,
                    source_names,
                    source,
                    held,
                )

            def is_null_slot(slot: Sequence[VbaToken]) -> bool:
                value = [tok for tok in slot if tok.kind is not TokenKind.COMMENT]
                if len(value) != 1:
                    return False
                if token_text(value[0]) == "null":
                    return True
                name = token_name(value[0])
                held = held_at(stmt, name.lower()) if name else None
                return held is not None and len(held) == 1 and token_text(held[0]) == "null"

            # `Dim x As Double: x = -1.5: Chr(x)`: a local known to hold a number
            # that is not whole, or Empty, stands in as that literal (issue #332,
            # measured in Excel 16.0). A whole number already comes through the
            # lookup.
            def stand_in(slot: list[VbaToken]) -> list[VbaToken]:
                value = [tok for tok in slot if tok.kind is not TokenKind.COMMENT]
                name = token_name(value[0]) if len(value) == 1 else None
                lower = name.lower() if name is not None else None
                # `a = Empty` reads as the keyword, which the literal checks know.
                reached = held_at(stmt, lower) if lower else None
                if reached is not None and len(reached) == 1 and token_text(reached[0]) == "empty":
                    return [replace(value[0], kind=TokenKind.KEYWORD, raw_text="Empty", canonical_text=None)]
                held = known.get(lower) if lower else None
                if (
                    held is None
                    or held.content_mutated
                    or held.kind != "number"
                    or _is_integer_value(held.value)
                ):
                    return slot
                # A whole-number type rounds what it is given, half to even:
                # `a = 2.5` with a a Long holds 2, `a = 1.5` holds 2 (issues #664
                # and #673, measured in Excel 16.0).
                whole = (normalize_type(env.get(lower or "")) or "") in (
                    "byte", "integer", "long", "longlong", "longptr",
                )
                number = held.value
                assert not isinstance(number, str)
                return _literal_tokens_for(bankers_round(number) if whole else number, value[0])

            # A bound of Len(s) reads the length s has as the loop starts.
            def atom_value(atom: CounterAtom, counter: LoopCounter) -> int | None:
                if atom.kind != "len":
                    return None
                return strings_at(values_at(counter.loop_node))[1].get(atom.name)

            def date_of(lower: str) -> UtcDate | None:
                # A Date or Variant local a straight line has just given a date
                # literal (issue #559).
                held = held_at(stmt, lower)
                declared = normalize_type(env.get(lower))
                if (
                    held is not None
                    and len(held) == 1
                    and held[0].kind is TokenKind.DATE_LITERAL
                    and declared in ("date", "variant")
                ):
                    return parse_date_literal(held[0].raw_text)
                return None

            def check(values: Mapping[str, int | float], report: PushFn) -> None:
                nonlocal counter_values
                counter_values = values
                string_calls = KnownStringCallContext(
                    known_strings=known_strings,
                    integer_value=lambda text: evaluate_integer_constant_expression(text, lookup),
                    shadowed=lambda name: runtime_callable_source_shadowed(name, source_names),
                    compare=compare,
                    date_of=date_of,
                )
                for hit in _runtime_argument_value_hits(
                    source,
                    stmt.span,
                    module_signatures,
                    env,
                    lookup,
                    string_calls,
                    source_names,
                    host,
                    declaration_of,
                    is_null_slot,
                    value_type,
                    stand_in,
                ):
                    raises = "'6': Overflow" if hit.error == 6 else "'5': Invalid procedure call or argument"
                    report(
                        "runtimeArgumentValue",
                        hit.message
                        if hit.message is not None
                        else f"Argument '{hit.parameter_name}' of '{hit.display_name}' is "
                        f"{_display_value(hit.value)}; this will raise Run-time error {raises}.",
                        hit.span,
                    )

                # A fixed-length string is always its declared length (issue #248).
                def fixed_length_of(slot: Sequence[VbaToken]) -> int | None:
                    return fixed_string_length(slot, symbols, member, types, lookup)

                for message_hit in _runtime_statement_value_hits(
                    source, stmt.span, lookup, known_string_lengths, known_strings, source_names, fixed_length_of
                ):
                    report("runtimeArgumentValue", message_hit.message, message_hit.span)
                for message_hit in _date_diff_past_date_range(source, stmt.span, string_calls, lookup):
                    report("runtimeConversionValue", message_hit.message, message_hit.span)

            check_each_counter_pass(source, stmt.span, counters.get(stmt), atom_value, check, push)
            counter_values = _NO_COUNTER_VALUES

        return visitor

    return factory


def _truncated(value: int | float | None) -> int | None:
    """A count DateAdd reads: it drops the fraction, so DateAdd("d", -0.6, #1/1/100#) runs."""
    return None if value is None else math.trunc(value)


def _literal_tokens_for(value: int | float, at: VbaToken) -> list[VbaToken]:
    """Literal tokens for a number, at the place of the token they stand in for: `-1.5` is `-` and `1.5`."""
    magnitude = abs(value)
    literal = replace(
        at,
        kind=TokenKind.INTEGER_LITERAL if _is_integer_value(magnitude) else TokenKind.FLOAT_LITERAL,
        raw_text=_js_string(magnitude).upper(),
        canonical_text=None,
    )
    if value < 0:
        return [replace(at, kind=TokenKind.OPERATOR, raw_text="-", end=at.start, canonical_text=None), literal]
    return [literal]


def _date_diff_past_date_range(
    source: str,
    span: Span,
    string_calls: KnownStringCallContext,
    constants: IntegerConstantLookup,
) -> list[_MessageHit]:
    """`CDate(DateDiff("s", #2/29/2000#, t))`: DateDiff gives a Variant, and CDate
    of a Variant number past the date range raises 13, where a Long past it
    raises 6 (XLIDE issue #559, measured in Excel 16.0). The dates must be known."""
    toks = statement_tokens_after_leading_label(source, span)
    out: list[_MessageHit] = []
    for i in range(len(toks) - 1):
        if (
            token_text(toks[i]) != "cdate"
            or toks[i + 1].raw_text != "("
            or not is_bare_or_vba_qualified_intrinsic_call(toks, i)
            or string_calls.shadowed("cdate")
        ):
            continue
        close = match_paren_from(toks, i + 1)
        arg = toks[i + 2 : close] if close > i + 2 else []
        if not any(
            token_text(tok) == "datediff" and _raw_text_at(arg, k + 1) == "(" for k, tok in enumerate(arg)
        ):
            continue
        folded = fold_known_string_calls(arg, string_calls)
        value = None if folded is None else evaluate_integer_constant_expression(folded, constants)
        if value is None or -657434 <= value <= 2958465:
            continue
        out.append(
            _MessageHit(
                f"CDate cannot convert {_js_string(value)}, the Variant DateDiff gives here, to a date: it is "
                "past the range of dates. This will raise Run-time error '13': Type mismatch.",
                Span(span.start + arg[0].start, span.start + arg[-1].end),
            )
        )
    return out


def _runtime_statement_value_hits(
    source: str,
    span: Span,
    constants: IntegerConstantLookup,
    known_string_lengths: Mapping[str, int],
    known_strings: Mapping[str, str],
    source_names: SourceNameScope,
    fixed_length_of: Callable[[Sequence[VbaToken]], int | None] | None = None,
) -> list[_MessageHit]:
    """Statement and operator forms that raise for a value the code states (XLIDE
    issue #118, each measured in Excel 16.0):

     - `Err.Raise 0` and `Err.Raise 65536`, `Error 0`: error 5. A number is
       valid from 1 to 65535.
     - `(-8) ^ (1 / 3)` and `0 ^ -1`: error 5. A negative base takes only a
       whole exponent; zero takes only a non-negative one.
     - `"b" Like "[z-a]"` and `"b" Like "[a-"`: error 93, Invalid pattern
       string. A reversed range or an unterminated character list.
    """
    toks = statement_tokens(source, span)
    if _is_declaration_like_statement(toks):
        return []
    out: list[_MessageHit] = []
    first = _token_at(toks, 0)
    # `On n GoTo L1, L2` and `On n GoSub`: an index outside 0 to 255 raises 5,
    # and text that is no number raises 13 (issue #499, measured in Excel
    # 16.0). 0, or one past the last label, falls through to the next line.
    jump = next(
        (i for i, tok in enumerate(toks) if i > 1 and token_text(tok) in ("goto", "gosub")),
        -1,
    )
    if (
        token_text(first) == "on"
        and token_text(_token_at(toks, 1)) not in ("error", "local")
        and jump > 1
    ):
        index = [tok for tok in toks[1:jump] if tok.kind is not TokenKind.COMMENT]
        value = _integer_group_value(source, span, index, constants)
        index_name = token_name(index[0]) if len(index) == 1 else None
        name = index_name.lower() if index_name is not None else None
        if len(index) == 1 and index[0].kind is TokenKind.STRING_LITERAL:
            text: str | None = string_literal_value(index[0].raw_text)
        elif name is not None:
            text = known_strings.get(name)
        else:
            text = None
        index_span = Span(span.start + index[0].start, span.start + index[-1].end)
        shown = " ".join(tok.raw_text for tok in toks[1:jump])
        verb = "GoTo" if token_text(toks[jump]) == "goto" else "GoSub"
        if value is not None and (value < 0 or value > 255):
            out.append(
                _MessageHit(
                    f"On {verb} takes an index from 0 to 255, and {shown} is {_js_string(value)}. {_RAISES_5}",
                    index_span,
                )
            )
        elif value is None and text is not None and _ASCII_DIGIT.search(text) is None:
            out.append(
                _MessageHit(
                    f'On {verb} takes a number for its index, and {shown} holds "{text}". This will raise '
                    "Run-time error '13': Type mismatch.",
                    index_span,
                )
            )
    # `Mid(s, 5, 1) = "x"` with s holding "abc": the statement form starts past
    # the end of the string, error 5 (issue #118). Only the length matters,
    # which an earlier Mid statement cannot have changed. A fixed-length
    # string's length is its declaration's, assigned or not (issue #248).
    # `Mid$` lexes as Mid and a `$` of its own.
    mid_open = 2 if _raw_text_at(toks, 1) == "$" else 1
    # MidB counts bytes, two to a character: `MidB(s, 9, 1) = "x"` on "abc"
    # starts past its six (issue #327, measured in Excel 16.0).
    is_bytes = token_text(first) == "midb"
    if (token_text(first) == "mid" or is_bytes) and _raw_text_at(toks, mid_open) == "(":
        close = match_paren_from(toks, mid_open)
        if close > 0 and _raw_text_at(toks, close + 1) == "=":
            split = split_arg_slots(toks[mid_open + 1 : close], span.start)
            target_slot = split.slots[0] if split.slots else None
            target_name = token_name(target_slot[0]) if target_slot is not None and len(target_slot) == 1 else None
            target = target_name.lower() if target_name is not None else None
            fixed = (
                fixed_length_of(target_slot)
                if target_slot and fixed_length_of is not None
                else None
            )
            characters = (
                fixed if fixed is not None else known_string_lengths.get(target) if target is not None else None
            )
            length = characters * 2 if characters is not None and is_bytes else characters
            start_slot = split.slots[1] if len(split.slots) > 1 else None
            # An empty array is truthy in JavaScript, so `Mid(s, , 1) = "x"`
            # reaches integerGroupValue upstream and throws there;
            # _integer_group_value raises the same way.
            start = (
                _integer_group_value(source, span, start_slot, constants)
                if start_slot is not None
                else None
            )
            if length is not None and start is not None and start > length and target_slot is not None:
                form = "MidB" if is_bytes else "Mid"
                unit = "byte(s)" if is_bytes else "character(s)"
                message = (
                    f"{form} statement start {_js_string(start)} is past the end of "
                    f"{target_slot[0].raw_text}, which is {length} {unit} long. {_RAISES_5}"
                    if fixed is None
                    else f"{form} statement start {_js_string(start)} is past the end of "
                    f"{''.join(tok.raw_text for tok in target_slot)}, a fixed-length string of "
                    f"{length} {unit}. {_RAISES_5}"
                )
                out.append(
                    _MessageHit(
                        message,
                        split.spans[1] if len(split.spans) > 1 else _token_span(span, toks[0]),
                    )
                )
    # `Err.Raise n` and `Error n`.
    number_index = -1
    form = ""
    if token_text(first) == "err" and _raw_text_at(toks, 1) == "." and token_text(_token_at(toks, 2)) == "raise":
        number_index = 3
        form = "Err.Raise"
    elif (
        token_text(first) == "error"
        and len(toks) > 1
        and not runtime_callable_source_shadowed("Error", source_names)
    ):
        number_index = 1
        form = "Error"
    if number_index > 0:
        group = _number_argument_group(toks, number_index)
        number = _integer_group_value(source, span, group, constants) if group is not None else None
        # Err.Raise takes 1 to 65535 or any negative Long: `vbObjectError + 513`
        # and an HRESULT such as &H80004002 are how a class raises its own errors
        # (issue #142, measured). The Error statement takes only 1 to 65535.
        if form == "Err.Raise":
            invalid = number is not None and (number == 0 or number > 65535 or number < -2147483648)
        else:
            invalid = number is not None and (number < 1 or number > 65535)
        if invalid and number is not None and group is not None:
            valid = (
                "1 to 65535, or a negative Long such as vbObjectError + n" if form == "Err.Raise" else "1 to 65535"
            )
            out.append(
                _MessageHit(
                    f"{form} {_js_string(number)} is not an error number: valid numbers are {valid}. {_RAISES_5}",
                    Span(span.start + group[0].start, span.start + group[-1].end),
                )
            )
    for i in range(1, len(toks) - 1):
        tok = toks[i]
        if tok.kind is TokenKind.OPERATOR and tok.raw_text == "^":
            # A local known to hold a number, and True or False, count too:
            # `z ^ -1` with z never assigned, `0 ^ True` (issue #331, measured in
            # Excel 16.0).
            def named(operand: VbaToken | None, beside: VbaToken | None) -> int | None:
                word = token_text(operand)
                if word in ("true", "false"):
                    return -1 if word == "true" else 0
                beside_raw = beside.raw_text if beside is not None else None
                name = (
                    token_name(operand)
                    if operand is not None and beside_raw != "(" and beside_raw != "."
                    else None
                )
                return constants.get(name.lower()) if name else None

            before = None if _raw_text_at(toks, i - 2) == "." else named(toks[i - 1], None)
            operand_before = _numeric_operand_before(toks, i)
            base = operand_before if operand_before is not None else before
            operand_after = _numeric_operand_after(toks, i)
            exponent = (
                operand_after
                if operand_after is not None
                else named(_token_at(toks, i + 1), _token_at(toks, i + 2))
            )
            if base is not None and exponent is not None:
                if base < 0 and not _is_integer_value(exponent):
                    out.append(
                        _MessageHit(
                            "A negative number raised to the fractional power "
                            f"{_js_string(exponent)} has no real value. {_RAISES_5}",
                            _token_span(span, tok),
                        )
                    )
                elif base == 0 and exponent < 0:
                    out.append(
                        _MessageHit(
                            f"Zero raised to the negative power {_js_string(exponent)} divides by zero. "
                            f"{_RAISES_5}",
                            _token_span(span, tok),
                        )
                    )
            continue
        pattern_tok = toks[i + 1]
        if token_text(tok) == "like" and pattern_tok.kind is TokenKind.STRING_LITERAL:
            pattern = string_literal_value(pattern_tok.raw_text)
            problem = _invalid_like_pattern(_utf16_units(pattern))
            # The matcher meets a bad list only with a character left to compare,
            # so the string matched decides it (issue #193).
            subject = _like_subject(toks, i, known_strings)
            if problem and subject is not None and _like_reaches_bad_list(subject, pattern):
                out.append(
                    _MessageHit(
                        f"The Like pattern {pattern_tok.raw_text} {problem}, and matching "
                        f"{_json_string(subject)} reaches it. This will raise Run-time error '93': "
                        "Invalid pattern string.",
                        _token_span(span, pattern_tok),
                    )
                )
    # SaveSetting with an empty AppName, Section or Key, and DeleteSetting with
    # an empty AppName, raise 5 (issue #700, measured in Excel 16.0).
    settings = token_text(first)
    if settings == "savesetting":
        setting_names: tuple[str, ...] = ("AppName", "Section", "Key")
    elif settings == "deletesetting":
        setting_names = ("AppName",)
    else:
        setting_names = ()
    statement = "SaveSetting" if settings == "savesetting" else "DeleteSetting"
    if setting_names and not runtime_callable_source_shadowed(statement, source_names):
        open_index = 1 if _raw_text_at(toks, 1) == "(" else -1
        args = split_top_level_token_groups(
            toks,
            1 if open_index < 0 else 2,
            ",",
            len(toks) if open_index < 0 else match_paren_from(toks, open_index),
        )
        for k, setting_name in enumerate(setting_names):
            arg = (
                [tok for tok in args[k] if tok.kind is not TokenKind.COMMENT] if k < len(args) else None
            )
            if (
                arg is not None
                and len(arg) == 1
                and (arg[0].raw_text == '""' or token_text(arg[0]) == "vbnullstring")
            ):
                out.append(
                    _MessageHit(
                        f"Argument '{setting_name}' of '{statement}' is {arg[0].raw_text}; this will raise "
                        "Run-time error '5': Invalid procedure call or argument.",
                        _token_span(span, arg[0]),
                    )
                )
    return out


def _number_argument_group(toks: Sequence[VbaToken], index: int) -> list[VbaToken] | None:
    """The tokens of the first argument after `index`, up to a top-level comma or the end."""
    group: list[VbaToken] = []
    depth = 0
    for k in range(index, len(toks)):
        raw = toks[k].raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
        elif raw == "," and depth == 0:
            break
        if toks[k].kind is not TokenKind.COMMENT:
            group.append(toks[k])
    return group if group else None


def _integer_group_value(
    source: str, span: Span, group: Sequence[VbaToken], constants: IntegerConstantLookup
) -> int | None:
    # Upstream reads group[0].start unguarded, so an empty group throws there and
    # the shared statement walk stops; the IndexError here does the same.
    return evaluate_integer_constant_expression(
        source[span.start + group[0].start : span.start + group[-1].end], constants
    )


def _numeric_operand_before(toks: Sequence[VbaToken], index: int) -> int | float | None:
    """A numeric literal (optionally signed and parenthesized) right before `index`."""
    end = index - 1
    if _raw_text_at(toks, end) == ")":
        depth = 0
        start = end
        while start >= 0:
            if toks[start].raw_text == ")":
                depth += 1
            if toks[start].raw_text == "(":
                depth -= 1
                if depth == 0:
                    break
            start -= 1
        if start < 0:
            return None
        return _numeric_literal_group_value(toks[start + 1 : end])
    return _numeric_literal_group_value([toks[end]])


def _numeric_operand_after(toks: Sequence[VbaToken], index: int) -> int | float | None:
    """A numeric literal (optionally signed and parenthesized) right after `index`."""
    start = index + 1
    if _raw_text_at(toks, start) == "(":
        close = match_paren_from(toks, start)
        if close < 0:
            return None
        return _numeric_literal_group_value(toks[start + 1 : close])
    group: list[VbaToken] = []
    if _raw_text_at(toks, start) in ("-", "+"):
        group.append(toks[start])
        start += 1
    if start < len(toks):
        group.append(toks[start])
    return _numeric_literal_group_value(group)


def _numeric_literal_group_value(group: Sequence[VbaToken]) -> int | float | None:
    """The value of `[sign] literal`, or of `a / b` with literal operands."""
    toks = [tok for tok in group if tok.kind is not TokenKind.COMMENT]
    sign = 1
    rest = toks
    if rest and rest[0].raw_text in ("-", "+"):
        sign = -1 if rest[0].raw_text == "-" else 1
        rest = rest[1:]
    if len(rest) == 1:
        value = _numeric_literal(rest[0])
        return None if value is None else sign * value
    if len(rest) == 3 and rest[1].raw_text == "/":
        a = _numeric_literal(rest[0])
        b = _numeric_literal(rest[2])
        return None if a is None or b is None or b == 0 else sign * (a / b)
    return None


def _numeric_literal(tok: VbaToken) -> int | float | None:
    if tok.kind is TokenKind.INTEGER_LITERAL:
        return parse_vba_integer_literal(tok.raw_text)
    if tok.kind is TokenKind.FLOAT_LITERAL:
        return _float_literal_value(tok.raw_text)
    return None


_FLOAT_SUFFIX = re.compile(r"[!#@]\Z")


def _float_literal_value(raw: str) -> float | None:
    """Number() of a float literal without its type suffix, when finite. A D
    exponent (`1D3`) is NaN there, as float() refuses it here."""
    try:
        value = float(_FLOAT_SUFFIX.sub("", raw))
    except ValueError:
        return None
    return value if math.isfinite(value) else None


_LIKE_ALONE_AFTER = frozenset({"if", "elseif", "while", "until", "and", "or", "not", "then"})


def _like_subject(toks: Sequence[VbaToken], like_index: int, known_strings: Mapping[str, str]) -> str | None:
    """The string Like matches at `like_index`: a string literal, or a local the
    procedure makes plain, standing alone on its left."""
    operand = _token_at(toks, like_index - 1)
    before = _token_at(toks, like_index - 2)
    alone = (
        before is None
        or before.raw_text in ("(", ",", "=")
        or token_text(before) in _LIKE_ALONE_AFTER
    )
    if operand is None or not alone:
        return None
    if operand.kind is TokenKind.STRING_LITERAL:
        return string_literal_value(operand.raw_text)
    name = token_name(operand)
    return known_strings.get(name.lower()) if name else None


def _like_reaches_bad_list(subject_text: str, pattern_text: str) -> bool:
    """Whether matching `subject` against `pattern` reaches a malformed character
    list with a character left to compare, which is when Like raises 93 (XLIDE
    issue #193, measured in Excel 16.0): "xy" Like "?[" raises, "x" Like "?[" is
    False, and so is "zb" Like "a[z-a]", which fails at the "a". A `*` with a
    character left reaches a bad list anywhere after it: "b" Like "*[" raises
    (issue #336). A comparison Option Compare Text could decide otherwise proves
    nothing.

    Upstream walks both strings by UTF-16 code unit; so does this."""
    subject = _utf16_units(subject_text)
    pattern = _utf16_units(pattern_text)
    p = 0
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        matches: Callable[[str], bool | None]
        if ch == "[":
            close = _index_of(pattern, "]", i + 1)
            body = None if close < 0 else pattern[i + 1 : close]
            if body is None or _invalid_like_pattern(["[", *body, "]"]):
                return p < len(subject)
            negated = body[:1] == ["!"]
            char_list = body[1:] if negated else body

            def list_matches(c: str, char_list: list[str] = char_list, negated: bool = negated) -> bool | None:
                exact = _char_list_has(char_list, c)
                folded = _char_list_has(_units_lower(char_list), _unit_lower(c)) or _char_list_has(
                    _units_upper(char_list), _unit_upper(c)
                )
                if exact != folded:
                    return None
                return (not exact) if negated else exact

            matches = list_matches
            i = close
        elif ch == "*":
            # A `*` with a character left reaches a bad list anywhere after it:
            # "b" Like "*[" and "b" Like "*[a][" raise (issue #336).
            return p < len(subject) and _invalid_like_pattern(pattern[i + 1 :]) is not None
        elif ch == "?":

            def any_matches(_c: str) -> bool | None:
                return True

            matches = any_matches
        elif ch == "#":

            def digit_matches(c: str) -> bool | None:
                return "0" <= c <= "9"

            matches = digit_matches
        else:

            def char_matches(c: str, ch: str = ch) -> bool | None:
                if c == ch:
                    return True
                return None if _unit_lower(c) == _unit_lower(ch) else False

            matches = char_matches
        if p >= len(subject):
            return False
        verdict = matches(subject[p])
        if verdict is not True:
            return False
        p += 1
        i += 1
    return False


def _char_list_has(char_list: Sequence[str], c: str) -> bool:
    """Whether a character list's body, ranges included, holds `c`.

    `c` may be more than one code unit once case-mapped ('SS'), as upstream's
    toUpperCase gives; it then compares as a string."""
    k = 0
    while k < len(char_list):
        if k + 1 < len(char_list) and char_list[k + 1] == "-" and k + 2 < len(char_list):
            if char_list[k] <= c <= char_list[k + 2]:
                return True
            k += 2
        elif char_list[k] == c:
            return True
        k += 1
    return False


def _invalid_like_pattern(pattern: Sequence[str]) -> str | None:
    """Why a Like pattern (as UTF-16 code units) raises error 93, or None when it is well formed."""
    i = 0
    while i < len(pattern):
        if pattern[i] != "[":
            i += 1
            continue
        close = _index_of(pattern, "]", i + 1)
        if close < 0:
            return "opens a character list it never closes"
        body = list(pattern[i + 1 : close])
        if body[:1] == ["!"]:
            body = body[1:]
        for k in range(1, len(body) - 1):
            if body[k] == "-" and body[k - 1] > body[k + 1]:
                return f"has the reversed range {body[k - 1]}-{body[k + 1]}"
        i = close + 1
    return None


# The built-ins that return Null for a Null argument before checking the rest (issue #364).
_NULL_RETURNING = frozenset({"mid", "left", "right", "instr", "strcomp"})

# The first-argument types whose Round raises 5 past 22 digits (issue #402,
# measured in Excel 16.0): Round(3#, 23), Round("3", 23) and Round(#1/2/2000#,
# 23) raise; an Integer, Long, Byte, Boolean, Currency or Decimal runs at any
# count, Round(3, 256) and Round(CCur(3.5), 256).
_ROUND_DIGIT_LIMITED = frozenset({"double", "single", "string", "date"})

# The type a number literal's suffix gives it; none leaves an integer literal
# whole and a float a Double.
_LITERAL_SUFFIX_TYPES: dict[str, str] = {
    "#": "double", "!": "single", "@": "currency", "%": "integer", "&": "long", "^": "longlong",
}

_UNSETTLED_DIVISION_TYPES = frozenset({"currency", "decimal", "variant", "string", "date"})


def _static_value_type(
    toks: list[VbaToken],
    env: Mapping[str, str],
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope,
    source: str,
    held: Callable[[str], list[VbaToken] | None],
    depth: int = 0,
) -> str | None:
    """The VBA type of a value, lowercase, or None where the forms below do not
    settle it: a literal (by its suffix), a declared local (a Variant by what a
    straight line just put in it), a call (by its return type), and a `/` of such
    operands, which gives a Double. A Currency or Decimal operand of `/` is left
    unsettled."""
    value = unwrap_outer_parens(toks)
    signed = value[1] if len(value) == 2 and value[0].raw_text in ("-", "+") else None
    single = value[0] if len(value) == 1 else signed
    if single is not None:
        if single.kind is TokenKind.STRING_LITERAL:
            return "string"
        if single.kind is TokenKind.DATE_LITERAL:
            return "date"
        if single.kind is TokenKind.INTEGER_LITERAL or single.kind is TokenKind.FLOAT_LITERAL:
            suffix_type = _LITERAL_SUFFIX_TYPES.get(single.raw_text[-1:])
            if suffix_type is not None:
                return suffix_type
            return "integer" if single.kind is TokenKind.INTEGER_LITERAL else "double"
        text = token_text(single)
        if text in ("true", "false"):
            return "boolean"
        single_name = None if signed is not None else token_name(single)
        lower = single_name.lower() if single_name is not None else None
        declared = normalize_type(env.get(lower)) if lower else None
        if declared == "variant" and depth == 0 and lower:
            assigned = held(lower)
            return (
                _static_value_type(assigned, env, module_signatures, source_names, source, held, 1)
                if assigned
                else None
            )
        return declared
    if (
        len(value) >= 3
        and token_name(value[0])
        and value[1].raw_text == "("
        and match_paren_from(value, 1) == len(value) - 1
    ):
        inferred = infer_expression_type(value, 0, env, module_signatures, source_names, source=source)
        return normalize_type(inferred.type_ if inferred is not None else None)
    # Operands split at each top-level `/`; any other top-level operator leaves it unsettled.
    operands: list[list[VbaToken]] = [[]]
    parens = 0
    for tok in value:
        parens += 1 if tok.raw_text == "(" else -1 if tok.raw_text == ")" else 0
        if parens == 0 and tok.raw_text == "/":
            operands.append([])
        else:
            operands[-1].append(tok)
    if len(operands) < 2:
        return None
    for operand in operands:
        operand_type = (
            _static_value_type(operand, env, module_signatures, source_names, source, held, depth)
            if operand
            else None
        )
        if operand_type is None or operand_type in _UNSETTLED_DIVISION_TYPES:
            return None
    return "double"


def _no_declaration(_lower: str) -> _LocalDeclaration | None:
    return None


def _runtime_argument_value_hits(
    source: str,
    span: Span,
    module_signatures: Mapping[str, CallableTypeSignature],
    env: Mapping[str, str],
    constants: IntegerConstantLookup,
    string_calls: KnownStringCallContext,
    source_names: SourceNameScope,
    host: str | None,
    declaration_of: Callable[[str], _LocalDeclaration | None] | None = None,
    is_null_slot: Callable[[Sequence[VbaToken]], bool] = lambda _slot: False,
    value_type: Callable[[Sequence[VbaToken] | None], str | None] = lambda _slot: None,
    stand_in: Callable[[list[VbaToken]], list[VbaToken]] = lambda slot: slot,
) -> list[_RuntimeArgumentValueHit]:
    toks = statement_tokens(source, span)
    if _is_declaration_like_statement(toks):
        return []
    hits: list[_RuntimeArgumentValueHit] = []
    for i in range(len(toks) - 1):
        found = _runtime_argument_value_call_at(toks, i, span, module_signatures, env, source_names, host)
        if found is None:
            continue
        # A local known to hold a number is read as that number (issue #332).
        call = _RuntimeArgumentValueCall(
            found.display_name, found.specs, [list(stand_in(slot)) for slot in found.slots]
        )
        first_canonical = call.specs[0].canonical_name if call.specs else None
        # `Mid(Null, 0)`, `InStr(0, Null, "a")`: a Null argument makes the call
        # return Null before the others are checked (issue #364, measured in
        # Excel 16.0). String checks its Character for Null first: String(-1,
        # Null) is Null (issue #409, measured in Excel 16.0). A Null Start is no
        # position: InStr(v, "abc", "a") with v Null raises 94, where a Null
        # string hands Null back (issue #332, measured).
        if first_canonical == "InStr" and len(call.slots) >= 3 and is_null_slot(call.slots[0]):
            start = [tok for tok in call.slots[0] if tok.kind is not TokenKind.COMMENT]
            hits.append(
                _RuntimeArgumentValueHit(
                    call.display_name,
                    "Start",
                    "Null",
                    Span(span.start + start[0].start, span.start + start[-1].end),
                    message="InStr's Start is Null here, which is no position. This will raise Run-time "
                    "error '94': Invalid use of Null.",
                )
            )
            continue
        null_character = first_canonical == "String" and len(call.slots) > 1 and is_null_slot(call.slots[1])
        if null_character or (
            (first_canonical or "").lower() in _NULL_RETURNING and any(is_null_slot(slot) for slot in call.slots)
        ):
            continue
        for spec in call.specs:
            slot = _runtime_argument_value_slot(call.slots, spec)
            literal = (
                _integer_argument_outside_bounds(source, slot, span.start, spec, constants, string_calls)
                if slot is not None
                else None
            )
            if literal is None:
                continue
            if (
                spec.skipped_by_empty_string
                and literal.error != 6
                and any(_known_empty_string(text, string_calls.known_strings) for text in call.slots[1:3])
            ):
                continue
            # Round's digit limit follows the value's type, not its fraction:
            # Round(3#, 23) raises 5 and Round(3, 23) runs (issue #402). A count
            # below 0 always raises.
            if (
                spec.canonical_name == "Round"
                and not isinstance(literal.value, str)
                and literal.value > 0
                and (value_type(call.slots[0] if call.slots else None) or "") not in _ROUND_DIGIT_LIMITED
            ):
                continue
            hits.append(
                _RuntimeArgumentValueHit(
                    call.display_name,
                    spec.parameter_name,
                    literal.value,
                    literal.span,
                    error=6 if literal.error == 6 else None,
                )
            )
        overflow = _date_add_past_maximum(source, span, call, constants, string_calls)
        if overflow is None:
            overflow = _date_serial_past_maximum(source, span, call, constants)
        if overflow is None:
            overflow = _argument_relation_hit(source, span, call, constants, declaration_of or _no_declaration)
        if overflow is not None:
            hits.append(overflow)
    return hits


def _significant(slot: Sequence[VbaToken]) -> list[VbaToken]:
    return [t for t in slot if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE]


def _numeric_slot_value(
    source: str,
    span: Span,
    slot: Sequence[VbaToken] | None,
    constants: IntegerConstantLookup,
) -> int | float | None:
    """A numeric argument's value: a signed literal, `a / b` of literals, or a constant expression."""
    toks = unwrap_outer_parens(_significant(slot)) if slot is not None else []
    if not toks or named_argument_slot(toks) is not None:
        return None
    literal = _numeric_literal_group_value(toks)
    return literal if literal is not None else _integer_group_value(source, span, toks, constants)


_RELATION_ERROR_TEXT = {11: "Division by zero", 9: "Subscript out of range", 13: "Type mismatch"}


def _argument_relation_hit(
    source: str,
    span: Span,
    call: _RuntimeArgumentValueCall,
    constants: IntegerConstantLookup,
    declaration_of: Callable[[str], _LocalDeclaration | None] = _no_declaration,
) -> _RuntimeArgumentValueHit | None:
    """Checks that read more than one argument, or an argument's kind rather than
    its bound (XLIDE issue #218, each measured in Excel 16.0):

     - Partition(Number, Start, Stop, Interval) raises 5 when Start is below 0,
       Stop is not above Start, or Interval is below 1, each rounded half to
       even: Partition(5, 0, 10, 0.6) runs, 0.4 raises.
     - The financial functions raise 5: Pmt with NPer 0; IPmt and PPmt with NPer
       or Per not above 0, or Per a whole period past NPer (Per 10.5 of 10 runs,
       11 raises); SLN with Life 0; SYD and DDB with Life or Period not above 0,
       or Period past Life; DDB with Factor not above 0; NPer where its log has
       no value (Rate at or below -1, Rate and Pmt both 0, or a ratio not above
       0); Rate with NPer not above 0. PV with Rate -1 divides by zero, error 11.
       Whether Rate's iteration converges otherwise is not judged: Rate(10, 100,
       1000) raises and Rate(9, 100, 1000) returns -1.73.
     - LBound and UBound of Array(...) or Split(...), which have one dimension,
       raise 9 for any other Dimension.
     - Join and Filter given a string or number where the array goes raise 13.
     - Switch with an odd number of arguments raises 5 (issue #219).
    """
    name = _strip_vba_prefix(call.display_name)
    name = (name[:-1] if name.endswith("$") else name).lower()
    if any(named_argument_slot(slot) is not None for slot in call.slots):
        return None

    def value(index: int) -> int | float | None:
        return _numeric_slot_value(source, span, call.slots[index] if index < len(call.slots) else None, constants)

    def slot_span(start: int, end: int | None = None) -> Span:
        # Upstream asserts both ends exist; an empty slot throws there, as the
        # IndexError does here.
        first = _significant(call.slots[start])[0]
        last = _significant(call.slots[start if end is None else end])[-1]
        return Span(span.start + first.start, span.start + last.end)

    def hit(message: str, at: Span, error: int = 5) -> _RuntimeArgumentValueHit:
        text = _RELATION_ERROR_TEXT.get(error, "Invalid procedure call or argument")
        return _RuntimeArgumentValueHit(
            call.display_name,
            "",
            "",
            at,
            message=f"{message} This will raise Run-time error '{error}': {text}.",
        )

    def present(count: int) -> bool:
        return len(call.slots) >= count and all(len(_significant(slot)) > 0 for slot in call.slots[:count])

    s = _js_string
    if name == "partition":
        if not present(4):
            return None
        start, stop, interval = (None if v is None else bankers_round(v) for v in (value(1), value(2), value(3)))
        if start is not None and start < 0:
            return hit(f"Partition's Start is {s(start)}; it must be 0 or more.", slot_span(1))
        if start is not None and stop is not None and stop <= start:
            return hit(f"Partition's Stop, {s(stop)}, is not above its Start, {s(start)}.", slot_span(1, 2))
        if interval is not None and interval < 1:
            return hit(f"Partition's Interval is {s(interval)}; it must be 1 or more.", slot_span(3))
        return None
    if name == "pmt":
        return (
            hit("Pmt over 0 periods (NPer 0) has no payment.", slot_span(1))
            if present(3) and value(1) == 0
            else None
        )
    if name in ("ipmt", "ppmt"):
        if not present(4):
            return None
        per = value(1)
        nper = value(2)
        label = call.display_name
        if nper is not None and nper <= 0:
            return hit(f"{label}'s NPer is {s(nper)}; it must be above 0.", slot_span(2))
        if per is not None and per <= 0:
            return hit(f"{label}'s Per is {s(per)}; it must be above 0.", slot_span(1))
        if per is not None and nper is not None and per >= nper + 1:
            return hit(f"{label}'s Per, {s(per)}, is past the last of its {s(nper)} periods.", slot_span(1))
        return None
    if name == "sln":
        return (
            hit("SLN over a Life of 0 has no depreciation.", slot_span(2))
            if present(3) and value(2) == 0
            else None
        )
    if name in ("syd", "ddb"):
        if not present(4):
            return None
        life = value(2)
        period = value(3)
        label = call.display_name
        if life is not None and life <= 0:
            return hit(f"{label}'s Life is {s(life)}; it must be above 0.", slot_span(2))
        if period is not None and period <= 0:
            return hit(f"{label}'s Period is {s(period)}; it must be above 0.", slot_span(3))
        if period is not None and life is not None and period > life:
            return hit(f"{label}'s Period, {s(period)}, is past its Life, {s(life)}.", slot_span(3))
        if name == "ddb" and present(5):
            factor = value(4)
            if factor is not None and factor <= 0:
                return hit(f"DDB's Factor is {s(factor)}; it must be above 0.", slot_span(4))
        return None
    if name == "nper":
        if not present(3) or len(call.slots) > 5:
            return None
        rate, pmt, pv = value(0), value(1), value(2)
        fv = value(3) if len(call.slots) >= 4 and present(4) else 0
        type_ = value(4) if len(call.slots) >= 5 and present(5) else 0
        if rate is None or pmt is None or pv is None or fv is None or type_ is None:
            return None
        if rate == 0:
            return (
                hit("NPer with a Rate and a Pmt of 0 has no number of periods.", slot_span(0, 1))
                if pmt == 0
                else None
            )
        if rate <= -1:
            return hit(
                f"NPer's Rate is {s(rate)}; the logarithm of 1 + Rate has no value at or below -1.", slot_span(0)
            )
        a = _js_divide(float(pmt) * (1 + float(rate) * (1 if type_ != 0 else 0)), float(rate))
        ratio = _js_divide(a - float(fv), a + float(pv))
        if math.isfinite(ratio) and ratio > 0:
            return None
        return hit(
            "No number of periods brings these payments to this value: the logarithm NPer takes has no value.",
            slot_span(0, min(len(call.slots), 5) - 1),
        )
    if name == "rate":
        nper = value(0) if present(3) else None
        return (
            hit(f"Rate's NPer is {s(nper)}; it must be above 0.", slot_span(0))
            if nper is not None and _is_integer_value(nper) and nper <= 0
            else None
        )
    if name == "pv":
        if not present(3):
            return None
        rate, nper = value(0), value(1)
        return (
            hit("PV with a Rate of -1 divides by (1 + Rate) ^ NPer, which is 0.", slot_span(0), 11)
            if rate == -1 and nper is not None and nper > 0
            else None
        )
    if name in ("lbound", "ubound"):
        if len(call.slots) != 2 or not present(2):
            return None
        array = _significant(call.slots[0])
        callee_name = token_name(array[0]) if array else None
        callee = callee_name.lower() if callee_name is not None else None
        one_dimension = (
            callee in ("array", "split")
            and _raw_text_at(array, 1) == "("
            and match_paren_from(array, 1) == len(array) - 1
        )
        dimension = value(1)
        return (
            hit(
                f"{call.display_name}'s Dimension is {s(dimension)}, but {array[0].raw_text}(...) has one "
                "dimension.",
                slot_span(1),
                9,
            )
            if one_dimension and dimension is not None and bankers_round(dimension) != 1
            else None
        )
    if name == "switch":
        # Switch takes condition-value pairs; an odd count compiles and raises 5
        # whatever the conditions are (issue #219).
        return (
            hit(
                "Switch takes its arguments in pairs, a condition and a value, but is given "
                f"{len(call.slots)}.",
                slot_span(0, len(call.slots) - 1),
            )
            if len(call.slots) % 2 == 1
            else None
        )
    if name in ("join", "filter"):
        # Join needs only its array (issue #239): `Join(5)`, `Join(Null)` and
        # Join of a Long or String local raise 13, as Filter of one does (issue
        # #242).
        join = name == "join"
        if not present(1 if join else 2):
            return None
        first = _significant(call.slots[0])
        scalar = len(first) == 1 and (
            first[0].kind in (TokenKind.STRING_LITERAL, TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL)
            or (
                join
                and (
                    first[0].kind is TokenKind.DATE_LITERAL
                    or token_text(first[0]) in ("null", "true", "false")
                )
            )
        )
        if scalar:
            return hit(f"{call.display_name} takes an array, but {first[0].raw_text} is not one.", slot_span(0), 13)
        declared = (
            declaration_of(first[0].raw_text.lower())
            if len(first) == 1 and first[0].kind is TokenKind.IDENTIFIER
            else None
        )
        if declared is None:
            return None
        if declared.dimensions == 0:
            return hit(
                f"{call.display_name} takes an array, but '{first[0].raw_text}' is declared As "
                f"{declared.as_type}.",
                slot_span(0),
                13,
            )
        # Join reads one dimension of Strings or Variants: an array of Long, or
        # of two dimensions, raises 5 (measured in Excel 16.0). Filter refuses
        # the same arrays with 13 (issue #242).
        verb = "joins" if join else "filters"
        error = 5 if join else 13
        if declared.dimensions > 1:
            return hit(
                f"{call.display_name} {verb} an array of one dimension, but '{first[0].raw_text}' has "
                f"{declared.dimensions}.",
                slot_span(0),
                error,
            )
        element = normalize_type(declared.as_type)
        return (
            hit(
                f"{call.display_name} {verb} Strings or Variants, but '{first[0].raw_text}' is an array of "
                f"{declared.as_type}.",
                slot_span(0),
                error,
            )
            if element != "string" and element != "variant"
            else None
        )
    return None


_MS_PER_DAY = 86400000
_DATE_ADD_UNITS: dict[str, int] = {
    "d": _MS_PER_DAY, "y": _MS_PER_DAY, "w": _MS_PER_DAY, "ww": 7 * _MS_PER_DAY,
    "h": 3600000, "n": 60000, "s": 1000,
}
_DATE_ADD_MONTHS: dict[str, int] = {"m": 1, "q": 3, "yyyy": 12}
_PAST_MAXIMUM = "past December 31, 9999"
_BEFORE_MINIMUM = "before January 1, 100"


def _date_add_past_maximum(
    source: str,
    span: Span,
    call: _RuntimeArgumentValueCall,
    constants: IntegerConstantLookup,
    string_calls: KnownStringCallContext | None = None,
) -> _RuntimeArgumentValueHit | None:
    """`DateAdd("d", 1, #12/31/9999#)` and `DateAdd("yyyy", -1, #1/1/100#)`: adding
    to a date past the last date VBA has (December 31, 9999) or before the first
    (January 1, 100) raises error 5 (XLIDE issues #118 and #262, measured in
    Excel 16.0 for every interval). The date is a literal or a DateSerial of
    literals from year 100 on; the count is a whole number. Months, quarters and
    years keep the day where the month has it, so only the month they reach
    decides."""
    if call.display_name != "DateAdd" or len(call.slots) < 3:
        return None
    interval_slot, number_slot, date_slot = (
        unwrap_outer_parens([t for t in slot if t.kind is not TokenKind.COMMENT]) for slot in call.slots[:3]
    )
    if len(interval_slot) != 1 or interval_slot[0].kind is not TokenKind.STRING_LITERAL or not date_slot:
        return None
    interval = string_literal_value(interval_slot[0].raw_text).lower()
    # A count or a date another call gives: `Year(#6/15/5000#)`,
    # `DateValue("12/31/9999")` (issue #510, measured in Excel 16.0).
    folded = fold_known_string_calls(number_slot, string_calls) if string_calls is not None and number_slot else None
    count: int | float | None
    if folded is not None:
        count = evaluate_integer_constant_expression(folded, constants)
    else:
        count = _integer_group_value(source, span, number_slot, constants)
        if count is None:
            count = _truncated(_numeric_literal_group_value(number_slot))
    if count is None or count == 0:
        return None
    # The count is held as a Long: 2147483648 of any interval overflows (issue
    # #332, measured in Excel 16.0).
    if abs(count) > 2147483647:
        first = number_slot[0]
        last = number_slot[-1]
        return _RuntimeArgumentValueHit(
            "DateAdd",
            "Number",
            _js_string(count),
            Span(span.start + first.start, span.start + last.end),
            error=6,
            message=f"DateAdd counts its intervals in a Long, and "
            f"{source[span.start + first.start : span.start + last.end]} is outside it. This will raise "
            "Run-time error '6': Overflow.",
        )
    if len(date_slot) == 1 and date_slot[0].kind is TokenKind.DATE_LITERAL:
        date = _epoch_ms(parse_date_literal(date_slot[0].raw_text))
    else:
        date = _date_serial_of_literals(source, span, date_slot, constants)
        if date is None:
            date = _epoch_ms(
                known_date(
                    date_slot,
                    lambda group: _integer_group_value(source, span, group, constants),
                    string_calls.date_of if string_calls is not None else None,
                )
            )
    if date is None:
        return None
    whole = int(count)
    past: str | None = None
    unit = _DATE_ADD_UNITS.get(interval)
    months = _DATE_ADD_MONTHS.get(interval)
    if unit is not None:
        result = date + whole * unit
        past = (
            _PAST_MAXIMUM
            if result >= _utc_ms(10000, 1, 1)
            else _BEFORE_MINIMUM
            if result < _utc_ms(100, 1, 1)
            else None
        )
    elif months is not None:
        # The months are counted in a Long that wraps: DateAdd("yyyy", 2147483647,
        # #1/1/2000#) gives 1/1/1999 (issue #332, measured).
        year, month, _day = _civil_from_days(date // _MS_PER_DAY)
        moved = whole * months
        remainder = abs(moved) % 2**32 * (1 if moved >= 0 else -1)
        month_number = year * 12 + (month - 1) + (remainder + 2**32 + 2**31) % 2**32 - 2**31
        past = (
            _PAST_MAXIMUM
            if month_number > 9999 * 12 + 11
            else _BEFORE_MINIMUM
            if month_number < 100 * 12
            else None
        )
    if past is None:
        return None
    first = date_slot[0]
    last = date_slot[-1]
    return _RuntimeArgumentValueHit(
        "DateAdd",
        "Date",
        f"{source[span.start + first.start : span.start + last.end]}, which the {_js_string(count)} {interval} "
        f"interval(s) carry {past}",
        Span(span.start + first.start, span.start + last.end),
    )


def _date_serial_of_literals(
    source: str, span: Span, toks: Sequence[VbaToken], constants: IntegerConstantLookup
) -> int | None:
    """`DateSerial(y, m, d)` of integer literals, year 100 or later, as UTC milliseconds."""
    index = 0
    if token_text(_token_at(toks, 0)) == "vba" and _raw_text_at(toks, 1) == ".":
        index = 2
    if (
        token_text(_token_at(toks, index)) != "dateserial"
        or _raw_text_at(toks, index + 1) != "("
        or match_paren_from(toks, index + 1) != len(toks) - 1
    ):
        return None
    parts = [
        _integer_group_value(source, span, group, constants)
        for group in split_top_level_token_groups(toks[index + 2 : len(toks) - 1], 0, ",")
    ]
    if len(parts) != 3:
        return None
    year, month, day = parts
    if year is None or month is None or day is None:
        return None
    if any(part < -32768 or part > 32767 for part in (year, month, day)):
        return None
    if year < 100:
        return None  # read as 19xx or 20xx
    return _utc_date_ms(year, month - 1, day)


def _date_serial_past_maximum(
    source: str,
    span: Span,
    call: _RuntimeArgumentValueCall,
    constants: IntegerConstantLookup,
) -> _RuntimeArgumentValueHit | None:
    """`DateSerial(9999, 13, 1)`: the month and day carry into the year, and a date
    past December 31, 9999 raises error 5 (XLIDE issue #189, measured in Excel
    16.0). The year alone decides nothing: DateSerial(10000, 0, 1) runs and is
    December 1, 9999. A year from 0 to 99, after the month carries, is read as
    19xx or 20xx, and is judged only where both readings agree."""
    if _strip_vba_prefix(call.display_name).lower() != "dateserial" or len(call.slots) != 3:
        return None
    parts: list[int | None] = []
    for slot in call.slots:
        toks = [t for t in slot if t.kind is not TokenKind.COMMENT]
        parts.append(None if not toks else _integer_group_value(source, span, toks, constants))
    year, month, day = parts
    if year is None or month is None or day is None:
        return None
    # A part past the Integer range overflows first (issue #218).
    if any(part < -32768 or part > 32767 for part in (year, month, day)):
        return None
    # VBA carries the month into the year first, then reads the year: 0 to 99
    # as 19xx or 20xx by the system's cut, below 0 counted from 2000. A year past
    # 9999 then raises 5 before the day is added; otherwise the day carries from
    # the first of the month, and the date must fall from January 1, 100 to
    # December 31, 9999. DateSerial(100, 0, 1) is year 99, December, and runs;
    # DateSerial(9999, 13, 0) raises 5 (issues #559 and #604, measured in Excel
    # 16.0 over a 1,716-call grid).
    carried = year + (month - 1) // 12
    month_index = (month - 1) % 12
    before = False
    if carried <= 9999:
        if carried < 0:
            readings = [carried + 2000]
        elif carried < 100:
            readings = [carried + 1900, carried + 2000]
        else:
            readings = [carried]
        outside: list[tuple[bool, bool]] = []
        for reading in readings:
            reached_year = _civil_from_days(_utc_date_ms(reading, month_index, day) // _MS_PER_DAY)[0]
            early = reached_year < 100
            outside.append((early, early or reached_year > 9999))
        # A two-digit year is judged only where both readings agree.
        if not all(is_outside for _early, is_outside in outside):
            return None
        before = outside[0][0]
    first = next(t for t in call.slots[0] if t.kind is not TokenKind.COMMENT)
    last = next(t for t in reversed(call.slots[2]) if t.kind is not TokenKind.COMMENT)
    return _RuntimeArgumentValueHit(
        call.display_name,
        "Year",
        f"{year} with month {month} and day {day}, a date {_BEFORE_MINIMUM if before else _PAST_MAXIMUM}",
        Span(span.start + first.start, span.start + last.end),
    )


def _runtime_argument_value_call_at(
    toks: Sequence[VbaToken],
    index: int,
    span: Span,
    module_signatures: Mapping[str, CallableTypeSignature],
    env: Mapping[str, str],
    source_names: SourceNameScope,
    host: str | None,
) -> _RuntimeArgumentValueCall | None:
    name = token_name(toks[index])
    if not name:
        return None
    # Only a bare `Left(...)` or a genuine `VBA.Left(...)` is the intrinsic. The
    # shared helper also rejects `obj.vba.Left(...)` via the third-token ('.')
    # check that this ad-hoc logic previously omitted.
    if not is_bare_or_vba_qualified_intrinsic_call(toks, index):
        return None
    # A bare call can be shadowed by a source symbol; a `VBA.`-qualified one
    # cannot, so only the bare form participates in the shadow gate below.
    qualifier = token_name(toks[index - 2]) if index >= 2 and toks[index - 1].raw_text == "." else None

    paren_index = index + 1
    suffix = ""
    if _raw_text_at(toks, paren_index) == "$":
        suffix = toks[paren_index].raw_text
        paren_index += 1
    if _raw_text_at(toks, paren_index) != "(":
        return None

    specs = _runtime_argument_value_specs(name, host)
    canonical_name = specs[0].canonical_name if specs else _RELATION_FUNCTIONS.get(name.lower())
    if not canonical_name:
        return None
    if suffix and not (specs and specs[0].string_suffix):
        return None
    # `Error (70000)` opening a statement is the Error statement, judged by
    # _runtime_statement_value_hits; only the function reads a message.
    if canonical_name == "Error" and index == 0:
        return None
    lower = canonical_name.lower()
    # A project InStr does not take the call from VBA's (issue #280).
    if (
        not qualifier
        and lower != "instr"
        and (
            lower in module_signatures
            or lower in env
            or runtime_callable_source_shadowed(name, source_names)
        )
    ):
        return None

    close = match_paren_from(toks, paren_index)
    if close < 0:
        return None
    inner = list(toks[paren_index + 1 : close])
    split = empty_arg_split() if not inner else split_arg_slots(inner, span.start)
    return _RuntimeArgumentValueCall(f"{canonical_name}{suffix}", specs, split.slots)


# Functions _argument_relation_hit judges, which have no single-argument bound.
_RELATION_FUNCTIONS: dict[str, str] = {
    canonical.lower(): canonical
    for canonical in (
        "Partition", "Pmt", "IPmt", "PPmt", "SLN", "SYD", "DDB", "NPer", "Rate", "PV",
        "LBound", "UBound", "Join", "Filter", "Switch",
    )
}


def whole_number_arguments(
    name: str, slots: Sequence[list[VbaToken]], host: str | None
) -> list[list[VbaToken]]:
    """The arguments of a VBA library call that take a whole number with bounds,
    Mid's Start or Space's Number: each a Long or Integer parameter (XLIDE issue
    #298). Empty for a call this table does not know."""
    out: list[list[VbaToken]] = []
    for spec in _runtime_argument_value_specs(name, host):
        if spec.minimum is None:
            continue
        slot = _runtime_argument_value_slot(slots, spec)
        if slot is not None and len(slot) > 0:
            out.append(slot)
    return out


def _str_conv_conversion_any_locale(value: int | float) -> bool:
    """Whether some locale accepts a StrConv Conversion (XLIDE issue #184, measured
    in Excel 16.0 with the LCIDs of English, Japanese and Chinese). vbWide 4,
    vbNarrow 8, vbKatakana 16 and vbHiragana 32 run only under an East Asian
    locale, so they are never judged here. What no locale accepts: vbUnicode 64 or
    vbFromUnicode 128 with any other value, vbWide with vbNarrow, vbKatakana with
    vbHiragana, and any value past 255 or below 0."""
    if value < 0 or value > 255:
        return False
    whole = int(value)
    if (whole & 192) != 0 and whole != 64 and whole != 128:
        return False
    return (whole & 12) != 12 and (whole & 48) != 48


# The longest string VBA builds: Left("abc", 1073741823) runs, 1073741824 raises 5.
_MAX_STRING_LENGTH = 1073741823

# The interval strings DateAdd, DateDiff and DatePart accept.
_DATE_INTERVALS = ("yyyy", "q", "m", "y", "d", "w", "ww", "h", "n", "s")

_Spec = _RuntimeArgumentValueSpec


def _specs_table(database_compare: tuple[int, ...]) -> dict[str, tuple[_RuntimeArgumentValueSpec, ...]]:
    """The bounds each runtime function's arguments must keep to compile-and-run
    clean. The first nine were VBE-oracle-backed from the start; the rest were
    measured one call at a time in Excel 16.0 (build 20326, 2026-09-26, XLIDE
    issue #118): every listed value raises error 5 every time, and the nearest
    value that runs - Mid("abc", 10), Round(1.5, 0), Weekday(Date, 7),
    Environ(1) - stays quiet.

    Issue #218 added, each measured in Excel 16.0 (build 20326, 2026-09-30):

     - Compare, for all six functions that take one, is 0, 1, or a locale ID
       from 3 up that Windows knows (1033 and 66567 run, 16383 and 65536 raise).
       Only a negative value, and 2 for InStr, StrComp and Filter outside
       Access, are refused everywhere; no upper bound holds.
     - A string is at most 1073741823 characters: Left, Right and Mid's Length
       and String and Space's Number past it raise 5, and so does Mid's Start
       past 1073741824.
     - CVErr takes 0 to 65535; LeftB, RightB, MidB, AscB and InStrB keep Left's,
       Mid's, Asc's and InStr's lower bounds.
     - Overflow, error 6: TimeSerial and DateSerial take Integers, ChrB a Byte,
       String's Number and InStr's Start a Long, and Error at most 65535.

    `database_compare` is (2,) outside Access, where vbDatabaseCompare is
    refused by InStr, StrComp and Filter."""
    m = _MAX_STRING_LENGTH

    def format_specs(canonical: str) -> tuple[_RuntimeArgumentValueSpec, ...]:
        # A Tristate takes vbUseDefault, vbTrue or vbFalse: -2 to 0 (issue #476,
        # measured in Excel 16.0).
        return (
            _Spec(canonical, "NumDigitsAfterDecimal", 1, minimum=-1),
            *(
                _Spec(canonical, parameter, 2 + k, minimum=-2, maximum=0)
                for k, parameter in enumerate(("IncludeLeadingDigit", "UseParensForNegativeNumbers", "GroupDigits"))
            ),
        )

    return {
        "left": (_Spec("Left", "Length", 1, minimum=0, maximum=m, string_suffix=True),),
        "right": (_Spec("Right", "Length", 1, minimum=0, maximum=m, string_suffix=True),),
        "string": (
            _Spec("String", "Number", 0, minimum=0, maximum=m, overflow_type="Long", string_suffix=True),
            _Spec("String", "Character", 1, empty_string_raises=True, string_suffix=True),
        ),
        "space": (_Spec("Space", "Number", 0, minimum=0, maximum=m, string_suffix=True),),
        "mid": (
            _Spec("Mid", "Start", 1, minimum=1, maximum=m + 1, string_suffix=True),
            _Spec("Mid", "Length", 2, minimum=0, maximum=m, string_suffix=True),
        ),
        "leftb": (_Spec("LeftB", "Length", 1, minimum=0, string_suffix=True),),
        "rightb": (_Spec("RightB", "Length", 1, minimum=0, string_suffix=True),),
        "midb": (
            _Spec("MidB", "Start", 1, minimum=1, string_suffix=True),
            _Spec("MidB", "Length", 2, minimum=0, string_suffix=True),
        ),
        "ascb": (_Spec("AscB", "String", 0, empty_string_raises=True),),
        "instrb": (
            _Spec("InStrB", "Start", 0, minimum=1, overflow_type="Long", minimum_slot_count=3, allow_named=False),
        ),
        "chrb": (_Spec("ChrB", "CharCode", 0, overflow_type="Byte", string_suffix=True),),
        "cverr": (_Spec("CVErr", "ErrorNumber", 0, minimum=0, maximum=65535),),
        "error": (_Spec("Error", "ErrorNumber", 0, maximum=65535, bounds_overflow=True, string_suffix=True),),
        "timeserial": tuple(
            _Spec("TimeSerial", parameter, k, overflow_type="Integer")
            for k, parameter in enumerate(("Hour", "Minute", "Second"))
        ),
        "filter": (_Spec("Filter", "Compare", 3, minimum=0, disallowed=database_compare),),
        "replace": (
            _Spec("Replace", "Start", 3, minimum=1),
            _Spec("Replace", "Count", 4, minimum=-1),
            # Replace takes a Compare of 2 even outside Access; -1 raises (issue #189).
            _Spec("Replace", "Compare", 5, minimum=0),
        ),
        "instr": (
            _Spec(
                "InStr", "Start", 0, minimum=1, overflow_type="Long", minimum_slot_count=3, allow_named=False,
                skipped_by_empty_string=True,
            ),
            _Spec(
                "InStr", "Compare", 3, minimum=0, disallowed=database_compare, minimum_slot_count=4,
                allow_named=False, skipped_by_empty_string=True,
            ),
        ),
        "instrrev": (
            _Spec("InStrRev", "Start", 2, overflow_type="Long", minimum=-1, disallowed=(0,)),
            _Spec("InStrRev", "Compare", 3, overflow_type="Long", minimum=0),
        ),
        "chr": (_Spec("Chr", "CharCode", 0, overflow_type="Long", minimum=0, maximum=255, string_suffix=True),),
        "chrw": (_Spec("ChrW", "CharCode", 0, minimum=-32768, maximum=65535),),
        "asc": (_Spec("Asc", "String", 0, empty_string_raises=True),),
        "ascw": (_Spec("AscW", "String", 0, empty_string_raises=True),),
        "sqr": (_Spec("Sqr", "Number", 0, minimum=0, fractional=True),),
        "log": (_Spec("Log", "Number", 0, exclusive_minimum=0, fractional=True),),
        "monthname": (_Spec("MonthName", "Month", 0, overflow_type="Long", minimum=1, maximum=12),),
        "weekdayname": (
            _Spec("WeekdayName", "Weekday", 0, overflow_type="Long", minimum=1, maximum=7),
            _Spec("WeekdayName", "FirstDayOfWeek", 2, overflow_type="Long", minimum=0, maximum=7),
        ),
        "weekday": (_Spec("Weekday", "FirstDayOfWeek", 1, overflow_type="Long", minimum=0, maximum=7),),
        "round": (_Spec("Round", "NumDigitsAfterDecimal", 1, minimum=0, maximum=22),),
        # No bound on the Year alone: _date_serial_past_maximum judges the whole
        # date the month and day carry it to (issue #189).
        "dateserial": tuple(
            _Spec("DateSerial", parameter, k, overflow_type="Integer")
            for k, parameter in enumerate(("Year", "Month", "Day"))
        ),
        "strcomp": (
            _Spec("StrComp", "Compare", 2, overflow_type="Long", minimum=0, disallowed=database_compare),
        ),
        "split": (
            _Spec("Split", "Limit", 2, minimum=-1),
            _Spec("Split", "Compare", 3, minimum=0),
        ),
        "strconv": (
            _Spec("StrConv", "Conversion", 1, overflow_type="Long", accepts=_str_conv_conversion_any_locale),
        ),
        "formatnumber": format_specs("FormatNumber"),
        "formatcurrency": format_specs("FormatCurrency"),
        "formatpercent": format_specs("FormatPercent"),
        # A rate of -1 divides by zero inside: error 5 (issue #476, measured in
        # Excel 16.0). Pmt and FV take it.
        "npv": (_Spec("NPV", "Rate", 0, disallowed=(-1,)),),
        "irr": (_Spec("IRR", "Guess", 1, disallowed=(-1,)),),
        "mirr": (
            _Spec("MIRR", "FinanceRate", 1, disallowed=(-1,)),
            _Spec("MIRR", "ReinvestRate", 2, disallowed=(-1,)),
        ),
        # Environ takes a variable's number from 1 to 255, past which it raises 5
        # and past an Integer 6, or a name, of which "" raises 5 (issue #700,
        # measured in Excel 16.0).
        "environ": (
            _Spec(
                "Environ", "Expression", 0, minimum=1, maximum=255, overflow_type="Integer", string_suffix=True
            ),
            _Spec("Environ", "Expression", 0, empty_string_raises=True, string_suffix=True),
        ),
        # Shell of "" and Dir with an attribute of 64 or more raise 5; GetSetting of
        # an empty AppName or Section too (issue #700, measured in Excel 16.0).
        "shell": (_Spec("Shell", "PathName", 0, empty_string_raises=True),),
        "dir": (_Spec("Dir", "Attributes", 1, minimum=0, maximum=63, string_suffix=True),),
        "getsetting": (
            _Spec("GetSetting", "AppName", 0, empty_string_raises=True),
            _Spec("GetSetting", "Section", 1, empty_string_raises=True),
        ),
        "dateadd": (_Spec("DateAdd", "Interval", 0, allowed_strings=_DATE_INTERVALS),),
        # A first day of the week runs from 0 to 7 and a first week of the year
        # from 0 to 3; DateDiff reads only the first (issue #262, measured in
        # Excel 16.0: DateDiff with a FirstWeekOfYear of 4 runs).
        "datediff": (
            _Spec("DateDiff", "Interval", 0, allowed_strings=_DATE_INTERVALS),
            _Spec("DateDiff", "FirstDayOfWeek", 3, minimum=0, maximum=7),
        ),
        "datepart": (
            _Spec("DatePart", "Interval", 0, allowed_strings=_DATE_INTERVALS),
            _Spec("DatePart", "FirstDayOfWeek", 2, minimum=0, maximum=7),
            _Spec("DatePart", "FirstWeekOfYear", 3, minimum=0, maximum=3),
        ),
        "format": (
            _Spec("Format", "FirstDayOfWeek", 2, minimum=0, maximum=7, string_suffix=True),
            _Spec("Format", "FirstWeekOfYear", 3, minimum=0, maximum=3, string_suffix=True),
        ),
        "formatdatetime": (_Spec("FormatDateTime", "NamedFormat", 1, minimum=0, maximum=4),),
    }


# 2 is vbDatabaseCompare, which only Access accepts in InStr, StrComp and Filter.
_SPECS = _specs_table((2,))
_SPECS_ACCESS = _specs_table(())


def _runtime_argument_value_specs(name: str, host: str | None) -> tuple[_RuntimeArgumentValueSpec, ...]:
    return (_SPECS_ACCESS if host == "access" else _SPECS).get(name.lower(), ())


def _known_empty_string(slot: Sequence[VbaToken], known_strings: Mapping[str, str]) -> bool:
    """Whether an argument is "", vbNullString, or a String local known to hold ""."""
    toks = unwrap_outer_parens(_significant(slot))
    if len(toks) != 1:
        return False
    if toks[0].kind is TokenKind.STRING_LITERAL:
        return string_literal_value(toks[0].raw_text) == ""
    name = token_name(toks[0])
    lower = name.lower() if name is not None else None
    return lower == "vbnullstring" or (lower is not None and known_strings.get(lower) == "")


def _runtime_argument_value_slot(
    slots: Sequence[list[VbaToken]], spec: _RuntimeArgumentValueSpec
) -> list[VbaToken] | None:
    if spec.minimum_slot_count is not None and len(slots) < spec.minimum_slot_count:
        return None
    positional_index = 0
    for slot in slots:
        named = named_argument_slot(slot)
        if named is not None:
            if not spec.allow_named:
                continue
            if named[0].lower() == spec.parameter_name.lower():
                return named[1]
            continue
        if positional_index == spec.argument_index:
            return slot
        positional_index += 1
    return None


def _integer_argument_outside_bounds(
    source: str,
    slot: Sequence[VbaToken],
    slice_start: int,
    spec: _RuntimeArgumentValueSpec,
    constants: IntegerConstantLookup,
    string_calls: KnownStringCallContext,
) -> _OutsideBounds | None:
    known_strings: Mapping[str, str] = string_calls.known_strings
    toks = unwrap_outer_parens(_significant(slot))
    if not toks:
        return None
    # String-valued bounds: an empty literal where the function needs a character
    # (Asc(""), String(3, "")), or a literal outside the words the function
    # accepts (DateAdd("x", ...)). A String local the procedure never assigns is
    # "" too (Asc(s)).
    if spec.empty_string_raises or spec.allowed_strings is not None:
        if len(toks) != 1:
            return None
        name = token_name(toks[0])
        known_name = name.lower() if name is not None else None
        # Asc(Empty) is Asc("") (issue #332, measured in Excel 16.0).
        if known_name == "empty" and spec.canonical_name == "Asc":
            known: str | None = ""
        elif known_name is not None:
            known = known_strings.get(known_name)
        else:
            known = None
        is_literal = toks[0].kind is TokenKind.STRING_LITERAL
        if not is_literal and known is None:
            return None
        text = string_literal_value(toks[0].raw_text) if is_literal else known or ""
        if spec.empty_string_raises:
            raises = len(text) == 0
        else:
            raises = not any(allowed.lower() == text.lower() for allowed in spec.allowed_strings or ())
        if not raises:
            return None
        if is_literal:
            value: str = toks[0].raw_text
        elif known_name == "empty":
            value = 'Empty, which it reads as ""'
        else:
            value = f'"{text}" ({toks[0].raw_text} is never given another value)'
        return _OutsideBounds(value, Span(slice_start + toks[0].start, slice_start + toks[0].end))
    sign = 1
    literal = toks[0]
    start = literal.start
    literal_value: int | float | None = None
    signed_literal = len(toks) == 2 and toks[0].raw_text in ("-", "+")
    if signed_literal:
        sign = -1 if toks[0].raw_text == "-" else 1
        literal = toks[1]
        start = toks[0].start
    if literal.kind is TokenKind.INTEGER_LITERAL and (len(toks) == 1 or signed_literal):
        raw_value = parse_vba_integer_literal(literal.raw_text)
        if raw_value is not None:
            literal_value = sign * raw_value
    elif literal.kind is TokenKind.FLOAT_LITERAL and (len(toks) == 1 or signed_literal):
        # Sqr(-4.5) and Log(0.0) raise like their whole-number neighbours.
        float_value = _float_literal_value(literal.raw_text)
        if float_value is not None:
            literal_value = sign * float_value
    # True passes -1, False and Empty 0: Left("abc", True) and Mid("abc", False)
    # raise 5 (issue #434, measured in Excel 16.0).
    word = token_text(toks[0]) if len(toks) == 1 else ""
    if literal_value is None and word in ("true", "false", "empty"):
        literal_value = -1 if word == "true" else 0
    # A Date passes its serial: Chr(#1/2/2000#) is Chr(36527), which raises 5
    # (issue #332, measured in Excel 16.0).
    if literal_value is None and len(toks) == 1 and toks[0].kind is TokenKind.DATE_LITERAL:
        literal_value = date_literal_serial(toks[0].raw_text)
    if literal_value is not None:
        verdict = _argument_value_verdict(literal_value, spec)
        if verdict is None:
            return None
        return _OutsideBounds(
            _shown_argument_value(literal_value, spec),
            Span(slice_start + start, slice_start + literal.end),
            6 if verdict == 6 else None,
        )

    # `InStr(s, " ") - 1` with s known is `0 - 1` (issue #201).
    folded = fold_known_string_calls(toks, string_calls)
    expression_value = evaluate_integer_constant_expression(
        folded if folded is not None else source[slice_start + toks[0].start : slice_start + toks[-1].end],
        constants,
    )
    # A local past the Long range overflows as it converts: Chr(a) with a Double
    # of 1E+300 (issue #336). argument-type-mismatch sees only its type.
    local = len(toks) == 1 and token_name(toks[0]) is not None
    long_min, long_max = _OVERFLOW_RANGES["Long"]
    past_long = (
        local
        and not spec.fractional
        and expression_value is not None
        and (expression_value < long_min or expression_value > long_max)
    )
    if expression_value is None:
        return None
    expression_verdict = 6 if past_long else _argument_value_verdict(expression_value, spec)
    if expression_verdict is None:
        return None
    return _OutsideBounds(
        _shown_argument_value(expression_value, spec),
        Span(slice_start + toks[0].start, slice_start + toks[-1].end),
        6 if expression_verdict == 6 else None,
    )


def _argument_value_verdict(raw_value: int | float, spec: _RuntimeArgumentValueSpec) -> int | None:
    """What passing `raw_value` does: runs (None), raises 5, or overflows (6). A
    whole-number parameter whose type no spec states is judged only inside the
    Long range, since past it the conversion overflows first and the typed
    signature's argument-type-mismatch says so."""
    if not spec.fractional:
        value = _passed_argument_value(raw_value, spec)
        low, high = _OVERFLOW_RANGES[spec.overflow_type or "Long"]
        if value < low or value > high:
            return 6 if spec.overflow_type else None
    if _integer_argument_value_in_bounds(raw_value, spec):
        return None
    return 6 if spec.bounds_overflow else 5


def _passed_argument_value(value: int | float, spec: _RuntimeArgumentValueSpec) -> int | float:
    """The value VBA passes: a whole-number parameter takes the argument rounded
    half to even, so Space(-0.5) is Space(0) and runs, and Mid(s, 0.5) is Mid(s,
    0) and raises (XLIDE issue #189, measured in Excel 16.0)."""
    return value if spec.fractional else bankers_round(value)


def _shown_argument_value(value: int | float, spec: _RuntimeArgumentValueSpec) -> int | float | str:
    """The argument as the message states it, with the rounded value VBA uses."""
    passed = _passed_argument_value(value, spec)
    # 1E+300 as VBA prints it, not 1e+300.
    shown: int | float | str = _js_to_exponential(value).replace("e", "E") if abs(value) >= 1e15 else value
    if passed == value:
        return shown
    return f"{_display_value(shown)}, which VBA rounds to {_js_string(passed)}"


def _integer_argument_value_in_bounds(raw_value: int | float, spec: _RuntimeArgumentValueSpec) -> bool:
    value = _passed_argument_value(raw_value, spec)
    if spec.minimum is not None and value < spec.minimum:
        return False
    if spec.exclusive_minimum is not None and value <= spec.exclusive_minimum:
        return False
    if spec.maximum is not None and value > spec.maximum:
        return False
    if spec.disallowed is not None and value in spec.disallowed:
        return False
    return not (spec.accepts is not None and _is_integer_value(value) and not spec.accepts(value))


_DECLARATION_HEADS = frozenset(
    {
        "dim", "static", "const", "private", "public", "friend", "declare",
        "sub", "function", "property", "type", "enum",
    }
)


def _is_declaration_like_statement(toks: Sequence[VbaToken]) -> bool:
    return token_text(_token_at(toks, 0)) in _DECLARATION_HEADS


# -- checkRuntimeConversionValues ------------------------------------------


@dataclass(frozen=True, slots=True)
class _RuntimeConversionValueHit:
    display_name: str
    name: str
    # What the literal cannot become: 'a number', 'Boolean', 'Date'.
    target: str
    span: Span


def check_runtime_conversion_values(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
    activity: ConditionalActivityTracker | None = None,
) -> ProcedureStatementVisitor:
    """Selected conversion functions compile with Variant-like arguments but can
    deterministically fail at runtime for literal values that cannot be converted.
    This first slice is intentionally narrow for CDate string literals that are
    plainly non-date text."""

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        # `s = "abc"` then `CLng(s)`: the string a local holds here (XLIDE issue #238).
        values_at = known_local_literal_values_at(source, member, symbols, activity)

        def visitor(stmt: LeafStatementNode) -> None:
            strings = picked_values(values_at(stmt), _plain_string)
            numbers = picked_values(values_at(stmt), _plain_number)
            for hit in _runtime_conversion_value_hits(source, stmt.span, source_names, strings, numbers):
                push(
                    "runtimeConversionValue",
                    f"{hit.display_name} cannot convert {hit.name} to {hit.target}. This will raise "
                    "Run-time error '13': Type mismatch.",
                    hit.span,
                )

        return visitor

    return factory


# The conversion functions and the kind each converts a string literal to (XLIDE
# issues #118 and #188, each measured in Excel 16.0). What converts is
# string_conversion.py's: `CLng("4x2")`, `CDbl("5%")`, `CBool(" True ")` and
# `CDate("March")` raise 13, and `CLng("&HFF")` and `CDbl("(5)")` run.
_CONVERSION_TARGETS: dict[str, str] = {
    "cbyte": "numeric", "cint": "numeric", "clng": "numeric", "clnglng": "numeric", "clngptr": "numeric",
    "csng": "numeric", "cdbl": "numeric", "ccur": "numeric", "cdec": "numeric", "sgn": "numeric",
    "cbool": "boolean",
    "cdate": "date", "cvdate": "date", "datevalue": "date", "timevalue": "date",
    # The math functions read a number the same way (issue #242): Abs("abc") and
    # Hex("abc") raise 13, and Oct("8") runs.
    "abs": "numeric", "sqr": "numeric", "int": "numeric", "fix": "numeric", "round": "numeric",
    "hex": "numeric", "oct": "numeric",
    # Str("abc"), Str("") and Str("True") raise 13, and Str("&H10") runs (issue
    # #332, measured in Excel 16.0).
    "str": "numeric",
    "exp": "numeric", "log": "numeric", "sin": "numeric", "cos": "numeric", "tan": "numeric", "atn": "numeric",
    "year": "date", "month": "date", "day": "date", "weekday": "date", "hour": "date", "minute": "date",
    "second": "date",
    "dateadd": "date", "datepart": "date", "datediff": "date",
    # FormatNumber("abc") and FormatDateTime("abc") raise 13 (issue #476,
    # measured in Excel 16.0).
    "formatnumber": "numeric", "formatcurrency": "numeric", "formatpercent": "numeric", "formatdatetime": "date",
}

# The arguments each function converts, where it is not the first (issue #218):
# DateAdd's Date is its third, DatePart's its second, and DateDiff converts its
# second and third.
_CONVERTED_SLOTS: dict[str, tuple[int, ...]] = {"dateadd": (2,), "datepart": (1,), "datediff": (1, 2)}

# Number arguments a String local converts into (issue #332, measured in Excel
# 16.0): `Chr(s)`, `Space(s)`, `String(s, "a")`, `Left("abc", s)` and the count
# of `DateAdd` raise 13 when s holds "x" or "".
_NUMBER_SLOTS: dict[str, tuple[int, ...]] = {
    "chr": (0,), "chrw": (0,), "space": (0,), "string": (0,), "left": (1,), "right": (1,), "dateadd": (1,),
}

# The functions that read a number as a Date serial, which runs from -657434
# (January 1, 100) to 2958465 (December 31, 9999): Year(2958466) and
# Day(-657435) raise 13, Year(2958465.9) and Year(-657434.9) run (issue #218,
# measured in Excel 16.0). CDate raises 6 there, which arithmetic-overflow
# reports.
_DATE_SERIAL_READERS = frozenset(
    {"year", "month", "day", "weekday", "hour", "minute", "second", "dateadd", "datepart", "datediff"}
)

# Upstream indexes plain object literals, which also answer the two lowercase
# names Object.prototype carries. A call by either name finds a "target" there,
# and then `.map` of what CONVERTED_SLOTS and NUMBER_SLOTS give back (a
# function, an object) throws.
_OBJECT_PROTOTYPE_KEYS = frozenset({"constructor", "__proto__"})

_WHOLE_NUMBER_TEXT = re.compile(r"[+-]?[0-9]+")
_EDGE_BLANKS = re.compile(r"^[ \t]+|[ \t]+$")


def _runtime_conversion_value_hits(
    source: str,
    span: Span,
    source_names: SourceNameScope,
    known_strings: Mapping[str, str] | None = None,
    known_numbers: Mapping[str, int | float] | None = None,
) -> list[_RuntimeConversionValueHit]:
    strings: Mapping[str, str] = known_strings if known_strings is not None else {}
    numbers: Mapping[str, int | float] = known_numbers if known_numbers is not None else {}
    toks = statement_tokens(source, span)
    if _is_declaration_like_statement(toks):
        return []
    hits: list[_RuntimeConversionValueHit] = []
    for i in range(len(toks) - 2):
        name = token_name(toks[i])
        if not name:
            continue
        lower = name.lower()
        number_slots = _NUMBER_SLOTS.get(lower)
        prototype = lower in _OBJECT_PROTOTYPE_KEYS
        target = _CONVERSION_TARGETS.get(lower) or ("numeric" if number_slots is not None or prototype else None)
        if not target:
            continue
        if toks[i + 1].raw_text != "(" or not is_bare_or_vba_qualified_intrinsic_call(toks, i):
            continue
        qualified = i >= 1 and toks[i - 1].raw_text == "."
        if not qualified and runtime_callable_source_shadowed(name, source_names):
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0:
            continue
        split = split_arg_slots(toks[i + 2 : close], span.start)
        if any(named_argument_slot(slot) is not None for slot in split.slots):
            continue
        if prototype:
            raise TypeError(f"{lower}: converted slots are not an array")
        converted = _CONVERTED_SLOTS.get(lower, (0,)) if lower in _CONVERSION_TARGETS else ()
        slots = [(index, target, False) for index in converted] + [
            (index, "numeric", True) for index in (number_slots or ())
        ]
        for index, slot_target, held_only in slots:
            slot = _significant(split.slots[index] if index < len(split.slots) else [])
            if index < len(split.spans):
                at: Span | None = split.spans[index]
            elif slot:
                at = Span(span.start + slot[0].start, span.start + slot[-1].end)
            else:
                at = None
            display_name = f"VBA.{name}" if qualified else name
            held = (
                strings.get(slot[0].raw_text.lower())
                if len(slot) == 1 and slot[0].kind is TokenKind.IDENTIFIER
                else None
            )
            # A number slot of Chr, Space or Left is judged on a String local only:
            # a literal there is argument-type-mismatch's (issue #332).
            if (
                len(slot) == 1
                and ((slot[0].kind is TokenKind.STRING_LITERAL and not held_only) or held is not None)
                and at is not None
            ):
                value = held if held is not None else string_literal_value(slot[0].raw_text)
                if slot_target == "date":
                    # A time out of range for every reader of a date, and DateValue
                    # or TimeValue of a whole number (issues #262, #444, measured
                    # in Excel 16.0): Year("25:00"), DateValue("12"), TimeValue("-1").
                    invalid = (
                        is_invalid_date_string(value)
                        or is_invalid_time_string(value)
                        or (
                            lower in ("timevalue", "datevalue")
                            and _WHOLE_NUMBER_TEXT.fullmatch(_EDGE_BLANKS.sub("", value)) is not None
                        )
                    )
                elif slot_target == "boolean":
                    invalid = is_invalid_boolean_string(value)
                else:
                    invalid = is_invalid_numeric_string(value)
                if invalid:
                    hits.append(
                        _RuntimeConversionValueHit(
                            display_name,
                            slot[0].raw_text
                            if held is None
                            else f"{slot[0].raw_text}, which holds {_json_string(held)} here,",
                            "Date" if slot_target == "date" else "Boolean" if slot_target == "boolean" else "a number",
                            at,
                        )
                    )
                continue
            # A number local known here reads the same way (issue #332).
            held_number = (
                numbers.get(slot[0].raw_text.lower())
                if len(slot) == 1 and slot[0].kind is TokenKind.IDENTIFIER
                else None
            )
            literal_serial = _signed_numeric_literal(slot)
            serial = (
                (literal_serial if literal_serial is not None else held_number)
                if lower in _DATE_SERIAL_READERS and not held_only
                else None
            )
            if serial is not None and at is not None and (serial >= 2958466 or serial <= -657435):
                hits.append(
                    _RuntimeConversionValueHit(
                        display_name,
                        f"{slot[0].raw_text}, which holds {_js_string(serial)} here,"
                        if held_number is not None and literal_serial is None
                        else _js_string(serial),
                        "a Date, whose serial numbers run from -657434 (January 1, 100) to 2958465 "
                        "(December 31, 9999)",
                        at,
                    )
                )
    return hits


def _signed_numeric_literal(slot: Sequence[VbaToken]) -> int | float | None:
    """The value of a lone numeric literal, optionally signed."""
    toks = unwrap_outer_parens(_significant(slot))
    signed = len(toks) == 2 and toks[0].raw_text in ("-", "+")
    if len(toks) != 1 and not signed:
        return None
    literal = toks[-1]
    if literal.kind is not TokenKind.INTEGER_LITERAL and literal.kind is not TokenKind.FLOAT_LITERAL:
        return None
    return _numeric_literal_group_value(toks)


# -- JavaScript value semantics ---------------------------------------------


def _js_string(value: int | float) -> str:
    """String(value) for a JavaScript number."""
    return js_number_to_string(value)


def _display_value(value: int | float | str) -> str:
    """A hit's value as upstream's template literal prints it."""
    return value if isinstance(value, str) else _js_string(value)


def _js_to_exponential(value: int | float) -> str:
    """Number.prototype.toExponential() with no digit count: the shortest digits
    that read back as the same double, as `d.ddde+n`."""
    number = float(value)
    if number == 0:
        return "0e+0"
    _, digit_tuple, exponent = Decimal(repr(abs(number))).normalize().as_tuple()
    assert isinstance(exponent, int)
    digits = "".join(str(digit) for digit in digit_tuple)
    power = exponent + len(digits) - 1
    mantissa = digits if len(digits) == 1 else f"{digits[0]}.{digits[1:]}"
    return f"{'-' if number < 0 else ''}{mantissa}e{'+' if power >= 0 else '-'}{abs(power)}"


def _js_divide(a: float, b: float) -> float:
    """JavaScript's `/`: division by zero gives an infinity or NaN rather than raising."""
    if b == 0:
        if a == 0 or math.isnan(a):
            return math.nan
        return math.copysign(math.inf, a) * math.copysign(1.0, b)
    return a / b


def _json_string(text: str) -> str:
    """JSON.stringify of a string."""
    return json.dumps(text, ensure_ascii=False)


def _is_integer_value(value: int | float | str) -> bool:
    """Number.isInteger."""
    if isinstance(value, str):
        return False
    return isinstance(value, int) or (math.isfinite(value) and value.is_integer())


def _strip_vba_prefix(name: str) -> str:
    """`name.replace(/^VBA\\./i, '')`."""
    return name[4:] if name[:4].lower() == "vba." else name


_ASCII_DIGIT = re.compile(r"[0-9]")


def _plain_string(_lower: str, value: KnownLocalValue) -> str | None:
    """A String local's text, where no Mid statement has rewritten it."""
    return value.value if value.kind == "string" and not value.content_mutated and isinstance(value.value, str) else None


def _plain_number(_lower: str, value: KnownLocalValue) -> int | float | None:
    """A number local's value, where nothing has rewritten it."""
    if value.kind != "number" or value.content_mutated or isinstance(value.value, str):
        return None
    return value.value


def _string_length(_lower: str, value: KnownLocalValue) -> int | None:
    """A String local's length in UTF-16 code units, rewritten by Mid or not."""
    return utf16_length(str(value.value)) if value.kind == "string" else None


def _utf16_units(text: str) -> list[str]:
    """The string as JavaScript indexes it: one entry per UTF-16 code unit."""
    units: list[str] = []
    for ch in text:
        code = ord(ch)
        if code > 0xFFFF:
            code -= 0x10000
            units.append(chr(0xD800 | (code >> 10)))
            units.append(chr(0xDC00 | (code & 0x3FF)))
        else:
            units.append(ch)
    return units


def _from_units(units: Sequence[str]) -> str:
    """UTF-16 code units back to a string, a surrogate pair rejoined."""
    return "".join(units).encode("utf-16-le", "surrogatepass").decode("utf-16-le", "surrogatepass")


def _units_lower(units: Sequence[str]) -> list[str]:
    """toLowerCase over a string held as code units."""
    return _utf16_units(_from_units(units).lower())


def _units_upper(units: Sequence[str]) -> list[str]:
    """toUpperCase over a string held as code units."""
    return _utf16_units(_from_units(units).upper())


def _unit_lower(unit: str) -> str:
    """toLowerCase of one code unit, which may come back longer."""
    return "".join(_units_lower([unit]))


def _unit_upper(unit: str) -> str:
    """toUpperCase of one code unit, which may come back longer ('SS')."""
    return "".join(_units_upper([unit]))


def _index_of(units: Sequence[str], unit: str, start: int) -> int:
    for k in range(start, len(units)):
        if units[k] == unit:
            return k
    return -1


def _raw_text_at(toks: Sequence[VbaToken], index: int) -> str | None:
    return toks[index].raw_text if 0 <= index < len(toks) else None


def _token_at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    return toks[index] if 0 <= index < len(toks) else None


def _token_span(base: Span, tok: VbaToken) -> Span:
    return Span(base.start + tok.start, base.start + tok.end)


# -- JavaScript Date arithmetic, in UTC milliseconds --------------------------


def _days_from_civil(year: int, month: int, day: int) -> int:
    """Days since 1970-01-01 of a proleptic Gregorian date, for any year."""
    y = year - 1 if month <= 2 else year
    era = y // 400
    year_of_era = y - era * 400
    day_of_year = (153 * ((month + 9) % 12) + 2) // 5 + day - 1
    day_of_era = year_of_era * 365 + year_of_era // 4 - year_of_era // 100 + day_of_year
    return era * 146097 + day_of_era - 719468


def _civil_from_days(days: int) -> tuple[int, int, int]:
    """The proleptic Gregorian (year, month, day) of a day count since 1970-01-01."""
    z = days + 719468
    era = z // 146097
    day_of_era = z - era * 146097
    year_of_era = (day_of_era - day_of_era // 1460 + day_of_era // 36524 - day_of_era // 146096) // 365
    day_of_year = day_of_era - (365 * year_of_era + year_of_era // 4 - year_of_era // 100)
    shifted_month = (5 * day_of_year + 2) // 153
    day = day_of_year - (153 * shifted_month + 2) // 5 + 1
    month = shifted_month + 3 if shifted_month < 10 else shifted_month - 9
    return (year_of_era + era * 400 + (1 if month <= 2 else 0), month, day)


def _utc_ms(year: int, month: int, day: int) -> int:
    """Date.UTC(year, month - 1, day) for a year past 99."""
    return _days_from_civil(year, month, day) * _MS_PER_DAY


def _utc_date_ms(year: int, month_index: int, day: int) -> int:
    """`new Date(0)`, then setUTCFullYear(year, month_index, 1) and
    setUTCDate(day): the month index and the day roll over into the next unit."""
    carried_year = year + month_index // 12
    return (_days_from_civil(carried_year, month_index % 12 + 1, 1) + day - 1) * _MS_PER_DAY


def _epoch_ms(date: UtcDate | None) -> int | None:
    """A date's getTime(), or None where there is no date. An invalid date (NaN)
    compares false against every bound upstream, so it decides nothing here
    either."""
    if date is None or math.isnan(date.time):
        return None
    return int(date.time)
