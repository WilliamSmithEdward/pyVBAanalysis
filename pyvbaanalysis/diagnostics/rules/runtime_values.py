"""Rule family: deterministic runtime argument / conversion values.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/runtimeValues.ts. Some
runtime-library arguments have deterministic value bounds even when the argument
type is valid (e.g. Left(s, -1) raises Run-time error 5), a few statement and
operator forms raise for a value the code states (`Err.Raise 0`, `0 ^ -1`, an
invalid Like pattern), and selected conversions fail for provably invalid literals
(CDate("abc") raises error 13).
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import (
    IntegerConstantLookup,
    evaluate_integer_constant_expression,
    parse_vba_integer_literal,
    resolve_raw_integer_constants,
)
from ...host.host_model import HostObjectModel
from ...js_compat import JS_WHITESPACE, js_number_to_string, js_trim, utf16_length
from ...lexer.token_helpers import match_paren_from
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureSignature, VbaSymbol
from ...types.type_inference import type_environment_for
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
from ..known_locals import ConstantOrKnownLocalLookup, known_local_literal_values
from ..walker import ProcedureStatementVisitor, token_name, token_text
from .shared import is_bare_or_vba_qualified_intrinsic_call


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
    # An empty string literal raises: Asc(""), String(3, "").
    empty_string_raises: bool = False
    # A string literal must be one of these (case-insensitive): DateAdd's interval.
    allowed_strings: tuple[str, ...] | None = None
    minimum_slot_count: int | None = None
    allow_named: bool = True
    # Whether the function also has a `$`-suffixed spelling.
    string_suffix: bool = False


@dataclass(frozen=True, slots=True)
class _RuntimeArgumentValueHit:
    display_name: str
    parameter_name: str
    value: int | float | str
    span: Span


@dataclass(frozen=True, slots=True)
class _RuntimeArgumentValueCall:
    display_name: str
    specs: tuple[_RuntimeArgumentValueSpec, ...]
    slots: list[list[VbaToken]]


_DECLARATION_HEADS = frozenset(
    {
        "dim", "static", "const", "private", "public", "friend", "declare",
        "sub", "function", "property", "type", "enum",
    }
)
_RAISES_5 = "This will raise Run-time error '5': Invalid procedure call or argument."
_FLOAT_SUFFIX = re.compile(r"[!#@]\Z")


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

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        constants = procedure_integer_constant_lookup(
            member, module_constants, symbols, project_visible_symbols, activity, host_model
        )
        # A local the procedure never assigns holds its default, and one whose
        # every assignment is one literal holds that (XLIDE issue #118): `Asc(s)`
        # with s never assigned is Asc(""), and `Mid(s, 5, 1) = "x"` after
        # s = "abc" starts past the end.
        known = known_local_literal_values(source, member, symbols, activity)
        known_strings: dict[str, str] = {}
        known_string_lengths: dict[str, int] = {}
        for lower, local in known.items():
            if local.kind == "string" and isinstance(local.value, str):
                known_string_lengths[lower] = utf16_length(local.value)
                if not local.content_mutated:
                    known_strings[lower] = local.value
        lookup = ConstantOrKnownLocalLookup(constants, known)

        def visitor(stmt: LeafStatementNode) -> None:
            for hit in _runtime_argument_value_hits(
                source, stmt.span, module_signatures, env, lookup, known_strings, source_names, host
            ):
                push(
                    "runtimeArgumentValue",
                    f"Argument '{hit.parameter_name}' of '{hit.display_name}' is "
                    f"{_display_value(hit.value)}; this will raise Run-time error '5': "
                    "Invalid procedure call or argument.",
                    hit.span,
                )
            for message, span in _runtime_statement_value_hits(
                source, stmt.span, lookup, known_string_lengths, source_names
            ):
                push("runtimeArgumentValue", message, span)

        return visitor

    return factory


def _runtime_statement_value_hits(
    source: str,
    span: Span,
    constants: IntegerConstantLookup,
    known_string_lengths: Mapping[str, int],
    source_names: SourceNameScope,
) -> list[tuple[str, Span]]:
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
    out: list[tuple[str, Span]] = []
    head = token_text(toks[0] if toks else None)
    # `Mid(s, 5, 1) = "x"` with s holding "abc": the statement form starts past
    # the end of the string, error 5 (XLIDE issue #118). Only the length matters,
    # which an earlier Mid statement cannot have changed.
    if head in ("mid", "mid$") and _raw_text_at(toks, 1) == "(":
        close = match_paren_from(toks, 1)
        if close > 0 and _raw_text_at(toks, close + 1) == "=":
            split = split_arg_slots(toks[2:close], span.start)
            target_slot = split.slots[0] if split.slots else None
            target_name = (
                token_name(target_slot[0]) if target_slot is not None and len(target_slot) == 1 else None
            )
            length = known_string_lengths.get(target_name.lower()) if target_name is not None else None
            start_slot = split.slots[1] if len(split.slots) > 1 else None
            # An empty array is truthy in JavaScript, so `Mid(s, , 1) = "x"` reaches
            # integerGroupValue upstream and throws there; _integer_group_value
            # raises the same way.
            start = (
                _integer_group_value(source, span, start_slot, constants)
                if start_slot is not None
                else None
            )
            if length is not None and start is not None and start > length and target_slot is not None:
                out.append(
                    (
                        f"Mid statement start {js_number_to_string(start)} is past the end of "
                        f"{target_slot[0].raw_text}, which is {length} character(s) long. {_RAISES_5}",
                        split.spans[1] if len(split.spans) > 1 else _token_span(span, toks[0]),
                    )
                )
    # `Err.Raise n` and `Error n`.
    number_index = -1
    form = ""
    if head == "err" and _raw_text_at(toks, 1) == "." and token_text(_token_at(toks, 2)) == "raise":
        number_index = 3
        form = "Err.Raise"
    elif head == "error" and len(toks) > 1 and not runtime_callable_source_shadowed("Error", source_names):
        number_index = 1
        form = "Error"
    if number_index > 0:
        group = _number_argument_group(toks, number_index)
        if group is not None:
            value = _integer_group_value(source, span, group, constants)
            if value is not None and (value < 1 or value > 65535):
                out.append(
                    (
                        f"{form} {js_number_to_string(value)} is not an error number: valid numbers "
                        f"are 1 to 65535. {_RAISES_5}",
                        Span(span.start + group[0].start, span.start + group[-1].end),
                    )
                )
    for i in range(1, len(toks) - 1):
        tok = toks[i]
        if tok.kind is TokenKind.OPERATOR and tok.raw_text == "^":
            base = _numeric_operand_before(toks, i)
            exponent = _numeric_operand_after(toks, i)
            if base is not None and exponent is not None:
                if base < 0 and not _is_integer(exponent):
                    out.append(
                        (
                            "A negative number raised to the fractional power "
                            f"{js_number_to_string(exponent)} has no real value. {_RAISES_5}",
                            _token_span(span, tok),
                        )
                    )
                elif base == 0 and exponent < 0:
                    out.append(
                        (
                            f"Zero raised to the negative power {js_number_to_string(exponent)} "
                            f"divides by zero. {_RAISES_5}",
                            _token_span(span, tok),
                        )
                    )
            continue
        pattern_tok = toks[i + 1]
        if token_text(tok) == "like" and pattern_tok.kind is TokenKind.STRING_LITERAL:
            problem = _invalid_like_pattern(string_literal_value(pattern_tok.raw_text))
            if problem:
                out.append(
                    (
                        f"The Like pattern {pattern_tok.raw_text} {problem}. This will raise "
                        "Run-time error '93': Invalid pattern string.",
                        _token_span(span, pattern_tok),
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
    if toks[end].raw_text == ")":
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


def _float_literal_value(raw: str) -> float | None:
    """Number() of a float literal without its type suffix, when finite. A D
    exponent (`1D3`) is NaN there, as float() refuses it here."""
    try:
        value = float(_FLOAT_SUFFIX.sub("", raw))
    except ValueError:
        return None
    return value if math.isfinite(value) else None


def _invalid_like_pattern(pattern: str) -> str | None:
    """Why a Like pattern raises error 93, or None when it is well formed.

    Upstream walks the pattern by UTF-16 code unit, so an astral character is its
    two surrogates here too, compared and reported the way upstream sees them."""
    units = _utf16_units(pattern)
    i = 0
    while i < len(units):
        if units[i] != "[":
            i += 1
            continue
        close = _index_of(units, "]", i + 1)
        if close < 0:
            return "opens a character list it never closes"
        body = units[i + 1 : close]
        if body[:1] == ["!"]:
            body = body[1:]
        for k in range(1, len(body) - 1):
            if body[k] == "-" and body[k - 1] > body[k + 1]:
                return f"has the reversed range {body[k - 1]}-{body[k + 1]}"
        i = close + 1
    return None


def _runtime_argument_value_hits(
    source: str,
    span: Span,
    module_signatures: Mapping[str, CallableTypeSignature],
    env: Mapping[str, str],
    constants: IntegerConstantLookup,
    known_strings: Mapping[str, str],
    source_names: SourceNameScope,
    host: str | None,
) -> list[_RuntimeArgumentValueHit]:
    toks = statement_tokens(source, span)
    if _is_declaration_like_statement(toks):
        return []
    hits: list[_RuntimeArgumentValueHit] = []
    for i in range(len(toks) - 1):
        call = _runtime_argument_value_call_at(toks, i, span, module_signatures, env, source_names, host)
        if call is None:
            continue
        for spec in call.specs:
            slot = _runtime_argument_value_slot(call.slots, spec)
            literal = (
                _integer_argument_outside_bounds(source, slot, span.start, spec, constants, known_strings)
                if slot is not None
                else None
            )
            if literal is None:
                continue
            value, hit_span = literal
            hits.append(_RuntimeArgumentValueHit(call.display_name, spec.parameter_name, value, hit_span))
        overflow = _date_add_past_maximum(source, span, call, constants)
        if overflow is not None:
            hits.append(overflow)
    return hits


def _date_add_past_maximum(
    source: str,
    span: Span,
    call: _RuntimeArgumentValueCall,
    constants: IntegerConstantLookup,
) -> _RuntimeArgumentValueHit | None:
    """`DateAdd("d", 1, #12/31/9999#)`: adding to a date literal past the last date
    VBA has (December 31, 9999) raises error 5 (XLIDE issue #118). Only a date
    literal with a positive whole-number count and a day, week, month or year
    interval is decided here."""
    if call.display_name != "DateAdd" or len(call.slots) < 3:
        return None
    interval_slot, number_slot, date_slot = (
        unwrap_outer_parens([t for t in slot if t.kind is not TokenKind.COMMENT])
        for slot in call.slots[:3]
    )
    if (
        len(interval_slot) != 1
        or interval_slot[0].kind is not TokenKind.STRING_LITERAL
        or len(date_slot) != 1
        or date_slot[0].kind is not TokenKind.DATE_LITERAL
    ):
        return None
    interval = string_literal_value(interval_slot[0].raw_text).lower()
    count = _integer_group_value(source, span, number_slot, constants)
    if count is None or count <= 0:
        return None
    date = _parse_date_literal(date_slot[0].raw_text)
    if date is None:
        return None
    result = _date_add(date, interval, count)
    if result is None or result <= _MAXIMUM_DATE:
        return None
    return _RuntimeArgumentValueHit(
        "DateAdd",
        "Date",
        f"{date_slot[0].raw_text}, which the {js_number_to_string(count)} {interval} interval(s) "
        "carry past December 31, 9999",
        Span(span.start + date_slot[0].start, span.start + date_slot[0].end),
    )


def _date_add(date: int, interval: str, count: int) -> int | None:
    """Upstream's JavaScript Date arithmetic on a UTC midnight, in days since the
    epoch: setUTCDate adds days, setUTCMonth and setUTCFullYear keep the day of the
    month and let it roll over (January 31 plus a month is March 2 or 3). A result
    past the range a Date holds is NaN there, which compares past the maximum as a
    large day count does here."""
    if interval in ("d", "y", "w"):
        return date + count
    if interval == "ww":
        return date + 7 * count
    year, month, day = _civil_from_days(date)
    if interval in ("m", "q"):
        month_index = month - 1 + (count if interval == "m" else 3 * count)
        return _days_from_civil(year + month_index // 12, month_index % 12 + 1, 1) + day - 1
    if interval == "yyyy":
        return _days_from_civil(year + count, month, 1) + day - 1
    return None


_WS = "[" + JS_WHITESPACE + "]"
_DATE_LITERAL = re.compile(
    "#" + _WS + "*([0-9]{1,2})/([0-9]{1,2})/([0-9]{4})" + _WS + "*"
    "(?:[0-9]{1,2}:[0-9]{2}(?::[0-9]{2})?" + _WS + "*(?:[Aa][Mm]|[Pp][Mm])?)?" + _WS + "*#"
)


def _parse_date_literal(raw: str) -> int | None:
    """A `#m/d/yyyy#` date literal as days since the epoch, or None for any other
    spelling. Date.UTC reads a year from 0 to 99 as 1900 plus the year, and rolls
    a day past the month's end into the next month."""
    match = _DATE_LITERAL.fullmatch(raw)
    if match is None:
        return None
    month = int(match.group(1))
    day = int(match.group(2))
    year = int(match.group(3))
    if month < 1 or month > 12 or day < 1 or day > 31:
        return None
    if 0 <= year <= 99:
        year += 1900
    return _days_from_civil(year, month, 1) + day - 1


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


_MAXIMUM_DATE = _days_from_civil(9999, 12, 31)


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
    qualifier = (
        token_name(toks[index - 2])
        if index >= 2 and toks[index - 1].raw_text == "."
        else None
    )

    paren_index = index + 1
    suffix = ""
    if _raw_text_at(toks, paren_index) == "$":
        suffix = toks[paren_index].raw_text
        paren_index += 1
    if _raw_text_at(toks, paren_index) != "(":
        return None

    specs = _runtime_argument_value_specs(name, host)
    if not specs:
        return None
    if suffix and not specs[0].string_suffix:
        return None
    lower = specs[0].canonical_name.lower()
    if not qualifier and (
        lower in module_signatures
        or lower in env
        or runtime_callable_source_shadowed(name, source_names)
    ):
        return None

    close = match_paren_from(toks, paren_index)
    if close < 0:
        return None
    inner = list(toks[paren_index + 1 : close])
    split = empty_arg_split() if not inner else split_arg_slots(inner, span.start)
    return _RuntimeArgumentValueCall(f"{specs[0].canonical_name}{suffix}", specs, split.slots)


# The interval strings DateAdd, DateDiff and DatePart accept.
_DATE_INTERVALS = ("yyyy", "q", "m", "y", "d", "w", "ww", "h", "n", "s")

# The bounds each runtime function's arguments must keep to compile-and-run clean.
# The first nine were VBE-oracle-backed from the start; the rest were measured one
# call at a time in Excel 16.0 (build 20326, 2026-09-26, XLIDE issue #118): every
# listed value raises error 5 every time, and the nearest value that runs -
# Mid("abc", 10), Round(1.5, 0), Weekday(Date, 7), Environ(1) - stays quiet.
# `vbDatabaseCompare` (2) is valid only where Access is the host, so InStr's
# Compare bound depends on the host.
_RUNTIME_ARGUMENT_VALUE_SPECS: dict[str, tuple[_RuntimeArgumentValueSpec, ...]] = {
    "left": (_RuntimeArgumentValueSpec("Left", "Length", 1, minimum=0, string_suffix=True),),
    "right": (_RuntimeArgumentValueSpec("Right", "Length", 1, minimum=0, string_suffix=True),),
    "string": (
        _RuntimeArgumentValueSpec("String", "Number", 0, minimum=0, string_suffix=True),
        _RuntimeArgumentValueSpec("String", "Character", 1, empty_string_raises=True, string_suffix=True),
    ),
    "space": (_RuntimeArgumentValueSpec("Space", "Number", 0, minimum=0, string_suffix=True),),
    "mid": (
        _RuntimeArgumentValueSpec("Mid", "Start", 1, minimum=1, string_suffix=True),
        _RuntimeArgumentValueSpec("Mid", "Length", 2, minimum=0, string_suffix=True),
    ),
    "replace": (
        _RuntimeArgumentValueSpec("Replace", "Start", 3, minimum=1),
        _RuntimeArgumentValueSpec("Replace", "Count", 4, minimum=-1),
    ),
    "instrrev": (_RuntimeArgumentValueSpec("InStrRev", "Start", 2, minimum=-1, disallowed=(0,)),),
    "chr": (_RuntimeArgumentValueSpec("Chr", "CharCode", 0, minimum=0, maximum=255, string_suffix=True),),
    "chrw": (_RuntimeArgumentValueSpec("ChrW", "CharCode", 0, maximum=65535),),
    "asc": (_RuntimeArgumentValueSpec("Asc", "String", 0, empty_string_raises=True),),
    "ascw": (_RuntimeArgumentValueSpec("AscW", "String", 0, empty_string_raises=True),),
    "sqr": (_RuntimeArgumentValueSpec("Sqr", "Number", 0, minimum=0),),
    "log": (_RuntimeArgumentValueSpec("Log", "Number", 0, exclusive_minimum=0),),
    "monthname": (_RuntimeArgumentValueSpec("MonthName", "Month", 0, minimum=1, maximum=12),),
    "weekdayname": (
        _RuntimeArgumentValueSpec("WeekdayName", "Weekday", 0, minimum=1, maximum=7),
        _RuntimeArgumentValueSpec("WeekdayName", "FirstDayOfWeek", 2, minimum=0, maximum=7),
    ),
    "weekday": (_RuntimeArgumentValueSpec("Weekday", "FirstDayOfWeek", 1, minimum=0, maximum=7),),
    "round": (_RuntimeArgumentValueSpec("Round", "NumDigitsAfterDecimal", 1, minimum=0),),
    "dateserial": (_RuntimeArgumentValueSpec("DateSerial", "Year", 0, maximum=9999),),
    "split": (_RuntimeArgumentValueSpec("Split", "Limit", 2, minimum=-1),),
    "formatnumber": (_RuntimeArgumentValueSpec("FormatNumber", "NumDigitsAfterDecimal", 1, minimum=-1),),
    "formatcurrency": (_RuntimeArgumentValueSpec("FormatCurrency", "NumDigitsAfterDecimal", 1, minimum=-1),),
    "formatpercent": (_RuntimeArgumentValueSpec("FormatPercent", "NumDigitsAfterDecimal", 1, minimum=-1),),
    "environ": (_RuntimeArgumentValueSpec("Environ", "Expression", 0, minimum=1, string_suffix=True),),
    "dateadd": (_RuntimeArgumentValueSpec("DateAdd", "Interval", 0, allowed_strings=_DATE_INTERVALS),),
    "datediff": (_RuntimeArgumentValueSpec("DateDiff", "Interval", 0, allowed_strings=_DATE_INTERVALS),),
    "datepart": (_RuntimeArgumentValueSpec("DatePart", "Interval", 0, allowed_strings=_DATE_INTERVALS),),
}
def _instr_specs(compare_maximum: int) -> tuple[_RuntimeArgumentValueSpec, ...]:
    return (
        _RuntimeArgumentValueSpec("InStr", "Start", 0, minimum=1, minimum_slot_count=3, allow_named=False),
        _RuntimeArgumentValueSpec(
            "InStr", "Compare", 3, minimum=0, maximum=compare_maximum, minimum_slot_count=4, allow_named=False
        ),
    )


_INSTR_SPECS = _instr_specs(1)
_INSTR_SPECS_ACCESS = _instr_specs(2)


def _runtime_argument_value_specs(name: str, host: str | None) -> tuple[_RuntimeArgumentValueSpec, ...]:
    lower = name.lower()
    if lower == "instr":
        return _INSTR_SPECS_ACCESS if host == "access" else _INSTR_SPECS
    return _RUNTIME_ARGUMENT_VALUE_SPECS.get(lower, ())


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
    known_strings: Mapping[str, str],
) -> tuple[int | float | str, Span] | None:
    toks = unwrap_outer_parens(
        [t for t in slot if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE]
    )
    if not toks:
        return None
    # String-valued bounds: an empty literal where the function needs a character
    # (Asc(""), String(3, "")), or a literal outside the words the function
    # accepts (DateAdd("x", ...)). A String local the procedure never assigns is
    # "" too (Asc(s)).
    if spec.empty_string_raises or spec.allowed_strings is not None:
        if len(toks) != 1:
            return None
        known_name = token_name(toks[0])
        known = known_strings.get(known_name.lower()) if known_name is not None else None
        is_literal = toks[0].kind is TokenKind.STRING_LITERAL
        if is_literal:
            text = string_literal_value(toks[0].raw_text)
        elif known is not None:
            text = known
        else:
            return None
        if spec.empty_string_raises:
            raises = len(text) == 0
        else:
            raises = not any(allowed.lower() == text.lower() for allowed in spec.allowed_strings or ())
        if not raises:
            return None
        value = (
            toks[0].raw_text
            if is_literal
            else f'"{text}" ({toks[0].raw_text} is never given another value)'
        )
        return (value, Span(slice_start + toks[0].start, slice_start + toks[0].end))
    sign = 1
    literal = toks[0]
    start = toks[0].start
    literal_value: int | float | None = None
    signed_literal = len(toks) == 2 and toks[0].raw_text in ("-", "+")
    if signed_literal:
        sign = -1 if toks[0].raw_text == "-" else 1
        literal = toks[1]
        start = toks[0].start
    if len(toks) == 1 or signed_literal:
        if literal.kind is TokenKind.INTEGER_LITERAL:
            raw_value = parse_vba_integer_literal(literal.raw_text)
            if raw_value is not None:
                literal_value = sign * raw_value
        elif literal.kind is TokenKind.FLOAT_LITERAL:
            # Sqr(-4.5) and Log(0.0) raise like their whole-number neighbours.
            float_value = _float_literal_value(literal.raw_text)
            if float_value is not None:
                literal_value = sign * float_value
    if literal_value is not None:
        if _integer_argument_value_in_bounds(literal_value, spec):
            return None
        return (literal_value, Span(slice_start + start, slice_start + literal.end))

    expression_value = evaluate_integer_constant_expression(
        source[slice_start + toks[0].start : slice_start + toks[-1].end], constants
    )
    if expression_value is None or _integer_argument_value_in_bounds(expression_value, spec):
        return None
    return (expression_value, Span(slice_start + toks[0].start, slice_start + toks[-1].end))


def _integer_argument_value_in_bounds(value: int | float, spec: _RuntimeArgumentValueSpec) -> bool:
    if spec.minimum is not None and value < spec.minimum:
        return False
    if spec.exclusive_minimum is not None and value <= spec.exclusive_minimum:
        return False
    if spec.maximum is not None and value > spec.maximum:
        return False
    return not (spec.disallowed is not None and value in spec.disallowed)


def _is_declaration_like_statement(toks: Sequence[VbaToken]) -> bool:
    return bool(toks) and token_text(toks[0]) in _DECLARATION_HEADS


# -- checkRuntimeConversionValues ------------------------------------------

# The conversion functions and what a string literal must look like to convert
# (XLIDE issue #118, each measured in Excel 16.0): the numeric conversions refuse
# letters-only and empty strings (`CLng("abc")`, `CDbl("")`) and take `"&H10"`;
# CBool takes True/False and numbers, and refuses `"yes"`; the date conversions
# refuse letters that name no month (`DateValue("abc")`).
_CONVERSION_TARGETS: dict[str, str] = {
    "cbyte": "numeric", "cint": "numeric", "clng": "numeric", "clnglng": "numeric", "clngptr": "numeric",
    "csng": "numeric", "cdbl": "numeric", "ccur": "numeric", "cdec": "numeric",
    "cbool": "boolean",
    "cdate": "date", "cvdate": "date", "datevalue": "date", "timevalue": "date",
}
# Upstream indexes a plain object literal, which also answers the two lowercase
# names Object.prototype carries; neither is 'date' nor 'boolean', so both read as
# a numeric conversion there (`[__proto__]("abc")` is reported).
_OBJECT_PROTOTYPE_KEYS = frozenset({"constructor", "__proto__"})

_MONTH_NAME = re.compile(
    r"\b(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?"
    r"|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b",
    re.IGNORECASE | re.ASCII,
)
_HAS_DIGIT = re.compile(r"[0-9]")
_NON_ASCII = re.compile(r"[^\x00-\x7F]")
_LETTERS_AND_SPACES = re.compile("[A-Za-z" + JS_WHITESPACE + "]+")


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
) -> ProcedureStatementVisitor:
    """Selected conversion functions compile with Variant-like arguments but can
    deterministically fail at runtime for literal values that cannot be converted.
    This first slice is intentionally narrow for CDate string literals that are
    plainly non-date text."""

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)

        def visitor(stmt: LeafStatementNode) -> None:
            for hit in _runtime_conversion_value_hits(source, stmt.span, source_names):
                push(
                    "runtimeConversionValue",
                    f"{hit.display_name} cannot convert {hit.name} to {hit.target}. This will raise "
                    "Run-time error '13': Type mismatch.",
                    hit.span,
                )

        return visitor

    return factory


def _runtime_conversion_value_hits(
    source: str, span: Span, source_names: SourceNameScope
) -> list[_RuntimeConversionValueHit]:
    toks = statement_tokens(source, span)
    if _is_declaration_like_statement(toks):
        return []
    hits: list[_RuntimeConversionValueHit] = []
    for i in range(len(toks) - 2):
        name = token_name(toks[i])
        if not name:
            continue
        lower = name.lower()
        target = _CONVERSION_TARGETS.get(lower) or ("prototype" if lower in _OBJECT_PROTOTYPE_KEYS else None)
        if target is None:
            continue
        if toks[i + 1].raw_text != "(" or not is_bare_or_vba_qualified_intrinsic_call(toks, i):
            continue
        qualified = i >= 1 and toks[i - 1].raw_text == "."
        if not qualified and runtime_callable_source_shadowed(name, source_names):
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0:
            continue
        split = split_arg_slots(list(toks[i + 2 : close]), span.start)
        first_slot = split.slots[0] if split.slots else []
        if len(first_slot) != 1 or first_slot[0].kind is not TokenKind.STRING_LITERAL:
            continue
        value = string_literal_value(first_slot[0].raw_text)
        if target == "date":
            invalid = _is_definitely_invalid_date_string(value)
        elif target == "boolean":
            invalid = _is_definitely_invalid_boolean_string(value)
        else:
            invalid = _is_definitely_non_numeric_string(value)
        if not invalid:
            continue
        hits.append(
            _RuntimeConversionValueHit(
                f"VBA.{name}" if qualified else name,
                first_slot[0].raw_text,
                "Date" if target == "date" else "Boolean" if target == "boolean" else "a number",
                split.spans[0]
                if split.spans
                else Span(span.start + first_slot[0].start, span.start + first_slot[0].end),
            )
        )
    return hits


def _is_definitely_non_numeric_string(value: str) -> bool:
    """Empty, or letters and spaces only: nothing VBA's numeric parser reads as a number."""
    trimmed = js_trim(value)
    return len(trimmed) == 0 or _LETTERS_AND_SPACES.fullmatch(trimmed) is not None


def _is_definitely_invalid_boolean_string(value: str) -> bool:
    """CBool takes True, False and anything numeric; letters that are neither raise 13."""
    trimmed = js_trim(value)
    if len(trimmed) == 0:
        return True
    return _LETTERS_AND_SPACES.fullmatch(trimmed) is not None and trimmed.lower() not in ("true", "false")


def _is_definitely_invalid_date_string(value: str) -> bool:
    trimmed = js_trim(value)
    if not trimmed:
        return True
    if _HAS_DIGIT.search(trimmed) or _NON_ASCII.search(trimmed):
        return False
    if _LETTERS_AND_SPACES.fullmatch(trimmed) is None:
        return False
    return _MONTH_NAME.search(trimmed) is None


# -- JavaScript value semantics ---------------------------------------------


def _display_value(value: int | float | str) -> str:
    """A hit's value as upstream's template literal prints it."""
    return value if isinstance(value, str) else js_number_to_string(value)


def _is_integer(value: int | float) -> bool:
    """Number.isInteger."""
    return isinstance(value, int) or value.is_integer()


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
