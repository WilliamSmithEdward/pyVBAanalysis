"""Call-site extraction and argument-list utilities.

Ported from xlide_vscode/src/analyzer/diagnostics/callExtraction.ts. Owns the
value model for a call (CallArguments, the callable-signature shapes, inferred
argument types), the argument-slot splitter that the call/argument rules share,
the statement-level call extractors (extract_call / extract_qualified_call), and
arity validation.
"""

from __future__ import annotations

import math
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field

from ..call.call_context import bare_call_statement_target
from ..js_compat import js_trim
from ..lexer.token_helpers import match_paren_from
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import Span
from ..symbols.symbol_model import qualified_procedure_key
from .context import PushFn, statement_tokens
from .walker import first_executable_token_index, strip_header_brackets, token_name, token_text

_STRING_LITERAL_QUOTES = (re.compile(r'^"'), re.compile(r'"$'))


@dataclass(slots=True)
class CallableParamType:
    name: str
    type_: str | None = None
    optional: bool = False
    param_array: bool = False
    is_array: bool | None = None
    by_ref: bool | None = None
    # A Variant parameter the function refuses Null for (error 94).
    null_raises: bool | None = None


@dataclass(slots=True)
class CallableTypeSignature:
    name: str
    params: list[CallableParamType] = field(default_factory=list)
    return_type: str | None = None
    # A Function or Property Get, which gives a value; None where unknown.
    valued: bool | None = None


_SCALAR_RESULT_TYPES = frozenset(
    (
        "string",
        "boolean",
        "date",
        "byte",
        "integer",
        "long",
        "longlong",
        "longptr",
        "single",
        "double",
        "currency",
        "decimal",
    )
)
_EMPTY_PARENS_END_RE = re.compile(r"\(\s*\)$")
_VBA_PREFIX_RE = re.compile(r"^vba\.")


def call_then_index(sig: CallableTypeSignature, call: CallArguments) -> bool:
    """`Arr(1)` with `Function Arr() As Variant`: a call that takes no argument,
    then an index into what it returns, when that is a Variant or an object. VBA
    compiles it and indexes the result, as a statement too, `Arr 1` and
    `Call Arr(1)` (XLIDE issue #609, measured in Excel 16.0). A scalar or a typed
    array result refuses the argument at compile."""
    if sig.valued is not True or len(sig.params) > 0 or len(call.slots) == 0:
        return False
    type_ = js_trim(sig.return_type).lower() if sig.return_type is not None else None
    if type_ is None or type_ == "variant":
        return True
    return not _EMPTY_PARENS_END_RE.search(type_) and _VBA_PREFIX_RE.sub("", type_) not in _SCALAR_RESULT_TYPES


@dataclass(slots=True)
class InferredArgumentType:
    type_: str
    label: str
    span: Span
    string_value: str | None = None
    numeric_value: float | None = None
    numeric_text: str | None = None
    # Set when numeric_value comes from a resolved named constant (e.g. xlChart)
    # rather than a literal, so overflow messages can name the constant instead
    # of calling its display name a "numeric literal".
    numeric_constant_name: str | None = None
    # A float literal's value (`3000000000#`, `2.5`). Kept apart from
    # numeric_value, whose range checks assume a whole number exactly held.
    float_value: float | None = None
    # The local whose known value numeric_value or float_value is, so the range
    # checks say whose value it is rather than call it a literal (XLIDE #332).
    held_by: str | None = None


@dataclass(slots=True)
class CallArguments:
    name: str
    name_span: Span
    slots: list[list[VbaToken]]
    slice_start: int
    qualifier: str | None = None
    lookup_key: str | None = None
    explicit_call: bool = False
    # True for `Take (i)` as a statement: the space makes the parentheses part of
    # the one argument, which is then passed as a copy, so ByRef exactness does
    # not apply to it (XLIDE issue #111).
    arguments_parenthesized: bool = False
    slot_spans: list[Span] | None = None


@dataclass(slots=True)
class ArgSplit:
    slots: list[list[VbaToken]]
    spans: list[Span]


def split_arg_slots(toks: Sequence[VbaToken], slice_start: int) -> ArgSplit:
    """Split top-level comma-separated argument slots, with a span per slot."""
    slots: list[list[VbaToken]] = [[]]
    spans: list[Span] = []
    depth = 0
    empty_marker: VbaToken | None = None

    def finish_slot(next_separator: VbaToken | None = None) -> None:
        spans.append(_argument_slot_span(slots[-1], empty_marker, next_separator, slice_start))

    for tok in toks:
        if tok.kind is TokenKind.PUNCTUATION and tok.raw_text == "(":
            depth += 1
        elif tok.kind is TokenKind.PUNCTUATION and tok.raw_text == ")":
            depth -= 1
        if tok.kind is TokenKind.PUNCTUATION and tok.raw_text == "," and depth == 0:
            finish_slot(tok)
            slots.append([])
            empty_marker = tok
        else:
            slots[-1].append(tok)
            empty_marker = None
    finish_slot()
    return ArgSplit(slots=slots, spans=spans)


def empty_arg_split() -> ArgSplit:
    return ArgSplit(slots=[], spans=[])


def _argument_slot_span(
    slot: Sequence[VbaToken],
    empty_marker: VbaToken | None,
    next_separator: VbaToken | None,
    slice_start: int,
) -> Span:
    if slot:
        return Span(slice_start + slot[0].start, slice_start + slot[-1].end)
    if empty_marker is not None:
        return Span(slice_start + empty_marker.start, slice_start + empty_marker.end)
    if next_separator is not None:
        return Span(slice_start + next_separator.start, slice_start + next_separator.end)
    return Span(slice_start, slice_start)


def is_named_slot(slot: Sequence[VbaToken]) -> bool:
    """True if a slot is a named argument (`name := value`).

    `:=` has one meaning in VBA, so the word before it is a parameter name whatever
    it spells. `Type` lexes as a keyword, and reading `Type:=xlLinkTypeExcelLinks`
    as positional reported correct calls (XLIDE 4.1.2)."""
    return (
        len(slot) >= 2
        and slot[0].kind in (TokenKind.IDENTIFIER, TokenKind.BRACKETED_IDENTIFIER, TokenKind.KEYWORD)
        and slot[1].kind is TokenKind.OPERATOR
        and slot[1].raw_text == ":="
    )


def callable_accepts_zero_arguments(sig: CallableTypeSignature) -> bool:
    return all(param.optional or param.param_array for param in sig.params)


def describe_arity(required: int, maximum: float) -> str:
    """Describe a procedure's acceptable argument-count range for a message."""
    if maximum == math.inf:
        return f"at least {required} argument{'' if required == 1 else 's'}"
    if required == maximum:
        return f"{required} argument{'' if required == 1 else 's'}"
    return f"between {required} and {int(maximum)} arguments"


def named_argument_slot(slot: Sequence[VbaToken]) -> tuple[str, list[VbaToken]] | None:
    """The (name, value-tokens) of a `name := value` slot, or None."""
    if not is_named_slot(slot):
        return None
    return (strip_header_brackets(slot[0].raw_text), list(slot[2:]))


def unwrap_outer_parens(toks: Sequence[VbaToken]) -> list[VbaToken]:
    """Strip one fully-enclosing pair of parentheses from a token list."""
    if len(toks) < 2 or toks[0].raw_text != "(":
        return list(toks)
    close = match_paren_from(toks, 0)
    return list(toks[1:-1]) if close == len(toks) - 1 else list(toks)


def string_literal_value(raw: str) -> str:
    """The text content of a string literal, with surrounding quotes removed."""
    return _STRING_LITERAL_QUOTES[1].sub("", _STRING_LITERAL_QUOTES[0].sub("", raw)).replace('""', '"')


def _significant_slice_tokens(source: str, span: Span) -> list[VbaToken]:
    # Reuse the per-pass statement-token cache (slice-relative offsets, byte-
    # identical to a raw tokenize of the slice) so a call statement is not
    # re-lexed once per call-validation rule.
    return statement_tokens(source, span)


def _explicit_call_inner_args(toks: list[VbaToken], open_search_start: int) -> list[VbaToken] | None:
    """Tokens strictly inside the `Call`-statement parentheses, or None if unbalanced."""
    depth = 0
    inner: list[VbaToken] = []
    for k in range(open_search_start, len(toks)):
        t = toks[k]
        if t.kind is TokenKind.PUNCTUATION and t.raw_text == "(":
            depth += 1
            if depth == 1:
                continue
        elif t.kind is TokenKind.PUNCTUATION and t.raw_text == ")":
            depth -= 1
            if depth == 0:
                return inner
        if depth >= 1:
            inner.append(t)
    return None


def extract_call(source: str, span: Span) -> CallArguments | None:
    """The callee and top-level argument slots of a bare/explicit call statement."""
    hit = bare_call_statement_target(source, span)
    if hit is None:
        return None
    slice_start = span.start
    toks = _significant_slice_tokens(source, span)
    start_index = first_executable_token_index(toks)
    rel_callee_start = hit.span.start - slice_start
    callee_idx = next((i for i, t in enumerate(toks) if t.start == rel_callee_start), -1)
    if callee_idx < 0:
        return None

    explicit_call = token_text(toks[start_index]) == "call"
    next_tok = toks[callee_idx + 1] if callee_idx + 1 < len(toks) else None
    if explicit_call:
        if next_tok is not None and next_tok.kind is TokenKind.PUNCTUATION and next_tok.raw_text == "(":
            inner = _explicit_call_inner_args(toks, callee_idx + 1)
            if inner is None:
                return None  # unbalanced - the parentheses rule reports this
            arg_toks = inner
        else:
            arg_toks = []  # `Call Foo` with no parameter list
    else:
        arg_toks = toks[callee_idx + 1 :]

    split = empty_arg_split() if not arg_toks else split_arg_slots(arg_toks, slice_start)
    return CallArguments(
        name=hit.name,
        name_span=hit.span,
        explicit_call=explicit_call,
        slots=split.slots,
        slot_spans=split.spans,
        slice_start=slice_start,
    )


def extract_qualified_call(
    source: str, span: Span, module_signatures: Mapping[str, CallableTypeSignature]
) -> CallArguments | None:
    """A module-qualified call statement (`ModuleName.Procedure ...`), proven by the
    project signature map; host/object member calls stay out of the validator."""
    slice_start = span.start
    toks = _significant_slice_tokens(source, span)
    if not toks:
        return None

    qualifier_idx = first_executable_token_index(toks)
    explicit_call = token_text(toks[qualifier_idx]) == "call"
    if explicit_call:
        qualifier_idx += 1
    qualifier = token_name(toks[qualifier_idx]) if qualifier_idx < len(toks) else None
    dot = toks[qualifier_idx + 1] if qualifier_idx + 1 < len(toks) else None
    callee = toks[qualifier_idx + 2] if qualifier_idx + 2 < len(toks) else None
    name = token_name(callee) if callee is not None else None
    if not qualifier or dot is None or dot.raw_text != "." or not name or callee is None:
        return None
    lookup_key = qualified_procedure_key(qualifier, name)
    if lookup_key not in module_signatures:
        return None

    next_tok = toks[qualifier_idx + 3] if qualifier_idx + 3 < len(toks) else None
    if explicit_call:
        if next_tok is not None and next_tok.kind is TokenKind.PUNCTUATION and next_tok.raw_text == "(":
            inner = _explicit_call_inner_args(toks, qualifier_idx + 3)
            if inner is None:
                return None
            arg_toks = inner
        else:
            arg_toks = []
    else:
        if next_tok is not None and next_tok.raw_text == "(":
            return None  # expression_calls handles parenthesized forms
        if next_tok is not None:
            gap = source[span.start + callee.end : span.start + next_tok.start]
            if not any(c.isspace() for c in gap):
                return None
        arg_toks = toks[qualifier_idx + 3 :]

    depth = 0
    for k in range(qualifier_idx + 3, len(toks)):
        raw = toks[k].raw_text
        if raw in ("(", "["):
            depth += 1
        elif raw in (")", "]"):
            depth -= 1
        elif depth == 0 and raw == "=":
            return None

    split = empty_arg_split() if not arg_toks else split_arg_slots(arg_toks, slice_start)
    return CallArguments(
        name=name,
        qualifier=qualifier,
        lookup_key=lookup_key,
        name_span=Span(span.start + callee.start, span.start + callee.end),
        explicit_call=explicit_call,
        slots=split.slots,
        slot_spans=split.spans,
        slice_start=slice_start,
    )


def _call_display_name(sig: CallableTypeSignature, call: CallArguments) -> str:
    return f"{call.qualifier}.{sig.name}" if call.qualifier else sig.name


def validate_arity(
    source: str, sig: CallableTypeSignature, call: CallArguments, push: PushFn
) -> None:
    """Validate a call's argument list against a procedure's parameters.

    Named arguments are checked by name (and ordering), skipping the positional
    count; otherwise the slot count is checked against the required minimum and the
    Optional/ParamArray-implied maximum. The quick-fix placeholder data the TS
    attaches is editor-only and is deferred (it does not affect which diagnostics
    fire), so this passes no diagnostic data.
    """
    if call_then_index(sig, call):
        return
    display_name = _call_display_name(sig, call)
    params = sig.params
    required = len(params)
    for k, param in enumerate(params):
        if param.optional or param.param_array:
            required = k
            break
    has_param_array = any(p.param_array for p in params)
    maximum: float = math.inf if has_param_array else len(params)

    named = [slot for slot in call.slots if is_named_slot(slot)]
    if named:
        saw_named = False
        slot_order_reported = False
        for i, slot in enumerate(call.slots):
            if is_named_slot(slot):
                saw_named = True
                continue
            if not saw_named:
                continue
            if slot:
                push(
                    "argumentCount",
                    f"A positional argument may not follow a named argument in the "
                    f"call to '{display_name}'.",
                    Span(call.slice_start + slot[0].start, call.slice_start + slot[-1].end),
                )
            else:
                push(
                    "argumentCount",
                    f"An omitted argument may not follow a named argument in the "
                    f"call to '{display_name}'.",
                    call.slot_spans[i] if call.slot_spans else call.name_span,
                )
            slot_order_reported = True
            break  # one syntax error per call, matching VBE
        param_names = {strip_header_brackets(p.name).lower() for p in params}
        # A procedure with a ParamArray takes no named argument at all, its fixed
        # parameters' names included: `PA(p0:=1)` (XLIDE issue #647, measured in
        # Excel 16.0).
        param_array = next((p for p in params if p.param_array), None)
        first_name = strip_header_brackets(named[0][0].raw_text).lower()
        if (
            param_array is not None
            and first_name
            and first_name in param_names
            and first_name != strip_header_brackets(param_array.name).lower()
        ):
            push(
                "argumentCount",
                f"'{display_name}' has a ParamArray, "
                f"'{strip_header_brackets(param_array.name)}', so no argument to it may "
                f"be named. This is a VBE compile error: Argument in ParamArray may not "
                f"be named.",
                Span(call.slice_start + named[0][0].start, call.slice_start + named[0][0].end),
            )
            return
        seen: set[str] = set()
        for slot in named:
            raw = strip_header_brackets(slot[0].raw_text)
            lower = raw.lower()
            if lower not in param_names:
                push(
                    "argumentCount",
                    f"Named argument not found: '{raw}' is not a parameter of '{display_name}'.",
                    Span(call.slice_start + slot[0].start, call.slice_start + slot[0].end),
                )
                continue
            # `Many(p:=1)`: a ParamArray takes no name (XLIDE issue #223, measured).
            named_param = next(
                (p for p in params if strip_header_brackets(p.name).lower() == lower), None
            )
            if named_param is not None and named_param.param_array:
                push(
                    "argumentCount",
                    f"'{raw}' is the ParamArray of '{display_name}', which may not be "
                    f"named. This is a VBE compile error: Argument in ParamArray may not "
                    f"be named.",
                    Span(call.slice_start + slot[0].start, call.slice_start + slot[0].end),
                )
                continue
            if lower in seen:
                push(
                    "argumentCount",
                    f"Named argument already specified: '{raw}' is supplied more than "
                    f"once to '{display_name}'.",
                    Span(call.slice_start + slot[0].start, call.slice_start + slot[0].end),
                )
                continue
            seen.add(lower)
        # `TTwo a:=1` leaves b out: every parameter that is not Optional needs an
        # argument, by position or by name (XLIDE issue #410, measured in Excel
        # 16.0). The names were all good, or a report above stands instead.
        if not slot_order_reported and len(seen) == len(named):
            first_named = next((k for k, s in enumerate(call.slots) if is_named_slot(s)), -1)
            missing = next(
                (
                    p
                    for k, p in enumerate(params)
                    if not p.optional
                    and not p.param_array
                    and not (k < first_named and len(call.slots[k]) > 0)
                    and strip_header_brackets(p.name).lower() not in seen
                ),
                None,
            )
            if missing is not None:
                push(
                    "argumentCount",
                    f"Parameter '{strip_header_brackets(missing.name)}' of '{display_name}' "
                    f"is not Optional, and the call gives it no argument, by position or by "
                    f"name. This is a VBE compile error: Argument not optional.",
                    call.name_span,
                )
        return  # positional count is not validated alongside named arguments

    n = len(call.slots)
    # `Opt(1, )`, `Many(1, )`, `Two(1, 2, )`: a list may not end in an empty
    # argument, whatever the parameters (XLIDE issue #223, measured in Excel
    # 16.0). `TwoSub 1,` and `Call TwoSub(1, )` too. Print is the exception:
    # `Debug.Print n,` ends in a separator of its own syntax.
    if n > 1 and len(call.slots[n - 1]) == 0 and sig.name.lower() != "print":
        push(
            "argumentCount",
            f"The call to '{display_name}' ends in an empty argument. This is a VBE "
            f"compile error: Syntax error.",
            call.slot_spans[n - 1] if call.slot_spans else call.name_span,
        )
        return
    # Per-slot 'Argument not optional' is only reported when the slot count is
    # in range: when it already violates min/max, the count check below emits the
    # single arity diagnostic instead, matching the named-arg path's one-error
    # policy (e.g. `Foo(, 2, 3)` against `Sub Foo(a, b)` yields one diagnostic).
    if required <= n <= maximum:
        for i in range(min(len(call.slots), len(params))):
            param = params[i]
            if not call.slots[i] and not param.optional and not param.param_array:
                name = strip_header_brackets(param.name)
                push(
                    "argumentCount",
                    f"Argument not optional: '{name}' is required by '{display_name}'.",
                    call.slot_spans[i] if call.slot_spans else call.name_span,
                )

    # A parameterless Function of the project named Left, InStr, InStrB or StrComp
    # does not hide VBA's when the call gives VBA's argument count:
    # `Left("abc", 1)` gives "a" (XLIDE issue #645, measured in Excel 16.0). Every
    # other VBA name, and Left with one or three arguments, is refused.
    if (
        len(params) == 0
        and not call.qualifier
        and n in _BUILT_INS_PAST_PARAMETERLESS.get(sig.name.lower(), ())
    ):
        return
    if n < required or n > maximum:
        push(
            "argumentCount",
            f"Wrong number of arguments to '{display_name}': expected "
            f"{describe_arity(required, maximum)}, but got {n}.",
            call.name_span,
        )


# VBA functions a parameterless project Function does not hide, by the argument
# counts that reach them (XLIDE issue #645).
_BUILT_INS_PAST_PARAMETERLESS: Mapping[str, tuple[int, ...]] = {
    "left": (2,),
    "instr": (2, 3),
    "instrb": (2,),
    "strcomp": (2,),
}
