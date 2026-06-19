"""Call-site extraction and argument-list utilities.

Ported from xlide_vscode/src/analyzer/diagnostics/callExtraction.ts. Owns the
value model for a call (CallArguments, the callable-signature shapes, inferred
argument types) and the argument-slot splitter that the call/argument rules
share. The statement-level call extractors (extract_call / extract_qualified_call)
and arity validation land with their consumer rules.
"""

from __future__ import annotations

import math
import re
from collections.abc import Sequence
from dataclasses import dataclass, field

from ..lexer.token_helpers import match_paren_from
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import Span
from .walker import strip_header_brackets

_STRING_LITERAL_QUOTES = (re.compile(r'^"'), re.compile(r'"$'))


@dataclass(slots=True)
class CallableParamType:
    name: str
    type_: str | None = None
    optional: bool = False
    param_array: bool = False
    is_array: bool | None = None
    by_ref: bool | None = None


@dataclass(slots=True)
class CallableTypeSignature:
    name: str
    params: list[CallableParamType] = field(default_factory=list)
    return_type: str | None = None


@dataclass(slots=True)
class InferredArgumentType:
    type_: str
    label: str
    span: Span
    string_value: str | None = None
    numeric_value: float | None = None
    numeric_text: str | None = None


@dataclass(slots=True)
class CallArguments:
    name: str
    name_span: Span
    slots: list[list[VbaToken]]
    slice_start: int
    qualifier: str | None = None
    lookup_key: str | None = None
    explicit_call: bool = False
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
    """True if a slot is a named argument (`name := value`)."""
    return (
        len(slot) >= 2
        and slot[0].kind in (TokenKind.IDENTIFIER, TokenKind.BRACKETED_IDENTIFIER)
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
