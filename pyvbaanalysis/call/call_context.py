"""Shared VBA call-site/context helpers.

Ported from xlide_vscode/src/analyzer/call/callContext.ts. VBA distinguishes
expression calls, explicit ``Call`` statements, and parenless call statements,
and those contexts have different parenthesis rules. This module owns the
classification the diagnostics call rules depend on; the completion / signature-
help surfaces in the TS file are not needed by the analyzer port.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from ..lexer.keyword_table import STATEMENT_KEYWORDS as _STATEMENT_KEYWORD_LIST
from ..lexer.token_helpers import (
    match_paren_from,
    statement_tokens,
    token_name,
    token_word,
    tokens_without_leading_line_number,
)
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize
from ..parser.nodes import Span

# Statement-like words that are not in MS-VBAL's statement-keyword list but still
# cannot be treated as bare parenless call targets.
_ADDITIONAL_NON_CALL_STATEMENT_LEADS = (
    "then", "property", "error", "line", "name", "kill",
    "mkdir", "rmdir", "chdir", "chdrive", "load", "unload",
)

# Lowercase statement-leading words that must not be parsed as bare calls.
STATEMENT_KEYWORDS: frozenset[str] = frozenset(
    [word.lower() for word in _STATEMENT_KEYWORD_LIST]
) | frozenset(_ADDITIONAL_NON_CALL_STATEMENT_LEADS)

_WHITESPACE = re.compile(r"\s")
_DIGITS = re.compile(r"^\d+$")


@dataclass(frozen=True, slots=True)
class BareCallStatementTarget:
    name: str
    span: Span


@dataclass(frozen=True, slots=True)
class ParenthesizedCallStatementTarget:
    name: str
    span: Span
    empty_parens_span: Span
    is_member: bool
    starts_with_leading_dot: bool
    callee_end_offset: int


@dataclass(frozen=True, slots=True)
class ExplicitCallStatementArgumentList:
    callee_end_offset: int
    first_argument_span: Span
    argument_span: Span


def _statement_tokens_after_leading_line_number(source: str, span: Span) -> list[VbaToken]:
    return tokens_without_leading_line_number(statement_tokens(source, span.start, span.end))


def _leading_line_number_token_count(tokens: list[VbaToken]) -> int:
    return (
        1
        if len(tokens) > 1
        and tokens[0].kind is TokenKind.INTEGER_LITERAL
        and _DIGITS.match(tokens[0].raw_text)
        else 0
    )


def bare_call_statement_target(source: str, span: Span) -> BareCallStatementTarget | None:
    """The callee of a bare call statement, or None.

    Member calls, assignments, labels, statement keywords, and implicit
    Application index/member forms are intentionally excluded.
    """
    toks = _statement_tokens_after_leading_line_number(source, span)
    if not toks:
        return None

    idx = 0
    explicit_call = token_word(toks[0]) == "call"
    if explicit_call:
        idx = 1

    callee = toks[idx] if idx < len(toks) else None
    if callee is None or callee.kind is not TokenKind.IDENTIFIER:
        return None
    if callee.raw_text.lower() in STATEMENT_KEYWORDS:
        return None

    result = BareCallStatementTarget(
        name=callee.raw_text,
        span=Span(span.start + callee.start, span.start + callee.end),
    )

    next_tok = toks[idx + 1] if idx + 1 < len(toks) else None
    if next_tok is None:
        if not explicit_call:
            j = span.start + callee.end
            while j < len(source) and source[j] in (" ", "\t"):
                j += 1
            if j < len(source) and source[j] == ":":
                return None
        return result

    r = next_tok.raw_text
    if r in (".", ":"):
        return None
    if not explicit_call and r == "(":
        return None
    if not explicit_call:
        gap = source[span.start + callee.end : span.start + next_tok.start]
        if _WHITESPACE.search(gap) is None:
            return None
    depth = 0
    for k in range(idx + 1, len(toks)):
        tr = toks[k].raw_text
        if tr in ("(", "["):
            depth += 1
        elif tr in (")", "]"):
            depth -= 1
        elif depth == 0 and tr == "=":
            return None
    return result


def explicit_call_statement_target(source: str, span: Span) -> BareCallStatementTarget | None:
    """The callee name and span of an explicit ``Call`` statement, or None when the statement is not a Call."""
    toks = _statement_tokens_after_leading_line_number(source, span)
    if len(toks) < 2 or token_word(toks[0]) != "call":
        return None
    name = token_name(toks[1])
    if not name:
        return None
    return BareCallStatementTarget(
        name=name, span=Span(span.start + toks[1].start, span.start + toks[1].end)
    )


def explicit_call_statement_argument_without_parens(source: str, span: Span) -> Span | None:
    """The span of the first stray argument when an explicit ``Call`` passes arguments without enclosing parentheses, or None."""
    found = _explicit_call_statement_argument_list_without_parens(source, span)
    return found.first_argument_span if found is not None else None


def _explicit_call_statement_argument_list_without_parens(
    source: str, span: Span
) -> ExplicitCallStatementArgumentList | None:
    raw_toks = [t for t in tokenize(source[span.start : span.end]) if t.kind is not TokenKind.NEWLINE]
    toks = [t for t in raw_toks if t.kind is not TokenKind.COMMENT]
    start = _leading_line_number_token_count(toks)
    if len(toks) == start or token_word(toks[start]) != "call":
        return None
    consumed = _consume_callable_chain(toks, start + 1)
    if consumed is None or consumed <= start + 1:
        return None
    stray = toks[consumed] if consumed < len(toks) else None
    if stray is None or stray.raw_text == ":":
        return None
    end = span.end
    for tok in raw_toks:
        if tok.start < stray.start:
            continue
        if tok.kind is TokenKind.COMMENT:
            end = span.start + tok.start
            break
        if tok.raw_text == ":":
            return None
    while end > span.start and source[end - 1] in (" ", "\t"):
        end -= 1
    arg_start = span.start + stray.start
    if end <= arg_start:
        return None
    callee = toks[consumed - 1]
    return ExplicitCallStatementArgumentList(
        callee_end_offset=span.start + callee.end,
        first_argument_span=Span(arg_start, span.start + stray.end),
        argument_span=Span(arg_start, end),
    )


def standalone_empty_parenthesized_call_statement(
    source: str, span: Span
) -> ParenthesizedCallStatementTarget | None:
    """A parenless call statement whose callee is immediately followed by an empty ``()`` and nothing else, or None (e.g. ``Foo()`` or ``obj.Foo()`` as a whole statement)."""
    toks = _statement_tokens_after_leading_line_number(source, span)
    if len(toks) < 3 or token_word(toks[0]) == "call" or _top_level_token_index(toks, "=") >= 0:
        return None
    for i in range(len(toks) - 2):
        name = token_name(toks[i])
        if not name or toks[i + 1].raw_text != "(":
            continue
        close = match_paren_from(toks, i + 1)
        if (
            close != i + 2
            or close != len(toks) - 1
            or not _is_complete_statement_chain_through_empty_call(toks, i, close)
        ):
            continue
        return ParenthesizedCallStatementTarget(
            name=name,
            is_member=i > 0 and toks[i - 1].raw_text == ".",
            starts_with_leading_dot=toks[0].raw_text == ".",
            callee_end_offset=span.start + toks[i].end,
            empty_parens_span=Span(span.start + toks[i + 1].start, span.start + toks[close].end),
            span=Span(span.start + toks[i].start, span.start + toks[close].end),
        )
    return None


def _consume_callable_chain(tokens: list[VbaToken], start: int) -> int | None:
    if start >= len(tokens) or not token_name(tokens[start]):
        return None
    i = start + 1
    while True:
        t = tokens[i] if i < len(tokens) else None
        if t is None:
            return i
        if t.raw_text == ".":
            if i + 1 >= len(tokens) or not token_name(tokens[i + 1]):
                return i
            i += 2
            continue
        if t.raw_text == "(":
            close = match_paren_from(tokens, i)
            if close < 0:
                return None
            i = close + 1
            continue
        return i


def _is_complete_statement_chain_through_empty_call(
    toks: list[VbaToken], callee_idx: int, close_idx: int
) -> bool:
    if callee_idx == 0:
        return bool(token_name(toks[0]))
    first = toks[0]
    i = 1
    if first.raw_text == ".":
        name_idx = 1
        if name_idx >= len(toks) or not token_name(toks[name_idx]):
            return False
        if name_idx == callee_idx:
            return (
                name_idx + 1 < len(toks)
                and toks[name_idx + 1].raw_text == "("
                and match_paren_from(toks, name_idx + 1) == close_idx
            )
        i = name_idx + 1
    elif not token_name(first):
        return False
    while i < len(toks):
        raw = toks[i].raw_text
        if raw == "(":
            close = match_paren_from(toks, i)
            if close < 0 or close >= callee_idx:
                return False
            i = close + 1
            continue
        if raw != ".":
            return False
        name_idx = i + 1
        if name_idx >= len(toks) or not token_name(toks[name_idx]):
            return False
        if name_idx == callee_idx:
            return (
                name_idx + 1 < len(toks)
                and toks[name_idx + 1].raw_text == "("
                and match_paren_from(toks, name_idx + 1) == close_idx
            )
        i = name_idx + 1
    return False


def _top_level_token_index(tokens: list[VbaToken], raw_text: str) -> int:
    depth = 0
    for i, tok in enumerate(tokens):
        raw = tok.raw_text
        if raw in ("(", "["):
            depth += 1
        elif raw in (")", "]"):
            depth -= 1
        elif depth == 0 and raw == raw_text:
            return i
    return -1
