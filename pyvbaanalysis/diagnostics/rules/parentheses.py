"""Rule family: parentheses the VBE reads otherwise, or refuses (XLIDE issue #236).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/parentheses.ts. Measured
in Excel 16.0 (build 20326, 2026-09-30) with a full compile.

A paren that follows a name, `)` or `]` calls or indexes; any other paren groups
an expression, and so does one after a statement's callee with a space between,
`Take (c)`.

 - collection-operand: an object in grouping parentheses is its default member,
   and a Collection's Item needs an index: `Set d = (c)`,
   `Set d = (New Collection)`, `c.Add (c)`, `Take (c)`, `Call Take((c))`,
   `With (New Collection)` -> "Argument not optional".
 - malformed-statement: `(Range("A1")).Address` and `(Application).Name`, a
   member of a grouping paren -> "Syntax error" ("Invalid or unqualified
   reference" after Print); `v = ()`; `UBound((a))`; `G((b:=1), a:=2)`, a named
   argument inside one; `Mid((s), 1, 1) = "x"`; `(n) = 1`, a statement that
   starts with one -> "Syntax error".
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass

from ...completion.member_access import MemberCompletionContext
from ...conditional import ConditionalActivityTracker
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...symbols.symbol_model import ModuleSymbols
from ...types.type_inference import object_value_needs_index, type_environment_for
from ...types.type_names import normalize_type
from ..context import PushFn, statement_tokens
from ..walker import (
    active_module_members,
    first_executable_token_index,
    for_each_statement_with_headers,
    match_paren_from,
    token_name,
    token_text,
)

# Keywords that an expression can follow: a paren after one groups.
_OPERAND_KEYWORDS: frozenset[str] = frozenset({
    "and", "or", "xor", "eqv", "imp", "not", "mod", "like", "is", "to", "step", "then", "else", "elseif",
    "if", "while", "until", "case", "print", "call", "return", "with", "in", "each", "set", "let",
    "select", "do", "loop", "for", "redim", "erase", "typeof", "new",
})

_MID_STATEMENTS: frozenset[str] = frozenset({"mid", "mid$", "midb", "midb$"})

_TYPE_CHARACTER_RE = re.compile(r"^[$%&!#@]\Z")
_VOWEL_HEAD_RE = re.compile(r"^[aeiou]", re.IGNORECASE | re.ASCII)

_NeedsIndex = Callable[[Sequence[VbaToken], int, int], "str | None"]


def _at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    return toks[index] if 0 <= index < len(toks) else None


def _raw(toks: Sequence[VbaToken], index: int) -> str | None:
    tok = _at(toks, index)
    return tok.raw_text if tok is not None else None


def check_parentheses(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        env = type_environment_for(symbols, member)

        def needs_index(
            toks: Sequence[VbaToken], from_: int, to: int, env: Mapping[str, str] = env
        ) -> str | None:
            # `New Collection`, or a variable of such a type.
            type_name: str | None
            if to - from_ == 2 and token_text(toks[from_]) == "new":
                type_name = token_name(toks[from_ + 1])
            elif to - from_ == 1:
                name = token_name(toks[from_])
                type_name = env.get(name.lower() if name else "")
            else:
                type_name = None
            return type_name if type_name and object_value_needs_index(type_name, member_ctx) else None

        def each_statement(stmt: LeafStatementNode, needs_index: _NeedsIndex = needs_index) -> None:
            _check_statement(source, stmt.span, needs_index, push)

        for_each_statement_with_headers(source, member.body, each_statement, activity)


def _check_statement(source: str, span: Span, needs_index: _NeedsIndex, push: PushFn) -> None:
    all_toks = statement_tokens(source, span)
    toks = all_toks[first_executable_token_index(all_toks) :]
    if len(toks) == 0:
        return

    def at(tok: VbaToken) -> Span:
        return Span(span.start + tok.start, span.start + tok.end)

    def syntax(message: str, where: Span, error: str = "Syntax error") -> None:
        push("malformedStatement", f"{message}. This is a VBE compile error: {error}.", where)

    if toks[0].raw_text == "(":
        syntax("A statement cannot start with a parenthesis", at(toks[0]))
        return
    head = token_text(toks[0])
    # Where a statement's callee ends: `Take (c)` and `c.Add (c)` group their
    # argument; `Take(c)` in an expression calls.
    callee = 0
    while _raw(toks, callee + 1) == "." and token_name(_at(toks, callee + 2)):
        callee += 2

    def grouping(i: int) -> bool:
        prev = _at(toks, i - 1)
        if prev is None:
            return True
        # `Mid$(`: a type character glued to a name is part of the name.
        before = _at(toks, i - 2)
        if (
            _TYPE_CHARACTER_RE.match(prev.raw_text) is not None
            and before is not None
            and before.end == prev.start
            and token_name(before)
        ):
            prev = before
        if i - 1 == callee and prev.end < toks[i].start and token_name(prev):
            return True
        if prev.raw_text in (")", "]") or prev.kind in (TokenKind.IDENTIFIER, TokenKind.BRACKETED_IDENTIFIER):
            return False
        return prev.kind is not TokenKind.KEYWORD or token_text(prev) in _OPERAND_KEYWORDS

    # A single group is already linear; avoid index allocation on the common path.
    first_open = next((k for k, tok in enumerate(toks) if tok.raw_text == "("), -1)
    if first_open < 0:
        return
    multiple = any(index > first_open and tok.raw_text == "(" for index, tok in enumerate(toks))
    facts = _parenthesis_facts(toks) if multiple else None
    for i in range(first_open, len(toks)):
        if toks[i].raw_text != "(":
            continue
        close = facts.closes[i] if facts is not None else match_paren_from(toks, i)
        if close < 0:
            return
        if not grouping(i):
            # `UBound((a))`, `UBound((a) + 0)`: the array must be a name.
            name = token_text(toks[i - 1])
            if name in ("ubound", "lbound") and _raw(toks, i + 1) == "(":
                syntax(f"{toks[i - 1].raw_text} takes an array by its name, not in parentheses", at(toks[i + 1]))
            # `Mid((s), 1, 1) = "x"`: the Mid statement writes into a variable.
            if i == 1 and head in _MID_STATEMENTS and _raw(toks, i + 1) == "(":
                syntax("The Mid statement writes into a variable, not a value in parentheses", at(toks[i + 1]))
            continue
        if close == i + 1:
            syntax("Empty parentheses hold no value", at(toks[i]))
            continue
        named = facts.next_named[i + 1] if facts is not None else _first_named_argument(toks, i + 1, close)
        # Without facts there is only one opening paren, so its contents cannot
        # contain a deeper group before this matching close.
        if named < close and (facts is None or facts.depths[named] == facts.depths[i] + 1):
            syntax("A named argument cannot stand inside parentheses", at(toks[named]))
        if _raw(toks, close + 1) == ".":
            after_print = token_text(_at(toks, i - 1)) == "print"
            syntax(
                "A member cannot be read from a value in parentheses: after Print this reads as a With member"
                if after_print
                else "A member cannot be read from a value in parentheses",
                at(toks[close + 1]),
                "Invalid or unqualified reference" if after_print else "Syntax error",
            )
            continue
        # A grouping paren always evaluates what it holds: an object whose default
        # member needs an index has no value there.
        type_name = needs_index(toks, i + 1, close)
        if type_name:
            where = Span(span.start + toks[i].start, span.start + toks[close].end)
            push(
                "collectionOperand",
                f"'{source[where.start : where.end]}' is the default member of {_article(type_name)} "
                f"{type_name}, whose Item needs an index. This is a VBE compile error: Argument not "
                "optional.",
                where,
            )


@dataclass(frozen=True, slots=True)
class _ParenthesisFacts:
    closes: list[int]
    depths: list[int]
    next_named: list[int]


def _parenthesis_facts(toks: Sequence[VbaToken]) -> _ParenthesisFacts:
    """Match groups and locate their first named token without rescanning nested slices."""
    count = len(toks)
    closes = [-1] * count
    depths = [0] * count
    next_named = [count] * (count + 1)
    stack: list[int] = []
    depth = 0
    for i, tok in enumerate(toks):
        depths[i] = depth
        if tok.raw_text == "(":
            stack.append(i)
            depth += 1
        elif tok.raw_text == ")":
            depth -= 1
            if stack:
                closes[stack.pop()] = i
    for i in range(count - 1, -1, -1):
        next_named[i] = i if toks[i].raw_text == ":=" else next_named[i + 1]
    return _ParenthesisFacts(closes, depths, next_named)


def _first_named_argument(toks: Sequence[VbaToken], from_: int, to: int) -> int:
    for i in range(from_, to):
        if toks[i].raw_text == ":=":
            return i
    return to


def _article(type_name: str) -> str:
    normalized = normalize_type(type_name)
    return "an" if _VOWEL_HEAD_RE.match(normalized if normalized is not None else type_name) else "a"
