"""Shared AST/statement traversal utilities for the diagnostics engine.

Ported from walker.ts. Pure: no rule logic, no diagnostics. The dataflow-coupled
helper localsPassedAsCallArguments is deferred until the dataflow layer (M7).
"""

from __future__ import annotations

import re
from collections.abc import Callable, Sequence
from typing import Protocol

from ..conditional import ConditionalActivityTracker
from ..lexer.token_helpers import match_paren_from, token_word
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import (
    BodyNode,
    LeafStatementNode,
    ModuleMember,
    ModuleNode,
    ProcedureNode,
    Span,
    VariableGroupNode,
    is_leaf_statement,
)
from .context import statement_tokens

# Re-export the lexer helpers under the walker's names so the diagnostics engine
# keeps one implementation (token_text == token_word).
token_text = token_word

__all__ = [
    "token_text",
    "match_paren_from",
    "statement_tokens",
    "is_inactive_node",
    "active_module_members",
    "for_each_statement",
    "ProcedureStatementVisitor",
    "walk_procedure_statements",
    "for_each_variable_group",
    "for_each_body_statement",
    "for_each_procedure_body_line",
    "next_line_start",
    "first_line_break_at_or_after",
    "statement_tokens_after_leading_label",
    "first_executable_token_index",
    "top_level_operator_index",
    "token_name",
    "strip_header_brackets",
    "absolute_span",
    "span_for_tokens",
    "bare_assignment_target",
    "set_assignment_target",
    "block_header_line_span",
    "block_footer_line_span",
    "declared_name_span",
    "first_token_span",
    "pluralize_count",
    "physical_line_span_at_offset",
]

_DECIMAL_RE = re.compile(r"^\d+$")


class _HasSpan(Protocol):
    @property
    def span(self) -> Span: ...


def is_inactive_node(activity: ConditionalActivityTracker | None, node: _HasSpan) -> bool:
    return activity is not None and activity.is_inactive(node.span)


def active_module_members(
    mod: ModuleNode, activity: ConditionalActivityTracker | None
) -> Sequence[ModuleMember]:
    if activity is None:
        return mod.members
    return [member for member in mod.members if not is_inactive_node(activity, member)]


def for_each_statement(
    body: Sequence[BodyNode],
    visit: Callable[[LeafStatementNode], None],
    activity: ConditionalActivityTracker | None = None,
) -> None:
    """Walk every leaf statement in a body, descending into nested blocks."""
    for node in body:
        if is_inactive_node(activity, node):
            continue
        if is_leaf_statement(node):
            visit(node)
        else:
            child = getattr(node, "body", None)
            if isinstance(child, list):
                for_each_statement(child, visit, activity)


# A per-procedure visitor of the shared statement walk: given a procedure, returns
# the per-statement callback to run inside it, or None to skip the procedure.
ProcedureStatementVisitor = Callable[[ProcedureNode], "Callable[[LeafStatementNode], None] | None"]


def _fan_out_statements(
    callbacks: list[Callable[[LeafStatementNode], None]],
) -> Callable[[LeafStatementNode], None]:
    def visit(stmt: LeafStatementNode) -> None:
        for callback in callbacks:
            callback(stmt)

    return visit


def walk_procedure_statements(
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    visitors: Sequence[ProcedureStatementVisitor],
) -> None:
    """Run every per-statement rule on ONE walk over active procedures/statements."""
    if len(visitors) == 0:
        return
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        callbacks: list[Callable[[LeafStatementNode], None]] = []
        for visitor in visitors:
            callback = visitor(member)
            if callback is not None:
                callbacks.append(callback)
        if len(callbacks) == 0:
            continue
        for_each_statement(member.body, _fan_out_statements(callbacks), activity)


def for_each_variable_group(
    body: Sequence[BodyNode],
    visit: Callable[[VariableGroupNode], None],
    activity: ConditionalActivityTracker | None = None,
) -> None:
    """Walk every VariableGroupNode in a body, descending into nested blocks."""
    for node in body:
        if is_inactive_node(activity, node):
            continue
        if isinstance(node, VariableGroupNode):
            visit(node)
        else:
            child = getattr(node, "body", None)
            if isinstance(child, list):
                for_each_variable_group(child, visit, activity)


def for_each_body_statement(
    body: Sequence[BodyNode],
    visit: Callable[[LeafStatementNode], None],
    activity: ConditionalActivityTracker | None = None,
) -> None:
    """Walk every leaf statement in a procedure body, descending into nested blocks."""
    for node in body:
        if is_inactive_node(activity, node):
            continue
        if is_leaf_statement(node):
            visit(node)
        else:
            child = getattr(node, "body", None)
            if isinstance(child, list):
                for_each_body_statement(child, visit, activity)


def for_each_procedure_body_line(
    source: str, procedure: ProcedureNode, visit: Callable[[Span], None]
) -> None:
    first_break = first_line_break_at_or_after(source, procedure.span.start)
    if first_break < 0 or first_break >= procedure.span.end:
        return
    line_start = next_line_start(source, first_break)
    while line_start < procedure.span.end:
        line_end = line_start
        while (
            line_end < procedure.span.end and source[line_end] != "\r" and source[line_end] != "\n"
        ):
            line_end += 1
        visit(Span(line_start, line_end))
        line_start = next_line_start(source, line_end)


def next_line_start(source: str, line_break_offset: int) -> int:
    if (
        line_break_offset < len(source)
        and source[line_break_offset] == "\r"
        and line_break_offset + 1 < len(source)
        and source[line_break_offset + 1] == "\n"
    ):
        return line_break_offset + 2
    return line_break_offset + 1


def first_line_break_at_or_after(source: str, start: int) -> int:
    for i in range(start, len(source)):
        ch = source[i]
        if ch == "\n" or ch == "\r":
            return i
    return -1


def statement_tokens_after_leading_label(source: str, span: Span) -> list[VbaToken]:
    toks = statement_tokens(source, span)
    first_executable = first_executable_token_index(toks)
    return toks[first_executable:] if first_executable > 0 else toks


def first_executable_token_index(toks: Sequence[VbaToken]) -> int:
    if len(toks) > 1 and toks[0].kind is TokenKind.INTEGER_LITERAL and _DECIMAL_RE.match(toks[0].raw_text):
        return 1
    if (
        len(toks) > 2
        and (toks[0].kind is TokenKind.IDENTIFIER or toks[0].kind is TokenKind.KEYWORD)
        and toks[1].raw_text == ":"
    ):
        return 2
    return 0


def top_level_operator_index(toks: Sequence[VbaToken], operator: str) -> int:
    depth = 0
    for i, tok in enumerate(toks):
        raw = tok.raw_text
        if raw == "(" or raw == "[":
            depth += 1
        elif raw == ")" or raw == "]":
            depth -= 1
        elif depth == 0 and tok.kind is TokenKind.OPERATOR and raw == operator:
            return i
    return -1


def strip_header_brackets(text: str) -> str:
    return text[1:-1] if text.startswith("[") and text.endswith("]") else text


def token_name(tok: VbaToken) -> str | None:
    if tok.kind is TokenKind.IDENTIFIER or tok.kind is TokenKind.KEYWORD:
        return tok.raw_text
    if tok.kind is TokenKind.BRACKETED_IDENTIFIER:
        return strip_header_brackets(tok.raw_text)
    return None


def absolute_span(base: Span, token: VbaToken) -> Span:
    return Span(base.start + token.start, base.start + token.end)


def span_for_tokens(toks: Sequence[VbaToken], slice_start: int) -> Span:
    """Absolute span covering a non-empty token slice (first.start .. last.end)."""
    return Span(slice_start + toks[0].start, slice_start + toks[-1].end)


def bare_assignment_target(
    source: str, span: Span
) -> tuple[str, Span, list[VbaToken]] | None:
    """For `name = ...` / `Let name = ...` returns (name, name-span, value tokens).

    Set (object) assignments and any LHS with a '.' or '(' are excluded.
    """
    toks = statement_tokens(source, span)
    i = first_executable_token_index(toks)
    if i < len(toks) and toks[i].kind is TokenKind.KEYWORD:
        kw = toks[i].raw_text.lower()
        if kw == "set":
            return None
        if kw == "let":
            i += 1
    name_tok = toks[i] if i < len(toks) else None
    if name_tok is None or name_tok.kind is not TokenKind.IDENTIFIER:
        return None
    nxt = toks[i + 1] if i + 1 < len(toks) else None
    if nxt is None or nxt.kind is not TokenKind.OPERATOR or nxt.raw_text != "=":
        return None
    return (
        name_tok.raw_text,
        Span(span.start + name_tok.start, span.start + name_tok.end),
        list(toks[i + 2 :]),
    )


def set_assignment_target(
    source: str, span: Span
) -> tuple[str, Span, list[VbaToken]] | None:
    toks = statement_tokens(source, span)
    i = first_executable_token_index(toks)
    if i >= len(toks) or token_text(toks[i]) != "set":
        return None
    name_tok = toks[i + 1] if i + 1 < len(toks) else None
    name = token_name(name_tok) if name_tok is not None else None
    if name_tok is None or not name:
        return None
    equals = toks[i + 2] if i + 2 < len(toks) else None
    if equals is None or equals.kind is not TokenKind.OPERATOR or equals.raw_text != "=":
        return None
    return (
        name,
        Span(span.start + name_tok.start, span.start + name_tok.end),
        list(toks[i + 3 :]),
    )


def block_header_line_span(source: str, span: Span) -> Span:
    nl = first_line_break_at_or_after(source, span.start)
    if nl < 0 or nl > span.end:
        return span
    return Span(span.start, nl)


def block_footer_line_span(source: str, span: Span) -> Span:
    start = span.end
    while start > span.start and source[start - 1] != "\n" and source[start - 1] != "\r":
        start -= 1
    return Span(start, span.end)


def declared_name_span(source: str, span: Span, name: str) -> Span:
    lower = name.lower()
    for tok in statement_tokens(source, span):
        candidate = token_name(tok)
        if candidate is not None and candidate.lower() == lower:
            return absolute_span(span, tok)
    return span


def first_token_span(source: str, span: Span) -> Span:
    toks = statement_tokens(source, span)
    return absolute_span(span, toks[0]) if toks else span


def pluralize_count(count: int, singular: str) -> str:
    return f"{count} {singular}{'' if count == 1 else 's'}"


def physical_line_span_at_offset(source: str, offset: int) -> Span:
    safe = max(0, min(offset, len(source)))
    before = source.rfind("\n", 0, max(0, safe - 1) + 1)
    start = 0 if before < 0 else before + 1
    after = source.find("\n", safe)
    end = len(source) if after < 0 else after
    if end > start and source[end - 1] == "\r":
        end -= 1
    return Span(start, end)
