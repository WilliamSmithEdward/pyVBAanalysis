"""Shared AST/statement traversal utilities for the diagnostics engine.

Ported from walker.ts. Pure: no rule logic, no diagnostics. The dataflow-coupled
helper tracked_locals_named_whole lives in the dataflow module; locals_named_whole
here reaches it through a function-local import. for_each_statement lives in
parser/statement_walk.py and is re-exported.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterator, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass, replace
from typing import Protocol, cast

from ..conditional import ConditionalActivityTracker, inactive_node_skip
from ..lexer.token_helpers import match_paren_from, token_word
from ..lexer.token_helpers import statement_tokens as lex_statement_tokens
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import (
    BodyNode,
    IfBlockNode,
    LeafStatementNode,
    ModuleMember,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    VariableGroupNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ..parser.statement_walk import for_each_statement
from .block_headers import block_header_statements
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
    "for_each_statement_with_headers",
    "block_header_statements",
    "ProcedureStatementVisitor",
    "ProcedureWalkHooks",
    "StatementHeaders",
    "walk_procedure_statements",
    "locals_named_whole",
    "for_each_variable_group",
    "for_each_body_statement",
    "for_each_procedure_body_line",
    "next_line_start",
    "first_line_break_at_or_after",
    "raw_expression_tokens",
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
_LEADING_HASH_RE = re.compile(r"^[ \t]*#")


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


def for_each_statement_with_headers(
    source: str,
    body: Sequence[BodyNode],
    visit: Callable[[LeafStatementNode], None],
    activity: ConditionalActivityTracker | None = None,
) -> None:
    """for_each_statement, with each block's header line visited as a statement of
    its own before the body, and a Do's `Loop While` line after it
    (block_header_statements, XLIDE issue #233). For a rule that judges an
    expression wherever it stands."""
    _walk_with_headers(
        body, activity, visit, visit, lambda node: block_header_statements(source, node)
    )


def _walk_with_headers(
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    visit: Callable[[LeafStatementNode], None],
    visit_header: Callable[[LeafStatementNode], None],
    headers_of: Callable[[BodyNode], tuple[StatementNode | None, StatementNode | None]],
    opening_of: Callable[[BodyNode], StatementNode | None] | None = None,
) -> None:
    """Every leaf statement of a body, with each block's headers around its body:
    `before` and the `opening_of` line ahead of it, `after` behind it. On an
    explicit stack, where upstream recurses once per block: a block pushes its
    `after` line and then its body, so the body is walked first."""
    stack: list[Iterator[BodyNode] | StatementNode] = [iter(body)]
    while stack:
        top = stack[-1]
        if isinstance(top, StatementNode):
            stack.pop()
            visit_header(top)
            continue
        for node in top:
            if is_inactive_node(activity, node):
                continue
            if is_leaf_statement(node):
                visit(node)
                continue
            child = getattr(node, "body", None)
            if isinstance(child, list):
                before, after = headers_of(node)
                if before is not None:
                    visit_header(before)
                if opening_of is not None:
                    opening = opening_of(node)
                    if opening is not None:
                        visit_header(opening)
                if after is not None:
                    stack.append(after)
                stack.append(iter(child))
                break
        else:
            stack.pop()


# A per-procedure visitor of the shared statement walk: given a procedure, returns
# the per-statement callback to run inside it, or None to skip the procedure.
ProcedureStatementVisitor = Callable[[ProcedureNode], "Callable[[LeafStatementNode], None] | None"]


@dataclass(frozen=True, slots=True)
class ProcedureWalkHooks:
    """Hooks of the shared statement and expression walks."""

    # Called before a procedure's factories run (incremental attribution).
    before_member: Callable[[ProcedureNode], None] | None = None
    # When it returns true, the member is skipped before its visitors are built.
    skip_body: Callable[[ProcedureNode], bool] | None = None


# Upstream's `{ source, takes }`: the module's text, and for each visitor whether
# it also takes block headers (block_header_statements). Without it no visitor does.
StatementHeaders = tuple[str, Sequence[bool]]


def walk_procedure_statements(
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    visitors: Sequence[ProcedureStatementVisitor],
    hooks: ProcedureWalkHooks | None = None,
    headers: StatementHeaders | None = None,
) -> None:
    """Run every per-statement rule on ONE walk over active procedures/statements."""
    takes_headers: Sequence[bool] = headers[1] if headers is not None else ()
    if len(visitors) == 0:
        return
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        if hooks is not None and hooks.before_member is not None:
            hooks.before_member(member)
        # Skipped before its visitors are built: an incremental pass walks the
        # edited procedure only, and asking every rule for a visitor for each of
        # the others cost a large module tens of milliseconds a keystroke. A
        # visitor is per member (cross-member state belongs in a run rule), so a
        # skipped one had nothing to contribute.
        if hooks is not None and hooks.skip_body is not None and hooks.skip_body(member):
            continue
        callbacks: list[Callable[[LeafStatementNode], None]] = []
        header_callbacks: list[Callable[[LeafStatementNode], None]] = []
        for k, visitor in enumerate(visitors):
            callback = visitor(member)
            if callback is not None:
                callbacks.append(callback)
                if k < len(takes_headers) and takes_headers[k]:
                    header_callbacks.append(callback)
        if len(callbacks) == 0:
            continue
        _walk_procedure_body(member.body, activity, callbacks, header_callbacks, headers)


def _walk_procedure_body(
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    callbacks: list[Callable[[LeafStatementNode], None]],
    header_callbacks: list[Callable[[LeafStatementNode], None]],
    headers: StatementHeaders | None,
) -> None:
    def visit(stmt: LeafStatementNode) -> None:
        for callback in callbacks:
            callback(stmt)

    if headers is None or not header_callbacks:
        for_each_statement(body, visit, activity)
        return
    source = headers[0]

    def visit_header(stmt: LeafStatementNode) -> None:
        # A header line goes to the visitors that take headers only.
        for callback in header_callbacks:
            callback(stmt)

    def opening_of(node: BodyNode) -> StatementNode | None:
        # A block If's own line: `If 10 / d > 1 Then` evaluates its condition as a
        # statement does (XLIDE issue #492). Its ElseIf lines are statements of the
        # body already.
        if not isinstance(node, IfBlockNode) or not node.branches:
            return None
        opening = node.branches[0].header_span
        return StatementNode(span=opening, raw=source[opening.start : opening.end])

    _walk_with_headers(
        body,
        activity,
        visit,
        visit_header,
        lambda node: block_header_statements(source, node),
        opening_of,
    )


def for_each_variable_group(
    body: Sequence[BodyNode],
    visit: Callable[[VariableGroupNode], None],
    activity: ConditionalActivityTracker | None = None,
) -> None:
    """Walk every VariableGroupNode in a body, descending into nested blocks."""
    for node in iter_body_nodes(body, inactive_node_skip(activity)):
        if isinstance(node, VariableGroupNode):
            visit(node)


def for_each_body_statement(
    body: Sequence[BodyNode],
    visit: Callable[[LeafStatementNode], None],
    activity: ConditionalActivityTracker | None = None,
) -> None:
    """Walk every leaf statement in a procedure body, descending into nested blocks."""
    for node in iter_body_nodes(body, inactive_node_skip(activity)):
        if is_leaf_statement(node):
            visit(node)


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


def raw_expression_tokens(text: str) -> list[VbaToken]:
    """Significant tokens of an expression the parser carried as its own string (an
    If condition, a For Each source, an Enum member value, a parameter default, a
    Const value). That text is not the module source, so it must not go through
    statement_tokens: the statement cache is keyed by source string and holds two
    of them, and each distinct raw string sent there evicted the module, so the next
    ordinary statement re-lexed the whole module (XLIDE issue #139, a 17x slowdown
    on real projects)."""
    # A `#` that opens a line is a directive to the lexer, and an expression never
    # opens one: `#12/31/9999#` is a date (XLIDE issue #255). It is lexed behind an
    # `=` and moved back, so every offset stays the text's.
    if not _LEADING_HASH_RE.match(text):
        return lex_statement_tokens(text, 0, len(text))
    return [
        replace(tok, start=tok.start - 1, end=tok.end - 1)
        for tok in lex_statement_tokens(f"={text}", 0, len(text) + 1)[1:]
    ]


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


def token_name(tok: VbaToken | None) -> str | None:
    if tok is None:
        return None
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


def statement_and_branch_spans(stmt: LeafStatementNode) -> list[Span]:
    """The statement's own span, plus the branches a single-line `If` executes.

    Rules that scan statement text need the branches explicitly: the walk itself
    does not descend into a single-line If, whose arms are part of one statement
    (XLIDE issue #46).
    """
    branches = getattr(stmt, "single_line_if_branches", None)
    return [stmt.span, *branches] if branches else [stmt.span]


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
    # A name that SPELLS a keyword is still a name. The lexer classifies `Text`,
    # `Read` and `Type` as keywords, so requiring an identifier here hid every
    # assignment to a variable or Function named one of them: `Function Read()`
    # assigning `Read = True` read as never assigning its own return (XLIDE issue
    # #46). The `=` that follows is what settles it: no VBA statement keyword is
    # followed by a bare `=` at statement start, and Set/Let are handled above.
    if name_tok is None or name_tok.kind not in (TokenKind.IDENTIFIER, TokenKind.KEYWORD):
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


def locals_named_whole(
    source: str,
    span: Span,
    tracked: Mapping[str, object] | AbstractSet[str],
    read_only_intrinsics: AbstractSet[str],
) -> dict[str, int]:
    """Every tracked local the statement names whole - bare, not the statement's
    own head, not a member access, not indexed - in an argument position: a call
    statement's argument, an argument to a function inside an expression, or an
    argument to a qualified member call. VBA passes by reference by default, so
    the callee may have assigned or allocated the caller's variable and its state
    is unknown from that point on (XLIDE #70). A mention the callee provably only
    reads is left out: the operand of `Is`, and the argument of an intrinsic in
    `read_only_intrinsics`. The value is the first such mention's absolute
    offset, so a rule can tell an access before the pass from one after it within
    the same statement."""
    from .callee_arguments import callee_keeps_argument
    from .dataflow import tracked_locals_named_whole

    # A procedure of the module that cannot change the argument keeps what is
    # known about it (XLIDE issue #449).
    keeps: object = callee_keeps_argument(source)
    return tracked_locals_named_whole(
        statement_tokens_after_leading_label(source, span),
        span.start,
        lambda lower: lower in tracked,
        read_only_intrinsics,
        frozenset(),
        cast("Callable[[str, int, str | None], bool]", keeps),
    )


def block_header_line_span(source: str, span: Span) -> Span:
    """The block's header line, with the lines a ` _` continues it onto."""
    nl = first_line_break_at_or_after(source, span.start)
    while 0 <= nl <= span.end and _ends_in_continuation(source, span.start, nl):
        nxt = nl + 2 if source[nl] == "\r" and nl + 1 < len(source) and source[nl + 1] == "\n" else nl + 1
        nl = first_line_break_at_or_after(source, nxt)
    if nl < 0 or nl > span.end:
        return span
    return Span(span.start, nl)


def _ends_in_continuation(source: str, start: int, nl: int) -> bool:
    i = nl - 1
    while i >= start and source[i] in (" ", "\t"):
        i -= 1
    return i > start and source[i] == "_" and source[i - 1] in (" ", "\t")


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
