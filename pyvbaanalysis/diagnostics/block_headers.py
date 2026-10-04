"""A block's own lines as statements (XLIDE issues #233 and #237).

Ported from xlide_vscode/src/analyzer/diagnostics/blockHeaders.ts. The line a
For, Select, Do, While or With opens with, a Do's `Loop While` line, and each If
arm's condition line. Rules judge the expressions there, and a walk that enters
a block counts the names they pass ByRef as ones the block changes.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import NamedTuple

from ..lexer.token_helpers import token_word
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize_cached
from ..parser.nodes import (
    BodyNode,
    DoBlockNode,
    ForBlockNode,
    IfBlockNode,
    IfBranchKind,
    SelectBlockNode,
    Span,
    StatementNode,
    WhileBlockNode,
    WithBlockNode,
)
from .context import statement_tokens

# The block kinds whose own line evaluates an expression: a For's bounds and
# step, a Select Case subject, a Do or While condition, a With subject.
_HEADER_BLOCKS = (ForBlockNode, SelectBlockNode, DoBlockNode, WhileBlockNode, WithBlockNode)


def _separator(tok: VbaToken) -> bool:
    return tok.kind is TokenKind.NEWLINE or tok.kind is TokenKind.COLON


class BlockHeaders(NamedTuple):
    """Upstream's `{ before?, after? }`: read by name or unpacked."""

    before: StatementNode | None = None
    after: StatementNode | None = None


_NO_HEADERS = BlockHeaders()


def block_header_statements(source: str, node: BodyNode) -> BlockHeaders:
    """A block's header line as a statement of its own, and a Do's `Loop While` or
    `Loop Until` line (XLIDE issue #233), as (before, after). The header ends at
    the end of its logical line or at a colon, so `If a Then With c: .Add 1: End
    With` gives `With c` alone, and a string holding a colon stays whole."""
    if not isinstance(node, _HEADER_BLOCKS):
        return _NO_HEADERS
    toks = tokenize_cached(source)

    def statement(from_: int, to: int) -> StatementNode:
        span = Span(toks[from_].start, toks[to].end)
        return StatementNode(span=span, raw=source[span.start : span.end])

    # The first token at or after the block's start.
    lo = 0
    hi = len(toks)
    while lo < hi:
        mid = (lo + hi) >> 1
        if toks[mid].start < node.span.start:
            lo = mid + 1
        else:
            hi = mid
    before: StatementNode | None = None
    after: StatementNode | None = None
    end = lo
    while (
        end + 1 < len(toks)
        and not _separator(toks[end + 1])
        and toks[end + 1].kind is not TokenKind.COMMENT
    ):
        end += 1
    if lo < len(toks) and not _separator(toks[lo]):
        before = statement(lo, end)
    if isinstance(node, DoBlockNode) and node.closed:
        # The last token of the block, and back to the start of its statement.
        last = hi
        while last < len(toks) and toks[last].end <= node.span.end:
            last += 1
        last -= 1
        while last > end and toks[last].kind is TokenKind.COMMENT:
            last -= 1
        first = last
        while first - 1 > end and not _separator(toks[first - 1]):
            first -= 1
        if first > end and token_word(toks[first]) == "loop" and first < last:
            after = statement(first, last)
    return BlockHeaders(before, after)


def block_header_leaves(source: str, node: BodyNode) -> list[StatementNode]:
    """Every line of a block that can change a variable it names, as a walk that
    enters the block must know: block_header_statements, and each If and ElseIf
    condition, `If TryGet(k, obj) Then` among them. A With's subject counts: its
    body acts on it through `.Add` and the like without naming it."""
    if isinstance(node, IfBlockNode):
        return [
            StatementNode(
                span=branch.header_span,
                raw=source[branch.header_span.start : branch.header_span.end],
            )
            for branch in node.branches
            if branch.branch_kind is not IfBranchKind.ELSE
        ]
    before, after = block_header_statements(source, node)
    return [stmt for stmt in (before, after) if stmt is not None]


def is_loop_block(node: BodyNode) -> bool:
    """True for a For, Do or While, whose body may run more than once."""
    return isinstance(node, (ForBlockNode, DoBlockNode, WhileBlockNode))


def select_arms(source: str, body: Sequence[BodyNode]) -> list[list[BodyNode]]:
    """A Select block's body as its arms: each `Case` line with the statements
    under it. Only one arm runs, so a walk enters each from the same state."""
    arms: list[list[BodyNode]] = []
    for node in body:
        head = ""
        if isinstance(node, StatementNode):
            first = next(
                (
                    tok
                    for tok in statement_tokens(source, node.span)
                    if tok.kind is not TokenKind.INTEGER_LITERAL
                ),
                None,
            )
            head = token_word(first)
        if head == "case" or not arms:
            arms.append([])
        arms[-1].append(node)
    return arms
