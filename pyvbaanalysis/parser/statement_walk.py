"""Ported from xlide_vscode/src/analyzer/parser/statementWalk.ts."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from typing import TYPE_CHECKING

from .nodes import BodyNode, LeafStatementNode, is_leaf_statement, iter_body_nodes

if TYPE_CHECKING:
    from ..conditional import ConditionalActivityTracker


def for_each_statement(
    body: Sequence[BodyNode],
    visit: Callable[[LeafStatementNode], None],
    activity: ConditionalActivityTracker | None = None,
) -> None:
    """Walk every leaf statement (Assignment/Call/Statement) in a body, descending
    into nested blocks."""
    # On an explicit stack (iter_body_nodes): upstream recurses once per block.
    skip = None if activity is None else (lambda node: activity.is_inactive(node.span))
    for node in iter_body_nodes(body, skip):
        if is_leaf_statement(node):
            visit(node)
