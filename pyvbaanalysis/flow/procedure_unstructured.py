"""Detection of control flow the structural branch-merge cannot model soundly.

Ported from xlide_vscode/src/analyzer/flow/procedureUnstructured.ts. Dataflow
rules fall back to the conservative straight-line walk for procedures whose flow
can skip or re-run assignments in ways the If/ElseIf/Else merge cannot see.
"""

from __future__ import annotations

from collections.abc import Sequence

from ..conditional import ConditionalActivityTracker, inactive_node_skip
from ..parser.nodes import BodyNode, ProcedureNode, is_leaf_statement, iter_body_nodes
from .procedure_labels import statement_has_unstructured_flow


def procedure_has_unstructured_flow(
    source: str,
    procedure: ProcedureNode,
    activity: ConditionalActivityTracker | None = None,
) -> bool:
    """True when a procedure contains label / GoTo / On Error / Resume flow.

    Any label, any GoTo / GoSub / On..GoTo / On..GoSub / Resume target, or any
    `On Error` / `Resume` statement (whose exception edges can bypass an
    assignment the merge would assume ran) forces the conservative straight-line
    dataflow, preserving the no-false-positive contract. Upstream memoizes the
    result per procedure node; the port recomputes it.
    """
    return _has_unstructured_statement(procedure.body, source, activity)


def _has_unstructured_statement(
    body: Sequence[BodyNode], source: str, activity: ConditionalActivityTracker | None
) -> bool:
    return any(
        is_leaf_statement(node) and statement_has_unstructured_flow(source, node.span)
        for node in iter_body_nodes(body, inactive_node_skip(activity))
    )
