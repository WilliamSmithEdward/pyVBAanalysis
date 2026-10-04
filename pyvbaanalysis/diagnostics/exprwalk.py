"""Shared expression-tree traversal for diagnostics (MS-VBAL 5.6).

Ported from exprWalk.ts. Owns the one canonical two-level walk over a procedure
body that expression-consuming rules share: find every root expression (assignment
sides, call callee + arguments, If/ElseIf conditions, nested block bodies) and
recurse each into its sub-expressions in pre-order.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence

from ..conditional import ConditionalActivityTracker, inactive_node_skip
from ..parser.nodes import (
    AssignmentNode,
    BinaryExpr,
    BodyNode,
    CallNode,
    ExprNode,
    IfBlockNode,
    IndexExpr,
    MemberAccessExpr,
    ModuleNode,
    ParenExpr,
    ProcedureNode,
    TypeOfIsExpr,
    UnaryExpr,
    iter_body_nodes,
)
from .walker import ProcedureWalkHooks, active_module_members

# A rule's per-procedure expression visitor: the factory does per-member setup and
# returns a callback invoked for every expression node in that member's body.
ProcedureExpressionVisitor = Callable[[ProcedureNode], Callable[[ExprNode], None]]


def _fan_out_expressions(
    visitors: list[Callable[[ExprNode], None]],
) -> Callable[[ExprNode], None]:
    def visit(expr: ExprNode) -> None:
        for v in visitors:
            v(expr)

    return visit


def walk_procedure_expressions(
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    factories: Sequence[ProcedureExpressionVisitor],
    hooks: ProcedureWalkHooks | None = None,
) -> None:
    """Run ONE shared expression walk per active procedure, dispatching to each visitor."""
    if len(factories) == 0:
        return
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        if hooks is not None and hooks.before_member is not None:
            hooks.before_member(member)
        # Skipped before its visitors are built, as in walk_procedure_statements.
        if hooks is not None and hooks.skip_body is not None and hooks.skip_body(member):
            continue
        visitors = [factory(member) for factory in factories]
        for_each_expression_in_body(member.body, activity, _fan_out_expressions(visitors))


def for_each_expression_in_body(
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    visit: Callable[[ExprNode], None],
) -> None:
    """Visit every expression node reachable in a body, skipping inactive regions."""
    for node in iter_body_nodes(body, inactive_node_skip(activity)):
        if isinstance(node, AssignmentNode):
            for_each_sub_expression(node.lhs, visit)
            for_each_sub_expression(node.rhs, visit)
        elif isinstance(node, CallNode):
            for_each_sub_expression(node.callee, visit)
            for arg in node.args:
                if arg.value is not None:
                    for_each_sub_expression(arg.value, visit)
        elif isinstance(node, IfBlockNode):
            # Arm statements live in the flat body, which the walk enters next.
            for branch in node.branches:
                if branch.condition is not None:
                    for_each_sub_expression(branch.condition, visit)


def for_each_sub_expression(expr: ExprNode, visit: Callable[[ExprNode], None]) -> None:
    """Visit expr and every nested sub-expression (pre-order)."""
    # A stack of the subtrees still to visit, not recursion: a left-deep chain such
    # as `a + b + c ...` or `a.b.c ...` nests one level per operand, which in one
    # long statement goes past Python's recursion limit.
    pending = [expr]
    while pending:
        node = pending.pop()
        visit(node)
        # Children are pushed last-first so they come off the stack in order.
        if isinstance(node, BinaryExpr):
            pending.append(node.right)
            pending.append(node.left)
        elif isinstance(node, UnaryExpr):
            pending.append(node.operand)
        elif isinstance(node, ParenExpr):
            pending.append(node.inner)
        elif isinstance(node, IndexExpr):
            for arg in reversed(node.args):
                if arg.value is not None:
                    pending.append(arg.value)
            pending.append(node.callee)
        elif isinstance(node, MemberAccessExpr):
            if node.object_ is not None:
                pending.append(node.object_)
        elif isinstance(node, TypeOfIsExpr):
            pending.append(node.operand)
        # LiteralExpr / IdentifierExpr / NewExpr / AddressOfExpr: leaves
