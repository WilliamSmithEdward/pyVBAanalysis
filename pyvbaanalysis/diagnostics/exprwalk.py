"""Shared expression-tree traversal for diagnostics (MS-VBAL 5.6).

Ported from exprWalk.ts. Owns the one canonical two-level walk over a procedure
body that expression-consuming rules share: find every root expression (assignment
sides, call callee + arguments, If/ElseIf conditions, nested block bodies) and
recurse each into its sub-expressions in pre-order.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence

from ..conditional import ConditionalActivityTracker
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
)
from .walker import active_module_members, is_inactive_node

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
) -> None:
    """Run ONE shared expression walk per active procedure, dispatching to each visitor."""
    if len(factories) == 0:
        return
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        visitors = [factory(member) for factory in factories]
        for_each_expression_in_body(member.body, activity, _fan_out_expressions(visitors))


def for_each_expression_in_body(
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    visit: Callable[[ExprNode], None],
) -> None:
    """Visit every expression node reachable in a body, skipping inactive regions."""
    for node in body:
        if is_inactive_node(activity, node):
            continue
        if isinstance(node, AssignmentNode):
            for_each_sub_expression(node.lhs, visit)
            for_each_sub_expression(node.rhs, visit)
        elif isinstance(node, CallNode):
            for_each_sub_expression(node.callee, visit)
            for arg in node.args:
                if arg.value is not None:
                    for_each_sub_expression(arg.value, visit)
        elif isinstance(node, IfBlockNode):
            for branch in node.branches:
                if branch.condition is not None:
                    for_each_sub_expression(branch.condition, visit)
            # Arm statements live in the flat body; recurse it for nested exprs.
            for_each_expression_in_body(node.body, activity, visit)
        else:
            child = getattr(node, "body", None)
            if isinstance(child, list):
                for_each_expression_in_body(child, activity, visit)


def for_each_sub_expression(expr: ExprNode, visit: Callable[[ExprNode], None]) -> None:
    """Visit expr and every nested sub-expression (pre-order)."""
    visit(expr)
    if isinstance(expr, BinaryExpr):
        for_each_sub_expression(expr.left, visit)
        for_each_sub_expression(expr.right, visit)
    elif isinstance(expr, UnaryExpr):
        for_each_sub_expression(expr.operand, visit)
    elif isinstance(expr, ParenExpr):
        for_each_sub_expression(expr.inner, visit)
    elif isinstance(expr, IndexExpr):
        for_each_sub_expression(expr.callee, visit)
        for arg in expr.args:
            if arg.value is not None:
                for_each_sub_expression(arg.value, visit)
    elif isinstance(expr, MemberAccessExpr):
        if expr.object_ is not None:
            for_each_sub_expression(expr.object_, visit)
    elif isinstance(expr, TypeOfIsExpr):
        for_each_sub_expression(expr.operand, visit)
    # LiteralExpr / IdentifierExpr / NewExpr / AddressOfExpr: leaves
