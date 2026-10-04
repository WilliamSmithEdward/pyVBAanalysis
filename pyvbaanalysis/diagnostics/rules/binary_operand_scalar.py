"""Rule family: non-scalar operand of a scalar-requiring binary operator.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/binaryOperandScalar.ts.
A bare array variable or a same-module user-defined Type value used as an operand
of a scalar-requiring binary operator is a VBE compile error ("Type mismatch").
Fires only on a plain identifier provably non-scalar (its shape is an array, or
its declared type names a Type in this module); everything else stays quiet.
A Nothing literal operand of any operator but Is (binary or unary) is "Invalid
use of object".
"""

from __future__ import annotations

from collections.abc import Callable, Mapping
from collections.abc import Set as AbstractSet

from ...parser.nodes import (
    BinaryExpr,
    ExprNode,
    IdentifierExpr,
    LiteralExpr,
    LiteralKind,
    ParenExpr,
    ProcedureNode,
    Span,
    UnaryExpr,
)
from ...symbols.symbol_model import ModuleSymbols
from ...types.type_inference import (
    DeclaredValueShape,
    declaration_shape_environment_for,
    same_module_type_names,
)
from ..context import PushFn
from ..exprwalk import ProcedureExpressionVisitor

# Binary operators that require scalar operands. Excludes Is (object operands,
# owned by is-operator-non-object) and Like (operands coerced to String).
_SCALAR_OPERAND_OPERATORS = frozenset(
    {
        "&", "+", "-", "*", "/", "\\", "^", "Mod",
        "=", "<>", "<", ">", "<=", ">=",
        "And", "Or", "Xor", "Eqv", "Imp",
    }
)


def check_binary_operand_scalar(symbols: ModuleSymbols, push: PushFn) -> ProcedureExpressionVisitor:
    udt_names = same_module_type_names(symbols)

    def factory(member: ProcedureNode) -> Callable[[ExprNode], None]:
        shapes = declaration_shape_environment_for(symbols, member)

        def visitor(expr: ExprNode) -> None:
            # Nothing is no value for any operator but Is: `Nothing + 1`,
            # `"a" & Nothing`, `-Nothing`, `Not Nothing` (issue #325, measured
            # in Excel 16.0).
            operation: tuple[str, ExprNode | None] | None = None
            if isinstance(expr, BinaryExpr) and expr.operator != "Is":
                found = next((side for side in (expr.left, expr.right) if _is_nothing(side)), None)
                operation = (expr.operator, found)
            elif isinstance(expr, UnaryExpr):
                operation = (expr.operator, expr.operand if _is_nothing(expr.operand) else None)
            if operation is not None and operation[1] is not None:
                push(
                    "nonScalarBinaryOperand",
                    f"The '{operation[0]}' operator needs a value, and Nothing is an object reference. "
                    "This is a VBE compile error: Invalid use of object.",
                    operation[1].span,
                )
                return
            if not isinstance(expr, BinaryExpr) or expr.operator not in _SCALAR_OPERAND_OPERATORS:
                return
            offender = _non_scalar_operand(expr.left, shapes, udt_names) or _non_scalar_operand(
                expr.right, shapes, udt_names
            )
            if offender is not None:
                span, detail = offender
                push(
                    "nonScalarBinaryOperand",
                    f"The '{expr.operator}' operator requires a scalar operand, but {detail}. "
                    "This will fail to compile with 'Type mismatch'.",
                    span,
                )

        return visitor

    return factory


def _is_nothing(expr: ExprNode) -> bool:
    inner = expr.inner if isinstance(expr, ParenExpr) else expr
    return isinstance(inner, LiteralExpr) and inner.literal_kind == LiteralKind.NOTHING


def _non_scalar_operand(
    expr: ExprNode, shapes: Mapping[str, DeclaredValueShape], udt_names: AbstractSet[str]
) -> tuple[Span, str] | None:
    # Only a bare identifier is typed: an indexed element a(i) is an IndexExpr
    # (a scalar element), member access / calls / parens are not the aggregate,
    # and literals are scalar.
    if not isinstance(expr, IdentifierExpr):
        return None
    shape = shapes.get(expr.name.lower())
    if shape is not None and shape.is_array:
        return (expr.span, f"'{expr.name}' is declared As an array")
    declared = shape.as_type if shape is not None else None
    if declared and declared.lower() in udt_names:
        return (expr.span, f"'{expr.name}' is declared As {declared} (a user-defined Type)")
    return None
