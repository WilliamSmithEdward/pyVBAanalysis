"""Rule family: TypeOf ... Is expression rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/typeOfIs.ts. The
TypeOf-missing-operand syntax check is host-free and ported. is-operator-non-object
is a BinaryExpr `Is` visitor: a provably-scalar operand of `Is` is a type error. It
fires on expression-reachable forms (`If x Is Nothing`, `b = x Is Nothing`) and is
silent on `Debug.Print x Is Nothing` — a reserved-name receiver statement parses as
a raw StatementNode, so the inner `Is` never reaches the expression walk. This
matches XLIDE byte-for-byte (verified by running XLIDE's analyzer: it too is silent
on the Debug.Print form and fires on the If form); the entire is-operator-non-object
oracle corpus happens to use the dormant Debug.Print form, so those cases are
documented as XLIDE-dormant in the test rather than satisfied.

typeof-is-always-false ports the SAFE v1: a simple-identifier operand whose declared
type resolves to a CONCRETE object class that is mutually incompatible with the
target object class makes `TypeOf x Is T` provably False. Object/Variant/scalar/
unknown/interface operands and any pair that could be assignment-compatible stay
quiet (no-FP). The object-compatibility check is a faithful port of
resolveKnownObjectAssignmentType + objectAssignmentIncompatibilityReason
(typeInference.ts), reusing the host alias resolver and project class members.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass

from ...completion.member_access import MemberCompletionContext
from ...conditional import ConditionalActivityTracker
from ...host import resolve_host_alias
from ...lexer.token_kinds import TokenKind
from ...lexer.tokenize import tokenize
from ...parser.nodes import (
    BinaryExpr,
    ExprNode,
    IdentifierExpr,
    LiteralExpr,
    LiteralKind,
    ProcedureNode,
    Span,
    TypeOfIsExpr,
)
from ...symbols.symbol_model import ModuleSymbols
from ...types.type_inference import type_environment_for
from ...types.type_names import is_known_scalar_type, normalize_type
from ..context import PushFn
from ..exprwalk import ProcedureExpressionVisitor

_SIMPLE_TYPE_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_TRAILING_ARRAY_RE = re.compile(r"\s*\(\s*\)\s*$")


def check_typeof_missing_operand(
    source: str, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """`TypeOf` requires an object expression before `Is`; `TypeOf Is Y` is a syntax error."""
    toks = [
        tok
        for tok in tokenize(source)
        if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE
    ]
    for i in range(len(toks) - 1):
        if (toks[i].canonical_text or toks[i].raw_text).lower() != "typeof":
            continue
        if (toks[i + 1].canonical_text or toks[i + 1].raw_text).lower() != "is":
            continue
        span = Span(toks[i].start, toks[i + 1].end)
        if activity is not None and activity.is_inactive(span):
            continue
        push("typeofMissingOperand", "'TypeOf' requires an object expression before 'Is'.", span)


# -- checkIsOperatorOperands (is-operator-non-object) -----------------------

_IS_NON_OBJECT_LITERAL_KINDS: dict[LiteralKind, str] = {
    LiteralKind.INTEGER: "integer",
    LiteralKind.FLOAT: "float",
    LiteralKind.STRING: "string",
    LiteralKind.DATE: "date",
    LiteralKind.BOOLEAN: "boolean",
}


def _non_object_operand(expr: ExprNode, env: dict[str, str]) -> tuple[Span, str] | None:
    """A provably non-object (scalar) operand of `Is`, or None."""
    if isinstance(expr, LiteralExpr):
        kind = _IS_NON_OBJECT_LITERAL_KINDS.get(expr.literal_kind)
        if kind is not None:
            return (expr.span, f"'{expr.raw}' is a {kind} literal")
        return None  # Nothing / Null / Empty -> not provably scalar
    if isinstance(expr, IdentifierExpr):
        declared = env.get(expr.name.lower())
        if not declared:
            return None  # undeclared / unknown -> quiet
        # Strip only a trailing array `()` marker -- NOT normalize_type's leading-`vb`
        # strip, which would wrongly collapse a user class named `vbLong` to a scalar.
        raw = _TRAILING_ARRAY_RE.sub("", declared).strip().lower()
        if is_known_scalar_type(raw):
            return (expr.span, f"'{expr.name}' is declared As {declared}")
        return None  # Variant / Object / class -> quiet
    return None  # member / call / paren / array / New / unary -> quiet (v1)


def check_is_operator_operands(symbols: ModuleSymbols, push: PushFn) -> ProcedureExpressionVisitor:
    """The `Is` operator requires object operands; a provably-scalar operand is an error."""

    def factory(member: ProcedureNode) -> Callable[[ExprNode], None]:
        env = type_environment_for(symbols, member)

        def visitor(expr: ExprNode) -> None:
            if not isinstance(expr, BinaryExpr) or expr.operator != "Is":
                return
            offender = _non_object_operand(expr.left, env) or _non_object_operand(expr.right, env)
            if offender is not None:
                span, detail = offender
                push(
                    "isOperatorNonObject",
                    f"The 'Is' operator requires object operands, but {detail}, "
                    "which is not an object.",
                    span,
                )

        return visitor

    return factory


# -- checkTypeOfIsCompatibility (always-False) ------------------------------


@dataclass(frozen=True, slots=True)
class _ObjectAssignmentType:
    kind: str  # 'generic' | 'host' | 'project'
    display: str
    key: str
    implements: tuple[str, ...] = ()


def _simple_type_name_for_assignment(type_text: str) -> str | None:
    trimmed = _TRAILING_ARRAY_RE.sub("", type_text).strip()
    return trimmed if _SIMPLE_TYPE_NAME_RE.match(trimmed) else None


def _resolve_known_object_assignment_type(
    type_text: str | None, member_ctx: MemberCompletionContext
) -> _ObjectAssignmentType | None:
    """Port of resolveKnownObjectAssignmentType: the object class a declared type
    names, classified as generic Object, a host type, or an unambiguous project
    class/document/userform. Returns None for Variant/scalar/unknown."""
    if not type_text:
        return None
    normalized = normalize_type(type_text)
    if not normalized or normalized == "variant":
        return None
    if normalized == "object":
        return _ObjectAssignmentType(kind="generic", display=type_text, key="object")
    if is_known_scalar_type(normalized):
        return None
    host = resolve_host_alias(type_text, member_ctx.model)
    if host:
        return _ObjectAssignmentType(kind="host", display=type_text, key=host.lower())
    simple = _simple_type_name_for_assignment(type_text)
    if not simple:
        return None
    lower = simple.lower()
    matches = [
        project_type
        for project_type in (member_ctx.project_class_members or [])
        if project_type.kind not in ("userType", "standardModule")
        and project_type.name.lower() == lower
    ]
    if len(matches) != 1:
        return None
    return _ObjectAssignmentType(
        kind="project",
        display=matches[0].name,
        key=lower,
        implements=tuple(matches[0].implements or []),
    )


def _implements_object_type(
    actual: _ObjectAssignmentType, expected: _ObjectAssignmentType
) -> bool:
    """Port of implementsObjectType: True when the actual project class declares
    `Implements <expected>` (honouring excel.-qualified host keys)."""
    expected_names = {expected.key}
    simple = _simple_type_name_for_assignment(expected.display)
    if simple:
        expected_names.add(simple.lower())
    last_segment = expected.key.split(".")[-1]
    if last_segment:
        expected_names.add(last_segment)
    for implemented in actual.implements:
        lower = implemented.lower()
        if lower in expected_names or f"excel.{lower}" in expected_names:
            return True
    return False


def _object_assignment_incompatible(
    expected_raw: str, actual_raw: str, member_ctx: MemberCompletionContext
) -> bool:
    """True when an `actual_raw`-typed object value is provably NOT assignable to an
    `expected_raw`-typed object reference. Port of the object-vs-object arm of
    objectAssignmentIncompatibilityReason (scalar/Variant/Nothing arms are not
    reachable here — both operands are already concrete object types)."""
    expected = _resolve_known_object_assignment_type(expected_raw, member_ctx)
    if expected is None or expected.kind == "generic":
        return False
    actual = _resolve_known_object_assignment_type(actual_raw, member_ctx)
    if actual is None or actual.kind == "generic":
        return False
    if expected.key == actual.key:
        return False
    if actual.kind == "project" and _implements_object_type(actual, expected):
        return False
    return True


def _is_implemented_by_any_project_class(
    operand_type: _ObjectAssignmentType, member_ctx: MemberCompletionContext
) -> bool:
    """True when any project class declares `Implements <operandType>` — so the
    operand could hold a subtype that is-a the target (stay quiet, no-FP)."""
    names = {operand_type.key, operand_type.display.lower()}
    for project_type in member_ctx.project_class_members or []:
        for implemented in project_type.implements or []:
            if implemented.lower() in names:
                return True
    return False


def _check_typeof_is(
    expr: TypeOfIsExpr, env: dict[str, str], member_ctx: MemberCompletionContext, push: PushFn
) -> None:
    if not isinstance(expr.operand, IdentifierExpr):
        return  # v1: only simple identifier operands have a known declared type
    operand_name = expr.operand.name
    declared = env.get(operand_name.lower())
    if not declared:
        return  # undeclared / unknown type -> quiet
    operand_type = _resolve_known_object_assignment_type(declared, member_ctx)
    target_type = _resolve_known_object_assignment_type(expr.type_name, member_ctx)
    if operand_type is None or target_type is None:
        return  # not both known object types -> quiet
    if operand_type.kind == "generic" or target_type.kind == "generic":
        return  # Object operand or `Is Object` -> quiet
    if operand_type.key == target_type.key:
        return  # same type (always True) -> quiet
    # Concrete-operand gate: an interface-typed operand could hold a subtype that
    # is-a the target, so only fire when no project class implements the operand.
    if operand_type.kind == "project" and _is_implemented_by_any_project_class(operand_type, member_ctx):
        return
    operand_can_be_target = not _object_assignment_incompatible(expr.type_name, declared, member_ctx)
    target_can_be_operand = not _object_assignment_incompatible(declared, expr.type_name, member_ctx)
    if operand_can_be_target or target_can_be_operand:
        return
    push(
        "typeOfIsAlwaysFalse",
        f"'TypeOf ... Is {target_type.display}' is always False: '{operand_name}' is declared "
        f"As {operand_type.display}, which is never {target_type.display}.",
        expr.span,
    )


def check_typeof_is_compatibility(
    symbols: ModuleSymbols, member_ctx: MemberCompletionContext, push: PushFn
) -> ProcedureExpressionVisitor:
    def factory(member: ProcedureNode) -> Callable[[ExprNode], None]:
        env = type_environment_for(symbols, member)

        def visitor(expr: ExprNode) -> None:
            if isinstance(expr, TypeOfIsExpr):
                _check_typeof_is(expr, env, member_ctx, push)

        return visitor

    return factory
