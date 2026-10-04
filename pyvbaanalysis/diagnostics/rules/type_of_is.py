"""Rule family: TypeOf ... Is expression rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/typeOfIs.ts. The
TypeOf-missing-operand syntax check is host-free and ported. is-operator-non-object
is a BinaryExpr `Is` visitor: a provably-scalar operand of `Is` is a type error. It
fires on expression-reachable forms (`If x Is Nothing`, `b = x Is Nothing`) and is
silent on `Debug.Print x Is Nothing`, a reserved-name receiver statement parses as
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

from ...completion.member_access import (
    KnownObjectAssignmentType,
    MemberCompletionContext,
    resolve_known_object_assignment_type,
    simple_type_name_for_assignment,
)
from ...conditional import ConditionalActivityTracker
from ...host.host_model import get_host_members, get_host_type, resolve_host_alias
from ...host.type_extensibility import host_type_resolves_when_compiling
from ...lexer.token_kinds import TokenKind
from ...lexer.tokenize import tokenize_cached
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
from ..call_extraction import InferredArgumentType
from ..context import PushFn
from ..exprwalk import ProcedureExpressionVisitor

_TRAILING_ARRAY_RE = re.compile(r"\s*\(\s*\)\s*$")


def check_typeof_missing_operand(
    source: str, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """`TypeOf` requires an object expression before `Is`; `TypeOf Is Y` is a syntax error."""
    toks = [
        tok
        for tok in tokenize_cached(source)
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


def _implements_object_type(
    actual: KnownObjectAssignmentType, expected: KnownObjectAssignmentType
) -> bool:
    """Port of implementsObjectType: True when the actual project class declares
    `Implements <expected>` (honouring excel.-qualified host keys)."""
    expected_names = {expected.key}
    simple = simple_type_name_for_assignment(expected.display)
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


def object_assignment_incompatibility_reason(
    expected_raw: str | None,
    actual: InferredArgumentType | None,
    member_ctx: MemberCompletionContext,
) -> str | None:
    """Port of objectAssignmentIncompatibilityReason: why an object-position target
    typed `expected_raw` cannot accept `actual`, or None when it can (the no-FP gate).

    Reused by the member-assignment Set branch. Returns the human reason string XLIDE
    emits; None whenever the operands are not both provably-incompatible object types
    (Variant/Nothing/generic/implements all stay quiet)."""
    expected = resolve_known_object_assignment_type(expected_raw, member_ctx)
    if expected is None or actual is None:
        return None
    actual_type = normalize_type(actual.type_)
    if not actual_type or actual_type in ("variant", "nothing"):
        return None
    if is_known_scalar_type(actual_type):
        return "An object assignment requires an object value."
    if expected.kind == "generic":
        return None
    actual_object = resolve_known_object_assignment_type(actual.type_, member_ctx)
    if actual_object is None or actual_object.kind == "generic":
        return None
    if expected.key == actual_object.key:
        return None
    if actual_object.kind == "host" and _HOST_VALUES_ALSO_OF_TYPE.get(actual_object.key) == expected.key:
        return None
    if actual_object.kind == "project" and _implements_object_type(actual_object, expected):
        return None
    # A Set between two class types is checked when it runs, by QueryInterface,
    # so it compiles whenever the object could support the target (XLIDE issue
    # #109). The class an interface is implemented by can hold the interface's
    # value (`Set c = o`, casting back), and two interfaces one class implements
    # can hold each other's (`Set b = o`). Only project interfaces are known here.
    if (
        expected.kind == "project"
        and actual_object.kind == "project"
        and _project_types_can_share_instance(expected, actual_object, member_ctx)
    ):
        return None
    return f"This object type is not compatible with {expected.display}."


_HAS_PARAMETERS_RE = re.compile(r"\([^)]")


def object_let_assignment_verdict(expected_raw: str | None, member_ctx: MemberCompletionContext) -> str:
    """Port of objectLetAssignmentVerdict: what a bare `name = value` does to a
    variable of a known object type (XLIDE issue #107, each case measured in Excel
    16.0). The VBE compiles it as a Let through the type's default member, so it is
    never "Set required" at compile time:

    - "lets": the type has a parameterless default member (Range's `_Default` is
      Value; a project class marks one with VB_UserMemId = 0), or is the generic
      Object, whose default member is looked up when it runs. `r = 5` writes A1.
      Nothing to report.
    - "argument": the default member takes an argument, so the VBE refuses the
      statement: `c = 5` on a Collection is "Argument not optional".
    - "noDefault": the type is fully known and has no default member, so the
      statement compiles and raises error 438 when it runs (`ws = 9`).
    - "unknown": the model cannot say. Nothing is reported.
    """
    expected = resolve_known_object_assignment_type(expected_raw, member_ctx)
    if expected is None:
        return "unknown"
    if expected.kind == "generic":
        return "argument" if expected.key == "collection" else "lets"
    if expected.kind == "project":
        project_type = next(
            (
                candidate
                for candidate in member_ctx.project_class_members or []
                if candidate.name.lower() == expected.key
            ),
            None,
        )
        if project_type is None or project_type.exhaustive is not True:
            return "unknown"
        default_member = next((member for member in project_type.members if member.default_member), None)
        if default_member is None:
            return "noDefault"
        takes_argument = bool(default_member.signature) and _HAS_PARAMETERS_RE.search(
            default_member.signature or ""
        ) is not None
        return "argument" if takes_argument else "lets"
    members = get_host_members(expected_raw or "", member_ctx.model)
    host_default = next((member for member in members if member["name"] == "_Default"), None)
    if host_default is not None:
        takes_argument = (
            host_default.get("kind") == "method"
            or _HAS_PARAMETERS_RE.search(host_default.get("signature") or "") is not None
        )
        return "argument" if takes_argument else "lets"
    return "noDefault" if _host_type_is_closed(expected_raw or "", member_ctx) else "unknown"


def _host_type_is_closed(type_name: str, member_ctx: MemberCompletionContext) -> bool:
    """Whether the host model's member list for the type proves a member absent:
    the list is complete AND the type library resolves members while compiling
    (the same two facts member-not-found needs)."""
    alias = resolve_host_alias(type_name, member_ctx.model)
    resolved = alias if alias is not None else type_name
    host_type = get_host_type(resolved, member_ctx.model)
    return (
        host_type is not None
        and host_type.get("exhaustive") is True
        and host_type_resolves_when_compiling(resolved)
    )


def _project_types_can_share_instance(
    expected: KnownObjectAssignmentType,
    actual: KnownObjectAssignmentType,
    member_ctx: MemberCompletionContext,
) -> bool:
    """Whether one project class can carry a value declared as the other: the
    expected class implements the actual type (a cast from an interface back to
    the class), or some project class implements both (a cast between two
    interfaces of one object)."""
    if _implements_object_type(expected, actual):
        return True
    wanted = {expected.key, actual.key}
    for project_type in member_ctx.project_class_members or []:
        implemented = {name.lower() for name in project_type.implements or []}
        if wanted <= implemented:
            return True
    return False


# Host types the model returns where the type library returns another, so a value
# of the model's type is also a value of the library's. Excel's library types the
# Charts and Worksheets properties of Application and Workbook as Sheets, and a
# Sheets object is what they return at run time. The model returns its own Charts
# and Worksheets there, whose members completion offers (XLIDE issue #90).
_HOST_VALUES_ALSO_OF_TYPE: dict[str, str] = {
    "excel.charts": "excel.sheets",
    "excel.worksheets": "excel.sheets",
}


def _is_implemented_by_any_project_class(
    operand_type: KnownObjectAssignmentType, member_ctx: MemberCompletionContext
) -> bool:
    """True when any project class declares `Implements <operandType>`, so the
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
    operand_type = resolve_known_object_assignment_type(declared, member_ctx)
    target_type = resolve_known_object_assignment_type(expr.type_name, member_ctx)
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
    # Mutual incompatibility: neither type is assignable to the other. If either
    # is, `TypeOf` could be True, so stay quiet.
    operand_can_be_target = (
        object_assignment_incompatibility_reason(
            expr.type_name, InferredArgumentType(declared, declared, expr.span), member_ctx
        )
        is None
    )
    target_can_be_operand = (
        object_assignment_incompatibility_reason(
            declared, InferredArgumentType(expr.type_name, expr.type_name, expr.span), member_ctx
        )
        is None
    )
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


# --- sync stubs (2f49b93): replaced as each group is ported ---


def check_type_of_missing_operand(*args: object, **kwargs: object) -> None:
    return None


def check_type_of_is_compatibility(*args: object, **kwargs: object) -> None:
    return None


def check_is_operands_in_conditions(*args: object, **kwargs: object) -> None:
    return None
