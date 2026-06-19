"""The ordered diagnostic rule registry.

Ported from registry.ts. A rule entry is a stable name plus exactly one execution
form (run / procedure_statements / procedure_expressions). The registry ORDER is a
hard contract: it is the diagnostic output order (run_rules buffers per rule and
flushes in registry order). Rule families are appended here as they are ported
(M6+); the engine skeleton starts with an empty registry.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from .context import PushFn, RulePassContext
from .exprwalk import ProcedureExpressionVisitor
from .rules.arrays import (
    check_array_bound_intrinsic_arguments,
    check_array_declaration_bounds,
    check_erase_targets,
    check_fixed_array_subscript_bounds,
    check_invalid_redim_targets,
    check_redim_impossible_bounds,
    check_redim_preserve_dimensions,
    check_unallocated_dynamic_array_access,
)
from .rules.assignments import check_const_assignment, check_mid_statement_literal_target
from .rules.binary_operand_scalar import check_binary_operand_scalar
from .rules.control_flow import (
    check_duplicate_case_else,
    check_duplicate_labels,
    check_else_branch_order,
    check_else_without_if,
    check_exit_statements,
    check_malformed_statements,
    check_statement_context,
    check_undefined_labels,
)
from .rules.declarations import (
    check_dim_initializer,
    check_duplicate_options,
    check_empty_type,
    check_fixed_length_string_bounds,
    check_identifier_too_long,
    check_invalid_identifier_starts,
    check_module_declarations_after_procedures,
    check_module_declarations_in_procedure_bodies,
    check_module_level_statements_outside_procedures,
    check_non_constant_const_values,
    check_non_constant_enum_member_values,
    check_option_placement,
    check_parameter_order,
    check_procedure_header,
    check_property_accessor_signatures,
    check_reserved_declaration_names,
    check_too_many_parameters,
    check_type_declaration_character_as_clause,
    check_udt_parameter_constraints,
    check_unexpected_declaration_tokens,
)
from .rules.expressions import check_unbalanced_parens
from .rules.duplicates import (
    check_duplicate_declarations,
    check_duplicate_enum_members,
    check_duplicate_module_members,
    check_duplicate_procedures,
    check_duplicate_type_fields,
)
from .rules.lexical import check_invalid_line_continuations, check_unterminated_strings
from .rules.module_kind import (
    check_declare_ptr_safe_for_win64,
    check_event_declaration_module_kind,
    check_friend_declarations,
    check_implements_statement_placement,
    check_me_outside_object_module,
    check_object_module_public_members,
    check_raise_event_targets,
    check_with_events_declarations,
)
from .rules.numeric_literals import check_suffixed_literal_overflow
from .rules.object_state import check_object_variable_not_set
from .rules.type_of_is import check_typeof_missing_operand
from .walker import ProcedureStatementVisitor


@dataclass(frozen=True, slots=True)
class DiagnosticRuleEntry:
    """One registered rule: a name and exactly one of the three execution forms."""

    name: str
    run: Callable[[RulePassContext, PushFn], None] | None = None
    procedure_statements: Callable[[RulePassContext, PushFn], ProcedureStatementVisitor] | None = None
    procedure_expressions: Callable[[RulePassContext, PushFn], ProcedureExpressionVisitor] | None = None


# The ordered table of active rules. The ORDER is the diagnostic output-order
# contract; entries are placed at their registry.ts positions as families are
# ported (gaps remain for not-yet-ported families).
DIAGNOSTIC_RULE_REGISTRY: tuple[DiagnosticRuleEntry, ...] = (
    DiagnosticRuleEntry(
        name="unterminatedStrings",
        run=lambda ctx, push: check_unterminated_strings(ctx.source, push),
    ),
    DiagnosticRuleEntry(
        name="invalidLineContinuations",
        run=lambda ctx, push: check_invalid_line_continuations(ctx.source, push),
    ),
    DiagnosticRuleEntry(
        name="duplicateProcedures",
        run=lambda ctx, push: check_duplicate_procedures(ctx.symbols.root.children or [], push),
    ),
    DiagnosticRuleEntry(
        name="duplicateDeclarations",
        run=lambda ctx, push: check_duplicate_declarations(ctx.symbols.root.children or [], push),
    ),
    DiagnosticRuleEntry(
        name="duplicateModuleMembers",
        run=lambda ctx, push: check_duplicate_module_members(ctx.symbols.root.children or [], push),
    ),
    DiagnosticRuleEntry(
        name="duplicateEnumMembers",
        run=lambda ctx, push: check_duplicate_enum_members(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="duplicateTypeFields",
        run=lambda ctx, push: check_duplicate_type_fields(ctx.source, ctx.mod, ctx.activity, push),
    ),
    # -- declarations family (positions 8-11) --
    DiagnosticRuleEntry(name="emptyType", run=lambda ctx, push: check_empty_type(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="tooManyParameters", run=lambda ctx, push: check_too_many_parameters(ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="identifierTooLong", run=lambda ctx, push: check_identifier_too_long(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="udtParameterConstraints", run=lambda ctx, push: check_udt_parameter_constraints(ctx.mod, ctx.activity, push)),
    # Position 12 deferred: ambiguousEnumMemberReferences (M9, needs project + host surfaces).
    DiagnosticRuleEntry(name="constAssignment", procedure_statements=lambda ctx, push: check_const_assignment(ctx.source, ctx.symbols, ctx.opts.project_visible_symbols, push)),
    # Positions 14-15 deferred: optionExplicit + undeclaredVariables (undeclared family, M9).
    DiagnosticRuleEntry(name="optionPlacement", run=lambda ctx, push: check_option_placement(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="duplicateOption", run=lambda ctx, push: check_duplicate_options(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="procedureHeader", run=lambda ctx, push: check_procedure_header(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="invalidIdentifierStarts", run=lambda ctx, push: check_invalid_identifier_starts(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="moduleDeclarationsInProcedureBodies", run=lambda ctx, push: check_module_declarations_in_procedure_bodies(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="moduleDeclarationsAfterProcedures", run=lambda ctx, push: check_module_declarations_after_procedures(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="moduleLevelStatementsOutsideProcedures", run=lambda ctx, push: check_module_level_statements_outside_procedures(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="reservedDeclarationNames", run=lambda ctx, push: check_reserved_declaration_names(ctx.source, ctx.mod, ctx.activity, push)),
    # Position 24 deferred: propertySetterValueParameters (object-value branch needs
    # resolveKnownObjectAssignmentType / host, M9).
    DiagnosticRuleEntry(name="propertyAccessorSignatures", run=lambda ctx, push: check_property_accessor_signatures(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="parameterOrder", run=lambda ctx, push: check_parameter_order(ctx.source, ctx.mod, ctx.activity, push)),
    # Positions 27-28 deferred: parameter defaults (memberCtx + inferArgumentType, M8).
    DiagnosticRuleEntry(name="constValueNotConstant", run=lambda ctx, push: check_non_constant_const_values(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="enumMemberNotConstant", run=lambda ctx, push: check_non_constant_enum_member_values(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="unbalancedParens", run=lambda ctx, push: check_unbalanced_parens(ctx.source, push)),
    # Positions 32-33 deferred: invalid-expression-syntax + division-by-zero
    # (call extraction / type inference, M8).
    DiagnosticRuleEntry(name="dimInitializer", run=lambda ctx, push: check_dim_initializer(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="invalidRedimTargets", procedure_statements=lambda ctx, push: check_invalid_redim_targets(ctx.source, ctx.mod, ctx.symbols, ctx.opts.project_visible_symbols, ctx.activity, push)),
    DiagnosticRuleEntry(name="redimImpossibleBounds", procedure_statements=lambda ctx, push: check_redim_impossible_bounds(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="arrayDeclarationImpossibleBounds", run=lambda ctx, push: check_array_declaration_bounds(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="redimPreserveDimensions", run=lambda ctx, push: check_redim_preserve_dimensions(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="unallocatedDynamicArrayAccess", run=lambda ctx, push: check_unallocated_dynamic_array_access(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="arraySubscriptOutOfBounds", run=lambda ctx, push: check_fixed_array_subscript_bounds(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="midStatementLiteralTarget", run=lambda ctx, push: check_mid_statement_literal_target(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push)),
    DiagnosticRuleEntry(name="eraseTargets", procedure_statements=lambda ctx, push: check_erase_targets(ctx.source, ctx.symbols, ctx.opts.project_visible_symbols, push)),
    DiagnosticRuleEntry(name="typeDeclarationCharacterAsClause", run=lambda ctx, push: check_type_declaration_character_as_clause(ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="unexpectedDeclarationTokens", run=lambda ctx, push: check_unexpected_declaration_tokens(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="fixedLengthStringBounds", run=lambda ctx, push: check_fixed_length_string_bounds(ctx.source, ctx.mod, ctx.activity, push)),
    # Positions 35-42 (arrays / object-state, M7) precede typeDeclarationCharacterAsClause
    # above and remain deferred.
    # -- moduleKind family (positions 46-53, self-contained subset) --
    DiagnosticRuleEntry(name="objectModulePublicMembers", run=lambda ctx, push: check_object_module_public_members(ctx.source, ctx.mod, ctx.module_kind, ctx.activity, push)),
    DiagnosticRuleEntry(name="eventDeclarationModuleKind", run=lambda ctx, push: check_event_declaration_module_kind(ctx.source, ctx.mod, ctx.module_kind, ctx.activity, push)),
    DiagnosticRuleEntry(name="meOutsideObjectModule", procedure_statements=lambda ctx, push: check_me_outside_object_module(ctx.module_kind, ctx.source, push)),
    DiagnosticRuleEntry(name="withEventsDeclarations", run=lambda ctx, push: check_with_events_declarations(ctx.source, ctx.mod, ctx.module_kind, ctx.activity, push)),
    DiagnosticRuleEntry(name="friendDeclarations", run=lambda ctx, push: check_friend_declarations(ctx.source, ctx.mod, ctx.module_kind, ctx.activity, push)),
    DiagnosticRuleEntry(name="implementsStatementPlacement", run=lambda ctx, push: check_implements_statement_placement(ctx.source, ctx.mod, ctx.module_kind, ctx.activity, push)),
    DiagnosticRuleEntry(name="raiseEventTargets", run=lambda ctx, push: check_raise_event_targets(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="declarePtrSafeForWin64", run=lambda ctx, push: check_declare_ptr_safe_for_win64(ctx.source, ctx.mod, ctx.opts.conditional_compilation, ctx.activity, push)),
    # Positions 54-58 deferred: eventHandlerModuleScope (completion), invalidAsTypeNames
    # (M9), and the parenthesized-call rules (memberCtx / call extraction, M8).
    # -- control-flow family (positions 59-66, self-contained subset) --
    DiagnosticRuleEntry(name="exitStatements", procedure_statements=lambda ctx, push: check_exit_statements(ctx.source, push)),
    DiagnosticRuleEntry(name="duplicateLabels", run=lambda ctx, push: check_duplicate_labels(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="undefinedLabels", run=lambda ctx, push: check_undefined_labels(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="elseBranchOrder", run=lambda ctx, push: check_else_branch_order(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="statementContext", run=lambda ctx, push: check_statement_context(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="duplicateCaseElse", run=lambda ctx, push: check_duplicate_case_else(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="malformedStatements", run=lambda ctx, push: check_malformed_statements(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(name="elseWithoutIf", run=lambda ctx, push: check_else_without_if(ctx.source, ctx.mod, ctx.activity, push)),
    # forEachLoopTypes (67, type inference + host) and scalarMemberAccess (69, type
    # env + member surface) are deferred to M8.
    DiagnosticRuleEntry(name="arrayBoundIntrinsicArguments", procedure_statements=lambda ctx, push: check_array_bound_intrinsic_arguments(ctx.source, ctx.symbols, ctx.opts.project_visible_symbols, push)),
    DiagnosticRuleEntry(name="objectVariableNotSet", run=lambda ctx, push: check_object_variable_not_set(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push)),
    # Positions 71-78 deferred: memberNotFound (M9), the argument/runtime/assignment
    # type rules (M8 foundation), and typeOfIsAlwaysFalse (M9 host).
    DiagnosticRuleEntry(name="typeofMissingOperand", run=lambda ctx, push: check_typeof_missing_operand(ctx.source, ctx.activity, push)),
    # Position 80 deferred: isOperatorNonObject (type_environment).
    DiagnosticRuleEntry(name="nonScalarBinaryOperand", procedure_expressions=lambda ctx, push: check_binary_operand_scalar(ctx.symbols, push)),
    # Position 82 deferred: argumentShapeMismatch (call extraction + inference).
    DiagnosticRuleEntry(name="suffixedLiteralOverflow", run=lambda ctx, push: check_suffixed_literal_overflow(ctx.source, ctx.activity, push)),
)
