"""The ordered diagnostic rule registry.

Ported from registry.ts. A rule entry is a stable name plus exactly one execution
form (run / procedure_statements / procedure_expressions). The registry ORDER is a
hard contract: it is the diagnostic output order (run_rules buffers per rule and
flushes in registry order), so entries sit at their registry.ts positions, and each
entry adapts the shared pass context to its rule function's signature as
registry.ts does. A few entries keep a port-only argument arrangement their rule
function has always had (noted at the entry).
"""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass

from ..host.host_model import get_excel_object_model
from ..parser.nodes import Span
from ..symbols.symbol_model import ModuleSymbolKind, VbaSymbolKind
from .callable_signatures import callable_type_signatures_for
from .context import PushFn, RulePassContext, own_object_member_names
from .exprwalk import ProcedureExpressionVisitor
from .model import VbaDiagnosticData
from .rules.access_data import check_access_data
from .rules.address_of_use import check_address_of_use
from .rules.argument_shape import check_argument_shape
from .rules.argument_types import check_argument_types
from .rules.arrays import (
    check_array_bound_intrinsic_arguments,
    check_array_declaration_bounds,
    check_erase_targets,
    check_fixed_array_subscript_bounds,
    check_invalid_redim_targets,
    check_redim_impossible_bounds,
    check_redim_preserve_dimensions,
    check_redim_type_change,
    check_unallocated_dynamic_array_access,
)
from .rules.assignments import (
    check_assignment_types,
    check_const_assignment,
    check_mid_statement_literal_target,
    check_missing_return_assignments,
    check_set_assignments,
)
from .rules.binary_operand_scalar import check_binary_operand_scalar
from .rules.by_name_calls import check_by_name_calls
from .rules.call_arity import check_argument_count
from .rules.class_instance_values import check_class_instance_values
from .rules.collection_state import check_collection_loop_counters, check_collection_state
from .rules.condition_values import check_condition_values
from .rules.control_flow import (
    check_duplicate_case_else,
    check_duplicate_labels,
    check_else_branch_order,
    check_else_without_if,
    check_exit_statements,
    check_for_each_loop_types,
    check_line_number_range,
    check_malformed_statements,
    check_reserved_labels,
    check_statement_context,
    check_undefined_labels,
)
from .rules.dead_code import (
    check_unreachable_code,
    check_unused_declarations,
    check_unused_private_procedures,
)
from .rules.declaration_forms import check_declaration_forms
from .rules.declaration_order import check_declaration_order
from .rules.declarations import (
    check_dim_initializer,
    check_duplicate_options,
    check_empty_type,
    check_fixed_length_string_bounds,
    check_identifier_too_long,
    check_invalid_as_type_names,
    check_invalid_identifier_starts,
    check_module_declarations_after_procedures,
    check_module_declarations_in_procedure_bodies,
    check_module_level_statements_outside_procedures,
    check_module_name,
    check_non_constant_const_values,
    check_non_constant_enum_member_values,
    check_non_constant_parameter_defaults,
    check_option_placement,
    check_option_statement_form,
    check_parameter_order,
    check_procedure_header,
    check_property_accessor_signatures,
    check_property_setter_value_parameters,
    check_reserved_declaration_names,
    check_too_many_parameters,
    check_type_declaration_character_as_clause,
    check_udt_parameter_constraints,
    check_unexpected_declaration_tokens,
)
from .rules.declares import check_declare_statements, check_unusable_declare_calls
from .rules.deleted_objects import check_deleted_objects
from .rules.deleted_settings import check_deleted_settings
from .rules.dictionary_state import check_dictionary_state
from .rules.directive_forms import check_directive_forms
from .rules.doc_comments import check_doc_comments
from .rules.document_names import check_document_names
from .rules.duplicates import (
    check_ambiguous_bare_procedure_calls,
    check_ambiguous_enum_member_references,
    check_duplicate_declarations,
    check_duplicate_enum_members,
    check_duplicate_module_members,
    check_duplicate_procedures,
    check_duplicate_type_fields,
    check_enum_member_name_clash,
    check_variable_procedure_name_clash,
)
from .rules.error_values import check_error_values
from .rules.event_handler_signatures import check_event_handler_signatures
from .rules.excel_session_state import check_excel_session_state
from .rules.expressions import (
    check_call_parens,
    check_division_by_zero_expressions,
    check_expression_call_parens,
    check_invalid_expression_syntax,
    check_string_arithmetic_operands,
    check_unbalanced_parens,
)
from .rules.file_paths import check_empty_file_paths
from .rules.file_statements import check_file_statements
from .rules.form_contents import check_form_contents
from .rules.handler_flow import check_handler_flow
from .rules.host_arguments import check_host_arguments, workbook_sheets_to_check
from .rules.implements_members import check_implements_members
from .rules.late_binding import check_late_bound_friend_member
from .rules.late_bound_members import check_runtime_member_not_found
from .rules.late_bound_objects import check_late_bound_objects
from .rules.lexical import check_invalid_line_continuations, check_unterminated_strings
from .rules.line_continuations import check_line_continuation_limits
from .rules.local_declaration_order import check_local_declaration_order
from .rules.locked_arrays import check_locked_arrays
from .rules.long_long_narrowing import check_long_long_narrowing
from .rules.malformed_lines import check_malformed_lines
from .rules.missing_reference import check_missing_library_reference, check_missing_scripting_reference
from .rules.module_kind import (
    check_declare_ptr_safe_for_win64,
    check_event_declaration_module_kind,
    check_event_handler_module_scope,
    check_friend_declarations,
    check_implements_statement_placement,
    check_me_outside_object_module,
    check_object_module_public_members,
    check_raise_event_arguments,
    check_raise_event_targets,
    check_with_events_declarations,
)
from .rules.module_members import check_module_member_forms
from .rules.numeric_literals import check_suffixed_literal_overflow
from .rules.object_state import check_object_variable_not_set, check_scalar_member_access
from .rules.object_values import check_object_default_values
from .rules.omitted_arguments import check_omitted_argument_reads
from .rules.overflow import check_overflow
from .rules.param_array_use import check_param_array_use
from .rules.parameter_defaults import check_parameter_default_values
from .rules.parentheses import check_parentheses
from .rules.property_use import check_invalid_property_use
from .rules.refused_declarations import check_refused_declarations
from .rules.runtime_values import check_runtime_argument_values, check_runtime_conversion_values
from .rules.statement_forms import check_statement_forms
from .rules.statement_types import check_statement_types
from .rules.stray_tokens import check_stray_characters
from .rules.type_field_arrays import check_type_field_arrays
from .rules.type_members import check_type_members
from .rules.type_of_is import (
    check_is_operands_in_conditions,
    check_is_operator_operands,
    check_type_of_is_compatibility,
    check_type_of_missing_operand,
)
from .rules.undeclared import (
    check_builtins_read_bare,
    check_member_not_found,
    check_non_callable_call_statement,
    check_option_explicit,
    check_undeclared_variables,
    check_unknown_call_statement,
)
from .rules.variant_values import check_variant_value_misuse
from .rules.vba_library_members import check_vba_library_members
from .walker import ProcedureStatementVisitor


@dataclass(frozen=True, slots=True)
class DiagnosticRuleEntry:
    """One registered rule: a name and exactly one of the three execution forms.

    `block_headers`: the statement visitor also takes each block's header line as
    a statement, a For's bounds, a Select Case subject, a Do, Loop or While
    condition, a With subject (XLIDE issue #233). For rules that judge an
    expression wherever it stands, and read no statement form.
    """

    name: str
    run: Callable[[RulePassContext, PushFn], None] | None = None
    procedure_statements: Callable[[RulePassContext, PushFn], ProcedureStatementVisitor] | None = None
    block_headers: bool = False
    procedure_expressions: Callable[[RulePassContext, PushFn], ProcedureExpressionVisitor] | None = None


def _active_only(ctx: RulePassContext, push: PushFn) -> PushFn:
    """`push` for a rule that reads the raw text rather than the parse: the VBE does
    not lex an inactive `#If` branch, so an unclosed string or a continuation into
    a blank line there compiles (XLIDE issue #234, measured in Excel 16.0)."""
    activity = ctx.activity
    if activity is None:
        return push

    def active_push(
        rule: str, message: str, span: Span, data: VbaDiagnosticData | None = None
    ) -> None:
        if not activity.is_inactive(span):
            push(rule, message, span, data)

    return active_push


def _host_name(ctx: RulePassContext) -> str | None:
    """`ctx.opts.hostModel?.hostName`."""
    model = ctx.opts.host_model
    return model.get("hostName") if model is not None else None


_CALLABLE_KINDS = frozenset(
    {
        VbaSymbolKind.SUB,
        VbaSymbolKind.FUNCTION,
        VbaSymbolKind.DECLARE,
        VbaSymbolKind.PROPERTY_GET,
        VbaSymbolKind.PROPERTY_LET,
        VbaSymbolKind.PROPERTY_SET,
    }
)


def _callable_names(ctx: RulePassContext) -> set[str]:
    """The module's own procedures and the project's, lowercased (the set
    registry.ts builds for documentNames and excelSessionState)."""
    names = {
        symbol.name.lower() for symbol in ctx.symbols.root.children or [] if symbol.kind in _CALLABLE_KINDS
    }
    names.update(name.lower() for name in (ctx.opts.project_procedures or {}))
    return names


_OPTION_EXPLICIT_RE = re.compile(r"^[ \t]*Option[ \t]+Explicit\b", re.IGNORECASE | re.MULTILINE | re.ASCII)


def _late_bound_friend_member(ctx: RulePassContext, push: PushFn) -> ProcedureStatementVisitor:
    """Cross-module rule: needs the project's class-member surfaces to know which
    member names are Friend-only (see AnalyzeModuleOptions)."""
    project_class_members = ctx.opts.project_class_members
    if project_class_members is None:
        return lambda member: None
    return check_late_bound_friend_member(
        ctx.source,
        ctx.symbols,
        ctx.opts.project_visible_symbols,
        project_class_members,
        ctx.opts.host_model,
        push,
    )


def _unknown_call_statement(ctx: RulePassContext, push: PushFn) -> ProcedureStatementVisitor:
    """Cross-module rule: only runs when the caller supplied the project's visible
    procedure names (see AnalyzeModuleOptions.known_procedures)."""
    known_procedures = ctx.opts.known_procedures
    if known_procedures is None:
        return lambda member: None
    return check_unknown_call_statement(
        ctx.source,
        ctx.symbols,
        known_procedures,
        ctx.opts.project_visible_symbols,
        ctx.opts.host_model,
        ctx.opts.designer_class,
        push,
        ctx.opts.project_class_members,
        own_object_member_names(ctx.opts),
    )


def _missing_scripting_reference(ctx: RulePassContext, push: PushFn) -> None:
    check_missing_scripting_reference(
        ctx.source,
        ctx.opts.referenced_libraries,
        {
            name.lower()
            for name in [
                *(surface.name for surface in ctx.member_ctx.project_class_members or []),
                *(symbol.name for symbol in ctx.symbols.root.children or []),
            ]
        },
        push,
    )


def _missing_library_reference(ctx: RulePassContext, push: PushFn) -> None:
    # The resolved model, not the raw option: a bare Excel project carries no
    # host_model, so the Excel default is supplied here, or the rule would know of
    # no library at all and stay silent. Port-only: the third argument says whether
    # the project's references are known at all, and a loose file's are not.
    check_missing_library_reference(
        ctx.source,
        ctx.opts.host_model if ctx.opts.host_model is not None else get_excel_object_model(),
        ctx.opts.referenced_hosts is not None,
        push,
        {
            name.lower()
            for name in [
                ctx.opts.module_name or "",
                *(surface.name for surface in ctx.member_ctx.project_class_members or []),
            ]
        },
    )


# The ordered table of active rules, in invocation order. Rules are independent:
# each only reads the shared context and reports through its own push.
DIAGNOSTIC_RULE_REGISTRY: tuple[DiagnosticRuleEntry, ...] = (
    DiagnosticRuleEntry(
        name="unterminatedStrings",
        run=lambda ctx, push: check_unterminated_strings(ctx.source, _active_only(ctx, push)),
    ),
    DiagnosticRuleEntry(
        name="invalidLineContinuations",
        run=lambda ctx, push: check_invalid_line_continuations(ctx.source, _active_only(ctx, push)),
    ),
    DiagnosticRuleEntry(
        name="duplicateProcedures",
        run=lambda ctx, push: check_duplicate_procedures(ctx.symbols.root.children or [], ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="duplicateDeclarations",
        run=lambda ctx, push: check_duplicate_declarations(ctx.symbols.root.children or [], ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="variableProcedureNameClash",
        run=lambda ctx, push: check_variable_procedure_name_clash(ctx.symbols.root.children or [], ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="enumMemberNameClash",
        run=lambda ctx, push: check_enum_member_name_clash(ctx.symbols.root.children or [], ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="duplicateModuleMembers",
        run=lambda ctx, push: check_duplicate_module_members(ctx.symbols.root.children or [], ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="duplicateEnumMembers",
        run=lambda ctx, push: check_duplicate_enum_members(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="duplicateTypeFields",
        run=lambda ctx, push: check_duplicate_type_fields(ctx.source, ctx.mod, ctx.activity, push),
    ),
    # Port-only: check_empty_type and check_identifier_too_long take the source first.
    DiagnosticRuleEntry(name="emptyType", run=lambda ctx, push: check_empty_type(ctx.source, ctx.mod, ctx.activity, push)),
    DiagnosticRuleEntry(
        name="refusedDeclarations",
        run=lambda ctx, push: check_refused_declarations(ctx.source, ctx.mod, ctx.module_kind, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="redimTypeChange",
        procedure_statements=lambda ctx, push: check_redim_type_change(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="tooManyParameters",
        run=lambda ctx, push: check_too_many_parameters(ctx.mod, ctx.module_kind, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="identifierTooLong",
        run=lambda ctx, push: check_identifier_too_long(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="udtParameterConstraints",
        run=lambda ctx, push: check_udt_parameter_constraints(ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="ambiguousBareProcedureCalls",
        procedure_statements=lambda ctx, push: check_ambiguous_bare_procedure_calls(
            ctx.source,
            ctx.symbols,
            ctx.module_name,
            ctx.opts.project_procedures,
            ctx.opts.project_visible_symbols,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="ambiguousEnumMemberReferences",
        run=lambda ctx, push: check_ambiguous_enum_member_references(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.activity,
            ctx.module_name,
            ctx.opts.known_procedures,
            ctx.opts.project_procedures,
            ctx.opts.project_class_members,
            ctx.opts.project_visible_symbols,
            ctx.opts.host_model,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="constAssignment",
        procedure_statements=lambda ctx, push: check_const_assignment(
            ctx.source, ctx.symbols, ctx.opts.project_visible_symbols, push
        ),
    ),
    # Port-only: check_option_explicit takes the source first.
    DiagnosticRuleEntry(
        name="optionExplicit",
        run=lambda ctx, push: check_option_explicit(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="localDeclarationOrder",
        run=lambda ctx, push: check_local_declaration_order(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.opts.project_visible_symbols,
            ctx.opts.host_model,
            ctx.activity,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="undeclaredVariables",
        run=lambda ctx, push: check_undeclared_variables(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.activity,
            ctx.opts.known_identifiers,
            ctx.opts.project_procedures,
            ctx.opts.project_class_members,
            ctx.opts.project_visible_symbols,
            ctx.opts.implicit_members,
            ctx.opts.module_kind,
            ctx.opts.host_model,
            ctx.opts.designer_class,
            ctx.opts.referenced_hosts,
            push,
            own_object_member_names(ctx.opts),
        ),
    ),
    DiagnosticRuleEntry(
        name="builtinsReadBare",
        run=lambda ctx, push: check_builtins_read_bare(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.activity,
            ctx.opts.project_visible_symbols,
            ctx.opts.module_kind,
            ctx.opts.host_model,
            ctx.opts.designer_class,
            ctx.opts.implicit_members,
            push,
            own_object_member_names(ctx.opts),
        ),
    ),
    DiagnosticRuleEntry(
        name="conditionValues",
        run=lambda ctx, push: check_condition_values(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="lockedArrays",
        run=lambda ctx, push: check_locked_arrays(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="deletedObjects",
        run=lambda ctx, push: check_deleted_objects(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="deletedSettings",
        run=lambda ctx, push: check_deleted_settings(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="handlerFlow",
        run=lambda ctx, push: check_handler_flow(
            ctx.source,
            ctx.mod,
            ctx.activity,
            push,
            ctx.opts.module_name if ctx.module_kind is ModuleSymbolKind.CLASS else None,
        ),
    ),
    DiagnosticRuleEntry(
        name="fileStatements",
        run=lambda ctx, push: check_file_statements(
            ctx.source, ctx.mod, ctx.activity, push, ctx.opts.project_opened_file_numbers
        ),
    ),
    DiagnosticRuleEntry(
        name="emptyFilePaths",
        procedure_statements=lambda ctx, push: check_empty_file_paths(
            ctx.source, ctx.symbols, ctx.activity, ctx.opts.project_visible_symbols, push
        ),
    ),
    DiagnosticRuleEntry(
        name="overflow",
        run=lambda ctx, push: check_overflow(
            ctx.source, ctx.mod, ctx.symbols, ctx.opts.project_visible_symbols, ctx.opts.host_model, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="hostArguments",
        block_headers=True,
        procedure_statements=lambda ctx, push: check_host_arguments(
            ctx.source, ctx.symbols, ctx.member_ctx, ctx.activity, push, workbook_sheets_to_check(ctx.opts)
        ),
    ),
    DiagnosticRuleEntry(
        name="collectionState",
        run=lambda ctx, push: check_collection_state(
            ctx.source,
            ctx.mod,
            ctx.activity,
            push,
            ctx.symbols,
            ctx.opts.project_integer_constants,
            ctx.opts.project_visible_symbols,
            ctx.opts.host_model,
        ),
    ),
    DiagnosticRuleEntry(
        name="byNameCalls",
        run=lambda ctx, push: check_by_name_calls(
            ctx.source, ctx.mod, ctx.symbols, ctx.member_ctx, ctx.opts.project_runnable_procedures, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="dictionaryState",
        run=lambda ctx, push: check_dictionary_state(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="documentNames",
        run=lambda ctx, push: check_document_names(
            ctx.source, ctx.mod, _host_name(ctx), _callable_names(ctx), ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="excelSessionState",
        run=lambda ctx, push: check_excel_session_state(
            ctx.source, ctx.mod, _callable_names(ctx), ctx.member_ctx, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="errorValues",
        run=lambda ctx, push: check_error_values(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="accessData",
        run=lambda ctx, push: check_access_data(ctx.source, ctx.mod, ctx.symbols, _host_name(ctx), ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="lateBoundObjectState",
        run=lambda ctx, push: check_late_bound_objects(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="variantValueMisuse",
        run=lambda ctx, push: check_variant_value_misuse(
            ctx.source, ctx.mod, ctx.symbols, ctx.activity, push, ctx.opts.project_visible_symbols
        ),
    ),
    DiagnosticRuleEntry(
        name="objectDefaultValue",
        procedure_statements=lambda ctx, push: check_object_default_values(
            ctx.source, ctx.symbols, ctx.member_ctx, push, ctx.activity
        ),
    ),
    DiagnosticRuleEntry(
        name="runtimeMemberNotFound",
        run=lambda ctx, push: check_runtime_member_not_found(
            ctx.source, ctx.mod, ctx.symbols, ctx.member_ctx, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="formContents",
        run=lambda ctx, push: check_form_contents(
            ctx.source, ctx.mod, ctx.opts.implicit_members, ctx.opts.project_name_mentions, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="classInstanceValues",
        run=lambda ctx, push: check_class_instance_values(
            ctx.source, ctx.mod, ctx.symbols, ctx.member_ctx, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="declarationForms",
        run=lambda ctx, push: check_declaration_forms(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="lineContinuationLimits",
        run=lambda ctx, push: check_line_continuation_limits(ctx.source, ctx.mod, _active_only(ctx, push)),
    ),
    DiagnosticRuleEntry(
        name="eventHandlerSignatures",
        run=lambda ctx, push: check_event_handler_signatures(
            ctx.mod, ctx.module_kind, ctx.opts, ctx.member_ctx, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="implementsMembers",
        run=lambda ctx, push: check_implements_members(
            ctx.source, ctx.mod, ctx.symbols, ctx.module_kind, ctx.opts.project_class_members, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="statementForms",
        run=lambda ctx, push: check_statement_forms(
            ctx.source, ctx.mod, ctx.symbols, ctx.opts.project_procedures, ctx.activity, push, ctx.member_ctx
        ),
    ),
    DiagnosticRuleEntry(
        name="vbaLibraryMembers",
        run=lambda ctx, push: check_vba_library_members(
            ctx.source, ctx.mod, ctx.symbols, ctx.opts.project_visible_symbols, ctx.activity, push, _host_name(ctx)
        ),
    ),
    DiagnosticRuleEntry(
        name="statementTypes",
        run=lambda ctx, push: check_statement_types(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.opts.project_visible_symbols,
            ctx.opts.project_types,
            ctx.activity,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="strayCharacters",
        run=lambda ctx, push: check_stray_characters(ctx.source, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="malformedLines",
        run=lambda ctx, push: check_malformed_lines(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="parentheses",
        run=lambda ctx, push: check_parentheses(ctx.source, ctx.mod, ctx.symbols, ctx.member_ctx, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="directiveForms",
        run=lambda ctx, push: check_directive_forms(ctx.source, ctx.mod, ctx.opts.conditional_compilation, push),
    ),
    DiagnosticRuleEntry(
        name="optionPlacement",
        run=lambda ctx, push: check_option_placement(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="duplicateOption",
        run=lambda ctx, push: check_duplicate_options(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="optionStatementForm",
        run=lambda ctx, push: check_option_statement_form(ctx.source, ctx.mod, ctx.opts, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="procedureHeader",
        run=lambda ctx, push: check_procedure_header(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="invalidIdentifierStarts",
        run=lambda ctx, push: check_invalid_identifier_starts(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="moduleDeclarationsInProcedureBodies",
        run=lambda ctx, push: check_module_declarations_in_procedure_bodies(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="moduleDeclarationsAfterProcedures",
        run=lambda ctx, push: check_module_declarations_after_procedures(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="moduleLevelStatementsOutsideProcedures",
        run=lambda ctx, push: check_module_level_statements_outside_procedures(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="reservedDeclarationNames",
        run=lambda ctx, push: check_reserved_declaration_names(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="moduleName",
        run=lambda ctx, push: check_module_name(ctx.source, ctx.opts.module_name, push, _host_name(ctx)),
    ),
    DiagnosticRuleEntry(
        name="propertySetterValueParameters",
        run=lambda ctx, push: check_property_setter_value_parameters(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="propertyAccessorSignatures",
        run=lambda ctx, push: check_property_accessor_signatures(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="parameterOrder",
        run=lambda ctx, push: check_parameter_order(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="parameterDefaultValues",
        run=lambda ctx, push: check_parameter_default_values(ctx.source, ctx.mod, ctx.activity, ctx.member_ctx, push),
    ),
    DiagnosticRuleEntry(
        name="parameterDefaultNotConstant",
        run=lambda ctx, push: check_non_constant_parameter_defaults(
            ctx.source, ctx.mod, ctx.activity, ctx.member_ctx, push
        ),
    ),
    DiagnosticRuleEntry(
        name="constValueNotConstant",
        run=lambda ctx, push: check_non_constant_const_values(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="enumMemberNotConstant",
        run=lambda ctx, push: check_non_constant_enum_member_values(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="unbalancedParens",
        run=lambda ctx, push: check_unbalanced_parens(ctx.source, push, ctx.activity),
    ),
    DiagnosticRuleEntry(
        name="invalidExpressionSyntax",
        procedure_statements=lambda ctx, push: check_invalid_expression_syntax(
            ctx.source, ctx.symbols, ctx.opts.project_visible_symbols, push
        ),
    ),
    DiagnosticRuleEntry(
        name="stringArithmeticOperands",
        procedure_statements=lambda ctx, push: check_string_arithmetic_operands(
            ctx.source, ctx.mod, ctx.symbols, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="divisionByZeroExpressions",
        block_headers=True,
        procedure_statements=lambda ctx, push: check_division_by_zero_expressions(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.opts.project_integer_constants,
            ctx.opts.project_visible_symbols,
            ctx.activity,
            push,
            ctx.opts.host_model,
        ),
    ),
    DiagnosticRuleEntry(
        name="dimInitializer",
        run=lambda ctx, push: check_dim_initializer(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="invalidRedimTargets",
        procedure_statements=lambda ctx, push: check_invalid_redim_targets(
            ctx.source, ctx.mod, ctx.symbols, ctx.opts.project_visible_symbols, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="redimImpossibleBounds",
        procedure_statements=lambda ctx, push: check_redim_impossible_bounds(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.opts.project_integer_constants,
            ctx.opts.project_visible_symbols,
            ctx.activity,
            push,
            ctx.opts.host_model,
        ),
    ),
    DiagnosticRuleEntry(
        name="arrayDeclarationImpossibleBounds",
        run=lambda ctx, push: check_array_declaration_bounds(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.opts.project_integer_constants,
            ctx.opts.project_visible_symbols,
            ctx.activity,
            push,
            ctx.opts.host_model,
        ),
    ),
    DiagnosticRuleEntry(
        name="redimPreserveDimensions",
        run=lambda ctx, push: check_redim_preserve_dimensions(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="unallocatedDynamicArrayAccess",
        run=lambda ctx, push: check_unallocated_dynamic_array_access(
            ctx.source, ctx.mod, ctx.symbols, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="collectionLoopCounters",
        run=lambda ctx, push: check_collection_loop_counters(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="arraySubscriptOutOfBounds",
        run=lambda ctx, push: check_fixed_array_subscript_bounds(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.activity,
            push,
            ctx.opts.project_integer_constants,
            ctx.opts.project_visible_symbols,
            ctx.opts.host_model,
        ),
    ),
    DiagnosticRuleEntry(
        name="declareStatements",
        run=lambda ctx, push: check_declare_statements(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="unusableDeclareCalls",
        run=lambda ctx, push: check_unusable_declare_calls(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="typeFieldArrays",
        run=lambda ctx, push: check_type_field_arrays(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.activity,
            push,
            ctx.opts.project_integer_constants,
            ctx.opts.project_visible_symbols,
            ctx.opts.host_model,
        ),
    ),
    DiagnosticRuleEntry(
        name="typeMembers",
        run=lambda ctx, push: check_type_members(ctx.source, ctx.mod, ctx.symbols, ctx.member_ctx, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="midStatementLiteralTarget",
        run=lambda ctx, push: check_mid_statement_literal_target(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="eraseTargets",
        procedure_statements=lambda ctx, push: check_erase_targets(
            ctx.source, ctx.symbols, ctx.opts.project_visible_symbols, push
        ),
    ),
    DiagnosticRuleEntry(
        name="typeDeclarationCharacterAsClause",
        run=lambda ctx, push: check_type_declaration_character_as_clause(ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="unexpectedDeclarationTokens",
        run=lambda ctx, push: check_unexpected_declaration_tokens(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="fixedLengthStringBounds",
        run=lambda ctx, push: check_fixed_length_string_bounds(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="objectModulePublicMembers",
        run=lambda ctx, push: check_object_module_public_members(
            ctx.source, ctx.mod, ctx.module_kind, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="eventDeclarationModuleKind",
        run=lambda ctx, push: check_event_declaration_module_kind(
            ctx.source, ctx.mod, ctx.module_kind, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="meOutsideObjectModule",
        procedure_statements=lambda ctx, push: check_me_outside_object_module(ctx.module_kind, ctx.source, push),
    ),
    DiagnosticRuleEntry(
        name="withEventsDeclarations",
        run=lambda ctx, push: check_with_events_declarations(
            ctx.source, ctx.mod, ctx.module_kind, ctx.activity, push, ctx.member_ctx.project_class_members or ()
        ),
    ),
    DiagnosticRuleEntry(
        name="friendDeclarations",
        run=lambda ctx, push: check_friend_declarations(ctx.source, ctx.mod, ctx.module_kind, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="implementsStatementPlacement",
        run=lambda ctx, push: check_implements_statement_placement(
            ctx.source, ctx.mod, ctx.module_kind, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="raiseEventTargets",
        run=lambda ctx, push: check_raise_event_targets(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="invalidPropertyUse",
        procedure_statements=lambda ctx, push: check_invalid_property_use(ctx.source, ctx.member_ctx, push),
    ),
    DiagnosticRuleEntry(
        name="raiseEventArguments",
        run=lambda ctx, push: check_raise_event_arguments(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="longLongNarrowing",
        run=lambda ctx, push: check_long_long_narrowing(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.opts.conditional_compilation,
            ctx.opts.host,
            ctx.opts.project_procedures,
            ctx.opts.project_visible_symbols,
            ctx.activity,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="addressOfUse",
        run=lambda ctx, push: check_address_of_use(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.opts.project_procedures,
            ctx.opts.project_class_members,
            ctx.opts.conditional_compilation,
            ctx.activity,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="declarePtrSafeForWin64",
        run=lambda ctx, push: check_declare_ptr_safe_for_win64(
            ctx.source, ctx.mod, ctx.opts.conditional_compilation, ctx.opts.host, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="eventHandlerModuleScope",
        run=lambda ctx, push: check_event_handler_module_scope(
            ctx.source, ctx.mod, ctx.module_name, ctx.module_kind, ctx.opts.document_type, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(
        name="invalidAsTypeNames",
        run=lambda ctx, push: check_invalid_as_type_names(ctx.source, ctx.mod, ctx.activity, ctx.opts, push),
    ),
    DiagnosticRuleEntry(
        name="callParens",
        procedure_statements=lambda ctx, push: check_call_parens(
            ctx.source,
            ctx.symbols,
            ctx.opts.project_procedures,
            ctx.opts.project_visible_symbols,
            ctx.member_ctx,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="expressionCallParens",
        procedure_statements=lambda ctx, push: check_expression_call_parens(
            ctx.source, ctx.symbols, ctx.opts.project_procedures, ctx.opts.project_visible_symbols, push
        ),
    ),
    DiagnosticRuleEntry(
        name="setAssignments",
        procedure_statements=lambda ctx, push: check_set_assignments(
            ctx.source, ctx.symbols, ctx.opts.project_visible_symbols, ctx.member_ctx, push, ctx.activity
        ),
    ),
    DiagnosticRuleEntry(
        name="exitStatements",
        procedure_statements=lambda ctx, push: check_exit_statements(ctx.source, push),
    ),
    DiagnosticRuleEntry(
        name="duplicateLabels",
        run=lambda ctx, push: check_duplicate_labels(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="declarationOrder",
        run=lambda ctx, push: check_declaration_order(
            ctx.source,
            ctx.mod,
            ctx.opts.module_name,
            ctx.opts.project_integer_constants,
            ctx.opts.project_class_members,
            ctx.activity,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="lineNumberRange",
        run=lambda ctx, push: check_line_number_range(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="undefinedLabels",
        run=lambda ctx, push: check_undefined_labels(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="reservedLabels",
        run=lambda ctx, push: check_reserved_labels(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="elseBranchOrder",
        run=lambda ctx, push: check_else_branch_order(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="statementContext",
        run=lambda ctx, push: check_statement_context(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="duplicateCaseElse",
        run=lambda ctx, push: check_duplicate_case_else(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="malformedStatements",
        run=lambda ctx, push: check_malformed_statements(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="elseWithoutIf",
        run=lambda ctx, push: check_else_without_if(ctx.source, ctx.mod, ctx.activity, push),
    ),
    # Port-only: check_for_each_loop_types takes the option fields it reads
    # rather than the whole options object.
    DiagnosticRuleEntry(
        name="forEachLoopTypes",
        run=lambda ctx, push: check_for_each_loop_types(
            ctx.mod,
            ctx.symbols,
            ctx.opts.project_visible_symbols,
            ctx.opts.project_types,
            ctx.member_ctx.model,
            ctx.activity,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="arrayBoundIntrinsicArguments",
        procedure_statements=lambda ctx, push: check_array_bound_intrinsic_arguments(
            ctx.source, ctx.symbols, ctx.opts.project_visible_symbols, push
        ),
    ),
    DiagnosticRuleEntry(
        name="scalarMemberAccess",
        procedure_statements=lambda ctx, push: check_scalar_member_access(
            ctx.source, ctx.symbols, ctx.opts.project_visible_symbols, push, ctx.member_ctx
        ),
    ),
    DiagnosticRuleEntry(
        name="objectVariableNotSet",
        run=lambda ctx, push: check_object_variable_not_set(
            ctx.source, ctx.mod, ctx.symbols, ctx.member_ctx, ctx.activity, push
        ),
    ),
    DiagnosticRuleEntry(name="missingScriptingReference", run=_missing_scripting_reference),
    DiagnosticRuleEntry(name="missingLibraryReference", run=_missing_library_reference),
    DiagnosticRuleEntry(
        name="memberNotFound",
        procedure_statements=lambda ctx, push: check_member_not_found(ctx.source, ctx.member_ctx, push),
    ),
    DiagnosticRuleEntry(
        name="moduleMemberForms",
        procedure_statements=lambda ctx, push: check_module_member_forms(
            ctx.source,
            ctx.symbols,
            ctx.member_ctx,
            ctx.opts.project_visible_symbols,
            _OPTION_EXPLICIT_RE.search(ctx.source) is not None,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="invalidParamArrayUse",
        procedure_statements=lambda ctx, push: check_param_array_use(
            ctx.source,
            callable_type_signatures_for(ctx.symbols, ctx.opts.project_procedures),
            push,
            ctx.symbols,
            ctx.member_ctx,
        ),
    ),
    DiagnosticRuleEntry(
        name="nonCallableCallStatement",
        procedure_statements=lambda ctx, push: check_non_callable_call_statement(
            ctx.source, ctx.symbols, ctx.opts.known_procedures, ctx.opts.project_visible_symbols, push
        ),
    ),
    DiagnosticRuleEntry(
        name="argumentCount",
        procedure_statements=lambda ctx, push: check_argument_count(
            ctx.source,
            ctx.symbols,
            ctx.opts.project_procedures,
            ctx.opts.project_visible_symbols,
            ctx.member_ctx,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="omittedArgumentReads",
        procedure_statements=lambda ctx, push: check_omitted_argument_reads(
            ctx.source, ctx.mod, ctx.symbols, ctx.activity, ctx.opts.project_visible_symbols, push
        ),
    ),
    DiagnosticRuleEntry(
        name="argumentTypes",
        procedure_statements=lambda ctx, push: check_argument_types(
            ctx.source,
            ctx.symbols,
            ctx.opts.project_procedures,
            ctx.opts.project_visible_symbols,
            ctx.member_ctx,
            push,
            ctx.activity,
        ),
    ),
    DiagnosticRuleEntry(
        name="runtimeArgumentValues",
        block_headers=True,
        procedure_statements=lambda ctx, push: check_runtime_argument_values(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.opts.project_procedures,
            ctx.opts.project_integer_constants,
            ctx.opts.project_visible_symbols,
            ctx.activity,
            push,
            ctx.opts.host_model,
        ),
    ),
    DiagnosticRuleEntry(
        name="runtimeConversionValues",
        block_headers=True,
        procedure_statements=lambda ctx, push: check_runtime_conversion_values(
            ctx.source, ctx.symbols, ctx.opts.project_visible_symbols, push, ctx.activity
        ),
    ),
    DiagnosticRuleEntry(
        name="assignmentTypes",
        run=lambda ctx, push: check_assignment_types(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.opts.project_visible_symbols,
            ctx.member_ctx,
            ctx.activity,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="typeOfIsAlwaysFalse",
        procedure_expressions=lambda ctx, push: check_type_of_is_compatibility(ctx.symbols, ctx.member_ctx, push),
    ),
    DiagnosticRuleEntry(
        name="typeofMissingOperand",
        run=lambda ctx, push: check_type_of_missing_operand(ctx.source, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="isOperatorNonObject",
        procedure_expressions=lambda ctx, push: check_is_operator_operands(ctx.symbols, push),
    ),
    DiagnosticRuleEntry(
        name="isOperandsInConditions",
        run=lambda ctx, push: check_is_operands_in_conditions(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="nonScalarBinaryOperand",
        procedure_expressions=lambda ctx, push: check_binary_operand_scalar(ctx.symbols, push),
    ),
    DiagnosticRuleEntry(
        name="argumentShapeMismatch",
        procedure_statements=lambda ctx, push: check_argument_shape(
            ctx.source,
            ctx.symbols,
            ctx.opts.project_procedures,
            ctx.opts.project_visible_symbols,
            ctx.member_ctx,
            push,
            ctx.mod,
            ctx.activity,
        ),
    ),
    DiagnosticRuleEntry(
        name="suffixedLiteralOverflow",
        run=lambda ctx, push: check_suffixed_literal_overflow(ctx.source, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="missingReturnAssignments",
        run=lambda ctx, push: check_missing_return_assignments(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.opts.project_procedures,
            ctx.activity,
            ctx.opts.module_name,
            ctx.opts.implemented_interfaces,
            push,
        ),
    ),
    DiagnosticRuleEntry(name="unknownCallStatement", procedure_statements=_unknown_call_statement),
    DiagnosticRuleEntry(name="lateBoundFriendMember", procedure_statements=_late_bound_friend_member),
    DiagnosticRuleEntry(
        name="unusedDeclarations",
        run=lambda ctx, push: check_unused_declarations(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        # A Private procedure is reachable from its own module only, so the module's
        # text decides; the project's string literals are consulted for a name
        # reached through Application.Run, OnTime and their kin.
        name="unusedPrivateProcedures",
        run=lambda ctx, push: check_unused_private_procedures(
            ctx.source,
            ctx.mod,
            ctx.symbols,
            ctx.module_kind,
            ctx.activity,
            ctx.opts.project_string_literal_words,
            push,
        ),
    ),
    DiagnosticRuleEntry(
        name="unreachableCode",
        run=lambda ctx, push: check_unreachable_code(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="docComments",
        run=lambda ctx, push: check_doc_comments(ctx.source, ctx.mod, ctx.activity, push),
    ),
)
