"""Rule family: assignment-statement rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/assignments.ts. Only the
Mid-statement literal-target rule is ported in M7; the type-coupled assignment
rules (assignment type mismatch, Set object types) land in M8.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence

from ...conditional import ConditionalActivityTracker
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, ProcKind, Span
from ...symbols.name_resolution import (
    BareIdentifierContext,
    BareIdentifierResolutionInput,
    BareIdentifierResolutionScope,
    resolve_bare_identifier_binding,
)
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureSignature, VbaSymbol, VbaSymbolKind
from ...types.type_inference import (
    DeclaredValueShape,
    SourceDeclaredShape,
    SourceDeclaredType,
    declaration_shape_environment_for,
    declared_shape_for_source_binding,
    declared_type_for_source_binding,
    declared_value_type_for_qualified_source_binding,
    declared_value_type_for_source_binding,
    procedure_symbol_for,
    type_environment_for,
)
from ...types.type_names import is_known_object_assignment_type, is_known_scalar_type, normalize_type
from ..argument_inference import (
    incompatibility_reason,
    infer_argument_type,
    nonnumeric_string_arithmetic_operand,
)
from ..call_extraction import (
    CallableTypeSignature,
    CallArguments,
    extract_call,
    extract_qualified_call,
    named_argument_slot,
)
from ..callable_signatures import (
    build_module_type_signatures,
    callable_signature_for_call,
    callable_type_signatures_for,
    source_name_scope_for,
)
from ..context import PushFn
from ..walker import (
    ProcedureStatementVisitor,
    active_module_members,
    bare_assignment_target,
    declared_name_span,
    for_each_statement,
    set_assignment_target,
    statement_tokens_after_leading_label,
    strip_header_brackets,
    token_name,
)


def check_const_assignment(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        proc_sym = procedure_symbol_for(symbols, member)

        def visitor(stmt: LeafStatementNode) -> None:
            hit = bare_assignment_target(source, stmt.span)
            if hit is None:
                return
            binding = resolve_bare_identifier_binding(
                BareIdentifierResolutionInput(
                    current_module=symbols,
                    name=hit[0],
                    context=BareIdentifierContext.ASSIGNMENT_TARGET,
                    enclosing_procedure=proc_sym,
                    project_visible_symbols=list(project_visible_symbols)
                    if project_visible_symbols
                    else [],
                )
            )
            if binding.scope is not BareIdentifierResolutionScope.AMBIGUOUS and any(
                d.kind is VbaSymbolKind.CONSTANT for d in binding.definitions
            ):
                push("constAssignment", f"Cannot assign to constant '{hit[0]}'.", hit[1])

        return visitor

    return factory


def check_assignment_types(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Bare scalar-assignment type compatibility (the member-access surface is M9)."""
    module_signatures = build_module_type_signatures(symbols)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        env = type_environment_for(symbols, member)
        shapes = declaration_shape_environment_for(symbols, member)
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        proc_sym = procedure_symbol_for(symbols, member)

        # The resolvers close over this iteration's proc_sym; for_each_statement
        # invokes the visitor synchronously below, so the closures always see the
        # current member's binding.
        def resolve_expression_type(name: str) -> SourceDeclaredType:
            return declared_value_type_for_source_binding(symbols, proc_sym, project_visible_symbols, name)

        def resolve_qualified_expression_type(qualifier: str, name: str) -> SourceDeclaredType:
            return declared_value_type_for_qualified_source_binding(
                symbols, project_visible_symbols, qualifier, name
            )

        def resolve_target_shape(name: str) -> SourceDeclaredShape:
            return declared_shape_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.ASSIGNMENT_TARGET
            )

        def resolve_source_shape(name: str) -> SourceDeclaredShape:
            return declared_shape_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.EXPRESSION
            )

        def visit(stmt: LeafStatementNode) -> None:
            assignment = bare_assignment_target(source, stmt.span)
            if assignment is None:
                return
            name, name_span, value_tokens = assignment
            target_type = declared_type_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.ASSIGNMENT_TARGET
            )
            expected = target_type.as_type if target_type.resolved else env.get(name.lower())
            if not expected:
                return
            if is_known_object_assignment_type(expected):
                push(
                    "setRequired",
                    f"Object assignment to '{name}' requires Set because it is declared as {expected}.",
                    name_span,
                )
                return
            array_source = _array_assignment_to_scalar_source(
                name, value_tokens, stmt.span.start, expected, shapes,
                resolve_target_shape, resolve_source_shape,
            )
            if array_source is not None:
                src_name, src_span = array_source
                push(
                    "arrayAssignmentToScalar",
                    f"Array variable '{src_name}' cannot be assigned to scalar '{name}'. "
                    "Assign an array element or use a Variant/array target.",
                    src_span,
                )
                return
            string_arithmetic = nonnumeric_string_arithmetic_operand(
                expected, value_tokens, stmt.span.start
            )
            if string_arithmetic is not None:
                push(
                    "stringArithmeticCoercion",
                    f"Assignment to '{name}' expects {expected}, but this numeric expression "
                    f"contains {string_arithmetic.label}. This will raise Run-time error '13': "
                    "Type mismatch.",
                    string_arithmetic.span,
                )
                return
            actual = infer_argument_type(
                value_tokens, stmt.span.start, env, module_signatures, source_names,
                resolve_expression_type, resolve_qualified_expression_type,
            )
            if actual is None:
                return
            reason = incompatibility_reason(expected, actual)
            if not reason:
                return
            push(
                "assignmentTypeMismatch",
                f"Assignment to '{name}' expects {expected}, but got {actual.label}. {reason}",
                actual.span,
            )

        for_each_statement(member.body, visit, activity)


def _array_assignment_to_scalar_source(
    name: str,
    value_tokens: list[VbaToken],
    base_offset: int,
    expected_type: str,
    shapes: Mapping[str, DeclaredValueShape],
    resolve_target_shape: Callable[[str], SourceDeclaredShape],
    resolve_source_shape: Callable[[str], SourceDeclaredShape],
) -> tuple[str, Span] | None:
    if not _is_known_scalar_assignment_target(name, expected_type, shapes, resolve_target_shape):
        return None
    if len(value_tokens) != 1:
        return None
    tok = value_tokens[0]
    source_name = token_name(tok)
    if not source_name:
        return None
    resolved = resolve_source_shape(source_name)
    source_shape = resolved.shape if resolved.resolved else shapes.get(source_name.lower())
    if source_shape is None or not source_shape.is_array:
        return None
    return (source_name, Span(base_offset + tok.start, base_offset + tok.end))


def _is_known_scalar_assignment_target(
    name: str,
    expected_type: str,
    shapes: Mapping[str, DeclaredValueShape],
    resolve_shape: Callable[[str], SourceDeclaredShape],
) -> bool:
    resolved = resolve_shape(name)
    target_shape = resolved.shape if resolved.resolved else shapes.get(name.lower())
    if target_shape is not None and target_shape.is_array:
        return False
    as_type = target_shape.as_type if target_shape is not None else None
    normalized = normalize_type(as_type if as_type is not None else expected_type)
    return normalized is not None and is_known_scalar_type(normalized)


def check_set_assignments(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """`Set x = ...` where x is a declared scalar requires an object variable.

    The object-valued target branch (Set against an incompatible object type) needs
    the host member surface and is deferred to M9.
    """

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        proc_sym = procedure_symbol_for(symbols, member)

        def visitor(stmt: LeafStatementNode) -> None:
            target = set_assignment_target(source, stmt.span)
            if target is None:
                return
            name, span, _value_tokens = target
            target_declared_type = declared_type_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.ASSIGNMENT_TARGET
            )
            expected = target_declared_type.as_type if target_declared_type.resolved else env.get(name.lower())
            target_type = normalize_type(expected)
            if not target_type or not is_known_scalar_type(target_type):
                return  # object/unknown target: object-type mismatch is deferred (M9)
            push(
                "setRequiresObject",
                f"Set assignment requires an object variable, but '{name}' is declared as {expected}.",
                span,
            )

        return visitor

    return factory


def check_missing_return_assignments(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """An untyped Function/Property Get with no return assignment silently returns Empty."""
    module_signatures = callable_type_signatures_for(symbols, project_procedures)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        if member.proc_kind not in (ProcKind.FUNCTION, ProcKind.PROPERTY_GET):
            continue
        if not member.closed or member.return_type:
            continue
        if _procedure_has_return_assignment(source, member, activity, module_signatures):
            continue
        proc_label = "Property Get" if member.proc_kind is ProcKind.PROPERTY_GET else "Function"
        push(
            "missingReturnAssignment",
            f"{proc_label} '{member.name}' has no return assignment; VBA will return the default "
            f"value. Assign to '{member.name}' before exit if a value is intended.",
            declared_name_span(source, member.span, member.name),
        )


def _procedure_has_return_assignment(
    source: str,
    proc: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    module_signatures: Mapping[str, CallableTypeSignature],
) -> bool:
    lower = proc.name.lower()
    found = False

    def visit(stmt: LeafStatementNode) -> None:
        nonlocal found
        if found:
            return
        bare = bare_assignment_target(source, stmt.span)
        if bare is not None and bare[0].lower() == lower:
            found = True
            return
        set_target = set_assignment_target(source, stmt.span)
        if set_target is not None and set_target[0].lower() == lower:
            found = True
            return
        call = extract_call(source, stmt.span)
        qualified = None if call else extract_qualified_call(source, stmt.span, module_signatures)
        effective = call or qualified
        if effective is not None and _call_passes_name_to_by_ref_param(effective, lower, module_signatures):
            found = True

    for_each_statement(proc.body, visit, activity)
    return found


def _call_passes_name_to_by_ref_param(
    call: CallArguments, lower_name: str, module_signatures: Mapping[str, CallableTypeSignature]
) -> bool:
    sig = callable_signature_for_call(call, module_signatures)
    if sig is None:
        return False
    positional_index = 0
    for slot in call.slots:
        named = named_argument_slot(slot)
        if named is not None:
            param = next(
                (p for p in sig.params if strip_header_brackets(p.name).lower() == named[0].lower()),
                None,
            )
            value_slot = named[1]
        else:
            param = sig.params[min(positional_index, len(sig.params) - 1)] if sig.params else None
            positional_index += 1
            value_slot = slot
        if param is None or not param.by_ref or not _single_slot_name_equals(value_slot, lower_name):
            continue
        return True
    return False


def _single_slot_name_equals(slot: list[VbaToken], lower_name: str) -> bool:
    toks = [t for t in slot if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE]
    return len(toks) == 1 and (token_name(toks[0]) or "").lower() == lower_name


_TYPE_CHAR_SUFFIX = re.compile(r"[$%&!#@]$")


def _mid_base_word(tok: VbaToken | None) -> str:
    """Suffix-stripped, lower-cased word for a token (keyword or identifier)."""
    if tok is None:
        return ""
    text = token_name(tok)
    if text is None:
        text = tok.raw_text
    return _TYPE_CHAR_SUFFIX.sub("", text.lower())


def check_mid_statement_literal_target(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    if _module_shadows_mid_intrinsic(symbols) or _module_redim_declares_mid_intrinsic(
        source, mod, activity
    ):
        return

    def visit(stmt: LeafStatementNode) -> None:
        hit = _mid_statement_literal_target_violation(source, stmt.span)
        if hit is not None:
            span, message = hit
            push("midStatementLiteralTarget", message, span)

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            for_each_statement(member.body, visit, activity)


def _module_shadows_mid_intrinsic(symbols: ModuleSymbols) -> bool:
    """True when a module declares any symbol that shadows the Mid/MidB intrinsic."""
    return any(
        _TYPE_CHAR_SUFFIX.sub("", sym.name.lower()) in ("mid", "midb") for sym in symbols.all
    )


def _module_redim_declares_mid_intrinsic(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None
) -> bool:
    """True when a ReDim implicitly declares an array named mid/midb (absent from symbols)."""
    found = False

    def visit(stmt: LeafStatementNode) -> None:
        nonlocal found
        if found:
            return
        toks = statement_tokens_after_leading_label(source, stmt.span)
        if not toks or _mid_base_word(toks[0]) != "redim":
            return
        start = 2 if len(toks) > 1 and _mid_base_word(toks[1]) == "preserve" else 1
        for group in split_top_level_token_groups(toks, start, ","):
            if group and _mid_base_word(group[0]) in ("mid", "midb"):
                found = True
                return

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            for_each_statement(member.body, visit, activity)
            if found:
                return True
    return False


def _mid_statement_literal_target_violation(source: str, span: Span) -> tuple[Span, str] | None:
    toks = statement_tokens_after_leading_label(source, span)
    if not toks:
        return None
    if _mid_base_word(toks[0]) not in ("mid", "midb"):
        return None
    # Handle both lexings of `Mid$`: a single `Mid$` token, or `Mid` then `$`.
    paren_index = 1
    if len(toks) > paren_index and toks[paren_index].raw_text == "$":
        paren_index = 2
    if paren_index >= len(toks) or toks[paren_index].raw_text != "(":
        return None
    close = match_paren_from(toks, paren_index)
    if close <= paren_index + 1:
        return None  # empty or unbalanced argument list
    # The Mid replacement-statement form: the matching `)` is followed by `=`.
    if close + 1 >= len(toks) or toks[close + 1].raw_text != "=":
        return None
    arg_toks = [tok for tok in toks[paren_index + 1 : close] if tok.kind is not TokenKind.COMMENT]
    slots = split_top_level_token_groups(arg_toks, 0, ",")
    target = slots[0] if slots else None
    if not target or len(target) != 1 or target[0].kind is not TokenKind.STRING_LITERAL:
        return None  # target is not exactly one string literal
    return (
        Span(span.start + target[0].start, span.start + target[0].end),
        "The target of a Mid statement must be a writable String variable, not a "
        "string literal. Assigning into a literal is a compile error.",
    )
