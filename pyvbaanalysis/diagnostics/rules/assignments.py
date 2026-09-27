"""Rule family: assignment-statement rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/assignments.ts: constant
targets, scalar and member-access assignment types, Set targets and their object
types, missing return assignments, and the Mid-statement literal target. The
rules that read a statement structurally also read the statements a single-line
`If` carries.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...completion.member_access import (
    MemberCompletionContext,
    resolve_exact_member_completion,
)
from ...completion.member_access import (
    is_known_object_assignment_type as is_known_object_assignment_type_ctx,
)
from ...conditional import ConditionalActivityTracker, inactive_node_skip
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ForBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    ProcKind,
    Span,
    iter_body_nodes,
)
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
from ...types.type_names import is_known_scalar_type, normalize_type
from ..argument_inference import (
    SourceDeclaredTypeResolver,
    SourceQualifiedDeclaredTypeResolver,
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
    SourceNameScope,
    build_module_type_signatures,
    callable_signature_for_call,
    callable_type_signatures_for,
    is_member_statement_chain_through,
    source_name_scope_for,
)
from ..context import PushFn
from ..walker import (
    ProcedureStatementVisitor,
    active_module_members,
    bare_assignment_target,
    declared_name_span,
    first_executable_token_index,
    for_each_statement,
    set_assignment_target,
    statement_and_branch_spans,
    statement_tokens,
    statement_tokens_after_leading_label,
    strip_header_brackets,
    token_name,
    token_text,
    top_level_operator_index,
)
from .type_of_is import object_assignment_incompatibility_reason, object_let_assignment_verdict


def check_const_assignment(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        proc_sym = procedure_symbol_for(symbols, member)

        def visitor(stmt: LeafStatementNode) -> None:
            for span in statement_and_branch_spans(stmt):
                check_span(span)

        def check_span(span: Span) -> None:
            hit = bare_assignment_target(source, span)
            if hit is None:
                return
            binding = resolve_bare_identifier_binding(
                BareIdentifierResolutionInput(
                    current_module=symbols,
                    name=hit[0],
                    context=BareIdentifierContext.ASSIGNMENT_TARGET,
                    enclosing_procedure=proc_sym,
                    project_visible_symbols=project_visible_symbols or (),
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
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Scalar and member-access assignment type compatibility (`x = v`, `obj.M = v`)."""
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
            for span in statement_and_branch_spans(stmt):
                check_assignment_span(span)

        def check_assignment_span(span: Span) -> None:
            assignment = bare_assignment_target(source, span)
            if assignment is None:
                return
            name, name_span, value_tokens = assignment
            target_type = declared_type_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.ASSIGNMENT_TARGET
            )
            expected = target_type.as_type if target_type.resolved else env.get(name.lower())
            if not expected:
                return
            if is_known_object_assignment_type_ctx(expected, member_ctx):
                # The VBE compiles a bare `=` to an object variable as a Let through
                # the type's default member (XLIDE issue #107): `r = 5` writes the
                # Range's Value. What is reported is what the default member makes
                # of it.
                verdict = object_let_assignment_verdict(expected, member_ctx)
                if verdict == "argument":
                    push(
                        "setRequired",
                        f"Assignment to '{name}' requires Set: the default member of {expected} "
                        "takes an argument, so a Let cannot reach it. This is a VBE compile error: "
                        "Argument not optional.",
                        name_span,
                    )
                elif verdict == "noDefault":
                    push(
                        "setRequired",
                        f"Assignment to '{name}' requires Set: {expected} has no default member for "
                        "a Let to reach. This will raise Run-time error '438': Object doesn't "
                        "support this property or method.",
                        name_span,
                    )
                return
            array_source = _array_assignment_to_scalar_source(
                name, value_tokens, span.start, expected, shapes,
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
            # A dynamic Byte array takes a String whole - `b = "abc"` copies the
            # string's bytes, and a String takes the array back (XLIDE issue #105,
            # measured in Excel 16.0). The element type is not what the value is
            # checked against there.
            resolved_target_shape = resolve_target_shape(name)
            target_shape = (
                resolved_target_shape.shape if resolved_target_shape.resolved else shapes.get(name.lower())
            )
            if target_shape is not None and target_shape.is_array and normalize_type(target_shape.as_type) == "byte":
                return
            string_arithmetic = nonnumeric_string_arithmetic_operand(
                expected, value_tokens, span.start
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
                value_tokens, span.start, env, module_signatures, source_names,
                resolve_expression_type, resolve_qualified_expression_type,
                source=source, member_ctx=member_ctx,
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
        check_member_assignment_types(
            source, member, env, module_signatures, source_names, member_ctx, activity,
            push, resolve_expression_type, resolve_qualified_expression_type,
        )


@dataclass(frozen=True, slots=True)
class _MemberAssignmentTarget:
    member: str
    label: str
    member_span: Span
    value_tokens: list[VbaToken]
    uses_set: bool


def _member_assignment_target(source: str, span: Span) -> _MemberAssignmentTarget | None:
    """Port of memberAssignmentTarget: an `obj.Member = value` / `Set obj.Member = ...`
    LHS whose last two tokens are `. Member`. Returns None for bare or compound LHS."""
    toks = statement_tokens(source, span)
    i = first_executable_token_index(toks)
    if i >= len(toks):
        return None
    uses_set = token_text(toks[i]) == "set"
    if uses_set or token_text(toks[i]) == "let":
        i += 1
    eq = top_level_operator_index(toks[i:], "=")
    if eq < 0:
        return None
    equals_index = i + eq
    lhs = toks[i:equals_index]
    if len(lhs) < 2:
        return None
    member_tok = lhs[-1]
    member_name = token_name(member_tok)
    if not member_name or lhs[-2].raw_text != ".":
        return None
    # A target is one receiver chain ending in the member. Anything else before
    # the `=` is another statement comparing the member: an ElseIf or Case
    # header, a single-line If's condition, a call given the comparison
    # (`Debug.Print w.Part = "a"`). ReDim's `ElseIf ReDimUI.SenderPart = "plus"
    # Then` compiles, and was reported as assigning to 'ElseIf
    # ReDimUI.SenderPart' (XLIDE #78).
    if not is_member_statement_chain_through(lhs, 0, len(lhs) - 1):
        return None
    if any(t.kind is TokenKind.OPERATOR and t.raw_text == "=" for t in lhs):
        return None
    return _MemberAssignmentTarget(
        member=member_name,
        label=source[span.start + lhs[0].start : span.start + member_tok.end].strip(),
        member_span=Span(span.start + member_tok.start, span.start + member_tok.end),
        value_tokens=list(toks[equals_index + 1 :]),
        uses_set=uses_set,
    )


def check_member_assignment_types(
    source: str,
    member: ProcedureNode,
    env: Mapping[str, str],
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    resolve_expression_type: SourceDeclaredTypeResolver | None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None,
) -> None:
    """Port of checkMemberAssignmentTypes: `obj.Member = value` type compatibility.

    Only source-backed project members carry writability, so a host member (whose
    writability is unknown) and an unresolved receiver both yield no diagnostic, the
    no-false-positive gate. The expected value type is the member's declared write
    type (falling back to its return type)."""
    if not member_ctx.project_class_members:
        return

    def check_statement(span: Span) -> None:
        assignment = _member_assignment_target(source, span)
        if assignment is None:
            return
        target = resolve_exact_member_completion(
            source, assignment.member, assignment.member_span.end, member_ctx
        )
        if target is None or target.writable is None:
            return
        if target.writable is False:
            push(
                "readonlyMemberAssignment",
                f"Cannot assign to read-only property '{assignment.label}'.",
                assignment.member_span,
            )
            return
        expected = target.write_type if target.write_type is not None else target.returns
        if assignment.uses_set:
            if expected and is_known_scalar_type(normalize_type(expected) or ""):
                push(
                    "setRequiresObject",
                    f"Set assignment requires an object-valued target, but '{assignment.label}' "
                    f"expects {expected}.",
                    assignment.member_span,
                )
                return
            # `Set h.Item = x` needs a Property Set; with only a Property Let the VBE
            # refuses it, "Invalid use of property" (XLIDE issue #107).
            if target.let_accessor and not target.set_accessor:
                push(
                    "setRequiresObject",
                    f"Set assignment to '{assignment.label}' needs a Property Set, but the property "
                    "declares only a Property Let. This is a VBE compile error: Invalid use of property.",
                    assignment.member_span,
                )
                return
            actual = infer_argument_type(
                assignment.value_tokens, span.start, env, module_signatures, source_names,
                resolve_expression_type, resolve_qualified_expression_type,
                source=source, member_ctx=member_ctx,
            )
            reason = object_assignment_incompatibility_reason(expected, actual, member_ctx)
            if reason:
                push(
                    "assignmentObjectTypeMismatch",
                    f"Object assignment to '{assignment.label}' expects {expected}, but got "
                    f"{actual.label if actual is not None else None}. {reason}",
                    actual.span if actual is not None else assignment.member_span,
                )
            return
        # A bare `=` to a project property calls its Property Let, whatever the
        # value's type: `h.Item = New Collection` compiles with `Property Let
        # Item(ByVal v As Object)` (XLIDE issue #107). Only a property with a Set and
        # no Let refuses it: "Invalid use of property".
        if target.set_accessor and not target.let_accessor:
            push(
                "setRequired",
                f"Assignment to '{assignment.label}' requires Set: the property declares a Property "
                "Set and no Property Let. This is a VBE compile error: Invalid use of property.",
                assignment.member_span,
            )
            return
        if not target.let_accessor and is_known_object_assignment_type_ctx(expected, member_ctx):
            push(
                "setRequired",
                f"Object assignment to '{assignment.label}' requires Set because it expects {expected}.",
                assignment.member_span,
            )
            return
        if not expected or not is_known_scalar_type(normalize_type(expected) or ""):
            return  # a Let of an object or unknown type: nothing provable about the value
        string_arithmetic = nonnumeric_string_arithmetic_operand(
            expected, assignment.value_tokens, span.start
        )
        if string_arithmetic is not None:
            push(
                "stringArithmeticCoercion",
                f"Assignment to '{assignment.label}' expects {expected}, but this numeric expression "
                f"contains {string_arithmetic.label}. This will raise Run-time error '13': "
                "Type mismatch.",
                string_arithmetic.span,
            )
            return
        actual = infer_argument_type(
            assignment.value_tokens, span.start, env, module_signatures, source_names,
            resolve_expression_type, resolve_qualified_expression_type,
            source=source, member_ctx=member_ctx,
        )
        if actual is None:
            return
        reason = incompatibility_reason(expected, actual)
        if not reason:
            return
        push(
            "assignmentTypeMismatch",
            f"Assignment to '{assignment.label}' expects {expected}, but got {actual.label}. {reason}",
            actual.span,
        )

    # This rule reads a statement structurally - what precedes its first `=` is
    # the target - so it takes a single-line If's branches as statements of their
    # own. Read whole, `If ok Then w.Part = 1` had the target `If ok Then w.Part`,
    # and `If w.Part = 1 Then Exit Sub`, which assigns nothing, had `If w.Part`.
    def visit(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            check_statement(span)

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
    target_scalar_type = _known_scalar_assignment_target_type(
        name, expected_type, shapes, resolve_target_shape
    )
    if not target_scalar_type:
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
    # VBA special case (MS-VBAL Let-statement rules): a Byte array is directly
    # assignable to a String scalar - the idiomatic encoding-conversion pattern
    # (`s = bytes`). Only Byte element types are exempt; every other element
    # type remains a compile error.
    if target_scalar_type == "string" and normalize_type(source_shape.as_type) == "byte":
        return None
    return (source_name, Span(base_offset + tok.start, base_offset + tok.end))


def _known_scalar_assignment_target_type(
    name: str,
    expected_type: str,
    shapes: Mapping[str, DeclaredValueShape],
    resolve_shape: Callable[[str], SourceDeclaredShape],
) -> str | None:
    resolved = resolve_shape(name)
    target_shape = resolved.shape if resolved.resolved else shapes.get(name.lower())
    if target_shape is not None and target_shape.is_array:
        return None
    as_type = target_shape.as_type if target_shape is not None else None
    normalized = normalize_type(as_type if as_type is not None else expected_type)
    return normalized if normalized is not None and is_known_scalar_type(normalized) else None


def check_set_assignments(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """`Set x = ...` where x is a declared scalar requires an object variable; a Set to
    an object target of a provably-incompatible object type is reported too."""
    module_signatures = build_module_type_signatures(symbols)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        proc_sym = procedure_symbol_for(symbols, member)

        def resolve_expression_type(name: str) -> SourceDeclaredType:
            return declared_value_type_for_source_binding(symbols, proc_sym, project_visible_symbols, name)

        def resolve_qualified_expression_type(qualifier: str, name: str) -> SourceDeclaredType:
            return declared_value_type_for_qualified_source_binding(
                symbols, project_visible_symbols, qualifier, name
            )

        def visitor(stmt: LeafStatementNode) -> None:
            for branch in statement_and_branch_spans(stmt):
                check_set_span(branch)

        def check_set_span(branch: Span) -> None:
            target = set_assignment_target(source, branch)
            if target is None:
                return
            name, span, value_tokens = target
            target_declared_type = declared_type_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.ASSIGNMENT_TARGET
            )
            expected = target_declared_type.as_type if target_declared_type.resolved else env.get(name.lower())
            target_type = normalize_type(expected)
            # `Set v = 5` is refused whatever v is: a literal is never an object
            # reference ("Object required", XLIDE issue #125, measured in Excel 16.0).
            literal = [tok for tok in value_tokens if tok.kind is not TokenKind.COMMENT]
            if (
                (not target_type or target_type == "variant")
                and len(literal) == 1
                and _is_scalar_literal_token(literal[0])
            ):
                push(
                    "setRequiresObject",
                    f"Set assigns an object reference, but {literal[0].raw_text} is a literal value. "
                    "This is a VBE compile error: Object required.",
                    Span(branch.start + literal[0].start, branch.start + literal[0].end),
                )
                return
            if not target_type or not is_known_scalar_type(target_type):
                if not is_known_object_assignment_type_ctx(expected, member_ctx):
                    return
                actual = infer_argument_type(
                    value_tokens, branch.start, env, module_signatures, source_names,
                    resolve_expression_type, resolve_qualified_expression_type,
                    source=source, member_ctx=member_ctx,
                )
                reason = object_assignment_incompatibility_reason(expected, actual, member_ctx)
                if reason:
                    push(
                        "assignmentObjectTypeMismatch",
                        f"Object assignment to '{name}' expects {expected}, but got "
                        f"{actual.label if actual is not None else None}. {reason}",
                        actual.span if actual is not None else span,
                    )
                return
            push(
                "setRequiresObject",
                f"Set assignment requires an object variable, but '{name}' is declared as {expected}.",
                span,
            )

        return visitor

    return factory


_DECLARATION_ONLY_RE = re.compile(r"^(Dim|Const|Static|ReDim)\b", re.IGNORECASE)
_ERR_RAISE_RE = re.compile(r"\bErr\s*\.\s*Raise\b", re.IGNORECASE)
_ERROR_STATEMENT_RE = re.compile(r"^Error\s", re.IGNORECASE)


def _return_is_not_expected(
    source: str,
    proc: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    is_interface: bool,
) -> bool:
    """Whether a missing return assignment is deliberate rather than unfinished.

    An empty body is a stub where the module is a contract: a class that another
    module declares with `Implements` states its members for the implementer to
    fill in, so every one of them is empty on purpose. An empty Function anywhere
    else is unfinished code and still reports. A body that raises never returns
    normally, so it owes no value either.
    """
    executable = 0
    raises = False

    def visit(stmt: LeafStatementNode) -> None:
        nonlocal executable, raises
        text = source[stmt.span.start : stmt.span.end].strip()
        if not text or text.startswith("'") or _DECLARATION_ONLY_RE.match(text):
            return
        executable += 1
        if _ERR_RAISE_RE.search(text) or _ERROR_STATEMENT_RE.match(text):
            raises = True

    for_each_statement(proc.body, visit, activity)
    return (executable == 0 and is_interface) or raises


def check_missing_return_assignments(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    activity: ConditionalActivityTracker | None,
    module_name: str | None,
    implemented_interfaces: AbstractSet[str] | None,
    push: PushFn,
) -> None:
    """A Function/Property Get with no return assignment silently returns the default."""
    module_signatures = callable_type_signatures_for(symbols, project_procedures)
    is_interface = (
        module_name is not None
        and implemented_interfaces is not None
        and module_name.lower() in implemented_interfaces
    )
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        if member.proc_kind not in (ProcKind.FUNCTION, ProcKind.PROPERTY_GET):
            continue
        if not member.closed:
            continue
        if _procedure_has_return_assignment(source, member, activity, module_signatures):
            continue
        if _return_is_not_expected(source, member, activity, is_interface):
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

    def assigns_in(span: Span) -> bool:
        bare = bare_assignment_target(source, span)
        if bare is not None and bare[0].lower() == lower:
            return True
        set_target = set_assignment_target(source, span)
        if set_target is not None and set_target[0].lower() == lower:
            return True
        if _return_assigned_by_statement_form(source, span, lower):
            return True
        call = extract_call(source, span)
        qualified = None if call else extract_qualified_call(source, span, module_signatures)
        effective = call or qualified
        return effective is not None and _call_passes_name_to_by_ref_param(
            effective, lower, module_signatures
        )

    def visit(stmt: LeafStatementNode) -> None:
        nonlocal found
        if found:
            return
        # The branch spans cover a single-line If, whose statements the walk
        # itself does not reach (XLIDE issue #46).
        for span in statement_and_branch_spans(stmt):
            if assigns_in(span) or _assigns_own_field(source, span, lower):
                found = True
                return

    for_each_statement(proc.body, visit, activity)
    # `For Count3 = 1 To 3` assigns the return variable as its counter (XLIDE issue
    # #115): the loop is a block, not a statement the walk above visits.
    return found or _for_loop_assigns(proc.body, lower, activity)


def _for_loop_assigns(
    body: Sequence[BodyNode], lower: str, activity: ConditionalActivityTracker | None
) -> bool:
    return any(
        isinstance(node, ForBlockNode)
        and node.control_variable is not None
        and node.control_variable.lower() == lower
        for node in iter_body_nodes(body, inactive_node_skip(activity))
    )


_TYPE_SUFFIX_CHARS = frozenset("$%&!#@")


def _return_assigned_by_statement_form(source: str, span: Span, lower: str) -> bool:
    """The statement forms besides `Name = value` that assign a Function's return
    variable (XLIDE issue #115, each measured in Excel 16.0): `For Name = 1 To 3`
    leaves the counter's final value, `ReDim Name(2)` sizes an array return,
    `Line Input #f, Name`, `Input #f, Name` and `Get #f, 1, Name` read into it, and
    `Name$ = "hi"` names it with its type-declaration character."""
    toks = statement_tokens(source, span)
    i = first_executable_token_index(toks)

    def at(index: int) -> VbaToken | None:
        return toks[index] if 0 <= index < len(toks) else None

    def is_name(tok: VbaToken | None) -> bool:
        name = token_name(tok)
        return name is not None and name.lower() == lower

    head = token_text(at(i))
    if head == "for":
        return is_name(at(i + 2)) if token_text(at(i + 1)) == "each" else is_name(at(i + 1))
    if head == "redim":
        k = i + 1
        if token_text(at(k)) == "preserve":
            k += 1
        depth = 0
        while k < len(toks):
            raw = toks[k].raw_text
            if raw == "(":
                depth += 1
            elif raw == ")":
                depth -= 1
            elif (
                depth == 0
                and is_name(toks[k])
                and (k == i + 1 or token_text(toks[k - 1]) == "preserve" or toks[k - 1].raw_text == ",")
            ):
                return True
            k += 1
        return False
    if head in ("line", "input", "get"):
        # Everything after the file number is a target (Line Input / Input) or the
        # third slot is (Get #f, rec, var); a plain name in one of them is the
        # assignment.
        rest = toks[i + 1 :]
        return any(is_name(tok) and index > 0 and rest[index - 1].raw_text == "," for index, tok in enumerate(rest))
    # `Name$ = value`: the suffix is glued to the name and the `=` follows.
    name_tok = at(i)
    suffix = at(i + 1)
    equals = at(i + 2)
    return (
        is_name(name_tok)
        and name_tok is not None
        and suffix is not None
        and suffix.start == name_tok.end
        and len(suffix.raw_text) == 1
        and suffix.raw_text in _TYPE_SUFFIX_CHARS
        and equals is not None
        and equals.raw_text == "="
    )


def _is_scalar_literal_token(tok: VbaToken) -> bool:
    """A literal that can never be an object reference: a number, string, date, True or False."""
    if tok.kind in (
        TokenKind.INTEGER_LITERAL,
        TokenKind.FLOAT_LITERAL,
        TokenKind.STRING_LITERAL,
        TokenKind.DATE_LITERAL,
    ):
        return True
    word = token_text(tok)
    return word in ("true", "false")


def _assigns_own_field(source: str, span: Span, lower: str) -> bool:
    """True for `Name.Field = value`, which fills in a UDT or object return field by
    field.

    `MsToSystemTime.wYear = ...` IS the return assignment. Reading only a bare
    `Name =` counts every such function as never assigning anything, which is nine
    false positives in one real workbook alone.
    """
    toks = statement_tokens(source, span)
    i = first_executable_token_index(toks)
    if i < len(toks) and toks[i].raw_text.lower() == "let":
        i += 1
    name = toks[i] if i < len(toks) else None
    nxt = toks[i + 1] if i + 1 < len(toks) else None
    if name is None or name.raw_text.lower() != lower:
        return False
    if nxt is None or nxt.raw_text != ".":
        return False
    return any(
        tok.kind is TokenKind.OPERATOR and tok.raw_text == "=" for tok in toks[i + 2 :]
    )


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
