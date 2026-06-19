"""Rule family: control-flow rules (self-contained slice).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/controlFlow.ts: Exit
statement kinds, statement context (If/Then, Case/loop-exit/leading-dot
placement), For/Next control-variable matching, duplicate Case Else, Else without
If, and malformed leaf statements. The label rules (need flow/procedureLabels),
the For Each element/source type rule (type inference + host), and the
conditional-compilation branch-order rule (shared cc helper) are deferred.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, replace

from ...conditional import ConditionalActivity, ConditionalActivityTracker
from ...lexer.token_kinds import TokenKind
from ...lexer.tokenize import tokenize
from ...parser.nodes import (
    AssignmentNode,
    BodyNode,
    CallNode,
    ConditionalDirectiveNode,
    DoBlockNode,
    ForBlockNode,
    IfBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    ProcKind,
    SelectBlockNode,
    Span,
    StatementNode,
    WhileBlockNode,
    WithBlockNode,
    is_leaf_statement,
)
from ..context import PushFn
from ..walker import (
    ProcedureStatementVisitor,
    absolute_span,
    active_module_members,
    for_each_statement,
    is_inactive_node,
    statement_tokens,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .shared import scan_conditional_compilation_branch_order

# -- checkExitStatements ---------------------------------------------------


def _expected_exit_word(kind: ProcKind) -> str:
    if kind is ProcKind.SUB:
        return "Sub"
    if kind is ProcKind.FUNCTION:
        return "Function"
    return "Property"


def _enclosing_proc_label(kind: ProcKind) -> str:
    if kind is ProcKind.SUB:
        return "Sub"
    if kind is ProcKind.FUNCTION:
        return "Function"
    return "Property procedure"


def _exit_target(source: str, span: Span) -> tuple[str, Span] | None:
    toks = statement_tokens_after_leading_label(source, span)
    if len(toks) < 2 or toks[0].raw_text.lower() != "exit":
        return None
    w = toks[1].raw_text.lower()
    if w == "sub":
        word = "Sub"
    elif w == "function":
        word = "Function"
    elif w == "property":
        word = "Property"
    else:
        return None  # Exit Do / Exit For etc.
    return (word, Span(span.start + toks[0].start, span.start + toks[1].end))


def check_exit_statements(source: str, push: PushFn) -> ProcedureStatementVisitor:
    """Exit Sub/Function/Property must match the enclosing procedure kind."""

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        expected = _expected_exit_word(member.proc_kind)
        label = _enclosing_proc_label(member.proc_kind)

        def visitor(stmt: LeafStatementNode) -> None:
            hit = _exit_target(source, stmt.span)
            if hit is not None and hit[0] != expected:
                word, hit_span = hit
                push(
                    "exitWrongProcedure",
                    f"'Exit {word}' is not valid inside a {label}; use 'Exit {expected}'.",
                    hit_span,
                )

        return visitor

    return factory


# -- checkStatementContext -------------------------------------------------


@dataclass(frozen=True, slots=True)
class _StatementContext:
    for_depth: int = 0
    do_depth: int = 0
    with_depth: int = 0
    select_depth: int = 0


def _exit_phrase_span(base: Span, first_start: int, target_end: int) -> Span:
    return Span(base.start + first_start, base.start + target_end)


def _check_context_statement(
    source: str, stmt: LeafStatementNode, ctx: _StatementContext, push: PushFn
) -> None:
    toks = statement_tokens_after_leading_label(source, stmt.span)
    if not toks:
        return
    first = toks[0]
    w0 = token_text(first)

    if w0 == "if" and not any(token_text(t) == "then" for t in toks):
        push("ifMissingThen", "If statement is missing 'Then'.", absolute_span(stmt.span, first))

    if w0 == "case" and ctx.select_depth == 0:
        push(
            "caseOutsideSelect",
            "'Case' can only appear inside a 'Select Case' block.",
            absolute_span(stmt.span, first),
        )

    if first.raw_text == "." and ctx.with_depth == 0:
        push(
            "memberAccessOutsideWith",
            "A statement that starts with '.' must be inside a With block.",
            absolute_span(stmt.span, first),
        )

    leading_member = toks[1] if len(toks) > 1 else None
    if first.raw_text == "." and ctx.with_depth > 0 and (leading_member is None or token_name(leading_member) is None):
        push(
            "invalidExpressionSyntax",
            "Incomplete member access: type a member name after '.'.",
            absolute_span(stmt.span, first),
        )

    if w0 == "exit":
        target = toks[1] if len(toks) > 1 else None
        target_word = token_text(target)
        if target is not None and target_word == "for" and ctx.for_depth == 0:
            push(
                "exitOutsideBlock",
                "'Exit For' can only appear inside a For loop.",
                _exit_phrase_span(stmt.span, first.start, target.end),
            )
        elif target is not None and target_word == "do" and ctx.do_depth == 0:
            push(
                "exitOutsideBlock",
                "'Exit Do' can only appear inside a Do loop.",
                _exit_phrase_span(stmt.span, first.start, target.end),
            )


def _check_for_next_control_variable(
    node: ForBlockNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    if (
        not node.control_variable
        or node.control_variable_span is None
        or not node.next_variable
        or node.next_variable_span is None
        or (activity is not None and activity.is_inactive(node.control_variable_span))
        or (activity is not None and activity.is_inactive(node.next_variable_span))
    ):
        return
    if node.control_variable.lower() == node.next_variable.lower():
        return
    push(
        "nextVariableMismatch",
        f"Next variable '{node.next_variable}' does not match active For control variable "
        f"'{node.control_variable}'.",
        node.next_variable_span,
    )


def _check_context_body(
    source: str,
    body: list[BodyNode],
    ctx: _StatementContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    for node in body:
        if is_inactive_node(activity, node):
            continue
        if isinstance(node, (StatementNode, AssignmentNode, CallNode)):
            _check_context_statement(source, node, ctx, push)
        elif isinstance(node, ForBlockNode):
            _check_for_next_control_variable(node, activity, push)
            _check_context_body(source, node.body, replace(ctx, for_depth=ctx.for_depth + 1), activity, push)
        elif isinstance(node, DoBlockNode):
            _check_context_body(source, node.body, replace(ctx, do_depth=ctx.do_depth + 1), activity, push)
        elif isinstance(node, WithBlockNode):
            _check_context_body(source, node.body, replace(ctx, with_depth=ctx.with_depth + 1), activity, push)
        elif isinstance(node, SelectBlockNode):
            _check_context_body(source, node.body, replace(ctx, select_depth=ctx.select_depth + 1), activity, push)
        elif isinstance(node, (IfBlockNode, WhileBlockNode)):
            _check_context_body(source, node.body, ctx, activity, push)
        # ConditionalDirective / VariableGroup: no context check.


# -- checkElseBranchOrder --------------------------------------------------


def check_else_branch_order(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    _check_conditional_compilation_else_branch_order(source, mod, push)
    _check_if_block_else_branch_order(source, mod, activity, push)


def _check_conditional_compilation_else_branch_order(
    source: str, mod: ModuleNode, push: PushFn
) -> None:
    # Conditional-compilation directives are checked structurally, regardless of
    # branch activity (a malformed #If block fails to compile either way).
    for issue in scan_conditional_compilation_branch_order(mod).issues:
        if issue.kind == "elseifAfterElse":
            message = (
                "'#ElseIf' cannot appear after '#Else' in the same conditional-compilation block."
            )
        else:
            message = "Only one '#Else' branch is allowed in a conditional-compilation block."
        push("elseBranchOrder", message, _conditional_directive_keyword_span(source, issue.directive))


def _check_if_block_else_branch_order(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_if_block_else_branch_order_in_body(source, member.body, activity, push)


def _check_if_block_else_branch_order_in_body(
    source: str, body: list[BodyNode], activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    for node in body:
        if is_inactive_node(activity, node):
            continue
        if isinstance(node, IfBlockNode):
            _check_single_if_block_else_branch_order(source, node, activity, push)
        child = getattr(node, "body", None)
        if isinstance(child, list):
            _check_if_block_else_branch_order_in_body(source, child, activity, push)


def _check_single_if_block_else_branch_order(
    source: str, node: IfBlockNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    seen_else = False
    for child in node.body:
        if is_inactive_node(activity, child) or not is_leaf_statement(child):
            continue
        toks = statement_tokens_after_leading_label(source, child.span)
        if not toks:
            continue
        word = token_text(toks[0])
        if word == "elseif" and seen_else:
            push(
                "elseBranchOrder",
                "'ElseIf' cannot appear after 'Else' in the same If block.",
                absolute_span(child.span, toks[0]),
            )
        elif word == "else":
            if seen_else:
                push(
                    "elseBranchOrder",
                    "Only one 'Else' branch is allowed in an If block.",
                    absolute_span(child.span, toks[0]),
                )
            seen_else = True


def _conditional_directive_keyword_span(source: str, directive: ConditionalDirectiveNode) -> Span:
    tokens = [
        tok
        for tok in tokenize(source[directive.span.start : directive.span.end])
        if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE
    ]
    marker = tokens[0] if tokens else None
    keyword = tokens[1] if len(tokens) > 1 else None
    if marker is not None and marker.kind is TokenKind.DIRECTIVE and keyword is not None:
        return Span(directive.span.start + marker.start, directive.span.start + keyword.end)
    return directive.span


def check_statement_context(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_context_body(source, member.body, _StatementContext(), activity, push)


# -- checkDuplicateCaseElse ------------------------------------------------


def _for_each_select_block(
    body: list[BodyNode],
    activity: ConditionalActivityTracker | None,
    visit: Callable[[SelectBlockNode], None],
) -> None:
    for node in body:
        if is_inactive_node(activity, node):
            continue
        if isinstance(node, SelectBlockNode):
            visit(node)
        child = getattr(node, "body", None)
        if isinstance(child, list):
            _for_each_select_block(child, activity, visit)


def check_duplicate_case_else(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    def visit_select(select: SelectBlockNode) -> None:
        seen_case_else = False
        for node in select.body:
            if not is_leaf_statement(node) or (
                activity is not None and activity.activity_for_span(node.span) is not ConditionalActivity.ACTIVE
            ):
                continue
            toks = statement_tokens_after_leading_label(source, node.span)
            case_tok = toks[0] if toks else None
            else_tok = toks[1] if len(toks) > 1 else None
            if case_tok is None or else_tok is None or token_text(case_tok) != "case" or token_text(else_tok) != "else":
                continue
            if seen_case_else:
                push(
                    "duplicateCaseElse",
                    "A 'Select Case' block can have only one 'Case Else'.",
                    Span(absolute_span(node.span, case_tok).start, absolute_span(node.span, else_tok).end),
                )
            else:
                seen_case_else = True

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _for_each_select_block(member.body, activity, visit_select)


# -- checkElseWithoutIf ----------------------------------------------------


def check_else_without_if(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    def visit(body: list[BodyNode], inside_if_block: bool) -> None:
        for node in body:
            if is_inactive_node(activity, node):
                continue
            if is_leaf_statement(node):
                if inside_if_block:
                    continue
                toks = statement_tokens_after_leading_label(source, node.span)
                first = toks[0] if toks else None
                word = token_text(first) if first is not None else ""
                if first is not None and (word == "else" or word == "elseif"):
                    push(
                        "elseWithoutIf",
                        f"'{first.raw_text}' can only appear inside an 'If' block.",
                        absolute_span(node.span, first),
                    )
                continue
            child = getattr(node, "body", None)
            if isinstance(child, list):
                visit(child, isinstance(node, IfBlockNode))

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            visit(member.body, False)


# -- checkMalformedStatements ----------------------------------------------


def check_malformed_statements(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    def inspect(stmt: LeafStatementNode) -> None:
        toks = statement_tokens(source, stmt.span)
        if not toks:
            return
        first = toks[0]
        second = toks[1] if len(toks) > 1 else None
        if (
            second is not None
            and second.raw_text == "="
            and first.kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL, TokenKind.STRING_LITERAL)
        ):
            push("invalidAssignmentTarget", "Cannot assign to a literal value.", absolute_span(stmt.span, first))
        if token_text(first) == "open" and not any(token_text(t) == "for" for t in toks):
            push(
                "openMissingForMode",
                "An 'Open' statement requires a 'For <mode>' clause.",
                absolute_span(stmt.span, first),
            )

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            for_each_statement(member.body, inspect, activity)
