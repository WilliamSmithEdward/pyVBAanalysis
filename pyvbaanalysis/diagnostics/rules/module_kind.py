"""Rule family: module-kind constraints (self-contained slice).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/moduleKind.ts: object-module
Public restrictions, Event/WithEvents/Friend/Implements placement, RaiseEvent
targets, Declare PtrSafe for Win64, and event-handler module scope (the last reads
the vendored event catalogue via completion.event_handlers).
"""

from __future__ import annotations

import re
from collections.abc import Callable, Sequence

from ...completion.event_handlers import (
    EventHandlerDocumentType,
    event_handler_document_type_for_context,
    event_handler_procedure_for_name,
)
from ...conditional import (
    ConditionalActivityTracker,
    ConditionalCompilationEnvironment,
    ConditionalValue,
    conditional_compiler_constants,
)
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    DeclareNode,
    EventNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    ProcKind,
    Span,
    StatementNode,
    TypeNode,
    VariableGroupNode,
)
from ...symbols.symbol_model import ModuleSymbolKind
from ..context import PushFn, is_object_module_kind
from ..walker import (
    ProcedureStatementVisitor,
    absolute_span,
    active_module_members,
    declared_name_span,
    first_token_span,
    for_each_procedure_body_line,
    for_each_statement,
    for_each_variable_group,
    is_inactive_node,
    statement_tokens,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

_DECIMAL_RE = re.compile(r"^\d+$")


def _is_public_modifier(value: str | None) -> bool:
    return value is not None and value.lower() == "public"


# -- checkObjectModulePublicMembers ----------------------------------------


def check_object_module_public_members(source: str, mod: ModuleNode, module_kind: ModuleSymbolKind, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    if not is_object_module_kind(module_kind):
        return

    def report(kind: str, span: Span) -> None:
        push(
            "objectModulePublicMember",
            f"Public {kind} are not allowed as Public members of object modules; "
            "VBE Compile rejects this declaration.",
            span,
        )

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode) and _is_public_modifier(member.modifier):
            for decl in member.declarations:
                span = declared_name_span(source, decl.span, decl.name)
                if member.is_const:
                    report("constants", span)
                elif decl.is_array:
                    report("arrays", span)
                elif decl.fixed_length is not None:
                    report("fixed-length strings", span)
            continue
        if isinstance(member, TypeNode) and _is_public_modifier(member.visibility):
            report("user-defined types", declared_name_span(source, member.span, member.name))
            continue
        if isinstance(member, DeclareNode) and _is_public_modifier(member.visibility):
            report("Declare statements", declared_name_span(source, member.span, member.name))


# -- checkEventDeclarationModuleKind ---------------------------------------


def check_event_declaration_module_kind(source: str, mod: ModuleNode, module_kind: ModuleSymbolKind, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    if is_object_module_kind(module_kind):
        return
    for member in active_module_members(mod, activity):
        if not isinstance(member, EventNode):
            continue
        push(
            "eventDeclarationModuleKind",
            f"Event declaration '{member.name}' is only valid in class, document, or UserForm modules.",
            declared_name_span(source, member.span, member.name),
        )


# -- checkMeOutsideObjectModule (per-statement) ----------------------------


def check_me_outside_object_module(module_kind: ModuleSymbolKind, source: str, push: PushFn) -> ProcedureStatementVisitor:
    if is_object_module_kind(module_kind):
        def skip(member: ProcedureNode) -> None:
            return None

        return skip

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None]:
        def visitor(stmt: LeafStatementNode) -> None:
            toks = statement_tokens(source, stmt.span)
            for i, tok in enumerate(toks):
                if token_text(tok) != "me":
                    continue
                if i > 0 and toks[i - 1].raw_text == ".":
                    continue  # a member named Me, not the Me keyword
                push(
                    "meOutsideObjectModule",
                    "'Me' is only valid in a class, document, or UserForm module.",
                    absolute_span(stmt.span, tok),
                )

        return visitor

    return factory


# -- checkWithEventsDeclarations -------------------------------------------


def check_with_events_declarations(source: str, mod: ModuleNode, module_kind: ModuleSymbolKind, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    def inspect(group: VariableGroupNode, inside_procedure: bool) -> None:
        if not group.with_events or is_inactive_node(activity, group):
            return
        for decl in group.declarations:
            name_span = declared_name_span(source, decl.span, decl.name)
            if inside_procedure:
                push("withEventsDeclaration", f"WithEvents variable '{decl.name}' must be declared at module level.", name_span)
                continue
            if not is_object_module_kind(module_kind):
                push("withEventsDeclaration", f"WithEvents variable '{decl.name}' is only valid in class, document, or UserForm modules.", name_span)
                continue
            if decl.is_new:
                push("withEventsDeclaration", f"WithEvents variable '{decl.name}' cannot be declared As New.", name_span)
            if decl.is_array:
                push("withEventsDeclaration", f"WithEvents variable '{decl.name}' cannot be an array.", name_span)

    def inspect_in_procedure(group: VariableGroupNode) -> None:
        inspect(group, True)

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            inspect(member, False)
        elif isinstance(member, ProcedureNode):
            for_each_variable_group(member.body, inspect_in_procedure, activity)


# -- checkFriendDeclarations -----------------------------------------------


def _has_friend_modifier(modifiers: Sequence[str]) -> bool:
    return any(m.lower() == "friend" for m in modifiers)


def _friend_keyword_span(source: str, span: Span) -> Span:
    for tok in statement_tokens_after_leading_label(source, span):
        if token_text(tok) == "friend":
            return absolute_span(span, tok)
    return first_token_span(source, span)


def check_friend_declarations(source: str, mod: ModuleNode, module_kind: ModuleSymbolKind, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            if _has_friend_modifier(member.modifiers) and not is_object_module_kind(module_kind):
                push(
                    "friendDeclaration",
                    f"Friend procedure '{member.name}' is only valid in class, document, or UserForm modules.",
                    _friend_keyword_span(source, member.span),
                )
            continue
        if not isinstance(member, VariableGroupNode) or member.modifier.lower() != "friend":
            continue
        push(
            "friendDeclaration",
            "Friend can only modify procedure declarations, not variables.",
            _friend_keyword_span(source, member.span),
        )


# -- checkImplementsStatementPlacement -------------------------------------


def _implements_statement_hit(source: str, span: Span) -> tuple[str, Span] | None:
    toks = statement_tokens_after_leading_label(source, span)
    if not toks or token_text(toks[0]) != "implements":
        return None
    first_name = token_name(toks[1]) if len(toks) > 1 else None
    if not first_name:
        return None
    name = first_name
    end_index = 1
    while True:
        dot = toks[end_index + 1] if end_index + 1 < len(toks) else None
        if dot is None or dot.raw_text != ".":
            break
        part = token_name(toks[end_index + 2]) if end_index + 2 < len(toks) else None
        if not part:
            break
        name += f".{part}"
        end_index += 2
    return (name, Span(span.start + toks[1].start, span.start + toks[end_index].end))


def check_implements_statement_placement(source: str, mod: ModuleNode, module_kind: ModuleSymbolKind, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    procedure_seen = False

    def report_procedure_placement(name: str, span: Span) -> None:
        push(
            "implementsStatementPlacement",
            f"Implements statement '{name}' must appear in the module declaration section before any procedure.",
            span,
        )

    def inspect_body_statement(stmt: LeafStatementNode) -> None:
        hit = _implements_statement_hit(source, stmt.span)
        if hit is not None:
            report_procedure_placement(hit[0], hit[1])

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            procedure_seen = True
            for_each_statement(member.body, inspect_body_statement, activity)
            continue
        if not isinstance(member, StatementNode):
            continue
        hit = _implements_statement_hit(source, member.span)
        if hit is None:
            continue
        name, span = hit
        if not is_object_module_kind(module_kind):
            push(
                "implementsStatementPlacement",
                f"Implements statement '{name}' is only valid in class, document, or UserForm modules.",
                span,
            )
            continue
        if procedure_seen:
            report_procedure_placement(name, span)


# -- checkRaiseEventTargets ------------------------------------------------


def _raise_event_statement_start_index(toks: Sequence[VbaToken]) -> int:
    start = 0
    if len(toks) > 1 and _DECIMAL_RE.match(toks[0].raw_text):
        start = 1
    elif (
        len(toks) > 2
        and (toks[0].kind is TokenKind.IDENTIFIER or toks[0].kind is TokenKind.KEYWORD)
        and toks[1].raw_text == ":"
    ):
        start = 2
    return start if (start < len(toks) and token_text(toks[start]) == "raiseevent") else -1


def _raise_event_target_hit(source: str, span: Span) -> tuple[str, Span] | None:
    toks = statement_tokens(source, span)
    start = _raise_event_statement_start_index(toks)
    if start < 0:
        return None
    name_tok = toks[start + 1] if start + 1 < len(toks) else None
    if name_tok is None:
        return None
    name = token_name(name_tok)
    if not name:
        return None
    return (name, Span(span.start + name_tok.start, span.start + name_tok.end))


def check_raise_event_targets(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    events: set[str] = set()
    for member in active_module_members(mod, activity):
        if isinstance(member, EventNode) and member.name:
            events.add(member.name.lower())

    def on_line(line_span: Span) -> None:
        if activity is not None and activity.is_inactive(line_span):
            return
        hit = _raise_event_target_hit(source, line_span)
        if hit is None or hit[0].lower() in events:
            return
        push(
            "raiseEventUndeclaredEvent",
            f"Event '{hit[0]}' is not declared in this module, so it cannot be raised with RaiseEvent.",
            hit[1],
        )

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            for_each_procedure_body_line(source, member, on_line)


# -- checkDeclarePtrSafeForWin64 -------------------------------------------


def _conditional_value_truthy(value: ConditionalValue | None) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    return isinstance(value, str) and len(value) > 0


def check_declare_ptr_safe_for_win64(source: str, mod: ModuleNode, conditional_compilation: ConditionalCompilationEnvironment | None, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    constants = conditional_compiler_constants(conditional_compilation)
    if not _conditional_value_truthy(constants.get("win64")):
        return
    for member in active_module_members(mod, activity):
        if not isinstance(member, DeclareNode) or member.ptr_safe:
            continue
        push(
            "declareMissingPtrSafe",
            f"Declare statement '{member.name}' must include PtrSafe when compiling for 64-bit Office.",
            declared_name_span(source, member.span, member.name),
        )


# -- checkEventHandlerModuleScope ------------------------------------------


def _describe_event_document_type(document_type: EventHandlerDocumentType | None) -> str:
    if document_type in ("workbook", "worksheet", "chart"):
        return document_type
    return "unknown"


def check_event_handler_module_scope(
    source: str,
    mod: ModuleNode,
    module_name: str,
    module_kind: ModuleSymbolKind,
    document_type: EventHandlerDocumentType | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """A Sub whose name matches an Excel event handler that this module's document
    type does not wire. Port of checkEventHandlerModuleScope: pure AST + the vendored
    event catalogue + module kind; no binder or host surface. A Sub named like an
    event in the wrong module (or any non-document module) behaves as an ordinary
    procedure, never as the wired event.
    """
    actual_document_type = event_handler_document_type_for_context(
        module_name, module_kind, document_type
    )
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode) or member.proc_kind is not ProcKind.SUB:
            continue
        event = event_handler_procedure_for_name(member.name)
        if event is None:
            continue
        if actual_document_type == event.document_type:
            continue
        module_description = (
            f"{_describe_event_document_type(actual_document_type)} document module"
            if module_kind is ModuleSymbolKind.DOCUMENT
            else f"{module_kind.value} module"
        )
        push(
            "eventHandlerWrongModule",
            f"'{event.name}' matches a {event.owner} event handler, but this "
            f"{module_description} is not where Excel wires that event. "
            "It will behave like an ordinary procedure here.",
            declared_name_span(source, member.span, member.name),
        )
