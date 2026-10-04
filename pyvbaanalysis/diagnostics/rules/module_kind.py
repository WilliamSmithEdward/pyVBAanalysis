"""Rule family: module-kind constraints (self-contained slice).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/moduleKind.ts: object-module
Public restrictions, Event/WithEvents/Friend/Implements placement, RaiseEvent
targets, Declare PtrSafe for Win64, and event-handler module scope (the last reads
the vendored event catalogue via completion.event_handlers).
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence

from ...completion.event_handlers import (
    EventHandlerDocumentType,
    event_handler_document_type_for_context,
    event_handler_procedure_for_name,
)
from ...conditional import (
    ConditionalActivityTracker,
    ConditionalCompilationEnvironment,
)
from ...conditional.conditional_compilation import compiler_constants_with_defaults
from ...constants.integer_constant_expression import parse_vba_integer_literal
from ...js_compat import js_number_to_string, js_trim
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    DeclareNode,
    EventNode,
    LeafStatementNode,
    ModuleNode,
    ParameterNode,
    ProcedureNode,
    ProcKind,
    Span,
    StatementNode,
    TypeNode,
    VariableGroupNode,
)
from ...symbols.symbol_model import ModuleSymbolKind, ModuleSymbols, VbaProjectClassMembers
from ...types.type_inference import type_environment_for
from ...types.type_names import is_known_scalar_type, normalize_type
from ..call_extraction import string_literal_value
from ..context import PushFn, is_object_module_kind
from ..string_conversion import is_invalid_numeric_string
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

    def report(kind: str, span: Span, what: str | None = None) -> None:
        push(
            "objectModulePublicMember",
            f"{what if what is not None else f'Public {kind}'} are not allowed as Public members "
            "of object modules; VBE Compile rejects this declaration.",
            span,
        )

    # A Public variable of the module's own Private Type is refused the same way,
    # and so is any Global (issue #212, measured in a class module).
    private_types: set[str] = set()
    for member in active_module_members(mod, activity):
        if isinstance(member, TypeNode) and member.visibility is not None and member.visibility.lower() == "private":
            private_types.add(member.name.lower())

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode) and member.modifier.lower() == "global":
            for decl in member.declarations:
                report("variables", declared_name_span(source, decl.span, decl.name), "Global variables")
            continue
        if isinstance(member, VariableGroupNode) and _is_public_modifier(member.modifier):
            for decl in member.declarations:
                span = declared_name_span(source, decl.span, decl.name)
                if member.is_const:
                    report("constants", span)
                elif decl.is_array:
                    report("arrays", span)
                elif decl.fixed_length is not None:
                    report("fixed-length strings", span)
                elif decl.as_type is not None and js_trim(decl.as_type).lower() in private_types:
                    report("user-defined types", span)
            continue
        # With no scope keyword a Type or a Declare is Public, and refused the same
        # way (issue #490, measured in Excel 16.0).
        if isinstance(member, TypeNode) and (not member.visibility or _is_public_modifier(member.visibility)):
            report("user-defined types", declared_name_span(source, member.span, member.name))
            continue
        if isinstance(member, DeclareNode) and (
            not member.visibility or _is_public_modifier(member.visibility)
        ):
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


def check_with_events_declarations(
    source: str,
    mod: ModuleNode,
    module_kind: ModuleSymbolKind,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    project_classes: Sequence[VbaProjectClassMembers] = (),
) -> None:
    # A project class whose member list is complete and holds no Event:
    # `WithEvents s As Plain` (issue #445, measured in Excel 16.0).
    def class_without_events(lower: str) -> bool:
        found = next((candidate for candidate in project_classes if candidate.name.lower() == lower), None)
        return (
            found is not None
            and found.kind == "class"
            and found.exhaustive is True
            and not any(member.kind == "event" for member in found.members)
        )

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
            # The type must be a class that sources events. `As Object` is refused
            # outright ("Expected: identifier") and `As Collection`, which has no
            # events, with "Object does not source automation events" (XLIDE issue
            # #124, measured in Excel 16.0). An intrinsic type is no object at all.
            # A host or project class is not judged here.
            normalized = normalize_type(decl.as_type)
            if normalized is None or normalized in ("object", "variant"):
                named = f"'{decl.as_type}'" if decl.as_type else "no type"
                push(
                    "withEventsDeclaration",
                    f"WithEvents variable '{decl.name}' must be declared As a specific class that raises events; {named} names none.",
                    name_span,
                )
            elif (
                normalized == "collection"
                or is_known_scalar_type(normalized)
                or class_without_events(normalized)
            ):
                push(
                    "withEventsDeclaration",
                    f"WithEvents variable '{decl.name}' is declared As {decl.as_type}, which does not source automation events.",
                    name_span,
                )

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
    # Procedures that precede the Implements under test AND could be compiled
    # beside it; a procedure in another arm of a `#If` chain never reaches the
    # compiler with it (XLIDE issue #58).
    procedures_above: list[Span] = []

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
            procedures_above.append(member.span)
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
        compiled_together = any(
            activity is None or not activity.mutually_exclusive(prior, member.span)
            for prior in procedures_above
        )
        if compiled_together:
            report_procedure_placement(name, span)


# -- checkRaiseEventTargets ------------------------------------------------


def _statement_segment_starts(toks: Sequence[VbaToken]) -> list[int]:
    """Indices where each ``:``-separated statement segment begins on a logical line.

    The first segment skips a leading line-number or ``Label:`` prefix; subsequent
    segments begin right after each top-level statement-separator colon. A colon
    inside parentheses/brackets is not a separator."""
    if not toks:
        return []
    starts: list[int] = []
    depth = 0
    segment_start = 0
    # Skip a leading line-number or `Label:` on the first segment.
    if len(toks) > 1 and _DECIMAL_RE.match(toks[0].raw_text):
        segment_start = 1
    elif (
        len(toks) > 2
        and (toks[0].kind is TokenKind.IDENTIFIER or toks[0].kind is TokenKind.KEYWORD)
        and toks[1].raw_text == ":"
    ):
        segment_start = 2
    starts.append(segment_start)
    for i in range(segment_start, len(toks)):
        raw = toks[i].raw_text
        if raw in ("(", "["):
            depth += 1
        elif raw in (")", "]"):
            depth -= 1
        elif raw == ":" and depth == 0 and i + 1 < len(toks):
            starts.append(i + 1)
    return starts


def _raise_event_target_hits(source: str, span: Span) -> list[tuple[str, Span, int | None]]:
    toks = statement_tokens(source, span)
    hits: list[tuple[str, Span, int | None]] = []
    # Each `:`-separated statement segment on the line may be its own RaiseEvent.
    for segment_start in _statement_segment_starts(toks):
        if segment_start >= len(toks) or token_text(toks[segment_start]) != "raiseevent":
            continue
        name_tok = toks[segment_start + 1] if segment_start + 1 < len(toks) else None
        name = token_name(name_tok) if name_tok is not None else None
        if not name or name_tok is None:
            continue
        hits.append(
            (
                name,
                Span(span.start + name_tok.start, span.start + name_tok.end),
                _raise_event_argument_count(toks, segment_start + 2),
            )
        )
    return hits


def _raise_event_argument_count(toks: Sequence[VbaToken], at: int) -> int | None:
    """How many arguments a RaiseEvent passes, from the token after the event's
    name: none when the segment ends there, else one per top-level comma group
    inside the parentheses. None when the list is not simply that."""
    following = toks[at] if at < len(toks) else None
    if following is None or following.kind in (TokenKind.COMMENT, TokenKind.COLON, TokenKind.NEWLINE):
        return 0
    if following.raw_text != "(":
        return None
    depth = 0
    count = 0
    saw_token = False
    for i in range(at, len(toks)):
        raw = toks[i].raw_text
        if raw == "(":
            depth += 1
            if depth == 1:
                continue
        elif raw == ")":
            depth -= 1
            if depth == 0:
                return count + 1 if saw_token else 0
        elif depth == 1 and raw == ",":
            count += 1
            continue
        if depth >= 1 and toks[i].kind is not TokenKind.COMMENT:
            saw_token = True
    return None


_RAISE_EVENT_RE = re.compile("raiseevent", re.IGNORECASE | re.ASCII)


def check_raise_event_targets(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    """`RaiseEvent` names an Event declared by the containing module, and passes it
    one argument per parameter. Measured in Excel 16.0 (issue #213): `RaiseEvent
    Changed(1, 2)` for `Event Changed(ByVal v As Long)` is "Wrong number of
    arguments or invalid property assignment", and `RaiseEvent Ev` for `Event
    Ev(ByVal a As Long)` is "Argument not optional". An Event takes no Optional or
    ParamArray parameter, so the count is exact."""
    # The scan below lexes every physical line of every procedure, which is wasted
    # on the great majority of modules that never raise an event (XLIDE issue
    # #139). A hit in a comment or string only means the scan runs.
    if _RAISE_EVENT_RE.search(source) is None:
        return
    events: dict[str, int | None] = {}
    for member in active_module_members(mod, activity):
        if isinstance(member, EventNode) and member.name:
            key = member.name.lower()
            exact = all(not param.optional and not param.param_array for param in member.params)
            events[key] = None if key in events or not exact else len(member.params)

    def on_line(line_span: Span) -> None:
        if activity is not None and activity.is_inactive(line_span):
            return
        # A single physical line can carry several `:`-separated statements
        # (e.g. `RaiseEvent A: RaiseEvent B`), so check every RaiseEvent on it.
        for name, hit_span, argument_count in _raise_event_target_hits(source, line_span):
            key = name.lower()
            if key in events:
                expected = events[key]
                if expected is not None and argument_count is not None and argument_count != expected:
                    plural = "" if expected == 1 else "s"
                    error = (
                        "Wrong number of arguments or invalid property assignment"
                        if argument_count > expected
                        else "Argument not optional"
                    )
                    push(
                        "raiseEventArgumentCount",
                        f"Event '{name}' takes {expected} argument{plural}, and RaiseEvent passes "
                        f"{argument_count}. This is a VBE compile error: {error}.",
                        hit_span,
                    )
                continue
            push(
                "raiseEventUndeclaredEvent",
                f"Event '{name}' is not declared in this module, so it cannot be raised with RaiseEvent.",
                hit_span,
            )

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            for_each_procedure_body_line(source, member, on_line)


# The whole-number ranges a ByVal Event parameter converts a literal to.
_WHOLE_RANGES: dict[str, tuple[int, int]] = {
    "byte": (0, 255),
    "integer": (-32768, 32767),
    "long": (-2147483648, 2147483647),
}


def check_raise_event_arguments(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """What a RaiseEvent passes its Event's parameters (issue #266, measured in
    Excel 16.0):

    - a named argument, `RaiseEvent Done(n:=1)`, is "Syntax error";
    - a variable of another type for a ByRef parameter, an Integer or a String for
      `n As Long`, is "ByRef argument type mismatch" (a literal, or the variable in
      parentheses, is a copy and compiles);
    - for a ByVal parameter of a number type, a string that is no number raises 13
      and a whole number outside the type raises 6 at run time.
    """
    if _RAISE_EVENT_RE.search(source) is None:
        return
    events: dict[str, list[ParameterNode] | None] = {}
    for member in active_module_members(mod, activity):
        if isinstance(member, EventNode) and member.name:
            key = member.name.lower()
            events[key] = None if key in events else list(member.params)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        env = type_environment_for(symbols, member)

        def on_line(line_span: Span, env: Mapping[str, str] = env) -> None:
            if activity is not None and activity.is_inactive(line_span):
                return
            toks = statement_tokens(source, line_span)
            for start in _statement_segment_starts(toks):
                params: list[ParameterNode] | None = None
                if start < len(toks) and token_text(toks[start]) == "raiseevent":
                    event_tok = toks[start + 1] if start + 1 < len(toks) else None
                    event_name_lower = token_name(event_tok)
                    params = events.get(event_name_lower.lower() if event_name_lower is not None else "")
                if params is None or start + 2 >= len(toks) or toks[start + 2].raw_text != "(":
                    continue
                close = match_paren_from(toks, start + 2)
                slots = (
                    []
                    if close < 0
                    else split_top_level_token_groups(
                        [tok for tok in toks[start + 3 : close] if tok.kind is not TokenKind.COMMENT], 0, ","
                    )
                )
                event_name = toks[start + 1].raw_text
                for k, slot in enumerate(slots):
                    _check_raise_event_slot(
                        params[k] if k < len(params) else None, slot, event_name, line_span, env, push
                    )

        for_each_procedure_body_line(source, member, on_line)


def _check_raise_event_slot(
    param: ParameterNode | None,
    slot: Sequence[VbaToken],
    event_name: str,
    line_span: Span,
    env: Mapping[str, str],
    push: PushFn,
) -> None:
    """One argument slot of checkRaiseEventArguments (upstream's slots.forEach body)."""
    if param is None or len(slot) == 0:
        return
    span = Span(line_span.start + slot[0].start, line_span.start + slot[-1].end)
    if len(slot) > 1 and slot[1].raw_text == ":=":
        push(
            "malformedStatement",
            f"RaiseEvent passes its arguments by position; '{slot[0].raw_text}:=' names one. This "
            "is a VBE compile error: Syntax error.",
            span,
        )
        return
    expected_normalized = normalize_type(param.as_type)
    expected = expected_normalized if expected_normalized is not None else "variant"
    if (
        not param.by_val
        and not param.is_array
        and len(slot) == 1
        and slot[0].kind is TokenKind.IDENTIFIER
        and expected != "variant"
    ):
        actual = normalize_type(env.get(slot[0].raw_text.lower()))
        if (
            actual
            and actual != "variant"
            and actual != expected
            and is_known_scalar_type(actual)
            and is_known_scalar_type(expected)
        ):
            push(
                "byRefArgumentTypeMismatch",
                f"'{slot[0].raw_text}' is declared As {_capitalized(actual)}, but parameter "
                f"'{param.name}' of Event '{event_name}' is ByRef As {param.as_type}. This is a VBE "
                "compile error: ByRef argument type mismatch.",
                span,
            )
        return
    value_range = _WHOLE_RANGES.get(expected) if param.by_val else None
    if value_range is None or len(slot) > 2:
        return
    if (
        len(slot) == 1
        and slot[0].kind is TokenKind.STRING_LITERAL
        and is_invalid_numeric_string(string_literal_value(slot[0].raw_text))
    ):
        push(
            "argumentTypeMismatch",
            f"Argument '{param.name}' of Event '{event_name}' expects {param.as_type}, but got "
            f"{slot[0].raw_text}, which is no number. This will raise Run-time error '13': Type "
            "mismatch.",
            span,
        )
        return
    negative = len(slot) == 2 and slot[0].raw_text == "-"
    literal = slot[1 if negative else 0] if len(slot) > (1 if negative else 0) else None
    raw = (
        parse_vba_integer_literal(literal.raw_text)
        if literal is not None
        and literal.kind is TokenKind.INTEGER_LITERAL
        and len(slot) == (2 if negative else 1)
        else None
    )
    value = None if raw is None else (-raw if negative else raw)
    if value is not None and (value < value_range[0] or value > value_range[1]):
        push(
            "argumentTypeMismatch",
            f"Argument '{param.name}' of Event '{event_name}' expects {param.as_type}, but got "
            f"{js_number_to_string(value)}, which does not fit. This will raise Run-time error '6': "
            "Overflow.",
            span,
        )


def _capitalized(type_name: str) -> str:
    return type_name[:1].upper() + type_name[1:]


# -- checkDeclarePtrSafeForWin64 -------------------------------------------


def _conditional_value_truthy(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    return isinstance(value, str) and len(value) > 0


def check_declare_ptr_safe_for_win64(
    source: str,
    mod: ModuleNode,
    conditional_compilation: ConditionalCompilationEnvironment | None,
    host: str | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """In 64-bit Office a Declare needs PtrSafe, "The code in this project must be
    updated for use on 64-bit systems". Win64 is read with the default compiler
    constants merged in, as the branch activity reads it, so an active Declare is
    checked unless the caller says Win64 is off (issue #215, measured in 64-bit
    Excel and Word). VB6 has no PtrSafe and is never 64-bit."""
    if host is not None and host.lower() == "vb6":
        return
    constants = compiler_constants_with_defaults(conditional_compilation)
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
    if document_type in ("workbook", "worksheet", "chart", "document"):
        return document_type
    if document_type == "userform":
        return "UserForm"
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
    """A Sub whose name matches an Office event handler that this module's document
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
            f"{module_description} is not where that event is wired. "
            "It will behave like an ordinary procedure here.",
            declared_name_span(source, member.span, member.name),
        )
