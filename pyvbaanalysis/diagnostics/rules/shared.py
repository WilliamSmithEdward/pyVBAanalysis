"""Helpers shared by more than one diagnostics rule family.

Ported incrementally from xlide_vscode/src/analyzer/diagnostics/rules/shared.ts as
families need them (the full file also has host/type-inference-coupled helpers -
the exhaustive member-surface resolver, read-reference scanner - that land with
their consumer rules in M8/M9).
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass

from ...call.call_context import bare_call_statement_target as call_statement_target
from ...completion import (
    MemberCompletionContext,
    MemberCompletionEntry,
    resolve_member_surface_at,
)
from ...conditional import ConditionalActivityTracker, collect_conditional_directives
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ConditionalDirectiveKind,
    ConditionalDirectiveNode,
    DoBlockNode,
    ModuleNode,
    Span,
    is_leaf_statement,
)
from ...symbols.symbol_model import VbaProjectClassMembers, qualified_procedure_key
from ..call_extraction import CallableTypeSignature
from ..context import statement_tokens
from ..walker import (
    absolute_span,
    block_footer_line_span,
    block_header_line_span,
    first_executable_token_index,
    is_inactive_node,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
    top_level_operator_index,
)


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


# -- exhaustive member-surface resolver ------------------------------------


@dataclass(frozen=True, slots=True)
class ExhaustiveMemberSurface:
    """An EXHAUSTIVE member surface: a complete member list that proves absence."""

    owner: str
    members: list[MemberCompletionEntry]


def resolve_exhaustive_member_surface(
    source: str, dot_end_offset: int, member_ctx: MemberCompletionContext
) -> ExhaustiveMemberSurface | None:
    """The member surface at a member-access dot, but ONLY when it is exhaustive.

    Ported from resolveExhaustiveMemberSurface (shared.ts). Returns None unless the
    resolved receiver type yields an exhaustive surface, the no-false-positive gate
    for member-not-found: a non-exhaustive host type, Object/Variant, or unresolved
    receiver produces no surface, so no member can be proven absent.
    """
    surface = resolve_member_surface_at(source, dot_end_offset, member_ctx)
    if surface is None or not surface.exhaustive:
        return None
    return ExhaustiveMemberSurface(owner=surface.owner, members=surface.members)


# -- name-token hits -------------------------------------------------------


@dataclass(frozen=True, slots=True)
class NameTokenHit:
    name: str
    span: Span
    bracketed: bool


def name_token_hit(base: Span, tok: VbaToken, name: str) -> NameTokenHit:
    return NameTokenHit(
        name=name, span=absolute_span(base, tok), bracketed=tok.kind is TokenKind.BRACKETED_IDENTIFIER
    )


def declaration_name_hit(source: str, span: Span, name: str) -> NameTokenHit | None:
    """The first token in the statement span whose name matches `name`."""
    lower = name.lower()
    for tok in statement_tokens(source, span):
        found = token_name(tok)
        if found is not None and found.lower() == lower:
            return name_token_hit(span, tok, found)
    return None


# -- module-level declaration-statement classifier -------------------------

# Visibility modifiers that may lead a module-level declaration in a procedure
# body (no Static - a Static local is legal inside a procedure).
_PROCEDURE_BODY_MODULE_DECLARATION_MODIFIERS: frozenset[str] = frozenset(
    {"public", "private", "friend", "global"}
)

DEFTYPE_KEYWORDS: frozenset[str] = frozenset(
    {
        "defbool", "defbyte", "defcur", "defdate", "defdbl", "defdec", "defint",
        "deflng", "deflnglng", "deflngptr", "defobj", "defsng", "defstr", "defvar",
    }
)


def leading_declaration_modifier_count(toks: Sequence[VbaToken]) -> int:
    i = 0
    while token_text(_at(toks, i)) in _PROCEDURE_BODY_MODULE_DECLARATION_MODIFIERS:
        i += 1
    return i


def module_declaration_statement_in_procedure(source: str, span: Span) -> tuple[str, Span] | None:
    """Classify a statement that is really a module-level declaration (Option,
    Attribute, Def*, visibility-led declaration, Type/Enum block) inside a body."""
    toks = statement_tokens_after_leading_label(source, span)
    first = toks[0] if toks else None
    head = token_text(first)
    if first is None:
        return None
    if head == "option":
        return ("Option statements", absolute_span(span, first))
    if head == "attribute":
        return ("Attribute statements", absolute_span(span, first))
    if head in DEFTYPE_KEYWORDS:
        label = (first.canonical_text if first.canonical_text is not None else first.raw_text) + " statements"
        return (label, absolute_span(span, first))
    modifier_count = leading_declaration_modifier_count(toks)
    declaration_head = token_text(_at(toks, modifier_count))
    if declaration_head == "type" or declaration_head == "enum":
        tok = _at(toks, modifier_count)
        label = "Type blocks" if declaration_head == "type" else "Enum blocks"
        return (label, absolute_span(span, tok) if tok is not None else absolute_span(span, first))
    if head in _PROCEDURE_BODY_MODULE_DECLARATION_MODIFIERS:
        label = (first.canonical_text if first.canonical_text is not None else first.raw_text) + " declarations"
        return (label, absolute_span(span, first))
    return None


# -- undeclared-reference scanner ------------------------------------------
#
# The conservative read-reference scanner the undeclared-variable rule rides
# over. The skip-set is exactly what prevents false positives: it removes every
# token position that is NOT a potential bare variable read (declaration spans,
# line labels, call targets and their argument identifiers, assignment LHS,
# qualified project member/callable qualifiers, New/Is/GoTo/RaiseEvent/AddressOf
# operands, and named-argument labels). Ported line-for-line from shared.ts.


@dataclass(frozen=True, slots=True)
class ValueReadReference:
    name: str
    span: Span


def for_each_undeclared_reference_span(
    source: str,
    body: Sequence[BodyNode],
    visit: Callable[[Span], None],
    activity: ConditionalActivityTracker | None = None,
) -> None:
    """Visit one span per executable statement / block-header / Do-footer line."""
    for node in body:
        if is_inactive_node(activity, node):
            continue
        if is_leaf_statement(node):
            visit(node.span)
            continue
        child = getattr(node, "body", None)
        if isinstance(child, list):
            visit(block_header_line_span(source, node.span))
            if isinstance(node, DoBlockNode):
                footer = block_footer_line_span(source, node.span)
                if footer.start > node.span.start:
                    visit(footer)
            for_each_undeclared_reference_span(source, child, visit, activity)


def value_read_references(
    source: str,
    span: Span,
    is_known_for_skip: Callable[[str], bool],
    module_signatures: Mapping[str, CallableTypeSignature],
    project_members: Sequence[VbaProjectClassMembers] | None,
) -> list[ValueReadReference]:
    """Bare identifier reads in a statement span, with declaration/call/etc. skipped."""
    toks = statement_tokens(source, span)
    out: list[ValueReadReference] = []
    skip = _undeclared_reference_skip_indexes(
        source, span, toks, is_known_for_skip, module_signatures, project_members
    )
    for i in range(len(toks)):
        if i in skip or not _is_potential_variable_reference_token(_at(toks, i)):
            continue
        if token_raw(_at(toks, i - 1)) == ".":
            continue
        name = token_name(toks[i])
        if not name:
            continue
        out.append(
            ValueReadReference(
                name=name, span=Span(span.start + toks[i].start, span.start + toks[i].end)
            )
        )
    return out


def token_raw(tok: VbaToken | None) -> str | None:
    return tok.raw_text if tok is not None else None


def _undeclared_reference_skip_indexes(
    source: str,
    span: Span,
    toks: Sequence[VbaToken],
    is_known: Callable[[str], bool],
    module_signatures: Mapping[str, CallableTypeSignature],
    project_members: Sequence[VbaProjectClassMembers] | None,
) -> set[int]:
    skip: set[int] = set()
    if len(toks) == 0:
        return skip
    if token_raw(_at(toks, 1)) == ":" or _is_line_label_only_statement(source, span, toks):
        skip.add(0)  # line label declaration
    first_executable = first_executable_token_index(toks)
    if module_declaration_statement_in_procedure(source, span):
        for i in range(first_executable, len(toks)):
            skip.add(i)
        return skip
    if token_text(_at(toks, first_executable)) == "implements":
        for i in range(first_executable + 1, len(toks)):
            skip.add(i)
        return skip

    call = call_statement_target(source, span)
    if call is not None:
        call_idx = next(
            (i for i, tok in enumerate(toks) if span.start + tok.start == call.span.start), -1
        )
        if call_idx >= 0:
            skip.add(call_idx)
            if not is_known(call.name):
                # Unknown call targets may be external procedures or unresolved call
                # errors; do not also guess about their argument identifiers.
                for i in range(call_idx + 1, len(toks)):
                    skip.add(i)

    assignment = _simple_assignment_lhs_identifier_index(toks)
    if assignment >= 0:
        skip.add(assignment)

    for i in range(len(toks)):
        word = token_text(toks[i])
        if _is_qualified_project_callable_qualifier(
            toks, i, module_signatures
        ) or _is_qualified_project_member_qualifier(toks, i, project_members):
            skip.add(i)
        if word == "new" and _is_potential_variable_reference_token(_at(toks, i + 1)):
            skip.add(i + 1)
        if (
            word == "is"
            and _has_earlier_type_of(toks, i)
            and _is_potential_variable_reference_token(_at(toks, i + 1))
        ):
            skip.add(i + 1)
        if _is_label_reference_keyword(word) and _is_potential_variable_reference_token(
            _at(toks, i + 1)
        ):
            skip.add(i + 1)
        if word == "raiseevent" and _is_potential_variable_reference_token(_at(toks, i + 1)):
            skip.add(i + 1)
        if word == "addressof" and _is_potential_variable_reference_token(_at(toks, i + 1)):
            skip.add(i + 1)
        if _is_named_argument_label(toks, i):
            skip.add(i)

    return skip


def _is_qualified_project_callable_qualifier(
    toks: Sequence[VbaToken],
    index: int,
    module_signatures: Mapping[str, CallableTypeSignature],
) -> bool:
    if not _is_potential_variable_reference_token(_at(toks, index)) or token_raw(
        _at(toks, index + 1)
    ) != ".":
        return False
    if not _is_potential_variable_reference_token(_at(toks, index + 2)):
        return False
    qualifier = token_name(toks[index])
    member = token_name(toks[index + 2])
    if not qualifier or not member:
        return False
    return qualified_procedure_key(qualifier, member) in module_signatures


def _is_qualified_project_member_qualifier(
    toks: Sequence[VbaToken],
    index: int,
    project_members: Sequence[VbaProjectClassMembers] | None,
) -> bool:
    if (
        not project_members
        or not _is_potential_variable_reference_token(_at(toks, index))
        or token_raw(_at(toks, index + 1)) != "."
    ):
        return False
    if not _is_potential_variable_reference_token(_at(toks, index + 2)):
        return False
    qualifier = token_name(toks[index])
    member = token_name(toks[index + 2])
    if not qualifier or not member:
        return False
    qualifier_lower = qualifier.lower()
    member_lower = member.lower()
    surface: VbaProjectClassMembers | None = None
    for candidate in project_members:
        if candidate.name.lower() != qualifier_lower:
            continue
        if surface is not None:
            return False
        surface = candidate
    if surface is None:
        return False
    if surface.kind == "standardModule":
        return True
    return any(candidate.name.lower() == member_lower for candidate in surface.members)


def _simple_assignment_lhs_identifier_index(toks: Sequence[VbaToken]) -> int:
    start = first_executable_token_index(toks)
    if token_text(_at(toks, start)) == "let" or token_text(_at(toks, start)) == "set":
        start += 1
    eq = top_level_operator_index(list(toks[start:]), "=")
    if eq != 1:
        return -1
    name_tok = _at(toks, start)
    return start if name_tok is not None and name_tok.kind is TokenKind.IDENTIFIER else -1


def _is_line_label_only_statement(source: str, span: Span, toks: Sequence[VbaToken]) -> bool:
    if len(toks) != 1 or not _is_potential_variable_reference_token(toks[0]):
        return False
    j = span.start + toks[0].end
    while j < len(source) and (source[j] == " " or source[j] == "\t"):
        j += 1
    return j < len(source) and source[j] == ":"


def _is_potential_variable_reference_token(tok: VbaToken | None) -> bool:
    return tok is not None and tok.kind in (
        TokenKind.IDENTIFIER,
        TokenKind.BRACKETED_IDENTIFIER,
    )


def _has_earlier_type_of(toks: Sequence[VbaToken], before: int) -> bool:
    for i in range(before):
        if token_text(toks[i]) == "typeof":
            return True
    return False


def _is_label_reference_keyword(word: str) -> bool:
    return word == "goto" or word == "gosub" or word == "resume"


def _is_named_argument_label(toks: Sequence[VbaToken], index: int) -> bool:
    if not _is_potential_variable_reference_token(_at(toks, index)):
        return False
    if token_raw(_at(toks, index + 1)) == ":=":
        return True
    return token_raw(_at(toks, index + 1)) == ":" and token_raw(_at(toks, index + 2)) == "="


# -- conditional-compilation branch-order scan -----------------------------


@dataclass(frozen=True, slots=True)
class ConditionalBranchOrderIssue:
    kind: str  # "elseifAfterElse" | "duplicateElse"
    directive: ConditionalDirectiveNode


@dataclass(frozen=True, slots=True)
class ConditionalBranchOrderScan:
    issues: list[ConditionalBranchOrderIssue]
    malformed_block_spans: list[Span]


@dataclass(slots=True)
class _ElseBranchFrame:
    seen_else: bool
    start: Span
    malformed: bool


def scan_conditional_compilation_branch_order(mod: ModuleNode) -> ConditionalBranchOrderScan:
    """Detect #ElseIf-after-#Else and duplicate #Else in conditional blocks, and the
    spans of malformed conditional blocks."""
    stack: list[_ElseBranchFrame] = []
    issues: list[ConditionalBranchOrderIssue] = []
    malformed_block_spans: list[Span] = []
    for occ in collect_conditional_directives(mod):
        directive = occ.directive
        kind = directive.directive_kind
        if kind is ConditionalDirectiveKind.IF:
            stack.append(_ElseBranchFrame(seen_else=False, start=directive.span, malformed=False))
        elif kind is ConditionalDirectiveKind.ELSE_IF:
            frame = stack[-1] if stack else None
            if frame is not None and frame.seen_else:
                frame.malformed = True
                issues.append(ConditionalBranchOrderIssue(kind="elseifAfterElse", directive=directive))
        elif kind is ConditionalDirectiveKind.ELSE:
            frame = stack[-1] if stack else None
            if frame is not None and frame.seen_else:
                frame.malformed = True
                issues.append(ConditionalBranchOrderIssue(kind="duplicateElse", directive=directive))
            if frame is not None:
                frame.seen_else = True
        elif kind is ConditionalDirectiveKind.END_IF:
            popped = stack.pop() if stack else None
            if popped is not None and popped.malformed:
                malformed_block_spans.append(Span(popped.start.start, directive.span.end))
        # Const, Unknown: no branch-order effect.
    return ConditionalBranchOrderScan(issues=issues, malformed_block_spans=malformed_block_spans)
