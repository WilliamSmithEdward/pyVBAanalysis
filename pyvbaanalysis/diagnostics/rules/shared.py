"""Helpers shared by more than one diagnostics rule family.

Ported incrementally from xlide_vscode/src/analyzer/diagnostics/rules/shared.ts as
families need them (the full file also has host/type-inference-coupled helpers -
the exhaustive member-surface resolver, read-reference scanner - that land with
their consumer rules in M8/M9).
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

from ...conditional import collect_conditional_directives
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import ConditionalDirectiveKind, ConditionalDirectiveNode, ModuleNode, Span
from ..context import statement_tokens
from ..walker import (
    absolute_span,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


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
