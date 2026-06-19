"""Rule family: assignment-statement rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/assignments.ts. Only the
Mid-statement literal-target rule is ported in M7; the type-coupled assignment
rules (assignment type mismatch, Set object types) land in M8.
"""

from __future__ import annotations

import re

from ...conditional import ConditionalActivityTracker
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...symbols.symbol_model import ModuleSymbols
from ..context import PushFn
from ..walker import (
    active_module_members,
    for_each_statement,
    statement_tokens_after_leading_label,
    token_name,
)

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
