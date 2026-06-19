"""Procedure-local control-flow label collection.

Ported from the analysis-core subset of
xlide_vscode/src/analyzer/flow/procedureLabels.ts. Collects label declarations
and references (GoTo / GoSub / Resume / On Error GoTo / On n GoTo / On n GoSub)
within a single procedure, honoring conditional-compilation activity. The
completion/definition editor helpers in the TypeScript source are out of scope
and are not ported.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Sequence
from dataclasses import dataclass

from ..conditional import ConditionalActivityTracker
from ..lexer.token_helpers import (
    split_top_level_token_groups,
    statement_tokens,
    token_name,
    token_word,
    tokens_without_leading_line_number,
)
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import BodyNode, LeafStatementNode, ProcedureNode, Span, is_leaf_statement

_DECIMAL_LABEL_RE = re.compile(r"^\d+$")


@dataclass(frozen=True, slots=True)
class VbaProcedureLabel:
    """A label declaration or reference site within a procedure."""

    key: str
    text: str
    span: Span
    kind: str  # "name" | "line"


@dataclass(frozen=True, slots=True)
class VbaProcedureLabelReference:
    """A label reference, tagged with the statement form that targets it."""

    key: str
    text: str
    span: Span
    kind: str  # "name" | "line"
    # goto | gosub | resume | on-error-goto | on-goto | on-gosub
    statement_kind: str


def collect_procedure_labels(
    source: str, procedure: ProcedureNode, activity: ConditionalActivityTracker | None = None
) -> dict[str, VbaProcedureLabel]:
    """The first declaration of each distinct label key in the procedure."""
    labels: dict[str, VbaProcedureLabel] = {}
    for label in collect_procedure_label_declarations(source, procedure, activity):
        if label.key not in labels:
            labels[label.key] = label
    return labels


def collect_procedure_label_declarations(
    source: str, procedure: ProcedureNode, activity: ConditionalActivityTracker | None = None
) -> list[VbaProcedureLabel]:
    """Every active label declaration in the procedure, in source order."""
    labels: list[VbaProcedureLabel] = []

    def visit(stmt: LeafStatementNode) -> None:
        label = _statement_label_declaration(source, stmt.span)
        if label is not None:
            labels.append(label)

    _for_each_procedure_statement(procedure.body, visit, activity)
    return labels


def collect_procedure_label_references(
    source: str, procedure: ProcedureNode, activity: ConditionalActivityTracker | None = None
) -> list[VbaProcedureLabelReference]:
    """Every active label reference in the procedure, in source order."""
    refs: list[VbaProcedureLabelReference] = []

    def visit(stmt: LeafStatementNode) -> None:
        refs.extend(statement_label_references(source, stmt.span))

    _for_each_procedure_statement(procedure.body, visit, activity)
    return refs


def statement_label_references(source: str, span: Span) -> list[VbaProcedureLabelReference]:
    """Label references targeted by a single statement (GoTo/GoSub/Resume/On...)."""
    toks = tokens_without_leading_line_number(statement_tokens(source, span.start, span.end))
    if not toks:
        return []
    if token_word(toks[0]) == "on":
        return _on_statement_label_references(toks, span)
    refs: list[VbaProcedureLabelReference] = []
    for i, tok in enumerate(toks):
        word = token_word(tok)
        if word in ("goto", "gosub"):
            if word == "goto" and _is_on_error_goto_disable_at(toks, i):
                continue
            ref = _label_reference_after(toks, span, i + 1, word)
            if ref is not None:
                refs.append(ref)
            continue
        if word == "resume":
            next_tok = _at(toks, i + 1)
            if next_tok is None or token_word(next_tok) == "next":
                continue
            ref = _label_reference_after(toks, span, i + 1, "resume")
            if ref is not None:
                refs.append(ref)
    return refs


def _on_statement_label_references(
    toks: Sequence[VbaToken], span: Span
) -> list[VbaProcedureLabelReference]:
    if token_word(_at(toks, 1)) == "error":
        if token_word(_at(toks, 2)) == "resume" and token_word(_at(toks, 3)) == "next":
            return []
        if token_word(_at(toks, 2)) != "goto":
            return []
        if _at(toks, 3) is None or _on_error_goto_disable_target(toks, 3):
            return []
        ref = _label_reference_after(toks, span, 3, "on-error-goto")
        return [ref] if ref is not None else []
    flow_index = _on_flow_index(toks)
    if flow_index < 0:
        return []
    statement_kind = "on-gosub" if token_word(toks[flow_index]) == "gosub" else "on-goto"
    refs: list[VbaProcedureLabelReference] = []
    for group in split_top_level_token_groups(toks, flow_index + 1, ","):
        ref = _label_reference_group(group, span, statement_kind)
        if ref is not None:
            refs.append(ref)
    return refs


def _on_flow_index(toks: Sequence[VbaToken]) -> int:
    for i, tok in enumerate(toks):
        if i > 0 and token_word(tok) in ("goto", "gosub"):
            return i
    return -1


def _on_error_goto_disable_target(toks: Sequence[VbaToken], index: int) -> bool:
    target = _at(toks, index)
    if target is None:
        return False
    if target.kind is TokenKind.INTEGER_LITERAL and _normalized_decimal_label(target.raw_text) == "0":
        return True
    next_tok = _at(toks, index + 1)
    return (
        target.raw_text == "-"
        and next_tok is not None
        and next_tok.kind is TokenKind.INTEGER_LITERAL
        and _normalized_decimal_label(next_tok.raw_text) == "1"
    )


def _is_on_error_goto_disable_at(toks: Sequence[VbaToken], goto_index: int) -> bool:
    return (
        token_word(_at(toks, goto_index - 2)) == "on"
        and token_word(_at(toks, goto_index - 1)) == "error"
        and _on_error_goto_disable_target(toks, goto_index + 1)
    )


def _label_reference_after(
    toks: Sequence[VbaToken], base: Span, index: int, statement_kind: str
) -> VbaProcedureLabelReference | None:
    group = list(toks[index:])
    end = -1
    for j, tok in enumerate(group):
        if tok.raw_text == "," or token_word(tok) == "else":
            end = j
            break
    selected = group[:end] if end >= 0 else group
    return _label_reference_group(selected, base, statement_kind)


def _label_reference_group(
    group: Sequence[VbaToken], base: Span, statement_kind: str
) -> VbaProcedureLabelReference | None:
    content = [tok for tok in group if tok.kind is not TokenKind.COMMENT]
    if len(content) != 1:
        return None
    label = _label_from_token(content[0], base)
    if label is None:
        return None
    return VbaProcedureLabelReference(
        key=label.key,
        text=label.text,
        span=label.span,
        kind=label.kind,
        statement_kind=statement_kind,
    )


def _label_from_token(tok: VbaToken, base: Span) -> VbaProcedureLabel | None:
    name = token_name(tok)
    if name:
        return VbaProcedureLabel(
            key=f"name:{name.lower()}", text=name, span=_absolute_span(base, tok), kind="name"
        )
    if tok.kind is TokenKind.INTEGER_LITERAL:
        normalized = _normalized_decimal_label(tok.raw_text)
        if normalized is not None:
            return VbaProcedureLabel(
                key=f"line:{normalized}", text=tok.raw_text, span=_absolute_span(base, tok), kind="line"
            )
    return None


def _statement_label_declaration(source: str, span: Span) -> VbaProcedureLabel | None:
    toks = statement_tokens(source, span.start, span.end)
    first = _at(toks, 0)
    if first is None:
        return None
    label = _label_from_token(first, span)
    if label is None:
        return None
    if first.kind is TokenKind.INTEGER_LITERAL:
        # A leading decimal integer is a line-label declaration whether or not a
        # statement follows it on the same line.
        return label
    second = _at(toks, 1)
    if second is not None and second.raw_text == ":":
        return label
    if len(toks) == 1 and _has_source_colon_after_token(source, span, first):
        return label
    return None


def _normalized_decimal_label(raw: str) -> str | None:
    if _DECIMAL_LABEL_RE.match(raw) is None:
        return None
    stripped = raw.lstrip("0")
    return stripped if stripped else "0"


def _has_source_colon_after_token(source: str, span: Span, tok: VbaToken) -> bool:
    i = span.start + tok.end
    while i < len(source) and source[i] in (" ", "\t"):
        i += 1
    return i < len(source) and source[i] == ":"


def _for_each_procedure_statement(
    body: Sequence[BodyNode],
    visit: Callable[[LeafStatementNode], None],
    activity: ConditionalActivityTracker | None,
) -> None:
    for node in body:
        if activity is not None and activity.is_inactive(node.span):
            continue
        if is_leaf_statement(node):
            visit(node)
        else:
            child = getattr(node, "body", None)
            if isinstance(child, list):
                _for_each_procedure_statement(child, visit, activity)


def _absolute_span(base: Span, token: VbaToken) -> Span:
    return Span(base.start + token.start, base.start + token.end)


def _at(tokens: Sequence[VbaToken], i: int) -> VbaToken | None:
    return tokens[i] if 0 <= i < len(tokens) else None
