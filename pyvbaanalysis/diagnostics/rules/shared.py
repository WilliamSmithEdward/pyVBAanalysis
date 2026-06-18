"""Shared helpers used across diagnostic rule families.

Ported incrementally from xlide_vscode/src/analyzer/diagnostics/rules/shared.ts as
families need them (the full file mixes host/type-inference-coupled helpers that
land with their consumer rules).
"""

from __future__ import annotations

from dataclasses import dataclass

from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import Span
from ..context import statement_tokens
from ..walker import absolute_span, token_name


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
