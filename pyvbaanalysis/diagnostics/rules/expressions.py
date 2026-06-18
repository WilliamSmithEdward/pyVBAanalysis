"""Rule family: expression-syntax rules (self-contained slice).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/expressions.ts: unbalanced
parentheses. The other rules in the family (invalid-expression-syntax,
division-by-zero, and the parenthesized/parenless call-shape rules) need call
extraction, type inference, and the member-completion context, so they land in M8.
"""

from __future__ import annotations

from ...lexer.token_kinds import TokenKind
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import Span
from ..context import PushFn


def check_unbalanced_parens(source: str, push: PushFn) -> None:
    """Every parenthesis must be matched within its logical statement (a `(` left
    open at a statement boundary, or a stray `)`, is a VBE Syntax error)."""
    toks = tokenize_cached(source)
    depth = 0
    open_offsets: list[int] = []
    flagged = False

    def flush() -> None:
        nonlocal depth, flagged
        if not flagged and depth > 0:
            off = open_offsets[0]
            push("unbalancedParens", "Unbalanced parentheses: a ')' is missing.", Span(off, off + 1))
        depth = 0
        open_offsets.clear()
        flagged = False

    for tok in toks:
        if tok.kind is TokenKind.NEWLINE:
            flush()
            continue
        if tok.kind is TokenKind.COLON and depth == 0:
            flush()
            continue
        if tok.kind is not TokenKind.PUNCTUATION:
            continue
        if tok.raw_text == "(":
            depth += 1
            open_offsets.append(tok.start)
        elif tok.raw_text == ")":
            if depth == 0:
                if not flagged:
                    push(
                        "unbalancedParens",
                        "Unbalanced parentheses: an unexpected ')' was found.",
                        Span(tok.start, tok.end),
                    )
                    flagged = True
            else:
                depth -= 1
                open_offsets.pop()
    flush()
