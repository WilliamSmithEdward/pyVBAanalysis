"""Rule: conditional-compilation directive forms the VBE refuses (XLIDE issue #130).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/directiveForms.ts. Measured
in Excel 16.0 (build 20326, 2026-09-26):

- duplicate-const-directive: `#Const FEATURE = 1` twice in one module ->
  "Duplicate definition".
- directive-trailing-statement: code after a colon on a directive line,
  `#If VBA7 Then: Debug.Print 1` -> "An # ElseIf, # Else, or # EndIf must be
  preceded by an # If clause" (the colon ends the directive, and what follows is no
  longer part of it).
"""

from __future__ import annotations

import re
from bisect import bisect_left

from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import (
    ConditionalDirectiveKind,
    ConditionalDirectiveNode,
    EnumNode,
    ModuleNode,
    ProcedureNode,
    Span,
    TypeNode,
    iter_body_nodes,
)
from ..context import PushFn

_LINE_TERMINATOR_RE = re.compile(r"[\r\n]")


def check_directive_forms(source: str, mod: ModuleNode, push: PushFn) -> None:
    directives: list[ConditionalDirectiveNode] = []
    for member in mod.members:
        if isinstance(member, ConditionalDirectiveNode):
            directives.append(member)
        elif isinstance(member, ProcedureNode):
            # Upstream recurses into every nested block body; iter_body_nodes visits
            # the same nodes in the same order on an explicit stack.
            for node in iter_body_nodes(member.body):
                if isinstance(node, ConditionalDirectiveNode):
                    directives.append(node)
        elif isinstance(member, (EnumNode, TypeNode)):
            if member.directives is not None:
                directives.extend(member.directives)
    defined: dict[str, ConditionalDirectiveNode] = {}
    for directive in directives:
        if directive.directive_kind is ConditionalDirectiveKind.CONST and directive.name:
            lower = directive.name.lower()
            earlier = defined.get(lower)
            if earlier is not None:
                push(
                    "duplicateConstDirective",
                    f"'#Const {directive.name}' is already defined in this module. This is a VBE "
                    "compile error: Duplicate definition.",
                    directive.name_span if directive.name_span is not None else directive.span,
                )
            else:
                defined[lower] = directive
    # A directive line that a colon continues: the tokens after the colon on
    # the directive's own physical line.
    tokens = tokenize_cached(source)
    for directive in directives:
        line_end = _line_end_after(source, directive.span.end)
        colon = -1
        # Upstream walks the module's tokens from the first one, skipping those that
        # start before the directive ends; the tokens are in source order, so a
        # binary search finds the same first token without the walk.
        for i in range(bisect_left(tokens, directive.span.end, key=_token_start), len(tokens)):
            tok = tokens[i]
            if tok.start >= line_end:
                break
            if tok.kind is TokenKind.COLON:
                colon = i
                break
        if colon < 0:
            continue
        following = tokens[colon + 1] if colon + 1 < len(tokens) else None
        if (
            following is not None
            and following.kind is not TokenKind.NEWLINE
            and following.kind is not TokenKind.COMMENT
            and following.start < line_end
        ):
            push(
                "directiveTrailingStatement",
                "A compiler directive takes the whole line: nothing may follow the ':' after it. "
                'This is a VBE compile error (the VBE reports "An # ElseIf, # Else, or # EndIf '
                'must be preceded by an # If clause").',
                Span(following.start, line_end),
            )


def _line_end_after(source: str, start: int) -> int:
    """Offset of the first line terminator at or after `start`."""
    found = _LINE_TERMINATOR_RE.search(source, start)
    return found.start() if found is not None else len(source)


def _token_start(tok: VbaToken) -> int:
    return tok.start
