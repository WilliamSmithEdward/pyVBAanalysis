"""Rule: conditional-compilation directive forms the VBE refuses (XLIDE issue #130).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/directiveForms.ts. Measured
in Excel 16.0 (build 20326, 2026-09-26):

- duplicate-const-directive: `#Const FEATURE = 1` twice in one module ->
  "Duplicate definition".
- directive-trailing-statement: code after a colon on a directive line,
  `#If VBA7 Then: Debug.Print 1` -> "An # ElseIf, # Else, or # EndIf must be
  preceded by an # If clause" (the colon ends the directive, and what follows is no
  longer part of it).
- null-directive-condition: `#If Null Then`, `#If Null = 1 Then`, or `#If N` after
  `#Const N = Null` -> "Invalid use of Null" (issue #208).
- directive-type-mismatch: a condition that is a string VBA cannot read as
  True or False (`#If "abc"`, `#If ""`, `#If " True "`), or that hands such a
  string to an operator (`#If Not "abc"`, `#If "abc" = 1`) -> "Type mismatch",
  even in an arm the VBE skips (Excel 16.0, 2026-10-12); and a `#Const` value
  that does the same (`#Const A = "abc" + 1`), even unread.
"""

from __future__ import annotations

import re

from ...conditional.conditional_compilation import (
    ConditionalActivityTracker,
    ConditionalCompilationEnvironment,
    invalid_condition_directives,
)
from ...lexer.token_helpers import (
    colon_directive_starts,
    directive_colon,
    first_token_at_or_after,
    token_word,
)
from ...lexer.token_kinds import TokenKind
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


# What a directive line a colon follows fails to do, by its keyword (XLIDE issue #1323).
_COLON_DIRECTIVE_EFFECT: dict[str, str] = {
    "if": "opens no #If block, so the #Else, #ElseIf or #End If after it has no #If",
    "elseif": "starts no new arm: it belongs to the arm above it",
    "else": "starts no new arm: it belongs to the arm above it",
    "end": "closes no #If block",
    "endif": "closes no #If block",
    "const": "defines no constant",
}


def check_directive_forms(
    source: str,
    mod: ModuleNode,
    conditional_compilation: ConditionalCompilationEnvironment | None,
    push: PushFn,
    activity: ConditionalActivityTracker | None = None,
) -> None:
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
    # A directive line a statement colon follows is no directive: the VBE
    # reads the whole line as code, refused where it is compiled and skipped
    # with a skipped arm (XLIDE issue #1323). The parser reads the same lines.
    starts = colon_directive_starts(source)
    if len(starts) > 0:
        tokens = tokenize_cached(source)
        # Upstream's Set iterates in insertion order, which is source order.
        for start in sorted(starts):
            if activity is not None and activity.is_inactive(Span(start, start)):
                continue
            hash_index = first_token_at_or_after(tokens, start)
            colon = directive_colon(tokens, hash_index)
            if colon is None:
                continue
            keyword = token_word(tokens[hash_index + 1] if hash_index + 1 < len(tokens) else None)
            effect = _COLON_DIRECTIVE_EFFECT.get(keyword, _COLON_DIRECTIVE_EFFECT["if"])
            # What follows the colon, or the colon itself when nothing does.
            line_end = _line_end_at_or_after(source, colon.start)
            after_index = tokens.index(colon, hash_index) + 1
            after = tokens[after_index] if after_index < len(tokens) else None
            from_offset = (
                after.start
                if after is not None
                and after.kind is not TokenKind.NEWLINE
                and after.start < line_end
                else colon.start
            )
            push(
                "directiveTrailingStatement",
                f"A ':' after a compiler directive makes the whole line no directive, so it "
                f"{effect}. A directive takes its line alone: drop the ':', and start any "
                "comment with '. This is a VBE compile error.",
                Span(from_offset, colon.end if from_offset == colon.start else line_end),
            )
    for item in invalid_condition_directives(mod, conditional_compilation):
        directive = item.directive
        kind = directive.directive_kind.value
        if item.reason == "null":
            push(
                "nullDirectiveCondition",
                f"This #{kind} condition is Null, which is neither True nor False. This is a VBE "
                "compile error: Invalid use of Null.",
                directive.span,
            )
        elif item.mismatch is not None:
            what = (
                "#Const value"
                if directive.directive_kind is ConditionalDirectiveKind.CONST
                else f"#{kind} condition"
            )
            conversion = item.mismatch.conversion
            as_text = "True or False" if conversion == "boolean" else f"a {conversion}"
            push(
                "directiveTypeMismatch",
                f"This {what} needs the string {_vba_string_literal(item.mismatch.value)} as "
                f"{as_text}, and VBA cannot convert it. This is a VBE compile error: Type mismatch.",
                directive.span,
            )


def _vba_string_literal(value: str) -> str:
    """A string as VBA writes it, shortened past 40 characters."""
    shown = f"{value[:40]}..." if len(value) > 40 else value
    return '"' + shown.replace('"', '""') + '"'


def _line_end_at_or_after(source: str, start: int) -> int:
    """Offset of the first line terminator at or after `start`.

    Private copy of upstream's vbaSourceScan.ts lineEndAtOrAfter (outside the
    analyzer tree the port mirrors).
    """
    found = _LINE_TERMINATOR_RE.search(source, start)
    return found.start() if found is not None else len(source)
