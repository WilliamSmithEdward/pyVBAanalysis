"""Rule family: characters and lines the VBE refuses outright (XLIDE issues #132 and #133).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/strayTokens.ts. Measured in
Excel 16.0 (build 20326, 2026-09-26):

- stray-character: `n = 1;` (a semicolon outside a Print or Write list), a
  backtick, braces, a lone `@`, `~`, `|` -> "Syntax error". A non-breaking space
  (U+00A0, what a web page or Word document pastes) is not whitespace to the VBE:
  between tokens it is a Syntax error and at the start of a line it becomes part of
  the name ("Variable not defined"). The lexer gives all of these the `unknown`
  kind, and `;` the punctuation kind.
- line-too-long: a physical line of 1024 characters is refused; 1023 compiles.
"""

from __future__ import annotations

import re

from ...conditional import ConditionalActivityTracker
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import Span
from ..context import PushFn
from ..walker import token_text

_MAX_LINE_LENGTH = 1023

# Statements whose lists take `;` and `,` as output separators.
_PRINT_LIKE: frozenset[str] = frozenset({"print", "write", "debug"})

# The type-declaration characters a name may carry glued to it (upstream's
# /^[$%&!#@]$/, which matches exactly one of them).
_TYPE_DECLARATION_CHARACTERS: frozenset[str] = frozenset({"$", "%", "&", "!", "#", "@"})

_NAME_KINDS = (TokenKind.IDENTIFIER, TokenKind.KEYWORD, TokenKind.BRACKETED_IDENTIFIER)

# /^\d+$/: JavaScript's `\d` is ASCII only.
_DIGITS_RE = re.compile(r"[0-9]+")

# U+00A0, spelled as a code point so the source stays plain ASCII.
_NO_BREAK_SPACE = chr(0xA0)


def check_stray_characters(
    source: str, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    tokens = tokenize_cached(source)
    statement_head: VbaToken | None = None
    after_hash = False
    previous: VbaToken | None = None
    for tok in tokens:
        if tok.kind is TokenKind.NEWLINE or tok.kind is TokenKind.COLON:
            statement_head = None
            after_hash = False
            previous = None
            continue
        if tok.kind is TokenKind.COMMENT:
            continue
        prior = previous
        previous = tok
        if statement_head is None:
            # A line number is not the statement: `10 Debug.Print "a"; "b"` (XLIDE
            # issue #143). The next token is the head.
            if (
                tok.kind is TokenKind.INTEGER_LITERAL
                and prior is None
                and _DIGITS_RE.fullmatch(tok.raw_text) is not None
            ):
                previous = None
                continue
            statement_head = tok
            # `#Const`, `#If` lines are directives; `Print #1, x` names a file.
            after_hash = tok.kind is TokenKind.DIRECTIVE
        elif tok.kind is TokenKind.KEYWORD and token_text(tok) in ("then", "else"):
            # A one-line If runs a statement after Then and another after Else:
            # `If x Then Debug.Print a; b` is a Print list (XLIDE issue #143).
            statement_head = None
            continue
        span = Span(tok.start, tok.end)
        if tok.kind is TokenKind.UNKNOWN:
            # A run of such characters (four NBSPs as an indent) is one finding.
            if (activity is not None and activity.is_inactive(span)) or (
                prior is not None and prior.kind is TokenKind.UNKNOWN and prior.end == tok.start
            ):
                continue
            # `1.5%`: the Integer suffix glued to a fractional literal is the
            # literal rule's finding, not a stray character. A type-declaration
            # character glued to a name (`Left$(`, `n%`, `x@`) is the name's suffix.
            if (
                tok.raw_text == "%"
                and prior is not None
                and prior.kind is TokenKind.FLOAT_LITERAL
                and prior.end == tok.start
            ):
                continue
            if (
                tok.raw_text in _TYPE_DECLARATION_CHARACTERS
                and prior is not None
                and prior.end == tok.start
                and prior.kind in _NAME_KINDS
            ):
                continue
            text = tok.raw_text
            if _is_no_break_space_run(text):
                push(
                    "strayCharacter",
                    "A non-breaking space (U+00A0) is not whitespace in VBA: the VBE reads it as "
                    "part of a name or refuses the line. Replace it with an ordinary space. This "
                    "is a VBE compile error.",
                    span,
                )
            else:
                push(
                    "strayCharacter",
                    f"'{text}' is not a character VBA uses here. This is a VBE compile error: "
                    "Syntax error.",
                    span,
                )
            continue
        if (
            tok.kind is TokenKind.PUNCTUATION
            and tok.raw_text == ";"
            and not after_hash
            and not (activity is not None and activity.is_inactive(span))
        ):
            head = token_text(statement_head) if statement_head is not None else ""
            if head not in _PRINT_LIKE:
                push(
                    "strayCharacter",
                    "A ';' ends no VBA statement: it separates items only in a Print or Write "
                    "list. Remove it. This is a VBE compile error: Syntax error.",
                    span,
                )
    line_start = 0
    line_index = 0
    while line_start <= len(source):
        line_end = source.find("\n", line_start)
        if line_end < 0:
            line_end = len(source)
        visible_end = line_end
        if visible_end > line_start and source[visible_end - 1] == "\r":
            visible_end -= 1
        length = visible_end - line_start
        if length > _MAX_LINE_LENGTH:
            span = Span(line_start + _MAX_LINE_LENGTH, visible_end)
            if not (activity is not None and activity.is_inactive(span)):
                push(
                    "lineTooLong",
                    f"Line {line_index + 1} is {length} characters long; the VBE accepts "
                    f"{_MAX_LINE_LENGTH}. Break it with a line continuation. This is a VBE "
                    "compile error.",
                    span,
                )
        if line_end >= len(source):
            break
        line_start = line_end + 1
        line_index += 1


def _is_no_break_space_run(text: str) -> bool:
    """Upstream's /^[\\u00A0]+$/: one or more U+00A0 and nothing else."""
    return len(text) > 0 and text.count(_NO_BREAK_SPACE) == len(text)
