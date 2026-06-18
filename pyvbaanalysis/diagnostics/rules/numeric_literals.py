"""Rule family: type-suffixed numeric literal overflow.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/numericLiterals.ts. A `%`
(Integer) type suffix on a decimal literal whose magnitude exceeds the Integer
range is a VBE compile-time Syntax error, intrinsic to the token. Only the `%`
suffix is checked - the `&` Long suffix is ambiguous with string concatenation, so
its overflow is deferred (a conservative false negative per no-false-positive).
"""

from __future__ import annotations

import re

from ...conditional import ConditionalActivityTracker
from ...lexer.token_kinds import TokenKind
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import Span
from ...types.type_names import numeric_literal_bounds
from ..context import PushFn

# A pure-decimal integer literal with the `%` (Integer) type suffix.
_INTEGER_SUFFIXED_LITERAL = re.compile(r"^(\d+)%$")


def check_suffixed_literal_overflow(
    source: str, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    bounds = numeric_literal_bounds("integer")
    if bounds is None:
        return
    for tok in tokenize_cached(source):
        if tok.kind is not TokenKind.INTEGER_LITERAL:
            continue
        match = _INTEGER_SUFFIXED_LITERAL.match(tok.raw_text)
        if match is None:
            continue
        # The digit run is unsigned, so only the upper bound is reachable.
        if int(match.group(1)) <= bounds.max:
            continue
        span = Span(tok.start, tok.end)
        if activity is not None and activity.is_inactive(span):
            continue
        push(
            "suffixedLiteralOverflow",
            f"The literal '{tok.raw_text}' is outside the {bounds.label} range "
            f"{bounds.min} to {bounds.max} of its '%' type suffix. VBE rejects this "
            "at compile time as a Syntax error.",
            span,
        )
