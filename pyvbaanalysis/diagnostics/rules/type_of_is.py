"""Rule family: TypeOf ... Is expression rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/typeOfIs.ts. The
TypeOf-missing-operand syntax check is host-free and ported. is-operator-non-object
is implemented but deferred: its oracle cases use `Debug.Print n Is Nothing`, and
the Python lexer classifies the intrinsic-object keyword `Debug` as a keyword (not
an identifier), so `Debug.<member>` does not parse as an expression and the
expression-AST walk never reaches the `Is` node. That is a lexer/parser parity
gap (agent.md Risk 4) to fix at the foundation level; typeof-is-always-false
additionally needs the host object-compatibility tables (M9).
"""

from __future__ import annotations

from ...conditional import ConditionalActivityTracker
from ...lexer.token_kinds import TokenKind
from ...lexer.tokenize import tokenize
from ...parser.nodes import Span
from ..context import PushFn


def check_typeof_missing_operand(
    source: str, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """`TypeOf` requires an object expression before `Is`; `TypeOf Is Y` is a syntax error."""
    toks = [
        tok
        for tok in tokenize(source)
        if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE
    ]
    for i in range(len(toks) - 1):
        if (toks[i].canonical_text or toks[i].raw_text).lower() != "typeof":
            continue
        if (toks[i + 1].canonical_text or toks[i + 1].raw_text).lower() != "is":
            continue
        span = Span(toks[i].start, toks[i + 1].end)
        if activity is not None and activity.is_inactive(span):
            continue
        push("typeofMissingOperand", "'TypeOf' requires an object expression before 'Is'.", span)
