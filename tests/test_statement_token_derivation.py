"""The shared-stream statement-token derivation is byte-identical to a slice lex.

cached_statement_tokens derives each statement's span-relative token view from
the module's one shared tokenization (with a fallback to lexing the slice).
This sweep proves the derived view matches a fresh statement_tokens lex
token-for-token - kind, raw text, and rebased offsets - across every leaf
statement of every oracle-corpus module, so the fast path can never change
which tokens a rule sees.
"""

from __future__ import annotations

from pyvbaanalysis.evidence import load_oracle_cases
from pyvbaanalysis.lexer.token_helpers import cached_statement_tokens, statement_tokens
from pyvbaanalysis.parser.nodes import LeafStatementNode, ProcedureNode
from pyvbaanalysis.parser.parse_module import parse_module


def _leaf_spans(node: object, out: list[tuple[int, int]]) -> None:
    body = getattr(node, "body", None)
    if isinstance(body, list):
        for child in body:
            if isinstance(child, LeafStatementNode):
                out.append((child.span.start, child.span.end))
            _leaf_spans(child, out)


def test_derived_statement_tokens_match_slice_lex() -> None:
    checked = 0
    for case in load_oracle_cases():
        for module in case.modules:
            source = module.source
            spans: list[tuple[int, int]] = []
            for member in parse_module(source).members:
                if isinstance(member, ProcedureNode):
                    _leaf_spans(member, spans)
            for start, end in spans:
                derived = cached_statement_tokens(source, start, end)
                fresh = statement_tokens(source, start, end)
                assert [(t.kind, t.raw_text, t.start, t.end) for t in derived] == [
                    (t.kind, t.raw_text, t.start, t.end) for t in fresh
                ], f"{case.id}: span {start}:{end} diverged"
                checked += 1
    # The corpus currently yields ~515 leaf statements; the floor just proves
    # the sweep is not silently skipping everything if the walker breaks.
    assert checked > 400
