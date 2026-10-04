"""Call-statement classification (callContext.ts parity)."""

from __future__ import annotations

from pyvbaanalysis.call.call_context import bare_call_statement_target
from pyvbaanalysis.parser.nodes import Span


def _target_name(source: str, statement: str) -> str | None:
    start = source.index(statement)
    target = bare_call_statement_target(source, Span(start, start + len(statement)))
    return target.name if target is not None else None


def test_name_and_colon_at_line_start_is_a_label() -> None:
    assert _target_name("Sub A()\n    L1:\nEnd Sub\n", "L1") is None


def test_name_and_colon_after_a_line_number_is_a_call() -> None:
    # XLIDE 2f49b93 (issue #230): in `10: L1:` the VBE reads L1 as a call.
    assert _target_name("Sub A()\n10: L1:\nEnd Sub\n", "L1") == "L1"
