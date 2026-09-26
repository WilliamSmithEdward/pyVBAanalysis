"""Collection-accessor chain resolution (XLIDE v2.5.10 parity).

Every indexed collection accessor resolves to its element type regardless of
member kind; the explicit element accessors (Item/_Default/Add) are not
re-indexed; a mixed-element collection resolves through its union surface; and
empty parentheses are a call, not collection indexing.
"""

from __future__ import annotations

from typing import Any

import pytest

from pyvbaanalysis import analyze_project
from pyvbaanalysis.completion import member_access
from pyvbaanalysis.completion.member_access import (
    MemberCompletionContext,
    resolve_receiver_type_at,
)
from pyvbaanalysis.lexer.token_kinds import TokenKind
from pyvbaanalysis.lexer.tokenize import tokenize, tokenize_cached
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind


def _receiver_type(body: str) -> str | None:
    """Receiver type at the trailing dot of the last line of ``body``."""
    source = f"Sub S()\n    Dim ws As Worksheet\n    {body}\nEnd Sub\n"
    offset = source.rindex(".") + 1
    return resolve_receiver_type_at(source, offset, MemberCompletionContext())


def test_method_kind_collection_accessor_resolves_element() -> None:
    # ChartObjects is a method-kind accessor returning the ChartObjects
    # collection; calling it with an index resolves to the element.
    assert _receiver_type("ws.ChartObjects(1).") == "Excel.ChartObject"


def test_uncalled_collection_member_keeps_collection_type() -> None:
    assert _receiver_type("ws.ChartObjects.") == "Excel.ChartObjects"


def test_item_is_not_reindexed() -> None:
    # SparklineGroups.Item(1) already returns the element (SparklineGroup);
    # it must not over-resolve one more level into Sparkline.
    assert (
        _receiver_type("ws.Range(\"A1\").SparklineGroups.Item(1).")
        == "Excel.SparklineGroup"
    )


def test_single_typed_collection_index_resolves_element() -> None:
    # Worksheet.Range("A1") keeps its concrete Range return type.
    assert _receiver_type("ws.Range(\"A1\").") == "Excel.Range"


# -- one analysis pass resolving every dot of a chain ------------------------


def _pass_context(source: str) -> MemberCompletionContext:
    """A context like the diagnostics pass builds: the AST and the shared stream."""
    return MemberCompletionContext(
        parsed_module=parse_module(source),
        source_tokens=[t for t in tokenize_cached(source) if t.kind is not TokenKind.COMMENT],
    )


@pytest.mark.parametrize(
    "chain",
    [
        'ws.Range("A1").Offset(1, 0).Resize(2).Font.Bold',
        'Application.Workbooks(1).Worksheets("a").Cells(1, 1).Value',
        'ws.Zzq.Range("A1").Value',
        '(ws).Range("A1").Value',
        "ws.ChartObjects(1).Chart.ChartArea.Format.Fill",
        'ws.Range("A1").SparklineGroups.Item(1).Axes',
    ],
)
def test_receivers_continued_from_the_previous_dot_match_resolving_each_afresh(chain: str) -> None:
    # One context resolves the dots in order, so each dot continues from the one
    # before it; a new context per dot resolves each chain from its root.
    source = (
        f"Sub S()\n    Dim ws As Worksheet\n    x = {chain}\n"
        '    With ws\n        y = .Range("A1").Font.Bold\n    End With\nEnd Sub\n'
    )
    dots = [t.end for t in tokenize(source) if t.raw_text == "."]
    shared = _pass_context(source)
    continued = [resolve_receiver_type_at(source, end, shared) for end in dots]
    afresh = [resolve_receiver_type_at(source, end, _pass_context(source)) for end in dots]
    assert continued == afresh
    assert "Excel.Range" in continued


def test_a_long_chain_resolves_each_member_a_bounded_number_of_times(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    # Every dot of a chain resolved its chain from the root: 400 members cost
    # 80,600 member lookups in one analysis, and 3,000 took 16 seconds.
    members = 400
    calls = 0
    original = member_access._resolve_any_member_return_type

    def counting(*args: Any, **kwargs: Any) -> Any:
        nonlocal calls
        calls += 1
        return original(*args, **kwargs)

    monkeypatch.setattr(member_access, "_resolve_any_member_return_type", counting)
    source = "Sub S()\n    Dim x As Variant\n    x = Application" + ".Application" * members + ".Name\nEnd Sub\n"
    analyze_project([ModuleInput("Module1", ModuleSymbolKind.STANDARD, source)], host="excel")
    assert calls < 4 * members
