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


# -- With scope ---------------------------------------------------------------

# With blocks nested, on one line, behind line numbers and labels, continued over
# two lines, a leading dot outside any With, and a second procedure.
WITH_MODULE = """Sub A()
    Dim ws As Worksheet
    With ws
        .Range("A1").Value = 1
        With .Range("B1")
            .Font.Bold = True
            .Offset(1, 0).Value = 2
        End With
        .Cells(1, 1).Value = 3
    End With
    With ws: .Range("C1").Value = 4: End With
10  With ws.Range("D1")
20      .Value = 5
L1: End With
    .Range("A1").Value = 6
End Sub
Sub B()
    Dim ws As Worksheet
    With ws _
        .Range("E1")
        .Font.Italic = True
    End With
End Sub
"""


def test_with_receivers_match_with_the_scan_cache_without_it_and_without_the_stream() -> None:
    source = WITH_MODULE
    dots = [t.end for t in tokenize(source) if t.raw_text == "."]
    cached = _pass_context(source)
    cached.with_scan_cache = {}
    with_cache = [resolve_receiver_type_at(source, end, cached) for end in dots]
    without_cache = [resolve_receiver_type_at(source, end, _pass_context(source)) for end in dots]
    relexed = [resolve_receiver_type_at(source, end, MemberCompletionContext()) for end in dots]
    assert with_cache == without_cache == relexed
    # `.Font` inside `With .Range("B1")`, and `.Value` inside the numbered block.
    assert "Excel.Range" in with_cache
    assert set(cached.with_scan_cache) == {source.index("Sub A"), source.index("Sub B")}


def test_a_with_block_is_scanned_once_per_procedure_per_pass(monkeypatch: pytest.MonkeyPatch) -> None:
    # The With stack at a leading dot was found by re-reading the procedure from
    # its first line, per dot: a 500-line With block took 18 seconds.
    scans = 0
    original = member_access._with_scan_tokens

    def counting(*args: Any, **kwargs: Any) -> Any:
        nonlocal scans
        scans += 1
        return original(*args, **kwargs)

    monkeypatch.setattr(member_access, "_with_scan_tokens", counting)
    lines = "".join(f"        .Cells({r}, 1).Value = {r}\n" for r in range(1, 301))
    source = "Sub Fill()\n    With Worksheets(1)\n" + lines + "    End With\nEnd Sub\n"
    analyze_project([ModuleInput("Module1", ModuleSymbolKind.STANDARD, source)], host="excel")
    assert scans == 1


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
