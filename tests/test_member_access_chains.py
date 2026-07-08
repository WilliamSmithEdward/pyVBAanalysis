"""Collection-accessor chain resolution (XLIDE v2.5.10 parity).

Every indexed collection accessor resolves to its element type regardless of
member kind; the explicit element accessors (Item/_Default/Add) are not
re-indexed; a mixed-element collection resolves through its union surface; and
empty parentheses are a call, not collection indexing.
"""

from __future__ import annotations

from pyvbaanalysis.completion.member_access import (
    MemberCompletionContext,
    resolve_receiver_type_at,
)


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
