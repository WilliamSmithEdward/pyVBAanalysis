"""heldObjects.ts parity: what object a local holds at a statement, and the
classes of a Collection local's items (XLIDE issues #246, #356, #447). Cases from
upstream's heldObjectSnapshots test."""

from __future__ import annotations

from pyvbaanalysis.conditional import ConditionalCompilationEnvironment, create_conditional_activity_tracker
from pyvbaanalysis.diagnostics.held_objects import HELD_VALUE, HeldObjects, held_objects_at
from pyvbaanalysis.parser.nodes import BodyNode, ProcedureNode, iter_body_nodes
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.symbols.build_module_symbols import BuildModuleSymbolsOptions, build_module_symbols
from pyvbaanalysis.symbols.symbol_model import ModuleSymbolKind


def _rows(body: str, vba7: bool | None = None) -> list[tuple[str, HeldObjects]]:
    source = "Sub P()\n" + body + "\nEnd Sub"
    mod = parse_module(source)
    symbols = build_module_symbols(
        "M", ModuleSymbolKind.STANDARD, source, BuildModuleSymbolsOptions(parsed_module=mod)
    )
    proc = next(member for member in mod.members if isinstance(member, ProcedureNode))
    activity = (
        None
        if vba7 is None
        else create_conditional_activity_tracker(
            mod, ConditionalCompilationEnvironment(compiler_constants={"VBA7": vba7})
        )
    )
    at = held_objects_at(source, proc, symbols, activity)
    nodes: list[BodyNode] = list(iter_body_nodes(proc.body))
    return [(source[node.span.start : node.span.end].strip(), at(node)) for node in nodes]


def _find(rows: list[tuple[str, HeldObjects]], text: str) -> HeldObjects:
    return next(value for row_text, value in rows if row_text == text)


def test_an_unchanged_snapshot_is_shared() -> None:
    body = "\n".join(f"Dim o{i} As New C" for i in range(100)) + "\n" + "x = 1\n" * 200
    reads = [value for text, value in _rows(body) if text == "x = 1"]
    assert len(reads) == 200
    assert len(reads[0].classes) == 100
    assert reads[199] is reads[0]


def test_collection_items_across_object_and_literal_additions() -> None:
    rows = _rows(
        "Dim c As New Collection\nc.Add New A\nx = c.Count\nc.Add 1\nx = c.Count\n"
        "c.Add New B, , Before:=1\nx = c.Count"
    )
    reads = [value.items.get("c") for text, value in rows if text == "x = c.Count"]
    assert reads == [["A"], ["A", HELD_VALUE], ["B", "A", HELD_VALUE]]


def test_branch_entry_items_and_earlier_snapshots_after_restoring() -> None:
    rows = _rows(
        "Dim c As New Collection\nc.Add New A\nx = 1\nIf flag Then\nc.Add New B\nx = 2\nElse\nx = 3\nEnd If\nx = 4"
    )
    assert _find(rows, "x = 1").items.get("c") == ["A"]
    assert _find(rows, "x = 2").items.get("c") == ["A", "B"]
    assert _find(rows, "x = 3").items.get("c") == ["A"]
    # A block that names c conservatively ends its item facts on exit.
    assert _find(rows, "x = 4").items.get("c") is None
    block = next(value for text, value in rows if text.startswith("If flag"))
    assert block.items.get("c") == ["A"]


def test_aliasing_a_collection_or_calling_a_changing_member() -> None:
    rows = _rows("Dim c As New Collection\nDim d As Object\nc.Add New A\nx = 1\nSet d = c\nx = 2\nc.Remove 1\nx = 3")
    assert _find(rows, "x = 1").items.get("c") == ["A"]
    assert _find(rows, "x = 2").items.get("c") is None
    assert _find(rows, "x = 2").classes.get("d") == "Collection"
    assert _find(rows, "x = 3").items.get("c") is None


def test_activity_environments_stay_independent() -> None:
    body = "Dim o As New A\n#If VBA7 Then\nSet o = New B\n#End If\nx = 1"
    assert _find(_rows(body, False), "x = 1").classes.get("o") == "A"
    assert _find(_rows(body, True), "x = 1").classes.get("o") == "B"
    assert _find(_rows(body, False), "x = 1").classes.get("o") == "A"


def test_a_changing_member_ends_items_and_a_new_collection_replaces_them() -> None:
    rows = _rows("Dim c As New Collection\nc.Add New A\nx = 1\nc.Remove 1\nx = 2\nSet c = New Collection\nx = 3")
    assert _find(rows, "x = 1").items.get("c") == ["A"]
    assert _find(rows, "x = 2").items.get("c") is None
    assert _find(rows, "x = 2").classes.get("c") == "Collection"
    assert _find(rows, "x = 3").items.get("c") == []


def test_host_objects_and_prog_ids() -> None:
    rows = _rows(
        "Dim a As Object, d As Object\nSet a = Application\nSet d = CreateObject(\"Scripting.Dictionary\")\nx = 1"
    )
    held = _find(rows, "x = 1")
    assert held.classes.get("a") == "Application"
    assert held.classes.get("d") == "Scripting.Dictionary"
