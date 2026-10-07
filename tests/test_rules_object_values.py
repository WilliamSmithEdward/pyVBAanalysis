"""Ported from xlide_vscode tests/objectReadIndexes.test.ts and
tests/objectReadTokenScans.test.ts (2f49b93): the objectDefaultValue rule's
per-statement token scans."""

from __future__ import annotations

import pytest

from pyvbaanalysis.completion import MemberCompletionContext
from pyvbaanalysis.diagnostics import walker
from pyvbaanalysis.diagnostics.rules.object_values import check_object_default_values
from pyvbaanalysis.host.host_registry import host_object_model_for_token
from pyvbaanalysis.parser.nodes import ProcedureNode, Span
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.symbols import BuildModuleSymbolsOptions, ModuleSymbolKind, build_module_symbols
from pyvbaanalysis.types.type_names import normalize_type

Hit = tuple[str, str, Span]


def _run(source: str, member_ctx: MemberCompletionContext) -> list[Hit]:
    mod = parse_module(source)
    symbols = build_module_symbols(
        "M", ModuleSymbolKind.STANDARD, source, BuildModuleSymbolsOptions(parsed_module=mod)
    )
    hits: list[Hit] = []

    def push(rule: str, message: str, span: Span, data: object = None) -> None:
        hits.append((rule, message, span))

    factory = check_object_default_values(source, symbols, member_ctx, push)
    for proc in mod.members:
        if not isinstance(proc, ProcedureNode):
            continue
        visitor = factory(proc)
        if visitor is not None:
            walker.for_each_statement_with_headers(source, proc.body, visitor)
    return hits


def _run_host(source: str, host: str = "excel") -> list[Hit]:
    return _run(source, MemberCompletionContext(model=host_object_model_for_token(host)))


def _run_body(body: str) -> tuple[str, list[Hit]]:
    source = "Sub P()\n" + body + "\nEnd Sub"
    return source, _run(source, MemberCompletionContext())


# --- objectReadIndexes.test.ts ---


def test_does_not_normalize_every_host_type_again_for_each_unrelated_statement(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    # Upstream spies on typeInference's normalizeType; here the rule module's
    # own binding is counted.
    calls = [0]
    real = normalize_type

    def counting(type_name: str | None) -> str | None:
        calls[0] += 1
        return real(type_name)

    monkeypatch.setattr("pyvbaanalysis.diagnostics.rules.object_values.normalize_type", counting)
    source = (
        "Sub P()\nDim n As Long\n"
        + "\n".join(f"Dim x{i} As Application" for i in range(100))
        + "\n"
        + "n = 1\n" * 1000
        + "End Sub"
    )
    assert _run_host(source) == []
    assert calls[0] < 500


def test_keeps_host_types_separate_across_procedures_that_shadow_module_variables() -> None:
    source = (
        "Dim x As Long\nSub A()\nDim x As Application\nSet x = Application\nDebug.Print X + 1\nEnd Sub\n"
        "Sub B()\nDim x As Long\nx = 2\nDebug.Print x + 1\nEnd Sub"
    )
    hits = _run_host(source)
    assert len(hits) == 1
    assert hits[0][0] == "assignmentTypeMismatch"
    assert source[hits[0][2].start : hits[0][2].end] == "X"


def test_retains_late_bound_collection_reads_and_word_document_reads_in_single_line_branches() -> (
    None
):
    collection = _run_host(
        "Sub P()\nDim x As Object\nSet x = New Collection\nDebug.Print X + 1\nEnd Sub"
    )
    assert any(hit[0] == "objectDefaultValue" and "'450'" in hit[1] for hit in collection)
    document = _run_host(
        "Sub P()\nDim x As Document\nSet x = ActiveDocument\nIf X Then Debug.Print 1 Else Debug.Print 2\nEnd Sub",
        "word",
    )
    assert len(document) == 1
    assert document[0][0] == "assignmentTypeMismatch"


# --- objectReadTokenScans.test.ts ---


def test_reports_a_whole_let_value_once() -> None:
    source, hits = _run_body("Dim c As New Collection\nDim v As Variant\nv = c")
    assert len(hits) == 1
    assert "'450'" in hits[0][1]
    assert source[hits[0][2].start : hits[0][2].end] == "c"


def test_preserves_whole_print_items_before_operator_reads_in_diagnostic_order() -> None:
    source, hits = _run_body("Dim ws As Worksheet\nDebug.Print ws; ws + 1; ws")
    at = source.index("Debug.Print")
    first = source.index("ws", at)
    middle = source.index("ws", first + 2)
    last = source.index("ws", middle + 2)
    assert [hit[2].start for hit in hits] == [first, last, middle]
    assert all(hit[0] == "objectDefaultValue" for hit in hits)


def test_keeps_repeated_operator_operands_as_separate_reads() -> None:
    source, hits = _run_body("Dim ws As Worksheet\nDebug.Print ws + ws + ws")
    assert len(hits) == 3
    assert len({hit[2].start for hit in hits}) == 3
    assert all(source[hit[2].start : hit[2].end] == "ws" for hit in hits)


def test_excludes_indexed_late_bound_values_while_retaining_whole_builtin_arguments_and_typed_indexing() -> (
    None
):
    assert _run_body("Dim x As Object\nDim v As Variant\nSet x = New Collection\nv = x(1)")[1] == []
    whole = _run_body("Dim x As Object\nSet x = New Collection\nDebug.Print CStr(x)")[1]
    assert len(whole) == 1
    assert "'450'" in whole[0][1]
    typed = _run_body("Dim ws As Worksheet\nDim v As Variant\nv = ws(1)")[1]
    assert len(typed) == 1
    assert "'438'" in typed[0][1]
