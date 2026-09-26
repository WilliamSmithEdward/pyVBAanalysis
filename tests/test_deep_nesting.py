"""Deeply nested blocks: the parser and the body walks run on explicit stacks.

Python stops a recursion at 1,000 frames by default. The parser spent two frames
per nested block, so analyze_project raised RecursionError on a procedure
nested 493 deep, and a score of rule and symbol walks spent one frame per level.
The VBE compiles blocks nested a thousand deep. These modules nest past the
limit and check that the analysis completes and still reaches the innermost
statements.
"""

from __future__ import annotations

import sys
from types import CodeType, FrameType
from typing import Any

import pytest

from pyvbaanalysis import analyze_project
from pyvbaanalysis.parser.nodes import (
    AssignmentNode,
    ForBlockNode,
    IfBlockNode,
    ProcedureNode,
    iter_body_nodes,
    iter_body_nodes_in_context,
)
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind

# Past the default recursion limit however many frames a walk spends per level.
DEPTH = 1100

# Opening and closing lines of one level of each block kind.
BLOCKS = {
    "If": ("If n = 0 Then", "End If"),
    "Else arm": ("If n = 0 Then\n    Else", "End If"),
    "ElseIf arm": ("If n = 0 Then\n    ElseIf n = 1 Then", "End If"),
    "For": ("For n = 1 To 1", "Next"),
    "For Each": ("For Each v In c", "Next"),
    "Do": ("Do", "Loop Until True"),
    "Do While": ("Do While n = 0", "Loop"),
    "While": ("While n = 0", "Wend"),
    "With": ("With c", "End With"),
    "Select Case": ("Select Case n\n    Case 0", "End Select"),
}


def _procedure_body(source: str) -> list:
    for member in parse_module(source).members:
        if isinstance(member, ProcedureNode):
            return member.body
    raise AssertionError("no procedure in source")


def test_iter_body_nodes_walks_blocks_in_source_order() -> None:
    body = _procedure_body(
        "Sub S()\n    a = 1\n    If x Then\n        b = 2\n        For i = 1 To 2\n"
        "            c = 3\n        Next\n    End If\n    d = 4\nEnd Sub\n"
    )
    walked = [
        node.lhs.name if isinstance(node, AssignmentNode) else type(node).__name__  # type: ignore[attr-defined]
        for node in iter_body_nodes(body)
    ]
    assert walked == ["a", "IfBlockNode", "b", "ForBlockNode", "c", "d"]


def test_iter_body_nodes_leaves_out_a_skipped_block_and_its_body() -> None:
    body = _procedure_body(
        "Sub S()\n    a = 1\n    If x Then\n        b = 2\n    End If\n    d = 4\nEnd Sub\n"
    )
    walked = [type(node).__name__ for node in iter_body_nodes(body, lambda node: isinstance(node, IfBlockNode))]
    assert walked == ["AssignmentNode", "AssignmentNode"]


def test_iter_body_nodes_in_context_gives_each_body_its_own_context() -> None:
    body = _procedure_body(
        "Sub S()\n    a = 1\n    For i = 1 To 2\n        b = 2\n        For j = 1 To 2\n"
        "            c = 3\n        Next\n        d = 4\n    Next\n    e = 5\nEnd Sub\n"
    )
    loops = [
        (node.lhs.name, depth)  # type: ignore[attr-defined]
        for node, depth in iter_body_nodes_in_context(
            body, 0, lambda block, depth: depth + 1 if isinstance(block, ForBlockNode) else depth
        )
        if isinstance(node, AssignmentNode)
    ]
    assert loops == [("a", 0), ("b", 1), ("c", 2), ("d", 1), ("e", 0)]


def _nested_module(opener: str, closer: str) -> str:
    return (
        "Option Explicit\nSub Main()\n    Dim n As Long, c As New Collection, v As Variant\n"
        + f"    {opener}\n" * DEPTH
        + "    Exit Sub\n    n = 2\n    missing = 1\n"
        + f"    {closer}\n" * DEPTH
        + "End Sub\n"
    )


@pytest.mark.parametrize("kind", sorted(BLOCKS))
def test_blocks_nested_past_the_recursion_limit_are_analyzed(kind: str) -> None:
    source = _nested_module(*BLOCKS[kind])
    result = analyze_project([ModuleInput("Module1", ModuleSymbolKind.STANDARD, source)], host="excel")
    found = {(d.code, source[d.span.start : d.span.end]) for d in result["Module1"]}
    # The innermost statements, reached by the unreachable-code walk and the
    # undeclared-variable walk.
    assert ("undeclared-variable", "missing") in found
    assert ("unreachable-code", "n = 2\n    missing = 1") in found


# One level of each kind, with the statements that switch on the dataflow rules
# (a tracked object, an array), the If-arm merge, and the Select and loop walks.
MIXED_LEVELS = [
    ("If n = 0 Then\n        Set o = New Collection\n        ReDim a(1)\n        a(0) = n\n    Else\n"
     "        Set o = Nothing", "        o.Add a(0)\n    End If"),
    ("For n = 1 To 2\n        s = s & CStr(n)", "    Next n"),
    ("With c", "    End With"),
    ("Do While n < 3\n        n = n + 1", "    Loop"),
    ("Select Case n\n    Case 1\n        Set o = Nothing\n    Case Else", "    End Select"),
    ("While n < 3", "        n = n + 1\n    Wend"),
]


def test_no_walk_recurses_once_per_nested_block() -> None:
    """A walk that recursed once per nested block would fail only past the
    recursion limit, and the rule isolation would swallow the error and drop
    that rule's findings. So count, per function, the most activations alive at
    once: across 150 levels no analyzer function comes near once per level."""
    depth = 150
    levels = [MIXED_LEVELS[i % len(MIXED_LEVELS)] for i in range(depth)]
    source = (
        "Option Explicit\nSub Main()\n"
        "    Dim n As Long, c As New Collection, o As Collection, a() As Long, s As String\n"
        + "".join(f"    {opener}\n" for opener, _ in levels)
        + "    Exit Sub\n    n = 2\n"
        + "".join(f"{closer}\n" for _, closer in reversed(levels))
        + "End Sub\n"
        # If blocks nested straight in each other's Else arms: the dataflow
        # merge walks every arm.
        + "Sub Merge()\n    Dim n As Long, o As Collection\n"
        + "    If n = 0 Then\n        Set o = New Collection\n    Else\n" * (depth // 6)
        + "        Set o = Nothing\n"
        + "    End If\n" * (depth // 6)
        + "    o.Add 1\nEnd Sub\n"
    )
    active: dict[CodeType, int] = {}
    peak: dict[CodeType, int] = {}

    def count(frame: FrameType, event: str, _arg: Any) -> None:
        if event == "call":
            code = frame.f_code
            active[code] = active.get(code, 0) + 1
            peak[code] = max(peak.get(code, 0), active[code])
        elif event == "return":
            code = frame.f_code
            active[code] = active.get(code, 1) - 1

    sys.setprofile(count)
    try:
        analyze_project([ModuleInput("Module1", ModuleSymbolKind.STANDARD, source)], host="excel")
    finally:
        sys.setprofile(None)
    # The rarest kind of level, and the Else arms of Merge, still nest 25 deep,
    # so a walk recursing into one kind of block only would reach 25.
    recursive = sorted(
        f"{code.co_filename}:{code.co_firstlineno} {code.co_name} ({n})"
        for code, n in peak.items()
        if n >= 20 and "pyvbaanalysis" in code.co_filename
    )
    assert recursive == []
