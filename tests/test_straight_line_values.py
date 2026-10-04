"""The straight-line walk: what reaches each statement, and what never runs."""

from __future__ import annotations

from collections.abc import Sequence

from pyvbaanalysis.diagnostics.straight_line_values import (
    ReachingAssignments,
    identity_assignment,
    straight_line_assignments,
    straight_line_exit,
    straight_line_unreachable,
)
from pyvbaanalysis.diagnostics.walker import raw_expression_tokens
from pyvbaanalysis.parser.nodes import BodyNode, ProcedureNode, is_leaf_statement, iter_body_nodes
from pyvbaanalysis.parser.parse_module import parse_module


def _procedure(source: str) -> ProcedureNode:
    for member in parse_module(source).members:
        if isinstance(member, ProcedureNode):
            return member
    raise AssertionError("no procedure")


def _held(state: ReachingAssignments | None) -> dict[str, str]:
    return {} if state is None else {k: " ".join(t.raw_text for t in v) for k, v in state.items()}


def _leaf(body: Sequence[BodyNode], source: str, text: str) -> BodyNode:
    for node in iter_body_nodes(body):
        if is_leaf_statement(node) and source[node.span.start : node.span.end] == text:
            return node
    raise AssertionError(text)


def test_a_later_assignment_does_not_hide_the_value_that_reaches() -> None:
    source = "Sub S()\n    d = 0\n    x = 10 / d\n    d = 2\n    y = d\nEnd Sub\n"
    proc = _procedure(source)
    reaching = straight_line_assignments(source, proc.body, None)
    assert _held(reaching.get(id(_leaf(proc.body, source, "x = 10 / d")))) == {"d": "0"}
    assert _held(reaching.get(id(_leaf(proc.body, source, "y = d")))) == {"d": "2", "x": "10 / d"}


def test_a_known_guard_leaves_the_rest_unreachable() -> None:
    source = (
        "Function F(d As Long) As Long\n    d = 0\n    If d = 0 Then Exit Function\n"
        "    F = 1 / d\nEnd Function\n"
    )
    proc = _procedure(source)
    dead = straight_line_unreachable(source, proc.body, None)
    assert id(_leaf(proc.body, source, "F = 1 / d")) in dead
    assert straight_line_exit(source, proc.body, None, {}).exit is None


def test_a_for_counter_holds_one_step_past_its_last_pass() -> None:
    source = "Sub S()\n    For i = 0 To 3\n        n = i\n    Next\n    x = i\nEnd Sub\n"
    proc = _procedure(source)
    reaching = straight_line_assignments(source, proc.body, None)
    assert _held(reaching.get(id(_leaf(proc.body, source, "x = i"))))["i"] == "4"


def test_a_do_counter_runs_to_its_condition() -> None:
    source = "Sub S()\n    i = 1\n    Do While i <= 5\n        i = i + 1\n    Loop\n    x = i\nEnd Sub\n"
    proc = _procedure(source)
    reaching = straight_line_assignments(source, proc.body, None)
    assert _held(reaching.get(id(_leaf(proc.body, source, "x = i"))))["i"] == "6"


def test_a_known_select_runs_one_arm() -> None:
    source = (
        "Sub S()\n    d = 0\n    Select Case d\n    Case 1\n        a = 1\n    Case 0\n        a = 2\n"
        "    End Select\nEnd Sub\n"
    )
    proc = _procedure(source)
    dead = straight_line_unreachable(source, proc.body, None)
    assert id(_leaf(proc.body, source, "a = 1")) in dead
    assert id(_leaf(proc.body, source, "a = 2")) not in dead
    exit_state = straight_line_exit(source, proc.body, None, {}).exit
    assert _held(exit_state) == {"d": "0", "a": "2"}


def test_identity_assignment() -> None:
    assert identity_assignment("d", raw_expression_tokens("d + 0"))
    assert identity_assignment("D", raw_expression_tokens("1 * d"))
    assert not identity_assignment("d", raw_expression_tokens("d + 1"))
    assert not identity_assignment("d", raw_expression_tokens("0 - d"))


def test_blocks_nested_past_the_recursion_limit_are_walked() -> None:
    # 1100 nested blocks, past the recursion limit.
    depth = 550
    source = (
        "Sub S()\n    n = 0\n"
        + "    If n = 0 Then\n    For i = 1 To 2\n" * depth
        + "    x = n\n"
        + "    Next\n    End If\n" * depth
        + "End Sub\n"
    )
    proc = _procedure(source)
    reaching = straight_line_assignments(source, proc.body, None)
    assert _held(reaching.get(id(_leaf(proc.body, source, "x = n"))))["n"] == "0"
