"""M7: the shared straight-line / branch-merged dataflow foundation (dataflow.ts parity).

The branch-merge join is the hardest single algorithm in the analyzer and the
no-false-positive soundness of object-variable-not-set and unallocated-array
depend on it, so it is gated here directly with a synthetic unset->set tracker
before any rule rides on it.
"""

from __future__ import annotations

from pyvbaanalysis.diagnostics.dataflow import (
    DataflowHooks,
    Lattice,
    tracked_locals_named_whole,
    walk_branch_merged_body,
    walk_straight_line_body,
)
from pyvbaanalysis.lexer.token_helpers import statement_tokens, token_name
from pyvbaanalysis.parser.nodes import ProcedureNode
from pyvbaanalysis.parser.parse_module import parse_module


def _first_procedure_body(source: str) -> list:
    mod = parse_module(source)
    for member in mod.members:
        if isinstance(member, ProcedureNode):
            return member.body
    raise AssertionError("no procedure in source")


def _assigns_x(source: str, stmt) -> bool:
    toks = statement_tokens(source, stmt.span.start, stmt.span.end)
    return len(toks) >= 2 and (token_name(toks[0]) or "").lower() == "x" and toks[1].raw_text == "="


def _track_x(source: str, *, merged: bool) -> str:
    """Run the chosen walk with an unset->set tracker on local `x`; return x's final state."""
    state = {"x": "unset"}

    def on_statement(stmt) -> None:
        if _assigns_x(source, stmt):
            state["x"] = "set"

    def touches(stmt):
        return ["x"] if _assigns_x(source, stmt) else []

    def demote(lower: str) -> None:
        if state.get(lower) == "unset":
            state[lower] = "unknown"

    def restore(snapshot) -> None:
        state.clear()
        state.update(snapshot)

    hooks = DataflowHooks(
        on_statement=on_statement,
        touches_in_statement=touches,
        demote_to_unknown=demote,
        snapshot_state=lambda: dict(state),
        restore_state=restore,
        set_state=lambda key, value: state.__setitem__(key, value),
        lattice=Lattice(init="unset", good="set", unknown="unknown"),
    )
    walk = walk_branch_merged_body if merged else walk_straight_line_body
    walk(_first_procedure_body(source), lambda node: False, hooks)
    return state["x"]


def test_straight_line_set() -> None:
    src = "Sub S\n    x = 1\nEnd Sub"
    assert _track_x(src, merged=False) == "set"
    assert _track_x(src, merged=True) == "set"


def test_if_without_else_demotes_on_both_walks() -> None:
    src = "Sub S\n    If c Then\n        x = 1\n    End If\nEnd Sub"
    assert _track_x(src, merged=False) == "unknown"
    assert _track_x(src, merged=True) == "unknown"


def test_balanced_if_set_on_every_arm_merges_to_good() -> None:
    # The precision win: branch merge proves x is set on every path.
    src = "Sub S\n    If c Then\n        x = 1\n    Else\n        x = 2\n    End If\nEnd Sub"
    assert _track_x(src, merged=True) == "set"
    # The conservative straight-line walk demotes it.
    assert _track_x(src, merged=False) == "unknown"


def test_if_set_on_one_arm_only_merges_to_unknown() -> None:
    src = "Sub S\n    If c Then\n        x = 1\n    Else\n        y = 2\n    End If\nEnd Sub"
    assert _track_x(src, merged=True) == "unknown"


def test_untouched_name_keeps_entry_state_across_block() -> None:
    # x is set before an If that never touches x, so it stays set under both walks.
    src = "Sub S\n    x = 1\n    If c Then\n        y = 2\n    End If\nEnd Sub"
    assert _track_x(src, merged=True) == "set"
    assert _track_x(src, merged=False) == "set"


def _significant(source: str) -> list:
    return statement_tokens(source, 0, len(source))


def test_tracked_locals_named_whole() -> None:
    tracked = {"x", "y"}

    def named(source: str, read_only: frozenset[str] = frozenset()) -> dict[str, int]:
        return tracked_locals_named_whole(_significant(source), 0, lambda name: name in tracked, read_only)

    assert named("Helper x") == {"x": 7}
    assert set(named("Call Helper(x, y)")) == {"x", "y"}
    # A top-level assignment is not a call statement, but an argument inside it is
    # still passed by reference (XLIDE issue #70).
    assert named("x = 1") == {}
    assert set(named("n = Fill(x)")) == {"x"}
    # A method call passes its arguments too; a member access or an indexed
    # element is not the variable itself.
    assert set(named("obj.Method x")) == {"x"}
    assert named("Helper a.x") == {}
    assert named("Helper x(0)") == {}
    # The operand of Is and the argument of a read-only intrinsic are only read.
    assert named("If x Is Nothing Then Helper") == {}
    assert named("n = UBound(x)", frozenset({"ubound"})) == {}
    # The first mention's offset tells an access before the pass from one after it.
    assert named("If Load(x) Then Debug.Print x(0)") == {"x": 8}
