"""Bounded reaching snapshots. Ported from upstream's tests/reachingSnapshot.test.ts."""

from __future__ import annotations

from collections.abc import Sequence

import pytest

from pyvbaanalysis.diagnostics.reaching_snapshot import with_reaching_value
from pyvbaanalysis.diagnostics.straight_line_values import ReachingAssignments
from pyvbaanalysis.diagnostics.walker import raw_expression_tokens
from pyvbaanalysis.lexer.token_kinds import VbaToken


def _value(n: int) -> Sequence[VbaToken]:
    return raw_expression_tokens(str(n))


def _initial(size: int) -> dict[str, Sequence[VbaToken]]:
    return {f"k{i}": _value(i) for i in range(size)}


def _compare(actual: ReachingAssignments, expected: ReachingAssignments) -> None:
    assert len(actual) == len(expected)
    assert list(actual.items()) == list(expected.items())
    assert list(actual.keys()) == list(expected.keys())
    assert list(actual.values()) == list(expected.values())
    for key, tokens in expected.items():
        assert key in actual
        assert actual.get(key) is tokens
        assert actual[key] is tokens
    assert "absent" not in actual
    assert actual.get("absent") is None


@pytest.mark.parametrize("size", [0, 63, 64, 1000])
def test_lookup_and_ordered_iteration_survive_updates_and_compaction(size: int) -> None:
    actual: ReachingAssignments = _initial(size)
    expected = dict(actual)
    snapshots: list[tuple[ReachingAssignments, dict[str, Sequence[VbaToken]]]] = []
    for i in range(150):
        if i % 20 == 0:
            snapshots.append((actual, dict(expected)))
        key = "k0" if i % 3 == 0 else f"new{i}"
        tokens = _value(i + 5000)
        actual = with_reaching_value(actual, key, tokens)
        expected[key] = tokens
        if i % 20 == 0 or i == 149:
            _compare(actual, expected)
    for kept, expected_then in snapshots:
        _compare(kept, expected_then)


def test_a_shared_base_is_detached_from_later_changes_to_the_initial_map() -> None:
    base = _initial(100)
    old = with_reaching_value(base, "n", _value(1))
    base.clear()
    assert len(old) == 101
    k0 = old.get("k0")
    assert k0 is not None and k0[0].raw_text == "0"


def test_a_materialized_state_keeps_deletion_and_reinsertion_order() -> None:
    base = _initial(100)
    old = with_reaching_value(base, "k3", _value(300))
    invalidated = dict(old)
    del invalidated["k3"]
    del invalidated["k4"]
    after = with_reaching_value(invalidated, "k3", _value(400))
    assert list(after.keys())[-1] == "k3"
    assert "k4" not in after
    k3 = old.get("k3")
    assert k3 is not None and k3[0].raw_text == "300"
    assert "k4" in old


def test_repeated_updates_of_one_local_share_the_large_base() -> None:
    first = with_reaching_value(_initial(1000), "n", _value(0))
    state = first
    for i in range(1, 1000):
        state = with_reaching_value(state, "n", _value(i))
    n = state.get("n")
    assert n is not None and n[0].raw_text == "999"
    assert type(state) is type(first) and type(state) is not dict
    assert len(state) == 1001
