"""Ported from xlide_vscode/src/analyzer/diagnostics/reachingSnapshot.ts."""

from __future__ import annotations

from collections.abc import Iterator, Mapping, Sequence
from typing import TYPE_CHECKING

from ..lexer.token_kinds import VbaToken

if TYPE_CHECKING:
    from .straight_line_values import ReachingAssignments

# Small states keep native dict behavior. Large states share their unchanged
# base and copy at most 32 overrides, with one bounded level of lookup.
SHARED_STATE_MIN_SIZE = 64
MAX_OVERRIDES = 32

class ReachingSnapshot(Mapping[str, Sequence[VbaToken]]):
    """A read-only state: a shared base dict plus a few overriding values."""

    __slots__ = ("_base", "_changes", "_size")

    def __init__(
        self,
        base: Mapping[str, Sequence[VbaToken]],
        changes: dict[str, Sequence[VbaToken]],
        size: int,
    ) -> None:
        self._base = base
        self._changes = changes
        self._size = size

    def __len__(self) -> int:
        return self._size

    def __getitem__(self, key: str) -> Sequence[VbaToken]:
        changes = self._changes
        return changes[key] if key in changes else self._base[key]

    def get(self, key: str, default: Sequence[VbaToken] | None = None) -> Sequence[VbaToken] | None:  # type: ignore[override]
        value = self._changes.get(key)
        if value is not None:
            return value
        return self._base.get(key, default)

    def __contains__(self, key: object) -> bool:
        return key in self._changes or key in self._base

    def with_value(self, key: str, value: Sequence[VbaToken]) -> ReachingAssignments:
        if key not in self._changes and len(self._changes) >= MAX_OVERRIDES:
            materialized = dict(self.entries())
            materialized[key] = value
            return materialized
        changes = dict(self._changes)
        changes[key] = value
        return ReachingSnapshot(self._base, changes, self._size + (0 if key in self else 1))

    def entries(self) -> Iterator[tuple[str, Sequence[VbaToken]]]:
        changes = self._changes
        for key, value in self._base.items():
            yield key, changes.get(key, value)
        for key, value in changes.items():
            if key not in self._base:
                yield key, value

    def __iter__(self) -> Iterator[str]:
        for key, _value in self.entries():
            yield key


def with_reaching_value(
    before: ReachingAssignments, key: str, value: Sequence[VbaToken]
) -> ReachingAssignments:
    """Change one value without copying every unaffected fact into the snapshot."""
    if isinstance(before, ReachingSnapshot):
        return before.with_value(key, value)
    if len(before) < SHARED_STATE_MIN_SIZE:
        updated = dict(before)
        updated[key] = value
        return updated
    # Detach the shared base once: callers may own a mutable initial dict.
    return ReachingSnapshot(dict(before), {key: value}, len(before) + (0 if key in before else 1))
