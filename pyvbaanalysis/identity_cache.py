"""A small identity-keyed memo for per-pass derived data.

The analyzer derives many views (name scopes, type environments, symbol
indexes) from objects that live for one analysis pass (ModuleSymbols, parsed
procedures, the project symbol list). Rebuilding a view per rule multiplies a
per-module cost by the rule count - the dominant cost on large modules - so
these views are memoized by the *identity* of the objects they derive from.

Identity keying mirrors the upstream WeakMap caches without requiring the key
objects to be hashable or weak-referenceable: entries hold a strong reference
to their key objects (so an id() can never be recycled while cached) and the
cache is a small MRU list, so at most `capacity` passes' worth of derived data
is ever retained.
"""

from __future__ import annotations

from typing import Any


class IdentityLru:
    """MRU-ordered memo keyed by the identities of one or more key objects."""

    __slots__ = ("_capacity", "_entries")

    def __init__(self, capacity: int = 8) -> None:
        self._capacity = capacity
        # (id-tuple, key objects kept alive, value); most recently used first.
        self._entries: list[tuple[tuple[int, ...], tuple[Any, ...], Any]] = []

    def get(self, *keys: Any) -> Any | None:
        ids = tuple(id(key) for key in keys)
        for i, (entry_ids, _keepalive, value) in enumerate(self._entries):
            if entry_ids == ids:
                if i > 0:
                    self._entries.insert(0, self._entries.pop(i))
                return value
        return None

    def put(self, value: Any, *keys: Any) -> Any:
        ids = tuple(id(key) for key in keys)
        self._entries.insert(0, (ids, keys, value))
        if len(self._entries) > self._capacity:
            self._entries.pop()
        return value
