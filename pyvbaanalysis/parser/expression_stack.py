"""Run expression readers on an explicit stack instead of Python call frames."""

from __future__ import annotations

from collections.abc import Generator
from typing import Any, TypeVar, cast

_T = TypeVar("_T")


def run_expression(task: Generator[Any, Any, _T]) -> _T:
    """A reader yields a child reader and receives its result before resuming."""
    stack = [task]
    value: Any = None
    error: BaseException | None = None
    try:
        while stack:
            try:
                child = stack[-1].send(value) if error is None else stack[-1].throw(error)
            except StopIteration as stop:
                stack.pop()
                value = stop.value
                error = None
            except BaseException as failure:
                stack.pop()
                if not stack:
                    raise
                error = failure
                value = None
            else:
                stack.append(child)
                value = None
                error = None
    finally:
        for pending in reversed(stack):
            pending.close()
    return cast(_T, value)
