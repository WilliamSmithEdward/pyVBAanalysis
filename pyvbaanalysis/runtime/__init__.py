"""Built-in VBA runtime metadata (ported from src/analyzer/runtime)."""

from __future__ import annotations

from .vba_runtime import (
    VBA_RUNTIME_FUNCTIONS,
    VbaRuntimeFunction,
    VbaRuntimeParam,
    resolve_runtime_function,
    runtime_allows_explicit_call,
)

__all__ = [
    "VBA_RUNTIME_FUNCTIONS",
    "VbaRuntimeFunction",
    "VbaRuntimeParam",
    "resolve_runtime_function",
    "runtime_allows_explicit_call",
]
