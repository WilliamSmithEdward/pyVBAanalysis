"""Built-in VBA runtime metadata (ported from src/analyzer/runtime)."""

from __future__ import annotations

from .vba_runtime import (
    VBA_RUNTIME_FUNCTIONS,
    VbaRuntimeConstant,
    VbaRuntimeFunction,
    VbaRuntimeObject,
    VbaRuntimeParam,
    resolve_runtime_constant,
    resolve_runtime_function,
    resolve_runtime_object,
    resolve_runtime_object_type,
    runtime_allows_explicit_call,
)

__all__ = [
    "VBA_RUNTIME_FUNCTIONS",
    "VbaRuntimeConstant",
    "VbaRuntimeFunction",
    "VbaRuntimeObject",
    "VbaRuntimeParam",
    "resolve_runtime_constant",
    "resolve_runtime_function",
    "resolve_runtime_object",
    "resolve_runtime_object_type",
    "runtime_allows_explicit_call",
]
