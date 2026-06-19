"""Shared VBA call-site/context helpers (ported from src/analyzer/call)."""

from __future__ import annotations

from .call_context import (
    BareCallStatementTarget,
    ExplicitCallStatementArgumentList,
    ParenthesizedCallStatementTarget,
    bare_call_statement_target,
    explicit_call_statement_argument_without_parens,
    explicit_call_statement_target,
    standalone_empty_parenthesized_call_statement,
)

__all__ = [
    "BareCallStatementTarget",
    "ExplicitCallStatementArgumentList",
    "ParenthesizedCallStatementTarget",
    "bare_call_statement_target",
    "explicit_call_statement_argument_without_parens",
    "explicit_call_statement_target",
    "standalone_empty_parenthesized_call_statement",
]
