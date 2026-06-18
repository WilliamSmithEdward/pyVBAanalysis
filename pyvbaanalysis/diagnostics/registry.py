"""The ordered diagnostic rule registry.

Ported from registry.ts. A rule entry is a stable name plus exactly one execution
form (run / procedure_statements / procedure_expressions). The registry ORDER is a
hard contract: it is the diagnostic output order (run_rules buffers per rule and
flushes in registry order). Rule families are appended here as they are ported
(M6+); the engine skeleton starts with an empty registry.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from .context import PushFn, RulePassContext
from .exprwalk import ProcedureExpressionVisitor
from .walker import ProcedureStatementVisitor


@dataclass(frozen=True, slots=True)
class DiagnosticRuleEntry:
    """One registered rule: a name and exactly one of the three execution forms."""

    name: str
    run: Callable[[RulePassContext, PushFn], None] | None = None
    procedure_statements: Callable[[RulePassContext, PushFn], ProcedureStatementVisitor] | None = None
    procedure_expressions: Callable[[RulePassContext, PushFn], ProcedureExpressionVisitor] | None = None


# The ordered table of all active rules. Empty until rule families are ported.
DIAGNOSTIC_RULE_REGISTRY: tuple[DiagnosticRuleEntry, ...] = ()
