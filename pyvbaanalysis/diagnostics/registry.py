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
from .rules.duplicates import (
    check_duplicate_declarations,
    check_duplicate_enum_members,
    check_duplicate_module_members,
    check_duplicate_procedures,
    check_duplicate_type_fields,
)
from .rules.lexical import check_invalid_line_continuations, check_unterminated_strings
from .walker import ProcedureStatementVisitor


@dataclass(frozen=True, slots=True)
class DiagnosticRuleEntry:
    """One registered rule: a name and exactly one of the three execution forms."""

    name: str
    run: Callable[[RulePassContext, PushFn], None] | None = None
    procedure_statements: Callable[[RulePassContext, PushFn], ProcedureStatementVisitor] | None = None
    procedure_expressions: Callable[[RulePassContext, PushFn], ProcedureExpressionVisitor] | None = None


# The ordered table of active rules. The ORDER is the diagnostic output-order
# contract; entries are placed at their registry.ts positions as families are
# ported (gaps remain for not-yet-ported families).
DIAGNOSTIC_RULE_REGISTRY: tuple[DiagnosticRuleEntry, ...] = (
    DiagnosticRuleEntry(
        name="unterminatedStrings",
        run=lambda ctx, push: check_unterminated_strings(ctx.source, push),
    ),
    DiagnosticRuleEntry(
        name="invalidLineContinuations",
        run=lambda ctx, push: check_invalid_line_continuations(ctx.source, push),
    ),
    DiagnosticRuleEntry(
        name="duplicateProcedures",
        run=lambda ctx, push: check_duplicate_procedures(ctx.symbols.root.children or [], push),
    ),
    DiagnosticRuleEntry(
        name="duplicateDeclarations",
        run=lambda ctx, push: check_duplicate_declarations(ctx.symbols.root.children or [], push),
    ),
    DiagnosticRuleEntry(
        name="duplicateModuleMembers",
        run=lambda ctx, push: check_duplicate_module_members(ctx.symbols.root.children or [], push),
    ),
    DiagnosticRuleEntry(
        name="duplicateEnumMembers",
        run=lambda ctx, push: check_duplicate_enum_members(ctx.source, ctx.mod, ctx.activity, push),
    ),
    DiagnosticRuleEntry(
        name="duplicateTypeFields",
        run=lambda ctx, push: check_duplicate_type_fields(ctx.source, ctx.mod, ctx.activity, push),
    ),
    # NOTE: ambiguousEnumMemberReferences (registry position 12) is deferred -
    # it needs type inference + host + runtime resolution (M8/M9).
)
