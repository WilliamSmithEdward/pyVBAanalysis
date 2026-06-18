"""The active VBA diagnostics engine entry point (MS-VBAL Phase 5).

Ported from analyzeModule.ts. analyze_module(source, opts) parses one module,
builds its symbols and conditional-compilation activity, then drives every rule in
the ordered DIAGNOSTIC_RULE_REGISTRY, buffering per rule and flushing in registry
order. It never throws: any internal failure yields an empty list.

The editor-only edit-span helpers (incompleteExpressionEditSpan and friends) are
completion-UI and out of scope. The member-completion context is deferred until
the member rules and host layer land.
"""

from __future__ import annotations

from collections.abc import Mapping

from ..conditional import create_conditional_activity_tracker
from ..parser.nodes import Span
from ..parser.parse_module import parse_module
from ..symbols.build_module_symbols import BuildModuleSymbolsOptions, build_module_symbols
from ..symbols.symbol_model import ModuleSymbolKind
from .context import AnalyzeModuleOptions, PushFn, RulePassContext
from .exprwalk import ProcedureExpressionVisitor, walk_procedure_expressions
from .model import DiagnosticSeverity, VbaDiagnostic, VbaDiagnosticData
from .registry import DIAGNOSTIC_RULE_REGISTRY
from .rule_metadata import DIAGNOSTIC_RULES, normalize_diagnostic_severity_override
from .walker import ProcedureStatementVisitor, walk_procedure_statements


def _severity_of(rule_name: str, overrides: Mapping[str, str] | None) -> DiagnosticSeverity | None:
    """Effective severity of a rule, or None when switched off."""
    meta = DIAGNOSTIC_RULES[rule_name]
    override_value = overrides.get(meta.code) if overrides is not None else None
    override = normalize_diagnostic_severity_override(meta.code, override_value)
    if override == "off":
        return None
    if override is not None:
        return DiagnosticSeverity(override)
    return meta.default_severity


def analyze_module(source: str, opts: AnalyzeModuleOptions | None = None) -> list[VbaDiagnostic]:
    """Analyze one VBA module source and return its active diagnostics. Never throws."""
    options = opts if opts is not None else AnalyzeModuleOptions()
    try:
        return _run_rules(source, options)
    except Exception:
        return []


def _run_rules(source: str, opts: AnalyzeModuleOptions) -> list[VbaDiagnostic]:
    module_name = opts.module_name or "Module"
    module_kind = opts.module_kind or ModuleSymbolKind.STANDARD
    overrides = opts.severity_overrides

    def push_into(sink: list[VbaDiagnostic]) -> PushFn:
        def push(
            rule: str, message: str, span: Span, data: VbaDiagnosticData | None = None
        ) -> None:
            severity = _severity_of(rule, overrides)
            if severity is None:
                return
            meta = DIAGNOSTIC_RULES[rule]
            sink.append(
                VbaDiagnostic(
                    code=meta.code,
                    message=message,
                    severity=severity,
                    span=span,
                    spec_reference=meta.spec_reference,
                    data=data,
                )
            )

        return push

    mod = opts.parsed_module if opts.parsed_module is not None else parse_module(source)
    ctx = RulePassContext(
        source=source,
        module_name=module_name,
        module_kind=module_kind,
        opts=opts,
        mod=mod,
        symbols=build_module_symbols(
            module_name,
            module_kind,
            source,
            BuildModuleSymbolsOptions(
                conditional_compilation=opts.conditional_compilation, parsed_module=mod
            ),
        ),
        activity=create_conditional_activity_tracker(mod, opts.conditional_compilation),
    )

    # Each rule reports into its own buffer; per-statement and per-expression rules
    # ride one shared walk each. Flushing buffers in registry order preserves the
    # rule-major diagnostic output order (a hard contract).
    buffers: list[list[VbaDiagnostic]] = []
    statement_visitors: list[ProcedureStatementVisitor] = []
    expression_visitors: list[ProcedureExpressionVisitor] = []
    for rule in DIAGNOSTIC_RULE_REGISTRY:
        buffer: list[VbaDiagnostic] = []
        buffers.append(buffer)
        push = push_into(buffer)
        if rule.run is not None:
            rule.run(ctx, push)
        if rule.procedure_statements is not None:
            statement_visitors.append(rule.procedure_statements(ctx, push))
        if rule.procedure_expressions is not None:
            expression_visitors.append(rule.procedure_expressions(ctx, push))

    walk_procedure_statements(ctx.mod, ctx.activity, statement_visitors)
    walk_procedure_expressions(ctx.mod, ctx.activity, expression_visitors)

    out: list[VbaDiagnostic] = []
    for buffer in buffers:
        out.extend(buffer)
    return out
