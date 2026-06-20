"""The active VBA diagnostics engine entry point (MS-VBAL Phase 5).

Ported from analyzeModule.ts. analyze_module(source, opts) parses one module,
builds its symbols and conditional-compilation activity, then drives every rule in
the ordered DIAGNOSTIC_RULE_REGISTRY, buffering per rule and flushing in registry
order. It never throws: any internal failure yields an empty list.

The member-completion context is assembled once per pass here
(diagnostic_member_completion_context) and shared through RulePassContext.member_ctx;
the member, object-state, call-shape, and type rules consume it. The editor-only
edit-span helpers (incompleteExpressionEditSpan and friends) are completion-UI and
out of scope.
"""

from __future__ import annotations

from collections.abc import Mapping

from ..completion import MemberCompletionContext
from ..conditional import create_conditional_activity_tracker
from ..lexer.token_kinds import TokenKind
from ..lexer.tokenize import tokenize
from ..parser.nodes import ModuleNode, Span
from ..parser.parse_module import parse_module
from ..symbols.build_module_symbols import BuildModuleSymbolsOptions, build_module_symbols
from ..symbols.symbol_model import ModuleSymbolKind
from .context import (
    AnalyzeModuleOptions,
    PushFn,
    RulePassContext,
    is_object_module_kind,
)
from .exprwalk import ProcedureExpressionVisitor, walk_procedure_expressions
from .inline_suppression import (
    DIRECTIVE_DIAGNOSTIC_CODE,
    filter_inline_suppressions,
    scan_inline_suppressions,
)
from .model import DiagnosticSeverity, VbaDiagnostic, VbaDiagnosticData
from .registry import DIAGNOSTIC_RULE_REGISTRY
from .rule_metadata import (
    DIAGNOSTIC_RULES,
    diagnostic_metadata_for_code,
    normalize_diagnostic_severity_override,
)
from .walker import ProcedureStatementVisitor, walk_procedure_statements


def _severity_of(
    rule_name: str, overrides: Mapping[str, str] | None, whole_project: bool = True
) -> DiagnosticSeverity | None:
    """Effective severity of a rule, or None when switched off."""
    meta = DIAGNOSTIC_RULES[rule_name]
    # A rule that needs every module to be correct (undeclared-variable, unknown-call,
    # member-not-found) stays silent on a partial project view: a symbol declared in an
    # unseen module is indistinguishable from an undefined one, so it would false-positive.
    if meta.requires_whole_project and not whole_project:
        return None
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
    # Match override codes case-insensitively: codes are canonically lowercase, and
    # validate_severity_overrides resolves them case-insensitively, so the apply path
    # must too (otherwise a mis-cased key validates but is then silently ignored).
    overrides = (
        {code.lower(): value for code, value in opts.severity_overrides.items()}
        if opts.severity_overrides is not None
        else None
    )
    whole_project = opts.whole_project

    def push_into(sink: list[VbaDiagnostic]) -> PushFn:
        def push(
            rule: str, message: str, span: Span, data: VbaDiagnosticData | None = None
        ) -> None:
            severity = _severity_of(rule, overrides, whole_project)
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
        member_ctx=diagnostic_member_completion_context(opts, source, mod),
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
    if not opts.inline_suppression:
        return out
    # Drop the diagnostics that '@pyvba-ignore directives suppress, then surface any
    # malformed directive as an analysis-suppression-directive diagnostic (itself
    # subject to severity overrides, never to inline suppression).
    scan = scan_inline_suppressions(source)
    out = filter_inline_suppressions(source, out, scan)
    if scan.issues:
        directive_meta = diagnostic_metadata_for_code(DIRECTIVE_DIAGNOSTIC_CODE)
        severity = (
            _severity_of(directive_meta.rule_name, overrides, whole_project)
            if directive_meta is not None
            else None
        )
        if directive_meta is not None and severity is not None:
            out.extend(
                VbaDiagnostic(
                    code=DIRECTIVE_DIAGNOSTIC_CODE,
                    message=message,
                    severity=severity,
                    span=span,
                    spec_reference=directive_meta.spec_reference,
                )
                for span, message in scan.issues
            )
    return out


def diagnostic_member_completion_context(
    opts: AnalyzeModuleOptions, source: str, mod: ModuleNode
) -> MemberCompletionContext:
    """Assemble the per-pass member-resolution context (analysisContext.ts mirror).

    Hard diagnostics disable Set-assignment refinement (VBE leaves those receivers
    late-bound). The context is primed with the per-pass AST and the shared
    full-source token stream so member resolution never re-parses or re-lexes per
    dotted reference. `me_project_type`/`me_type` are derived from the module
    identity; `code_names` is left unset (the diagnostics pass has no code-name map).
    """
    ctx = MemberCompletionContext(
        project_class_members=opts.project_class_members,
        allow_set_assignment_refinement=False,
        model=opts.host_model,
        parsed_module=mod,
        source_tokens=[t for t in tokenize(source) if t.kind is not TokenKind.COMMENT],
    )
    me_project_type = _me_project_type_for(opts.module_name, opts.module_kind)
    if me_project_type:
        ctx.me_project_type = me_project_type
    me_type = _me_host_type_for(opts.module_name, opts.module_kind)
    if me_type:
        ctx.me_type = me_type
    return ctx


def _me_project_type_for(
    module_name: str | None, module_kind: ModuleSymbolKind | None
) -> str | None:
    return module_name if module_name and is_object_module_kind(module_kind) else None


def _me_host_type_for(
    module_name: str | None, module_kind: ModuleSymbolKind | None
) -> str | None:
    if not module_name or module_kind is not ModuleSymbolKind.DOCUMENT:
        return None
    return "Excel.Workbook" if module_name.lower() == "thisworkbook" else None
