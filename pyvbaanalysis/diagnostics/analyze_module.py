"""The active VBA diagnostics engine entry point (MS-VBAL Phase 5).

Ported from analyzeModule.ts. analyze_module(source, opts) parses one module,
builds its symbols and conditional-compilation activity, then drives every rule in
the ordered DIAGNOSTIC_RULE_REGISTRY, buffering per rule and flushing in registry
order. It never raises: a failure of its own costs the findings of what failed and
is reported to opts.on_internal_error. A run-time error found in a statement that
never runs is dropped, as upstream does. The list it returns has then been through
the steps XLIDE takes before showing findings (module_analysis.py), unless
opts.raw_rule_output asks for the rules' own list.

The member-completion context is assembled once per pass here
(diagnostic_member_completion_context) and shared through RulePassContext.member_ctx;
the member, object-state, call-shape, and type rules consume it.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping
from dataclasses import replace

from ..completion import MemberCompletionContext
from ..conditional import create_conditional_activity_tracker
from ..host.host_registry import (
    host_knowledge_is_absent,
    host_object_model_for_token,
    host_object_model_for_tokens,
)
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize_cached
from ..parser.nodes import ExprNode, LeafStatementNode, ModuleNode, ProcedureNode, Span, iter_body_nodes
from ..parser.parse_module import parse_module
from ..symbols.build_module_symbols import BuildModuleSymbolsOptions, build_module_symbols
from ..symbols.symbol_model import ModuleSymbolKind
from ..types.type_inference import dead_branch_spans_in, unreachable_statements_in
from .context import (
    AnalysisFailure,
    AnalyzeModuleOptions,
    PushFn,
    RulePassContext,
    is_object_module_kind,
    statement_tokens,
)
from .exprwalk import ProcedureExpressionVisitor, walk_procedure_expressions
from .inline_suppression import (
    DIRECTIVE_DIAGNOSTIC_CODE,
    filter_inline_suppressions,
    scan_inline_suppressions,
)
from .model import DiagnosticEvidenceKind, DiagnosticSeverity, VbaDiagnostic, VbaDiagnosticData
from .module_analysis import deduplicate_diagnostics, drop_handled_runtime_errors
from .module_state import remember_project_written_names
from .registry import DIAGNOSTIC_RULE_REGISTRY
from .rule_metadata import (
    DIAGNOSTIC_RULES,
    diagnostic_metadata_for_code,
    normalize_diagnostic_severity_override,
)
from .rules.expressions import incomplete_member_access, is_non_unary_binary_operator
from .walker import (
    ProcedureStatementVisitor,
    absolute_span,
    active_module_members,
    physical_line_span_at_offset,
    walk_procedure_statements,
)


def _severity_of(
    rule_name: str,
    overrides: Mapping[str, str] | None,
    whole_project: bool = True,
    host_known: bool = True,
) -> DiagnosticSeverity | None:
    """Effective severity of a rule, or None when switched off."""
    meta = DIAGNOSTIC_RULES[rule_name]
    # A rule that needs every module to be correct (undeclared-variable, unknown-call,
    # member-not-found, late-bound-friend-member) stays silent on a partial project view:
    # a symbol declared in an unseen module is indistinguishable from an undefined one,
    # so it would false-positive. An unmodelled host is the same situation one level up:
    # with no object model, a name the host injects is indistinguishable from an
    # undefined one, so the same rules stay silent rather than report every host global.
    if meta.requires_whole_project and not (whole_project and host_known):
        return None
    override_value = overrides.get(meta.code) if overrides is not None else None
    override = normalize_diagnostic_severity_override(meta.code, override_value)
    if override == "off":
        return None
    if override is not None:
        return DiagnosticSeverity(override)
    return meta.default_severity


def analyze_module(source: str, opts: AnalyzeModuleOptions | None = None) -> list[VbaDiagnostic]:
    """Analyze one VBA module source and return its active diagnostics.

    Never raises: a failure of its own costs the findings of what failed (one rule,
    one rule's walk, or the whole pass) and is reported to opts.on_internal_error,
    so a caller can tell a module that is clean from one that was not fully checked.
    """
    given = opts if opts is not None else AnalyzeModuleOptions()
    report = _internal_error_reporter(given)
    try:
        return _run_rules(
            source, with_resolved_host_model(_with_usable_project_procedures(given, report)), report
        )
    except Exception as err:
        report(err, AnalysisFailure("analysis"))
        return []


_ReportInternalError = Callable[[BaseException, AnalysisFailure], None]


def _internal_error_reporter(opts: AnalyzeModuleOptions) -> _ReportInternalError:
    callback = opts.on_internal_error

    def report(error: BaseException, where: AnalysisFailure) -> None:
        if callback is None:
            return
        try:
            callback(error, where)
        except Exception:
            # The caller's callback failing must not make analysis raise.
            pass

    return report


def _with_usable_project_procedures(
    opts: AnalyzeModuleOptions, report: _ReportInternalError
) -> AnalyzeModuleOptions:
    """project_procedures is a mapping from lowercased name to signatures
    (ProjectIndex.procedure_signatures). A list passed in its place (the editor
    contexts carry one under the same name upstream) made rules raise on first use
    while the rest reported as usual. It is left out and reported instead:
    converting it would not give the same checks, since the list holds the module's
    own private procedures too."""
    procedures: object = opts.project_procedures
    if procedures is None or isinstance(procedures, Mapping):
        return opts
    given = "a list" if isinstance(procedures, (list, tuple)) else type(procedures).__name__
    report(
        TypeError(
            "project_procedures must be a mapping of lowercased name to signatures "
            f"(ProjectIndex.procedure_signatures), not {given}"
        ),
        AnalysisFailure("options"),
    )
    return replace(opts, project_procedures=None)


def _guard_statement_visitor(
    visitor: ProcedureStatementVisitor, rule: str, report: _ReportInternalError
) -> ProcedureStatementVisitor:
    """A rule's statement visitor that cannot stop the shared walk for every other
    rule: its first failure is reported, and the rule sits out the rest of the
    module while the others walk on."""
    failed = False

    def fail(err: BaseException) -> None:
        nonlocal failed
        failed = True
        report(err, AnalysisFailure("statement-walk", rule))

    def guarded(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        if failed:
            return None
        try:
            callback = visitor(member)
        except Exception as err:
            fail(err)
            return None
        if callback is None:
            return None
        visit = callback

        def guarded_visit(stmt: LeafStatementNode) -> None:
            if failed:
                return
            try:
                visit(stmt)
            except Exception as err:
                fail(err)

        return guarded_visit

    return guarded


def _guard_expression_visitor(
    visitor: ProcedureExpressionVisitor, rule: str, report: _ReportInternalError
) -> ProcedureExpressionVisitor:
    """The expression-walk counterpart of _guard_statement_visitor."""
    failed = False

    def fail(err: BaseException) -> None:
        nonlocal failed
        failed = True
        report(err, AnalysisFailure("expression-walk", rule))

    def skip(expr: ExprNode) -> None:
        return None

    def guarded(member: ProcedureNode) -> Callable[[ExprNode], None]:
        if failed:
            return skip
        try:
            visit = visitor(member)
        except Exception as err:
            fail(err)
            return skip

        def guarded_visit(expr: ExprNode) -> None:
            if failed:
                return
            try:
                visit(expr)
            except Exception as err:
                fail(err)

        return guarded_visit

    return guarded


def with_resolved_host_model(opts: AnalyzeModuleOptions) -> AnalyzeModuleOptions:
    """Resolve the `host` token, and any referenced libraries, into a host_model
    once, up front, so every host_model consumer inherits the caller's choice. An
    explicit host_model wins; absent all three, the Excel defaults ride as they
    always have.

    A referenced library is resolved alongside the host, in declaration order,
    which is the order VBA itself resolves an ambiguous name in. An absent host is
    Excel here too: upstream merges only the tokens it is given, so a referenced
    library with no host named would drop Excel from the merge, and "absent means
    Excel" is the contract every other path in this port keeps.
    """
    if opts.host_model is not None:
        return opts
    referenced = list(opts.referenced_hosts or ())
    if opts.host is None and not referenced:
        return opts
    resolved = (
        host_object_model_for_token(opts.host)
        if not referenced
        else host_object_model_for_tokens([opts.host or "excel", *referenced])
    )
    if resolved is None:
        return opts
    return replace(opts, host_model=resolved)


def _run_rules(
    source: str, opts: AnalyzeModuleOptions, report: _ReportInternalError
) -> list[VbaDiagnostic]:
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
    host_known = not host_knowledge_is_absent(opts.host_model)

    def push_into(sink: list[VbaDiagnostic]) -> PushFn:
        def push(
            rule: str, message: str, span: Span, data: VbaDiagnosticData | None = None
        ) -> None:
            severity = _severity_of(rule, overrides, whole_project, host_known)
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
    remember_project_written_names(ctx.symbols, opts.project_written_names)

    # Each rule reports into its own buffer; per-statement and per-expression rules
    # ride one shared walk each. Flushing buffers in registry order preserves the
    # rule-major diagnostic output order (a hard contract).
    buffers: list[list[VbaDiagnostic]] = []
    statement_visitors: list[ProcedureStatementVisitor] = []
    takes_headers: list[bool] = []
    expression_visitors: list[ProcedureExpressionVisitor] = []
    for rule in DIAGNOSTIC_RULE_REGISTRY:
        buffer: list[VbaDiagnostic] = []
        buffers.append(buffer)
        push = push_into(buffer)
        # Isolate each rule: one rule throwing during construction or its eager
        # run() must not discard every other rule's diagnostics for the module.
        try:
            if rule.run is not None:
                rule.run(ctx, push)
            if rule.procedure_statements is not None:
                statement_visitors.append(
                    _guard_statement_visitor(rule.procedure_statements(ctx, push), rule.name, report)
                )
                takes_headers.append(rule.block_headers)
            if rule.procedure_expressions is not None:
                expression_visitors.append(
                    _guard_expression_visitor(rule.procedure_expressions(ctx, push), rule.name, report)
                )
        except Exception as err:
            # Degrade only this rule; keep the rest of the pass intact.
            report(err, AnalysisFailure("rule", rule.name))

    # Each rule's visitor is guarded on its own; what these catch is the walk
    # itself failing, which ends it for every rule from there on.
    try:
        walk_procedure_statements(ctx.mod, ctx.activity, statement_visitors, None, (ctx.source, takes_headers))
    except Exception as err:
        report(err, AnalysisFailure("statement-walk"))
    try:
        walk_procedure_expressions(ctx.mod, ctx.activity, expression_visitors)
    except Exception as err:
        report(err, AnalysisFailure("expression-walk"))

    out: list[VbaDiagnostic] = []
    for buffer in buffers:
        out.extend(buffer)
    out = _without_runtime_errors_in_dead_code(out, ctx)
    if not opts.raw_rule_output:
        out = drop_handled_runtime_errors(out, source, mod, ctx.symbols, ctx.activity)
    if opts.inline_suppression:
        out = _apply_inline_suppression(source, out, overrides, whole_project, host_known)
    return out if opts.raw_rule_output else deduplicate_diagnostics(out)


# The codes of the rules that report an error raised when a statement runs.
_RUNTIME_ERROR_CODES: frozenset[str] = frozenset(
    meta.code
    for meta in DIAGNOSTIC_RULES.values()
    if meta.diagnostic_kind is DiagnosticEvidenceKind.DETERMINISTIC_RUNTIME_ERROR
)


def _without_runtime_errors_in_dead_code(
    diagnostics: list[VbaDiagnostic], ctx: RulePassContext
) -> list[VbaDiagnostic]:
    """Drop a run-time error found in a statement that never runs: under a guard the
    code decides against (`x = 2: If x = 1 Then y = Sqr(-1)`), after `GoTo` or Exit,
    or in a loop of no pass (XLIDE issues #406, #430, measured in Excel 16.0). Each
    rule used to decide this on its own, and the ones that judge a literal did not."""
    if not any(d.code in _RUNTIME_ERROR_CODES for d in diagnostics):
        return diagnostics
    dead: list[Span] = []
    for member in active_module_members(ctx.mod, ctx.activity):
        if isinstance(member, ProcedureNode):
            # The port's set holds the unreachable nodes' ids.
            unreachable = unreachable_statements_in(ctx.source, member, ctx.symbols, ctx.activity)
            if unreachable:
                dead.extend(node.span for node in iter_body_nodes(member.body) if id(node) in unreachable)
            dead.extend(dead_branch_spans_in(ctx.source, member, ctx.symbols, ctx.activity))
    if not dead:
        return diagnostics
    return [
        d
        for d in diagnostics
        if d.code not in _RUNTIME_ERROR_CODES
        or not any(d.span.start >= span.start and d.span.end <= span.end for span in dead)
    ]


def _apply_inline_suppression(
    source: str,
    out: list[VbaDiagnostic],
    overrides: Mapping[str, str] | None,
    whole_project: bool,
    host_known: bool,
) -> list[VbaDiagnostic]:
    """Drop the diagnostics that '@pyvba-ignore directives suppress, then surface any
    malformed directive as an analysis-suppression-directive diagnostic (itself
    subject to severity overrides, never to inline suppression)."""
    scan = scan_inline_suppressions(source)
    out = filter_inline_suppressions(source, out, scan)
    if scan.issues:
        directive_meta = diagnostic_metadata_for_code(DIRECTIVE_DIAGNOSTIC_CODE)
        severity = (
            _severity_of(directive_meta.rule_name, overrides, whole_project, host_known)
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
    dotted reference, and with a With-scan cache so each procedure's `With` stack
    is scanned once rather than once per leading-dot member. `me_project_type` and
    `me_type` are derived from the module identity; `code_names` is left unset (the
    diagnostics pass has no code-name map).
    """
    ctx = MemberCompletionContext(
        project_class_members=opts.project_class_members,
        allow_set_assignment_refinement=False,
        model=opts.host_model,
        parsed_module=mod,
        source_tokens=[t for t in tokenize_cached(source) if t.kind is not TokenKind.COMMENT],
        # The With-scan index of a procedure is built once per pass, not once per
        # leading-dot member: without the cache a 2,000-line With block cost
        # 1.8 s, with it 0.3 s, and the findings are identical (XLIDE issue #134).
        with_scan_cache={},
        receiver_type_cache={},
        receiver_chain_cache={},
    )
    me_project_type = _me_project_type_for(opts.module_name, opts.module_kind)
    if me_project_type:
        ctx.me_project_type = me_project_type
    me_type = _me_host_type_for(opts.module_name, opts.module_kind, opts.host, opts.designer_class)
    if me_type:
        ctx.me_type = me_type
    return ctx


def _me_project_type_for(
    module_name: str | None, module_kind: ModuleSymbolKind | None
) -> str | None:
    return module_name if module_name and is_object_module_kind(module_kind) else None


def _me_host_type_for(
    module_name: str | None,
    module_kind: ModuleSymbolKind | None,
    host: str | None = None,
    designer_class: str | None = None,
) -> str | None:
    # A designer class is what the module IS, whatever kind it is listed as: an
    # Access form's `Me` reaches Requery from `Access.Form`, which the module's own
    # text never declares. The module's project type still applies alongside, so
    # `Me` keeps its own procedures too.
    if designer_class:
        return designer_class
    if not module_name or module_kind is not ModuleSymbolKind.DOCUMENT:
        return None
    lower = module_name.lower()
    token = (host or "excel").lower()
    if token == "excel":
        return "Excel.Workbook" if lower == "thisworkbook" else None
    if token == "word":
        return "Word.Document" if lower == "thisdocument" else None
    # PowerPoint has no document modules; other hosts' document surfaces are
    # unmodelled, and silence beats a wrong type.
    return None


def incomplete_expression_edit_span(source: str, offset: int) -> Span | None:
    """The span of the expression being typed at `offset` that is not yet complete:
    a member access with nothing after its dot, a trailing binary operator, or the
    first unmatched open parenthesis. The editor uses it to hold back the findings
    such half-typed text would raise."""
    line = physical_line_span_at_offset(source, offset)
    statement = _active_statement_span_on_line(source, line, offset)
    found = incomplete_member_access(source, statement, include_leading_dot=True)
    if found is not None:
        return found
    found = _trailing_binary_operator_edit_span(source, statement, offset)
    if found is not None:
        return found
    return _unmatched_open_paren_edit_span(source, statement, offset)


def _active_statement_span_on_line(source: str, line: Span, offset: int) -> Span:
    safe_offset = max(0, min(offset, len(source)))
    depth = 0
    start = line.start
    end = line.end
    for tok in statement_tokens(source, line):
        if tok.kind is TokenKind.PUNCTUATION:
            if tok.raw_text == "(":
                depth += 1
            elif tok.raw_text == ")":
                depth = max(0, depth - 1)
            continue
        if tok.kind is not TokenKind.COLON or depth != 0:
            continue
        colon = absolute_span(line, tok)
        if safe_offset <= colon.start:
            end = colon.start
            break
        start = colon.end
    return Span(start, end)


_BLANK_TAIL_RE = re.compile(r"[ \t]*")


def _trailing_binary_operator_edit_span(source: str, span: Span, offset: int) -> Span | None:
    toks = statement_tokens(source, span)
    last = toks[-1] if toks else None
    if last is None or not is_non_unary_binary_operator(last):
        return None
    active = absolute_span(span, last)
    if offset < active.start:
        return None
    cursor_tail = source[active.end : max(active.end, min(offset, span.end))]
    return active if _BLANK_TAIL_RE.fullmatch(cursor_tail) else None


def _unmatched_open_paren_edit_span(source: str, span: Span, offset: int) -> Span | None:
    stack: list[VbaToken] = []
    for tok in statement_tokens(source, span):
        if tok.kind is not TokenKind.PUNCTUATION:
            continue
        if tok.raw_text == "(":
            stack.append(tok)
        elif tok.raw_text == ")" and stack:
            stack.pop()
    if not stack:
        return None
    active = absolute_span(span, stack[0])
    return active if offset >= active.start else None
