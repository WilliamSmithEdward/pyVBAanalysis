"""What XLIDE does to a module's findings before it shows them.

Ported from xlide_vscode/src/vbaModuleAnalysis.ts. XLIDE's live diagnostics, its
analysis command and its project-wide analysis all run the rules through
analyzeVbaModuleSource, and two of its steps change the list the rules return:

- A deterministic runtime-error finding in a stretch of a procedure under
  `On Error Resume Next` is dropped (XLIDE issue #106). The error is raised and
  handled there, which is usually the point: `n = UBound(a)` under Resume Next is
  the common test for an allocated array. A stretch runs from the statement to the
  next `On Error Resume Next`, `On Error GoTo label` or `On Error GoTo 0` statement
  or the end of the procedure, in source order; branches are not modelled, so a
  Resume Next inside an If arm covers what follows the arm. An On Error statement
  that never runs sets nothing, and one inside a running error handler covers
  nothing. A handler no error can reach is dropped from in full.
- Findings with the same code and span are merged, the later one taking the
  earlier one's place.

analyze_module applies both unless AnalyzeModuleOptions.raw_rule_output is set.
The wrapper's other steps stay out of the port: the structural block-balance pass,
the @xlide-test directive check and the expected-error suppression that belong to
XLIDE's test runner, XLIDE's own @xlide-analysis suppression comments, where the
port has '@pyvba-ignore, and the analysisFailures record, where the port reports
through AnalyzeModuleOptions.on_internal_error.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from collections.abc import Set as AbstractSet

from ..conditional import ConditionalActivityTracker
from ..js_compat import JS_WHITESPACE
from ..lexer.token_kinds import TokenKind
from ..parser.nodes import BodyNode, ModuleNode, ProcedureNode, Span, StatementNode
from ..symbols.symbol_model import ModuleSymbols
from ..types.type_inference import unreachable_statements_in
from .model import DiagnosticEvidenceKind, VbaDiagnostic
from .rule_metadata import diagnostic_metadata_for_code
from .rules.handler_flow import error_handler_extents, on_error_mode
from .walker import statement_tokens_after_leading_label

_LEAVING_WORDS = frozenset({"goto", "exit", "end", "resume", "return"})

# /^\s*([A-Za-z_][A-Za-z0-9_]*|\d+)\s*:?/ with JavaScript's `\s` and `\d`.
_S = "[" + JS_WHITESPACE + "]"
_HANDLER_LABEL_RE = re.compile("^" + _S + "*([A-Za-z_][A-Za-z0-9_]*|[0-9]+)" + _S + "*:?")


def _block_body(node: BodyNode) -> list[BodyNode] | None:
    body = getattr(node, "body", None)
    return body if isinstance(body, list) else None


def on_error_resume_next_spans(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> list[Span]:
    """The stretches of each procedure under an active `On Error Resume Next`, and
    the error handlers no error can reach."""
    out: list[Span] = []
    for member in mod.members:
        if not isinstance(member, ProcedureNode):
            continue
        handlers: list[tuple[int, int, bool]] = []
        running = error_handler_extents(source, member)
        # An On Error statement a known guard keeps from running sets nothing:
        # `If False Then ... On Error Resume Next ... End If`, a For of no pass, a
        # Case that cannot match, the line after a GoTo (XLIDE issue #486). The walk
        # is run only where one could be dead.
        # The port's unreachable set holds node ids.
        dead: AbstractSet[int] | None = None

        def never_runs(node: BodyNode, nested: bool, before: Sequence[BodyNode]) -> bool:
            nonlocal dead
            # `If True Then GoTo L` before it can leave it dead too (XLIDE issue #673).
            may_be_dead = nested or any(
                isinstance(earlier, StatementNode)
                and any(
                    tok.raw_text.lower() in _LEAVING_WORDS
                    for tok in statement_tokens_after_leading_label(source, earlier.span)
                )
                for earlier in before
            )
            if not may_be_dead:
                return False
            if dead is None:
                dead = unreachable_statements_in(source, member, symbols, activity)
            return id(node) in dead

        # Bodies are walked on an explicit stack; the order handlers are found in
        # does not matter, since they are sorted below.
        stack: list[tuple[Sequence[BodyNode], bool]] = [(member.body, False)]
        while stack:
            body, nested = stack.pop()
            for index, node in enumerate(body):
                if isinstance(node, StatementNode):
                    # Read after any line label: `10 On Error Resume Next`.
                    mode = on_error_mode(statement_tokens_after_leading_label(source, node.span))
                    # GoTo 0 turns handling off and ends the stretch; GoTo -1 only
                    # clears the error, and Resume Next stays (XLIDE issue #313).
                    if mode in ("resume-next", "goto-label", "goto-0") and not never_runs(
                        node, nested, body[:index]
                    ):
                        start = node.span.start
                        handlers.append(
                            (
                                start,
                                node.span.end,
                                mode == "resume-next"
                                and not any(extent.start <= start < extent.end for extent in running),
                            )
                        )
                else:
                    inner = _block_body(node)
                    if inner is not None:
                        stack.append((inner, True))
        out.extend(_unentered_handlers(source, member, running))
        handlers.sort(key=lambda handler: handler[0])
        for i, (_, end, resume_next) in enumerate(handlers):
            if resume_next:
                until = handlers[i + 1][0] if i + 1 < len(handlers) else member.span.end
                out.append(Span(end, until))
    return out


def _unentered_handlers(source: str, member: ProcedureNode, running: Sequence[Span]) -> list[Span]:
    """The handlers no error can reach: every `On Error GoTo H` naming the handler's
    label is followed at once by Exit or End, so nothing raises while it is set and
    what the handler would do never runs (XLIDE issue #556)."""
    safe_after: dict[str, bool] = {}
    stack: list[Sequence[BodyNode]] = [member.body]
    while stack:
        body = stack.pop()
        for index, node in enumerate(body):
            inner = _block_body(node)
            if inner is not None:
                stack.append(inner)
                continue
            if not isinstance(node, StatementNode):
                continue
            toks = [
                tok
                for tok in statement_tokens_after_leading_label(source, node.span)
                if tok.kind is not TokenKind.COMMENT
            ]
            if on_error_mode(toks) != "goto-label":
                continue
            label = toks[3].raw_text.lower() if len(toks) > 3 else None
            following = body[index + 1] if index + 1 < len(body) else None
            head = (
                [
                    tok.raw_text.lower()
                    for tok in statement_tokens_after_leading_label(source, following.span)
                    if tok.kind is not TokenKind.COMMENT
                ]
                if isinstance(following, StatementNode)
                else []
            )
            first = head[0] if head else None
            leaves = len(toks) == 4 and (first == "exit" or (first == "end" and len(head) == 1))
            if label:
                safe_after[label] = safe_after.get(label, True) and leaves
    out: list[Span] = []
    for extent in running:
        match = _HANDLER_LABEL_RE.match(source[extent.start : extent.end])
        label = match.group(1).lower() if match is not None else None
        if label and safe_after.get(label) is True:
            out.append(extent)
    return out


def _is_deterministic_runtime_error(code: str) -> bool:
    meta = diagnostic_metadata_for_code(code)
    return (
        meta is not None
        and meta.diagnostic_kind is DiagnosticEvidenceKind.DETERMINISTIC_RUNTIME_ERROR
    )


def drop_handled_runtime_errors(
    diagnostics: list[VbaDiagnostic],
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> list[VbaDiagnostic]:
    """The findings without the deterministic runtime errors that start inside a
    stretch under `On Error Resume Next` or a handler no error reaches."""
    runtime = [_is_deterministic_runtime_error(d.code) for d in diagnostics]
    if not any(runtime):
        return diagnostics
    spans = on_error_resume_next_spans(source, mod, symbols, activity)
    if not spans:
        return diagnostics
    return [
        d
        for d, is_runtime in zip(diagnostics, runtime, strict=True)
        if not (is_runtime and _starts_inside_any(d.span, spans))
    ]


def _starts_inside_any(span: Span, stretches: Sequence[Span]) -> bool:
    return any(stretch.start <= span.start < stretch.end for stretch in stretches)


def deduplicate_diagnostics(diagnostics: list[VbaDiagnostic]) -> list[VbaDiagnostic]:
    """One finding per code and span: a later one replaces the earlier in place."""
    out: list[VbaDiagnostic] = []
    index_by_key: dict[tuple[str, int, int], int] = {}
    for d in diagnostics:
        key = (d.code, d.span.start, d.span.end)
        at = index_by_key.get(key)
        if at is None:
            index_by_key[key] = len(out)
            out.append(d)
        else:
            out[at] = d
    return out
