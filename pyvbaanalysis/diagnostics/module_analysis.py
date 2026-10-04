"""What XLIDE does to a module's findings before it shows them.

Ported from xlide_vscode/src/vbaModuleAnalysis.ts. XLIDE's live diagnostics, its
analysis command and its project-wide analysis all run the rules through
analyzeVbaModuleSource, and two of its steps change the list the rules return:

- A deterministic runtime-error finding in a stretch of a procedure under
  `On Error Resume Next` is dropped (XLIDE issue #106). The error is raised and
  handled there, which is usually the point: `n = UBound(a)` under Resume Next is
  the common test for an allocated array. A stretch runs from the statement to the
  next `On Error` statement or the end of the procedure, in source order; branches
  are not modelled, so a Resume Next inside an If arm covers what follows the arm.
- Findings with the same code and span are merged, the later one taking the
  earlier one's place.

analyze_module applies both unless AnalyzeModuleOptions.raw_rule_output is set.
The wrapper's other steps stay out of the port: the structural block-balance pass,
the @xlide-test directive check and the expected-error suppression that belong to
XLIDE's test runner, and XLIDE's own @xlide-analysis suppression comments, where
the port has '@pyvba-ignore.
"""

from __future__ import annotations

import re
from collections.abc import Sequence

from ..js_compat import JS_WHITESPACE
from ..parser.nodes import ModuleNode, ProcedureNode, Span, StatementNode, iter_body_nodes
from .model import DiagnosticEvidenceKind, VbaDiagnostic
from .rule_metadata import diagnostic_metadata_for_code

# /^\s*On\s+(?:Local\s+)?Error\s+(Resume\s+Next|GoTo\b)/i, with JavaScript's `\s`
# and its ASCII-only `\b` and case folding.
_S = "[" + JS_WHITESPACE + "]"
_ON_ERROR_RE = re.compile(
    "^" + _S + "*On" + _S + "+(?:Local" + _S + "+)?Error" + _S + "+(Resume" + _S + r"+Next|GoTo\b)",
    re.IGNORECASE | re.ASCII,
)


def on_error_resume_next_spans(mod: ModuleNode) -> list[Span]:
    """The stretches of each procedure under an active `On Error Resume Next`."""
    out: list[Span] = []
    for member in mod.members:
        if not isinstance(member, ProcedureNode):
            continue
        handlers: list[tuple[int, int, bool]] = []
        for node in iter_body_nodes(member.body):
            if not isinstance(node, StatementNode):
                continue
            match = _ON_ERROR_RE.match(node.raw)
            if match is not None:
                resume_next = match.group(1).lower().startswith("resume")
                handlers.append((node.span.start, node.span.end, resume_next))
        handlers.sort(key=lambda handler: handler[0])
        for i, (_, end, resume_next) in enumerate(handlers):
            if resume_next:
                until = handlers[i + 1][0] if i + 1 < len(handlers) else member.span.end
                out.append(Span(end, until))
    return out


def _is_deterministic_runtime_error(code: str) -> bool:
    meta = diagnostic_metadata_for_code(code)
    return (
        meta is not None
        and meta.diagnostic_kind is DiagnosticEvidenceKind.DETERMINISTIC_RUNTIME_ERROR
    )


def drop_handled_runtime_errors(
    diagnostics: list[VbaDiagnostic], mod: ModuleNode
) -> list[VbaDiagnostic]:
    """The findings without the deterministic runtime errors that start inside a
    stretch under `On Error Resume Next`."""
    runtime = [_is_deterministic_runtime_error(d.code) for d in diagnostics]
    if not any(runtime):
        return diagnostics
    spans = on_error_resume_next_spans(mod)
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


# --- sync stubs (2f49b93): replaced as each group is ported ---


class VbaModuleAnalysisDiagnostic:
    pass


class VbaModuleAnalysisInput:
    pass


class VbaModuleAnalysisResult:
    pass


class VbaModuleAnalysisFailure:
    pass


def analyze_vba_module_source(*args: object, **kwargs: object) -> object:
    raise NotImplementedError("analyzeVbaModuleSource not ported yet")
