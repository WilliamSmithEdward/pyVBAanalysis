"""Detection of control flow the structural branch-merge cannot model soundly.

Ported from xlide_vscode/src/analyzer/flow/procedureUnstructured.ts. Dataflow
rules fall back to the conservative straight-line walk for procedures whose flow
can skip or re-run assignments in ways the If/ElseIf/Else merge cannot see.
"""

from __future__ import annotations

from collections.abc import Sequence

from ..conditional import ConditionalActivityTracker
from ..lexer.token_helpers import statement_tokens, token_word, tokens_without_leading_line_number
from ..parser.nodes import BodyNode, ProcedureNode, Span, is_leaf_statement
from .procedure_labels import (
    collect_procedure_label_declarations,
    collect_procedure_label_references,
)


def procedure_has_unstructured_flow(
    source: str,
    procedure: ProcedureNode,
    activity: ConditionalActivityTracker | None = None,
) -> bool:
    """True when a procedure contains label / GoTo / On Error / Resume flow.

    Any label, any GoTo / GoSub / On..GoTo / On..GoSub / Resume target, or any
    `On Error` / `Resume` statement (whose exception edges can bypass an
    assignment the merge would assume ran) forces the conservative straight-line
    dataflow, preserving the no-false-positive contract.
    """
    if collect_procedure_label_references(source, procedure, activity):
        return True
    if collect_procedure_label_declarations(source, procedure, activity):
        return True
    # On Error Resume Next / On Error GoTo 0 / bare Resume carry no label, so the
    # collectors above miss them; scan for them directly.
    return _has_on_error_or_resume_statement(procedure.body, source, activity)


def _has_on_error_or_resume_statement(
    body: Sequence[BodyNode], source: str, activity: ConditionalActivityTracker | None
) -> bool:
    for node in body:
        if activity is not None and activity.is_inactive(node.span):
            continue
        if is_leaf_statement(node):
            if _is_on_error_or_resume(source, node.span):
                return True
        else:
            child = getattr(node, "body", None)
            if isinstance(child, list) and _has_on_error_or_resume_statement(child, source, activity):
                return True
    return False


def _is_on_error_or_resume(source: str, span: Span) -> bool:
    toks = tokens_without_leading_line_number(statement_tokens(source, span.start, span.end))
    if not toks:
        return False
    first = token_word(toks[0])
    if first == "resume":
        return True
    return first == "on" and len(toks) > 1 and token_word(toks[1]) == "error"
