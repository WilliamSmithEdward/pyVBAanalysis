"""Rule family: runtime errors that control flow itself raises (XLIDE issue #117).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/handlerFlow.ts.

Each construct below compiles and raises every time it runs, measured in Excel 16.0
(build 20326, 2026-09-26):

- handler-fall-through: a procedure whose normal path runs off the end into its
  error handler, where `Err.Raise Err.Number` re-raises with no error pending.
  Err.Number is 0 there, and `Err.Raise 0` is error 5. A handler that ends with
  Resume Next does not raise; it just runs once too often, so only the re-raising
  form is reported.
- resume-without-error: a `Resume` statement in a procedure that never installs a
  handler with `On Error GoTo <label>`. Nothing can be pending when it runs, and
  Resume then raises error 20.
- return-without-gosub: a GoSub target entered by falling into it from the
  statement above, whose `Return` then has no GoSub to return to, error 3.
- recursive-property-accessor: a Property Get that reads `Me.Name` for its own
  Name, or a Property Let/Set that assigns `Name = value` to its own Name (in a
  Let, the name is the property, not a return variable). Each calls itself without
  end, error 28.

Fall-through is proven only from a plain statement directly above the label: a
block above it may or may not leave the procedure, and nothing is reported for it.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Sequence
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker, inactive_node_skip
from ...flow.procedure_labels import (
    collect_procedure_label_references,
    statement_label_declaration,
)
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    ProcKind,
    Span,
    is_leaf_statement,
    iter_body_nodes,
)
from ..context import PushFn
from ..walker import (
    active_module_members,
    bare_assignment_target,
    first_executable_token_index,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

# The heads of the plain statements that always leave the place they stand.
_LEAVING_HEADS = frozenset({"exit", "goto", "return", "resume", "end", "stop", "error"})

# /^0+$/ and /^0*1$/, used with fullmatch.
_ZEROS_RE = re.compile(r"0+")
_ONE_RE = re.compile(r"0*1")


@dataclass(frozen=True, slots=True)
class _TopLevelEntry:
    """One top-level entry of a procedure body: a leaf statement or an opaque block."""

    node: BodyNode
    leaf: LeafStatementNode | None
    # The label this statement declares, lower-cased, when it does.
    label: str | None = None
    label_span: Span | None = None


def check_handler_flow(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """Report the runtime errors the control flow of a procedure raises."""
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        entries = _top_level_entries(source, member.body, activity)
        _check_resume_without_error(source, member, activity, push)
        _check_fall_through_into_targets(source, member, entries, activity, push)
        _check_recursive_property(source, member, entries, push)


def _top_level_entries(
    source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> list[_TopLevelEntry]:
    out: list[_TopLevelEntry] = []
    for node in body:
        if activity is not None and activity.is_inactive(node.span):
            continue
        if not is_leaf_statement(node):
            out.append(_TopLevelEntry(node, None))
            continue
        label = statement_label_declaration(source, node.span)
        out.append(
            _TopLevelEntry(
                node,
                node,
                label.key if label is not None else None,
                label.span if label is not None else None,
            )
        )
    return out


def _leaves_unconditionally(source: str, stmt: LeafStatementNode) -> bool:
    """Whether a plain statement always leaves the place it stands: Exit, GoTo, Return,
    Resume, End, Err.Raise, Error, Stop."""
    toks = statement_tokens_after_leading_label(source, stmt.span)
    head = token_text(_at(toks, 0))
    if head in _LEAVING_HEADS:
        # `End If` and the like never reach here (they close blocks), so `End`
        # alone is the End statement.
        return True
    return head == "err" and _raw_at(toks, 1) == "." and token_text(_at(toks, 2)) == "raise"


def _label_body(entries: Sequence[_TopLevelEntry], index: int) -> list[_TopLevelEntry]:
    """The statements of a label's body: from the label to the next label or the end."""
    out: list[_TopLevelEntry] = []
    for k in range(index, len(entries)):
        if k > index and entries[k].label is not None:
            break
        out.append(entries[k])
    return out


def _check_fall_through_into_targets(
    source: str,
    proc: ProcedureNode,
    entries: Sequence[_TopLevelEntry],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    references = collect_procedure_label_references(source, proc, activity)
    handler_labels = {ref.key for ref in references if ref.statement_kind == "on-error-goto"}
    gosub_labels = {
        ref.key for ref in references if ref.statement_kind == "gosub" or ref.statement_kind == "on-gosub"
    }
    for i, entry in enumerate(entries):
        if entry.label is None or entry.label_span is None:
            continue
        # The flow above must be a plain statement that does not leave; a
        # block above proves nothing. The first statement of the body is
        # entered directly.
        above = entries[i - 1] if i > 0 else None
        falls_in = above is None or (
            above.leaf is not None and not _leaves_unconditionally(source, above.leaf)
        )
        if not falls_in:
            continue
        body = _label_body(entries, i)
        if entry.label in handler_labels:
            reraise = _first_index(
                body, lambda one: one.leaf is not None and _reraises_pending_error(source, one.leaf)
            )
            exits_first = _first_index(
                body,
                lambda one: one.leaf is not None
                and _leaves_unconditionally(source, one.leaf)
                and not _reraises_pending_error(source, one.leaf),
            )
            if reraise >= 0 and (exits_first < 0 or reraise < exits_first):
                push(
                    "handlerFallThrough",
                    f"Execution falls into error handler '{_label_text(source, entry)}' with no error "
                    "pending, and 'Err.Raise Err.Number' then raises with Err.Number 0. This will "
                    "raise Run-time error '5': Invalid procedure call or argument. Put an Exit "
                    f"{_procedure_word(proc)} before the label.",
                    entry.label_span,
                )
        if entry.label in gosub_labels:
            returns = _first_index(
                body,
                lambda one: one.leaf is not None
                and token_text(_at(statement_tokens_after_leading_label(source, one.leaf.span), 0))
                == "return",
            )
            exits_first = _first_index(
                body, lambda one: one.leaf is not None and _leaves_unconditionally(source, one.leaf)
            )
            if returns >= 0 and returns == exits_first:
                push(
                    "returnWithoutGosub",
                    f"Execution falls into GoSub target '{_label_text(source, entry)}' from the "
                    "statement above it, and its Return then has no GoSub to return to. This will "
                    "raise Run-time error '3': Return without GoSub. Put an Exit "
                    f"{_procedure_word(proc)} before the label.",
                    entry.label_span,
                )


def _reraises_pending_error(source: str, stmt: LeafStatementNode) -> bool:
    """`Err.Raise Err.Number` (with or without further arguments)."""
    toks = statement_tokens_after_leading_label(source, stmt.span)
    return (
        token_text(_at(toks, 0)) == "err"
        and _raw_at(toks, 1) == "."
        and token_text(_at(toks, 2)) == "raise"
        and token_text(_at(toks, 3)) == "err"
        and _raw_at(toks, 4) == "."
        and token_text(_at(toks, 5)) == "number"
    )


def _label_text(source: str, entry: _TopLevelEntry) -> str:
    if entry.label_span is not None:
        return source[entry.label_span.start : entry.label_span.end]
    return entry.label if entry.label is not None else ""


def _procedure_word(proc: ProcedureNode) -> str:
    if proc.proc_kind is ProcKind.SUB:
        return "Sub"
    if proc.proc_kind is ProcKind.FUNCTION:
        return "Function"
    return "Property"


def _check_resume_without_error(
    source: str,
    proc: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    resumes: list[Span] = []
    installs_handler = False
    # Upstream recurses into every nested block body; the explicit-stack walk
    # visits the same leaf statements in the same order.
    for node in iter_body_nodes(proc.body, inactive_node_skip(activity)):
        if not is_leaf_statement(node):
            continue
        for span in statement_and_branch_spans(node):
            toks = statement_tokens_after_leading_label(source, span)
            head = token_text(_at(toks, 0))
            if (
                head == "on"
                and any(token_text(tok) == "error" for tok in toks)
                and any(token_text(tok) == "goto" for tok in toks)
            ):
                # `On Error GoTo 0` and `On Error GoTo -1` install nothing; a label
                # does. The lexer gives `-1` as two tokens (XLIDE issue #142).
                target = toks[-1]
                zero = (
                    target.kind is TokenKind.INTEGER_LITERAL
                    and _ZEROS_RE.fullmatch(target.raw_text) is not None
                )
                minus_one = (
                    target.kind is TokenKind.INTEGER_LITERAL
                    and _ONE_RE.fullmatch(target.raw_text) is not None
                    and _raw_at(toks, len(toks) - 2) == "-"
                )
                if not zero and not minus_one:
                    installs_handler = True
            if head == "resume":
                resumes.append(Span(span.start + toks[0].start, span.start + toks[0].end))
    if installs_handler:
        return
    for span in resumes:
        push(
            "resumeWithoutError",
            "'Resume' runs with no error handler installed in this procedure, so no error is "
            "pending. This will raise Run-time error '20': Resume without error.",
            span,
        )


def _check_recursive_property(
    source: str,
    proc: ProcedureNode,
    entries: Sequence[_TopLevelEntry],
    push: PushFn,
) -> None:
    if proc.proc_kind not in (ProcKind.PROPERTY_GET, ProcKind.PROPERTY_LET, ProcKind.PROPERTY_SET):
        return
    lower = proc.name.lower()
    for entry in entries:
        if entry.leaf is None:
            continue
        toks = statement_tokens_after_leading_label(source, entry.leaf.span)
        if proc.proc_kind is ProcKind.PROPERTY_GET:
            # `Name = Me.Name`: the bare Name is the return variable, `Me.Name`
            # the property, which is this procedure.
            for i in range(len(toks) - 2):
                member_name = token_name(toks[i + 2])
                if (
                    token_text(toks[i]) == "me"
                    and toks[i + 1].raw_text == "."
                    and member_name is not None
                    and member_name.lower() == lower
                    and len(proc.params) == 0
                    and _raw_at(toks, i + 3) != "("
                ):
                    push(
                        "recursivePropertyAccessor",
                        f"Property Get '{proc.name}' reads 'Me.{proc.name}', which is itself: the "
                        "call never returns. This will raise Run-time error '28': Out of stack space.",
                        _absolute_range(entry.leaf.span, toks[i], toks[i + 2]),
                    )
            continue
        # Property Let/Set: `Name = value` assigns the property, which is this
        # procedure, and Property Let has no return variable to mean instead.
        target = bare_assignment_target(source, entry.leaf.span)
        first = first_executable_token_index(toks)
        if target is not None and target[0].lower() == lower and len(proc.params) == 1:
            accessor = "Property Let" if proc.proc_kind is ProcKind.PROPERTY_LET else "Property Set"
            push(
                "recursivePropertyAccessor",
                f"{accessor} '{proc.name}' assigns '{proc.name}', which is itself: the call never "
                "returns. This will raise Run-time error '28': Out of stack space.",
                Span(entry.leaf.span.start + toks[first].start, entry.leaf.span.start + toks[first].end),
            )


def _absolute_range(base: Span, first: VbaToken, last: VbaToken) -> Span:
    return Span(base.start + first.start, base.start + last.end)


def _first_index(
    entries: Sequence[_TopLevelEntry], predicate: Callable[[_TopLevelEntry], bool]
) -> int:
    """Array.prototype.findIndex: the first entry the predicate accepts, or -1."""
    for k, entry in enumerate(entries):
        if predicate(entry):
            return k
    return -1


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) outside the list."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    """`toks[i]?.rawText`."""
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


# --- sync stubs (2f49b93): replaced as each group is ported ---


OnErrorMode = object


def on_error_mode(*args: object, **kwargs: object) -> None:
    return None


def error_handler_extents(*args: object, **kwargs: object) -> None:
    return None
