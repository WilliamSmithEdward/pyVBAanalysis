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
- unbounded-recursion: a Sub or Function that calls itself, or another procedure
  of the module that calls it back, before anything that could leave (XLIDE issue
  #240): `Sub S(): S: End Sub`, `F = F() + 1`, and `F = F(n - 1)` with no base
  case. Error 28.

Fall-through is proven only from a plain statement directly above the label: a
block above it may or may not leave the procedure, and nothing is reported for it.

Upstream recurses over nested block bodies (bodyLists, resetsHandler, the Resume
walk); the port walks them on explicit stacks in the same order.
"""

from __future__ import annotations

import dataclasses
import re
from collections.abc import Callable, Iterable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass
from typing import Literal

from ...conditional import ConditionalActivityTracker, inactive_node_skip
from ...flow.procedure_labels import (
    VbaProcedureLabel,
    VbaProcedureLabelReference,
    collect_procedure_label_references,
    statement_label_declarations,
)
from ...js_compat import js_trim
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    IfBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    ProcKind,
    Span,
    StatementNode,
    VariableGroupNode,
    WithBlockNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ..context import PushFn
from ..straight_line_values import straight_line_unreachable
from ..walker import (
    active_module_members,
    bare_assignment_target,
    block_header_line_span,
    first_executable_token_index,
    statement_and_branch_spans,
    statement_tokens,
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
    # The labels this statement declares: none, one, or a line number and a name
    # (`10 L1:`).
    labels: tuple[VbaProcedureLabel, ...] = ()


def check_handler_flow(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    class_name: str | None = None,
) -> None:
    """Report the runtime errors the control flow of a procedure raises."""
    # `New Class1` inside Class1 makes another of this class (XLIDE issue #613).
    own_class = class_name.lower() if class_name is not None else None
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        entries = _top_level_entries(source, member.body, activity)
        _check_resume_without_error(source, member, activity, push)
        # A label in a block falls in from the statement above it there, the way
        # one at the top level does (XLIDE issue #237). The procedure's label
        # references are read once, not once per block (XLIDE issue #322).
        references = collect_procedure_label_references(source, member, activity)
        for body in _body_lists(member.body):
            _check_fall_through_into_targets(
                source,
                member,
                entries if body is member.body else _top_level_entries(source, body, activity),
                references,
                push,
            )
        _check_recursive_property(source, member, entries, own_class, push)
    _check_unbounded_recursion(source, mod, activity, own_class, push)


def _note_self_alias(
    toks: Sequence[VbaToken], selves: set[str], own_class: str | None
) -> None:
    """The names a statement at the top of a body makes another way to this object,
    or to a new one of its class that does the same: `Set o = Me`, `Set o = New
    Class1` (XLIDE issue #613). Me is always one."""
    words = [tok for tok in toks if tok.kind is not TokenKind.COMMENT]
    name_token = token_name(_at(words, 1))
    name = (
        name_token.lower()
        if token_text(_at(words, 0)) == "set" and _raw_at(words, 2) == "=" and name_token is not None
        else None
    )
    if not name:
        return
    value = words[3:]
    self_ = (len(value) == 1 and token_text(value[0]) in selves) or (
        own_class is not None
        and len(value) == 2
        and token_text(value[0]) == "new"
        and token_text(value[1]) == own_class
    )
    if self_:
        selves.add(name)
    else:
        selves.discard(name)


def _body_lists(body: Sequence[BodyNode]) -> list[Sequence[BodyNode]]:
    """A procedure's body, and every body a block in it holds: each If arm alone.

    Upstream recurses; this is the same pre-order walk on an explicit stack."""
    out: list[Sequence[BodyNode]] = []
    stack: list[Sequence[BodyNode]] = [body]
    while stack:
        current = stack.pop()
        out.append(current)
        children: list[Sequence[BodyNode]] = []
        for node in current:
            if isinstance(node, IfBlockNode):
                children.extend(branch.body for branch in node.branches)
                continue
            child = getattr(node, "body", None)
            if isinstance(child, list):
                children.append(child)
        stack.extend(reversed(children))
    return out


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
        out.append(_TopLevelEntry(node, node, tuple(statement_label_declarations(source, node.span))))
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
        if k > index and len(entries[k].labels) > 0:
            break
        out.append(entries[k])
    return out


def _check_fall_through_into_targets(
    source: str,
    proc: ProcedureNode,
    entries: Sequence[_TopLevelEntry],
    references: Sequence[VbaProcedureLabelReference],
    push: PushFn,
) -> None:
    handler_labels = {ref.key for ref in references if ref.statement_kind == "on-error-goto"}
    gosub_labels = {
        ref.key for ref in references if ref.statement_kind == "gosub" or ref.statement_kind == "on-gosub"
    }
    named = {ref.key for ref in references}
    for i, entry in enumerate(entries):
        if len(entry.labels) == 0:
            continue
        # The flow above must be a plain statement that does not leave, and that
        # itself runs; a block above proves nothing. The first statement of the
        # body is entered directly.
        above = entries[i - 1] if i > 0 else None
        falls_in = above is None or (
            above.leaf is not None
            and not _leaves_unconditionally(source, above.leaf)
            and _runs(source, entries, i - 1, named)
        )
        if not falls_in:
            continue
        body = _label_body(entries, i)
        handler = next((label for label in entry.labels if label.key in handler_labels), None)
        if handler is not None:
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
                    f"Execution falls into error handler '{handler.text}' with no error "
                    "pending, and 'Err.Raise Err.Number' then raises with Err.Number 0. This will "
                    "raise Run-time error '5': Invalid procedure call or argument. Put an Exit "
                    f"{_procedure_word(proc)} before the label.",
                    handler.span,
                )
        target = next((label for label in entry.labels if label.key in gosub_labels), None)
        if target is not None:
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
                    f"Execution falls into GoSub target '{target.text}' from the "
                    "statement above it, and its Return then has no GoSub to return to. This will "
                    "raise Run-time error '3': Return without GoSub. Put an Exit "
                    f"{_procedure_word(proc)} before the label.",
                    target.span,
                )


def _runs(
    source: str, entries: Sequence[_TopLevelEntry], index: int, named: AbstractSet[str]
) -> bool:
    """Whether the top-level entry at `index` can run: false when a statement above
    it leaves unconditionally with nothing between that a statement can jump to. A
    label no statement names is no way in: in `Exit Function`, `Skip:`, `x = 1`, the
    assignment is dead, and falls into nothing (XLIDE issue #203, measured in Excel
    16.0)."""
    for k in range(index, -1, -1):
        entry = entries[k]
        if any(label.key in named for label in entry.labels):
            return True
        if k < index and entry.leaf is not None and _leaves_unconditionally(source, entry.leaf):
            return False
        if entry.leaf is None:
            return True  # a block may or may not leave
    return True


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
    # A handler below Exit Function that nothing jumps to never runs, as when its
    # On Error line is commented out (XLIDE issue #421, measured in Excel 16.0).
    dead: frozenset[int] | None = None
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
                if dead is None:
                    dead = _node_ids(straight_line_unreachable(source, proc.body, activity))
                if id(node) not in dead:
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


def _node_ids(dead: Iterable[object]) -> frozenset[int]:
    """straightLineUnreachable's node set, keyed by node identity. Body nodes are
    mutable dataclasses and so unhashable; the set holds either the nodes or their
    ids, and both read the same here."""
    return frozenset(one if isinstance(one, int) else id(one) for one in dead)


# How an `On Error` statement sets error handling.
OnErrorMode = Literal["resume-next", "goto-label", "goto-0", "goto-minus-1"]


def on_error_mode(toks: Sequence[VbaToken]) -> OnErrorMode | None:
    """The mode an `On [Local] Error` statement sets, from its tokens after any line
    label."""
    words = [tok for tok in toks if tok.kind is not TokenKind.COMMENT]
    i = 1 if token_text(_at(words, 0)) == "on" else -1
    if i < 0:
        return None
    if token_text(_at(words, i)) == "local":
        i += 1
    if token_text(_at(words, i)) != "error":
        return None
    if token_text(_at(words, i + 1)) == "resume" and token_text(_at(words, i + 2)) == "next":
        return "resume-next"
    if token_text(_at(words, i + 1)) != "goto":
        return None
    # The lexer gives `-1` as two tokens (XLIDE issue #142).
    target = words[i + 2 :]
    if (
        len(target) == 1
        and target[0].kind is TokenKind.INTEGER_LITERAL
        and _ZEROS_RE.fullmatch(target[0].raw_text) is not None
    ):
        return "goto-0"
    if (
        len(target) == 2
        and target[0].raw_text == "-"
        and target[1].kind is TokenKind.INTEGER_LITERAL
        and _ONE_RE.fullmatch(target[1].raw_text) is not None
    ):
        return "goto-minus-1"
    return "goto-label"


def error_handler_extents(source: str, proc: ProcedureNode) -> list[Span]:
    """The stretches of a procedure where one of its error handlers is running.
    There, `On Error Resume Next` and `On Error GoTo label` do not take effect, and
    the next error goes to the caller (XLIDE issue #199, measured in Excel 16.0).

    A stretch starts at a label that only `On Error GoTo` names, below a statement
    that leaves, so that nothing but an error reaches it: a label execution can fall
    into, or that a GoTo names, runs with no handler active. It ends at an `On Error
    GoTo -1` anywhere, which ends the handler, or at the next label a statement
    names, which may be entered from outside; a label nothing names, like the number
    on every line of numbered code, does not end it. Nor does an Exit inside an If:
    the code after the If runs only when the If did not leave, and the handler is
    still running there. Code after a Resume or Exit at the top level is reached
    only through a named label, so those need no rule of their own.
    """
    # A handler inside a block runs to the procedure's end the same way (XLIDE
    # issue #237).
    kinds_by_label: dict[str, set[str]] = {}
    for ref in collect_procedure_label_references(source, proc, None):
        kinds_by_label.setdefault(ref.key, set()).add(ref.statement_kind)
    out: list[Span] = []
    for body in _body_lists(proc.body):
        out.extend(_handler_extents_in(source, proc, body, kinds_by_label))
    return out


def _handler_extents_in(
    source: str,
    proc: ProcedureNode,
    body: Sequence[BodyNode],
    kinds_by_label: Mapping[str, AbstractSet[str]],
) -> list[Span]:
    entries = _top_level_entries(source, body, None)
    out: list[Span] = []
    for i, entry in enumerate(entries):
        # What names any of the line's labels: `10 H:` is entered by GoTo 10 as well.
        kinds: set[str] = set()
        for label in entry.labels:
            kinds.update(kinds_by_label.get(label.key, ()))
        if len(kinds) != 1 or "on-error-goto" not in kinds:
            continue
        above = entries[i - 1] if i > 0 else None
        if above is None or above.leaf is None or not _leaves_unconditionally(source, above.leaf):
            continue
        end = proc.span.end
        for k in range(i, len(entries)):
            one = entries[k]
            entered = k > i and any(label.key in kinds_by_label for label in one.labels)
            if entered or _resets_handler(source, one.node):
                end = one.node.span.start
                break
        out.append(Span(entry.node.span.start, end))
    return out


def _resets_handler(source: str, node: BodyNode) -> bool:
    """Whether a statement, or any statement in a block, is `On Error GoTo -1`.

    Upstream recurses into block bodies; iter_body_nodes walks the same nodes."""
    for one in iter_body_nodes([node]):
        if not is_leaf_statement(one):
            continue
        if any(
            on_error_mode(statement_tokens_after_leading_label(source, span)) == "goto-minus-1"
            for span in statement_and_branch_spans(one)
        ):
            return True
    return False


@dataclass(frozen=True, slots=True)
class _OnceLeaf:
    leaf: LeafStatementNode
    # Whether it stands in the body of a `With Me` at the top level.
    within: bool


def _me_token(tok: VbaToken) -> VbaToken:
    """`{ ...tok, kind: 'identifier', rawText: 'Me', end: tok.start }`."""
    return dataclasses.replace(tok, kind=TokenKind.IDENTIFIER, raw_text="Me", end=tok.start)


def _check_recursive_property(
    source: str,
    proc: ProcedureNode,
    entries: Sequence[_TopLevelEntry],
    own_class: str | None,
    push: PushFn,
) -> None:
    if proc.proc_kind not in (ProcKind.PROPERTY_GET, ProcKind.PROPERTY_LET, ProcKind.PROPERTY_SET):
        return
    lower = proc.name.lower()
    # Me, and a local set to it (XLIDE issue #613).
    selves = {"me"}
    # The statements that run once each: the top level, and the body of a `With Me`
    # there, whose `.Value` is Me's (XLIDE issue #613).
    leaves: list[_OnceLeaf] = []
    for entry in entries:
        if entry.leaf is not None:
            leaves.append(_OnceLeaf(entry.leaf, False))
            continue
        if isinstance(entry.node, WithBlockNode):
            header = statement_tokens_after_leading_label(
                source, block_header_line_span(source, entry.node.span)
            )
            if len(header) == 2 and token_text(header[1]) in selves:
                for child in entry.node.body:
                    if is_leaf_statement(child) and not (
                        # A JS array is truthy even when empty.
                        isinstance(child, StatementNode)
                        and child.single_line_if_branches is not None
                    ):
                        leaves.append(_OnceLeaf(child, True))
    for once in leaves:
        leaf = once.leaf
        own = statement_tokens_after_leading_label(source, leaf.span)
        if not once.within:
            _note_self_alias(own, selves, own_class)
        # Inside `With Me`, a leading dot is Me's.
        toks: list[VbaToken]
        if once.within and _raw_at(own, 0) == ".":
            toks = [_me_token(own[0]), *own]
        elif once.within:
            toks = []
            for i, tok in enumerate(own):
                if (
                    tok.raw_text == "."
                    and i > 0
                    and own[i - 1].kind not in (TokenKind.IDENTIFIER, TokenKind.BRACKETED_IDENTIFIER)
                    and own[i - 1].raw_text != ")"
                ):
                    toks.append(_me_token(tok))
                toks.append(tok)
        else:
            toks = own
        if proc.proc_kind is ProcKind.PROPERTY_GET:
            # `Name = Me.Name`: the bare Name is the return variable, `Me.Name` the
            # property, which is this procedure. `Me.Name = 7` assigns, and calls the
            # Let (XLIDE issue #613).
            for i in range(len(toks) - 2):
                assigned = _raw_at(toks, i + 3) == "=" and (
                    i == 0
                    or (token_text(toks[i - 1]) or toks[i - 1].raw_text) in ("then", "else", "set", ":")
                )
                # `Item = Me.Item(i)`: the same arguments again (XLIDE issue #613).
                same_arguments = (
                    len(proc.params) > 0
                    and _raw_at(toks, i + 3) == "("
                    and _same_arguments_at(toks, i + 3, proc)
                )
                member_name = token_name(toks[i + 2])
                if (
                    token_text(toks[i]) in selves
                    and toks[i + 1].raw_text == "."
                    and member_name is not None
                    and member_name.lower() == lower
                    and not assigned
                    and ((len(proc.params) == 0 and _raw_at(toks, i + 3) != "(") or same_arguments)
                ):
                    push(
                        "recursivePropertyAccessor",
                        f"Property Get '{proc.name}' reads 'Me.{proc.name}', which is itself: the "
                        "call never returns. This will raise Run-time error '28': Out of stack space.",
                        _absolute_range(leaf.span, toks[i], toks[i + 2]),
                    )
                # `Value = Value() + 1`: with its parentheses the name calls the
                # property, where bare it is the return variable (XLIDE issue #338).
                own_name = token_name(toks[i])
                if (
                    own_name is not None
                    and own_name.lower() == lower
                    and _raw_at(toks, i - 1) != "."
                    and toks[i + 1].raw_text == "("
                    and toks[i + 2].raw_text == ")"
                    and len(proc.params) == 0
                ):
                    push(
                        "recursivePropertyAccessor",
                        f"Property Get '{proc.name}' calls '{proc.name}()', which is itself: the "
                        "call never returns. This will raise Run-time error '28': Out of stack space.",
                        _absolute_range(leaf.span, toks[i], toks[i + 2]),
                    )
            continue
        # `Me.Value = v` in the Let, `Set Me.Items = v` in the Set: the property
        # assigned through Me is this procedure (XLIDE issue #338).
        me_at = 1 if token_text(_at(toks, 0)) == "set" else 0
        set_form = me_at == 1
        assigned_name = token_name(_at(toks, me_at + 2))
        if (
            token_text(_at(toks, me_at)) in selves
            and _raw_at(toks, me_at + 1) == "."
            and assigned_name is not None
            and assigned_name.lower() == lower
            and _raw_at(toks, me_at + 3) == "="
            and set_form == (proc.proc_kind is ProcKind.PROPERTY_SET)
            and len(proc.params) == 1
        ):
            accessor = "Property Let" if proc.proc_kind is ProcKind.PROPERTY_LET else "Property Set"
            push(
                "recursivePropertyAccessor",
                f"{accessor} '{proc.name}' assigns 'Me.{proc.name}', which is itself: the call "
                "never returns. This will raise Run-time error '28': Out of stack space.",
                _absolute_range(leaf.span, toks[me_at], toks[me_at + 2]),
            )
            continue
        # Property Let/Set: `Name = value` assigns the property, which is this
        # procedure, and Property Let has no return variable to mean instead.
        target = bare_assignment_target(source, leaf.span)
        first = first_executable_token_index(toks)
        if target is not None and target[0].lower() == lower and len(proc.params) == 1:
            accessor = "Property Let" if proc.proc_kind is ProcKind.PROPERTY_LET else "Property Set"
            push(
                "recursivePropertyAccessor",
                f"{accessor} '{proc.name}' assigns '{proc.name}', which is itself: the call never "
                "returns. This will raise Run-time error '28': Out of stack space.",
                Span(leaf.span.start + toks[first].start, leaf.span.start + toks[first].end),
            )


# Words that may leave a procedure, or hand an error to a handler, before a call is
# reached.
_LEAVING_WORDS = frozenset(
    {"exit", "goto", "gosub", "return", "resume", "end", "stop", "error", "raise", "on"}
)

# The words after End that close a block rather than end the program.
_BLOCK_ENDS = frozenset({"if", "select", "with", "sub", "function", "property", "type", "enum"})


def _may_leave(toks: Sequence[VbaToken]) -> bool:
    """Whether any of these tokens may leave the procedure or install a handler."""
    return any(
        token_text(tok) in _LEAVING_WORDS
        and not (token_text(tok) == "end" and token_text(_at(toks, i + 1)) in _BLOCK_ENDS)
        for i, tok in enumerate(toks)
    )


@dataclass(frozen=True, slots=True)
class _FirstCall:
    """A call one procedure makes before anything could leave it."""

    callee: str
    span: Span


@dataclass(frozen=True, slots=True)
class _ProcedureCall:
    callee: str
    first: VbaToken
    last: VbaToken


def _check_unbounded_recursion(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    own_class: str | None,
    push: PushFn,
) -> None:
    """A Sub or Function whose every run calls itself, directly or through other
    procedures of the module that do the same, never returns (XLIDE issue #240,
    measured in Excel 16.0): error 28. A call counts only at the top of the body,
    ahead of any statement or block that could leave or install a handler, and
    outside a single-line If."""
    procedures: dict[str, ProcedureNode] = {}
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode) and member.proc_kind in (ProcKind.SUB, ProcKind.FUNCTION):
            procedures[member.name.lower()] = member
    first_calls: dict[str, _FirstCall] = {}
    for lower, proc in procedures.items():
        call = _first_unconditional_call(source, proc, procedures, activity, own_class)
        if call is not None:
            first_calls[lower] = call
    for lower, call in first_calls.items():
        # Follow the first calls until one repeats; report each procedure on the cycle.
        path = [lower]
        next_ = call.callee
        while next_ in first_calls and next_ not in path:
            path.append(next_)
            next_ = first_calls[next_].callee
        if next_ != lower:
            continue
        name = procedures[lower].name
        # 'A' calls 'B', which calls 'C', which calls 'A'.
        chain = [f"'{procedures[one].name}'" for one in [*path[1:], lower]]
        through = "calls itself" if len(path) == 1 else f"calls {', which calls '.join(chain)},"
        push(
            "unboundedRecursion",
            f"'{name}' {through} before anything could make it return: the calls never end. "
            "This will raise Run-time error '28': Out of stack space.",
            call.span,
        )


def _first_unconditional_call(
    source: str,
    proc: ProcedureNode,
    procedures: Mapping[str, ProcedureNode],
    activity: ConditionalActivityTracker | None,
    own_class: str | None = None,
) -> _FirstCall | None:
    selves = {"me"}
    for entry in _top_level_entries(source, proc.body, activity):
        if isinstance(entry.node, VariableGroupNode):
            continue
        if entry.leaf is None:
            # A block may leave inside.
            if _may_leave(statement_tokens(source, entry.node.span)):
                return None
            continue
        toks = statement_tokens_after_leading_label(source, entry.leaf.span)
        if _may_leave(toks):
            return None
        if token_text(_at(toks, 0)) == "if":
            continue  # a single-line If runs its call on some paths only
        _note_self_alias(toks, selves, own_class)
        call = _procedure_call_in(toks, procedures, selves)
        if call is not None:
            return _FirstCall(call.callee, _absolute_range(entry.leaf.span, call.first, call.last))
    return None


def _is_private(proc: ProcedureNode) -> bool:
    return any(modifier.lower() == "private" for modifier in proc.modifiers)


def _procedure_call_in(
    toks: Sequence[VbaToken],
    procedures: Mapping[str, ProcedureNode],
    selves: AbstractSet[str] = frozenset({"me"}),
) -> _ProcedureCall | None:
    """The first call a statement makes to a Sub or Function of the module: a call
    statement (`S`, `S 1`, `Call S(1)`), or a Function with its parentheses in an
    expression (`F()`, `F(n - 1)`). Inside F, a bare `F` is the return value.
    `F(1)`, when F takes no arguments, calls a Variant F and indexes what comes back
    (measured in Excel 16.0)."""
    # `CallByName Me, "Go", VbMethod` calls Go (XLIDE issue #613, measured in Excel
    # 16.0).
    by_name = next(
        (
            i
            for i, tok in enumerate(toks)
            if token_text(tok) == "callbyname" and _raw_at(toks, i - 1) != "."
        ),
        -1,
    )
    if by_name >= 0:
        open_ = by_name + 1 if _raw_at(toks, by_name + 1) == "(" else -1
        close = match_paren_from(toks, open_) if open_ > 0 else len(toks)
        args = split_top_level_token_groups(toks, open_ + 1 if open_ > 0 else by_name + 1, ",", close)
        second = args[1] if len(args) > 1 else None
        target = (
            second[0].raw_text[1:-1].lower()
            if second is not None and len(second) == 1 and second[0].kind is TokenKind.STRING_LITERAL
            else None
        )
        callee = procedures.get(target) if target else None
        call_type = "".join(token_text(tok) for tok in args[2]) if len(args) > 2 else None
        if (
            len(args) == 3
            and len(args[0]) == 1
            and token_text(args[0][0]) in selves
            and callee is not None
            and target is not None
            and second is not None
            and (call_type == "vbmethod" or call_type == "1")
            and not _is_private(callee)
            and len(callee.params) == 0
        ):
            return _ProcedureCall(target, toks[by_name], second[0])
    head = 1 if token_text(_at(toks, 0)) == "call" else 0
    head_token_name = token_name(_at(toks, head))
    head_name = head_token_name.lower() if head_token_name is not None else None
    head_proc = procedures.get(head_name) if head_name else None
    assigns = any(tok.raw_text == "=" for tok in toks) and head == 0
    if (
        head_proc is not None
        and head_name is not None
        and not assigns
        and _raw_at(toks, head + 1) != "."
        and _raw_at(toks, head + 1) != "!"
    ):
        return _ProcedureCall(head_name, toks[0], toks[head])
    # `Me.Go`, `Twice = Me.Twice`: through Me a member is always called, never the
    # return variable (XLIDE issue #338, measured in Excel 16.0). Me reaches a Public
    # or Friend member only.
    for i in range(2, len(toks)):
        name = token_name(toks[i])
        lower = name.lower() if name is not None else None
        callee = procedures.get(lower) if lower else None
        if (
            callee is not None
            and lower is not None
            and toks[i - 1].raw_text == "."
            and token_text(toks[i - 2]) in selves
            and _raw_at(toks, i - 3) != "."
            and not _is_private(callee)
            and (len(callee.params) == 0 or _raw_at(toks, i + 1) == "(")
        ):
            return _ProcedureCall(lower, toks[i - 2], toks[i])
    for i in range(1, len(toks) - 1):
        name = token_name(toks[i])
        lower = name.lower() if name is not None else None
        callee = procedures.get(lower) if lower else None
        if (
            callee is None
            or lower is None
            or toks[i + 1].raw_text != "("
            or toks[i - 1].raw_text == "."
            or toks[i - 1].raw_text == "!"
        ):
            continue
        # With arguments it takes none of, a Variant Function calls itself and
        # indexes what comes back; a typed one does not compile.
        variant = not callee.return_type or js_trim(callee.return_type).lower() == "variant"
        if len(callee.params) == 0 and _raw_at(toks, i + 2) != ")" and not variant:
            continue
        return _ProcedureCall(lower, toks[i], toks[i])
    return None


def _same_arguments_at(toks: Sequence[VbaToken], open_: int, proc: ProcedureNode) -> bool:
    """Whether the argument list opening at `open_` passes the procedure's own
    parameters, in order."""
    close = match_paren_from(toks, open_)
    args = split_top_level_token_groups(toks, open_ + 1, ",", close) if close > open_ + 1 else []
    if len(args) != len(proc.params):
        return False
    for arg, param in zip(args, proc.params):
        name = token_name(arg[0]) if len(arg) == 1 else None
        if name is None or name.lower() != param.name.lower():
            return False
    return True


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
