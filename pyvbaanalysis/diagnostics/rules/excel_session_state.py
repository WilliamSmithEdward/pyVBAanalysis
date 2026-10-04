"""Rule family: what earlier statements of a procedure leave in Excel, and the
1004 a later one then raises (XLIDE issue #308).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/excelSessionState.ts.

Measured in Excel 16.0 (build 20430, 2026-10-02); each compiles and raises
every time it runs.

 - `Application.CutCopyMode = False` empties Excel's clipboard, and a
   PasteSpecial onto a range before anything is copied again raises 1004,
   "PasteSpecial method of Range class failed". A Copy, a Cut, or a call into
   code that may copy ends what is known.
 - Two sheets the procedure added, `Set w = Worksheets.Add`, cannot share a
   name: sheet names ignore case, so `w1.Name = "Aa"` then `w2.Name = "aa"`
   raises 1004, "That name is already taken".
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterable, Mapping, Sequence
from collections.abc import Set as AbstractSet

from ...completion.member_access import MemberCompletionContext, resolve_receiver_type_at
from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import statement_label_declaration
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import BodyNode, ModuleNode, ProcedureNode, Span, StatementNode, is_leaf_statement
from ..call_extraction import string_literal_value
from ..context import PushFn
from ..dataflow import BlockEnteringState, walk_entering_blocks
from ..walker import (
    active_module_members,
    set_assignment_target,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .shared import names_in

# The state key for an empty clipboard; no VBA name holds a `#`.
_CLIPBOARD = "#clipboard"

# Words that may put something on the clipboard, or run code that does.
_CLIPBOARD_WORDS: frozenset[str] = frozenset(
    {"copy", "cut", "copypicture", "run", "sendkeys", "execute", "doevents", "call"}
)

_SHEETS_ADD = re.compile(
    r"^(?:(?:activeworkbook|thisworkbook|application)\.)?(?:worksheets|sheets)\.add(?:\(\))?\Z"
)


def check_excel_session_state(
    source: str,
    mod: ModuleNode,
    callables: AbstractSet[str],
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    model = member_ctx.model
    host = model.get("hostName") if model is not None else None
    if host is not None and host != "Excel":
        return

    # Anything that may copy, or a call into the project's own code.
    def may_copy(toks: Sequence[VbaToken]) -> bool:
        for tok in toks:
            word = token_text(tok)
            if word in _CLIPBOARD_WORDS or (token_name(tok) is not None and word in callables):
                return True
        return False

    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        _check_procedure(source, member, may_copy, member_ctx, activity, push)


def _check_procedure(
    source: str,
    member: ProcedureNode,
    may_copy: Callable[[Sequence[VbaToken]], bool],
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    # A sheet the procedure added, by variable, and the name it gave it ('' unknown).
    state: dict[str, str] = {}

    def forget(names: Iterable[str]) -> None:
        for lower in names:
            state.pop(lower, None)

    def visit(node: BodyNode) -> None:
        if not is_leaf_statement(node):
            return
        toks = statement_tokens_after_leading_label(source, node.span)
        if statement_label_declaration(source, node.span) or token_text(_at(toks, 0)) == "gosub":
            state.clear()
        if isinstance(node, StatementNode) and node.single_line_if_branches:
            forget(names_in(source, node.span))
            state.pop(_CLIPBOARD, None)
            return

        def at(first: VbaToken, last: VbaToken) -> Span:
            return Span(node.span.start + first.start, node.span.start + last.end)

        # `Application.CutCopyMode = False` (or 0).
        if (
            len(toks) == 5
            and token_text(toks[0]) == "application"
            and toks[1].raw_text == "."
            and token_text(toks[2]) == "cutcopymode"
            and toks[3].raw_text == "="
            and (token_text(toks[4]) == "false" or toks[4].raw_text == "0")
        ):
            state[_CLIPBOARD] = ""
            return
        if may_copy(toks):
            state.pop(_CLIPBOARD, None)
        paste = next(
            (
                k
                for k, tok in enumerate(toks)
                if token_text(tok) == "pastespecial" and k >= 1 and toks[k - 1].raw_text == "."
            ),
            -1,
        )
        if (
            paste > 0
            and _CLIPBOARD in state
            and resolve_receiver_type_at(source, node.span.start + toks[paste - 1].end, member_ctx)
            == "Excel.Range"
        ):
            push(
                "pasteWithNothingCopied",
                "Nothing is copied here: Application.CutCopyMode = False emptied the clipboard, and no "
                "Copy or Cut came after it. This will raise Run-time error '1004': PasteSpecial method "
                "of Range class failed.",
                at(toks[paste], toks[paste]),
            )
        # `Set w = Worksheets.Add`: a sheet of its own.
        set_target = set_assignment_target(source, node.span)
        if set_target is not None:
            lower = set_target[0].lower()
            forget(names_in(source, node.span))
            value = "".join(
                token_text(tok) for tok in set_target[2] if tok.kind is not TokenKind.COMMENT
            )
            if _SHEETS_ADD.search(value) is not None:
                state[lower] = ""
            return
        # `w.Name = "Aa"`.
        sheet_name = token_name(toks[0]) if len(toks) == 5 else None
        sheet = (
            sheet_name.lower()
            if sheet_name is not None
            and toks[1].raw_text == "."
            and token_text(toks[2]) == "name"
            and toks[3].raw_text == "="
            else None
        )
        if sheet is not None and sheet in state:
            name = (
                string_literal_value(toks[4].raw_text) if toks[4].kind is TokenKind.STRING_LITERAL else ""
            )
            taken = (
                None
                if name == ""
                else next(
                    (
                        (other, held)
                        for other, held in state.items()
                        if other != sheet and other != _CLIPBOARD and held.lower() == name.lower()
                    ),
                    None,
                )
            )
            if taken is not None:
                push(
                    "sheetNameInvalid",
                    f"\"{name}\" is the name the code gave '{taken[0]}', another sheet it added, and "
                    "sheet names ignore case. This will raise Run-time error '1004': That name is "
                    "already taken.",
                    at(toks[4], toks[4]),
                )
            state[sheet] = name
            return
        forget([lower for lower in names_in(source, node.span) if lower != _CLIPBOARD])

    def snapshot() -> dict[str, str]:
        return dict(state)

    def restore(saved: Mapping[str, str]) -> None:
        state.clear()
        for lower, held in saved.items():
            state[lower] = held

    def touches(stmt: BodyNode) -> set[str]:
        names = set(names_in(source, stmt.span))
        if may_copy(statement_tokens_after_leading_label(source, stmt.span)):
            names.add(_CLIPBOARD)
        return names

    walk_entering_blocks(
        source,
        member.body,
        lambda node: activity is not None and activity.is_inactive(node.span),
        visit,
        BlockEnteringState(snapshot=snapshot, restore=restore, forget=forget, touches=touches),
    )


def _at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """`toks[index]`, or None out of range the way a JavaScript index reads undefined."""
    return toks[index] if 0 <= index < len(toks) else None
