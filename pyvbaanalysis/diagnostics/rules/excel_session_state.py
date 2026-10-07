"""Workbook names for freshly added local sheets owned by ThisWorkbook.

Clipboard contents and sheet contents remain runtime facts (XLIDE 11.1.0).
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from collections.abc import Set as AbstractSet

from ...completion.member_access import MemberCompletionContext
from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import statement_label_declaration
from ...lexer.token_kinds import TokenKind
from ...parser.nodes import BodyNode, ModuleNode, ProcedureNode, Span, StatementNode, VariableGroupNode, is_leaf_statement
from ..call_extraction import string_literal_value
from ..context import PushFn
from ..dataflow import BlockEnteringState, walk_entering_blocks
from ..walker import active_module_members, for_each_variable_group, set_assignment_target, statement_tokens_after_leading_label, token_name, token_text
from .shared import names_in


def check_excel_session_state(source: str, mod: ModuleNode, callables: AbstractSet[str], member_ctx: MemberCompletionContext, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    host = member_ctx.model.get("hostName") if member_ctx.model is not None else None
    if host is not None and host != "Excel":
        return
    shadows_workbook = any(any(decl.name.lower() == "thisworkbook" for decl in node.declarations) if isinstance(node, VariableGroupNode) else (getattr(node, "name", None) or "").lower() == "thisworkbook" for node in mod.members)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        if re.search(r"\bon\s+error\b", source[member.span.start:member.span.end], re.IGNORECASE):
            continue
        locals_: set[str] = set()

        def collect(group: VariableGroupNode) -> None:
            if not group.is_const and group.modifier.lower() != "static":
                locals_.update(decl.name.lower() for decl in group.declarations)

        for_each_variable_group(member.body, collect, activity)
        if shadows_workbook or "thisworkbook" in locals_ or any(param.name.lower() == "thisworkbook" for param in member.params):
            continue
        state: dict[str, str] = {}

        def forget(names: Iterable[str]) -> None:
            for lower in names:
                state.pop(lower, None)

        def visit(node: BodyNode) -> None:
            if not is_leaf_statement(node):
                return
            if statement_label_declaration(source, node.span) or (isinstance(node, StatementNode) and node.single_line_if_branches):
                state.clear()
            toks = statement_tokens_after_leading_label(source, node.span)
            target = set_assignment_target(source, node.span)
            if target:
                lower = target[0].lower()
                value = "".join(token_text(tok) for tok in target[2] if tok.kind is not TokenKind.COMMENT)
                if lower in locals_ and re.fullmatch(r"thisworkbook\.(?:worksheets|sheets)\.add(?:\(\))?", value):
                    for name in state:
                        state[name] = ""
                    state[lower] = ""
                else:
                    state.clear()
                return
            sheet_name = token_name(toks[0]) if len(toks) == 5 and toks[1].raw_text == "." and token_text(toks[2]) == "name" and toks[3].raw_text == "=" else None
            sheet = sheet_name.lower() if sheet_name else None
            if sheet is not None and sheet in state and toks[4].kind is TokenKind.STRING_LITERAL:
                name = string_literal_value(toks[4].raw_text)
                taken = next((other for other, held in state.items() if other != sheet and held and held.lower() == name.lower()), None)
                if taken:
                    push("sheetNameInvalid", f"\"{name}\" is the name the code gave '{taken}', another sheet it added to ThisWorkbook, and sheet names ignore case. This will raise Run-time error '1004': That name is already taken.", Span(node.span.start + toks[4].start, node.span.start + toks[4].end))
                state[sheet] = name
                return
            state.clear()

        def restore(saved: dict[str, str]) -> None:
            state.clear()
            state.update(saved)

        walk_entering_blocks(source, member.body, lambda node: activity is not None and activity.is_inactive(node.span), visit, BlockEnteringState(snapshot=lambda: dict(state), restore=restore, forget=forget, touches=lambda stmt: names_in(source, stmt.span) | state.keys()))
