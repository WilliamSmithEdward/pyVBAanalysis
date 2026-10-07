"""Collection-independent form-content diagnostics from XLIDE 11.1.0.

Designer pages and initial list contents cannot establish runtime bounds.
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence

from ...conditional import ConditionalActivityTracker
from ...lexer.token_helpers import match_paren_from
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import ModuleNode, ProcedureNode, Span, VariableGroupNode, iter_body_nodes, LeafStatementNode
from ...symbols.project_index import FormControlInfo
from ..context import PushFn
from ..walker import active_module_members, for_each_statement, statement_and_branch_spans, statement_tokens, token_name, token_text

_Judged = Mapping[str, FormControlInfo]
_DIGITS = re.compile(r"^[0-9]+\Z")


def check_form_contents(source: str, mod: ModuleNode, controls: Sequence[FormControlInfo] | None, name_mentions: Mapping[str, int] | None, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    if not controls:
        return
    known = {control.name.lower(): control for control in controls}
    for proc in active_module_members(mod, activity):
        if not isinstance(proc, ProcedureNode):
            continue
        available = {name: control for name, control in known.items() if not _declares_name(proc, name)}

        def visit(stmt: LeafStatementNode) -> None:
            for span in statement_and_branch_spans(stmt):
                toks = statement_tokens(source, span)
                start = 2 if token_text(_at(toks, 0)) == "me" and _raw(toks, 1) == "." else 0
                for i in range(len(toks)):
                    control = _named_control(toks, i, available)
                    if control is None or _raw(toks, i + 1) != ".":
                        continue
                    member = token_text(_at(toks, i + 2))
                    is_list = control.type in ("MSForms.ListBox", "MSForms.ComboBox")

                    def at(end: int) -> Span:
                        return Span(span.start + toks[i].start, span.start + toks[end].end)

                    if is_list and i == start and member == "listindex" and _raw(toks, i + 3) == "=":
                        rhs = toks[i + 4:]
                        r = 2 if token_text(_at(rhs, 0)) == "me" and _raw(rhs, 1) == "." else 0
                        name = token_name(_at(rhs, r))
                        same_count = len(rhs) == r + 3 and name is not None and name.lower() == control.name.lower() and _raw(rhs, r + 1) == "." and token_text(_at(rhs, r + 2)) == "listcount"
                        value = _signed_literal(rhs)
                        if same_count or (value is not None and value < -1):
                            push("hostPropertyValueOutOfRange", f"ListIndex runs from -1 to {control.name}.ListCount - 1; this value is outside it. This will raise Run-time error '380': Could not set the ListIndex property. Invalid property value.", at(len(toks) - 1))
                    if ((is_list and member in ("list", "selected")) or (control.type == "MSForms.MultiPage" and member == "pages")) and _raw(toks, i + 3) == "(":
                        close = match_paren_from(toks, i + 3)
                        index = _signed_literal(toks[i + 4:close]) if close > 0 else None
                        if index is None or index >= 0:
                            continue
                        if member == "pages":
                            push("runtimeArgumentValue", "Pages indexes start at 0. This will raise Run-time error '5': Invalid procedure call or argument.", at(close))
                        elif member == "selected" and control.type == "MSForms.ListBox" and _raw(toks, close + 1) == "=":
                            push("hostPropertyValueOutOfRange", "Selected indexes start at 0. This will raise Run-time error '380': Could not set the Selected property. Invalid property value.", at(close))
                        elif member == "list" and i != start:
                            push("hostPropertyValueOutOfRange", "List row indexes start at 0. This will raise Run-time error '381': Could not get the List property. Invalid property array index.", at(close))

        for_each_statement(proc.body, visit, activity)


def _at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """`toks[index]`, or None out of range the way a JavaScript index reads undefined."""
    return toks[index] if 0 <= index < len(toks) else None


def _raw(toks: Sequence[VbaToken], index: int) -> str | None:
    tok = _at(toks, index)
    return tok.raw_text if tok is not None else None


def _declares_name(member: ProcedureNode, lower: str) -> bool:
    """Whether the procedure declares a parameter or local of this name, which hides the control."""
    if any(param.name.lower() == lower for param in member.params):
        return True
    # An If block's flat body holds every arm's statements, so its arms need no
    # second visit.
    return any(
        isinstance(node, VariableGroupNode) and any(decl.name.lower() == lower for decl in node.declarations)
        for node in iter_body_nodes(member.body)
    )


def _named_control(toks: Sequence[VbaToken], i: int, controls: _Judged) -> FormControlInfo | None:
    """The control a token names: `L1` or `Me.L1`, never `f.L1`, which is another
    instance's. None for any other token."""
    name = token_name(toks[i])
    lower = name.lower() if name is not None else None
    if not lower or toks[i].kind is not TokenKind.IDENTIFIER or lower not in controls:
        return None
    if _raw(toks, i - 1) == "." and not (token_text(_at(toks, i - 2)) == "me" and _raw(toks, i - 3) != "."):
        return None
    return controls.get(lower)


def _signed_literal(group: Sequence[VbaToken]) -> int | None:
    """A whole-number literal, a minus sign allowed."""
    toks = [tok for tok in group if tok.kind is not TokenKind.COMMENT]
    negative = len(toks) == 2 and toks[0].raw_text == "-"
    tok = _at(toks, 1 if negative else 0)
    if (
        tok is None
        or len(toks) != (2 if negative else 1)
        or tok.kind is not TokenKind.INTEGER_LITERAL
        or _DIGITS.search(tok.raw_text) is None
    ):
        return None
    return -int(tok.raw_text) if negative else int(tok.raw_text)
