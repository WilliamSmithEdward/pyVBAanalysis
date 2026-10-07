"""Rule: DeleteSetting of a key, section or application the procedure already
deleted (XLIDE issue #700, measured in Excel 16.0).

`DeleteSetting "App", "S", "k"` a second time raises 5, Invalid procedure call or
argument, and so does deleting a key of a section or application deleted above,
or that section or application again. A SaveSetting of the application in
between, or a call, a label or a block that may save or delete settings, ends
what is known. Under On Error Resume Next nothing is reported, and the setting
is gone either way.

Deleting a key the code never saved raises 5 too, but the registry keeps
settings from earlier runs, so that is not for a static rule.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/deletedSettings.ts.
"""

from __future__ import annotations

import re
from collections.abc import Iterator, Sequence

from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import statement_label_declaration
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import BodyNode, ModuleNode, ProcedureNode, Span, is_leaf_statement
from ..call_extraction import string_literal_value
from ..context import PushFn
from ..walker import active_module_members, is_inactive_node, statement_tokens_after_leading_label, token_text

_DELETE_SETTING = re.compile(r"\bdeletesetting\b", re.IGNORECASE | re.ASCII)


def _setting_arguments(toks: Sequence[VbaToken]) -> list[str | None]:
    """The literal arguments of a SaveSetting or DeleteSetting, lowercased; None for one that is not a literal."""
    open_index = 1 if len(toks) > 1 and toks[1].raw_text == "(" else -1
    end = len(toks) if open_index < 0 else match_paren_from(toks, open_index)
    if end < 0:
        return [None]
    out: list[str | None] = []
    for arg in split_top_level_token_groups(toks, 1 if open_index < 0 else 2, ",", end):
        value = [tok for tok in arg if tok.kind is not TokenKind.COMMENT]
        out.append(
            string_literal_value(value[0].raw_text).lower()
            if len(value) == 1 and value[0].kind is TokenKind.STRING_LITERAL
            else None
        )
    return out


def _forget_application(gone: set[str], app: str) -> None:
    for fact in list(gone):
        if fact.startswith(f"{app}|"):
            gone.discard(fact)


def check_deleted_settings(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode) or _DELETE_SETTING.search(
            source[member.span.start : member.span.end]
        ) is None:
            continue
        # What the procedure deleted, as `app|section|key`, with `*` for all.
        gone: set[str] = set()
        resume_next = False
        # Upstream recurses into each block between two clears; the None entry
        # is the clear after a block's body.
        stack: list[Iterator[BodyNode] | None] = [iter(member.body)]
        while stack:
            top = stack[-1]
            if top is None:
                stack.pop()
                gone.clear()
                continue
            node = next(top, None)
            if node is None:
                stack.pop()
                continue
            if is_inactive_node(activity, node):
                continue
            if not is_leaf_statement(node):
                # A block may save or delete settings on one path only.
                gone.clear()
                body = getattr(node, "body", None)
                stack.append(None)
                if isinstance(body, list):
                    stack.append(iter(body))
                continue
            if statement_label_declaration(source, node.span):
                gone.clear()
            toks = statement_tokens_after_leading_label(source, node.span)
            head = token_text(toks[0] if toks else None)
            second = toks[1].raw_text if len(toks) > 1 else None
            if head == "on" and token_text(toks[1] if len(toks) > 1 else None) == "error":
                resume_next = token_text(toks[2] if len(toks) > 2 else None) == "resume"
                continue
            if head == "savesetting" and second != "=":
                app = _setting_arguments(toks)[0]
                if app is None:
                    gone.clear()
                else:
                    _forget_application(gone, app)
                continue
            if head == "deletesetting" and second != "=":
                args = _setting_arguments(toks)
                app = args[0]
                section = args[1] if len(args) > 1 else None
                key = args[2] if len(args) > 2 else None
                if (
                    app is None
                    or len(args) > 3
                    or (len(args) > 1 and section is None)
                    or (len(args) > 2 and key is None)
                ):
                    # A setting the rule cannot name: what it deletes is not known.
                    if app is None:
                        gone.clear()
                    else:
                        _forget_application(gone, app)
                    continue
                covered = (
                    f"{app}|*|*" in gone
                    or (section is not None and f"{app}|{section}|*" in gone)
                    or (key is not None and f"{app}|{section}|{key}" in gone)
                )
                if covered and not resume_next:
                    what = "key" if key is not None else "section" if section is not None else "application"
                    push(
                        "runtimeArgumentValue",
                        f"DeleteSetting finds no such {what}: this procedure deleted it above. This will raise "
                        "Run-time error '5': Invalid procedure call or argument.",
                        Span(node.span.start + toks[0].start, node.span.start + toks[-1].end),
                    )
                gone.add(f"{app}|{section if section is not None else '*'}|{key if key is not None else '*'}")
                continue
            # A call may save or delete settings.
            gone.clear()
