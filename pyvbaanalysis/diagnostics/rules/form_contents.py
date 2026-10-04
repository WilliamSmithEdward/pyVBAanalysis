"""What a UserForm's designer puts in its controls, against what the form's own
code asks of them (XLIDE issue #315, each case measured in Excel 16.0).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/formContents.ts.

  Mp.Pages(9) and Mp.Pages("Page9") on a two-page MultiPage   error 5
  Mp.Value = 5 on the same MultiPage                          error 380
  L1.AddItem "a" then L1.ListIndex = 5 on an unbound ListBox  error 380
  L1.Selected(5) = True after the same AddItem                error 380
  L1.List(5) read after the same AddItem                      error 381
  L1.ListIndex = L1.ListCount, whatever the list holds        error 380

The designer knows a MultiPage's pages and whether a list has a RowSource.
Code anywhere can add pages or items, so a control is judged only inside the
one procedure of its form that names it: no other procedure or module names
it, and nothing in the project reaches a form's Controls collection.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterator, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import statement_label_declaration
from ...lexer.token_helpers import identifier_words, match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import (
    BodyNode,
    IfBlockNode,
    ModuleNode,
    ProcedureNode,
    Span,
    VariableGroupNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ...symbols.project_index import FormControlInfo
from ..call_extraction import string_literal_value
from ..context import PushFn
from ..walker import (
    active_module_members,
    block_header_line_span,
    is_inactive_node,
    statement_and_branch_spans,
    statement_tokens,
    token_name,
    token_text,
)

# A list or MultiPage the designer knows the contents of, by lowercased name.
_Judged = Mapping[str, FormControlInfo]

# Members of a list control that read it without changing what it holds.
_LIST_READS: frozenset[str] = frozenset(
    {
        "listcount", "listindex", "value", "text", "name", "visible", "enabled", "tag", "locked",
        "top", "left", "width", "height", "setfocus", "boundcolumn", "textcolumn", "multiselect",
    }
)

_DIGITS = re.compile(r"^[0-9]+\Z")


def _at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """`toks[index]`, or None out of range the way a JavaScript index reads undefined."""
    return toks[index] if 0 <= index < len(toks) else None


def _raw(toks: Sequence[VbaToken], index: int) -> str | None:
    tok = _at(toks, index)
    return tok.raw_text if tok is not None else None


def check_form_contents(
    source: str,
    mod: ModuleNode,
    controls: Sequence[FormControlInfo] | None,
    name_mentions: Mapping[str, int] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    if not controls:
        return
    lists: dict[str, FormControlInfo] = {}
    page_sets: dict[str, FormControlInfo] = {}
    for control in controls:
        lower = control.name.lower()
        if control.type == "MSForms.ListBox" or control.type == "MSForms.ComboBox":
            lists[lower] = control
        elif control.type == "MSForms.MultiPage" and control.pages is not None:
            page_sets[lower] = control
    if len(lists) == 0 and len(page_sets) == 0:
        return
    # Code that reaches a form's Controls, here or in any other module, can
    # reach every control without naming it.
    controls_reached = (name_mentions.get("controls", 0) if name_mentions is not None else 0) > 0
    where = _mention_spans(source, set(lists.keys()) | set(page_sets.keys()))
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        own_only = _own_only_test(member, name_mentions, controls_reached, where)
        _check_list_against_count(source, member, lists, own_only, activity, push)
        _check_pages(source, member, page_sets, own_only, activity, push)


def _own_only_test(
    member: ProcedureNode,
    name_mentions: Mapping[str, int] | None,
    controls_reached: bool,
    where: Mapping[str, list[int]],
) -> Callable[[str], bool]:
    def own_only(lower: str) -> bool:
        return (
            name_mentions is not None
            and not controls_reached
            and name_mentions.get(lower, 0) == 1
            and all(member.span.start <= at < member.span.end for at in where.get(lower, []))
            and not _declares_name(member, lower)
        )

    return own_only


def _mention_spans(source: str, names: AbstractSet[str]) -> dict[str, list[int]]:
    """Where the module names each of `names`, as an identifier or a word in a string."""
    out: dict[str, list[int]] = {}
    for token in tokenize_cached(source):
        if token.kind is TokenKind.IDENTIFIER:
            words = [token.raw_text.lower()]
        elif token.kind is TokenKind.STRING_LITERAL:
            words = identifier_words(token.raw_text)
        else:
            continue
        for word in words:
            if word in names:
                out.setdefault(word, []).append(token.start)
    return out


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


def _jumps(source: str, member: ProcedureNode) -> bool:
    """Whether the procedure has a label or a jump, which can run a line twice."""
    for node in iter_body_nodes(member.body):
        if is_leaf_statement(node):
            head = token_text(_at(statement_tokens(source, node.span), 0))
            if (
                statement_label_declaration(source, node.span)
                or head == "goto"
                or head == "gosub"
                or head == "resume"
                or head == "on"
            ):
                return True
    return False


def _children(node: BodyNode) -> list[BodyNode]:
    """A block's statements: each If arm's, in order, or its body."""
    if isinstance(node, IfBlockNode):
        return [child for branch in node.branches for child in branch.body]
    body = getattr(node, "body", None)
    return body if isinstance(body, list) else []


def _walk(body: Sequence[BodyNode], activity: ConditionalActivityTracker | None) -> Iterator[tuple[BodyNode, bool]]:
    """Every active node in the order upstream's recursive visit takes them,
    with whether it sits at the procedure's top level. Runs on an explicit
    stack: each If arm's statements, or a block's body, after the block."""
    stack: list[tuple[Iterator[BodyNode], bool]] = [(iter(body), True)]
    while stack:
        nodes, top_level = stack[-1]
        for node in nodes:
            if is_inactive_node(activity, node):
                continue
            yield node, top_level
            if not is_leaf_statement(node):
                stack.append((iter(_children(node)), False))
                break
        else:
            stack.pop()


def _check_list_against_count(
    source: str,
    member: ProcedureNode,
    lists: _Judged,
    own_only: Callable[[str], bool],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Follows each unbound list's item count down the procedure's top level,
    from empty: AddItem adds one, Clear empties it, `List = Array(...)` sets it.
    A line that may change it any other way, or a block that names it, ends
    what is known."""
    if len(lists) == 0:
        return
    counts: dict[str, int | None] = {}
    for lower, control in lists.items():
        counts[lower] = 0 if control.list_starts_empty is True and own_only(lower) else None
    judged = not _jumps(source, member)
    for node, top_level in _walk(member.body, activity):
        if not is_leaf_statement(node):
            # A block may run its lines any number of times.
            toks = statement_tokens(source, node.span)
            for i in range(len(toks)):
                named = _named_control(toks, i, lists)
                if named is not None:
                    counts[named.name.lower()] = None
            continue
        spans = statement_and_branch_spans(node)
        for span in spans:
            toks = statement_tokens(source, span)
            _list_statement(
                toks, span, lists, counts if judged and top_level and len(spans) == 1 else {}, counts, push
            )


def _list_statement(
    toks: Sequence[VbaToken],
    span: Span,
    lists: _Judged,
    known: Mapping[str, int | None],
    counts: dict[str, int | None],
    push: PushFn,
) -> None:
    """One statement against the counts known before it: reports what it asks
    past the list, then moves `counts` on. `known` is empty where the statement
    may run other than once, in order."""

    def at(first: VbaToken, last: VbaToken | None = None) -> Span:
        end = last if last is not None else first
        return Span(span.start + first.start, span.start + end.end)

    for i in range(len(toks)):
        control = _named_control(toks, i, lists)
        if control is None:
            continue
        lower = control.name.lower()
        # `L1.ListIndex = L1.ListCount` is past any list (measured).
        member = token_text(_at(toks, i + 2)) if _raw(toks, i + 1) == "." else None
        statement_start = 2 if token_text(_at(toks, 0)) == "me" else 0
        target = i == statement_start
        if target and member == "listindex" and _raw(toks, i + 3) == "=" and len(toks) >= i + 7:
            rest = toks[i + 4 :]
            rest_start = 2 if token_text(_at(rest, 0)) == "me" and _raw(rest, 1) == "." else 0
            rest_name = token_name(_at(rest, rest_start))
            if (
                len(rest) == rest_start + 3
                and rest_name is not None
                and rest_name.lower() == lower
                and rest[rest_start + 1].raw_text == "."
                and token_text(rest[rest_start + 2]) == "listcount"
            ):
                push(
                    "hostPropertyValueOutOfRange",
                    f"ListIndex runs from -1 to {control.name}.ListCount - 1, so {control.name}.ListCount "
                    "is past the list. This will raise Run-time error '380': Could not set the ListIndex "
                    "property. Invalid property value.",
                    at(toks[i], toks[-1]),
                )
                return
        count = known.get(lower)
        if not target:
            # Read in an expression: `L1.List(5)` past the list raises 381.
            if member == "list" and _raw(toks, i + 3) == "(" and count is not None:
                close = match_paren_from(toks, i + 3)
                args = split_top_level_token_groups(toks, i + 4, ",", close) if close > 0 else []
                index = _signed_literal(args[0]) if len(args) == 1 else None
                if index is not None and (index < 0 or index >= count):
                    push(
                        "hostPropertyValueOutOfRange",
                        f"{control.name} {_holds(count)} here, so List({index}) is past it. This will raise "
                        "Run-time error '381': Could not get the List property. Invalid property array index.",
                        at(toks[i], toks[close]),
                    )
                continue
            if member is None or (
                member not in _LIST_READS and member != "list" and member != "column" and member != "selected"
            ):
                counts[lower] = None
            continue
        args = split_top_level_token_groups(toks, i + 3, ",")
        if member == "additem":
            counts[lower] = None if count is None else count + 1
        elif member == "clear":
            counts[lower] = None if counts.get(lower) is None else 0
        elif member == "removeitem":
            index = _signed_literal(args[0]) if len(args) == 1 else None
            counts[lower] = (
                count - 1 if count is not None and index is not None and 0 <= index < count else None
            )
        elif member == "list":
            # `L1.List = Array("a", "b")` holds two.
            value = toks[i + 4 :]
            close = (
                match_paren_from(value, 1)
                if token_text(_at(value, 0)) == "array" and _raw(value, 1) == "("
                else -1
            )
            items = (
                split_top_level_token_groups(value, 2, ",", close)
                if close == len(value) - 1 and _raw(toks, i + 3) == "="
                else None
            )
            counts[lower] = (
                (0 if close == 2 else len(items)) if items is not None and counts.get(lower) is not None else None
            )
        elif member == "listindex":
            value_index = _signed_literal(toks[i + 4 :]) if _raw(toks, i + 3) == "=" else None
            if count is not None and value_index is not None and (value_index < -1 or value_index >= count):
                runs = "only to -1" if count == 0 else f"from -1 to {count - 1}"
                push(
                    "hostPropertyValueOutOfRange",
                    f"{control.name} {_holds(count)} here, so ListIndex runs {runs}; {value_index} is outside "
                    "it. This will raise Run-time error '380': Could not set the ListIndex property. Invalid "
                    "property value.",
                    at(toks[i], toks[-1]),
                )
        elif member == "selected":
            close = match_paren_from(toks, i + 3) if _raw(toks, i + 3) == "(" else -1
            index = _signed_literal(toks[i + 4 : close]) if close > 0 else None
            if (
                control.type == "MSForms.ListBox"
                and count is not None
                and index is not None
                and (index < 0 or index >= count)
                and _raw(toks, close + 1) == "="
            ):
                push(
                    "hostPropertyValueOutOfRange",
                    f"{control.name} {_holds(count)} here, so Selected({index}) is past it. This will raise "
                    "Run-time error '380': Could not set the Selected property. Invalid property value.",
                    at(toks[i], toks[close]),
                )
        elif member is None or member not in _LIST_READS:
            counts[lower] = None


def _holds(count: int) -> str:
    if count == 0:
        return "holds no items"
    return f"holds {count} item{'' if count == 1 else 's'}"


@dataclass(frozen=True, slots=True)
class _PageUse:
    control: FormControlInfo
    toks: Sequence[VbaToken]
    i: int
    span: Span


def _check_pages(
    source: str,
    member: ProcedureNode,
    page_sets: _Judged,
    own_only: Callable[[str], bool],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """`Mp.Pages(9)`, `Mp.Pages("Page9")` and `Mp.Value = 5` on a MultiPage whose
    pages the designer lists, in a procedure that cannot have changed them: it
    names the MultiPage only as `Mp.Pages(...)`, `Mp.Pages.Count` or one of its
    own properties."""
    candidates = {lower: control for lower, control in page_sets.items() if own_only(lower)}
    if len(candidates) == 0:
        return
    uses: list[_PageUse] = []
    unsure: set[str] = set()
    for node, _top_level in _walk(member.body, activity):
        # A block's opening line names what it reads like any statement; a With
        # over the MultiPage can then reach it unnamed.
        leaf = is_leaf_statement(node)
        spans = statement_and_branch_spans(node) if is_leaf_statement(node) else [block_header_line_span(source, node.span)]
        for span in spans:
            toks = statement_tokens(source, span)
            with_header = not leaf and token_text(_at(toks, 0)) == "with"
            for i in range(len(toks)):
                control = _named_control(toks, i, candidates)
                if control is None:
                    continue
                following = token_text(_at(toks, i + 2)) if _raw(toks, i + 1) == "." else None
                if (
                    with_header
                    or following is None
                    or (
                        following == "pages"
                        and _raw(toks, i + 3) != "("
                        and not (_raw(toks, i + 3) == "." and token_text(_at(toks, i + 4)) == "count")
                    )
                ):
                    unsure.add(control.name.lower())
                else:
                    uses.append(_PageUse(control, toks, i, span))
    for use in uses:
        control, use_toks, i, span = use.control, use.toks, use.i, use.span
        if control.name.lower() in unsure:
            continue
        pages = control.pages
        assert pages is not None

        def at(tok: VbaToken, span: Span = span) -> Span:
            return Span(span.start + tok.start, span.start + tok.end)

        following = token_text(_at(use_toks, i + 2))
        count = len(pages)
        plural = "" if count == 1 else "s"
        if following == "pages" and _raw(use_toks, i + 3) == "(":
            close = match_paren_from(use_toks, i + 3)
            arg = use_toks[i + 4 : close] if close > 0 else []
            index = _signed_literal(arg)
            if index is not None and (index < 0 or index >= count):
                push(
                    "runtimeArgumentValue",
                    f"The MultiPage {control.name} has {count} page{plural}, indexed 0 to {count - 1}; "
                    f"{index} is none of them. This will raise Run-time error '5': Invalid procedure call "
                    "or argument.",
                    at(arg[-1]),
                )
            elif len(arg) == 1 and arg[0].kind is TokenKind.STRING_LITERAL:
                name = string_literal_value(arg[0].raw_text)
                if not any(page.lower() == name.lower() for page in pages):
                    push(
                        "runtimeArgumentValue",
                        f"The MultiPage {control.name} has no page named \"{name}\". This will raise "
                        "Run-time error '5': Invalid procedure call or argument.",
                        at(arg[0]),
                    )
        elif (
            following == "value"
            and _raw(use_toks, i + 3) == "="
            and (i == 0 or (i == 2 and token_text(_at(use_toks, 0)) == "me"))
        ):
            value = _signed_literal(use_toks[i + 4 :])
            if value is not None and value >= count:
                push(
                    "hostPropertyValueOutOfRange",
                    f"The MultiPage {control.name} has {count} page{plural}, so Value runs from 0 to "
                    f"{count - 1}; {value} is past it. This will raise Run-time error '380': Could not set "
                    "the Value property. Invalid property value.",
                    Span(span.start + use_toks[i].start, span.start + use_toks[-1].end),
                )
