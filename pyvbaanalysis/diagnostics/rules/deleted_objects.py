"""Rule: an object used after the statement that deletes or closes it (XLIDE
issue #294, each measured in Excel 16.0).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/deletedObjects.ts.

`ws.Delete` leaves the variable referring to a sheet that is gone: a member of it
then raises -2147221080, "Method 'Name' of object '_Worksheet' failed", and a
Range taken from it before raises 424. A closed Workbook raises the same
-2147221080; a deleted Shape or Name 424; an Unlisted ListObject 1004.
`ws Is Nothing` still runs and gives False, and a new `Set` ends what is known.

A sheet or Range taken from a workbook before it closed is gone with it, and so
are a closed Word document (5825) and the Ranges taken from it, a closed
PowerPoint presentation, a deleted Slide and a Slide of a closed presentation
(-2147188720) (issue #683, measured in Excel, Word and PowerPoint 16.0).

What Erase leaves in a Variant that held an array is the array rules' (issue
#420).

Only a straight run of a procedure's statements is followed: a block that names
a variable ends what is known of it.
"""

from __future__ import annotations

import re
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from bisect import bisect_right

from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import statement_label_declaration
from ...lexer.token_helpers import match_paren_from
from ...lexer.token_kinds import VbaToken, TokenKind
from ...parser.nodes import BodyNode, ModuleNode, ProcedureNode, Span, is_leaf_statement
from ...symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbolKind
from ...types.type_inference import procedure_symbol_for
from ...types.type_names import normalize_type
from ..context import PushFn
from ..walker import (
    active_module_members,
    is_inactive_node,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)


@dataclass(frozen=True, slots=True)
class _Ending:
    verb: str
    error: str
    cancellable: bool = False


# What deletes or closes each type, and what a member of it then raises.
_ENDINGS: dict[str, _Ending] = {
    "worksheet": _Ending("delete", "'-2147221080': Method 'MEMBER' of object '_Worksheet' failed", True),
    "workbook": _Ending("close", "'-2147221080': Method 'MEMBER' of object '_Workbook' failed", True),
    "shape": _Ending("delete", "'424': Object required"),
    "name": _Ending("delete", "'424': Object required"),
    "listobject": _Ending("unlist", "'1004': Application-defined or object-defined error"),
    # Word and PowerPoint (issue #683, measured in Word and PowerPoint 16.0).
    "document": _Ending("close", "'5825': Object has been deleted", True),
    "presentation": _Ending(
        "close", "'-2147188720': Presentation (unknown member) : Object does not exist", True
    ),
    "slide": _Ending("delete", "'-2147188720': Slide (unknown member) : Object does not exist"),
}

# Members of a sheet that give a Range of it.
_SHEET_RANGES: frozenset[str] = frozenset({"range", "cells", "rows", "columns", "usedrange"})

# Members of a workbook that give one of its sheets.
_WORKBOOK_SHEETS: frozenset[str] = frozenset({"sheets", "worksheets", "activesheet"})

_DERIVED_HOW_RE = re.compile(r"^a (?:range|sheet|slide) ")


def _derived_ending(owner_type: str, type_: str, chain: Sequence[str]) -> tuple[str, str] | None:
    """What an object taken from another becomes when that one ends, by the owner's
    type and its own (issues #294 and #683, measured in Excel, Word and PowerPoint
    16.0), as (kind, error): a Range of a deleted sheet or of a closed workbook's
    sheet raises 424, the sheet itself the Worksheet's error; a Range of a closed
    document 5825; a Slide of a closed presentation the Slide's error. `chain` is
    the member names after the owner, `ws.Range` as ['range']."""
    head = chain[0] if chain else None
    if owner_type == "worksheet" and type_ == "range" and head in _SHEET_RANGES:
        return "a range", "'424': Object required"
    if owner_type == "workbook" and head in _WORKBOOK_SHEETS:
        if type_ == "worksheet" and len(chain) == 1:
            return "a sheet", _ENDINGS["worksheet"].error
        if type_ == "range" and len(chain) == 2 and chain[1] in _SHEET_RANGES:
            return "a range", "'424': Object required"
    if owner_type == "document" and type_ == "range":
        return "a range", _ENDINGS["document"].error
    if owner_type == "presentation" and type_ == "slide" and head == "slides":
        return "a slide", _ENDINGS["slide"].error
    return None


@dataclass(frozen=True, slots=True)
class _Ended:
    # How it ended, for the message: "deleted on line 7".
    how: str
    error: str


@dataclass(frozen=True, slots=True)
class _Derived:
    owner: str
    # How the message names it: "a range".
    kind: str
    error: str


@dataclass(slots=True)
class _Run:
    """One straight run of statements and what it knows."""

    nodes: Iterator[BodyNode]
    ended: dict[str, _Ended] = field(default_factory=dict)
    ranges_of: dict[str, _Derived] = field(default_factory=dict)


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined below 0 and past the end."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw(tok: VbaToken | None) -> str | None:
    return tok.raw_text if tok is not None else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name is not None else None


def check_deleted_objects(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        if re.search(r"\bon\s+error\b", source[member.span.start:member.span.end], re.IGNORECASE):
            continue
        proc_sym = procedure_symbol_for(symbols, member)
        locals_ = [
            child
            for child in (proc_sym.children if proc_sym is not None else None) or []
            if child.kind is VbaSymbolKind.LOCAL_VARIABLE
            and child.visibility is not SymbolVisibility.STATIC
        ]
        type_of: dict[str, str] = {}
        for child in locals_:
            if not child.is_array:
                normalized = normalize_type(child.as_type)
                type_of[child.name.lower()] = normalized if normalized is not None else "variant"
        if not any(
            type_ in _ENDINGS or type_ == "range" or type_ == "variant"
            for type_ in type_of.values()
        ):
            continue
        _check_procedure(source, member.body, type_of, activity, push)


def _check_procedure(
    source: str,
    body: Sequence[BodyNode],
    type_of: Mapping[str, str],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    # Upstream runs each block's body as a nested run, recursively; the runs
    # are kept on an explicit stack, in the same order.
    stack: list[_Run] = [_Run(iter(body))]
    line_starts: list[int] | None = None
    while stack:
        run = stack[-1]
        node = next(run.nodes, None)
        if node is None:
            stack.pop()
            continue
        if is_inactive_node(activity, node):
            continue
        ended = run.ended
        ranges_of = run.ranges_of
        if not is_leaf_statement(node):
            # A block may change what it names; its own run starts afresh.
            text = source[node.span.start : node.span.end].lower()
            for lower in [*ended.keys(), *ranges_of.keys()]:
                if re.search(rf"\b{re.escape(lower)}\b", text, re.ASCII):
                    ended.pop(lower, None)
                    ranges_of.pop(lower, None)
            child = getattr(node, "body", None)
            if isinstance(child, list):
                stack.append(_Run(iter(child)))
            continue
        # A label is reached from wherever a jump to it runs: an error
        # handler from any line after its On Error GoTo, the Delete's
        # included (issue #584). Nothing known before it holds there.
        if statement_label_declaration(source, node.span) is not None:
            ended.clear()
            ranges_of.clear()
        toks = statement_tokens_after_leading_label(source, node.span)
        head = token_text(_at(toks, 0))
        # `Set x = ...` gives x a new object.
        if head == "set" and token_name(_at(toks, 1)) and _raw(_at(toks, 2)) == "=":
            target = toks[1].raw_text.lower()
            ended.pop(target, None)
            owner = _lower_name(_at(toks, 3))
            derived = (
                _derived_ending(
                    type_of.get(owner, ""), type_of.get(target, ""), _member_chain(toks, 4)
                )
                if owner
                else None
            )
            if owner and derived is not None:
                ranges_of[target] = _Derived(owner, derived[0], derived[1])
            else:
                ranges_of.pop(target, None)
            continue
        # `ws.Delete`, `wb.Close False`, `lo.Unlist`.
        subject = _lower_name(_at(toks, 0))
        ending = _ENDINGS.get(type_of.get(subject, "")) if subject else None
        if (
            subject
            and ending is not None
            and not ending.cancellable
            and _raw(_at(toks, 1)) == "."
            and token_text(_at(toks, 2)) == ending.verb
            and (len(toks) == 3 or ending.verb == "close")
        ):
            verbed = (
                "closed"
                if ending.verb == "close"
                else "unlisted"
                if ending.verb == "unlist"
                else "deleted"
            )
            if line_starts is None:
                line_starts = [0, *(match.end() for match in re.finditer(r"\r\n|\r|\n", source))]
            how = f"{verbed} on line {bisect_right(line_starts, node.span.start)}"
            ended[subject] = _Ended(how, ending.error)
            # What was taken from it, and from that in turn: a sheet of a
            # closed workbook and a range of that sheet.
            _end_from(subject, toks[0].raw_text, how, ended, ranges_of)
            continue
        for i, tok in enumerate(toks):
            passed_lower = _lower_name(tok)
            first_arg = i == 1 or (i == 2 and head == "call")
            before_token = _at(toks, i - 1)
            after = _at(toks, i + 1)
            if passed_lower and (before_token is not None and (before_token.raw_text in ("(", ",") or (first_arg and before_token.kind is TokenKind.IDENTIFIER))) and (after is None or after.raw_text in (")", ",")):
                ended.pop(passed_lower, None)
                ranges_of.pop(passed_lower, None)
        _report(toks, node.span.start, ended, push)
        # What the statement assigns or passes whole is no longer known.
        assigned = _lower_name(_at(toks, 1 if head == "let" else 0))
        assign_at = 2 if head == "let" else 1
        if assigned and _raw(_at(toks, assign_at)) == "=":
            ended.pop(assigned, None)
        for i, tok in enumerate(toks):
            whole = _lower_name(tok)
            before = _raw(_at(toks, i - 1))
            after = _at(toks, i + 1)
            if (
                whole
                and (before == "(" or before == ",")
                and (after is None or after.raw_text == ")" or after.raw_text == ",")
            ):
                ended.pop(whole, None)


def _end_from(
    owner: str,
    shown: str,
    how: str,
    ended: dict[str, _Ended],
    ranges_of: Mapping[str, _Derived],
) -> None:
    """Ends what was taken from `owner`, and from that in turn, depth first as
    upstream's recursive endFrom does."""
    stack: list[Iterator[tuple[str, _Derived]]] = [iter(list(ranges_of.items()))]
    owners = [owner]
    while stack:
        for taken, derived in stack[-1]:
            if derived.owner == owners[-1] and taken not in ended:
                ended[taken] = _Ended(
                    f"{derived.kind} of '{shown}', which was {how}", derived.error
                )
                stack.append(iter(list(ranges_of.items())))
                owners.append(taken)
                break
        else:
            stack.pop()
            owners.pop()


def _member_chain(toks: Sequence[VbaToken], from_: int) -> list[str]:
    """The member names of the chain starting at `.` at `from_`, arguments skipped:
    `.Sheets(1).Range("A1")` gives ['sheets', 'range']."""
    out: list[str] = []
    i = from_
    while _raw(_at(toks, i)) == "." and token_name(_at(toks, i + 1)):
        out.append(token_text(toks[i + 1]))
        i += 2
        if _raw(_at(toks, i)) == "(":
            close = match_paren_from(toks, i)
            if close < 0:
                return []
            i = close + 1
    return out if i == len(toks) else []


def _report(toks: Sequence[VbaToken], start: int, ended: dict[str, _Ended], push: PushFn) -> None:
    for i, tok in enumerate(toks):
        lower = _lower_name(tok)
        before = _raw(_at(toks, i - 1))
        if not lower or before == "." or before == "!":
            continue
        at = Span(start + tok.start, start + tok.end)
        gone = ended.get(lower)
        if gone is not None and _raw(_at(toks, i + 1)) == "." and token_name(_at(toks, i + 2)):
            name = toks[i + 2].raw_text
            verb = "is" if _DERIVED_HOW_RE.search(gone.how) else "was"
            what = f"'{tok.raw_text}' {verb} {gone.how}"
            push(
                "objectUsedAfterDelete",
                f"{what}, so its {name} is gone. This will raise Run-time error "
                f"{gone.error.replace('MEMBER', name, 1)}.",
                at,
            )
            del ended[lower]
