"""What a module's code may do to its workbook's sheets while it runs.

Ported from xlide_vscode/src/analyzer/symbols/sheetChanges.ts (XLIDE issue
#229). The sheets a workbook was saved with say which names and indexes
``ThisWorkbook.Sheets(...)`` can reach, but only until code adds, copies or
renames one. This scan finds those operations in the text, so the check judges
a sheet only when no code in the project could have made it.

It errs towards finding too much: a With block's ``.Add`` counts as adding a
sheet whatever the With names, and so does a ``.Copy`` given an argument that
no range word explains.
"""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from typing import Literal

from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize_cached


@dataclass(frozen=True, slots=True)
class WorkbookSheetInfo:
    """A sheet of the saved workbook, as the host read it from the file."""

    name: str
    kind: Literal["worksheet", "chartsheet", "dialogsheet", "macrosheet"]


@dataclass(frozen=True, slots=True)
class SheetChanges:
    # Code adds or copies a sheet, so the count and the default names can grow.
    adds_sheets: bool = False
    # Lowercased names that code assigns to some `.Name` as a literal.
    names_assigned: frozenset[str] = field(default_factory=frozenset)
    # Code assigns some `.Name` a value that is not a literal, so any name can appear.
    assigns_computed_name: bool = False


_SHEET_COLLECTIONS = frozenset(("sheets", "worksheets", "charts"))
_RANGE_WORDS = frozenset(
    (
        "range", "cells", "rows", "columns", "usedrange", "selection", "destination",
        "currentregion", "entirerow", "entirecolumn", "offset", "resize",
    )
)
# Tokens after which a member chain begins a statement rather than an expression.
_STATEMENT_OPENERS = frozenset(("then", "else"))


def sheet_changes_in(source: str) -> SheetChanges:
    """The sheet operations in one module's code."""
    names: set[str] = set()
    adds_sheets = False
    assigns_computed_name = False
    statement: list[VbaToken] = []

    def flush() -> None:
        nonlocal adds_sheets, assigns_computed_name
        found_adds, found_computed = _scan_statement(statement, names)
        adds_sheets = adds_sheets or found_adds
        assigns_computed_name = assigns_computed_name or found_computed
        statement.clear()

    for token in tokenize_cached(source):
        if token.kind is TokenKind.COMMENT:
            continue
        if token.kind is TokenKind.NEWLINE or token.kind is TokenKind.COLON:
            flush()
            continue
        statement.append(token)
    flush()
    return SheetChanges(
        adds_sheets=adds_sheets,
        names_assigned=frozenset(names),
        assigns_computed_name=assigns_computed_name,
    )


def merge_sheet_changes(all_changes: Iterable[SheetChanges]) -> SheetChanges:
    """The union of several modules' sheet operations."""
    names: set[str] = set()
    adds_sheets = False
    assigns_computed_name = False
    for changes in all_changes:
        adds_sheets = adds_sheets or changes.adds_sheets
        assigns_computed_name = assigns_computed_name or changes.assigns_computed_name
        names.update(changes.names_assigned)
    return SheetChanges(
        adds_sheets=adds_sheets,
        names_assigned=frozenset(names),
        assigns_computed_name=assigns_computed_name,
    )


def _scan_statement(toks: Sequence[VbaToken], names: set[str]) -> tuple[bool, bool]:
    adds_sheets = False
    assigns_computed_name = False

    def lower(i: int) -> str:
        return toks[i].raw_text.lower() if 0 <= i < len(toks) else ""

    for i, tok in enumerate(toks):
        if tok.raw_text != ".":
            continue
        member = lower(i + 1)
        if member in ("add", "add2"):
            # `Worksheets.Add`, or `.Add` inside a With whose object is not in view.
            before = toks[i - 1] if i > 0 else None
            if (
                before is None
                or before.kind is TokenKind.KEYWORD
                or before.raw_text.lower() in _SHEET_COLLECTIONS
            ):
                adds_sheets = True
        elif (
            member == "copy"
            and i + 2 < len(toks)
            and not any(t.raw_text.lower() in _RANGE_WORDS for t in toks)
        ):
            # `ws.Copy After:=...` makes a sheet; `ws.Copy` alone makes a workbook.
            adds_sheets = True
        elif member == "name" and i + 2 < len(toks) and toks[i + 2].raw_text == "=":
            rhs = toks[i + 3 :]
            # A lone literal, maybe followed by a single-line If's Else.
            if (
                rhs
                and rhs[0].kind is TokenKind.STRING_LITERAL
                and (len(rhs) == 1 or rhs[1].raw_text.lower() in _STATEMENT_OPENERS)
            ):
                names.add(rhs[0].raw_text[1:-1].replace('""', '"').lower())
            elif _starts_statement(toks, i):
                assigns_computed_name = True
    return adds_sheets, assigns_computed_name


def _starts_statement(toks: Sequence[VbaToken], dot: int) -> bool:
    """True when the member chain ending at the ``.`` at ``dot`` starts the
    statement, as an assignment target does, rather than sitting inside a
    condition such as ``If ws.Name = s Then``."""
    i = dot - 1
    while i >= 0:
        text = toks[i].raw_text
        if text == ")":
            depth = 0
            while i >= 0:
                if toks[i].raw_text == ")":
                    depth += 1
                if toks[i].raw_text == "(":
                    depth -= 1
                    if depth == 0:
                        break
                i -= 1
            i -= 1
            continue
        # `Me.Name = s` in a sheet's own module renames that sheet.
        if (
            toks[i].kind is TokenKind.IDENTIFIER
            or toks[i].kind is TokenKind.BRACKETED_IDENTIFIER
            or text == "."
            or text == "!"
            or text.lower() == "me"
        ):
            i -= 1
            continue
        break
    return i < 0 or toks[i].raw_text.lower() in _STATEMENT_OPENERS
