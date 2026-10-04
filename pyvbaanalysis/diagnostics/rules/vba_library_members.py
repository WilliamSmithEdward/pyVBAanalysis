"""Rule: names after `VBA.` that the VBA library does not have (XLIDE issue #369),
and after `Excel.` in an Excel project (issue #305). Measured in Excel 16.0,
compiled with Debug > Compile:

 - `VBA.Nosuch`, `VBA.Strings.Nosuch`, `VBA.VbMsgBoxResult.vbNosuch` and
   `VBA.Global.Left$(...)` are "Method or data member not found": the library,
   its modules, its enums and its hidden Global class are closed (see
   runtime/vba_library_names.py, generated from VBE7.DLL).
 - `VBA.Err.LastDllError = 5` is "Can't assign to read-only property", as the
   unqualified `Err.LastDllError = 5` is. Err's other members are not judged:
   `VBA.Err.Nosuch` compiles.

A project that declares a name VBA of its own is not judged.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/vbaLibraryMembers.ts.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from typing import AbstractSet

from ...conditional import ConditionalActivityTracker
from ...host.excel_library_names import EXCEL_LIBRARY_NAMES
from ...lexer.token_kinds import VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...runtime.vba_library_names import VBA_ERR_READ_ONLY, VBA_LIBRARY_CONTAINERS, VBA_LIBRARY_NAMES
from ...symbols.symbol_model import ModuleSymbols, VbaSymbol
from ..context import PushFn
from ..walker import (
    active_module_members,
    first_executable_token_index,
    for_each_statement,
    statement_and_branch_spans,
    statement_tokens,
    token_name,
    token_text,
)


def check_vba_library_members(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    host_name: str | None = None,
) -> None:
    def shadows(lower: str) -> bool:
        return any(
            symbol.name.lower() == lower or any(child.name.lower() == lower for child in (symbol.children or []))
            for symbol in (symbols.root.children or [])
        ) or any(
            symbol.name.lower() == lower or symbol.module_name.lower() == lower
            for symbol in (project_visible_symbols or [])
        )

    if host_name is None or host_name == "Excel":
        if not shadows("excel"):
            _check_excel_library_names(source, mod, activity, push)
    if shadows("vba"):
        return
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue

        def visit(stmt: LeafStatementNode) -> None:
            for span in statement_and_branch_spans(stmt):
                toks = statement_tokens(source, span)

                def at(k: int, toks: Sequence[VbaToken] = toks, span: Span = span) -> Span:
                    return Span(span.start + toks[k].start, span.start + toks[k].end)

                for i in range(len(toks) - 2):
                    if (
                        token_text(toks[i]) != "vba"
                        or (i >= 1 and toks[i - 1].raw_text == ".")
                        or toks[i + 1].raw_text != "."
                    ):
                        continue
                    first = _name_at(toks, i + 2)
                    if first is None:
                        continue
                    if first.lower not in VBA_LIBRARY_NAMES:
                        push("memberNotFound", _missing(first, "the VBA library", VBA_LIBRARY_NAMES), at(i + 2))
                        continue
                    if _raw_text_at(toks, first.next) != ".":
                        continue
                    second = _name_at(toks, first.next + 1)
                    if second is None:
                        continue
                    container = VBA_LIBRARY_CONTAINERS.get(first.lower)
                    if container is not None and second.lower not in container:
                        push(
                            "memberNotFound",
                            _missing(second, f"VBA.{first.shown}", container),
                            at(first.next + 1),
                        )
                        continue
                    # `VBA.Err.LastDllError = 5`, the assignment's target.
                    head = first_executable_token_index(toks)
                    starts_target = i == head or (
                        i == head + 1 and token_text(_token_at(toks, head)) in ("let", "set")
                    )
                    if (
                        first.lower == "err"
                        and starts_target
                        and second.lower in VBA_ERR_READ_ONLY
                        and _raw_text_at(toks, second.next) == "="
                    ):
                        push(
                            "readonlyMemberAssignment",
                            f"Cannot assign to read-only property 'VBA.Err.{second.shown}'. This is a VBE "
                            "compile error: Can't assign to read-only property.",
                            at(first.next + 1),
                        )

        for_each_statement(member.body, visit, activity)


def _check_excel_library_names(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """`Excel.Nope`, `Excel.Version`: a name after `Excel.` that is no type, enum
    constant or Global member of the library (XLIDE issue #305, measured in Excel
    16.0). A name after As or New is a type, which the VBE refuses otherwise, and
    a member of what `Excel.X` gives is not judged here."""
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue

        def visit(stmt: LeafStatementNode) -> None:
            for span in statement_and_branch_spans(stmt):
                toks = statement_tokens(source, span)
                for i in range(len(toks) - 2):
                    before = _token_at(toks, i - 1)
                    if (
                        token_text(toks[i]) != "excel"
                        or (before is not None and before.raw_text == ".")
                        or toks[i + 1].raw_text != "."
                        or token_text(before) in ("as", "new")
                    ):
                        continue
                    name = token_name(toks[i + 2])
                    if name and name.lower() not in EXCEL_LIBRARY_NAMES:
                        push(
                            "memberNotFound",
                            f"'{name}' is not a member of the Excel library. This is a VBE compile error: Method or "
                            "data member not found.",
                            Span(span.start + toks[i + 2].start, span.start + toks[i + 2].end),
                        )

        for_each_statement(member.body, visit, activity)


@dataclass(frozen=True, slots=True)
class _NameAt:
    # The name, lowercased, with a `$` glued to it.
    lower: str
    shown: str
    # The index after it.
    next: int


def _missing(name: _NameAt, where: str, names: AbstractSet[str]) -> str:
    """Why a name is refused: none of that name, or a `$` form of a name that has
    none, `VBA.Asc$` (measured in Excel 16.0)."""
    if name.lower.endswith("$") and name.lower[:-1] in names:
        return (
            f"'{name.shown}' has no String form with $ in {where}. This is a VBE compile error: "
            "Type-declaration character does not match declared data type."
        )
    return f"'{name.shown}' is not a member of {where}. This is a VBE compile error: Method or data member not found."


def _name_at(toks: Sequence[VbaToken], k: int) -> _NameAt | None:
    """The name at `k`, lowercased, with a `$` glued to it, and the index after it."""
    name = token_name(_token_at(toks, k))
    if not name:
        return None
    following = _token_at(toks, k + 1)
    dollar = following is not None and following.raw_text == "$" and following.start == toks[k].end
    suffix = "$" if dollar else ""
    return _NameAt(name.lower() + suffix, toks[k].raw_text + suffix, k + (2 if dollar else 1))


def _raw_text_at(toks: Sequence[VbaToken], index: int) -> str | None:
    return toks[index].raw_text if 0 <= index < len(toks) else None


def _token_at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    return toks[index] if 0 <= index < len(toks) else None
