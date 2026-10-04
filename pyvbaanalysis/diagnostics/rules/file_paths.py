"""Rule: a file statement or function given an empty path (XLIDE issue #262).

Each was measured in Excel 16.0 (build 20326, 2026-10-01): `Open ""` raises 75,
`FileLen("")` 53 and `MkDir ""` 76, and some take a path of spaces the same way
while Dir("") runs and Dir(" ") raises 53. Only the pairs measured are reported.
The path is a string literal, or a String local or Const the statement is known
to see holding one: a String never assigned holds "".

Ported from xlide_vscode/src/analyzer/diagnostics/rules/filePaths.ts.
"""

from __future__ import annotations

import json
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ProcedureNode, Span
from ...symbols.symbol_model import ModuleSymbols, VbaSymbol
from ...types.type_inference import known_local_literal_values_at, string_constants_in_scope
from ..call_extraction import string_literal_value
from ..callable_signatures import SourceNameScope, runtime_callable_source_shadowed, source_name_scope_for
from ..context import PushFn
from ..known_locals import KnownLocalValue
from ..walker import (
    ProcedureStatementVisitor,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .shared import is_bare_or_vba_qualified_intrinsic_call


@dataclass(frozen=True, slots=True)
class _PathErrors:
    """The error a statement or function raises for "" and for a path of spaces, where measured."""

    display: str
    empty: int | None = None
    blank: int | None = None


_ERRORS: dict[str, _PathErrors] = {
    "open": _PathErrors("Open", empty=75, blank=53),
    "mkdir": _PathErrors("MkDir", empty=76, blank=76),
    "chdir": _PathErrors("ChDir", empty=76, blank=76),
    "rmdir": _PathErrors("RmDir", empty=76),
    "kill": _PathErrors("Kill", empty=53),
    "setattr": _PathErrors("SetAttr", empty=53),
    "filecopy": _PathErrors("FileCopy", empty=75),
    "name": _PathErrors("Name", empty=75),
    "filelen": _PathErrors("FileLen", empty=53, blank=53),
    "filedatetime": _PathErrors("FileDateTime", empty=53),
    "getattr": _PathErrors("GetAttr", empty=53),
    "dir": _PathErrors("Dir", blank=53),
}

_FUNCTIONS = frozenset({"filelen", "filedatetime", "getattr", "dir"})

_ERROR_TEXT: dict[int, str] = {53: "File not found", 75: "Path/File access error", 76: "Path not found"}

_SPACES = re.compile(" +")

_OPEN_CLAUSE_WORDS = frozenset({"for", "access", "shared", "lock", "as"})


def check_empty_file_paths(
    source: str,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        # Built on first use, as upstream's `??=` builds them.
        values_at: list[Callable[[LeafStatementNode], Mapping[str, KnownLocalValue]]] = []
        consts: list[Mapping[str, str]] = []
        source_names: list[SourceNameScope] = []

        def visitor(stmt: LeafStatementNode) -> None:
            def string_of(tok: VbaToken) -> tuple[str, bool] | None:
                """The text a path token holds, and whether a local or Const holds it."""
                if tok.kind is TokenKind.STRING_LITERAL:
                    return (string_literal_value(tok.raw_text), False)
                name = token_name(tok)
                lower = name.lower() if name is not None else None
                if not lower:
                    return None
                if not values_at:
                    values_at.append(known_local_literal_values_at(source, member, symbols, activity))
                local = values_at[0](stmt).get(lower)
                if local is not None:
                    if local.kind == "string" and not local.content_mutated and isinstance(local.value, str):
                        return (local.value, True)
                    return None
                if not consts:
                    consts.append(string_constants_in_scope(symbols, member))
                constant = consts[0].get(lower)
                return None if constant is None else (constant, True)

            def check(span_start: int, which: str, path: Sequence[VbaToken]) -> None:
                errors = _ERRORS[which]
                if len(path) != 1:
                    return
                known = string_of(path[0])
                if known is None:
                    return
                value, held = known
                kind = "empty" if value == "" else "blank" if _SPACES.fullmatch(value) is not None else None
                error = (errors.empty if kind == "empty" else errors.blank) if kind is not None else None
                if kind is None or error is None:
                    return
                if held:
                    given = f"'{path[0].raw_text}', which holds {json.dumps(value, ensure_ascii=False)} here"
                else:
                    given = "an empty path" if kind == "empty" else "a path of spaces"
                push(
                    "emptyFilePath",
                    f"{errors.display} is given {given}. This will raise Run-time error '{error}': "
                    f"{_ERROR_TEXT[error]}.",
                    Span(span_start + path[0].start, span_start + path[0].end),
                )

            for span in statement_and_branch_spans(stmt):
                toks = statement_tokens_after_leading_label(source, span)
                head = token_text(toks[0] if toks else None)
                for path in _statement_paths(head, toks):
                    check(span.start, head, path)
                for i in range(len(toks) - 1):
                    name = token_text(toks[i])
                    if (
                        name not in _FUNCTIONS
                        or toks[i + 1].raw_text != "("
                        or not is_bare_or_vba_qualified_intrinsic_call(toks, i)
                    ):
                        continue
                    if not (i >= 1 and toks[i - 1].raw_text == "."):
                        if not source_names:
                            source_names.append(source_name_scope_for(symbols, member, project_visible_symbols))
                        if runtime_callable_source_shadowed(name, source_names[0]):
                            continue
                    close = match_paren_from(toks, i + 1)
                    args = split_top_level_token_groups(toks[i + 2 : close], 0, ",") if close > 0 else []
                    if len(args) >= 1 and (name == "dir" or len(args) == 1):
                        check(span.start, name, args[0])

        return visitor

    return factory


def _statement_paths(head: str, toks: Sequence[VbaToken]) -> list[list[VbaToken]]:
    """The path arguments a file statement names."""
    second = toks[1].raw_text if len(toks) > 1 else None
    if head == "open":
        end = next((k for k, tok in enumerate(toks) if k > 0 and token_text(tok) in _OPEN_CLAUSE_WORDS), -1)
        return [list(toks[1:end])] if end > 1 else []
    if head in ("mkdir", "chdir", "rmdir", "kill"):
        return [] if second == "=" else [list(toks[1:])]
    if head in ("setattr", "filecopy"):
        if second == "=":
            return []
        args = split_top_level_token_groups(toks[1:], 0, ",")
        if len(args) != 2:
            return []
        return [args[0]] if head == "setattr" else args
    if head == "name":
        as_index = next((k for k, tok in enumerate(toks) if token_text(tok) == "as"), -1)
        return [list(toks[1:as_index])] if as_index > 1 and second != "=" else []
    return []
