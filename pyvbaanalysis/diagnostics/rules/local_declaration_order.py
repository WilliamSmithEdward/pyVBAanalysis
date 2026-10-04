"""Rule: a procedure's own Dim, Static or Const covers only the lines after it.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/localDeclarationOrder.ts.
Measured by full compile in 64-bit Excel 16.0 (build 20326, 2026-09-30):

- Under Option Explicit, a name used above its local declaration, with no
  module or project declaration of that name, is "Variable not defined":
  `Debug.Print K` then `Const K = 1`, `x = 1` then `Dim x`, `Const A = B
  + 1` then `Const B = 1`, `Dim a(N)` then `Const N = 5`, a use inside an
  earlier loop. VBA does not hoist a declaration.
- Where the earlier use found something else, a module-level declaration
  of that name or, without Option Explicit, an implicit variable, the
  local declaration is "Duplicate declaration in current scope". Without
  Option Explicit, an earlier use inside a Const's value is "Constant
  expression required" instead, since an implicit variable is no constant.
- Scope is the whole procedure: a Const inside an If block, used after
  the block, compiles.
"""

from __future__ import annotations

import math
import re
from collections.abc import Sequence
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...host.host_model import HostObjectModel, resolve_host_global
from ...js_compat import js_trim
from ...lexer.token_helpers import first_token_at_or_after
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import ModuleNode, OptionNode, ProcedureNode, Span, VariableGroupNode
from ...runtime import resolve_runtime_function
from ...symbols.name_resolution import BareIdentifierContext, BareIdentifierResolutionScope
from ...symbols.symbol_model import ModuleSymbols, VbaSymbol
from ...types.type_inference import source_identifier_binding
from ..context import PushFn
from ..walker import (
    active_module_members,
    for_each_variable_group,
    is_inactive_node,
    token_name,
    token_text,
)


@dataclass(frozen=True, slots=True)
class _LocalDeclaration:
    name: str
    name_span: Span
    # Where the declaring statement begins; a use before it is too early.
    start: int
    is_const: bool


# Words after which a name is a label, a type or an object, never a value read.
_NOT_A_VALUE_AFTER = frozenset({"as", "new", "goto", "gosub", "resume"})

_EXPLICIT_RE = re.compile(r"^explicit\b", re.IGNORECASE | re.ASCII)
_LINE_BREAK_RE = re.compile(r"\r|\n")


def check_local_declaration_order(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    host_model: HostObjectModel | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    option_explicit = any(
        isinstance(member, OptionNode) and _EXPLICIT_RE.search(js_trim(member.option_text))
        for member in active_module_members(mod, activity)
    )
    tokens: Sequence[VbaToken] | None = None
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        declarations: dict[str, _LocalDeclaration] = {}
        last_declaration = -math.inf

        def visit(group: VariableGroupNode) -> None:
            nonlocal last_declaration
            for decl in group.declarations:
                key = decl.name.lower()
                if not is_inactive_node(activity, decl) and key not in declarations:
                    last_declaration = max(last_declaration, group.span.start)
                    declarations[key] = _LocalDeclaration(
                        name=decl.name,
                        name_span=decl.name_span if decl.name_span is not None else decl.span,
                        start=group.span.start,
                        is_const=group.is_const,
                    )

        for_each_variable_group(member.body, visit, activity)
        if len(declarations) == 0:
            continue
        params = {param.name.lower() for param in member.params}
        header_end = source.find("\n", member.span.start)
        body_start = member.span.end if header_end < 0 else header_end
        if tokens is None:
            tokens = tokenize_cached(source)
        toks = tokens
        reported: set[str] = set()
        const_spans: list[Span] | None = None
        const_cursor = 0

        def in_const_value(offset: int) -> bool:
            nonlocal const_spans, const_cursor
            if const_spans is None:
                const_spans = []
                seen: set[int] = set()
                for declaration in declarations.values():
                    if not declaration.is_const or declaration.start in seen:
                        continue
                    seen.add(declaration.start)
                    const_spans.append(
                        Span(declaration.start, _const_statement_end(source, declaration.start))
                    )
                const_spans.sort(key=lambda span: span.start)
            # Uses arrive in token order; physical Const line ends never decrease.
            while const_cursor < len(const_spans) and const_spans[const_cursor].start <= offset:
                const_cursor += 1
            return const_cursor > 0 and const_spans[const_cursor - 1].end > offset

        i = first_token_at_or_after(toks, body_start)
        while i < len(toks) and toks[i].start < last_declaration:
            tok = toks[i]
            i += 1
            name = token_name(tok) if tok.kind is TokenKind.IDENTIFIER else None
            key = name.lower() if name is not None else None
            declaration = declarations.get(key) if key else None
            if (
                not key
                or declaration is None
                or tok.start >= declaration.start
                or key in params
                or key in reported
            ):
                continue
            if not _is_value_use(toks, i - 1) or (
                activity is not None and activity.is_inactive(Span(tok.start, tok.end))
            ):
                continue
            # A name VBA or the host already gives a meaning is not measured here.
            if resolve_runtime_function(declaration.name) is not None or (
                host_model is not None and resolve_host_global(declaration.name, host_model)
            ):
                continue
            reported.add(key)
            outer = source_identifier_binding(
                symbols,
                None,
                project_visible_symbols,
                declaration.name,
                BareIdentifierContext.EXPRESSION,
            )
            unresolved = outer.scope is BareIdentifierResolutionScope.UNRESOLVED
            if option_explicit and unresolved:
                what = "Const" if declaration.is_const else "declaration"
                push(
                    "undeclaredVariable",
                    f"Variable not defined: '{tok.raw_text}'. It is declared further down the "
                    "procedure, and a declaration covers only the lines after it; move the "
                    f"{what} above this line.",
                    Span(tok.start, tok.end),
                )
            elif not option_explicit and unresolved and in_const_value(tok.start):
                push(
                    "constValueNotConstant",
                    f"'{tok.raw_text}' is declared further down the procedure, so here it is an "
                    "implicit variable, which a Const cannot take its value from. This is a VBE "
                    "compile error: Constant expression required.",
                    Span(tok.start, tok.end),
                )
            else:
                where = (
                    "became an implicit variable"
                    if unresolved
                    else "named the module's declaration"
                )
                push(
                    "duplicateDeclaration",
                    f"'{declaration.name}' is used above this declaration, where it {where}. "
                    "This is a VBE compile error: Duplicate declaration in current scope.",
                    declaration.name_span,
                )


def _is_value_use(tokens: Sequence[VbaToken], i: int) -> bool:
    """Whether the name at `i` reads a variable, rather than naming a member, a
    label, a type or an argument."""
    before = tokens[i - 1] if i - 1 >= 0 else None
    after = tokens[i + 1] if i + 1 < len(tokens) else None
    if before is not None and (
        before.raw_text == "." or before.raw_text == "!" or token_text(before) in _NOT_A_VALUE_AFTER
    ):
        return False
    if after is not None and after.raw_text == ":=":
        return False
    # `Label:` at the start of a line.
    at_line_start = before is None or before.kind is TokenKind.NEWLINE
    return not (at_line_start and after is not None and after.kind is TokenKind.COLON)


def _const_statement_end(source: str, start: int) -> int:
    """Where the Const statement starting at `start` ends: its line's end."""
    found = _LINE_BREAK_RE.search(source, start)
    return len(source) if found is None else found.start()
