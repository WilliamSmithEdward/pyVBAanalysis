"""Rule family: Declare statements (XLIDE issue #254).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/declares.ts. Every case
was measured in Excel 16.0 (build 20326, 2026-10-01).

- invalid-proc-header: what the VBE refuses in a Declare's own line. A
  Declare Sub given an As clause, a Declare Function given both a type
  character and an As clause, or one returning `String * n`, is
  "Expected: end of statement"; one returning `As Any` is "Expected: type
  name"; a Lib or Alias that is not a string literal, a Const included, is
  "Expected: string constant". A parameter `As String * n`, of a Declare
  or of a procedure, is "Expected array".
- unusable-declare: a Declare that compiles and fails on every call. CDecl
  raises 49, Bad DLL calling convention; `Lib ""` or a Lib of spaces 48,
  File not found; an Alias of nothing or spaces 453, and `Alias "#0"` 452,
  Can't find DLL entry point. A Declare never called runs, so each call is
  reported, not the Declare. A real library or entry point the machine
  lacks (53, 453) depends on the machine and is left alone.
"""

from __future__ import annotations

import re
from collections.abc import Sequence

from ...conditional import ConditionalActivityTracker
from ...js_compat import js_trim
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    DeclareNode,
    LeafStatementNode,
    ModuleNode,
    ParameterNode,
    ProcedureNode,
    Span,
)
from ...symbols.symbol_model import ModuleSymbols
from ...types.type_inference import procedure_symbol_for
from ..call_extraction import string_literal_value
from ..context import PushFn
from ..walker import (
    active_module_members,
    for_each_statement,
    match_paren_from,
    statement_tokens,
    token_name,
    token_text,
)

_ZERO_ORDINAL_RE = re.compile(r"#0+")


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) outside the list."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    """`toks[i]?.rawText`."""
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


def check_declare_statements(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    for member in active_module_members(mod, activity):
        if isinstance(member, DeclareNode):
            _check_declare_line(source, member, push)
        if isinstance(member, (DeclareNode, ProcedureNode)):
            for param in member.params:
                _check_fixed_length_parameter(source, param, push)


def _check_declare_line(source: str, declare: DeclareNode, push: PushFn) -> None:
    toks = statement_tokens(source, declare.span)

    def refuse(tok: VbaToken, what: str, error: str) -> None:
        push(
            "invalidProcedureHeader",
            f"Declare '{declare.name}': {what}. This is a VBE compile error: {error}.",
            Span(declare.span.start + tok.start, declare.span.start + tok.end),
        )

    open_ = -1
    for i, tok in enumerate(toks):
        word = token_text(tok)
        following = _at(toks, i + 1)
        if (
            (word == "lib" or word == "alias")
            and following is not None
            and following.kind is not TokenKind.STRING_LITERAL
        ):
            refuse(
                following,
                f"{tok.raw_text} takes a string literal, and '{following.raw_text}' is none",
                "Expected: string constant",
            )
            return
        if tok.raw_text == "(" and open_ < 0:
            open_ = i
    close = -1 if open_ < 0 else match_paren_from(toks, open_)
    as_ = None if close < 0 else _at(toks, close + 1)
    if as_ is None or token_text(as_) != "as":
        return
    if not declare.is_function:
        refuse(as_, "a Sub returns nothing, so it takes no As clause", "Expected: end of statement")
    elif declare.type_suffix:
        refuse(
            as_,
            f"the type character '{declare.type_suffix}' already gives the return type",
            "Expected: end of statement",
        )
    elif token_text(_at(toks, close + 2)) == "any":
        refuse(
            toks[close + 2], "As Any is for a parameter, not a return type", "Expected: type name"
        )
    elif _raw_at(toks, close + 3) == "*":
        refuse(
            toks[close + 3],
            "a Declare Function cannot return a fixed-length String",
            "Expected: end of statement",
        )


def _check_fixed_length_parameter(source: str, param: ParameterNode, push: PushFn) -> None:
    """`ByVal s As String * 4`: no parameter takes a fixed-length String."""
    toks = statement_tokens(source, param.span)
    star = next(
        (
            i
            for i, tok in enumerate(toks)
            if tok.raw_text == "*"
            and token_text(_at(toks, i - 1)) == "string"
            and token_text(_at(toks, i - 2)) == "as"
        ),
        -1,
    )
    if star < 0:
        return
    length = _at(toks, star + 1)
    push(
        "invalidProcedureHeader",
        f"Parameter '{param.name}' is declared As String * "
        f"{length.raw_text if length is not None else 'n'}, and no parameter takes a "
        "fixed-length String. This is a VBE compile error: Expected array.",
        Span(
            param.span.start + toks[star - 1].start,
            param.span.start + (length if length is not None else toks[star]).end,
        ),
    )


def _failure(source: str, declare: DeclareNode) -> str | None:
    """Why every call of a Declare fails, or None."""
    toks = statement_tokens(source, declare.span)

    def value_after(word: str) -> str | None:
        i = next((k for k, tok in enumerate(toks) if token_text(tok) == word), -1)
        following = _at(toks, i + 1) if i >= 0 else None
        if following is not None and following.kind is TokenKind.STRING_LITERAL:
            return string_literal_value(following.raw_text)
        return None

    if any(token_text(tok) == "cdecl" for tok in toks):
        return (
            "is declared CDecl, which Office on Windows refuses: every call to it raises "
            "Run-time error '49': Bad DLL calling convention"
        )
    lib = value_after("lib")
    if lib is not None and js_trim(lib) == "":
        return (
            f"names no library, Lib \"{lib}\": every call to it raises Run-time error '48': "
            "File not found"
        )
    alias = value_after("alias")
    if alias is not None and js_trim(alias) == "":
        return (
            f'names no entry point, Alias "{alias}": every call to it raises Run-time error '
            "'453': Can't find DLL entry point"
        )
    if alias is not None and _ZERO_ORDINAL_RE.fullmatch(alias):
        return (
            f"names entry point {alias}, an ordinal no library has: every call to it raises "
            "Run-time error '452': Can't find DLL entry point 0"
        )
    return None


def check_unusable_declare_calls(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    failing: dict[str, str] = {}
    for member in active_module_members(mod, activity):
        why = _failure(source, member) if isinstance(member, DeclareNode) else None
        if why and isinstance(member, DeclareNode):
            failing[member.name.lower()] = why
    if len(failing) == 0:
        return
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        # A parameter or local of the same name hides the Declare.
        proc_sym = procedure_symbol_for(symbols, member)
        hidden = {
            child.name.lower()
            for child in (proc_sym.children if proc_sym is not None else None) or []
        }

        # A single-line If's statement holds its branches.
        def visit(stmt: LeafStatementNode, hidden: set[str] = hidden) -> None:
            toks = statement_tokens(source, stmt.span)
            for i, tok in enumerate(toks):
                name = token_name(tok)
                lower = name.lower() if name is not None else None
                why = failing.get(lower) if lower and lower not in hidden else None
                if not why or _raw_at(toks, i - 1) == "." or _raw_at(toks, i - 1) == "!":
                    continue
                push(
                    "unusableDeclare",
                    f"'{tok.raw_text}' {why}.",
                    Span(stmt.span.start + tok.start, stmt.span.start + tok.end),
                )

        for_each_statement(member.body, visit, activity)
