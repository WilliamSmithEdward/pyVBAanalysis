"""Rule family: expression-syntax rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/expressions.ts: unbalanced
parentheses and division by a provably-zero divisor. The remaining rules in the
family (invalid-expression-syntax and the parenthesized/parenless call-shape
rules) need the member-completion context and runtime-function surfaces, so they
stay deferred until those host layers land.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence

from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import (
    IntegerConstantLookup,
    resolve_raw_integer_constants,
)
from ...lexer.token_helpers import match_paren_from
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...symbols.symbol_model import ModuleSymbols, VbaSymbol
from ...types.type_inference import procedure_symbol_for
from ..callable_signatures import scoped_integer_constant_lookup
from ..const_expr import (
    collect_body_literal_integer_constants,
    collect_module_literal_integer_constants,
    fold_integer_expression_tokens,
)
from ..context import PushFn, statement_tokens
from ..walker import ProcedureStatementVisitor, token_name, token_text


def check_unbalanced_parens(source: str, push: PushFn) -> None:
    """Every parenthesis must be matched within its logical statement (a `(` left
    open at a statement boundary, or a stray `)`, is a VBE Syntax error)."""
    toks = tokenize_cached(source)
    depth = 0
    open_offsets: list[int] = []
    flagged = False

    def flush() -> None:
        nonlocal depth, flagged
        if not flagged and depth > 0:
            off = open_offsets[0]
            push("unbalancedParens", "Unbalanced parentheses: a ')' is missing.", Span(off, off + 1))
        depth = 0
        open_offsets.clear()
        flagged = False

    for tok in toks:
        if tok.kind is TokenKind.NEWLINE:
            flush()
            continue
        if tok.kind is TokenKind.COLON and depth == 0:
            flush()
            continue
        if tok.kind is not TokenKind.PUNCTUATION:
            continue
        if tok.raw_text == "(":
            depth += 1
            open_offsets.append(tok.start)
        elif tok.raw_text == ")":
            if depth == 0:
                if not flagged:
                    push(
                        "unbalancedParens",
                        "Unbalanced parentheses: an unexpected ')' was found.",
                        Span(tok.start, tok.end),
                    )
                    flagged = True
            else:
                depth -= 1
                open_offsets.pop()
    flush()


_TYPE_SUFFIX = re.compile(r"[!#@%&^]$")
_D_EXPONENT = re.compile(r"[dD]")
_HEX = re.compile(r"^&[hH]([0-9A-Fa-f]+)$")
_OCTAL = re.compile(r"^&[oO]([0-7]+)$")
_FLOAT = re.compile(r"^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")


def check_division_by_zero_expressions(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_integer_constants: Mapping[str, str | None] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """`/`, `\\`, or `Mod` against a provably-zero divisor raises Run-time error 11."""
    project_constants = resolve_raw_integer_constants(project_integer_constants or {}, {})
    module_constants = collect_module_literal_integer_constants(mod, activity, project_constants)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        procedure_constants = dict(module_constants)
        collect_body_literal_integer_constants(member.body, procedure_constants, activity)
        proc_sym = procedure_symbol_for(symbols, member)
        constants = scoped_integer_constant_lookup(
            procedure_constants, symbols, proc_sym, project_visible_symbols
        )

        def visitor(stmt: LeafStatementNode) -> None:
            for operator, span in _division_by_zero_divisors(source, stmt.span, constants):
                push(
                    "divisionByZero",
                    f"Expression uses '{operator}' with a zero divisor. This will raise "
                    "Run-time error '11': Division by zero.",
                    span,
                )

        return visitor

    return factory


def _division_by_zero_divisors(
    source: str, span: Span, constants: IntegerConstantLookup
) -> list[tuple[str, Span]]:
    toks = statement_tokens(source, span)
    hits: list[tuple[str, Span]] = []
    for i, tok in enumerate(toks):
        operator = _division_by_zero_operator_label(tok)
        if operator is None:
            continue
        divisor = _zero_divisor_token(source, span, toks, i + 1, constants)
        if divisor:
            hits.append((operator, _absolute_token_group_span(span, divisor)))
    return hits


def _division_by_zero_operator_label(tok: VbaToken) -> str | None:
    text = token_text(tok)
    if text in ("/", "\\"):
        return text
    return "Mod" if text == "mod" else None


def _zero_divisor_token(
    source: str,
    span: Span,
    toks: list[VbaToken],
    start: int,
    constants: IntegerConstantLookup,
) -> list[VbaToken] | None:
    if start >= len(toks):
        return None
    first = toks[start]
    if first.raw_text == "(":
        close = match_paren_from(toks, start)
        if close < 0:
            return None
        return _zero_divisor_expression(source, span, toks, start + 1, close, constants)
    if first.kind is TokenKind.OPERATOR and first.raw_text in ("+", "-"):
        signed = _zero_divisor_atom_token_group(toks, start + 1, constants)
        return [first, *signed] if signed else None
    return _zero_divisor_atom_token_group(toks, start, constants)


def _zero_divisor_expression(
    source: str,
    span: Span,
    toks: list[VbaToken],
    start: int,
    end_exclusive: int,
    constants: IntegerConstantLookup,
) -> list[VbaToken] | None:
    if start >= end_exclusive:
        return None
    folded = fold_integer_expression_tokens(source, span, toks, start, end_exclusive, constants)
    if folded == 0:
        return toks[start:end_exclusive]
    if toks[start].raw_text == "(":
        close = match_paren_from(toks, start)
        if close == end_exclusive - 1:
            return _zero_divisor_expression(source, span, toks, start + 1, close, constants)
    if (
        end_exclusive == start + 2
        and toks[start].kind is TokenKind.OPERATOR
        and toks[start].raw_text in ("+", "-")
        and _is_zero_divisor_atom(toks[start + 1], constants)
    ):
        return [toks[start], toks[start + 1]]
    if end_exclusive == start + 1 and _is_zero_divisor_atom(toks[start], constants):
        return [toks[start]]
    return None


def _zero_divisor_atom_token_group(
    toks: list[VbaToken], start: int, constants: IntegerConstantLookup
) -> list[VbaToken] | None:
    if start >= len(toks):
        return None
    first = toks[start]
    first_name = token_name(first)
    member = toks[start + 2] if start + 2 < len(toks) else None
    member_name = token_name(member) if member is not None else None
    if (
        first_name
        and start + 1 < len(toks)
        and toks[start + 1].raw_text == "."
        and member is not None
        and member_name
    ):
        return (
            [first, toks[start + 1], member]
            if constants.get(f"{first_name}.{member_name}".lower()) == 0
            else None
        )
    return [first] if _is_zero_divisor_atom(first, constants) else None


def _is_zero_divisor_atom(tok: VbaToken | None, constants: IntegerConstantLookup) -> bool:
    if _is_zero_numeric_literal(tok):
        return True
    name = token_name(tok) if tok is not None else None
    return name is not None and constants.get(name.lower()) == 0


def _is_zero_numeric_literal(tok: VbaToken | None) -> bool:
    if tok is None or tok.kind not in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL):
        return False
    normalized = _D_EXPONENT.sub("E", _TYPE_SUFFIX.sub("", tok.raw_text))
    hex_match = _HEX.match(normalized)
    if hex_match:
        return int(hex_match.group(1), 16) == 0
    octal_match = _OCTAL.match(normalized)
    if octal_match:
        return int(octal_match.group(1), 8) == 0
    if _FLOAT.match(normalized) is None:
        return False
    return float(normalized) == 0


def _absolute_token_group_span(base: Span, toks: list[VbaToken]) -> Span:
    return Span(base.start + toks[0].start, base.start + toks[-1].end)
