"""Rule family: statement forms the VBE refuses while compiling (XLIDE issue #125).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/statementForms.ts.

Measured in Excel 16.0 (build 20326, 2026-09-25):

- collection-operand: `x = c + 1` with c As New Collection -> "Argument not
  optional". A Collection's default member Item takes an index, so the bare
  variable has no value for the operator.
- sub-used-as-value: `x = Foo` where Foo is a Sub -> "Expected Function or
  variable".
- rem-after-then: `If x Then Rem note` -> "Syntax error". Rem starts a comment
  only at the start of a statement.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet

from ...conditional import ConditionalActivityTracker
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureSignature, VbaSymbolKind
from ...types.type_inference import procedure_symbol_for, type_environment_for
from ...types.type_names import normalize_type
from ..context import PushFn
from ..walker import (
    active_module_members,
    bare_assignment_target,
    for_each_statement,
    set_assignment_target,
    statement_and_branch_spans,
    statement_tokens,
    token_name,
    token_text,
)

_SCALAR_OPERATORS = frozenset({"=", "<", ">", "<=", ">=", "<>", "+", "-", "*", "/", "\\", "&", "^"})


def check_statement_forms(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Report a Rem after Then, a Collection used as an operand, and a Sub read as a value."""
    # Subs of this module, and of the project's standard modules, by name;
    # a name that is also a Function or a module-level variable anywhere is
    # not judged.
    subs: set[str] = set()
    not_subs: set[str] = set()
    for symbol in symbols.root.children or []:
        lower = symbol.name.lower()
        if symbol.kind is VbaSymbolKind.SUB:
            subs.add(lower)
        else:
            not_subs.add(lower)
    if project_procedures is not None:
        for key, signatures in project_procedures.items():
            for signature in signatures:
                (subs if signature.kind is VbaSymbolKind.SUB else not_subs).add(key.lower())
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        env = type_environment_for(symbols, member)
        locals_: set[str] = set()
        proc_sym = procedure_symbol_for(symbols, member)
        for child in (proc_sym.children if proc_sym is not None else None) or []:
            locals_.add(child.name.lower())

        def visit(stmt: LeafStatementNode) -> None:
            _check_statement(source, stmt, env, locals_, subs, not_subs, push)

        for_each_statement(member.body, visit, activity)


def _check_statement(
    source: str,
    stmt: LeafStatementNode,
    env: Mapping[str, str],
    locals_: AbstractSet[str],
    subs: AbstractSet[str],
    not_subs: AbstractSet[str],
    push: PushFn,
) -> None:
    for span in statement_and_branch_spans(stmt):
        toks = statement_tokens(source, span)
        if token_text(_at(toks, 0)) == "if":
            then = _find_index(toks, lambda tok: token_text(tok) == "then")
            if then > 0 and token_text(_at(toks, then + 1)) == "rem":
                push(
                    "remAfterThen",
                    "'Rem' cannot follow 'Then' on one line: a Rem comment starts only at the "
                    "start of a statement. This is a VBE compile error: Syntax error.",
                    _token_span(span, toks[then + 1]),
                )
        target = bare_assignment_target(source, span)
        # A Set's `=` is the assignment too: `Set c = New Collection` is no operand.
        assigns = target is not None or set_assignment_target(source, span) is not None
        eq = _find_index(toks, lambda tok: tok.raw_text == "=") if assigns else -1
        for i, tok in enumerate(toks):
            name = token_name(tok)
            if not name or _raw_at(toks, i - 1) == "." or _raw_at(toks, i + 1) == ":=" or i == eq - 1:
                continue
            lower = name.lower()
            following = _raw_at(toks, i + 1)
            if normalize_type(env.get(lower)) == "collection" and following != "(" and following != ".":
                previous = None if i - 1 == eq else _at(toks, i - 1)
                operator = next(
                    (
                        candidate
                        for candidate in (_at(toks, i + 1), previous)
                        if candidate is not None
                        and (
                            (
                                candidate.kind is TokenKind.OPERATOR
                                and candidate.raw_text in _SCALAR_OPERATORS
                            )
                            or token_text(candidate) == "mod"
                        )
                    ),
                    None,
                )
                if operator is not None:
                    push(
                        "collectionOperand",
                        f"'{name}' is a Collection: its default member Item needs an index, so "
                        f"'{operator.raw_text}' has no value to work on. This is a VBE compile "
                        "error: Argument not optional.",
                        _token_span(span, tok),
                    )
                    continue
            if (
                target is not None
                and i > eq
                and lower in subs
                and lower not in not_subs
                and lower not in locals_
                and lower not in env
            ):
                push(
                    "subUsedAsValue",
                    f"'{name}' is a Sub, which returns nothing, so it cannot be used as a value. "
                    "This is a VBE compile error: Expected Function or variable.",
                    _token_span(span, tok),
                )


def _token_span(base: Span, tok: VbaToken) -> Span:
    return Span(base.start + tok.start, base.start + tok.end)


def _find_index(toks: Sequence[VbaToken], predicate: Callable[[VbaToken], bool]) -> int:
    """Array.prototype.findIndex: the first token the predicate accepts, or -1."""
    for k, tok in enumerate(toks):
        if predicate(tok):
            return k
    return -1


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) outside the list."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    """`toks[i]?.rawText`."""
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None
