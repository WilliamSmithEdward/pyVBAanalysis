"""Rule family: a standard module's member used against its kind, through the
module's name or bare (XLIDE issue #423).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/moduleMembers.ts. Each
is a compile error, measured in Excel 16.0 (build 20430, 2026-10-02) with M in
Module2 and the use in Module1:

- `Module2.M = 9` with M a Sub: "Expected Function or variable"; with M a
  Function or Declare returning a type of VBA's own: "Function call on
  left-hand side of assignment must return Variant or Object".
- `Main = Module2.M` with M a Sub: "Expected Function or variable".
- `Module2.M` or `Call Module2.M` with M a variable, a Const or an Enum
  member: "Expected procedure, not variable".
- `M`, `Call M`, `Module2.M` or `Call Module2.M` with M a Property Get and
  no Let or Set: "Invalid use of property"; `M = 9`: "Can't assign to
  read-only property".
- `M = 9` or `Main = M` with M a Type: "Variable not defined", under
  Option Explicit.

The bare forms of the first three are judged elsewhere already.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence

from ...completion.member_access import MemberCompletionContext
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ProcedureNode, Span
from ...symbols.name_resolution import BareIdentifierContext, BareIdentifierResolutionScope
from ...symbols.symbol_model import (
    ModuleSymbolKind,
    ModuleSymbols,
    SymbolVisibility,
    VbaSymbol,
    VbaSymbolKind,
)
from ...types.type_inference import procedure_symbol_for, source_identifier_binding
from ...types.type_names import is_known_scalar_type, normalize_type
from ..context import PushFn
from ..walker import (
    ProcedureStatementVisitor,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

# The kinds of symbol a bare name reads as a value.
_VALUE_KINDS = frozenset(
    {
        VbaSymbolKind.MODULE_VARIABLE,
        VbaSymbolKind.CONSTANT,
        VbaSymbolKind.FUNCTION,
        VbaSymbolKind.PROPERTY_GET,
        VbaSymbolKind.DECLARE,
    }
)

_COMPARING_HEADS = frozenset({"if", "elseif", "do", "loop", "while", "select", "case", "for"})

_PROPERTY_KINDS = (
    VbaSymbolKind.PROPERTY_GET,
    VbaSymbolKind.PROPERTY_LET,
    VbaSymbolKind.PROPERTY_SET,
)


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) outside the list."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    """`toks[i]?.rawText`."""
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


def _module_member(
    module_name: str, lower: str, symbols: ModuleSymbols, visible: Sequence[VbaSymbol]
) -> list[VbaSymbol]:
    """The symbols named `lower` that a standard module declares, Public or by default."""
    own = symbols.module_name.lower() == module_name.lower()
    pool = (
        [
            sym
            for sym in symbols.all
            if not sym.container_name or sym.kind is VbaSymbolKind.ENUM_MEMBER
        ]
        if own
        else [sym for sym in visible if sym.module_name.lower() == module_name.lower()]
    )
    return [
        sym
        for sym in pool
        if sym.name.lower() == lower
        and sym.kind is not VbaSymbolKind.LOCAL_VARIABLE
        and sym.kind is not VbaSymbolKind.PARAMETER
        and (
            own
            or (
                sym.visibility is not SymbolVisibility.PRIVATE
                and sym.visibility is not SymbolVisibility.DIM
            )
        )
    ]


def _property_only(found: Sequence[VbaSymbol]) -> bool:
    """Whether every symbol of the name is a Property Get, Let or Set."""
    return len(found) > 0 and all(sym.kind in _PROPERTY_KINDS for sym in found)


def _get_only(found: Sequence[VbaSymbol]) -> bool:
    return (
        len(found) > 0
        and any(sym.kind is VbaSymbolKind.PROPERTY_GET for sym in found)
        and all(sym.kind is VbaSymbolKind.PROPERTY_GET for sym in found)
    )


def _procedure_assignment(found: Sequence[VbaSymbol], name: str) -> str | None:
    """Why a Sub, or a Function of a type of VBA's own, cannot be assigned."""
    if len(found) != 1:
        return None
    sym = found[0]
    if sym.kind is VbaSymbolKind.SUB or (
        sym.kind is VbaSymbolKind.DECLARE and sym.declare_kind == "Sub"
    ):
        return (
            f"'{name}' is a Sub, which has no value to assign. This is a VBE compile error: "
            "Expected Function or variable."
        )
    returns = normalize_type(sym.as_type)
    is_function = sym.kind is VbaSymbolKind.FUNCTION or (
        sym.kind is VbaSymbolKind.DECLARE and sym.declare_kind == "Function"
    )
    if (
        is_function
        and returns
        and returns != "variant"
        and is_known_scalar_type(returns)
        and not sym.is_array
    ):
        return (
            f"'{name}' is a Function returning {sym.as_type}, and a call cannot be assigned to. "
            "This is a VBE compile error: Function call on left-hand side of assignment must "
            "return Variant or Object."
        )
    return None


def check_module_member_forms(
    source: str,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    explicit: bool,
    push: PushFn,
) -> ProcedureStatementVisitor:
    modules = {
        type_.name.lower()
        for type_ in (member_ctx.project_class_members or [])
        if type_.kind == "standardModule"
    }
    if symbols.module_kind is ModuleSymbolKind.STANDARD:
        modules.add(symbols.module_name.lower())
    visible: Sequence[VbaSymbol] = project_visible_symbols or []

    def for_member(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        proc_sym = procedure_symbol_for(symbols, member)
        locals_ = {
            name.lower()
            for name in [
                member.name,
                *(param.name for param in member.params),
                *(
                    child.name
                    for child in (proc_sym.children if proc_sym is not None else None) or []
                ),
            ]
        }

        def is_module(tok: VbaToken | None) -> bool:
            name = token_name(tok)
            lower = name.lower() if name is not None else None
            return lower is not None and lower in modules and lower not in locals_

        def at(span: Span, from_: VbaToken, to: VbaToken) -> Span:
            return Span(span.start + from_.start, span.start + to.end)

        def check(span: Span) -> None:
            toks = statement_tokens_after_leading_label(source, span)
            head = token_text(_at(toks, 0))
            first = 1 if head == "call" else 0
            assign_at = (
                -1
                if head in _COMPARING_HEADS or head == "set"
                else next(
                    (
                        k
                        for k, tok in enumerate(toks)
                        if tok.raw_text == "=" and tok.kind is TokenKind.OPERATOR
                    ),
                    -1,
                )
            )
            # `Module2.M` at the statement's start: called, or assigned.
            first_tok = _at(toks, first)
            member_tok = _at(toks, first + 2)
            module_word = token_name(first_tok)
            member_word = token_name(member_tok)
            if (
                is_module(first_tok)
                and _raw_at(toks, first + 1) == "."
                and member_word is not None
                and first_tok is not None
                and member_tok is not None
                and module_word is not None
            ):
                found = _module_member(module_word, member_word.lower(), symbols, visible)
                label = f"{first_tok.raw_text}.{member_tok.raw_text}"
                where = at(span, first_tok, member_tok)
                if first == 0 and assign_at == first + 3:
                    why = _procedure_assignment(found, label)
                    if why:
                        push("assignmentToProcedureName", why, where)
                    return
                statement_call = (
                    assign_at < 0
                    and _raw_at(toks, first + 3) != "."
                    and _raw_at(toks, first + 3) != "!"
                )
                if (
                    statement_call
                    and len(found) == 1
                    and found[0].kind
                    in (
                        VbaSymbolKind.MODULE_VARIABLE,
                        VbaSymbolKind.CONSTANT,
                        VbaSymbolKind.ENUM_MEMBER,
                    )
                ):
                    kind = (
                        "a module variable"
                        if found[0].kind is VbaSymbolKind.MODULE_VARIABLE
                        else "a constant"
                        if found[0].kind is VbaSymbolKind.CONSTANT
                        else "an Enum member"
                    )
                    push(
                        "nonCallableCallStatement",
                        f"Cannot call '{label}' because it resolves to {kind}, not a Sub or "
                        "Function. This is a VBE compile error: Expected procedure, not variable.",
                        where,
                    )
                    return
                if statement_call and _property_only(found):
                    push(
                        "invalidPropertyUse",
                        f"'{label}' is a property, and a statement cannot call one. This is a VBE "
                        "compile error: Invalid use of property.",
                        where,
                    )
                return
            # `Main = Module2.M` with M a Sub: the whole value.
            if (
                assign_at > 0
                and len(toks) == assign_at + 4
                and is_module(_at(toks, assign_at + 1))
                and _raw_at(toks, assign_at + 2) == "."
            ):
                module_tok = toks[assign_at + 1]
                name_tok = toks[assign_at + 3]
                name_word = token_name(name_tok)
                found = _module_member(
                    token_name(module_tok) or "",
                    name_word.lower() if name_word is not None else "",
                    symbols,
                    visible,
                )
                if len(found) == 1 and found[0].kind is VbaSymbolKind.SUB:
                    push(
                        "subUsedAsValue",
                        f"'{module_tok.raw_text}.{name_tok.raw_text}' is a Sub, which returns "
                        "nothing, so it cannot be used as a value. This is a VBE compile error: "
                        "Expected Function or variable.",
                        at(span, module_tok, name_tok),
                    )
                return
            # Bare `M`, `Call M` and `M = 9`, with M a standard module's Property Get.
            bare = token_name(first_tok)
            if (
                not bare
                or first_tok is None
                or _raw_at(toks, first + 1) == "."
                or _raw_at(toks, first + 1) == "("
            ):
                return
            binding = source_identifier_binding(
                symbols, proc_sym, project_visible_symbols, bare, BareIdentifierContext.CALL
            )
            in_module = [sym for sym in binding.definitions if sym.module_name.lower() in modules]
            # Inside the property, its name is its value: `M = 1` in Get M.
            own = bare.lower() == member.name.lower()
            if (
                not own
                and binding.scope is not BareIdentifierResolutionScope.AMBIGUOUS
                and len(in_module) == len(binding.definitions)
                and _property_only(in_module)
            ):
                where = at(span, first_tok, first_tok)
                from .assignments import _getter_may_return_object
                from ...types.type_inference import def_type_of
                getter = next((sym for sym in in_module if sym.kind is VbaSymbolKind.PROPERTY_GET), None)
                getter_type = getter.as_type or (def_type_of(symbols, getter.name) if getter.module_name.lower() == symbols.module_name.lower() else None) if getter else None
                if assign_at == first + 1 and first == 0 and _get_only(in_module) and not _getter_may_return_object(getter_type, member_ctx):
                    push(
                        "readonlyMemberAssignment",
                        f"'{bare}' has a Property Get and no Property Let, so it cannot be "
                        "assigned. This is a VBE compile error: Can't assign to read-only "
                        "property.",
                        where,
                    )
                elif assign_at < 0 and head != "set":
                    push(
                        "invalidPropertyUse",
                        f"'{bare}' is a property, and a statement cannot call one. This is a VBE "
                        "compile error: Invalid use of property.",
                        where,
                    )
                return
            # `M = 9` and `Main = M` with M a Type: no variable of that name.
            if not explicit or assign_at < 0:
                return
            # `TypeName(M)` reads M as a value too (XLIDE issue #639).
            type_name_arguments = [
                k
                for k in range(len(toks))
                if token_text(_at(toks, k - 2)) == "typename"
                and _raw_at(toks, k - 1) == "("
                and _raw_at(toks, k + 1) == ")"
            ]
            for index in [0, assign_at + 1, *type_name_arguments]:
                tok = _at(toks, index)
                name = token_name(tok)
                as_argument = index in type_name_arguments
                if (
                    not name
                    or tok is None
                    or (
                        not as_argument
                        and (
                            (index == 0 and assign_at != 1)
                            or (index > 0 and len(toks) != assign_at + 2)
                        )
                    )
                ):
                    continue
                # A variable, Const or Function of the name elsewhere in the
                # project is what the name reads: `Public Zq As Long` in Module2
                # beside a Private Type Zq here runs (XLIDE issue #639, measured in
                # Excel 16.0).
                lower = name.lower()
                if any(
                    sym.name.lower() == lower
                    and sym.kind in _VALUE_KINDS
                    and sym.visibility is not SymbolVisibility.PRIVATE
                    and sym.module_name.lower() != symbols.module_name.lower()
                    for sym in visible
                ):
                    continue
                found_binding = source_identifier_binding(
                    symbols,
                    proc_sym,
                    project_visible_symbols,
                    name,
                    BareIdentifierContext.ASSIGNMENT_TARGET
                    if index == 0
                    else BareIdentifierContext.EXPRESSION,
                )
                if (
                    found_binding.scope is BareIdentifierResolutionScope.UNRESOLVED
                    or (
                        len(found_binding.definitions) > 0
                        and all(sym.kind is VbaSymbolKind.TYPE for sym in found_binding.definitions)
                    )
                ) and _type_named(name, symbols, visible):
                    push(
                        "undeclaredVariable",
                        f"Variable not defined: '{name}'. It names a user-defined type, not a "
                        "variable. This is a VBE compile error.",
                        at(span, tok, tok),
                    )

        def visit(stmt: LeafStatementNode) -> None:
            for span in statement_and_branch_spans(stmt):
                check(span)

        return visit

    return for_member


def _type_named(name: str, symbols: ModuleSymbols, visible: Sequence[VbaSymbol]) -> bool:
    """Whether a Type of that name is visible: the module's own, or another standard
    module's Public one."""
    lower = name.lower()
    return any(
        sym.kind is VbaSymbolKind.TYPE and sym.name.lower() == lower
        for sym in [*(symbols.root.children or []), *visible]
    )
