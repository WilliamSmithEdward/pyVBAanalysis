"""Rule: a member of a project class instance used where what it holds, or what
it is, cannot serve (XLIDE issue #414).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/classInstanceValues.ts.
Each case measured in Excel 16.0 on an instance the procedure itself makes,
`Dim c As New Class1` or `Set c = New Class1`:

  c.M.Add 1, c.M(1), c.M & "x" with M an object field never set   91
  Main = c.M with M a Function that returns Nothing               91
  Main = c.M(1), c.M(1) = 2 with M a Get returning 1              13
  Set o = c.M with the same Get                                   424
  c.M.Count with M a Variant field never assigned                 424

Late-bound misuse of a member, a Sub assigned or a Set-only property read, is
runtime-member-not-found's.

The instance is followed only while the procedure keeps it to itself: any use of
it other than `c.Member`, or a second Set, ends the rule's reading of it. A field
assigned through it, `Set c.M = ...`, is not judged.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...completion import MemberCompletionContext
from ...conditional import ConditionalActivityTracker
from ...lexer.token_helpers import match_paren_from, top_level_equals_index
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...symbols.symbol_model import (
    ModuleSymbols,
    SymbolVisibility,
    VbaProjectClassMember,
    VbaProjectClassMembers,
    VbaSymbolKind,
)
from ...types.type_inference import procedure_symbol_for
from ...types.type_names import normalize_type
from ..context import PushFn
from ..walker import (
    active_module_members,
    for_each_statement,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

# Operators that read a value: beside one, an object member gives its default.
_VALUE_OPERATORS: frozenset[str] = frozenset(
    {"&", "+", "-", "*", "/", "\\", "^", "mod", "<", ">", "<=", ">=", "<>"}
)


@dataclass(frozen=True, slots=True)
class _Instance:
    name: str
    type: VbaProjectClassMembers


@dataclass(frozen=True, slots=True)
class _Statement:
    span: Span
    toks: list[VbaToken]


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined below 0 and past the end."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw(tok: VbaToken | None) -> str | None:
    return tok.raw_text if tok is not None else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name is not None else None


def check_class_instance_values(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    classes = {
        type_.name.lower(): type_
        for type_ in member_ctx.project_class_members or []
        if type_.kind == "class"
    }
    if not classes:
        return
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        proc_sym = procedure_symbol_for(symbols, member)
        locals_ = [
            child
            for child in (proc_sym.children if proc_sym is not None else None) or []
            if child.kind is VbaSymbolKind.LOCAL_VARIABLE
            and not child.is_array
            and child.visibility is not SymbolVisibility.STATIC
        ]
        if not locals_:
            continue
        # Every statement of the procedure, its single-line If arms apart.
        statements: list[_Statement] = []

        def collect(stmt: LeafStatementNode) -> None:
            for span in statement_and_branch_spans(stmt):
                statements.append(
                    _Statement(span, statement_tokens_after_leading_label(source, span))
                )

        for_each_statement(member.body, collect, activity)
        instances: dict[str, _Instance] = {}
        for local in locals_:
            lower = local.name.lower()
            declared = normalize_type(local.as_type)
            sets = [
                stmt
                for stmt in statements
                if token_text(_at(stmt.toks, 0)) == "set"
                and _lower_name(_at(stmt.toks, 1)) == lower
                and _raw(_at(stmt.toks, 2)) == "="
            ]
            type_: VbaProjectClassMembers | None = None
            if local.is_auto_instantiated and declared and declared in classes and len(sets) == 0:
                type_ = classes.get(declared)
            elif (
                len(sets) == 1
                and len(sets[0].toks) == 5
                and token_text(sets[0].toks[3]) == "new"
                and token_text(sets[0].toks[4]) in classes
                and (
                    declared is None
                    or declared == "object"
                    or declared == "variant"
                    or declared == token_text(sets[0].toks[4])
                )
            ):
                type_ = classes.get(token_text(sets[0].toks[4]))
            if type_ is None:
                continue
            # Kept to itself: every other mention is `c.Member`.
            own = all(
                _kept_to_itself(stmt.toks, i, lower, sets)
                for stmt in statements
                for i in range(len(stmt.toks))
            )
            if own:
                instances[lower] = _Instance(local.name, type_)
        if not instances:
            continue
        # Fields the procedure assigns through the instance hold what it gave.
        assigned: set[str] = set()
        for stmt in statements:
            toks = stmt.toks
            first = 1 if token_text(_at(toks, 0)) in ("set", "let") else 0
            lower_name = _lower_name(_at(toks, first))
            if (
                lower_name
                and lower_name in instances
                and _raw(_at(toks, first + 1)) == "."
                and token_name(_at(toks, first + 2))
            ):
                close = (
                    match_paren_from(toks, first + 3)
                    if _raw(_at(toks, first + 3)) == "("
                    else first + 2
                )
                if _raw(_at(toks, close + 1)) == "=":
                    assigned.add(f"{lower_name}.{token_text(toks[first + 2])}")
        for stmt in statements:
            _check_statement(stmt.span, stmt.toks, instances, assigned, push)


def _kept_to_itself(toks: list[VbaToken], i: int, lower: str, sets: Sequence[_Statement]) -> bool:
    tok = toks[i]
    if (
        tok.kind is not TokenKind.IDENTIFIER
        or tok.raw_text.lower() != lower
        or _raw(_at(toks, i - 1)) == "."
    ):
        return True
    return (
        _raw(_at(toks, i + 1)) == "."
        or (len(sets) == 1 and toks is sets[0].toks and i == 1)
        or token_text(_at(toks, 0)) == "dim"
    )


def _check_statement(
    span: Span,
    toks: Sequence[VbaToken],
    instances: Mapping[str, _Instance],
    assigned: AbstractSet[str],
    push: PushFn,
) -> None:
    head = token_text(_at(toks, 0))
    i = 0
    while i + 2 < len(toks):
        lower = _lower_name(toks[i])
        instance = instances.get(lower) if lower else None
        if (
            instance is None
            or _raw(_at(toks, i - 1)) == "."
            or toks[i + 1].raw_text != "."
            or not token_name(toks[i + 2])
        ):
            i += 1
            continue
        name = token_text(toks[i + 2])
        member = next(
            (candidate for candidate in instance.type.members if candidate.name.lower() == name),
            None,
        )
        if member is None:
            i += 1
            continue
        indexed = _raw(_at(toks, i + 3)) == "("
        close = match_paren_from(list(toks), i + 3) if indexed else i + 2
        if close < 0:
            i += 1
            continue
        after = _at(toks, close + 1)
        before = _at(toks, i - 1)
        first = 1 if head in ("set", "let") else 0
        target = i == first and _raw(after) == "="
        set_read = head == "set" and _raw(_at(toks, 2)) == "=" and i == 3 and close == len(toks) - 1
        plain_read = (
            head != "set"
            and i >= 2
            and toks[i - 1].raw_text == "="
            and i - 1 == top_level_equals_index(toks)
            and close == len(toks) - 1
        )
        member_of = _raw(after) == "." and not indexed
        label = "".join(tok.raw_text for tok in toks[i : close + 1])
        at = Span(span.start + toks[i].start, span.start + toks[close].end)
        shown = f"'{label}'"
        field_assigned = f"{lower}.{name}" in assigned
        # What it holds, through any binding.
        if member.known_value == "nothing" and not field_assigned and not target:
            operand = (
                after is not None and (token_text(after) or after.raw_text) in _VALUE_OPERATORS
            ) or (
                before is not None and (token_text(before) or before.raw_text) in _VALUE_OPERATORS
            )
            if (
                member_of
                or indexed
                or (operand and _is_object_type(member, "object"))
                or (plain_read and member.kind == "method")
            ):
                why = (
                    f"the Function {member.name} returns nothing else"
                    if member.kind == "method"
                    else f"nothing in {instance.type.name} sets {member.name}"
                )
                push(
                    "objectVariableNotSet",
                    f"{shown} is Nothing here: {why}. This will raise Run-time error '91': Object variable "
                    "or With block variable not set.",
                    at,
                )
                i += 1
                continue
        if member.known_value == "empty" and not field_assigned and member_of:
            push(
                "variantValueMisuse",
                f"{shown} is Empty here: nothing in {instance.type.name} assigns {member.name}, so it has no "
                "members. This will raise Run-time error '424': Object required.",
                at,
            )
            i += 1
            continue
        if member.known_value == "scalar":
            if indexed and (target or not member_of):
                push(
                    "variantValueMisuse",
                    f"{member.name} gives a single value, so {shown} has no element to "
                    f"{'assign' if target else 'read'}. This will raise Run-time error '13': Type mismatch.",
                    at,
                )
                i += 1
                continue
            if set_read and not indexed:
                push(
                    "variantValueMisuse",
                    f"{member.name} gives a single value, not an object, so Set has nothing to assign. This "
                    "will raise Run-time error '424': Object required.",
                    at,
                )
                i += 1
                continue
        i += 1


def _is_object_type(member: VbaProjectClassMember, type_: str) -> bool:
    return normalize_type(member.returns) == type_
