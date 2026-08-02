"""Rule family: late binding.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/lateBinding.ts. A
receiver whose static type is Variant or Object dispatches through IDispatch at
runtime. Friend and Private members are not on a class's dispatch interface, so
a late-bound call to one raises runtime error 438 - and the VBA compiler says
nothing, because it cannot know the runtime type either. That combination makes
the failure easy to ship: the code compiles, and the call only dies on the
first execution that reaches it.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass

from ...host.host_model import is_host_member_name
from ...lexer.token_kinds import VbaToken
from ...parser.nodes import LeafStatementNode, ProcedureNode, Span
from ...runtime import resolve_runtime_object, resolve_runtime_object_type
from ...symbols.symbol_model import (
    ModuleSymbols,
    SymbolVisibility,
    VbaProjectClassMembers,
    VbaSymbol,
)
from ...types.type_inference import (
    declaration_shape_environment_for,
    declared_value_type_for_source_binding,
    procedure_symbol_for,
)
from ...types.type_names import normalize_type
from ..context import PushFn
from ..walker import ProcedureStatementVisitor, statement_tokens, token_name

# Declared types whose members can only be reached through IDispatch.
_LATE_BOUND_TYPES = frozenset({"variant", "object"})


def check_late_bound_friend_member(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    project_class_members: Sequence[VbaProjectClassMembers],
    push: PushFn,
) -> ProcedureStatementVisitor:
    """Per-statement rule: member access on a late-bound receiver where the member
    name is Friend-only across the project's class modules.

    Deliberately narrow. Project class members that are `Private` never reach
    the project member surfaces at all (the index drops them), so "unknown
    member" and "Private member" are indistinguishable here - and the VBE oracle
    says an unknown member on a Variant/Object receiver must stay silent,
    because the runtime type is unknowable and the call may well be legal.
    Firing only on names that resolve EXCLUSIVELY to Friend members keeps the
    rule on the one case where there is provably no legal late-bound target
    anywhere in scope."""
    friend_only = _friend_only_member_names(project_class_members)
    if not friend_only:
        return lambda member: None

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        shapes = declaration_shape_environment_for(symbols, member)
        proc_sym = procedure_symbol_for(symbols, member)

        def declared_type_of(name: str) -> str | None:
            shape = shapes.get(name.lower())
            if shape is not None:
                # An entry with no `As` clause is an implicit Variant.
                return shape.as_type if shape.as_type is not None else "Variant"
            binding = declared_value_type_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name
            )
            # `resolved` with no type also covers ambiguous multi-definition
            # bindings, which are not proof of anything - stay quiet.
            return binding.as_type if binding.resolved else None

        def visitor(stmt: LeafStatementNode) -> None:
            toks = statement_tokens(source, stmt.span)
            for access in _late_bound_member_accesses(toks, declared_type_of):
                owners = friend_only.get(access.member.lower())
                if owners is None:
                    continue
                push(
                    "lateBoundFriendMember",
                    f"'{access.member}' is a Friend member of {_describe_owners(owners)}, "
                    f"so it is not on the dispatch interface. {access.receiver_description} "
                    f"is late bound ({access.receiver_type}), so this raises run-time error "
                    f"438 at run time even though it compiles. Assign to a variable of the "
                    f"class type first, then access the member through that.",
                    Span(stmt.span.start + access.token.start, stmt.span.start + access.token.end),
                )

        return visitor

    return factory


@dataclass(frozen=True, slots=True)
class _LateBoundAccess:
    member: str
    token: VbaToken
    receiver_type: str
    receiver_description: str


_TypeResolver = Callable[[str], "str | None"]


def _late_bound_member_accesses(
    toks: Sequence[VbaToken], declared_type_of: _TypeResolver
) -> list[_LateBoundAccess]:
    """Find ``<late-bound receiver>.Member`` in one statement's tokens.

    Two receiver shapes are recognized, both of which appear in real code:
    a bare identifier declared As Variant/As Object/with no type; and a
    ``Collection`` subscript or ``.Item(...)`` call, whose result is Variant.
    The second is the shape that hides the bug in practice, because the
    collection itself is strongly typed and only its element type is lost."""
    out: list[_LateBoundAccess] = []
    for i in range(1, len(toks) - 1):
        if toks[i].raw_text != ".":
            continue
        member_name = token_name(toks[i + 1])
        if not member_name:
            continue
        receiver = _receiver_before(toks, i, declared_type_of)
        if receiver is None:
            continue
        receiver_type, receiver_description = receiver
        out.append(
            _LateBoundAccess(
                member=member_name,
                token=toks[i + 1],
                receiver_type=receiver_type,
                receiver_description=receiver_description,
            )
        )
    return out


def _receiver_before(
    toks: Sequence[VbaToken], dot_index: int, declared_type_of: _TypeResolver
) -> tuple[str, str] | None:
    prev = toks[dot_index - 1] if dot_index >= 1 else None
    if prev is None:
        return None

    if prev.raw_text == ")":
        open_index = _matching_open_paren(toks, dot_index - 1)
        if open_index is None:
            return None
        return _collection_element_receiver(toks, open_index, declared_type_of)

    # A bare identifier, and not itself the tail of a longer chain: a chain like
    # `a.b.c` would need b's return type, which is a different question.
    if dot_index >= 2 and toks[dot_index - 2].raw_text == ".":
        return None
    name = token_name(prev)
    if not name:
        return None
    normalized = normalize_type(declared_type_of(name))
    if not normalized or normalized not in _LATE_BOUND_TYPES:
        return None
    # A name that is also a built-in runtime object (Err, Debug) is not a
    # late-bound local no matter what a same-named declaration says.
    if resolve_runtime_object(name) is not None:
        return None
    return ("As Object" if normalized == "object" else "As Variant", f"'{name}'")


def _collection_element_receiver(
    toks: Sequence[VbaToken], open_index: int, declared_type_of: _TypeResolver
) -> tuple[str, str] | None:
    """Recognize ``coll(i)`` and ``coll.Item(i)`` where ``coll`` is declared As
    Collection. ``Collection.Item`` returns Variant, so the element arrives late
    bound however strongly typed the objects inside the collection actually are."""
    root_index = open_index - 1
    via_item = False
    before_paren = token_name(toks[root_index]) if root_index >= 0 else None
    if (
        before_paren
        and before_paren.lower() == "item"
        and root_index >= 1
        and toks[root_index - 1].raw_text == "."
    ):
        via_item = True
        root_index -= 2
    if root_index >= 1 and toks[root_index - 1].raw_text == ".":
        return None
    root_name = token_name(toks[root_index]) if root_index >= 0 else None
    if not root_name:
        return None
    if normalize_type(declared_type_of(root_name)) != "collection":
        return None
    return (
        "Collection.Item returns Variant",
        f"'{root_name}.Item(...)'" if via_item else f"'{root_name}(...)'",
    )


def _matching_open_paren(toks: Sequence[VbaToken], close_index: int) -> int | None:
    depth = 0
    for i in range(close_index, -1, -1):
        text = toks[i].raw_text
        if text == ")":
            depth += 1
        elif text == "(":
            depth -= 1
            if depth == 0:
                return i
    return None


def _friend_only_member_names(
    project_class_members: Sequence[VbaProjectClassMembers],
) -> dict[str, list[str]]:
    """Member names that appear ONLY as Friend members of project class modules.

    A name that is also Public anywhere, or that exists in the host object model
    or the VBA runtime objects, could legally dispatch to that other target, so
    it is excluded - the runtime type inside a Variant is unknowable and the
    rule must not guess."""
    friend_owners: dict[str, list[str]] = {}
    disqualified: set[str] = set()
    for owner in project_class_members:
        # Only plain class modules are source-exhaustive. Document modules and
        # UserForms also expose host/designer members we cannot enumerate.
        exhaustive_class = owner.kind == "class" and owner.exhaustive is not False
        for member in owner.members:
            lower = member.name.lower()
            if member.visibility is SymbolVisibility.FRIEND and exhaustive_class:
                friend_owners.setdefault(lower, []).append(owner.name)
            else:
                disqualified.add(lower)
    out: dict[str, list[str]] = {}
    for lower, owners in friend_owners.items():
        if lower in disqualified:
            continue
        if is_host_member_name(lower):
            continue
        if _is_runtime_object_member_name(lower):
            continue
        out[lower] = owners
    return out


_RUNTIME_OBJECT_MEMBER_NAMES: set[str] = set()


def _is_runtime_object_member_name(lower: str) -> bool:
    if not _RUNTIME_OBJECT_MEMBER_NAMES:
        for name in ("Collection", "Err", "Debug", "Dictionary"):
            obj = resolve_runtime_object(name) or resolve_runtime_object_type(name)
            for member in (obj or {}).get("members", []):
                _RUNTIME_OBJECT_MEMBER_NAMES.add(member["name"].lower())
    return lower in _RUNTIME_OBJECT_MEMBER_NAMES


def _describe_owners(owners: Sequence[str]) -> str:
    unique = list(dict.fromkeys(owners))
    if len(unique) == 1:
        return f"class '{unique[0]}'"
    return "classes " + ", ".join(f"'{name}'" for name in unique)
