"""Member-access type/surface resolver (reduced port of memberAccess.ts).

Given VBA source and an offset just after a member-access dot, this resolves the
type of the receiver expression and returns the verified member surface available
on it. The diagnostics consume exactly one seam,
``resolve_member_surface_at(source, offset, ctx) -> ResolvedMemberSurface`` (wrapped
by ``resolve_exhaustive_member_surface`` in ``rules/shared``), and read only
``surface.owner`` (the message string) and member ``.name`` (a lowercased
membership test).

This is a DELIBERATELY REDUCED subset of the 1394-line memberAccess.ts. Dropped
verbatim: every completion-UX path (``resolveMemberCompletions``,
``resolveMemberCompletionNamed``, ``resolveMemberDefinitionsAt``,
``completionFromSurfaceMember``, signature/doc/returns rendering, the typed-prefix
filtering, the partial-identifier peel, and all VbaDoc plumbing). What remains is
exactly the receiver-type resolution and the member-surface construction the
no-false-positive member-not-found contract rides on. The EXHAUSTIVE flag is never
synthesized: host surfaces use the host model's ``exhaustive`` flag and project
surfaces use ``VbaProjectClassMembers.exhaustive`` verbatim.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass

from ..host import (
    HostObjectModel,
    get_host_members,
    get_host_type,
    resolve_host_alias,
    resolve_host_global,
)
from ..host.host_model import HostMember
from ..lexer.token_helpers import IDENT_RE, is_ident_like
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize
from ..parser.nodes import (
    BodyNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    VariableGroupNode,
    is_leaf_statement,
)
from ..parser.parse_module import parse_module
from ..runtime import resolve_runtime_object, resolve_runtime_object_type
from ..symbols.symbol_model import VbaProjectClassMember, VbaProjectClassMembers
from ..types.type_names import is_known_scalar_type, normalize_type
from .cursor_context import completion_significant_tokens

_PROJECT_TYPE_PREFIX = "project:"
_COMBINED_TYPE_PREFIX = "combined:"
_COMBINED_TYPE_SEPARATOR = "|"
_UNION_TYPE_PREFIX = "union:"
_UNION_TYPE_SEPARATOR = "|"


@dataclass(slots=True)
class MemberCompletionContext:
    """Project/module facts the resolver needs that come from outside the source."""

    # Lowercased worksheet/document code name -> qualified host type, by CODE NAME.
    code_names: dict[str, str] | None = None
    # Qualified host type that `Me` resolves to in the current module.
    me_type: str | None = None
    # Project object type that `Me` resolves to in the current class/document module.
    me_project_type: str | None = None
    # Source-declared workbook object members and visible UDT fields, keyed by type.
    project_class_members: Sequence[VbaProjectClassMembers] | None = None
    # True/default lets generic Object/Variant receivers narrow from preceding
    # simple Set assignments. Hard diagnostics disable this because VBA still
    # compile-binds those receivers late.
    allow_set_assignment_refinement: bool = False
    # Host object model to resolve against. Defaults to the Excel model when None.
    model: HostObjectModel | None = None
    # Pre-parsed AST of the analyzed source, when the caller already holds one.
    parsed_module: ModuleNode | None = None
    # Full-source significant tokens (comments removed, newlines kept), used to
    # slice the prefix token stream by offset instead of re-lexing per reference.
    source_tokens: Sequence[VbaToken] | None = None


@dataclass(frozen=True, slots=True)
class MemberCompletionEntry:
    """One member of a resolved surface. The diagnostics read only ``name``."""

    name: str
    kind: str


@dataclass(frozen=True, slots=True)
class ResolvedMemberSurface:
    """The seam consumed by the diagnostics: owner string + members + exhaustive."""

    owner: str
    members: list[MemberCompletionEntry]
    exhaustive: bool


@dataclass(slots=True)
class _ReceiverChainSegment:
    name: str
    has_arguments: bool


@dataclass(slots=True)
class _ReceiverChain:
    segments: list[_ReceiverChainSegment]
    start_index: int


@dataclass(frozen=True, slots=True)
class _ResolvedMemberReturn:
    type: str
    kind: str


@dataclass(slots=True)
class _MemberSurface:
    """Internal surface (with raw member dicts) before flattening to the seam."""

    owner: str
    members: list[MemberCompletionEntry]
    exhaustive: bool


@dataclass(frozen=True, slots=True)
class _SetAssignment:
    name: str
    value_tokens: list[VbaToken]
    offset: int


@dataclass(frozen=True, slots=True)
class _DeclaredBinding:
    as_type: str | None


def _word(token: VbaToken) -> str:
    return token.raw_text


def _is_boundary(token: VbaToken) -> bool:
    """A logical-line boundary: a newline or a statement-separating colon."""
    return token.kind is TokenKind.NEWLINE or token.raw_text == ":"


def _at(tokens: Sequence[VbaToken], i: int) -> VbaToken | None:
    return tokens[i] if 0 <= i < len(tokens) else None


# -- public seam -----------------------------------------------------------


def resolve_member_surface_at(
    source: str, offset: int, ctx: MemberCompletionContext | None = None
) -> ResolvedMemberSurface | None:
    """Resolve the complete source/host member surface at a member-access dot.

    Includes empty-but-exhaustive project surfaces (a class with no public members
    is still exhaustive). Returns None when the receiver cannot be resolved.
    """
    ctx = ctx if ctx is not None else MemberCompletionContext()
    current_type = resolve_receiver_type_at(source, offset, ctx)
    if current_type is None:
        return None
    surface = _member_surface_for_type(current_type, ctx)
    if surface is None:
        return None
    return ResolvedMemberSurface(
        owner=surface.owner, members=list(surface.members), exhaustive=surface.exhaustive
    )


def resolve_receiver_type_at(
    source: str, offset: int, ctx: MemberCompletionContext | None = None
) -> str | None:
    """The qualified type whose members are accessible at the dot ending before
    ``offset``, or None when the cursor is not in a member-access position or the
    receiver cannot be resolved. A partially typed member name after the dot is
    ignored."""
    ctx = ctx if ctx is not None else MemberCompletionContext()
    tokens = _prefix_significant_tokens(source, offset, ctx)
    if len(tokens) == 0:
        return None
    i = len(tokens) - 1
    if is_ident_like(tokens[i]) and i > 0 and tokens[i - 1].raw_text == ".":
        i -= 1
    if i < 0 or tokens[i].raw_text != ".":
        return None
    return _receiver_type_from_tokens(tokens, i, source, offset, ctx)


# -- prefix token slicing --------------------------------------------------


def _prefix_significant_tokens(
    source: str, offset: int, ctx: MemberCompletionContext
) -> list[VbaToken]:
    """Significant prefix tokens for ``offset``. When the context carries
    full-source tokens and a token ends exactly at ``offset``, slice the shared
    stream instead of re-lexing the prefix; the cut sits on a token boundary so
    the two paths produce identical tokens."""
    shared = ctx.source_tokens
    if shared and len(shared) > 0:
        lo = 0
        hi = len(shared) - 1
        found = -1
        while lo <= hi:
            mid = (lo + hi) >> 1
            if shared[mid].end <= offset:
                found = mid
                lo = mid + 1
            else:
                hi = mid - 1
        if found >= 0 and shared[found].end == offset:
            return list(shared[: found + 1])
    return completion_significant_tokens(source, offset)


# -- receiver-type resolution ----------------------------------------------


def _receiver_type_from_tokens(
    tokens: Sequence[VbaToken],
    dot_index: int,
    source: str,
    offset: int,
    ctx: MemberCompletionContext,
) -> str | None:
    """Walk the receiver chain ending at the dot ``tokens[dot_index]`` and resolve
    it to a qualified host/project type, threading return types through each hop."""
    chain = _collect_receiver_chain(tokens, dot_index - 1)
    explicit_receiver = _receiver_type_from_chain(chain, source, offset, ctx)
    if explicit_receiver:
        return explicit_receiver
    grouped = _receiver_type_from_parenthesized_receiver(
        tokens, dot_index - 1, source, offset, ctx
    )
    if grouped:
        return grouped
    implicit_with_chain = _collect_implicit_with_chain(tokens, dot_index - 1)
    if implicit_with_chain is None:
        return None
    return _receiver_type_from_implicit_with_chain(
        _with_receiver_type_at(source, tokens[dot_index].end, ctx), implicit_with_chain, ctx
    )


def _receiver_type_from_implicit_with_chain(
    with_type: str | None,
    chain: list[_ReceiverChainSegment],
    ctx: MemberCompletionContext,
) -> str | None:
    current_type = with_type
    for segment in chain:
        if not current_type:
            return None
        resolved = _resolve_any_member_return_type(current_type, segment.name, ctx)
        if resolved is None:
            return None
        current_type = _apply_default_member_return_type(
            resolved.type, segment.has_arguments and resolved.kind != "method", ctx
        )
    return current_type


def _receiver_type_from_expression_tokens(
    tokens: Sequence[VbaToken], source: str, offset: int, ctx: MemberCompletionContext
) -> str | None:
    if len(tokens) == 0:
        return None
    chain = _collect_receiver_chain_with_start(tokens, len(tokens) - 1)
    if chain is None:
        return None
    prefix = tokens[: chain.start_index]
    if len(prefix) > 0 and not (
        len(prefix) == 1 and prefix[0].raw_text.lower() == "new"
    ):
        return None
    return _receiver_type_from_chain(chain.segments, source, offset, ctx)


def _receiver_type_from_parenthesized_receiver(
    tokens: Sequence[VbaToken],
    end_index: int,
    source: str,
    offset: int,
    ctx: MemberCompletionContext,
) -> str | None:
    if end_index < 0 or tokens[end_index].raw_text != ")":
        return None
    open_index = _match_paren_left(tokens, end_index)
    if open_index < 0:
        return None
    expression_tokens = list(tokens[open_index + 1 : end_index])
    return _receiver_type_from_expression_tokens(
        expression_tokens, source, offset, ctx
    ) or _receiver_type_from_parenthesized_receiver(
        expression_tokens, len(expression_tokens) - 1, source, offset, ctx
    )


def _receiver_type_from_chain(
    chain: list[_ReceiverChainSegment],
    source: str,
    offset: int,
    ctx: MemberCompletionContext,
) -> str | None:
    if len(chain) == 0:
        return None
    root = chain[0]
    root_type = _resolve_root(root.name, source, offset, ctx)
    if not root_type:
        return None
    current_type: str | None = _apply_default_member_return_type(
        root_type, root.has_arguments, ctx
    )
    s = 1
    while s < len(chain) and current_type:
        segment = chain[s]
        resolved = _resolve_any_member_return_type(current_type, segment.name, ctx)
        if resolved is None:
            return None
        current_type = _apply_default_member_return_type(
            resolved.type, segment.has_arguments and resolved.kind != "method", ctx
        )
        s += 1
    return current_type


# -- chain collection ------------------------------------------------------


def _collect_receiver_chain(
    tokens: Sequence[VbaToken], end_index: int
) -> list[_ReceiverChainSegment]:
    chain = _collect_receiver_chain_with_start(tokens, end_index)
    return chain.segments if chain is not None else []


def _collect_receiver_chain_with_start(
    tokens: Sequence[VbaToken], end_index: int
) -> _ReceiverChain | None:
    segments: list[_ReceiverChainSegment] = []
    i = end_index
    pending_has_arguments = False
    start_index = -1
    while True:
        if i >= 0 and _is_boundary(tokens[i]):
            return None
        if i >= 0 and tokens[i].raw_text == ")":
            open_index = _match_paren_left(tokens, i)
            if open_index < 0:
                return None
            pending_has_arguments = True
            i = open_index - 1
            continue
        if i < 0 or not is_ident_like(tokens[i]):
            return None
        start_index = i
        segments.insert(0, _ReceiverChainSegment(_word(tokens[i]), pending_has_arguments))
        pending_has_arguments = False
        i -= 1
        if i >= 0 and tokens[i].raw_text == ".":
            i -= 1
            continue
        break
    return _ReceiverChain(segments, start_index)


def _collect_implicit_with_chain(
    tokens: Sequence[VbaToken], end_index: int
) -> list[_ReceiverChainSegment] | None:
    if end_index < 0 or _is_boundary(tokens[end_index]):
        return []
    segments: list[_ReceiverChainSegment] = []
    i = end_index
    pending_has_arguments = False
    while True:
        if i >= 0 and _is_boundary(tokens[i]):
            return None
        if i >= 0 and tokens[i].raw_text == ")":
            open_index = _match_paren_left(tokens, i)
            if open_index < 0:
                return None
            pending_has_arguments = True
            i = open_index - 1
            continue
        if i < 0 or not is_ident_like(tokens[i]):
            return None
        segments.insert(0, _ReceiverChainSegment(_word(tokens[i]), pending_has_arguments))
        pending_has_arguments = False
        i -= 1
        if i >= 0 and tokens[i].raw_text == ".":
            prior = i - 1
            if prior < 0 or _is_boundary(tokens[prior]):
                return segments
            i = prior
            continue
        return None


def _match_paren_left(tokens: Sequence[VbaToken], close_index: int) -> int:
    """Index of the '(' matching the ')' at ``close_index``, or -1."""
    depth = 0
    for i in range(close_index, -1, -1):
        t = tokens[i].raw_text
        if t == ")":
            depth += 1
        elif t == "(":
            depth -= 1
            if depth == 0:
                return i
    return -1


# -- root resolution -------------------------------------------------------


def _resolve_root(
    root: str, source: str, offset: int, ctx: MemberCompletionContext
) -> str | None:
    model = ctx.model
    lower = root.lower()

    if lower == "me":
        project_key = _project_key_for_type_name(ctx.me_project_type, ctx)
        if ctx.me_type:
            return _combined_type_key(project_key, ctx.me_type) if project_key else ctx.me_type
        return _project_type_key(project_key) if project_key else None

    declared = _find_declared_binding(source, offset, root, ctx)
    if declared is not None:
        if declared.as_type:
            declared_object_type = _resolve_declared_object_type(declared.as_type, ctx, model)
            if declared_object_type:
                return declared_object_type
            if not _is_generic_object_declaration(declared.as_type):
                return None
        return (
            None
            if ctx.allow_set_assignment_refinement is False
            else _find_set_assigned_object_type(source, offset, root, ctx)
        )

    project_surface = _project_class_members_by_name(ctx).get(lower)
    project_key = lower if project_surface is not None else None
    runtime_object = resolve_runtime_object(root)
    if runtime_object is not None:
        return runtime_object.get("type")
    as_global = resolve_host_global(root, model)
    if as_global:
        return _combined_type_key(project_key, as_global) if project_key else as_global
    as_code = (ctx.code_names or {}).get(lower)
    if as_code:
        return _combined_type_key(project_key, as_code) if project_key else as_code
    if project_surface is not None and project_surface.kind == "standardModule":
        return _project_type_key(lower)
    return (
        None
        if ctx.allow_set_assignment_refinement is False
        else _find_set_assigned_object_type(source, offset, root, ctx)
    )


# -- With-scope ------------------------------------------------------------


@dataclass(slots=True)
class _ActiveWithExpression:
    tokens: list[VbaToken]
    slice_start: int


def _with_receiver_type_at(
    source: str, offset: int, ctx: MemberCompletionContext
) -> str | None:
    current_type: str | None = None
    for expression in _active_with_expressions_at(source, offset, ctx):
        explicit_type = _receiver_type_from_expression_tokens(
            expression.tokens, source, expression.slice_start, ctx
        )
        if explicit_type:
            current_type = explicit_type
            continue
        implicit_chain = _collect_implicit_with_chain(
            expression.tokens, len(expression.tokens) - 1
        )
        if implicit_chain is None:
            return None
        current_type = _receiver_type_from_implicit_with_chain(
            current_type, implicit_chain, ctx
        )
        if not current_type:
            return None
    return current_type


def _active_with_expressions_at(
    source: str, offset: int, ctx: MemberCompletionContext
) -> list[_ActiveWithExpression]:
    scan_text, scan_slice_start = _active_with_scan_window(source, offset, ctx)
    stack: list[_ActiveWithExpression] = []
    statement: list[VbaToken] = []

    def flush() -> None:
        nonlocal statement
        _process_with_stack_statement(statement, stack, scan_slice_start)
        statement = []

    for token in tokenize(scan_text):
        if token.kind is TokenKind.COMMENT:
            continue
        if _is_boundary(token):
            flush()
            continue
        statement.append(token)
    flush()
    return stack


def _active_with_scan_window(
    source: str, offset: int, ctx: MemberCompletionContext
) -> tuple[str, int]:
    safe_offset = max(0, offset)
    module = ctx.parsed_module if ctx.parsed_module is not None else parse_module(source)
    enclosing = _enclosing_procedure(module, safe_offset)
    if enclosing is None:
        return (source[:safe_offset], 0)
    return (source[enclosing.span.start : safe_offset], enclosing.span.start)


def _process_with_stack_statement(
    statement: Sequence[VbaToken], stack: list[_ActiveWithExpression], slice_start: int
) -> None:
    start = _statement_executable_start(statement)
    first = _at(statement, start)
    if first is None:
        return
    first_word = _word(first).lower()
    if first_word == "with":
        stack.append(
            _ActiveWithExpression(
                tokens=list(statement[start + 1 :]),
                slice_start=slice_start + first.start,
            )
        )
        return
    if first_word == "end" and _word(_at(statement, start + 1) or first).lower() == "with":
        if stack:
            stack.pop()


def _statement_executable_start(statement: Sequence[VbaToken]) -> int:
    if (
        len(statement) > 1
        and statement[0].kind is TokenKind.INTEGER_LITERAL
        and re.match(r"^\d+$", statement[0].raw_text)
    ):
        return 1
    if len(statement) > 2 and is_ident_like(statement[0]) and statement[1].raw_text == ":":
        return 2
    return 0


# -- member surface --------------------------------------------------------


def _member_surface_for_type(
    type_name: str, ctx: MemberCompletionContext
) -> _MemberSurface | None:
    union = _parse_union_type_key(type_name)
    if union is not None:
        surfaces = [
            surface
            for surface in (_member_surface_for_type(item, ctx) for item in union)
            if surface is not None
        ]
        if len(surfaces) == 0:
            return None
        return _MemberSurface(
            owner=" | ".join(_display_type_name(item) for item in union),
            members=_merge_completion_members(*[surface.members for surface in surfaces]),
            exhaustive=all(surface.exhaustive for surface in surfaces),
        )
    combined = _parse_combined_type_key(type_name)
    if combined is not None:
        project_key, host_type_name = combined
        project_type = _project_class_members_by_name(ctx).get(project_key)
        host_type = get_host_type(host_type_name, ctx.model)
        if project_type is None and host_type is None:
            return None
        return _MemberSurface(
            owner=project_type.name if project_type is not None else host_type_name,
            members=_merge_completion_members(
                _project_member_entries(project_type),
                _host_member_entries(get_host_members(host_type_name, ctx.model)),
            ),
            exhaustive=(
                _project_source_surface_complete_when_merged_with_host(project_type)
                and host_type is not None
                and host_type.get("exhaustive") is True
            ),
        )
    if type_name.startswith(_PROJECT_TYPE_PREFIX):
        project_type = _project_class_members_by_name(ctx).get(
            type_name[len(_PROJECT_TYPE_PREFIX) :]
        )
        if project_type is None:
            return None
        return _MemberSurface(
            owner=project_type.name,
            members=_project_member_entries(project_type),
            exhaustive=(
                project_type.exhaustive
                if project_type.exhaustive is not None
                else project_type.kind == "class"
            ),
        )
    runtime_object = resolve_runtime_object_type(type_name)
    if runtime_object is not None:
        return _MemberSurface(
            owner=runtime_object.get("name", type_name),
            members=[
                MemberCompletionEntry(name=m["name"], kind=m.get("kind", "property"))
                for m in (runtime_object.get("members") or [])
            ],
            exhaustive=runtime_object.get("exhaustive") is True,
        )
    host_type = get_host_type(type_name, ctx.model)
    return _MemberSurface(
        owner=type_name,
        members=_host_member_entries(get_host_members(type_name, ctx.model)),
        exhaustive=host_type is not None and host_type.get("exhaustive") is True,
    )


def _host_member_entries(members: Sequence[HostMember]) -> list[MemberCompletionEntry]:
    return [
        MemberCompletionEntry(name=m["name"], kind=m.get("kind", "property")) for m in members
    ]


def _project_member_entries(
    project_type: VbaProjectClassMembers | None,
) -> list[MemberCompletionEntry]:
    if project_type is None:
        return []
    return [
        MemberCompletionEntry(name=m.name, kind=m.kind) for m in project_type.members
    ]


# -- member-return chaining ------------------------------------------------


def _resolve_any_member_return_type(
    owner_type: str, member_name: str, ctx: MemberCompletionContext
) -> _ResolvedMemberReturn | None:
    union = _parse_union_type_key(owner_type)
    if union is not None:
        resolved = [
            item
            for item in (
                _resolve_any_member_return_type(item, member_name, ctx) for item in union
            )
            if item is not None
        ]
        if len(resolved) == 0:
            return None
        return _ResolvedMemberReturn(
            type=_type_key_for([item.type for item in resolved]),
            kind="method" if all(item.kind == "method" for item in resolved) else "property",
        )
    combined = _parse_combined_type_key(owner_type)
    if combined is not None:
        project_key, host_type_name = combined
        project_type = _project_class_members_by_name(ctx).get(project_key)
        project_member = _project_member_by_name(project_type, member_name)
        if project_member is not None and project_member.returns:
            type_ = _resolve_declared_object_type(project_member.returns, ctx, ctx.model)
            return _ResolvedMemberReturn(type=type_, kind=project_member.kind) if type_ else None
        return _host_member_return(host_type_name, member_name, ctx.model)
    if not owner_type.startswith(_PROJECT_TYPE_PREFIX):
        runtime_object = resolve_runtime_object_type(owner_type)
        if runtime_object is not None:
            lower = member_name.lower()
            member = next(
                (m for m in (runtime_object.get("members") or []) if m["name"].lower() == lower),
                None,
            )
            if member is not None and member.get("returns"):
                return _ResolvedMemberReturn(
                    type=member["returns"], kind=member.get("kind", "property")
                )
            return None
        return _host_member_return(owner_type, member_name, ctx.model)
    project_type = _project_class_members_by_name(ctx).get(
        owner_type[len(_PROJECT_TYPE_PREFIX) :]
    )
    project_member = _project_member_by_name(project_type, member_name)
    if project_member is None or not project_member.returns:
        return None
    type_ = _resolve_declared_object_type(project_member.returns, ctx, ctx.model)
    return _ResolvedMemberReturn(type=type_, kind=project_member.kind) if type_ else None


def _apply_default_member_return_type(
    type_name: str | None, has_arguments: bool, ctx: MemberCompletionContext
) -> str | None:
    if not type_name or not has_arguments:
        return type_name
    union = _parse_union_type_key(type_name)
    if union is not None:
        return _type_key_for(
            [
                (_host_member_return(item, "Item", ctx.model) or _ResolvedMemberReturn(item, "")).type
                for item in union
            ]
        )
    resolved = _host_member_return(type_name, "Item", ctx.model)
    return resolved.type if resolved is not None else type_name


def _host_member_return(
    owner_type: str, member_name: str, model: HostObjectModel | None
) -> _ResolvedMemberReturn | None:
    lower = member_name.lower()
    member = next(
        (m for m in get_host_members(owner_type, model) if m["name"].lower() == lower), None
    )
    if member is None:
        return None
    if member.get("returns"):
        return _ResolvedMemberReturn(type=member["returns"], kind=member.get("kind", "property"))
    returns_any_of = member.get("returnsAnyOf")
    if returns_any_of:
        return _ResolvedMemberReturn(
            type=_type_key_for(returns_any_of), kind=member.get("kind", "property")
        )
    return None


# -- declared-type resolution ----------------------------------------------


def _resolve_declared_object_type(
    declared_type: str, ctx: MemberCompletionContext, model: HostObjectModel | None
) -> str | None:
    host = resolve_host_alias(declared_type, model)
    if host:
        return host
    key = _project_key_for_type_name(declared_type, ctx)
    if key:
        code_name_host = (ctx.code_names or {}).get(key)
        return _combined_type_key(key, code_name_host) if code_name_host else _project_type_key(key)
    return None


def _project_key_for_type_name(
    type_name: str | None, ctx: MemberCompletionContext
) -> str | None:
    if not type_name:
        return None
    simple = _simple_type_name(type_name)
    key = simple.lower() if simple else None
    if not key:
        return None
    project_type = _project_class_members_by_name(ctx).get(key)
    return key if project_type is not None and project_type.kind != "standardModule" else None


def _simple_type_name(type_text: str) -> str | None:
    trimmed = type_text.strip()
    return trimmed if IDENT_RE.match(trimmed) else None


def _project_type_key(lower_name: str) -> str:
    return f"{_PROJECT_TYPE_PREFIX}{lower_name}"


def _combined_type_key(project_key: str, host_type: str) -> str:
    return f"{_COMBINED_TYPE_PREFIX}{project_key}{_COMBINED_TYPE_SEPARATOR}{host_type}"


def _parse_combined_type_key(type_name: str) -> tuple[str, str] | None:
    if not type_name.startswith(_COMBINED_TYPE_PREFIX):
        return None
    body = type_name[len(_COMBINED_TYPE_PREFIX) :]
    sep = body.find(_COMBINED_TYPE_SEPARATOR)
    if sep < 1 or sep >= len(body) - 1:
        return None
    return (body[:sep], body[sep + 1 :])


def _parse_union_type_key(type_name: str) -> list[str] | None:
    if not type_name.startswith(_UNION_TYPE_PREFIX):
        return None
    parts = [
        item
        for item in type_name[len(_UNION_TYPE_PREFIX) :].split(_UNION_TYPE_SEPARATOR)
        if len(item) > 0
    ]
    return parts if len(parts) > 0 else None


def _type_key_for(types: Sequence[str]) -> str:
    out: list[str] = []
    seen: set[str] = set()
    for type_ in types:
        for item in _parse_union_type_key(type_) or [type_]:
            key = item.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(item)
    return out[0] if len(out) == 1 else f"{_UNION_TYPE_PREFIX}{_UNION_TYPE_SEPARATOR.join(out)}"


def _display_type_name(type_name: str) -> str:
    dot = type_name.rfind(".")
    return type_name[dot + 1 :] if dot >= 0 else type_name


def _merge_completion_members(
    *member_groups: Sequence[MemberCompletionEntry],
) -> list[MemberCompletionEntry]:
    out: list[MemberCompletionEntry] = []
    seen: set[str] = set()
    for members in member_groups:
        for member in members:
            key = member.name.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(member)
    return out


def _project_source_surface_complete_when_merged_with_host(
    project_type: VbaProjectClassMembers | None,
) -> bool:
    if project_type is None:
        return True
    if project_type.kind == "userform":
        return False
    return True


def _project_class_members_by_name(
    ctx: MemberCompletionContext,
) -> dict[str, VbaProjectClassMembers]:
    out: dict[str, VbaProjectClassMembers] = {}
    ambiguous: set[str] = set()
    for type_ in ctx.project_class_members or []:
        key = type_.name.lower()
        if key in ambiguous:
            continue
        if key in out:
            del out[key]
            ambiguous.add(key)
            continue
        out[key] = type_
    return out


def _project_member_by_name(
    project_type: VbaProjectClassMembers | None, member_name: str
) -> VbaProjectClassMember | None:
    if project_type is None:
        return None
    lower = member_name.lower()
    return next((m for m in project_type.members if m.name.lower() == lower), None)


def _is_generic_object_declaration(declared_type: str) -> bool:
    simple = _simple_type_name(declared_type)
    lower = simple.lower() if simple else None
    return lower == "object" or lower == "variant"


# -- Set-assignment refinement ---------------------------------------------


def _find_set_assigned_object_type(
    source: str, offset: int, name: str, ctx: MemberCompletionContext
) -> str | None:
    module = ctx.parsed_module if ctx.parsed_module is not None else parse_module(source)
    lower = name.lower()
    enclosing = _enclosing_procedure(module, offset)

    if enclosing is not None:
        hit = _latest_set_assignment_in_body(enclosing.body, source, offset, lower)
        if hit is not None:
            return _receiver_type_from_expression_tokens(hit.value_tokens, source, hit.offset, ctx)

    latest: _SetAssignment | None = None
    for member in module.members:
        if not is_leaf_statement(member) or member.span.end > offset:
            continue
        hit = _set_assignment(source, member)
        if hit is not None and hit.name.lower() == lower:
            latest = hit
    return (
        _receiver_type_from_expression_tokens(latest.value_tokens, source, latest.offset, ctx)
        if latest is not None
        else None
    )


def _latest_set_assignment_in_body(
    body: Sequence[BodyNode], source: str, offset: int, lower_name: str
) -> _SetAssignment | None:
    latest: _SetAssignment | None = None
    for node in body:
        if is_leaf_statement(node):
            if node.span.end > offset:
                continue
            hit = _set_assignment(source, node)
            if hit is not None and hit.name.lower() == lower_name:
                latest = hit
        else:
            child = getattr(node, "body", None)
            if isinstance(child, list):
                hit = _latest_set_assignment_in_body(child, source, offset, lower_name)
                if hit is not None:
                    latest = hit
    return latest


def _set_assignment(source: str, stmt: LeafStatementNode) -> _SetAssignment | None:
    tokens = [
        t
        for t in tokenize(source[stmt.span.start : stmt.span.end])
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]
    i = 0
    if (
        len(tokens) >= 2
        and tokens[0].kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD)
        and tokens[1].raw_text == ":"
    ):
        i = 2
    if i >= len(tokens) or tokens[i].raw_text.lower() != "set":
        return None
    name_token = _at(tokens, i + 1)
    if name_token is None or name_token.kind is not TokenKind.IDENTIFIER:
        return None
    equals = _at(tokens, i + 2)
    if equals is None or equals.kind is not TokenKind.OPERATOR or equals.raw_text != "=":
        return None
    return _SetAssignment(
        name=name_token.raw_text,
        value_tokens=list(tokens[i + 3 :]),
        offset=stmt.span.start,
    )


# -- declared bindings -----------------------------------------------------


def _find_declared_binding(
    source: str, offset: int, name: str, ctx: MemberCompletionContext
) -> _DeclaredBinding | None:
    module = ctx.parsed_module if ctx.parsed_module is not None else parse_module(source)
    lower = name.lower()
    enclosing = _enclosing_procedure(module, offset)

    if enclosing is not None:
        for param in enclosing.params:
            if param.name.lower() == lower:
                return _DeclaredBinding(as_type=param.as_type)
        local = _find_in_body(enclosing.body, lower)
        if local is not None:
            return local

    for mem in module.members:
        if isinstance(mem, VariableGroupNode):
            hit = _match_group(mem, lower)
            if hit is not None:
                return hit
    return None


def _find_in_body(body: Sequence[BodyNode], lower: str) -> _DeclaredBinding | None:
    for node in body:
        if isinstance(node, VariableGroupNode):
            hit = _match_group(node, lower)
            if hit is not None:
                return hit
        else:
            child = getattr(node, "body", None)
            if isinstance(child, list):
                hit = _find_in_body(child, lower)
                if hit is not None:
                    return hit
    return None


def _match_group(group: VariableGroupNode, lower: str) -> _DeclaredBinding | None:
    for decl in group.declarations:
        if decl.name.lower() == lower:
            return _DeclaredBinding(as_type=decl.as_type)
    return None


# -- AST helpers -----------------------------------------------------------


def _enclosing_procedure(module: ModuleNode, offset: int) -> ProcedureNode | None:
    for mem in module.members:
        if (
            isinstance(mem, ProcedureNode)
            and offset >= mem.span.start
            and offset <= mem.span.end
        ):
            return mem
    return None


# -- object-assignment type gate (for objectState M9 wiring) ---------------


def is_known_object_assignment_type(
    type_name: str | None, ctx: MemberCompletionContext
) -> bool:
    """True when a declared type names an object that Set-binds and supports
    members. Ported from resolveKnownObjectAssignmentType (typeInference.ts): the
    generic ``Object`` type, any host alias (Excel/Office) resolved through the
    host model, or an unambiguous project class/document/userform type qualify.
    ``Variant`` and the scalar types do not. Used to widen the host-free
    is_known_object_assignment_type so that e.g. ``Dim ws As Worksheet`` counts as
    an object variable."""
    if not type_name:
        return False
    normalized = normalize_type(type_name)
    if not normalized or normalized == "variant":
        return False
    if normalized == "object":
        return True
    if is_known_scalar_type(normalized):
        return False
    if resolve_host_alias(type_name, ctx.model) is not None:
        return True
    simple = _simple_type_name(type_name)
    if not simple:
        return False
    lower = simple.lower()
    matches = [
        project_type
        for project_type in (ctx.project_class_members or [])
        if project_type.kind not in ("userType", "standardModule")
        and project_type.name.lower() == lower
    ]
    return len(matches) == 1
