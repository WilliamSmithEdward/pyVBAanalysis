"""Member-access type/surface resolver (memberAccess.ts).

Given VBA source and an offset just after a member-access dot, this resolves the
type of the receiver expression and returns the verified member surface available
on it. The diagnostics consume the exhaustive surface (member-not-found),
``resolve_exact_member_completion`` / ``resolve_member_completion_named`` for the one
member a call or assignment names, with its returns, writability and call
signature, and the project-surface lookups (private members, project types and
class members at a receiver). The object-assignment type resolution of
typeInference.ts lives here too, below the surfaces it reads.

Documentation rendering is not ported, and neither are the implicit control
members an editor passes for the form being edited (the diagnostics context never
carries them). The EXHAUSTIVE flag is never synthesized: host surfaces use the host
model's ``exhaustive`` flag and project surfaces use
``VbaProjectClassMembers.exhaustive`` verbatim.
"""

from __future__ import annotations

import re
from bisect import bisect_left, bisect_right
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, replace
from typing import cast

from ..host import (
    HostObjectModel,
    get_host_members,
    get_host_type,
    resolve_host_alias,
    resolve_host_global,
)
from ..host.host_model import (
    HostMember,
    get_host_enum_members,
    resolve_host_enum,
    resolve_host_global_member,
    resolve_host_member_signature,
)
from ..host.msforms import VBA_USERFORM_TYPE, msforms_control_members, resolve_msforms_type_name
from ..host.type_extensibility import host_type_resolves_when_compiling
from ..identity_cache import IdentityLru
from ..lexer.token_helpers import is_ident_like, is_identifier
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize
from ..parser.nodes import (
    BodyNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    ProcKind,
    VariableGroupNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ..parser.parse_module import parse_module
from ..runtime import resolve_runtime_object, resolve_runtime_object_type, resolve_vba_library_qualifier
from ..host.host_default_members import HOST_DEFAULT_MEMBERS
from ..symbols.symbol_model import (
    VbaProjectClassMember,
    VbaProjectClassMemberDefinition,
    VbaProjectClassMembers,
    VbaSymbolAttribute,
    is_access_designer_class,
)
from ..types.type_names import is_known_scalar_type, normalize_type
from .cursor_context import completion_significant_tokens

_PROJECT_TYPE_PREFIX = "project:"
# Receiver key for a host enumeration used as a qualifier: `XlAxisType.xlCategory`.
_HOST_ENUM_PREFIX = "hostEnum:"
# Receiver key for a VBA library enum or module of constants: `VbMsgBoxResult.vbYes`.
_VBA_LIBRARY_PREFIX = "vbaLibrary:"
_COMBINED_TYPE_PREFIX = "combined:"
_COMBINED_TYPE_SEPARATOR = "|"
_UNION_TYPE_PREFIX = "union:"
_UNION_TYPE_SEPARATOR = "|"
_TRAILING_EMPTY_PARENS_RE = re.compile(r"\s*\(\s*\)\s*$")
_AS_OBJECT_SIGNATURE_RE = re.compile(r"\bAs Object\s*$", re.IGNORECASE)


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
    # Per-pass memo of the active `With` stack, keyed by enclosing procedure
    # start. Callers resolving many references against one unchanging source pass
    # a fresh dict; the scan is then paid once per procedure rather than once per
    # leading-dot member. Must be discarded whenever the source changes.
    with_scan_cache: dict[int, _WithScanIndex] | None = None
    # Receiver-chain prefix results for one analysis pass, keyed by the chain's
    # root token offset and the number of segments resolved (XLIDE issue #135).
    # The caller owns its lifetime: one source, one pass.
    receiver_type_cache: dict[tuple[int, int], str | None] | None = None
    # Receiver chains already collected, keyed by the dot token's offset (#135).
    receiver_chain_cache: dict[int, _ReceiverChain] | None = None
    # Upstream's memberSurfaceCache (XLIDE issue #139) has no field here: the port
    # has memoized surfaces per (project types, model) since before, in
    # _SURFACES_CACHE below, which serves every context of one pass.


@dataclass(frozen=True, slots=True)
class MemberCompletionEntry:
    """One member of a resolved surface. Mirrors XLIDE's CompletionMemberSource: most
    diagnostics read only ``name``/``kind``, the assignment-type rule reads
    ``returns``/``writable``/``write_type`` from the exact resolved member, and the
    member-call rules read its ``signature``."""

    name: str
    kind: str
    returns: str | None = None
    writable: bool | None = None
    write_type: str | None = None
    # The verified call signature, when the source or the host metadata has one.
    signature: str | None = None
    # The setters a project property declares (XLIDE issue #107); None for host
    # members.
    let_accessor: bool | None = None
    set_accessor: bool | None = None
    # The type a host property declares, when it is not a chainable object.
    declared_type: str | None = None
    # The read/write contract the type library states for a host property.
    access: str | None = None
    # A user-defined type's field that holds an array (XLIDE issue #417).
    is_array: bool | None = None
    # Qualified type the member belongs to. Filled in when a member is resolved
    # (resolve_member_completion_named and resolve_member_completions); a raw
    # surface entry carries "".
    owner: str = ""
    # True when the owner member surface is complete enough to prove absence.
    surface_exhaustive: bool | None = None
    # Source declaration locations, when this member comes from project code.
    definitions: Sequence[VbaProjectClassMemberDefinition] | None = None
    # True when exported source marks this member as the VBA default member.
    default_member: bool | None = None
    # A project method declared as a Sub, which gives no value (XLIDE issue #414).
    sub: bool | None = None
    # How many parameters a project property's Let declares, the value's
    # included (XLIDE issue #414).
    let_param_count: int | None = None
    # What a project class member is known to hold or return: "nothing", "empty"
    # or "scalar" (XLIDE issue #414).
    known_value: str | None = None
    # Exported attribute lines attached to this member.
    attributes: Sequence[VbaSymbolAttribute] | None = None
    # Marked hidden in the type library: resolved, but never offered.
    hidden: bool | None = None


# Upstream's name for one resolved member.
MemberCompletion = MemberCompletionEntry


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


ReceiverChain = _ReceiverChain


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
    # The members by lowercased name, built the first time one is looked up: a host
    # type has hundreds of members and a module asks for one name per reference.
    by_lower_name: dict[str, MemberCompletionEntry] | None = None


@dataclass(frozen=True, slots=True)
class ExhaustiveMemberSurface:
    """An exhaustive surface reduced to what a member-existence check needs."""

    owner: str
    surface: _MemberSurface

    def has_member(self, member_name: str) -> bool:
        return _surface_member_named(self.surface, member_name) is not None


def _surface_member_named(surface: _MemberSurface, member_name: str) -> MemberCompletionEntry | None:
    by_name = surface.by_lower_name
    if by_name is None:
        by_name = {}
        # First occurrence wins, as the linear search it replaces did.
        for member in surface.members:
            by_name.setdefault(member.name.lower(), member)
        surface.by_lower_name = by_name
    return by_name.get(member_name.lower())


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
        owner=surface.owner,
        members=[
            _completion_from_surface_member(current_type, surface, member, ctx)
            for member in surface.members
        ],
        exhaustive=surface.exhaustive,
    )


def resolve_exact_member_completion(
    source: str, member_name: str, member_end_offset: int, ctx: MemberCompletionContext | None = None
) -> MemberCompletionEntry | None:
    """Port of resolveExactMemberCompletion / resolveMemberCompletionNamed: the single
    resolved member named ``member_name`` whose access dot ends before
    ``member_end_offset``, carrying its ``returns``/``writable``/``write_type``.

    Returns None when the receiver does not resolve or the member is absent from the
    surface. The assignment-type rule uses the member's writable/write_type to decide
    read-only and value-type compatibility (the no-FP gate: an unresolved member or a
    member whose writability is unknown yields no diagnostic)."""
    return resolve_member_completion_named(source, member_end_offset, member_name, ctx)


@dataclass(frozen=True, slots=True)
class _SurfaceHit:
    current_type: str
    surface: _MemberSurface
    typed_prefix: str


def _member_surface_at_dot(
    source: str,
    offset: int,
    ctx: MemberCompletionContext,
    prefix_tokens: Sequence[VbaToken] | None = None,
) -> _SurfaceHit | None:
    # Keep newline tokens: they mark statement boundaries so a dangling
    # member-access dot on a previous line is not merged into this chain.
    tokens = prefix_tokens if prefix_tokens is not None else _prefix_significant_tokens(source, offset, ctx)
    if len(tokens) == 0:
        return None
    i = len(tokens) - 1
    typed_prefix = ""
    if is_ident_like(tokens[i]) and i > 0 and tokens[i - 1].raw_text == ".":
        typed_prefix = tokens[i].raw_text
        i -= 1
    if i < 0 or tokens[i].raw_text != ".":
        return None
    current_type = _receiver_type_from_tokens(tokens, i, source, offset, ctx)
    if not current_type:
        return None
    surface = _member_surface_for_type(current_type, ctx)
    if surface is None:
        return None
    return _SurfaceHit(current_type, surface, typed_prefix)


def _completion_from_surface_member(
    current_type: str,
    surface: _MemberSurface,
    member: MemberCompletionEntry,
    ctx: MemberCompletionContext,
) -> MemberCompletionEntry:
    """A surface member as a resolved completion: its owner, the surface's
    exhaustiveness, and its call signature. Documentation is not ported."""
    signature = (
        member.signature
        if member.signature is not None
        else _signature_for_member(current_type, member.name, ctx)
    )
    return replace(
        member, signature=signature, owner=surface.owner, surface_exhaustive=surface.exhaustive
    )


def resolve_member_completions(
    source: str, offset: int, ctx: MemberCompletionContext | None = None
) -> list[MemberCompletionEntry]:
    """The member completions available at ``offset``: the surface's members whose
    names start with the typed prefix, hidden ones left out (XLIDE issue #56)."""
    ctx = ctx if ctx is not None else MemberCompletionContext()
    hit = _member_surface_at_dot(source, offset, ctx)
    if hit is None:
        return []
    lower_prefix = hit.typed_prefix.lower()
    return [
        _completion_from_surface_member(hit.current_type, hit.surface, member, ctx)
        for member in hit.surface.members
        if member.name.lower().startswith(lower_prefix) and not member.hidden
    ]


def resolve_member_completion_named(
    source: str, offset: int, member_name: str, ctx: MemberCompletionContext | None = None
) -> MemberCompletionEntry | None:
    """The single member named ``member_name`` at ``offset``, without building rows
    for the whole member surface."""
    ctx = ctx if ctx is not None else MemberCompletionContext()
    hit = _member_surface_at_dot(source, offset, ctx)
    if hit is None:
        return None
    member = _surface_member_named(hit.surface, member_name)
    return (
        _completion_from_surface_member(hit.current_type, hit.surface, member, ctx)
        if member is not None
        else None
    )


def resolve_host_member_kind_at(
    source: str, offset: int, member_name: str, ctx: MemberCompletionContext | None = None
) -> str | None:
    """The kind of the HOST member named ``member_name`` ending at ``offset``, or
    None when the receiver is not a host object or carries no such member."""
    ctx = ctx if ctx is not None else MemberCompletionContext()
    hit = _member_surface_at_dot(source, offset, ctx)
    if hit is None or not any(
        get_host_type(type_, ctx.model) for type_ in _host_receiver_types_of(hit.current_type)
    ):
        return None
    lower_name = member_name.lower()
    member = next((m for m in hit.surface.members if m.name.lower() == lower_name), None)
    return member.kind if member is not None else None


def _host_receiver_types_of(receiver_type: str) -> list[str]:
    """The host types a receiver key denotes (XLIDE issue #44)."""
    union = _parse_union_type_key(receiver_type)
    if union is not None:
        return [host for item in union for host in _host_receiver_types_of(item)]
    combined = _parse_combined_type_key(receiver_type)
    if combined is not None:
        return [combined[1]]
    return [] if receiver_type.startswith(_PROJECT_TYPE_PREFIX) else [receiver_type]


def resolve_member_definitions_at(
    source: str,
    offset: int,
    member_name: str,
    ctx: MemberCompletionContext | None = None,
    prefix_tokens: Sequence[VbaToken] | None = None,
) -> Sequence[VbaProjectClassMemberDefinition]:
    """The source definition locations of the member named ``member_name`` ending
    at ``offset``. Bails on a cheap character scan when no member-access dot
    precedes the name."""
    ctx = ctx if ctx is not None else MemberCompletionContext()
    safe_offset = max(0, min(offset, len(source)))
    if not preceded_by_member_access_dot(source, safe_offset - len(member_name)):
        return []
    # Only trust supplied tokens that end exactly with the member name.
    last = prefix_tokens[-1] if prefix_tokens else None
    tokens = (
        prefix_tokens
        if last is not None
        and last.end == safe_offset
        and last.raw_text.lower() == member_name.lower()
        else None
    )
    hit = _member_surface_at_dot(source, safe_offset, ctx, tokens)
    if hit is None:
        return []
    lower_name = member_name.lower()
    member = next((m for m in hit.surface.members if m.name.lower() == lower_name), None)
    return (member.definitions if member is not None else None) or []


def preceded_by_member_access_dot(source: str, name_start: int) -> bool:
    """True when the identifier starting at ``name_start`` is preceded by a
    member-access dot, allowing for whitespace and `_` line continuations."""
    i = name_start - 1
    while True:
        while i >= 0 and source[i] in (" ", "\t"):
            i -= 1
        if i < 0:
            return False
        ch = source[i]
        if ch == ".":
            return True
        if ch in ("\n", "\r"):
            if ch == "\n" and i > 0 and source[i - 1] == "\r":
                i -= 1
            i -= 1
            while i >= 0 and source[i] in (" ", "\t"):
                i -= 1
            if i < 0 or source[i] != "_":
                return False
            i -= 1
            continue
        return False


def _project_key_of_receiver(current_type: str) -> str | None:
    combined = _parse_combined_type_key(current_type)
    if combined is not None:
        return combined[0]
    if current_type.startswith(_PROJECT_TYPE_PREFIX):
        return current_type[len(_PROJECT_TYPE_PREFIX) :]
    return None


def private_member_owner_at(
    source: str, offset: int, member_name: str, ctx: MemberCompletionContext | None = None
) -> str | None:
    """The owner to name when a reference reaches a Private member of a project
    module through an object (`Sheet1.Secret()`, `Me.Secret()`), where the rest of
    the surface cannot prove absence: VBA refuses each, "Method or data member not
    found" (XLIDE issue #219). None when the member is not Private there, or the
    surface has a public member of that name."""
    ctx = ctx if ctx is not None else MemberCompletionContext()
    current_type = resolve_receiver_type_at(source, offset, ctx)
    if not current_type:
        return None
    project_key = _project_key_of_receiver(current_type)
    project_type = _project_class_members_by_name(ctx).get(project_key) if project_key else None
    lower = member_name.lower()
    if project_type is None or not any(
        name.lower() == lower for name in (project_type.private_members or [])
    ):
        return None
    surface = _member_surface_for_type(current_type, ctx)
    return None if surface is not None and _surface_member_named(surface, member_name) else project_type.name


def project_type_at(
    source: str, offset: int, ctx: MemberCompletionContext | None = None
) -> VbaProjectClassMembers | None:
    """The project type (class, form, document) a receiver resolves to, if any."""
    ctx = ctx if ctx is not None else MemberCompletionContext()
    current_type = resolve_receiver_type_at(source, offset, ctx)
    if not current_type:
        return None
    project_key = _project_key_of_receiver(current_type)
    return _project_class_members_by_name(ctx).get(project_key) if project_key else None


def project_class_member_at(
    source: str, offset: int, member_name: str, ctx: MemberCompletionContext | None = None
) -> VbaProjectClassMember | None:
    """The member of a project class module a reference reaches, if the receiver
    is one."""
    project_type = project_type_at(source, offset, ctx)
    if project_type is None or project_type.kind != "class":
        return None
    lower = member_name.lower()
    return next((m for m in project_type.members if m.name.lower() == lower), None)


def resolve_exhaustive_member_surface_at(
    source: str, offset: int, ctx: MemberCompletionContext | None = None
) -> ExhaustiveMemberSurface | None:
    """The member surface of the receiver ending at ``offset`` when the surface can
    prove a member absent, without building a completion row for every member
    (XLIDE issue #139). The member-not-found rules ask this for every dot in a
    module and only ever test one name against it."""
    ctx = ctx if ctx is not None else MemberCompletionContext()
    current_type = resolve_receiver_type_at(source, offset, ctx)
    if not current_type:
        return None
    surface = _member_surface_for_type(current_type, ctx)
    if surface is None or not surface.exhaustive:
        return None
    return ExhaustiveMemberSurface(surface.owner, surface)


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
            # Receiver chains never cross a logical statement boundary and every
            # consumer walks backward stopping at one, so the prefix starts at
            # the previous newline token instead of copying the whole module
            # prefix per lookup (O(offset) copies dominated big-module passes).
            # Line continuations are trivia, not newline tokens, so a continued
            # statement stays intact.
            return list(shared[_line_start_indexes(shared)[found] : found + 1])
    return completion_significant_tokens(source, offset)


# Where _prefix_significant_tokens starts its slice, for each token of a shared
# stream: the last newline token before it, kept as the explicit boundary, or 0.
# Found once per stream rather than by walking back per lookup, which was
# quadratic in the length of one long statement.
_LINE_START_INDEX_CACHE = IdentityLru(capacity=4)


def _line_start_indexes(shared: Sequence[VbaToken]) -> list[int]:
    cached = _LINE_START_INDEX_CACHE.get(shared)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    starts: list[int] = []
    last_newline = 0
    for i, token in enumerate(shared):
        starts.append(last_newline)
        if token.kind is TokenKind.NEWLINE:
            last_newline = i
    return _LINE_START_INDEX_CACHE.put(starts, shared)  # type: ignore[no-any-return]


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
    # A dot whose chain is the previous dot's plus one member takes that dot's
    # chain and adds the member, instead of walking the whole chain back again
    # (XLIDE issue #135: a 4,000-member chain took 2.4 s, each dot re-walking it).
    chain = _chain_extended_from_previous_dot(tokens, dot_index, ctx)
    if chain is None:
        chain = _collect_receiver_chain_with_start(tokens, dot_index - 1)
    if chain is not None and ctx.receiver_chain_cache is not None:
        ctx.receiver_chain_cache[tokens[dot_index].start] = chain
    # The chain's root resolves the same at every dot of one statement, so the
    # type walk resumes from the longest prefix already resolved.
    cache_base: int | None = None
    if chain is not None and ctx.receiver_type_cache is not None:
        root_token = _at(tokens, chain.start_index)
        cache_base = root_token.start if root_token is not None else None
    explicit_receiver = _receiver_type_from_chain(
        chain.segments if chain is not None else [], source, offset, ctx, cache_base
    )
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


def is_explicit_element_accessor(name: str) -> bool:
    """Members whose declared return IS the already-resolved element/result: the
    default member (Item/_Default) and the creation method Add. A call to one of
    these must not be element-indexed again, or a collection whose element is
    itself a collection (e.g. SparklineGroups.Item(1)) over-resolves one level."""
    lower = name.lower()
    return lower == "item" or lower == "_default" or lower == "add"


def _receiver_type_from_implicit_with_chain(
    with_type: str | None,
    chain: list[_ReceiverChainSegment],
    ctx: MemberCompletionContext,
) -> str | None:
    current_type = with_type
    for segment in chain:
        if not current_type:
            return None
        current_type = _advance_receiver_type(current_type, segment, ctx)
    return current_type


def _advance_receiver_type(
    current_type: str, segment: _ReceiverChainSegment, ctx: MemberCompletionContext
) -> str | None:
    """The type one member further along a receiver chain, or None when the member
    does not resolve."""
    resolved = _resolve_any_member_return_type(current_type, segment.name, ctx)
    if resolved is None:
        return None
    # A member called with arguments indexes into its return type; when that
    # type is a host collection, _apply_default_member_return_type resolves the
    # element (and no-ops otherwise). This holds for method-kind accessors too
    # (e.g. ws.ChartObjects(1).Chart), so it must not be gated on kind. But
    # Item/_Default/Add already return the resolved element/result, so they are
    # not re-indexed (avoids over-resolving SparklineGroups.Item(1) one level).
    # A member that takes an argument of its own is what it returns:
    # Shapes.Range(Array("A")) is a ShapeRange, not a Shape (XLIDE issue #197).
    return _apply_default_member_return_type(
        resolved.type,
        segment.has_arguments
        and not is_explicit_element_accessor(segment.name)
        and not member_takes_own_arguments(_signature_for_member(current_type, segment.name, ctx)),
        ctx,
    )


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
    # After a name or another list the parentheses hold arguments, not a grouped
    # receiver: `k.Wrap(r).Caption` with k late-bound is Wrap's result, whatever r
    # is (XLIDE issue #594).
    before = _at(tokens, open_index - 1)
    if before is not None and (is_ident_like(before) or before.raw_text in (")", "]")):
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
    cache_base: int | None = None,
) -> str | None:
    if len(chain) == 0:
        return None
    # Prefix results of this chain, keyed by the root token's offset and the
    # number of segments resolved: the longest cached prefix is where the walk
    # resumes, and every prefix reached is stored for the next dot.
    cache: dict[tuple[int, int], str | None] | None = None
    base = 0
    if cache_base is not None:
        cache = ctx.receiver_type_cache
        base = cache_base
    resume_at = 0
    current_type: str | None = None
    if cache is not None:
        for s in range(len(chain), 0, -1):
            key = (base, s)
            if key in cache:
                current_type = cache[key]
                resume_at = s
                break
    if resume_at == 0:
        root = chain[0]
        root_type = _resolve_root(root.name, source, offset, ctx)
        if not root_type:
            if cache is not None:
                cache[(base, 1)] = None
            return None
        current_type = _apply_default_member_return_type(root_type, root.has_arguments, ctx)
        if cache is not None:
            cache[(base, 1)] = current_type
        resume_at = 1
    s = resume_at
    while s < len(chain) and current_type:
        # An unresolved member stores None, which also ends the walk.
        current_type = _advance_receiver_type(current_type, chain[s], ctx)
        if cache is not None:
            cache[(base, s + 1)] = current_type
        s += 1
    return current_type


# -- chain collection ------------------------------------------------------


def _chain_extended_from_previous_dot(
    tokens: Sequence[VbaToken], dot_index: int, ctx: MemberCompletionContext
) -> _ReceiverChain | None:
    """The chain for the dot at ``dot_index`` when the token before it is a plain
    member name that follows an already-resolved dot: that dot's cached chain plus
    this member. Any other shape (a root, a boundary) is collected the long way."""
    cache = ctx.receiver_chain_cache
    if cache is None:
        return None
    member = _at(tokens, dot_index - 1)
    if member is not None and is_ident_like(member):
        previous_dot = _at(tokens, dot_index - 2)
        if previous_dot is None or previous_dot.raw_text != ".":
            return None
        previous = cache.get(previous_dot.start)
        if previous is None:
            return None
        return _ReceiverChain(
            [*previous.segments, _ReceiverChainSegment(_word(member), False)], previous.start_index
        )
    # Port-only: upstream extends only a plain member and collects `name(args)` the
    # long way, so a chain of calls such as `.Offset(1, 0).Offset(1, 0)...` stayed
    # quadratic (a 1,000-link chain took 3.4 s). The segment ending at the `)` is
    # the one the long walk would collect there, so the chain is the same.
    if member is None or member.raw_text != ")":
        return None
    last = _receiver_segment_ending_at(tokens, dot_index - 1)
    if last is None:
        return None
    segment, name_index = last
    previous_dot = _at(tokens, name_index - 1)
    if previous_dot is None or previous_dot.raw_text != ".":
        return None
    previous = cache.get(previous_dot.start)
    if previous is None:
        return None
    return _ReceiverChain([*previous.segments, segment], previous.start_index)


def _collect_receiver_chain_with_start(
    tokens: Sequence[VbaToken], end_index: int
) -> _ReceiverChain | None:
    segments: list[_ReceiverChainSegment] = []
    i = end_index
    start_index = -1
    while True:
        last = _receiver_segment_ending_at(tokens, i)
        if last is None:
            return None
        segment, start_index = last
        segments.insert(0, segment)
        i = start_index - 1
        if i >= 0 and tokens[i].raw_text == ".":
            i -= 1
            continue
        break
    return _ReceiverChain(segments, start_index)


def _receiver_segment_ending_at(
    tokens: Sequence[VbaToken], end_index: int
) -> tuple[_ReceiverChainSegment, int] | None:
    """The receiver-chain segment ending at ``end_index``, a name with any
    argument lists after it, and the index of the name; None when the tokens
    there are not one."""
    i = end_index
    has_arguments = False
    while True:
        if i >= 0 and _is_boundary(tokens[i]):
            return None
        if i >= 0 and tokens[i].raw_text == ")":
            open_index = _match_paren_left(tokens, i)
            if open_index < 0:
                return None
            # Empty parens foo() are a call with no index, not collection indexing;
            # only a non-empty argument list resolves to an element (matches the
            # assignment-inference path's argument_tokens check).
            if open_index < i - 1:
                has_arguments = True
            i = open_index - 1
            continue
        if i < 0 or not is_ident_like(tokens[i]):
            return None
        return _ReceiverChainSegment(_word(tokens[i]), has_arguments), i


# `Me` and `Debug` are the VBA keywords that can terminate a receiver expression
# (`Me.`, `Debug.Print` inside a With, XLIDE issue #184); every other keyword
# before a dot (In, To, Then, ...) introduces a fresh expression, so the dot is a
# leading implicit-With member access.
_RECEIVER_TAIL_KEYWORDS: frozenset[str] = frozenset({"me", "debug"})


def precedes_leading_member_dot(token: VbaToken) -> bool:
    """True when ``token`` (the token immediately before a ``.``) means the dot is
    a LEADING implicit-With member-access dot rather than ``receiver.member``. A
    dot is explicit only when preceded by something that terminates a receiver
    expression: a plain identifier, a ``[Foo]`` foreign-name escape, ``Me``, or a
    closing ``)``/``]``. Anything else - a statement boundary, an operator
    (``=``, ``&``, ``+``, ...), ``(``/``,``, or an expression-introducing keyword
    (``In``, ``To``, ``Then``, ...) - starts a new expression where ``.member``
    binds to the active ``With`` block (e.g. ``For Each wb In .Workbooks``,
    ``Set x = .Foo``)."""
    if token.kind is TokenKind.IDENTIFIER or token.kind is TokenKind.BRACKETED_IDENTIFIER:
        return False
    if token.raw_text == ")" or token.raw_text == "]":
        return False
    if token.kind is TokenKind.KEYWORD and token.raw_text.lower() in _RECEIVER_TAIL_KEYWORDS:
        return False
    return True


def _collect_implicit_with_chain(
    tokens: Sequence[VbaToken], end_index: int
) -> list[_ReceiverChainSegment] | None:
    if end_index < 0 or precedes_leading_member_dot(tokens[end_index]):
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
            # Empty parens foo() are a call with no index, not collection indexing;
            # only a non-empty argument list resolves to an element (matches the
            # assignment-inference path's argument_tokens check).
            if open_index < i - 1:
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
            if prior < 0 or precedes_leading_member_dot(tokens[prior]):
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
    # A member of the host's hidden Global interface with a typed return is a
    # receiver too: Excel's Union(a, b) yields a Range, Word's RecentFiles a
    # RecentFiles (XLIDE #34). Ranked with the other host-injected names.
    global_member = resolve_host_global_member(root, model)
    as_global_member = global_member.get("returns") if global_member is not None else None
    if as_global_member:
        return _combined_type_key(project_key, as_global_member) if project_key else as_global_member
    # VBA's own enums and modules of constants reach their constants too:
    # `VbMsgBoxResult.vbYes`, `ColorConstants.vbRed`. VBA is first in every
    # project's references, so it has a name a host shares: `Constants.vbCrLf` in
    # Excel is VBA's module, not Excel's Constants enum.
    as_vba_library = resolve_vba_library_qualifier(root) if project_key is None else None
    if as_vba_library is not None and as_vba_library.constants is not None:
        return f"{_VBA_LIBRARY_PREFIX}{as_vba_library.name}"
    # An enum name reaches its own constants: `XlAxisType.xlCategory` is ordinary
    # VBA and is how a reader tells one library's xlNone from another's.
    as_enum = resolve_host_enum(root, model) if project_key is None else None
    if as_enum is not None:
        return f"{_HOST_ENUM_PREFIX}{as_enum.get('displayName', root)}"
    # A standard module's name reaches its members, and so does a class, form or
    # Enum name: forms carry a default instance, factory-style classes are
    # addressed by name as a matter of course, and `Corner.TopLeft` is ordinary
    # VBA. Misusing a class that is not predeclared is the diagnostics' concern.
    # A document module whose host type is unknown (no code name reached the
    # analyzer) still reaches its own code. Its surface is never exhaustive, so
    # only a Private member is provably out of reach.
    if project_surface is not None and project_surface.kind in (
        "standardModule",
        "class",
        "userform",
        "enum",
        "document",
    ):
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
    scan = _active_with_scan_window(source, offset, ctx)
    index = _with_scan_index(source, scan, ctx)
    # Resume from the last complete statement before `offset` rather than from
    # the top of the procedure, then finish the partial statement the offset
    # sits in. Same answer, paid once per procedure instead of once per dot.
    at = _last_boundary_at_or_before(index.boundaries, offset)
    stack = [] if at < 0 else list(index.stacks[at])
    statement: list[VbaToken] = []

    def flush() -> None:
        nonlocal statement
        _process_with_stack_statement(statement, stack, index.slice_start)
        statement = []

    for i in range(0 if at < 0 else index.resume_at[at], len(index.tokens)):
        token = index.tokens[i]
        # Boundaries are absolute; the fallback lexer numbers its tokens from the
        # start of the sliced procedure, so the window's own start is added back.
        if token.end + index.slice_start > offset:
            break
        if token.kind is TokenKind.COMMENT:
            continue
        if _is_boundary(token):
            flush()
            continue
        statement.append(token)
    flush()
    return stack


@dataclass(frozen=True, slots=True)
class _WithScanWindow:
    """The source range a `With` scan reads: the procedure enclosing the offset, or
    at module level everything before it, which has no procedure to key an index
    on (procedure_start is then -1)."""

    slice_start: int
    procedure_start: int
    window_end: int


@dataclass(frozen=True, slots=True)
class _WithScanIndex:
    """The active `With` stack after every complete statement of one procedure.

    Walking the procedure from its start for each leading-dot member is quadratic
    in the procedure's length: a 500-line With block took 18 seconds. Callers that
    resolve many references against one source pass a `with_scan_cache`, and then
    each procedure is walked once."""

    tokens: Sequence[VbaToken]
    slice_start: int
    # End offset of each complete statement, ascending.
    boundaries: list[int]
    # Stack after the statement ending at the same position in `boundaries`.
    stacks: list[list[_ActiveWithExpression]]
    # Token index to resume scanning from, per boundary.
    resume_at: list[int]


def _with_scan_index(source: str, scan: _WithScanWindow, ctx: MemberCompletionContext) -> _WithScanIndex:
    cache = ctx.with_scan_cache if scan.procedure_start >= 0 else None
    cached = cache.get(scan.procedure_start) if cache is not None else None
    if cached is not None:
        return cached
    tokens, slice_start = _with_scan_tokens(source, scan, ctx)
    boundaries: list[int] = []
    stacks: list[list[_ActiveWithExpression]] = []
    resume_at: list[int] = []
    stack: list[_ActiveWithExpression] = []
    statement: list[VbaToken] = []
    for i, token in enumerate(tokens):
        if token.kind is TokenKind.COMMENT:
            continue
        if not _is_boundary(token):
            statement.append(token)
            continue
        _process_with_stack_statement(statement, stack, slice_start)
        statement = []
        boundaries.append(token.end + slice_start)
        stacks.append(list(stack))
        resume_at.append(i + 1)
    index = _WithScanIndex(tokens, slice_start, boundaries, stacks, resume_at)
    if cache is not None:
        cache[scan.procedure_start] = index
    return index


def _last_boundary_at_or_before(boundaries: Sequence[int], offset: int) -> int:
    """Index of the last boundary at or before `offset`, or -1."""
    return bisect_right(boundaries, offset) - 1


def _with_scan_tokens(
    source: str, scan: _WithScanWindow, ctx: MemberCompletionContext
) -> tuple[Sequence[VbaToken], int]:
    """Tokens of the scan window, with the offset their positions count from.

    Re-lexing the enclosing procedure for every leading-dot member is quadratic in
    the procedure's length. When the caller holds the full-source stream, the
    window is a slice of it; its tokens are then at absolute offsets, so the
    slice start the callers add is zero."""
    shared = ctx.source_tokens
    if not shared:
        return tokenize(source[scan.slice_start : scan.window_end]), scan.slice_start
    first = bisect_left(shared, scan.slice_start, key=lambda token: token.start)
    last = bisect_right(shared, scan.window_end, lo=first, key=lambda token: token.end)
    return shared[first:last], 0


def _active_with_scan_window(
    source: str, offset: int, ctx: MemberCompletionContext
) -> _WithScanWindow:
    safe_offset = max(0, offset)
    module = ctx.parsed_module if ctx.parsed_module is not None else parse_module(source)
    enclosing = _enclosing_procedure(module, safe_offset)
    if enclosing is None:
        return _WithScanWindow(slice_start=0, procedure_start=-1, window_end=safe_offset)
    return _WithScanWindow(
        slice_start=enclosing.span.start,
        procedure_start=enclosing.span.start,
        window_end=enclosing.span.end,
    )


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


# A surface depends only on its type key, the project's type list and the host model,
# and the member-call rules ask for the same few receivers once per call. Rebuilding
# a project class's entry list per lookup was the dominant cost of a project pass, so
# surfaces are memoized per (project types, model) pair, whose objects live for the
# pass. Consumers never mutate a surface.
_SURFACES_CACHE = IdentityLru()


def _member_surface_for_type(
    type_name: str, ctx: MemberCompletionContext
) -> _MemberSurface | None:
    surfaces: dict[str, _MemberSurface | None] | None = _SURFACES_CACHE.get(
        ctx.project_class_members, ctx.model
    )
    if surfaces is None:
        surfaces = _SURFACES_CACHE.put({}, ctx.project_class_members, ctx.model)
    if type_name not in surfaces:
        surfaces[type_name] = _build_member_surface_for_type(type_name, ctx)
    return surfaces[type_name]


def _build_member_surface_for_type(
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
            # A union is what the library declares Object - ActiveSheet, a Sheets
            # item - so VBA binds its members when it runs and a name on none of
            # the parts is not a compile error (XLIDE issue #114: `ActiveSheet.asdf`
            # compiles). Never exhaustive.
            exhaustive=False,
        )
    if type_name.startswith(_VBA_LIBRARY_PREFIX):
        qualifier = resolve_vba_library_qualifier(type_name[len(_VBA_LIBRARY_PREFIX) :])
        if qualifier is None or qualifier.constants is None:
            return None
        # A constant declares a type rather than returning a chainable object, so
        # none of these members carries `returns`.
        return _MemberSurface(
            owner=qualifier.name,
            members=[
                MemberCompletionEntry(name=constant.get("name", ""), kind="property")
                for constant in qualifier.constants
            ],
            # Its members are exactly the type library's, so one can be proved absent.
            exhaustive=True,
        )
    if type_name.startswith(_HOST_ENUM_PREFIX):
        enum_name = type_name[len(_HOST_ENUM_PREFIX) :]
        constants = get_host_enum_members(enum_name, ctx.model)
        if not constants:
            return None
        return _MemberSurface(
            owner=enum_name,
            members=[
                MemberCompletionEntry(name=constant.get("name", ""), kind="property")
                for constant in constants
            ],
            # An enum's members are exactly its constants, so this surface can prove
            # one absent, unlike the object types, which never can.
            exhaustive=True,
        )
    combined = _parse_combined_type_key(type_name)
    if combined is not None:
        project_key, host_type_name = combined
        project_type = _project_class_members_by_name(ctx).get(project_key)
        host_type = get_host_type(host_type_name, ctx.model)
        # A form's `Me` can be combined:<form>|MSForms.UserForm, and MSForms is no
        # part of a host model, so the base surface comes from the forms metadata
        # when the host type names a forms class.
        forms_members = msforms_control_members(host_type_name)
        base_members = (
            forms_members if forms_members is not None else get_host_members(host_type_name, ctx.model)
        )
        if project_type is None and host_type is None and forms_members is None:
            return None
        return _MemberSurface(
            owner=project_type.name if project_type is not None else host_type_name,
            members=_merge_completion_members(
                _project_member_entries(project_type),
                _host_member_entries(base_members),
            ),
            # A form's own `Me` follows the rule its qualified name does (XLIDE #26):
            # the forms base plus an index-proven control list proves absence.
            # Other combined surfaces keep the host-exhaustive gate, with no
            # extensibility gate, unlike a bare host type: a document module such
            # as ThisWorkbook is the project's own class, and the VBE refuses a
            # member it lacks even though Excel's Workbook interface is extensible
            # (oracle case workbook_unknown_member_compile).
            # So does an Access form's or report's, whose list is its TypeInfo
            # stream's (XLIDE issue #206).
            exhaustive=(
                project_type is not None and project_type.exhaustive is True
                if forms_members is not None or is_access_designer_class(host_type_name)
                else _project_source_surface_complete_when_merged_with_host(project_type)
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
        designer_class = project_type.designer_class
        designer_members = (
            get_host_members(designer_class, ctx.model)
            if project_type.kind in ("userform", "document") and designer_class
            else []
        )
        if designer_class and len(designer_members) > 0:
            # An Access form or report is its own library's class, not a UserForm:
            # `Form_Orders.Requery` reaches Access.Form's members, and Show and Hide
            # are not among them. Exhaustive when the index holds the design's
            # member list (XLIDE issue #206). So is a VB6 form: `Form1.Cls` reaches
            # VB.Form (#358). A worksheet reaches Excel.Worksheet's, a closed
            # interface, so `Sheet1.Nope` is refused while compiling (#225).
            designer_type = get_host_type(designer_class, ctx.model)
            return _MemberSurface(
                owner=project_type.name,
                members=_merge_completion_members(
                    _project_member_entries(project_type),
                    _host_member_entries(designer_members),
                ),
                exhaustive=project_type.exhaustive is True
                and (
                    project_type.kind != "document"
                    or (
                        designer_type is not None
                        and designer_type.get("exhaustive") is True
                        and host_type_resolves_when_compiling(designer_class)
                    )
                ),
            )
        if project_type.kind == "userform":
            # A form IS an MSForms.UserForm wherever it is reached from, so a
            # qualified reference from another module gets Show, Hide and the rest
            # of the form surface beside the form's code and controls (XLIDE #22).
            # Exhaustive exactly when the index proved the control list: the merged
            # surface then proves absence the way the VBE's compiler does (#26).
            return _MemberSurface(
                owner=project_type.name,
                members=_merge_completion_members(
                    _project_member_entries(project_type),
                    _host_member_entries(msforms_control_members(VBA_USERFORM_TYPE) or []),
                ),
                exhaustive=project_type.exhaustive is True,
            )
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
                MemberCompletionEntry(
                    name=m["name"],
                    kind=m.get("kind", "property"),
                    returns=m.get("returns"),
                    writable=m.get("writable"),
                    write_type=m.get("writeType"),
                    signature=m.get("signature"),
                )
                for m in (runtime_object.get("members") or [])
            ],
            exhaustive=runtime_object.get("exhaustive") is True,
        )
    control_members = msforms_control_members(type_name)
    if control_members is not None:
        return _MemberSurface(
            owner=type_name,
            members=_host_member_entries(control_members),
            # Not exhaustive: this list is for offering members, and treating it as
            # complete would let absence become a diagnostic about form code.
            exhaustive=False,
        )
    host_type = get_host_type(type_name, ctx.model)
    return _MemberSurface(
        owner=type_name,
        members=_host_member_entries(get_host_members(type_name, ctx.model)),
        # A complete member list proves absence only where the type library says
        # VBA resolves against the interface while compiling. Most of Excel's
        # object model is extensible, so `Application.Match`, a worksheet function
        # on no interface at all, is ordinary VBA, and calling it absent reported
        # working code as an error.
        exhaustive=(
            host_type is not None
            and host_type.get("exhaustive") is True
            and host_type_resolves_when_compiling(type_name)
        ),
    )


def _host_member_entries(members: Sequence[HostMember]) -> list[MemberCompletionEntry]:
    # Host metadata never proves source-writability, so writable/write_type stay None
    # (the assignment-type rule treats writable === undefined as "cannot decide").
    return [
        MemberCompletionEntry(
            name=m["name"],
            kind=m.get("kind", "property"),
            returns=m.get("returns"),
            signature=m.get("signature"),
            declared_type=m.get("declaredType"),
            access=_host_member_access(m),
            hidden=m.get("hidden"),
        )
        for m in members
    ]


def _host_member_access(member: HostMember) -> str | None:
    """The read/write contract a host member's metadata states, if any."""
    access = cast("Mapping[str, object]", member).get("access")
    return access if isinstance(access, str) else None


def _project_member_entries(
    project_type: VbaProjectClassMembers | None,
) -> list[MemberCompletionEntry]:
    if project_type is None:
        return []
    return [
        MemberCompletionEntry(
            name=m.name,
            kind=m.kind,
            returns=m.returns,
            writable=m.writable,
            write_type=m.write_type,
            signature=m.signature,
            let_accessor=m.let_accessor,
            set_accessor=m.set_accessor,
            is_array=True if m.is_array else None,
            definitions=m.definitions,
            default_member=m.default_member,
            sub=m.sub,
            let_param_count=_let_params_of(m),
            known_value=m.known_value,
            attributes=m.attributes,
        )
        for m in project_type.members
    ]


def _let_params_of(member: VbaProjectClassMember) -> int | None:
    """How many parameters a project property's Let declares, if the member says."""
    params = (member.procedure_params or {}).get("propertyLet")
    return len(params) if params is not None else None


def _signature_for_member(
    type_name: str, member_name: str, ctx: MemberCompletionContext
) -> str | None:
    """The call signature of `member_name` on the receiver `type_name` when the
    surface member itself carries none: a union answers only when every type that
    has one agrees, and a host type also consults the model's signature table."""
    union = _parse_union_type_key(type_name)
    if union is not None:
        signatures = [
            signature
            for signature in (_signature_for_member(item, member_name, ctx) for item in union)
            if signature
        ]
        return signatures[0] if len(set(signatures)) == 1 else None
    combined = _parse_combined_type_key(type_name)
    if combined is not None:
        project_key, host_type_name = combined
        project_signature = _project_member_signature(project_key, member_name, ctx)
        if project_signature is not None:
            return project_signature
        return resolve_host_member_signature(host_type_name, member_name, ctx.model)
    if type_name.startswith(_PROJECT_TYPE_PREFIX):
        return _project_member_signature(type_name[len(_PROJECT_TYPE_PREFIX) :], member_name, ctx)
    runtime_object = resolve_runtime_object_type(type_name)
    if runtime_object is not None:
        lower = member_name.lower()
        runtime_member = next(
            (m for m in (runtime_object.get("members") or []) if m["name"].lower() == lower), None
        )
        return runtime_member.get("signature") if runtime_member is not None else None
    return resolve_host_member_signature(type_name, member_name, ctx.model)


def _project_member_signature(
    project_key: str, member_name: str, ctx: MemberCompletionContext
) -> str | None:
    project_member = _project_member_by_name(
        _project_class_members_by_name(ctx).get(project_key), member_name
    )
    return project_member.signature if project_member is not None else None


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
        return _msforms_member_return(host_type_name, member_name) or _host_member_return(
            host_type_name, member_name, ctx.model
        )
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
        return _msforms_member_return(owner_type, member_name) or _host_member_return(
            owner_type, member_name, ctx.model
        )
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


def _msforms_member_return(owner_type: str, member_name: str) -> _ResolvedMemberReturn | None:
    """`Views.SelectedItem.` chains into the returned object's own MSForms surface
    (XLIDE #32): the member's bare return name ("Tab", "Font") resolves to its
    qualified type exactly when the forms metadata carries that surface, so a
    primitive or unmodelled return ends the chain instead of guessing."""
    lower = member_name.lower()
    member = next(
        (m for m in msforms_control_members(owner_type) or [] if m["name"].lower() == lower), None
    )
    if member is None or not member.get("returns"):
        return None
    type_ = resolve_msforms_type_name(f"MSForms.{member['returns']}")
    return _ResolvedMemberReturn(type=type_, kind=member.get("kind", "property")) if type_ else None


def _host_member_return(
    owner_type: str, member_name: str, model: HostObjectModel | None
) -> _ResolvedMemberReturn | None:
    lower = member_name.lower()
    members = get_host_members(owner_type, model)
    member = next((m for m in members if m["name"].lower() == lower), None)
    if member is None:
        return None
    if member.get("returns"):
        # An accessor the library declares `As Object` is late bound however well
        # the model knows its element: `Worksheets(1).NoSuchMember` compiles
        # (XLIDE issue #114). A one-part union keeps the element for completion
        # and chaining without closing its surface. The hand-written collections
        # carry the repaired type on Item, so the library's word is read off
        # `_Default` as well.
        default_member = (
            next((m for m in members if m["name"] == "_Default"), None) if lower == "item" else None
        )
        declared_object = any(
            candidate is not None
            and (
                candidate.get("declaredType") == "Object"
                or _AS_OBJECT_SIGNATURE_RE.search(candidate.get("signature") or "") is not None
            )
            for candidate in (member, default_member)
        )
        return _ResolvedMemberReturn(
            type=f"{_UNION_TYPE_PREFIX}{member['returns']}" if declared_object else member["returns"],
            kind=member.get("kind", "property"),
        )
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
    # The project's own declarations are consulted BEFORE the referenced type
    # libraries, which is what VBA does. The Excel object model owns many ordinary
    # nouns (Point, Border, Font, Shape, Style, Name), so a class declared as one of
    # those got the library type's members instead of its own (XLIDE #11).
    key = _project_key_for_type_name(declared_type, ctx)
    if key:
        code_name_host = (ctx.code_names or {}).get(key)
        return _combined_type_key(key, code_name_host) if code_name_host else _project_type_key(key)
    return resolve_host_alias(declared_type, model) or resolve_msforms_type_name(declared_type)


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
    # A standard module is not a type anything is declared against, and an Enum is
    # a VALUE type: `Dim c As Corner` is a Long, not an object. Both are member
    # surfaces so `Module.Member` and `Corner.TopLeft` resolve, but neither may
    # answer here or a plain enum variable would look like an object.
    return (
        key
        if project_type is not None and project_type.kind not in ("standardModule", "enum")
        else None
    )


def _simple_type_name(type_text: str) -> str | None:
    trimmed = type_text.strip()
    return trimmed if is_identifier(trimmed) else None


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
    late_bound = False
    for type_ in types:
        parts = _parse_union_type_key(type_)
        if parts is not None:
            late_bound = True
        for item in parts if parts is not None else [type_]:
            key = item.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(item)
    # A one-part union stays a union: it marks a value the library declares
    # Object, whose members bind when it runs (XLIDE issue #114).
    if len(out) == 1 and not late_bound:
        return out[0]
    return f"{_UNION_TYPE_PREFIX}{_UNION_TYPE_SEPARATOR.join(out)}"


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


_PROJECT_TYPES_BY_NAME_CACHE = IdentityLru()


_NO_PROJECT_CLASS_MEMBERS: tuple[VbaProjectClassMembers, ...] = ()


def _project_class_members_by_name(
    ctx: MemberCompletionContext,
) -> dict[str, VbaProjectClassMembers]:
    return project_class_members_index(
        ctx.project_class_members if ctx.project_class_members is not None else _NO_PROJECT_CLASS_MEMBERS
    )


def project_class_members_index(
    project_class_members: Sequence[VbaProjectClassMembers],
) -> dict[str, VbaProjectClassMembers]:
    """The project's surfaces by lowercased name, indexed once per list. A name two
    surfaces share answers neither, since it cannot be told which one a reference
    means. Callers treat the result as read-only."""
    cached = _PROJECT_TYPES_BY_NAME_CACHE.get(project_class_members)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    out: dict[str, VbaProjectClassMembers] = {}
    ambiguous: set[str] = set()
    for type_ in project_class_members:
        key = type_.name.lower()
        if key in ambiguous:
            continue
        if key in out:
            del out[key]
            ambiguous.add(key)
            continue
        out[key] = type_
    return _PROJECT_TYPES_BY_NAME_CACHE.put(out, project_class_members)  # type: ignore[no-any-return]


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
    # The last one in source order wins, nested blocks included.
    latest: _SetAssignment | None = None
    for node in iter_body_nodes(body):
        if is_leaf_statement(node):
            if node.span.end > offset:
                continue
            hit = _set_assignment(source, node)
            if hit is not None and hit.name.lower() == lower_name:
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
    """A local variable, parameter, module-level variable or module procedure named
    `name`, preferring the procedure that encloses `offset`. Untyped declarations
    still shadow globals, so callers learn of them even with no `As` text."""
    module = ctx.parsed_module if ctx.parsed_module is not None else parse_module(source)
    lower = name.lower()
    enclosing = _enclosing_procedure(module, offset)

    if enclosing is not None:
        for param in enclosing.params:
            if param.name.lower() == lower:
                return _DeclaredBinding(as_type=param.as_type)
        local = _body_bindings(enclosing).get(lower)
        if local is not None:
            return local

    bindings = _module_bindings(module)
    variable = bindings.variables.get(lower)
    return variable if variable is not None else bindings.procedures.get(lower)


# Receiver lookups run once per dotted reference, and every member call a rule
# checks resolves one, so the declarations each lookup scanned are indexed once per
# parsed module and procedure instead. The first declaration of a name wins, as the
# scans did.
_MODULE_BINDINGS_CACHE = IdentityLru()
_BODY_BINDINGS_CACHE = IdentityLru(capacity=16)


@dataclass(frozen=True, slots=True)
class _ModuleBindings:
    variables: dict[str, _DeclaredBinding]
    procedures: dict[str, _DeclaredBinding]


def _module_bindings(module: ModuleNode) -> _ModuleBindings:
    """The module-level variables, then the module's procedures, as receivers.

    The module's own procedures shadow the host's globals, and this is where the two
    used to disagree: a module VARIABLE named `rows` resolved from its declaration,
    while `Public Property Get rows() As Widget` fell through to Excel's global
    `Rows`, so `rows.Where(p)` was measured against `Excel.Range` (XLIDE issue #68).
    The names that collide are the ones every workbook uses: rows, columns, cells,
    selection, names, sheets, application.

    A Function or Property Get yields its return type. A Sub, or a Property with
    only Let/Set, yields nothing readable, but it still shadows the global, so it
    binds with no type rather than letting the host answer for it."""
    cached = _MODULE_BINDINGS_CACHE.get(module)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    variables: dict[str, _DeclaredBinding] = {}
    readable: dict[str, _DeclaredBinding] = {}
    shadows: dict[str, _DeclaredBinding] = {}
    for mem in module.members:
        if isinstance(mem, VariableGroupNode):
            _add_group_bindings(mem, variables)
        elif isinstance(mem, ProcedureNode):
            lower = mem.name.lower()
            if mem.proc_kind in (ProcKind.FUNCTION, ProcKind.PROPERTY_GET):
                readable.setdefault(lower, _DeclaredBinding(as_type=mem.return_type or None))
            else:
                shadows.setdefault(lower, _DeclaredBinding(as_type=None))
    bindings = _ModuleBindings(variables=variables, procedures={**shadows, **readable})
    return _MODULE_BINDINGS_CACHE.put(bindings, module)  # type: ignore[no-any-return]


def _body_bindings(proc: ProcedureNode) -> dict[str, _DeclaredBinding]:
    """A procedure body's declarations, recursing into block nodes."""
    cached = _BODY_BINDINGS_CACHE.get(proc)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    out: dict[str, _DeclaredBinding] = {}
    _collect_body_bindings(proc.body, out)
    return _BODY_BINDINGS_CACHE.put(out, proc)  # type: ignore[no-any-return]


def _collect_body_bindings(body: Sequence[BodyNode], out: dict[str, _DeclaredBinding]) -> None:
    for node in iter_body_nodes(body):
        if isinstance(node, VariableGroupNode):
            _add_group_bindings(node, out)


def _add_group_bindings(group: VariableGroupNode, out: dict[str, _DeclaredBinding]) -> None:
    for decl in group.declarations:
        out.setdefault(decl.name.lower(), _DeclaredBinding(as_type=decl.as_type))


# -- AST helpers -----------------------------------------------------------


# Procedure spans are disjoint and ordered, so per-module the lookup is a
# binary search over (start, end, node) rows instead of an O(members) scan per
# dotted reference (which multiplied out on big modules).
_ENCLOSING_PROCEDURE_INDEX_CACHE = IdentityLru()


def _procedure_span_index(module: ModuleNode) -> list[tuple[int, int, ProcedureNode]]:
    cached = _ENCLOSING_PROCEDURE_INDEX_CACHE.get(module)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    rows = [
        (mem.span.start, mem.span.end, mem)
        for mem in module.members
        if isinstance(mem, ProcedureNode)
    ]
    return _ENCLOSING_PROCEDURE_INDEX_CACHE.put(rows, module)  # type: ignore[no-any-return]


def _enclosing_procedure(module: ModuleNode, offset: int) -> ProcedureNode | None:
    # Ends are increasing (source order), so the earliest row with end >= offset
    # is exactly the first member the linear scan would have matched; any earlier
    # row has end < offset and could never contain it.
    rows = _procedure_span_index(module)
    lo = 0
    hi = len(rows) - 1
    first = len(rows)
    while lo <= hi:
        mid = (lo + hi) >> 1
        if rows[mid][1] >= offset:
            first = mid
            hi = mid - 1
        else:
            lo = mid + 1
    if first < len(rows):
        start, end, node = rows[first]
        if start <= offset <= end:
            return node
    return None


# -- object-assignment types -----------------------------------------------


@dataclass(frozen=True, slots=True)
class KnownObjectAssignmentType:
    """The object class a declared type names."""

    # 'generic' | 'host' | 'project'. A generic type short-circuits compatibility
    # in both directions, which is what the untyped `Object` needs and what keeps
    # `Collection` from inventing mismatch errors while still requiring `Set`.
    kind: str
    display: str
    key: str
    implements: tuple[str, ...] = ()


ProjectTypeLookup = Callable[[str], VbaProjectClassMembers | None]


def resolve_known_object_assignment_type(
    type_name: str | None,
    ctx: MemberCompletionContext,
    project_type_lookup: ProjectTypeLookup | None = None,
) -> KnownObjectAssignmentType | None:
    """The object class a declared type names, when it names one that Set-binds and
    supports members: the generic `Object`, VBA's `Collection`, a host alias
    resolved through the host model, a DAO type, or an unambiguous project class,
    document or form. None for Variant, the scalar types and anything unknown.
    Ported from resolveKnownObjectAssignmentType (typeInference.ts);
    ``project_type_lookup`` is the per-pass index create_object_assignment_type_resolver
    passes."""
    if not type_name:
        return None
    normalized = normalize_type(type_name)
    if not normalized or normalized == "variant":
        return None
    if normalized == "object":
        return KnownObjectAssignmentType(kind="generic", display=type_name, key="object")
    if is_known_scalar_type(normalized):
        return None
    # VBA's own creatable class. It belongs to no host model and to no project, so
    # neither lookup below reaches it, and `Dim c As Collection : c = ...` read as
    # clean while refusing to compile.
    if normalized == "collection":
        return KnownObjectAssignmentType(kind="generic", display=type_name, key="collection")
    host = resolve_host_alias(type_name, ctx.model)
    if host:
        return KnownObjectAssignmentType(kind="host", display=type_name, key=host.lower())
    library = library_object_type(type_name)
    if library:
        return KnownObjectAssignmentType(kind="host", display=type_name, key=library.lower())
    simple = simple_type_name_for_assignment(type_name)
    if not simple:
        return None
    lower = simple.lower()
    match: VbaProjectClassMembers | None
    if project_type_lookup is not None:
        match = project_type_lookup(lower)
    else:
        matches = [
            project_type
            for project_type in (ctx.project_class_members or [])
            # userType and enum are VALUE types: `Dim c As Corner` is a Long, not
            # an object, so neither can make an assignment require Set.
            if project_type.kind not in ("userType", "enum", "standardModule")
            and project_type.name.lower() == lower
        ]
        match = matches[0] if len(matches) == 1 else None
    if match is None:
        return None
    return KnownObjectAssignmentType(
        kind="project",
        display=match.name,
        key=lower,
        implements=tuple(match.implements or []),
    )


def is_known_object_assignment_type(
    type_name: str | None, ctx: MemberCompletionContext
) -> bool:
    """True when a declared type names an object that Set-binds and supports members
    (see resolve_known_object_assignment_type)."""
    return resolve_known_object_assignment_type(type_name, ctx) is not None


def simple_type_name_for_assignment(type_text: str) -> str | None:
    """The bare type name of an As clause, without a trailing `()`, when it is one
    identifier."""
    trimmed = _TRAILING_EMPTY_PARENS_RE.sub("", type_text).strip()
    return trimmed if is_identifier(trimmed) else None


_DAO_QUALIFIED_RE = re.compile(r"^dao\.", re.IGNORECASE)
_LIBRARY_TYPES_BY_LOWER: dict[str, str] | None = None


def library_object_type(type_name: str | None) -> str | None:
    """A DAO type named with its library, `DAO.Recordset`, as the default-member
    table keys it. DAO has no host model, so this is how a variable of a DAO type is
    known to hold an object whose default member the table gives (XLIDE issue
    #464). Port of typeInference.ts' private libraryObjectType, kept here so the
    object-assignment resolver above can reach it."""
    global _LIBRARY_TYPES_BY_LOWER
    if not type_name or _DAO_QUALIFIED_RE.match(type_name.strip()) is None:
        return None
    if _LIBRARY_TYPES_BY_LOWER is None:
        _LIBRARY_TYPES_BY_LOWER = {
            key.lower(): key for key in HOST_DEFAULT_MEMBERS if key.startswith("DAO.")
        }
    return _LIBRARY_TYPES_BY_LOWER.get(type_name.strip().lower())


# -- member signatures -----------------------------------------------------

_DECLARED_RETURN_RE = re.compile(r"\)\s+As\s+([A-Za-z0-9_.]+)\s*$", re.IGNORECASE)
_LATE_BOUND_RETURN_RE = re.compile(r"^(?:Object|Variant)$", re.IGNORECASE)


def member_takes_own_arguments(signature: str | None) -> bool:
    """Whether a member called with arguments takes them itself, so the call is
    what the member declares it returns: a member that declares a parameter and a
    specific return type. GetSpellingSuggestions("helo") is a SpellingSuggestions,
    Shapes.Range(Array("A")) a ShapeRange (XLIDE issue #197). A member with no
    parameters passes the arguments to what it returns: Shapes.Placeholders(1) is a
    Shape. So does one the library declares As Object."""
    if not signature_declares_parameters(signature):
        return False
    match = _DECLARED_RETURN_RE.search(signature or "")
    declared = match.group(1) if match is not None else None
    return declared is not None and _LATE_BOUND_RETURN_RE.match(declared) is None


def signature_declares_parameters(signature: str | None) -> bool:
    """Whether a member signature label declares at least one parameter."""
    return len(_signature_parameters(signature)) > 0


def is_late_bound_type_key(type_key: str) -> bool:
    """Whether a receiver type key is late bound: a value the library declares
    Object, whose members bind when the code runs (`Worksheets(1)`), so the VBE
    checks nothing about them at compile time."""
    return type_key.startswith(_UNION_TYPE_PREFIX)


def _signature_parameters(signature: str | None) -> list[str]:
    """The parameters of a member signature label, trimmed, in order."""
    open_index = signature.find("(") if signature else -1
    if not signature or open_index < 0:
        return []
    depth = 0
    start = open_index + 1
    params: list[str] = []
    for i in range(open_index, len(signature)):
        ch = signature[i]
        if ch in ("(", "["):
            depth += 1
        elif ch in (")", "]"):
            depth -= 1
            if depth == 0 and ch == ")":
                params.append(signature[start:i])
                break
        elif ch == "," and depth == 1:
            params.append(signature[start:i])
            start = i + 1
    return [param.strip() for param in params if param.strip()]


def ms_forms_control_members(type_name: str) -> list[HostMember] | None:
    """Members of `MSForms.ComboBox` and friends, and of `MSForms.UserForm` for the
    form itself (upstream's msFormsControlMembers; the port keeps it in
    host/msforms.py)."""
    return msforms_control_members(type_name)
