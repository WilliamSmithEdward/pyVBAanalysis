"""Rule family: module-level declarations that refer ahead or around a cycle
(XLIDE issue #211).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/declarationOrder.ts.
Measured in 64-bit Excel 16.0 (build 20326, 2026-09-30):

- declaration-forward-reference: a Const, an Enum member, an array bound
  or a `String * N` length that uses a Const or Enum member declared
  further down its module, or itself: "Constant expression required". A
  Type member of a Type declared further down: "Forward reference to
  user-defined type". With the order swapped each compiles. A variable
  of a later Type is fine, and so is order across modules.
- circular-declaration-dependency: a Type with a member of its own type,
  and a cycle of Consts or Types through other modules: "Circular
  dependencies between modules".
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import enum_member_raw_expression
from ...js_compat import js_trim
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize
from ...parser.nodes import EnumNode, ModuleNode, Span, TypeNode, VariableGroupNode
from ...symbols.symbol_model import VbaProjectClassMembers
from ..context import PushFn
from ..walker import (
    absolute_span,
    active_module_members,
    is_inactive_node,
    match_paren_from,
    statement_tokens,
    token_name,
    token_text,
)


@dataclass(slots=True)
class _Declarations:
    """Where each module-level constant and Type is declared; None for a name declared twice."""

    constants: dict[str, int | None]
    # Enum name -> member -> where the member is declared.
    enums: dict[str, dict[str, int]]
    types: dict[str, int | None]
    # The value expression of each constant and Enum member, for the cycle walk.
    expressions: dict[str, str]


def check_declaration_order(
    source: str,
    mod: ModuleNode,
    module_name: str | None,
    project_integer_constants: Mapping[str, str | None] | None,
    project_class_members: Sequence[VbaProjectClassMembers] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    declared = _module_declarations(mod, activity)
    if len(declared.constants) == 0 and len(declared.types) == 0:
        return

    def check(
        tokens: Sequence[VbaToken], base: Span, site_start: int, own_name: str | None
    ) -> None:
        for ref in _constant_references(tokens, declared):
            if ref.start < site_start:
                continue
            self_ = ref.start == site_start and ref.name.lower() == (
                own_name.lower() if own_name is not None else None
            )
            push(
                "declarationForwardReference",
                f"'{ref.name}' is defined in terms of itself. This is a VBE compile error: "
                "Constant expression required."
                if self_
                else f"'{ref.name}' is declared further down the module, and a constant "
                "expression can only use what is declared above it. This is a VBE compile error: "
                "Constant expression required.",
                absolute_span(base, ref.token),
            )

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            for decl in member.declarations:
                if is_inactive_node(activity, decl):
                    continue
                tokens = statement_tokens(source, decl.span)
                check(
                    _tokens_after_equals(tokens)
                    if member.is_const
                    else _declaration_shape_tokens(tokens),
                    decl.span,
                    decl.span.start,
                    decl.name if member.is_const else None,
                )
        elif isinstance(member, EnumNode):
            for enum_member in member.members:
                if enum_member.value_raw is not None and not is_inactive_node(
                    activity, enum_member
                ):
                    check(
                        _tokens_after_equals(statement_tokens(source, enum_member.span)),
                        enum_member.span,
                        enum_member.span.start,
                        enum_member.name,
                    )
        elif isinstance(member, TypeNode):
            type_start = declared.types.get(member.name.lower())
            for field in member.fields:
                if is_inactive_node(activity, field):
                    continue
                tokens = statement_tokens(source, field.span)
                check(_declaration_shape_tokens(tokens), field.span, field.span.start, None)
                as_type = _as_type_token(tokens)
                target = declared.types.get(token_text(as_type)) if as_type is not None else None
                if as_type is None or target is None or type_start is None or target < type_start:
                    continue
                if target == type_start:
                    push(
                        "circularDeclarationDependency",
                        f"Type '{member.name}' has a member of its own type. This is a VBE compile "
                        "error: Circular dependencies between modules.",
                        absolute_span(field.span, as_type),
                    )
                else:
                    push(
                        "declarationForwardReference",
                        f"Type '{as_type.raw_text}' is declared further down the module. This is a "
                        "VBE compile error: Forward reference to user-defined type.",
                        absolute_span(field.span, as_type),
                    )

    if module_name is not None and project_integer_constants and len(project_integer_constants) > 0:
        _check_constant_cycles_across_modules(
            mod, module_name, declared, project_integer_constants, activity, push
        )
    own_lower = module_name.lower() if module_name is not None else None
    if project_class_members is not None and any(
        surface.kind == "userType" and surface.module_name.lower() != own_lower
        for surface in project_class_members
    ):
        _check_type_cycles_across_modules(
            source, mod, module_name, project_class_members, activity, push
        )


def _module_declarations(
    mod: ModuleNode, activity: ConditionalActivityTracker | None
) -> _Declarations:
    constants: dict[str, int | None] = {}
    enums: dict[str, dict[str, int]] = {}
    types: dict[str, int | None] = {}
    expressions: dict[str, str] = {}

    def place(table: dict[str, int | None], name: str, start: int) -> None:
        key = name.lower()
        table[key] = None if key in table else start

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode) and member.is_const:
            for decl in member.declarations:
                if not is_inactive_node(activity, decl):
                    place(constants, decl.name, decl.span.start)
                    if decl.default_raw is not None:
                        expressions[decl.name.lower()] = decl.default_raw
        elif isinstance(member, EnumNode):
            members: dict[str, int] = {}
            previous: str | None = None
            for enum_member in member.members:
                if is_inactive_node(activity, enum_member):
                    continue
                place(constants, enum_member.name, enum_member.span.start)
                members[enum_member.name.lower()] = enum_member.span.start
                expressions[enum_member.name.lower()] = enum_member_raw_expression(
                    enum_member.value_raw, previous
                )
                previous = enum_member.name
            enums[member.name.lower()] = members
        elif isinstance(member, TypeNode):
            place(types, member.name, member.span.start)
    return _Declarations(constants, enums, types, expressions)


@dataclass(frozen=True, slots=True)
class _ConstantReference:
    name: str
    token: VbaToken
    start: int


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) outside the list."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    """`toks[i]?.rawText`."""
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


def _constant_references(
    tokens: Sequence[VbaToken], declared: _Declarations
) -> list[_ConstantReference]:
    """The module's constants an expression names: a bare name, or an Enum member
    qualified by its Enum. A name after any other `.` is some other object's
    member and names nothing here."""
    out: list[_ConstantReference] = []
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        name = token_name(tok)
        if not name or tok.kind is TokenKind.BRACKETED_IDENTIFIER or _raw_at(tokens, i - 1) == ".":
            i += 1
            continue
        if _raw_at(tokens, i + 1) == ".":
            member = _at(tokens, i + 2)
            enum_members = declared.enums.get(name.lower())
            start = (
                enum_members.get(token_text(member))
                if member is not None and enum_members is not None
                else None
            )
            if start is not None and member is not None:
                out.append(_ConstantReference(f"{name}.{member.raw_text}", member, start))
            i += 3
            continue
        constant_start = declared.constants.get(name.lower())
        if constant_start is not None:
            out.append(_ConstantReference(name, tok, constant_start))
        i += 1
    return out


def _tokens_after_equals(tokens: Sequence[VbaToken]) -> list[VbaToken]:
    """The tokens after a declaration's top-level `=`: a Const's or Enum member's value."""
    depth = 0
    for i, tok in enumerate(tokens):
        raw = tok.raw_text
        depth += 1 if raw == "(" else -1 if raw == ")" else 0
        if depth == 0 and raw == "=" and tok.kind is TokenKind.OPERATOR:
            return [t for t in tokens[i + 1 :] if t.kind is not TokenKind.COMMENT]
    return []


def _declaration_shape_tokens(tokens: Sequence[VbaToken]) -> list[VbaToken]:
    """The tokens a declaration's shape reads constants from: its array bounds and
    the length of a `String * N`. The type after `As` is not among them."""
    out: list[VbaToken] = []
    open_ = next((k for k, tok in enumerate(tokens) if tok.raw_text == "("), -1)
    as_at = next((k for k, tok in enumerate(tokens) if token_text(tok) == "as"), -1)
    if open_ >= 0 and (as_at < 0 or open_ < as_at):
        close = match_paren_from(tokens, open_)
        out.extend(tokens[open_ + 1 : len(tokens) if close < 0 else close])
    star = next(
        (k for k, tok in enumerate(tokens) if k > as_at and as_at >= 0 and tok.raw_text == "*"),
        -1,
    )
    if star >= 0:
        out.extend(tok for tok in tokens[star + 1 :] if tok.kind is not TokenKind.COMMENT)
    return out


def _as_type_token(tokens: Sequence[VbaToken]) -> VbaToken | None:
    """The type name after `As`, unless it is qualified (another module's) or a New."""
    as_at = next((k for k, tok in enumerate(tokens) if token_text(tok) == "as"), -1)
    name = None if as_at < 0 else _at(tokens, as_at + 1)
    if name is None or not token_name(name) or _raw_at(tokens, as_at + 2) == ".":
        return None
    return name


@dataclass(frozen=True, slots=True)
class _Dependency:
    module: str
    key: str
    written: str


@dataclass(slots=True)
class _Frame:
    edges: list[_Dependency]
    index: int
    via: str | None


_PLAIN_NUMBER_RE = re.compile(r"[-+]?[0-9]+(\.[0-9]+)?")


def _check_constant_cycles_across_modules(
    mod: ModuleNode,
    module_name: str,
    declared: _Declarations,
    external: Mapping[str, str | None],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Consts and Enum members that reach themselves through another module:
    `Public Const A1 = B1 + 1` here and `Public Const B1 = A1 + 1` in Module2.
    The other modules' values come from the project's exported constants,
    keyed by bare and by qualified name; a value another module could fold is
    already a number there and ends the walk."""
    own = module_name.lower()
    # The module that exports each bare name, where exactly one does.
    owners: dict[str, str | None] = {}
    for key in external:
        dot = key.find(".")
        if dot > 0:
            bare = key[dot + 1 :]
            module = key[:dot]
            owners[bare] = None if bare in owners and owners[bare] != module else module
    modules = {value for value in owners.values() if value is not None}

    prepared: dict[str, list[_Dependency]] = {}

    def dependencies(raw: str, scope: str) -> list[_Dependency]:
        cache_key = scope + "\0" + raw
        cached = prepared.get(cache_key)
        if cached is not None:
            return cached
        out: list[_Dependency] = []
        # Upstream lexes the raw value with the uncached tokenize, as here.
        tokens = [
            tok
            for tok in tokenize(raw)
            if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE
        ]
        i = 0
        while i < len(tokens):
            name = token_name(tokens[i])
            if not name or _raw_at(tokens, i - 1) == ".":
                i += 1
                continue
            written = tokens[i].raw_text
            qualifier: str | None = None
            qualified_name = token_name(_at(tokens, i + 2))
            if _raw_at(tokens, i + 1) == "." and name.lower() in modules and qualified_name:
                qualifier = name.lower()
                name = qualified_name
                written = f"{tokens[i].raw_text}.{tokens[i + 2].raw_text}"
                i += 2
            lower = name.lower()
            target: tuple[str, str] | None = None
            owner = owners.get(lower)
            if qualifier is not None:
                if qualifier == own:
                    target = (own, lower)
                elif f"{qualifier}.{lower}" in external:
                    target = (qualifier, lower)
            elif scope == own:
                if lower in declared.expressions:
                    target = (own, lower)
                elif owner:
                    target = (owner, lower)
            elif f"{scope}.{lower}" in external:
                target = (scope, lower)
            elif lower in declared.expressions:
                target = (own, lower)
            elif owner:
                target = (owner, lower)
            i += 1
            if target is None:
                continue
            out.append(_Dependency(target[0], target[1], written))
        prepared[cache_key] = out
        return out

    def reaches_itself(start: str) -> str | None:
        raw = declared.expressions.get(start)
        if raw is None:
            return None
        seen: set[str] = set()
        # Explicit DFS frames preserve reference order and the first foreign name
        # without consuming the call stack for long dependency chains.
        stack: list[_Frame] = [_Frame(dependencies(raw, own), 0, None)]
        while stack:
            frame = stack[-1]
            if frame.index == len(frame.edges):
                stack.pop()
                continue
            edge = frame.edges[frame.index]
            frame.index += 1
            module, key, written = edge.module, edge.key, edge.written
            if module == own and key == start and frame.via is not None:
                return frame.via
            id_ = module + "." + key
            if id_ in seen:
                continue
            seen.add(id_)
            nxt = declared.expressions.get(key) if module == own else external.get(id_)
            if nxt is None or _PLAIN_NUMBER_RE.fullmatch(js_trim(nxt)):
                continue
            stack.append(
                _Frame(
                    dependencies(nxt, module),
                    0,
                    frame.via if frame.via is not None else (None if module == own else written),
                )
            )
        return None

    for member in active_module_members(mod, activity):
        names: list[tuple[str, Span]] = []
        if isinstance(member, VariableGroupNode) and member.is_const:
            for decl in member.declarations:
                if not is_inactive_node(activity, decl):
                    names.append(
                        (decl.name, decl.name_span if decl.name_span is not None else decl.span)
                    )
        elif isinstance(member, EnumNode):
            for enum_member in member.members:
                if not is_inactive_node(activity, enum_member):
                    names.append(
                        (
                            enum_member.name,
                            enum_member.name_span
                            if enum_member.name_span is not None
                            else enum_member.span,
                        )
                    )
        for name, span in names:
            via = reaches_itself(name.lower())
            if via:
                push(
                    "circularDeclarationDependency",
                    f"'{name}' depends on itself through '{via}' in another module. This is a VBE "
                    "compile error: Circular dependencies between modules.",
                    span,
                )


def _check_type_cycles_across_modules(
    source: str,
    mod: ModuleNode,
    module_name: str | None,
    project_class_members: Sequence[VbaProjectClassMembers],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Types that contain themselves through another module's Type:
    `Public Type T1: a As T2` here and `Public Type T2: b As T1` in Module2."""
    own = module_name.lower() if module_name is not None else None
    by_name: dict[str, VbaProjectClassMembers | None] = {}
    for surface in project_class_members:
        if surface.kind != "userType":
            continue
        key = surface.name.lower()
        by_name[key] = None if key in by_name else surface
    for member in active_module_members(mod, activity):
        if not isinstance(member, TypeNode):
            continue
        start = member.name.lower()
        for field in member.fields:
            if is_inactive_node(activity, field):
                continue
            as_type = _as_type_token(statement_tokens(source, field.span))
            first = by_name.get(token_text(as_type)) if as_type is not None else None
            if as_type is None or first is None or first.module_name.lower() == own:
                continue
            # by_name resolves each unambiguous name to one canonical surface.
            # Mark it when queued so shared descendants enter only once.
            seen: set[int] = {id(first)}
            queue: list[VbaProjectClassMembers] = [first]
            cycle = False
            head = 0
            while head < len(queue) and not cycle:
                surface = queue[head]
                for field_member in surface.members:
                    type_name = field_member.returns.lower() if field_member.returns else None
                    if not type_name:
                        continue
                    nxt = by_name.get(type_name)
                    if type_name == start and (nxt is None or nxt.module_name.lower() == own):
                        cycle = True
                        break
                    if nxt is not None and nxt.module_name.lower() != own and id(nxt) not in seen:
                        seen.add(id(nxt))
                        queue.append(nxt)
                head += 1
            if cycle:
                push(
                    "circularDeclarationDependency",
                    f"Type '{member.name}' contains itself through "
                    f"{first.module_name}.{first.name}. This is a VBE compile error: Circular "
                    "dependencies between modules.",
                    absolute_span(field.span, as_type),
                )
