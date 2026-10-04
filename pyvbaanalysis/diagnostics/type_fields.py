"""The user-defined types a module declares, and the fields a dotted path
reaches through them (XLIDE issue #248): `t.kids(0).vals` from `Dim t As Outer`.

Ported from xlide_vscode/src/analyzer/diagnostics/typeFields.ts.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Callable, Iterator, Mapping, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING

from ..conditional import ConditionalActivityTracker
from ..constants.integer_constant_expression import (
    IntegerConstantLookup,
    evaluate_integer_constant_expression,
)
from ..identity_cache import IdentityLru
from ..js_compat import js_trim
from ..lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import (
    BodyNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    TypeNode,
    WithBlockNode,
    is_leaf_statement,
)
from ..symbols.symbol_model import ModuleSymbols, VbaSymbol, VbaSymbolKind
from .context import statement_tokens
from .walker import (
    active_module_members,
    is_inactive_node,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

if TYPE_CHECKING:
    from .rules.arrays import ArrayDimensionBound


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) out of range."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return None if tok is None else tok.raw_text


@dataclass(frozen=True, slots=True)
class TypeFieldInfo:
    name: str
    # The declared type, lowercased; None when the field has none.
    type: str | None
    # The declared type as written, for a message.
    type_name: str | None
    is_array: bool
    # A fixed array field's bounds. An implicit lower bound is 0 whatever Option
    # Base says: `t.arr(0)` runs on `arr(3)` under Option Base 1 (measured in
    # Excel 16.0).
    dims: tuple[ArrayDimensionBound, ...] | None = None
    # A fixed array field whose bounds are not known here: `vals(1 To N)` with N
    # a Public Const of another module (issue #366).
    sized: bool | None = None
    # The raw length of a `String * n` field.
    fixed_length: str | None = None


# Each Type the module declares, by lowercased name, to its fields by lowercased name.
ModuleTypes = Mapping[str, Mapping[str, TypeFieldInfo]]

_MODULE_TYPES = IdentityLru()


def type_key(as_type: str | None) -> str | None:
    """A declared type's name, lowercased, without the `Module1.` qualifier."""
    if as_type is None:
        return None
    name = js_trim(js_trim(as_type).split(".")[-1]).lower()
    return name or None


def module_types(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None
) -> ModuleTypes:
    cached = _MODULE_TYPES.get(mod)
    if cached is not None and cached[0] == source and cached[1] is activity:
        result: ModuleTypes = cached[2]
        return result
    from .const_expr import collect_module_literal_integer_constants
    from .rules.arrays import parse_fixed_array_bounds_for_decl

    out: dict[str, dict[str, TypeFieldInfo]] = {}
    constants: IntegerConstantLookup | None = None
    for member in active_module_members(mod, activity):
        if not isinstance(member, TypeNode):
            continue
        fields: dict[str, TypeFieldInfo] = {}
        for field in member.fields:
            # Bounds a module Const or Enum member states are read too, and
            # bounds that cannot be read still make the field fixed (issue #366).
            bounds = _bounds_tokens(source, field.span) if field.is_array else None
            dims: Sequence[ArrayDimensionBound] | None = None
            if field.is_array and bounds:
                dims = parse_fixed_array_bounds_for_decl(source, field, 0)
                if dims is None:
                    if constants is None:
                        constants = collect_module_literal_integer_constants(mod, activity)
                    dims = _constant_dimensions(bounds, constants)
            sized = not dims and len(bounds or ()) > 0
            fields[field.name.lower()] = TypeFieldInfo(
                name=field.name,
                type=type_key(field.as_type),
                type_name=js_trim(field.as_type) if field.as_type is not None else None,
                is_array=field.is_array,
                dims=(
                    tuple(dataclasses.replace(dim, explicit_lower=True) for dim in dims)
                    if dims
                    else None
                ),
                sized=True if sized else None,
                fixed_length=field.fixed_length,
            )
        out[member.name.lower()] = fields
    _MODULE_TYPES.put((source, activity, out), mod)
    return out


def is_fixed_array_field(field: TypeFieldInfo) -> bool:
    """Whether an array field is fixed, its bounds known or not."""
    return field.is_array and (field.dims is not None or field.sized is True)


def _bounds_tokens(source: str, span: Span) -> list[VbaToken] | None:
    """The tokens between a declaration's parentheses, comments dropped; None without them."""
    toks = statement_tokens(source, span)
    open_ = next((i for i, tok in enumerate(toks) if tok.raw_text == "("), -1)
    close = -1 if open_ < 0 else match_paren_from(toks, open_)
    if close < 0:
        return None
    return [tok for tok in toks[open_ + 1 : close] if tok.kind is not TokenKind.COMMENT]


def _constant_dimensions(
    bounds: Sequence[VbaToken], constants: IntegerConstantLookup
) -> list[ArrayDimensionBound] | None:
    """The bounds `1 To N, N * 2` state with every name a module Const or Enum
    member. An implicit lower bound is 0."""
    from .rules.arrays import ArrayDimensionBound

    def text(part: Sequence[VbaToken]) -> str:
        return " ".join(tok.raw_text for tok in part)

    out: list[ArrayDimensionBound] = []
    for dim in split_top_level_token_groups(bounds, 0, ","):
        to = next((i for i, tok in enumerate(dim) if token_text(tok) == "to"), -1)
        lower = 0 if to < 0 else evaluate_integer_constant_expression(text(dim[:to]), constants)
        upper = evaluate_integer_constant_expression(text(dim if to < 0 else dim[to + 1 :]), constants)
        if lower is None or upper is None or len(dim) == 0:
            return None
        out.append(ArrayDimensionBound(lower=lower, upper=upper, explicit_lower=True))
    return out if out else None


# Per ModuleSymbols: each procedure (by identity, held alive here) to its names.
_VARIABLES = IdentityLru()


def variable_symbol_in(symbols: ModuleSymbols, proc: ProcedureNode, lower: str) -> VbaSymbol | None:
    """The variable a name means inside a procedure: a parameter or local first,
    then a module variable."""
    by_proc: dict[int, tuple[ProcedureNode, dict[str, VbaSymbol | None]]] | None = _VARIABLES.get(symbols)
    if by_proc is None:
        by_proc = {}
        _VARIABLES.put(by_proc, symbols)
    entry = by_proc.get(id(proc))
    if entry is None or entry[0] is not proc:
        from ..types.type_inference import procedure_symbol_for

        # Module variables first, so a parameter or local of the same name replaces one.
        names: dict[str, VbaSymbol | None] = {}
        for child in symbols.root.children or []:
            if child.kind is VbaSymbolKind.MODULE_VARIABLE and child.name.lower() not in names:
                names[child.name.lower()] = child
        proc_symbol = procedure_symbol_for(symbols, proc)
        own = (proc_symbol.children if proc_symbol is not None else None) or []
        for child in reversed(own):
            names[child.name.lower()] = (
                child
                if child.kind is VbaSymbolKind.PARAMETER or child.kind is VbaSymbolKind.LOCAL_VARIABLE
                else None
            )
        entry = (proc, names)
        by_proc[id(proc)] = entry
    return entry[1].get(lower)


@dataclass(frozen=True, slots=True)
class TypeRoot:
    """A value of a module Type that a chain of fields starts from."""

    type: str
    # The text so far, as written: `ts(1)`, `t`.
    display: str
    # Index of the `.` that follows the root.
    dot: int
    # The variable and fields so far, lowercased, while no subscript intervenes.
    path: str | None = None


def variable_root(
    toks: Sequence[VbaToken], i: int, variable: VbaSymbol | None, types: ModuleTypes
) -> TypeRoot | None:
    """The Type value a variable at `i` is, when a `.` follows it: `t.` with `Dim
    t As Outer`, or `ts(1).` with `Dim ts(2) As Outer`."""
    type_ = type_key(variable.as_type) if variable is not None else None
    if variable is None or not type_ or type_ not in types or variable.fixed_length is not None:
        return None
    if not variable.is_array:
        return (
            TypeRoot(type=type_, path=variable.name.lower(), display=toks[i].raw_text, dot=i + 1)
            if _raw(toks, i + 1) == "."
            else None
        )
    close = match_paren_from(toks, i + 1) if _raw(toks, i + 1) == "(" else -1
    if close > 0 and _raw(toks, close + 1) == ".":
        return TypeRoot(
            type=type_, display="".join(tok.raw_text for tok in toks[i : close + 1]), dot=close + 1
        )
    return None


def fixed_string_length(
    toks: Sequence[VbaToken],
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    types: ModuleTypes,
    constants: IntegerConstantLookup,
) -> float | None:
    """The declared length of the fixed-length string the tokens name: `s` with
    `Dim s As String * 3`, an element `s(1)` of an array of them, or a field
    `t.name` declared `name As String * 3`. A Const length is read too."""
    name = token_name(_at(toks, 0))
    lower = name.lower() if name is not None else None
    variable = variable_symbol_in(symbols, proc, lower) if lower else None
    raw: str | None
    if variable is not None and variable.fixed_length is not None:
        element = _raw(toks, 1) == "(" and match_paren_from(toks, 1) == len(toks) - 1
        raw = variable.fixed_length if (element if variable.is_array else len(toks) == 1) else None
    else:
        root = variable_root(toks, 0, variable, types)
        chain = field_chain(toks, root, types) if root is not None else []
        last = chain[-1] if chain else None
        whole = (
            last is not None
            and (last.close if last.close is not None else last.at) == len(toks) - 1
            and (not last.field.is_array or last.open is not None)
        )
        raw = last.field.fixed_length if whole and last is not None else None
    return None if raw is None else evaluate_integer_constant_expression(raw, constants)


def is_leading_dot(toks: Sequence[VbaToken], i: int) -> bool:
    """A `.` that opens an expression, which a With's subject qualifies."""
    prev = _at(toks, i - 1)
    if prev is None:
        return True
    if prev.kind is TokenKind.KEYWORD:
        return token_text(prev) != "me"
    # `FillArr .dyn`: a space after a name opens a call's argument.
    if (
        prev.kind is TokenKind.IDENTIFIER or prev.kind is TokenKind.BRACKETED_IDENTIFIER
    ) and prev.end < toks[i].start:
        return True
    return prev.kind is TokenKind.OPERATOR or prev.raw_text in ("(", ",", ":", "=")


@dataclass(frozen=True, slots=True)
class WithSubject:
    """What a With block's leading `.` stands for."""

    # The subject as written: `t`, `t.kids(0)`.
    display: str
    # The variable and fields, lowercased, while no subscript intervenes.
    path: str | None = None
    # The module Type the subject is a value of.
    type: str | None = None
    # The field the subject ends at, when it is one: `With t.o` is the object field o.
    field: TypeFieldInfo | None = None


def type_root_at(
    toks: Sequence[VbaToken],
    i: int,
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    types: ModuleTypes,
    subject: WithSubject | None,
) -> TypeRoot | None:
    """The Type value a chain starts from at `i`: a variable, or a leading `.`
    inside a With of a Type value."""
    if toks[i].raw_text == ".":
        if subject is not None and subject.type and is_leading_dot(toks, i):
            return TypeRoot(type=subject.type, path=subject.path, display=subject.display, dot=i)
        return None
    name = token_name(toks[i])
    lower = name.lower() if name is not None else None
    if not lower or _raw(toks, i - 1) == "." or _raw(toks, i - 1) == "!":
        return None
    return variable_root(toks, i, variable_symbol_in(symbols, proc, lower), types)


def walk_with_subjects(
    source: str,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    types: ModuleTypes,
    subject: WithSubject | None,
    visit: Callable[[LeafStatementNode, WithSubject | None], None],
) -> None:
    """Every statement of a body, block headers included, with the subject of the
    innermost With around it when that is a value of a module Type or a field of
    one.

    Upstream recurses per block; this walks an explicit stack in the same order.
    Each frame holds the body's nodes, the subject inside it, and the block's
    trailing line to visit, with the outer subject, once the body is done."""
    from .block_headers import block_header_statements

    frames: list[tuple[Iterator[BodyNode], WithSubject | None, LeafStatementNode | None, WithSubject | None]] = [
        (iter(body), subject, None, None)
    ]
    while frames:
        nodes, current, _after, _outer = frames[-1]
        for node in nodes:
            if is_inactive_node(activity, node):
                continue
            if is_leaf_statement(node):
                visit(node, current)
                continue
            child = getattr(node, "body", None)
            if not isinstance(child, list):
                continue
            headers = block_header_statements(source, node)
            before = headers.before
            if before is not None:
                visit(before, current)
            if isinstance(node, WithBlockNode):
                inner = (
                    with_subject(
                        statement_tokens_after_leading_label(source, before.span),
                        symbols,
                        proc,
                        types,
                        current,
                    )
                    if before is not None
                    else None
                )
            else:
                inner = current
            frames.append((iter(child), inner, headers.after, current))
            break
        else:
            _, _, after, outer = frames.pop()
            if after is not None:
                visit(after, outer)


def with_subject(
    toks: Sequence[VbaToken],
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    types: ModuleTypes,
    outer: WithSubject | None,
) -> WithSubject | None:
    """`With t`, `With t.kid`, `With .kids(0)`, `With t.o`: what the header names, or None."""
    if token_text(_at(toks, 0)) != "with" or len(toks) < 2:
        return None
    last = len(toks) - 1
    if last == 1:
        name = token_name(toks[1])
        variable = variable_symbol_in(symbols, proc, name.lower() if name is not None else "")
        type_ = type_key(variable.as_type) if variable is not None and not variable.is_array else None
        if type_ and type_ in types and variable is not None:
            return WithSubject(type=type_, path=variable.name.lower(), display=toks[1].raw_text)
        return None
    root = type_root_at(toks, 1, symbols, proc, types, outer)
    chain = field_chain(toks, root, types) if root is not None else []
    step = chain[-1] if chain else None
    end = (step.close if step.close is not None else step.at) if step is not None else -1
    if step is None or end != last or (step.field.is_array and step.open is None):
        return None
    type_ = step.field.type if step.field.type is not None and step.field.type in types else None
    if step.open is None:
        display = step.display
    else:
        assert step.close is not None
        display = step.display + "".join(tok.raw_text for tok in toks[step.open : step.close + 1])
    return WithSubject(
        display=display,
        path=step.path if step.open is None else None,
        type=type_,
        field=step.field,
    )


def with_subjects_in(
    source: str,
    proc: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    symbols: ModuleSymbols,
    types: ModuleTypes,
) -> dict[int, WithSubject]:
    """The subject of the innermost With around each statement, block headers
    included, by the statement's start."""
    out: dict[int, WithSubject] = {}

    def record(stmt: LeafStatementNode, subject: WithSubject | None) -> None:
        if subject is not None:
            out[stmt.span.start] = subject

    walk_with_subjects(source, proc.body, activity, symbols, proc, types, None, record)
    return out


@dataclass(slots=True)
class FieldStep:
    """One field a chain reaches."""

    field: TypeFieldInfo
    # `t.kid.dyn`, lowercased; None once a subscript intervenes.
    path: str | None
    # `t.kids(0).vals`, as written.
    display: str
    # Index of the field's name.
    at: int
    # The parentheses right after the field, when it is indexed.
    open: int | None = None
    close: int | None = None


def field_chain(toks: Sequence[VbaToken], root: TypeRoot, types: ModuleTypes) -> list[FieldStep]:
    """The fields a chain reaches from a Type value whose `.` is at `root.dot`:
    `t.kids(0).vals(9)` reaches kids, then vals. It stops at a name that is no
    field of a module Type, and at an array field used whole."""
    steps: list[FieldStep] = []
    type_: str | None = root.type
    path = root.path
    display = root.display
    i = root.dot
    while type_ and _raw(toks, i) == ".":
        name = token_name(_at(toks, i + 1))
        lower = name.lower() if name is not None else None
        fields = types.get(type_) if lower else None
        field = fields.get(lower) if fields is not None and lower else None
        if field is None or not lower:
            break
        path = None if path is None else f"{path}.{lower}"
        display = f"{display}.{toks[i + 1].raw_text}"
        step = FieldStep(field=field, path=path, display=display, at=i + 1)
        steps.append(step)
        i += 2
        if field.is_array:
            close = match_paren_from(toks, i) if _raw(toks, i) == "(" else -1
            if close < 0:
                break
            step.open = i
            step.close = close
            path = None
            display += "".join(tok.raw_text for tok in toks[i : close + 1])
            i = close + 1
        elif _raw(toks, i) == "(":
            break
        type_ = field.type if field.type is not None and field.type in types else None
    return steps
