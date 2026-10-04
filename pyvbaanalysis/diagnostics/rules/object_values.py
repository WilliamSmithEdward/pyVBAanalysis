"""Rule: an object read as a value when its type has no default member to give one.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/objectValues.ts (XLIDE
issue #183). Measured in Excel 16.0 (build 20326, 2026-09-29); each compiles
and raises every time it runs.

- A class with no default member: `s = c`, `v = c` into a Variant,
  `Debug.Print c`, `c & "x"`, `c + 1`, `If c = 1` -> 438, Object doesn't
  support this property or method. So do a Worksheet and a Workbook.
- A Collection, whose default member Item needs an index: `v = c` and
  `Debug.Print c` -> 450, Wrong number of arguments or invalid property
  assignment. With an operator, or into a String or other typed value, it
  does not compile, which is collection-operand's. Excel's collections whose
  default is their Item do the same: Hyperlinks, Areas, Borders, Windows,
  Workbooks, Shapes (issue #221).
- Excel's other objects with no default member raise 438 as a Worksheet
  does: Workbook, Font, Interior, Validation, Window, PageSetup, Border,
  Shape, Hyperlink (issue #221).
- An object variable still Nothing raises 91 first.

A class's default member read the wrong way (issue #256, measured in Excel
16.0): one whose first parameter is required, read with no argument, raises
449, Argument not optional; one with no parameter that returns a Collection
gives the Collection, whose own default needs an index, 450. A class with no
default member indexed, `c(1)`, raises 438. For Each over a class asks its -4
member (`VB_UserMemId = -4`) for an enumerator: with none it raises 438, and
with one returning a Collection rather than an object it raises 451.

`Set o = c`, passing `c` to a Variant parameter and `c Is Nothing` read no
value and run. The write side, `Let c = ...`, is set-required's.
"""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from typing import NamedTuple

from ...completion import MemberCompletionContext, is_known_object_assignment_type
from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import parse_vba_integer_literal
from ...js_compat import js_trim
from ...lexer.token_helpers import match_paren_from
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ForBlockNode,
    LeafStatementNode,
    ProcedureNode,
    ProcKind,
    Span,
    StatementNode,
    iter_body_nodes,
)
from ...symbols.symbol_model import ModuleSymbols, VbaProjectClassMember, VbaProjectClassMembers
from ...types.type_inference import (
    SourceDeclaredType,
    dao_whole_value_error,
    object_value_needs_index,
    procedure_symbol_for,
    type_environment_for,
    type_field_declared_type,
)
from ...types.type_names import is_known_scalar_type, normalize_type
from ..argument_inference import infer_member_expression_type
from ..context import PushFn, statement_tokens
from ..held_objects import ACTIVE_SHEET_HELD, HeldObjects, held_objects_at
from ..walker import (
    ProcedureStatementVisitor,
    bare_assignment_target,
    first_executable_token_index,
    statement_and_branch_spans,
    token_name,
    token_text,
)
from .shared import ONE_VALUE_BUILTINS, builtin_name_before
from .type_of_is import object_let_assignment_verdict

_SCALAR_OPERATORS = frozenset({"=", "<", ">", "<=", ">=", "<>", "+", "-", "*", "/", "\\", "&", "^"})

_NOTHING_91 = ", or '91' while it is Nothing"


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """toks[i] as JavaScript reads it: undefined (None) outside the list."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


class _DefaultFacts(NamedTuple):
    has_default: bool
    problem: str | None


class _ValueRead(NamedTuple):
    tok: VbaToken
    operator: bool
    into_typed_value: bool = False


class _Hit(NamedTuple):
    tok: VbaToken
    rule: str
    message: str


class _RangeHit(NamedTuple):
    start: int
    end: int
    rule: str
    message: str


class _CollectionArgument(NamedTuple):
    start: int
    end: int
    fn: str
    what: str
    compiles: bool


class _CreatedObject(NamedTuple):
    type: str
    span: Span


def check_object_default_values(
    source: str,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    push: PushFn,
    activity: ConditionalActivityTracker | None = None,
) -> ProcedureStatementVisitor:
    """Reads of an object's value where its type has no default member to give one."""
    # Facts belong to this invocation, so a later project metadata update is read anew.
    project_classes: dict[str, VbaProjectClassMembers | None] = {}

    def class_for_type(type_: str | None) -> VbaProjectClassMembers | None:
        lower = js_trim(type_).split(".")[-1].lower() if type_ is not None else None
        if not lower:
            return None
        if lower not in project_classes:
            # Preserve the first class, including an incomplete surface that
            # prevents a later same-name class from proving absence.
            found = next(
                (
                    candidate
                    for candidate in member_ctx.project_class_members or []
                    if candidate.kind == "class" and candidate.name.lower() == lower
                ),
                None,
            )
            project_classes[lower] = (
                found if found is not None and found.exhaustive is True else None
            )
        return project_classes.get(lower)

    # Keyed by identity, as upstream's Map is; the class object is kept with
    # its value so its id cannot be reused within this invocation.
    default_reads: dict[int, tuple[VbaProjectClassMembers, _DefaultFacts]] = {}

    def default_facts_for(cls: VbaProjectClassMembers) -> _DefaultFacts:
        entry = default_reads.get(id(cls))
        if entry is None:
            member = next(
                (candidate for candidate in cls.members if candidate.default_member), None
            )
            entry = (cls, _DefaultFacts(member is not None, _default_read_problem(member)))
            default_reads[id(cls)] = entry
        return entry[1]

    def has_default(cls: VbaProjectClassMembers) -> bool:
        return default_facts_for(cls).has_default

    enumerators: dict[int, tuple[VbaProjectClassMembers, VbaProjectClassMember | None]] = {}

    def enumerator_for(cls: VbaProjectClassMembers) -> VbaProjectClassMember | None:
        entry = enumerators.get(id(cls))
        if entry is None:
            entry = (
                cls,
                next((member for member in cls.members if _dispatch_id(member) == -4), None),
            )
            enumerators[id(cls)] = entry
        return entry[1]

    module_auto_instanced: set[str] = set()
    module_names = {child.name.lower() for child in symbols.root.children or []}
    for child in symbols.root.children or []:
        if child.is_auto_instantiated:
            module_auto_instanced.add(child.name.lower())

    def factory(proc: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, proc)
        proc_lower = proc.name.lower()
        # An array of a class is no object of it: `ReDim a(1 To 16)` and
        # `a(1)` on `Dim a() As K1` read no default member (issue #738).
        proc_sym = procedure_symbol_for(symbols, proc)
        own = (proc_sym.children if proc_sym is not None else None) or []
        own_names = {child.name.lower() for child in own} | {
            param.name.lower() for param in proc.params
        }
        arrays = (
            {child.name.lower() for child in own if child.is_array}
            | {param.name.lower() for param in proc.params if param.is_array}
            | {
                child.name.lower()
                for child in symbols.root.children or []
                if child.is_array and child.name.lower() not in own_names
            }
        )
        verdicts: dict[str, str] = {}

        def verdict_for(lower: str) -> str:
            verdict = verdicts.get(lower)
            if verdict is None:
                # The procedure's own name is its return value only as a target;
                # read, it is a recursive call.
                type_ = None if lower == proc_lower or lower in arrays else env.get(lower)
                verdict = object_let_assignment_verdict(type_, member_ctx) if type_ else "unknown"
                verdicts[lower] = verdict
            return verdict

        # An `As New` variable is never Nothing when read. A local shadows a
        # module-level variable of the same name.
        auto_instanced = set(module_auto_instanced)
        for child in own:
            if child.is_auto_instantiated:
                auto_instanced.add(child.name.lower())
            else:
                auto_instanced.discard(child.name.lower())

        def is_object_variable(name: str) -> bool:
            type_ = env.get(name.lower())
            return type_ is not None and is_known_object_assignment_type(type_, member_ctx)

        # A Let target of a declared scalar type: `s = c` with s a String.
        def is_typed_value(name: str) -> bool:
            lower = name.lower()
            type_ = normalize_type(proc.return_type if lower == proc_lower else env.get(lower))
            return type_ is not None and is_known_scalar_type(type_)

        def is_collection(lower: str) -> bool:
            return lower != proc_lower and normalize_type(env.get(lower)) == "collection"

        def class_of(lower: str) -> VbaProjectClassMembers | None:
            return (
                None if lower == proc_lower or lower in arrays else class_for_type(env.get(lower))
            )

        _check_for_each_enumerators(proc.body, class_of, enumerator_for, push)
        # An Object holding a Collection, `Set x = New Collection` with x As
        # Object, is late bound: its value read raises 450 when it runs, and
        # a Let to it 438 (issue #415, measured in Excel 16.0).
        late_bound = {lower for lower, type_ in env.items() if normalize_type(type_) == "object"}
        # An Object holding any other host object, `Set o = Range("A1").Font`,
        # reads as that object does, late (issue #685).
        held_at: list[Callable[[BodyNode], HeldObjects]] = []

        def declared_name(name: str) -> bool:
            return name in env or name in module_names

        def class_of_value(value: Sequence[VbaToken], offset: int) -> str | None:
            return _host_chain_type(source, value, offset, member_ctx, declared_name)

        def held_of(stmt: BodyNode, lower: str) -> str | None:
            if not held_at:
                held_at.append(held_objects_at(source, proc, symbols, activity, class_of_value))
            return held_at[0](stmt).classes.get(lower)

        # Excel's Application and Names, and Word's Document, by their own
        # rules (issues #415 and #438). An Object holding one is read late.
        host = _host_name(member_ctx)
        reads_by_type: AbstractSet[str] = (
            _EXCEL_DEFAULT_READS
            if host == "Excel"
            else _WORD_DEFAULT_READS
            if host == "Word"
            else frozenset()
        )
        host_defaults: dict[str, str] = {}
        for lower, type_ in env.items():
            if lower == proc_lower:
                continue
            key = normalize_type(type_) or ""
            if key in reads_by_type or (len(reads_by_type) > 0 and key == "object"):
                host_defaults[lower] = key

        def visitor(stmt: LeafStatementNode) -> None:
            if late_bound:
                _check_held_objects(
                    source, stmt, late_bound, lambda lower: held_of(stmt, lower), member_ctx, push
                )
            if host_defaults:

                def held(lower: str) -> str | None:
                    value = held_of(stmt, lower)
                    return value.lower() if value is not None else None

                for span in statement_and_branch_spans(stmt):
                    for hit in _host_default_reads(source, span, host_defaults, held):
                        push(
                            hit.rule,
                            hit.message,
                            Span(span.start + hit.tok.start, span.start + hit.tok.end),
                        )
            # `If c Then` on a Collection is condition-values' (issues #268, #424).
            condition = statement_tokens(source, stmt.span)
            # `If ws Then`, `ws(1)`, `CStr(ws)`: a type with no default member
            # has no value there either (issue #415, measured in Excel 16.0 on
            # a Worksheet, a Workbook and a Font).
            for tok in _no_default_reads(
                condition,
                token_text(_at(condition, 0)) in ("if", "elseif"),
                lambda lower: verdict_for(lower) == "noDefault" and not class_of(lower),
            ):
                lower = tok.raw_text.lower()
                # Keyed by the raw text, as upstream's `env.get(lower)!` is; a
                # bracketed name misses and interpolates as "undefined".
                type_ = env.get(lower, "undefined")
                nothing = "" if lower in auto_instanced else _NOTHING_91
                push(
                    "objectDefaultValue",
                    f"'{tok.raw_text}' is {_article(type_)} {type_}, which has no default member, so it has "
                    f"no value to read here. This will raise Run-time error '438': Object doesn't support "
                    f"this property or method{nothing}.",
                    Span(stmt.span.start + tok.start, stmt.span.start + tok.end),
                )
            for span in statement_and_branch_spans(stmt):
                for chain_hit in _host_chain_reads(
                    source,
                    span,
                    member_ctx,
                    lambda lower: lower in env or lower in module_names,
                    is_object_variable,
                ):
                    push(
                        chain_hit.rule,
                        chain_hit.message,
                        Span(span.start + chain_hit.start, span.start + chain_hit.end),
                    )
                for arg in _collection_arguments(
                    statement_tokens(source, span), is_collection, module_names
                ):
                    at = Span(span.start + arg.start, span.start + arg.end)
                    if arg.compiles:
                        push(
                            "objectDefaultValue",
                            f"{arg.what} is a Collection: its default member Item needs an index, so {arg.fn} "
                            f"has no value to read. This will raise Run-time error '450': Wrong number of "
                            f"arguments or invalid property assignment.",
                            at,
                        )
                    else:
                        push(
                            "collectionOperand",
                            f"{arg.what} is a Collection: its default member Item needs an index, so {arg.fn} "
                            f"has no value to take. This is a VBE compile error: Argument not optional.",
                            at,
                        )
                created = _new_object_let_into_variant(
                    source,
                    span,
                    env,
                    proc,
                    arrays,
                    lambda variable, field: type_field_declared_type(
                        symbols, procedure_symbol_for(symbols, proc), None, variable, field
                    ),
                )
                if created is not None:
                    created_verdict = object_let_assignment_verdict(created.type, member_ctx)
                    if created_verdict == "noDefault" or (
                        created_verdict == "argument"
                        and normalize_type(created.type) == "collection"
                    ):
                        push(
                            "objectDefaultValue",
                            f"'New {created.type}' is assigned without Set, so its value is read, and "
                            f"{created.type} has no default member to give one. This will raise Run-time "
                            f"error '438': Object doesn't support this property or method."
                            if created_verdict == "noDefault"
                            else f"'New {created.type}' is assigned without Set, so its value is read, and a "
                            f"Collection's default member Item needs an index. This will raise Run-time error "
                            f"'450': Wrong number of arguments or invalid property assignment.",
                            created.span,
                        )
                for indexed in _indexed_without_default(
                    statement_tokens(source, span), class_of, has_default
                ):
                    push(
                        "objectDefaultValue",
                        indexed.message,
                        Span(span.start + indexed.tok.start, span.start + indexed.tok.end),
                    )
                if_head = (
                    isinstance(stmt, StatementNode)
                    and stmt.single_line_if_branches is not None
                    and span is stmt.span
                )
                for read in _value_reads(source, span, if_head, is_object_variable, is_typed_value):
                    read_name = token_name(read.tok)
                    assert read_name is not None
                    lower = read_name.lower()
                    cls = class_of(lower)
                    wrong_way = (
                        default_facts_for(cls).problem
                        if cls is not None and not read.operator and not read.into_typed_value
                        else None
                    )
                    if wrong_way:
                        assert cls is not None
                        push(
                            "objectDefaultValue",
                            f"'{read.tok.raw_text}' is {_article(cls.name)} {cls.name}, {wrong_way}",
                            Span(span.start + read.tok.start, span.start + read.tok.end),
                        )
                        continue
                    verdict = verdict_for(lower)
                    if verdict != "noDefault" and verdict != "argument":
                        continue
                    type_ = env[lower]
                    # DAO checks the missing index itself (issue #464).
                    dao_error = (
                        dao_whole_value_error(type_)
                        if verdict == "argument" and not read.operator and not read.into_typed_value
                        else None
                    )
                    if dao_error:
                        nothing = "" if lower in auto_instanced else _NOTHING_91
                        push(
                            "objectDefaultValue",
                            f"'{read.tok.raw_text}' is {_article(type_)} {type_}: read whole, its default member "
                            f"reaches an Item that needs an index, so it has no value to read here. This will "
                            f"raise Run-time error {dao_error}{nothing}.",
                            Span(span.start + read.tok.start, span.start + read.tok.end),
                        )
                        continue
                    # A default member that needs an index raises 450 read as a
                    # value; with an operator, or into a typed value, it is a
                    # compile error, collection-operand's.
                    if verdict == "argument" and (
                        read.operator
                        or read.into_typed_value
                        or not object_value_needs_index(type_, member_ctx)
                    ):
                        continue
                    nothing = "" if lower in auto_instanced else _NOTHING_91
                    push(
                        "objectDefaultValue",
                        f"'{read.tok.raw_text}' is {_article(type_)} {type_}, which has no default member, so it "
                        f"has no value to read here. This will raise Run-time error '438': Object doesn't "
                        f"support this property or method{nothing}."
                        if verdict == "noDefault"
                        else f"'{read.tok.raw_text}' is {_article(type_)} {type_}: its default member Item needs "
                        f"an index, so it has no value to read here. This will raise Run-time error '450': "
                        f"Wrong number of arguments or invalid property assignment{nothing}.",
                        Span(span.start + read.tok.start, span.start + read.tok.end),
                    )

        return visitor

    return factory


def _host_name(member_ctx: MemberCompletionContext) -> str:
    """`memberCtx.model?.hostName ?? 'Excel'`."""
    model = member_ctx.model
    name = model.get("hostName") if model is not None else None
    return name if isinstance(name, str) else "Excel"


# Excel's Application gives its Name, "Microsoft Excel", as its value: set,
# `x + 1`, `x = 0` and `If x Then` raise 13, and `x(1)` does not compile, since
# Name takes no argument. Names gives its Item, whose argument the call needs:
# `v = x` set raises 449, and `x & "a"` does not compile, Type mismatch (issue
# #415, each measured in Excel 16.0).
_EXCEL_DEFAULT_READS = frozenset({"application", "names"})
# Word's Document gives its Name, a file name, which is no number: set, `x + 1`
# and `If x Then` raise 13. Through an Object, `x(1)` raises 451 and `x = 5`
# 5861, "'Name' is a read only property" (issue #438, measured in Word 16.0).
# Typed, those two are compile errors found elsewhere.
_WORD_DEFAULT_READS = frozenset({"document"})


class _NameDefault(NamedTuple):
    what: str
    value: str


# What each type with a String default gives as its value, for the message.
_NAME_DEFAULTS: dict[str, _NameDefault] = {
    "application": _NameDefault("the Application", 'its Name, "Microsoft Excel"'),
    "document": _NameDefault("a Document", "its Name, a file name"),
}
# Upstream's table is a plain object, so `type in NAME_DEFAULTS` and
# `NAME_DEFAULTS[type]` also see Object.prototype's lowercase keys, whose
# what and value interpolate as "undefined" (kept for parity).
_JS_PROTOTYPE_KEYS = frozenset({"constructor", "__proto__"})
_PROTOTYPE_NAME_DEFAULT = _NameDefault("undefined", "undefined")


def _name_default(type_: str) -> _NameDefault | None:
    found = _NAME_DEFAULTS.get(type_)
    if found is None and type_ in _JS_PROTOTYPE_KEYS:
        return _PROTOTYPE_NAME_DEFAULT
    return found


_ARITHMETIC = frozenset({"-", "*", "/", "\\", "^", "mod"})
_COMPARISONS = frozenset({"=", "<>", "<", ">", "<=", ">=", "+"})


def _host_default_reads(
    source: str,
    span: Span,
    typed: Mapping[str, str],
    held: Callable[[str], str | None],
) -> list[_Hit]:
    toks = statement_tokens(source, span)
    first = first_executable_token_index(toks)
    if token_text(_at(toks, first)) == "set":
        return []
    target = bare_assignment_target(source, span)
    eq = next((k for k, tok in enumerate(toks) if tok.raw_text == "="), -1) if target else -1
    then = (
        next((k for k, tok in enumerate(toks) if token_text(tok) == "then"), -1)
        if token_text(_at(toks, first)) in ("if", "elseif")
        else -1
    )

    def numeric(tok: VbaToken | None) -> bool:
        return tok is not None and (
            tok.kind is TokenKind.INTEGER_LITERAL or tok.kind is TokenKind.FLOAT_LITERAL
        )

    def op(tok: VbaToken | None) -> str:
        return (token_text(tok) or tok.raw_text) if tok is not None else ""

    out: list[_Hit] = []
    # `x = 5` on an Object holding a Document (issue #438).
    target_lower = target[0].lower() if target else None
    if (
        target_lower
        and typed.get(target_lower) == "object"
        and held(target_lower) == "document"
        and _raw(toks, first) != "."
    ):
        tok = next(
            candidate for candidate in toks if (token_name(candidate) or "").lower() == target_lower
        )
        return [
            _Hit(
                tok,
                "objectDefaultValue",
                f"'{tok.raw_text}' holds a Document, whose default member Name no Let reaches. This will "
                f"raise Run-time error '5861': 'Name' is a read only property.",
            )
        ]
    for i in range(first, len(toks)):
        name = token_name(toks[i])
        lower = name.lower() if name is not None else None
        declared = typed.get(lower) if lower else None
        if not declared or i == eq - 1 or _raw(toks, i - 1) == "." or _raw(toks, i + 1) == ".":
            continue
        assert lower is not None
        # An Object is read as what it holds, late.
        late = declared == "object"
        type_ = held(lower) if late else declared
        if not type_ or not (_name_default(type_) is not None or (not late and type_ == "names")):
            continue
        before = None if i - 1 == eq else _at(toks, i - 1)
        after = _at(toks, i + 1)
        named = _name_default(type_)
        if named is not None:
            if after is not None and after.raw_text == "(":
                # Typed, a Document's index is argument-count's from its type library.
                if late and type_ == "document":
                    out.append(
                        _Hit(
                            toks[i],
                            "objectDefaultValue",
                            f"'{toks[i].raw_text}' holds a Document, whose default member Name takes no "
                            f"argument. This will raise Run-time error '451': Property let procedure not "
                            f"defined and property get procedure did not return an object.",
                        )
                    )
                elif not late and type_ == "application":
                    out.append(
                        _Hit(
                            toks[i],
                            "argumentCount",
                            f"'{toks[i].raw_text}' is the Application, whose default member Name takes no "
                            f"argument. This is a VBE compile error: Wrong number of arguments or invalid "
                            f"property assignment.",
                        )
                    )
                continue
            if held(lower) != type_:
                continue
            condition = then > 0 and i == first + 1 and i + 1 == then
            arithmetic = op(after) in _ARITHMETIC or op(before) in _ARITHMETIC
            numeric_compare = (op(after) in _COMPARISONS and numeric(_at(toks, i + 2))) or (
                op(before) in _COMPARISONS and before is not None and numeric(_at(toks, i - 2))
            )
            if condition or arithmetic or numeric_compare:
                out.append(
                    _Hit(
                        toks[i],
                        "assignmentTypeMismatch",
                        f"'{toks[i].raw_text}' {'holds' if late else 'is'} {named.what}, whose value is "
                        f"{named.value}, which is not a number. This will raise Run-time error '13': Type "
                        f"mismatch.",
                    )
                )
            continue
        # Names.
        if op(after) == "&" or op(before) == "&":
            out.append(
                _Hit(
                    toks[i],
                    "objectDefaultValue",
                    f"'{toks[i].raw_text}' is a Names collection, whose default member Item needs its "
                    f"argument, so '&' has no value to join. This is a VBE compile error: Type mismatch.",
                )
            )
            continue
        if i == eq + 1 and i == len(toks) - 1 and held(lower) == "names":
            out.append(
                _Hit(
                    toks[i],
                    "objectDefaultValue",
                    f"'{toks[i].raw_text}' is a Names collection, whose default member Item needs its "
                    f"argument, so it has no value to read here. This will raise Run-time error '449': "
                    f"Argument not optional.",
                )
            )
    return out


# Excel types with no default member that the model does not list in full, so
# it cannot tell: WorksheetFunction, and Comment, which Range.Comment gives as
# Nothing where the cell has none (issue #685, each read as a value measured in
# Excel 16.0: 438, and 91 with no comment).
_EXCEL_NO_DEFAULT = frozenset({"worksheetfunction", "comment"})


class _HeldProblem(NamedTuple):
    what: str
    error: str


def _held_value_error(held: str, member_ctx: MemberCompletionContext) -> _HeldProblem | None:
    """What reading an Object as a value raises when it holds this class: a
    Collection, Sheets, a Dictionary or Hyperlinks need an index or key, 450;
    Names its argument, 449; a type with no default member, 438 (issues #415
    and #685, measured in Excel 16.0). None where the class is not judged here:
    Application and Document have rules of their own."""
    key = _host_type_key(held)
    needs_index = "This will raise Run-time error '450': Wrong number of arguments or invalid property assignment."
    if key == "collection":
        return _HeldProblem("a Collection, whose default member Item needs an index", needs_index)
    if key == "sheets":
        return _HeldProblem(
            "a Sheets collection, whose default member Item needs an index", needs_index
        )
    if key == "scripting.dictionary":
        return _HeldProblem("a Dictionary, whose default member Item needs a key", needs_index)
    if key == "names":
        return _HeldProblem(
            "a Names collection, whose default member Item needs its argument",
            "This will raise Run-time error '449': Argument not optional.",
        )
    if key == "worksheet or chart":
        return _HeldProblem(
            "a Worksheet or a Chart, neither of which has a default member",
            "This will raise Run-time error '438': Object doesn't support this property or method.",
        )
    if key is None or key == "application" or key == "document":
        return None
    if key in _EXCEL_NO_DEFAULT:
        none = ", or '91' while it holds no comment" if key == "comment" else ""
        return _HeldProblem(
            f"{_article(held)} {held}, which has no default member",
            f"This will raise Run-time error '438': Object doesn't support this property or method{none}.",
        )
    # A class of the project has rules of its own (issue #256).
    name = held.split(".")[-1].lower()
    if any(cls.name.lower() == name for cls in member_ctx.project_class_members or []):
        return None
    verdict = object_let_assignment_verdict(held, member_ctx)
    if verdict == "noDefault":
        return _HeldProblem(
            f"{_article(held)} {held}, which has no default member",
            "This will raise Run-time error '438': Object doesn't support this property or method.",
        )
    if verdict == "argument":
        return _HeldProblem(
            f"{_article(held)} {held}, whose default member Item needs an index", needs_index
        )
    return None


def _check_held_objects(
    source: str,
    stmt: LeafStatementNode,
    late_bound: AbstractSet[str],
    held_of: Callable[[str], str | None],
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> None:
    """`x + 1`, `If x Then`, `CStr(x)` and `x = 5` on an Object local that holds a
    Collection or a host object here."""
    toks = statement_tokens(source, stmt.span)
    if (
        not any((token_name(tok) or "").lower() in late_bound for tok in toks)
        or token_text(_at(toks, 0)) == "set"
    ):
        return
    target = bare_assignment_target(source, stmt.span)
    if target is not None:
        target_name, target_span, _ = target
        target_held = held_of(target_name.lower()) if target_name.lower() in late_bound else None
        if target_held is not None and target_held.lower() == "collection":
            push(
                "objectDefaultValue",
                f"'{target_name}' holds a Collection, whose default member Item needs an index, so a Let "
                f"cannot reach it. This will raise Run-time error '438': Object doesn't support this "
                f"property or method.",
                target_span,
            )
            return

    def is_late_bound(name: str) -> bool:
        return name.lower() in late_bound

    if_head = isinstance(stmt, StatementNode) and stmt.single_line_if_branches is not None
    reads = [
        *(
            read.tok
            for read in _value_reads(
                source, stmt.span, if_head, lambda _name: False, lambda _name: False
            )
        ),
        *_no_default_reads(
            toks, token_text(_at(toks, 0)) in ("if", "elseif"), is_late_bound, False
        ),
    ]
    for tok in reads:
        name = token_name(tok)
        lower = name.lower() if name is not None else None
        held = held_of(lower) if lower and is_late_bound(lower) else None
        problem = _held_value_error(held, member_ctx) if held else None
        if problem is not None:
            push(
                "objectDefaultValue",
                f"'{tok.raw_text}' holds {problem.what}, so it has no value to read here. {problem.error}",
                Span(stmt.span.start + tok.start, stmt.span.start + tok.end),
            )


def _no_default_reads(
    toks: Sequence[VbaToken],
    if_head: bool,
    judged: Callable[[str], bool],
    include_indexed: bool = True,
) -> list[VbaToken]:
    """Plain names read as a value where _value_reads does not look: the whole
    condition of an If or ElseIf, an index `x(1)` with no member after it, and a
    whole argument of a built-in that reads one value. Offsets are the
    statement's. Indexed reads can be excluded when judging the whole object."""
    out: list[VbaToken] = []
    then = (
        next((k for k, tok in enumerate(toks) if token_text(tok) == "then"), -1) if if_head else -1
    )
    second = token_name(_at(toks, 1))
    if then == 2 and second and judged(second.lower()):
        out.append(toks[1])
    for i in range(len(toks)):
        name = token_name(toks[i])
        if not name or not judged(name.lower()) or _raw(toks, i - 1) == ".":
            continue
        close = (
            match_paren_from(toks, i + 1) if include_indexed and _raw(toks, i + 1) == "(" else -1
        )
        indexed = close > i + 2 and _raw(toks, close + 1) != "." and _raw(toks, close + 1) != "="
        argument = (
            _raw(toks, i + 1) != "("
            and _raw(toks, i + 1) != "."
            and (_raw(toks, i - 1) or "") in ("(", ",")
            and (_raw(toks, i + 1) or "") in (")", ",")
            and token_text(_at(toks, builtin_name_before(toks, i))) in ONE_VALUE_BUILTINS
        )
        if indexed or argument:
            out.append(toks[i])
    return out


# Built-ins whose argument takes a value, measured in Excel 16.0 with a
# Collection (issue #242): a typed parameter refuses it while compiling,
# `Len(c)`, `CStr(c)`, `Abs(c)`; a Variant one asks the Item for a value at run
# time and raises 450, `InStr(c, "a")`, `Format(c)`, `Hex(c)`. `TypeName(c)`
# and `IsNumeric(c)` read no value.
_REFUSING_BUILTINS = frozenset(
    {
        "len", "cstr", "val", "clng", "cdbl", "cint", "cbool", "cdate", "trim$", "ucase$", "lcase$",
        "instrrev", "asc", "chr", "abs",
    }
)  # fmt: skip
_VALUE_READING_BUILTINS = frozenset(
    {
        "instr",
        "format",
        "ucase",
        "lcase",
        "trim",
        "ltrim",
        "rtrim",
        "left",
        "right",
        "mid",
        "cvar",
        "strcomp",
        "hex",
    }
)


def _collection_arguments(
    toks: Sequence[VbaToken],
    is_collection: Callable[[str], bool],
    module_names: AbstractSet[str],
) -> list[_CollectionArgument]:
    """A Collection, a variable or `New Collection`, as the first argument of
    one of those built-ins. Offsets are the statement's."""
    out: list[_CollectionArgument] = []
    for i in range(len(toks) - 2):
        # `Trim$(` lexes as Trim and a `$` of its own.
        suffixed = toks[i + 1].raw_text == "$"
        fn = toks[i].raw_text.lower() + ("$" if suffixed else "")
        open_ = i + 2 if suffixed else i + 1
        refuses = fn in _REFUSING_BUILTINS
        if (
            (not refuses and fn not in _VALUE_READING_BUILTINS)
            or _raw(toks, open_) != "("
            or toks[i].raw_text.lower() in module_names
        ):
            continue
        qualified = _raw(toks, i - 1) == "."
        if qualified and token_text(_at(toks, i - 2)) != "vba":
            continue
        a = _at(toks, open_ + 1)
        if a is None:
            continue
        created = token_text(a) == "new" and token_text(_at(toks, open_ + 2)) == "collection"
        last = open_ + 2 if created else open_ + 1
        closes = _raw(toks, last + 1) == ")" or _raw(toks, last + 1) == ","
        a_name = token_name(a)
        name = a_name.lower() if a_name is not None else None
        if not closes or (not created and (not name or not is_collection(name))):
            continue
        out.append(
            _CollectionArgument(
                a.start,
                toks[last].end,
                toks[i].raw_text + ("$" if suffixed else ""),
                "'New Collection'" if created else f"'{a.raw_text}'",
                not refuses,
            )
        )
    return out


def _value_reads(
    source: str,
    span: Span,
    if_head: bool,
    is_object_variable: Callable[[str], bool],
    is_typed_value: Callable[[str], bool],
) -> list[_ValueRead]:
    """The names a statement reads as a value: the whole value of a Let
    (`s = c`), an item Debug.Print prints (`Debug.Print "a"; c`), and an operand
    of a scalar operator (`c & "x"`). A name followed by `(` or `.`, or after a
    `.`, is a call or a member access, not the object's own value."""
    toks = statement_tokens(source, span)
    first = first_executable_token_index(toks)
    head = token_text(_at(toks, first))
    if head == "set":
        return []
    out: list[_ValueRead] = []

    def plain_name(i: int) -> bool:
        return (
            token_name(_at(toks, i)) is not None
            and _raw(toks, i - 1) != "."
            and _raw(toks, i + 1) != "("
            and _raw(toks, i + 1) != "."
        )

    target = bare_assignment_target(source, span)
    eq = next((k for k, tok in enumerate(toks) if tok.raw_text == "="), -1) if target else -1
    if target is not None:
        target_name = target[0]
        value = [tok for tok in target[2] if tok.kind is not TokenKind.COMMENT]
        # A Let into an object variable is set-required's.
        if (
            len(value) == 1
            and eq + 1 == next((k for k, tok in enumerate(toks) if tok.start == value[0].start), -1)
            and plain_name(eq + 1)
            and not is_object_variable(target_name)
        ):
            out.append(_ValueRead(value[0], False, is_typed_value(target_name)))
    if (
        head == "debug"
        and _raw(toks, first + 1) == "."
        and token_text(_at(toks, first + 2)) == "print"
    ):
        depth = 0
        for i in range(first + 3, len(toks)):
            raw = toks[i].raw_text
            depth += 1 if raw == "(" else -1 if raw == ")" else 0
            before = _raw(toks, i - 1)
            after_tok = _at(toks, i + 1)
            after = after_tok.raw_text if after_tok is not None else None
            alone = (i == first + 3 or before == "," or before == ";") and (
                after is None
                or after == ","
                or after == ";"
                or (after_tok is not None and after_tok.kind is TokenKind.COMMENT)
            )
            if depth == 0 and alone and plain_name(i):
                out.append(_ValueRead(toks[i], False))
    then = (
        next((k for k, tok in enumerate(toks) if token_text(tok) == "then"), -1) if if_head else -1
    )
    limit = then if then > 0 else len(toks)

    def is_operator(tok: VbaToken | None) -> bool:
        return tok is not None and (
            (tok.kind is TokenKind.OPERATOR and tok.raw_text in _SCALAR_OPERATORS)
            or token_text(tok) == "mod"
        )

    # Whole Let/Print values have no adjacent scalar operator. Each token
    # below is visited once, so operator reads cannot duplicate earlier reads.
    for i in range(first, limit):
        if i == eq - 1 or not plain_name(i):
            continue
        previous = None if i - 1 == eq else _at(toks, i - 1)
        if is_operator(_at(toks, i + 1)) or is_operator(previous):
            out.append(_ValueRead(toks[i], True))
    return out


def _host_chain_type(
    source: str,
    chain: Sequence[VbaToken],
    offset: int,
    member_ctx: MemberCompletionContext,
    declared: Callable[[str], bool],
) -> str | None:
    """The host type an Excel expression of globals and members gives:
    `ActiveWorkbook.Names` a Names, `Range("A1").Font` a Font, and ActiveSheet
    a Worksheet or a Chart. None for anything the source declares, or the
    model does not type."""
    head_name = token_name(_at(chain, 0))
    head = head_name.lower() if head_name is not None else None
    if not head or declared(head) or _host_name(member_ctx) != "Excel":
        return None
    if len(chain) == 1:
        if head == "activesheet":
            return ACTIVE_SHEET_HELD
        if head == "activeworkbook" or head == "thisworkbook":
            return "Workbook"
        return "WorksheetFunction" if head == "worksheetfunction" else None
    if not any(tok.raw_text == "." for tok in chain):
        return None
    inferred = infer_member_expression_type(source, list(chain), offset, member_ctx)
    return inferred.type_ if inferred is not None else None


def _host_type_key(type_: str | None) -> str | None:
    """A host type's lowercased name, Excel's prefix off; the Worksheets property
    gives a Sheets object (issue #404)."""
    normalized = normalize_type(type_)
    if normalized is None:
        return None
    if normalized.startswith("excel."):
        normalized = normalized[len("excel.") :]
    return "sheets" if normalized == "worksheets" else normalized


def _host_chain_reads(
    source: str,
    span: Span,
    member_ctx: MemberCompletionContext,
    declared: Callable[[str], bool],
    is_object_variable: Callable[[str], bool],
) -> list[_RangeHit]:
    """A host object read whole as a value, `v = ActiveWorkbook.Names`,
    `CStr(Range("A1").Font)` or `ActiveWorkbook & ""` (issue #685, measured in
    Excel 16.0): the expression's type, from the host's model, decides. One
    with no default member raises 438. Names, whose Item needs its argument,
    raises 449 read whole and does not compile with `&` or in a built-in (Type
    mismatch); Sheets raises 450 and does not compile there (Argument not
    optional). ActiveSheet is a Worksheet or a Chart, and neither has a default
    member."""
    toks = statement_tokens(source, span)
    first = first_executable_token_index(toks)
    if token_text(_at(toks, first)) == "set" or _host_name(member_ctx) != "Excel":
        return []
    target = bare_assignment_target(source, span)
    eq = (
        next((k for k, tok in enumerate(toks) if tok.raw_text == "="), -1)
        if target is not None and not is_object_variable(target[0])
        else -1
    )
    out: list[_RangeHit] = []
    i = first if eq < 0 else eq + 1
    while i < len(toks):
        if not token_name(toks[i]) or _raw(toks, i - 1) == "." or _raw(toks, i - 1) == "!":
            i += 1
            continue
        # The chain from here: names, `.member` and `(arguments)`.
        end = i
        while end + 1 < len(toks):
            if toks[end + 1].raw_text == "(":
                close = match_paren_from(toks, end + 1)
                if close < 0:
                    break
                end = close
            elif toks[end + 1].raw_text == "." and token_name(_at(toks, end + 2)):
                end += 2
            else:
                break
        chain = toks[i : end + 1]
        start = i
        before = _at(toks, start - 1)
        after = _at(toks, end + 1)
        whole = start == eq + 1 and eq >= 0 and end == len(toks) - 1
        joined = (before is not None and before.raw_text == "&") or (
            after is not None and after.raw_text == "&"
        )
        builtin = (
            before is not None
            and before.raw_text == "("
            and after is not None
            and after.raw_text == ")"
            and token_text(_at(toks, start - 2)) == "cstr"
            and _raw(toks, start - 3) != "."
        )
        # Not read whole here, its parts may be: `CStr(ActiveWorkbook)`.
        if not whole and not joined and not builtin:
            i += 1
            continue
        type_ = _host_chain_type(source, chain, span.start, member_ctx, declared)
        normalized = _host_type_key(type_)
        if not type_ or not normalized:
            i += 1
            continue
        i = end
        shown = "".join(tok.raw_text for tok in chain)
        at_start, at_end = chain[0].start, chain[-1].end
        # Through ActiveSheet, an Object, the rest is bound late: what would
        # not compile raises when it runs.
        late = token_text(chain[0]) == "activesheet" and len(chain) > 1
        verdict = (
            "noDefault"
            if normalized == "worksheet or chart" or normalized in _EXCEL_NO_DEFAULT
            else "argument"
            if normalized == "sheets"
            else object_let_assignment_verdict(type_, member_ctx)
        )
        if verdict == "noDefault":
            what = (
                "a Worksheet or a Chart, neither of which has"
                if normalized == "worksheet or chart"
                else f"{_article(type_)} {type_}, which has"
            )
            none = ", or '91' where the cell has no comment" if normalized == "comment" else ""
            out.append(
                _RangeHit(
                    at_start,
                    at_end,
                    "objectDefaultValue",
                    f"'{shown}' is {what} no default member, so it has no value to read here. This will raise "
                    f"Run-time error '438': Object doesn't support this property or method{none}.",
                )
            )
        elif normalized == "names":
            out.append(
                _RangeHit(
                    at_start,
                    at_end,
                    "objectDefaultValue",
                    f"'{shown}' is a Names collection, whose default member Item needs its argument, so it has "
                    f"no value to read here. This will raise Run-time error '449': Argument not optional."
                    if whole or late
                    else f"'{shown}' is a Names collection, whose default member Item needs its argument, so it "
                    f"has no value to take here. This is a VBE compile error: Type mismatch.",
                )
            )
        elif verdict == "argument" and (normalized == "sheets" or late):
            what = "a Sheets collection" if normalized == "sheets" else f"{_article(type_)} {type_}"
            if whole or late:
                out.append(
                    _RangeHit(
                        at_start,
                        at_end,
                        "objectDefaultValue",
                        f"'{shown}' is {what}, whose default member Item needs an index, so it has no value to "
                        f"read here. This will raise Run-time error '450': Wrong number of arguments or invalid "
                        f"property assignment.",
                    )
                )
            else:
                out.append(
                    _RangeHit(
                        at_start,
                        at_end,
                        "collectionOperand",
                        f"'{shown}' is {what}, whose default member Item needs an index, so it has no value to "
                        f"take here. This is a VBE compile error: Argument not optional.",
                    )
                )
        i += 1
    return out


def _new_object_let_into_variant(
    source: str,
    span: Span,
    env: Mapping[str, str],
    proc: ProcedureNode,
    arrays: AbstractSet[str],
    field_type: Callable[[str, str], SourceDeclaredType | None],
) -> _CreatedObject | None:
    """`v = New Collection` into a Variant, or into the Function's own result: a
    Let, which reads the new object's default value (issue #219, measured in
    Excel 16.0; `Set v = New Collection` runs). Into a typed scalar the VBE
    refuses it while compiling, which is not judged here."""
    target = bare_assignment_target(source, span)
    if target is None:
        return _new_object_let_into_variant_part(
            statement_tokens(source, span), span, env, arrays, field_type
        )
    value = [tok for tok in target[2] if tok.kind is not TokenKind.COMMENT]
    if len(value) != 2 or token_text(value[0]) != "new" or not token_name(value[1]):
        return None
    lower = target[0].lower()
    is_result = proc.proc_kind is ProcKind.FUNCTION and lower == proc.name.lower()
    if not is_result and lower not in env:
        return None
    declared = normalize_type(proc.return_type if is_result else env.get(lower))
    if (declared is not None and declared != "variant") or (is_result and proc.type_suffix):
        return None
    return _CreatedObject(
        value[1].raw_text, Span(span.start + value[0].start, span.start + value[1].end)
    )


def _new_object_let_into_variant_part(
    all_toks: Sequence[VbaToken],
    span: Span,
    env: Mapping[str, str],
    arrays: AbstractSet[str],
    field_type: Callable[[str, str], SourceDeclaredType | None],
) -> _CreatedObject | None:
    """`v(0) = New Collection` with v an array of Variant, and `t.v = New
    Collection` with v a Variant field: the Let reads the object's value as a
    whole variable's does (issue #306, measured in Excel 16.0: 450)."""
    toks = [tok for tok in all_toks if tok.kind is not TokenKind.COMMENT]
    i = first_executable_token_index(toks)
    if token_text(_at(toks, i)) == "let":
        i += 1
    name = token_name(_at(toks, i))
    if not name:
        return None
    equals = -1
    declared: str | None = None
    if _raw(toks, i + 1) == "(" and name.lower() in arrays:
        equals = match_paren_from(toks, i + 1) + 1
        declared = env.get(name.lower())
    elif _raw(toks, i + 1) == "." and token_name(_at(toks, i + 2)) and _raw(toks, i + 3) == "=":
        field = field_type(name, toks[i + 2].raw_text)
        if not field:
            return None
        equals = i + 3
        declared = field.as_type
    value = toks[equals + 1 :]
    type_ = normalize_type(declared)
    if (
        equals <= i
        or _raw(toks, equals) != "="
        or (type_ is not None and type_ != "variant")
        or len(value) != 2
        or token_text(value[0]) != "new"
        or not token_name(value[1])
    ):
        return None
    return _CreatedObject(
        value[1].raw_text, Span(span.start + value[0].start, span.start + value[1].end)
    )


_USER_MEM_ID_RE = re.compile(r"vb_(var)?usermemid", re.IGNORECASE | re.ASCII)


def _dispatch_id(member: VbaProjectClassMember) -> int | None:
    """The DISPID a member's attribute gives it: 0 for the default, -4 for the enumerator."""
    attr = next(
        (
            candidate
            for candidate in member.attributes or []
            if _USER_MEM_ID_RE.fullmatch(candidate.name)
        ),
        None,
    )
    raw = js_trim(attr.value_raw) if attr is not None else ""
    value = (
        parse_vba_integer_literal(raw[1:])
        if raw.startswith("-")
        else parse_vba_integer_literal(raw)
    )
    if value is None:
        return None
    return -value if raw.startswith("-") else value


_FIRST_PARAMETER_RE = re.compile(r"[^(]*\(([^,)]*)")
_PARAMARRAY_RE = re.compile(r"paramarray\b", re.IGNORECASE | re.ASCII)


def _default_read_problem(member: VbaProjectClassMember | None) -> str | None:
    """Why reading a class's default member with no argument fails, or None."""
    if member is None:
        return None
    match = _FIRST_PARAMETER_RE.match(member.signature or "")
    first = js_trim(match.group(1)) if match is not None else ""
    # The signature writes an Optional parameter in brackets: `Item([i As Long = 1])`.
    if first != "" and not first.startswith("[") and not _PARAMARRAY_RE.match(first):
        return (
            f"whose default member {member.name} takes an argument this read does not give. This will raise "
            f"Run-time error '449': Argument not optional."
        )
    if first == "" and normalize_type(member.returns) == "collection":
        return (
            f"whose default member {member.name} returns a Collection, and a Collection's default member Item "
            f"needs an index. This will raise Run-time error '450': Wrong number of arguments or invalid "
            f"property assignment."
        )
    return None


def _indexed_without_default(
    toks: Sequence[VbaToken],
    class_of: Callable[[str], VbaProjectClassMembers | None],
    has_default: Callable[[VbaProjectClassMembers], bool],
) -> list[_Hit]:
    """`c(1)` on a class with no default member to take the index."""
    out: list[_Hit] = []
    for i in range(len(toks) - 1):
        name = token_name(toks[i])
        lower = name.lower() if name is not None else None
        cls = (
            class_of(lower)
            if lower and toks[i + 1].raw_text == "(" and _raw(toks, i - 1) != "."
            else None
        )
        if cls is not None and not has_default(cls):
            out.append(
                _Hit(
                    toks[i],
                    "objectDefaultValue",
                    f"'{toks[i].raw_text}' is {_article(cls.name)} {cls.name}, which has no default member to "
                    f"take an index. This will raise Run-time error '438': Object doesn't support this property "
                    f"or method.",
                )
            )
    return out


def _is_simple_name(text: str) -> bool:
    """Upstream's `/^[\\p{L}_][\\p{L}\\p{N}_]*$/u` (no marks, unlike IDENT_RE)."""
    if not text:
        return False
    for k, ch in enumerate(text):
        category = unicodedata.category(ch)[0]
        if ch == "_" or category == "L" or (k > 0 and category == "N"):
            continue
        return False
    return True


def _check_for_each_enumerators(
    body: Sequence[BodyNode],
    class_of: Callable[[str], VbaProjectClassMembers | None],
    enumerator_for: Callable[[VbaProjectClassMembers], VbaProjectClassMember | None],
    push: PushFn,
) -> None:
    """`For Each v In c` over a class: the -4 member it needs, and what that returns.

    Upstream recurses per block body; this walks the same nodes in the same
    order on an explicit stack."""
    for node in iter_body_nodes(body):
        if (
            not isinstance(node, ForBlockNode)
            or not node.each
            or node.source_expression_span is None
        ):
            continue
        over = js_trim(node.source_expression or "")
        cls = class_of(over.lower()) if _is_simple_name(over) else None
        enumerator = enumerator_for(cls) if cls is not None else None
        returns = normalize_type(enumerator.returns if enumerator is not None else None)
        if cls is not None and enumerator is None:
            push(
                "objectDefaultValue",
                f"'{over}' is {_article(cls.name)} {cls.name}, which has no member marked VB_UserMemId = -4 for "
                f"For Each to ask for its elements. This will raise Run-time error '438': Object doesn't support "
                f"this property or method.",
                node.source_expression_span,
            )
        elif cls is not None and enumerator is not None and returns == "collection":
            push(
                "objectDefaultValue",
                f"'{over}' is {_article(cls.name)} {cls.name}, whose enumerator {enumerator.name} returns a "
                f"Collection, not the enumerator object For Each needs. This will raise Run-time error '451': "
                f"Property let procedure not defined and property get procedure did not return an object.",
                node.source_expression_span,
            )


_VOWEL_START_RE = re.compile(r"[aeiou]", re.IGNORECASE | re.ASCII)


def _article(type_: str) -> str:
    return "an" if _VOWEL_START_RE.match(type_) else "a"
