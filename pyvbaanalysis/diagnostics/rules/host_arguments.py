"""Rule family: host object model arguments the code proves wrong (XLIDE issue #122).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/hostArguments.ts.

Office collections are 1-based, a cell has a row and a column of at least
1, and a sheet name has a spelling Excel enforces. Every case here was
measured through pyVBAharness on 2026-09-26 in Excel, Word and PowerPoint
16.0 (build 20326): each compiles and raises every time it runs, whatever
the document holds.

 - host-argument-out-of-range
     Index 0 or below into a host collection: Worksheets(0), Sheets(0),
     Workbooks(0), Names(0), Charts(0), Windows(0), ListObjects(0),
     Comments(0), Hyperlinks(0), CommandBars(0), AddIns(0), Worksheets.Item(0)
     -> 9 in Excel; Shapes(0) -> -2147024809; PivotTables(0) -> 1004; a
     Range-valued collection (Rows, Columns, Areas, Cells) -> 1004.
     Paragraphs(0), Documents(0), Tables(0), Sections(0), Words(0), ...
     -> 5941 in Word (Shapes(0) -> -2147024809). Slides(0), Presentations(0),
     Shapes(0), Designs(0), Slides.Add 0 -> -2147188160 in PowerPoint.
     Cells(0, 1), Cells(1, 0), Cells(0) -> 1004. Range("A0"), Range("$A$0"),
     Range("Sheet1!A0"), Range("0:0"), Range("A1048577"), Range("XFE1"),
     Range("A0", "B2") -> 1004; Range("A:A") and Range("XFD1048576") run.
     Range("A1").Offset(-1, 0), Range("A1").Offset(0, -1) -> 1004;
     Range("B2").Offset(-1, -1) runs. Resize(0, 1), Resize(1, 0), Resize(0),
     Resize(-1, 1) -> 1004. ActiveDocument.Range(-1, 0), Range(0, -1),
     Range(1, 0) -> 4608 in Word; Range(0, 0) runs.
 - sheet-name-invalid
     Worksheets(1).Name = "a:b", "", a 32-character name, or a name holding
     any of : \\ / ? * [ ] -> 1004. A 31-character name runs.
 - multi-cell-range-as-scalar
     A multi-cell address literal is an array when read as a value:
     `s = Range("A1:B2")` with s As String (or Long, Integer, Double,
     Boolean, Date), `s = Range("A1:B2").Value`, `Range("A1:B2") = 5`,
     `< 5`, `+ 1`, `& "x"` -> 13. A single-cell address runs.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...completion.member_access import MemberCompletionContext, resolve_receiver_type_at
from ...constants.integer_constant_expression import parse_vba_integer_literal
from ...host.host_model import (
    HostObjectModel,
    get_host_members,
    get_host_type,
    resolve_host_global,
    resolve_host_global_member,
    resolve_host_member,
)
from ...js_compat import JS_WHITESPACE, utf16_length
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ProcedureNode, Span
from ...symbols.symbol_model import ModuleSymbols
from ...types.type_inference import procedure_symbol_for, type_environment_for
from ...types.type_names import normalize_type
from ..call_extraction import string_literal_value
from ..context import PushFn
from ..walker import (
    ProcedureStatementVisitor,
    bare_assignment_target,
    first_executable_token_index,
    match_paren_from,
    statement_tokens,
    token_name,
    token_text,
)

_EXCEL_MAX_ROW = 1048576
_EXCEL_MAX_COLUMN = 16384
_SHEET_NAME_MAX = 31
_SHEET_NAME_FORBIDDEN = re.compile(r"[:\\/?*\[\]]")

_SCALAR_TYPES: frozenset[str] = frozenset(
    {
        "string", "long", "integer", "double", "single", "boolean", "date", "byte", "currency",
        "longlong",
    }
)

_SCALAR_OPERATORS: frozenset[str] = frozenset(
    {"=", "<", ">", "<=", ">=", "<>", "+", "-", "*", "/", "\\", "&", "^"}
)

# JavaScript's `\s` rather than Python's: the sheet-prefix pattern must split an
# address exactly where upstream's does. No character of the set is special
# inside a class.
_SHEET_PREFIX = re.compile("^(?:'[^']*'|[^!'" + JS_WHITESPACE + "]+)!")
# `[0-9]` and `\Z`, not `\d` and `$`: JavaScript's `\d` is ASCII only and its
# `$` never matches before a trailing newline.
_A1_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)\Z")
_A1_COLUMN = re.compile(r"^\$?([A-Za-z]{1,3})\Z")
_A1_ROW = re.compile(r"^\$?([0-9]+)\Z")

_CONDITION_HEADS: frozenset[str] = frozenset(
    {"if", "elseif", "while", "until", "do", "loop", "select", "case"}
)


@dataclass(frozen=True, slots=True)
class _HostCallee:
    # The member's own name as written.
    name: str
    # Qualified host type the callee's value has, when the model says.
    returns: str | None
    # Qualified host type of the receiver, or 'global' for a bare name.
    receiver: str
    # Index of the name token.
    name_index: int
    # Index of the `(` after the name, or -1 for a paren-less statement call.
    open_index: int
    # Index of the matching `)`, or the last token for a paren-less call.
    close_index: int
    # Top-level argument groups.
    args: list[list[VbaToken]]


@dataclass(frozen=True, slots=True)
class _RaisedError:
    number: str
    text: str


@dataclass(frozen=True, slots=True)
class _CellOrigin:
    row: int
    column: int
    text: str


@dataclass(frozen=True, slots=True)
class _A1Area:
    text: str
    valid: bool
    multi_cell: bool
    row: int | None = None
    column: int | None = None


def check_host_arguments(
    source: str, symbols: ModuleSymbols, member_ctx: MemberCompletionContext, push: PushFn
) -> ProcedureStatementVisitor:
    """Per-statement rule: host collection indexes, cell coordinates, address
    literals, sheet names and multi-cell values the code proves wrong."""
    model = member_ctx.model
    host_name = model.get("hostName") if model is not None else None
    host = host_name if host_name is not None else "Excel"
    if host != "Excel" and host != "Word" and host != "PowerPoint":
        return lambda proc: None
    module_names: set[str] = set()
    for child in symbols.root.children or []:
        module_names.add(child.name.lower())

    def factory(proc: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, proc)
        source_names = set(module_names)
        proc_sym = procedure_symbol_for(symbols, proc)
        for child in (proc_sym.children if proc_sym is not None else None) or []:
            source_names.add(child.name.lower())

        # A single-line If is one span here: its branches would otherwise be
        # walked twice, once inside the whole statement and once on their own.
        def visitor(stmt: LeafStatementNode) -> None:
            _check_span(source, stmt.span, host, model, member_ctx, env, source_names, push)

        return visitor

    return factory


def _check_span(
    source: str,
    span: Span,
    host: str,
    model: HostObjectModel | None,
    member_ctx: MemberCompletionContext,
    env: Mapping[str, str],
    source_names: AbstractSet[str],
    push: PushFn,
) -> None:
    toks = statement_tokens(source, span)

    def at(first: int, last: int) -> Span:
        return Span(span.start + toks[first].start, span.start + toks[last].end)

    if host == "Excel":
        _check_sheet_name_assignment(source, span, toks, member_ctx, push)
    for i in range(len(toks)):
        callee = _host_callee_at(source, span, toks, i, model, member_ctx, source_names)
        if callee is None:
            continue
        callee_span = at(callee.name_index, callee.close_index)
        lower = callee.name.lower()
        # Index 0 into a 1-based collection, `Worksheets(0)` or `Worksheets.Item(0)`.
        collection: str | None = None
        if lower == "item" and _is_collection_type(callee.receiver, model):
            collection = callee.receiver
        elif (
            callee.returns
            and _is_collection_type(callee.returns, model)
            and callee.open_index > 0
        ):
            collection = callee.returns
        if collection and len(callee.args) == 1 and lower != "cells" and lower != "range":
            index = _integer_literal_value(callee.args[0])
            if index is not None and index < 1:
                error = _collection_index_error(host, collection, model)
                push(
                    "hostArgumentOutOfRange",
                    f"Index {index} is never an element: {host} collections start at 1. "
                    f"This will raise Run-time error '{error.number}': {error.text}.",
                    at(callee.open_index + 1, callee.close_index - 1),
                )
                continue
        if host == "Excel":
            _check_excel_callee(source, span, toks, callee, callee_span, env, push)
        elif host == "Word":
            if lower == "range" and callee.receiver == "Word.Document" and callee.open_index > 0:
                start = _integer_literal_value(callee.args[0]) if len(callee.args) > 0 else None
                end = _integer_literal_value(callee.args[1]) if len(callee.args) > 1 else None
                bad = (
                    (start is not None and start < 0)
                    or (end is not None and end < 0)
                    or (start is not None and end is not None and end < start)
                )
                if bad:
                    push(
                        "hostArgumentOutOfRange",
                        "Document.Range takes character positions from 0, with End at or "
                        "after Start. This will raise Run-time error '4608': Value out of range.",
                        at(callee.open_index + 1, callee.close_index - 1),
                    )
        elif lower == "add" and callee.receiver == "PowerPoint.Slides" and len(callee.args) >= 1:
            index = _integer_literal_value(callee.args[0])
            if index is not None and index < 1:
                push(
                    "hostArgumentOutOfRange",
                    "Slides.Add places the new slide at Index, which starts at 1. This will "
                    "raise Run-time error '-2147188160': Integer out of range.",
                    _arg_span(span, callee.args[0]),
                )


def _host_callee_at(
    source: str,
    span: Span,
    toks: Sequence[VbaToken],
    i: int,
    model: HostObjectModel | None,
    member_ctx: MemberCompletionContext,
    source_names: AbstractSet[str],
) -> _HostCallee | None:
    """The host member a name-with-arguments at `toks[i]` calls: a bare host
    global (`Worksheets(0)`, `Cells(0, 1)`), a member of the hidden Global
    interface (`Rows(0)`, `Names(0)`), or a member of a resolved receiver
    (`ActiveDocument.Paragraphs(0)`, `Range("A1").Offset(-1, 0)`). A name the
    module or procedure declares is the source's, not the host's."""
    name = token_name(toks[i])
    if not name:
        return None
    qualified = i >= 1 and toks[i - 1].raw_text == "."
    parenthesized = i + 1 < len(toks) and toks[i + 1].raw_text == "("
    open_index = -1
    if parenthesized:
        open_index = i + 1
        close_index = match_paren_from(toks, open_index)
        if close_index < 0:
            return None
        args = _split_top_level(toks[open_index + 1 : close_index])
    elif (
        qualified
        and i == _first_executable_token_index_of_member_call(toks, i)
        and i + 1 < len(toks)
        and toks[i + 1].raw_text != "."
        and toks[i + 1].raw_text != "="
    ):
        # Statement form: `ActivePresentation.Slides.Add 0, ppLayoutBlank`.
        close_index = len(toks) - 1
        args = _split_top_level(toks[i + 1 :])
    else:
        return None
    if not qualified:
        if name.lower() in source_names or not parenthesized:
            return None
        global_type = resolve_host_global(name, model)
        if global_type:
            return _HostCallee(name, global_type, "global", i, open_index, close_index, args)
        global_member = resolve_host_global_member(name, model)
        if global_member is not None:
            return _HostCallee(
                name, global_member.get("returns"), "global", i, open_index, close_index, args
            )
        return None
    receiver = _host_receiver_type(
        resolve_receiver_type_at(source, span.start + toks[i - 1].end, member_ctx), model
    )
    if not receiver:
        return None
    member = resolve_host_member(receiver, name, model)
    if member is None:
        return None
    return _HostCallee(name, member.get("returns"), receiver, i, open_index, close_index, args)


def _host_receiver_type(resolved: str | None, model: HostObjectModel | None) -> str | None:
    """The host type a resolved receiver names. A one-part union - what a
    collection's Object-declared Item gives, `Worksheets(1)` (XLIDE issue #114) - is
    its part; a union of several types is not judged."""
    if not resolved:
        return None
    parts = resolved[len("union:") :].split("|") if resolved.startswith("union:") else [resolved]
    if len(parts) != 1 or get_host_type(parts[0], model) is None:
        return None
    return parts[0]


def _first_executable_token_index_of_member_call(toks: Sequence[VbaToken], name_index: int) -> int:
    """The member call's name index when the statement is `a.b.Name args`, else -1."""
    j = name_index
    while j >= 2 and toks[j - 1].raw_text == "." and token_name(toks[j - 2]):
        j -= 2
        if j >= 1 and toks[j - 1].raw_text == ")":
            open_index = _open_paren_for(toks, j - 1)
            if open_index < 1:
                return -1
            j = open_index
    return name_index if j == first_executable_token_index(toks) else -1


def _open_paren_for(toks: Sequence[VbaToken], close: int) -> int:
    """Index of the first `(` whose match is `toks[close]`, or -1.

    Upstream tries every `(` from the start of the statement. Only one `(` can
    match a given `)`, and walking back from it finds that one without the
    quadratic scan, which cost a 3,000-link `.Offset(1, 0)` chain 3.5 s."""
    if not 0 <= close < len(toks) or toks[close].raw_text != ")":
        return -1
    depth = 0
    for k in range(close, -1, -1):
        raw = toks[k].raw_text
        if raw == ")":
            depth += 1
        elif raw == "(":
            depth -= 1
            if depth == 0:
                return k if match_paren_from(toks, k) == close else -1
    return -1


def _is_collection_type(type_name: str, model: HostObjectModel | None) -> bool:
    members = get_host_members(type_name, model)
    return any(m["name"] == "Item" for m in members) and any(m["name"] == "Count" for m in members)


def _collection_index_error(
    host: str, collection: str, model: HostObjectModel | None
) -> _RaisedError:
    bare = collection[collection.find(".") + 1 :]
    if host == "PowerPoint":
        return _RaisedError("-2147188160", "Integer out of range")
    if bare == "Shapes":
        return _RaisedError(
            "-2147024809", "The index into the specified collection is out of bounds"
        )
    if host == "Word":
        return _RaisedError("5941", "The requested member of the collection does not exist")
    if bare == "PivotTables" or bare == "Range":
        return _RaisedError("1004", "Application-defined or object-defined error")
    item = resolve_host_member(collection, "Item", model)
    if item is not None and item.get("returns") == "Excel.Range":
        return _RaisedError("1004", "Application-defined or object-defined error")
    return _RaisedError("9", "Subscript out of range")


def _check_excel_callee(
    source: str,
    span: Span,
    toks: Sequence[VbaToken],
    callee: _HostCallee,
    callee_span: Span,
    env: Mapping[str, str],
    push: PushFn,
) -> None:
    lower = callee.name.lower()
    if callee.open_index < 0:
        return
    args_span = (
        Span(
            span.start + toks[callee.open_index + 1].start,
            span.start + toks[callee.close_index - 1].end,
        )
        if callee.close_index > callee.open_index + 1
        else callee_span
    )
    if lower == "cells" and callee.returns == "Excel.Range":
        for arg in callee.args:
            value = _integer_literal_value(arg)
            if value is not None and value < 1:
                push(
                    "hostArgumentOutOfRange",
                    f"Cells takes a row and a column of at least 1; {value} names no cell. "
                    "This will raise Run-time error '1004': Application-defined or "
                    "object-defined error.",
                    _arg_span(span, arg),
                )
                return
        return
    if lower == "resize" and callee.receiver == "Excel.Range":
        for arg in callee.args:
            value = _integer_literal_value(arg)
            if value is not None and value < 1:
                push(
                    "hostArgumentOutOfRange",
                    f"Resize needs at least one row and one column; {value} gives none. "
                    "This will raise Run-time error '1004': Application-defined or "
                    "object-defined error.",
                    _arg_span(span, arg),
                )
                return
        return
    if lower == "offset" and callee.receiver == "Excel.Range":
        origin = _single_cell_receiver(toks, callee.name_index - 1)
        if origin is None:
            return
        row_offset = _integer_literal_value(callee.args[0]) if len(callee.args) > 0 else 0
        column_offset = _integer_literal_value(callee.args[1]) if len(callee.args) > 1 else 0
        if row_offset is None or column_offset is None:
            return
        row = _js_add(origin.row, row_offset)
        column = _js_add(origin.column, column_offset)
        if row < 1 or column < 1:
            push(
                "hostArgumentOutOfRange",
                f"Offset({row_offset}, {column_offset}) from {origin.text} lands at row {row}, "
                f"column {column}, off the sheet. This will raise Run-time error '1004': "
                "Application-defined or object-defined error.",
                args_span,
            )
        return
    if lower == "range" and callee.returns == "Excel.Range":
        areas = [
            _parse_a1_address(string_literal_value(arg[0].raw_text))
            if len(arg) == 1 and arg[0].kind is TokenKind.STRING_LITERAL
            else None
            for arg in callee.args
        ]
        for k in range(len(callee.args)):
            area = areas[k]
            if area is not None and not area.valid:
                push(
                    "hostArgumentOutOfRange",
                    f'"{area.text}" is not a cell address Excel accepts: rows run 1 to '
                    f"{_EXCEL_MAX_ROW} and columns A to XFD. This will raise Run-time error "
                    "'1004': Method 'Range' of object failed.",
                    _arg_span(span, callee.args[k]),
                )
                return
        first = areas[0] if len(areas) > 0 else None
        if len(callee.args) == 1 and first is not None and first.valid and first.multi_cell:
            _check_multi_cell_as_scalar(source, span, toks, callee, first.text, env, push)


def _check_multi_cell_as_scalar(
    source: str,
    span: Span,
    toks: Sequence[VbaToken],
    callee: _HostCallee,
    address: str,
    env: Mapping[str, str],
    push: PushFn,
) -> None:
    """`Range("A1:B2")` read as a value is a two-dimensional array (XLIDE issue #122):
    assigned to a scalar variable, or combined with a scalar operator, it is a
    type mismatch. `.Value` after it changes nothing."""
    end = callee.close_index
    if _raw_at(toks, end + 1) == "." and token_text(_token_at(toks, end + 2)) in (
        "value",
        "value2",
    ):
        end += 2
    if _raw_at(toks, end + 1) == "." or _raw_at(toks, end + 1) == "(":
        return  # a member or an index: not the range read as a value
    start = (
        callee.name_index
        if callee.receiver == "global"
        else _receiver_start(toks, callee.name_index)
    )
    value_span = Span(span.start + toks[start].start, span.start + toks[end].end)

    def message(use: str) -> str:
        return (
            f'Range("{address}") read as a value is a two-dimensional array, {use}. '
            "This will raise Run-time error '13': Type mismatch."
        )

    bare = bare_assignment_target(source, span)
    if bare is not None:
        eq = _index_of_equals(toks)
        if eq == start - 1 and end == len(toks) - 1:
            declared = env.get(bare[0].lower())
            target = normalize_type(declared)
            if target and target in _SCALAR_TYPES:
                push(
                    "multiCellRangeAsScalar",
                    message(f"which a {declared} variable cannot hold"),
                    value_span,
                )
            return
        if eq < 0 or start <= eq:
            return
    else:
        head = token_text(toks[first_executable_token_index(toks)])
        if head not in _CONDITION_HEADS:
            return
        # Only the condition of a one-line If is judged here: a range after Then or
        # Else belongs to that branch's own statement, `If r Is Nothing Then Set r =
        # ws.Range("A1:P36")` (XLIDE issue #140).
        then = (
            next((i for i, tok in enumerate(toks) if token_text(tok) == "then"), -1)
            if head == "if"
            else -1
        )
        if then > 0 and start > then:
            return
    # The operator on either side, never the assignment's own `=`.
    eq_index = _index_of_equals(toks) if bare is not None else -1
    before = None if start - 1 == eq_index else _token_at(toks, start - 1)
    after = _token_at(toks, end + 1)
    operator = next(
        (
            tok
            for tok in (after, before)
            if tok is not None
            and (
                (tok.kind is TokenKind.OPERATOR and tok.raw_text in _SCALAR_OPERATORS)
                or token_text(tok) == "mod"
            )
        ),
        None,
    )
    if operator is not None:
        push(
            "multiCellRangeAsScalar",
            message(f"which '{operator.raw_text}' cannot combine with a scalar"),
            value_span,
        )


def _check_sheet_name_assignment(
    source: str,
    span: Span,
    toks: Sequence[VbaToken],
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> None:
    """`Worksheets(1).Name = "a:b"`: the receiver's type and the literal decide."""
    n = len(toks)
    if (
        n < 4
        or toks[n - 1].kind is not TokenKind.STRING_LITERAL
        or toks[n - 2].raw_text != "="
        or token_text(toks[n - 3]) != "name"
        or toks[n - 4].raw_text != "."
    ):
        return
    resolved = resolve_receiver_type_at(source, span.start + toks[n - 4].end, member_ctx)
    # `Worksheets(1)` is a one-part union, `Sheets(1)` a Worksheet-or-Chart union: both are sheets.
    parts: list[str] = []
    if resolved:
        parts = (
            resolved[len("union:") :].split("|") if resolved.startswith("union:") else [resolved]
        )
    if len(parts) == 0 or not all(
        part == "Excel.Worksheet" or part == "Excel.Chart" for part in parts
    ):
        return
    name = string_literal_value(toks[n - 1].raw_text)
    # JavaScript measures a string in UTF-16 code units, as Excel counts a name.
    length = utf16_length(name)
    problem: str | None = None
    if length == 0:
        problem = "a sheet name cannot be blank"
    elif length > _SHEET_NAME_MAX:
        problem = (
            f"a sheet name has at most {_SHEET_NAME_MAX} characters, and this one has {length}"
        )
    elif _SHEET_NAME_FORBIDDEN.search(name) is not None:
        problem = "a sheet name cannot contain any of : \\ / ? * [ ]"
    if problem:
        push(
            "sheetNameInvalid",
            f"Excel refuses this name: {problem}. This will raise Run-time error '1004': "
            "You typed an invalid name for a sheet or chart.",
            Span(span.start + toks[n - 1].start, span.start + toks[n - 1].end),
        )


def _single_cell_receiver(toks: Sequence[VbaToken], dot_index: int) -> _CellOrigin | None:
    """The single-cell literal `Range("B2")` ending at `toks[close_index]`, when
    that is the receiver."""
    if _raw_at(toks, dot_index) != "." or _raw_at(toks, dot_index - 1) != ")":
        return None
    close = dot_index - 1
    open_index = _open_paren_for(toks, close)
    if (
        open_index < 1
        or token_text(toks[open_index - 1]) != "range"
        or close != open_index + 2
        or toks[open_index + 1].kind is not TokenKind.STRING_LITERAL
    ):
        return None
    area = _parse_a1_address(string_literal_value(toks[open_index + 1].raw_text))
    if (
        area is None
        or not area.valid
        or area.multi_cell
        or area.row is None
        or area.column is None
    ):
        return None
    return _CellOrigin(area.row, area.column, f'Range("{area.text}")')


def _receiver_start(toks: Sequence[VbaToken], name_index: int) -> int:
    """Index of the first token of the receiver chain that ends at the dot before `name_index`."""
    j = name_index
    while j >= 2 and toks[j - 1].raw_text == ".":
        j -= 1
        if toks[j - 1].raw_text == ")":
            open_index = _open_paren_for(toks, j - 1)
            j = open_index if open_index >= 1 else j
        j -= 1
    return j


def _parse_a1_address(text: str) -> _A1Area | None:
    """Parses an A1-style address literal: `A1`, `$A$1`, `A1:B2`, `A:A`, `1:1`,
    with an optional `Sheet1!` or `'My Sheet'!` prefix. Anything else (a name,
    an R1C1 address, a union) is not judged."""
    body = _SHEET_PREFIX.sub("", text, count=1)
    parts = body.split(":")
    if len(parts) > 2:
        return None
    cell_matches = [_A1_CELL.match(part) for part in parts]
    cells = [match for match in cell_matches if match is not None]
    if len(cells) == len(cell_matches):
        rows = [_decimal_number(match.group(2)) for match in cells]
        columns = [_column_number(match.group(1)) for match in cells]
        valid = all(1 <= row <= _EXCEL_MAX_ROW for row in rows) and all(
            1 <= column <= _EXCEL_MAX_COLUMN for column in columns
        )
        multi_cell = len(cells) == 2 and (rows[0] != rows[1] or columns[0] != columns[1])
        return _A1Area(text, valid, multi_cell, rows[0], columns[0])
    if len(parts) == 2:
        column_matches = [_A1_COLUMN.match(part) for part in parts]
        columns_only = [match for match in column_matches if match is not None]
        if len(columns_only) == len(column_matches):
            valid = all(
                _column_number(match.group(1)) <= _EXCEL_MAX_COLUMN for match in columns_only
            )
            return _A1Area(text, valid, True)
        row_matches = [_A1_ROW.match(part) for part in parts]
        rows_only = [match for match in row_matches if match is not None]
        if len(rows_only) == len(row_matches):
            valid = all(
                1 <= _decimal_number(match.group(1)) <= _EXCEL_MAX_ROW for match in rows_only
            )
            return _A1Area(text, valid, True)
    return None


def _decimal_number(digits: str) -> int:
    """The value of a run of ASCII digits, as JavaScript's Number() reads it for
    every comparison made here. A run too long for int() to convert (Python caps
    the digits it will read) names a row past any sheet, so a stand-in past the
    last row answers the same."""
    significant = digits.lstrip("0")
    if len(significant) > 15:
        return _EXCEL_MAX_ROW + 1
    return int(significant) if significant else 0


def _column_number(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _integer_literal_value(arg: Sequence[VbaToken]) -> int | None:
    """The whole-number value of an argument that is a literal, optionally negated."""
    toks = [tok for tok in arg if tok.kind is not TokenKind.COMMENT]
    if len(toks) == 1 and toks[0].kind is TokenKind.INTEGER_LITERAL:
        return parse_vba_integer_literal(toks[0].raw_text)
    if len(toks) == 2 and toks[0].raw_text == "-" and toks[1].kind is TokenKind.INTEGER_LITERAL:
        value = parse_vba_integer_literal(toks[1].raw_text)
        return None if value is None else -value
    return None


def _split_top_level(toks: Sequence[VbaToken]) -> list[list[VbaToken]]:
    out: list[list[VbaToken]] = []
    current: list[VbaToken] = []
    depth = 0
    for tok in toks:
        if tok.raw_text == "(":
            depth += 1
        elif tok.raw_text == ")":
            depth -= 1
        if tok.raw_text == "," and depth == 0:
            out.append(current)
            current = []
            continue
        current.append(tok)
    if len(current) > 0 or len(out) > 0:
        out.append(current)
    return out


def _arg_span(span: Span, arg: Sequence[VbaToken]) -> Span:
    return Span(span.start + arg[0].start, span.start + arg[-1].end)


def _index_of_equals(toks: Sequence[VbaToken]) -> int:
    """Index of the first `=` token, or -1."""
    for k, tok in enumerate(toks):
        if tok.raw_text == "=":
            return k
    return -1


def _token_at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """`toks[index]`, or None out of range the way a JavaScript index reads undefined."""
    return toks[index] if 0 <= index < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], index: int) -> str | None:
    tok = _token_at(toks, index)
    return tok.raw_text if tok is not None else None


def _js_add(a: int, b: int) -> int:
    """`a + b` as JavaScript computes it: in doubles, so a sum past 2**53 rounds
    the way the upstream message prints it. Both operands are safe integers."""
    return int(float(a) + float(b))


# --- sync stubs (2f49b93): replaced as each group is ported ---


def range_method_owner(*args: object, **kwargs: object) -> None:
    return None


def literal_intersect_is_nothing(*args: object, **kwargs: object) -> None:
    return None


class WorkbookSheetsCheck:
    pass


def workbook_sheets_to_check(*args: object, **kwargs: object) -> None:
    return None
