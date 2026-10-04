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
     Cells(0, 1), Cells(1, 0), Cells(0) -> 1004. Range("$A$0"),
     Range("A$0"), Range("Sheet1!$A$0"), Range("0:0"), Range("$XFE:$XFE")
     -> 1004; Range("A:A") and Range("XFD1048576") run. Range("A0"),
     Range("A1048577") and Range("XFE1") raise 1004 too, but only while
     no workbook name is spelled that way, so they are not judged (measured
     2026-10-01): A0, XFE1, A1048577 and XFE are names Names.Add accepts,
     and Range then finds them, alone, in "A0:D4", in Range("A0", "D4")
     and after "Sheet1!". A name cannot hold `$` or start with a digit.
     Range("A1").Offset(-1, 0), Range("A1").Offset(0, -1) -> 1004;
     Range("B2").Offset(-1, -1) runs. Resize(0, 1), Resize(1, 0), Resize(0),
     Resize(-1, 1) -> 1004. Past the bottom and right edges (XLIDE issue #182,
     measured 2026-09-29): Cells(1048577, 1), Cells(1, 16385), Rows(1048577),
     Columns(16385), Range("A1048576").Offset(1, 0),
     Range("A2").Resize(1048576), Range("B2").Cells(1048576, 1) and
     Range("A5").Rows(1048573) -> 1004; Cells(1048576, 16384) and
     Range("A2").Offset(0) or Offset(-1) run. On a range, Cells, Item, Rows
     and Columns count from its top-left cell (XLIDE issue #275, measured
     2026-10-01): Range("C3").Cells(-1, -1) is A1, Range("B2:C3").Rows(0)
     is B1:C1, and one index w columns wide is Cells((i - 1) \\ w + 1,
     (i - 1) Mod w + 1), so Range("B2:C3").Cells(0) is A2. Each raises
     1004 only where it lands above row 1 or left of column A:
     Range("A1").Cells(0), Range("B2").Cells(-1). A range the code does not
     spell out, a variable or ActiveCell, is not judged.
     ActiveDocument.Range(-1, 0), Range(0, -1),
     Range(1, 0) -> 4608 in Word; Range(0, 0) runs.
     Range("") and Range(" ") -> 1004 (XLIDE issue #276). A name that looks
     like an address past the sheet, ZZZZ1, may be a workbook name, and
     Range("A1:ZZZZ1") then runs, so it is not judged.
     ActiveSheet.Shapes(0) and Sheets(1).Shapes(0) -> -2147024809: a
     Worksheet and a Chart both have Shapes (XLIDE issue #276).
 - sheet-name-invalid
     Worksheets(1).Name = "a:b", "", a 32-character name, or a name holding
     any of : \\ / ? * [ ] -> 1004. A 31-character name runs. "History" in
     any case, and an apostrophe first or last, -> 1004, and so does a
     name String$, Space$ or & spell out (XLIDE issue #276, measured
     2026-10-01); "History ", "History1" and "a'b" run.
 - multi-cell-range-as-scalar
     A multi-cell address literal is an array when read as a value:
     `s = Range("A1:B2")` with s As String (or Long, Integer, Double,
     Boolean, Date), `s = Range("A1:B2").Value`, `Range("A1:B2") = 5`,
     `< 5`, `+ 1`, `& "x"` -> 13. A single-cell address runs.

Upstream walks block bodies by recursion; the walks here run on explicit
stacks (Python stops at 1000 frames), and the Range chain a callee is read
through is followed iteratively and memoized per statement, with the same
results.
"""

from __future__ import annotations

import dataclasses
import math
import re
import unicodedata
from collections.abc import Callable, Iterator, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...completion.member_access import MemberCompletionContext, resolve_receiver_type_at
from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import parse_vba_integer_literal
from ...flow.procedure_labels import jump_target_label_declaration
from ...host.host_model import (
    HostObjectModel,
    get_host_members,
    get_host_type,
    resolve_host_global,
    resolve_host_global_member,
    resolve_host_member,
)
from ...host.word_builtin_styles import WORD_BUILTIN_STYLES
from ...js_compat import JS_WHITESPACE, js_number, js_number_to_string, js_trim, utf16_length
from ...lexer.token_helpers import split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    LeafStatementNode,
    ProcedureNode,
    Span,
    WithBlockNode,
    is_leaf_statement,
)
from ...symbols.sheet_changes import SheetChanges, WorkbookSheetInfo
from ...symbols.symbol_model import ModuleSymbols
from ...types.type_inference import (
    known_local_literal_values_at,
    procedure_symbol_for,
    type_environment_for,
)
from ...types.type_names import normalize_type
from ..call_extraction import string_literal_value
from ..context import AnalyzeModuleOptions, PushFn
from ..known_locals import KnownLocalValue
from ..known_string_calls import StringFoldContext, fold_string_expression
from ..loop_counters import check_each_counter_pass, loop_counters_at
from ..walker import (
    ProcedureStatementVisitor,
    bare_assignment_target,
    block_header_line_span,
    first_executable_token_index,
    match_paren_from,
    statement_tokens,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .worksheet_function_arguments import worksheet_function_refusal

# A number the code fixes: JavaScript's one number type, as int or float here.
Number = int | float
ValueOf = Callable[[Sequence[VbaToken]], "int | float | None"]
StringOf = Callable[[Sequence[VbaToken]], "str | None"]

_EXCEL_MAX_ROW = 1048576
_EXCEL_MAX_COLUMN = 16384

# Range members whose arguments are a cell address, a row and column, an
# offset or a size, not an index into the range they return. Each has its own
# check. `Range("A2").Offset(0)` and `Offset(-1)` run (XLIDE issue #182).
_RANGE_COORDINATE_MEMBERS: frozenset[str] = frozenset({"cells", "range", "offset", "resize"})
# The members that count from a range's own top-left cell (XLIDE issue #275).
_RANGE_RELATIVE_MEMBERS: frozenset[str] = frozenset({"cells", "item", "rows", "columns"})
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
_WS = JS_WHITESPACE
_SHEET_PREFIX = re.compile("^(?:'[^']*'|[^!'" + _WS + "]+)!")
# `[0-9]` and `\Z`, not `\d` and `$`: JavaScript's `\d` is ASCII only and its
# `$` never matches before a trailing newline.
_A1_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)\Z")
_A1_COLUMN = re.compile(r"^\$?([A-Za-z]{1,3})\Z")
_A1_ROW = re.compile(r"^\$?([0-9]+)\Z")
_R1C1_BODY = re.compile(r"^R[0-9]+C[0-9]+\Z", re.IGNORECASE | re.ASCII)
_AREAS_ADDRESS = re.compile(r"^\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?\Z")
_INTEGER_SUFFIX = re.compile(r"[%&^]\Z")
_PROTECT_IN_SOURCE = re.compile(r"\.[" + _WS + r"]*protect\b", re.IGNORECASE | re.ASCII)
_SHEETS_ADD_VALUE = re.compile(r"^(?:\w+\.)*(?:worksheets|sheets)\.add(?:\(.*\))?\Z", re.ASCII)
_SHEETS_CALL_VALUE = re.compile(r"^(?:\w+\.)*(?:worksheets|sheets)\(", re.ASCII)
_SHEETS_VALUE = re.compile(r"^(?:\w+\.)*(?:worksheets|sheets)", re.ASCII)
_LETTERS_COLUMN = re.compile(r"^\$?([A-Za-z]{1,3})\Z")
_LETTERS_ONLY = re.compile(r"^[A-Za-z]{1,3}\Z")
_ASCII_DIGIT = re.compile("[0-9]")
_LEADING_DIGIT = re.compile("^[0-9]")
_JS_SPACE = re.compile("[" + _WS + "]")
_R1C1_NAME = re.compile(r"^[Rr][0-9]+[Cc][0-9]+\Z")
_WHOLE_ROWS = re.compile(r"^\$?([0-9]+):\$?([0-9]+)\Z")
_WHOLE_COLUMNS = re.compile(r"^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})\Z")
_LINE_BREAKS = re.compile("\r\n|\r|\n")
_SENTENCE_STOPS = re.compile("[.!?]")

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
class _A1Area:
    text: str
    valid: bool
    multi_cell: bool
    # Empty or spaces only: no address and no name (XLIDE issue #276).
    blank: bool = False
    # Every part past the sheet could be a workbook name: it starts with a
    # letter and holds no `$`. A0, XFE1, A1048577 and XFE are names Excel
    # accepts, and Range finds them, alone, in `A0:D4` or after `Sheet1!`.
    may_be_name: bool = False
    row: int | None = None
    column: int | None = None
    # The second cell of `A1:B2`.
    end_row: int | None = None
    end_column: int | None = None


@dataclass(frozen=True, slots=True)
class _CellBlock:
    """A block of cells: its top-left row and column and its size."""

    row: Number
    column: Number
    rows: Number
    width: Number
    # The expression that names it, as written.
    text: str = ""
    # Rows or columns from EntireRow, EntireColumn, Rows(n) or Columns(n): one
    # index in Item then counts rows or columns: `Columns(4).EntireColumn.Item(0)`
    # is column C (XLIDE issue #556, measured in Excel 16.0).
    mode: str | None = None


def _n(value: Number) -> str:
    """A number as JavaScript's template literals print it."""
    return js_number_to_string(value)


def _at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """`toks[index]`, or None out of range the way a JavaScript index reads undefined."""
    return toks[index] if 0 <= index < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], index: int) -> str | None:
    tok = _at(toks, index)
    return tok.raw_text if tok is not None else None


def _kind_at(toks: Sequence[VbaToken], index: int) -> TokenKind | None:
    tok = _at(toks, index)
    return tok.kind if tok is not None else None


def _word(words: Sequence[str], index: int) -> str | None:
    return words[index] if 0 <= index < len(words) else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name is not None else None


def _is_integral(value: object) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, int):
        return True
    return isinstance(value, float) and math.isfinite(value) and value.is_integer()


def _significant(arg: Sequence[VbaToken]) -> list[VbaToken]:
    return [tok for tok in arg if tok.kind is not TokenKind.COMMENT]


def _is_inactive(activity: ConditionalActivityTracker | None, node: BodyNode) -> bool:
    return activity is not None and activity.is_inactive(node.span)


def _block_body(node: BodyNode) -> list[BodyNode] | None:
    body = getattr(node, "body", None)
    return body if isinstance(body, list) else None


def check_host_arguments(
    source: str,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    sheets: WorkbookSheetsCheck | None = None,
) -> ProcedureStatementVisitor:
    """Per-statement rule: host collection indexes, cell coordinates, address
    literals, sheet names and multi-cell values the code proves wrong."""
    model = member_ctx.model
    host_name = model.get("hostName") if model is not None else None
    host = host_name if host_name is not None else "Excel"
    if host != "Excel" and host != "Word" and host != "PowerPoint":
        return lambda proc: None
    workbook = sheets if host == "Excel" else None
    module_names: set[str] = set()
    module_arrays: set[str] = set()
    for child in symbols.root.children or []:
        if child.is_array:
            module_arrays.add(child.name.lower())
    for child in symbols.root.children or []:
        module_names.add(child.name.lower())
    # A Public procedure of another module named Cells, Worksheets or Range
    # takes the call from Excel's (XLIDE issue #280, measured in Excel 16.0).
    for project_type in member_ctx.project_class_members or []:
        if project_type.kind == "standardModule":
            for project_member in project_type.members:
                module_names.add(project_member.name.lower())

    def factory(proc: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        return _procedure_visitor(
            source, proc, host, model, member_ctx, activity, push, workbook, symbols, module_names, module_arrays
        )

    return factory


def _procedure_visitor(
    source: str,
    proc: ProcedureNode,
    host: str,
    model: HostObjectModel | None,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    workbook: WorkbookSheetsCheck | None,
    symbols: ModuleSymbols,
    module_names: AbstractSet[str],
    module_arrays: AbstractSet[str],
) -> Callable[[LeafStatementNode], None]:
    env = type_environment_for(symbols, proc)
    source_names = set(module_names)
    proc_sym = procedure_symbol_for(symbols, proc)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        source_names.add(child.name.lower())
    # A single-line If is one span here: its branches would otherwise be walked
    # twice, once inside the whole statement and once on their own. Array
    # variables, so a range read into one is named as an array. A local shadows
    # a module-level name.
    arrays = set(module_arrays)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if child.is_array:
            arrays.add(child.name.lower())
        else:
            arrays.discard(child.name.lower())
    # `Cells(r, 1)` inside `For r = 0 To 3`: each pass's value (XLIDE issue #200).
    counters = loop_counters_at(source, proc.body, activity)
    # A local known to hold one value here: `r = 0` then `Cells(r, 1)` (XLIDE
    # issue #345, measured in Excel 16.0). Each per-procedure view is built on
    # first use, as upstream's `??=` builds it.
    values_at: list[Callable[[LeafStatementNode], Mapping[str, KnownLocalValue]]] = []
    documents_at: list[dict[int, dict[str, _NewDocument]]] = []
    sheets_at: list[dict[int, _ActiveSheetState]] = []
    facts_at: list[dict[int, _SheetFacts]] = []
    # A sheet the code just protected (XLIDE issue #471).
    if host == "Excel" and _PROTECT_IN_SOURCE.search(source[proc.span.start : proc.span.end]) is not None:
        _check_protected_sheets(source, proc, activity, push)

    def visitor(stmt: LeafStatementNode) -> None:
        stmt_counters = counters.get(stmt)
        if not values_at:
            values_at.append(known_local_literal_values_at(source, proc, symbols, activity))
        known = values_at[0](stmt)

        def held_by(arg: Sequence[VbaToken]) -> KnownLocalValue | None:
            toks = _significant(arg)
            held = known.get(_lower_name(toks[0]) or "") if len(toks) == 1 else None
            return held if held is not None and not held.content_mutated else None

        def known_number(arg: Sequence[VbaToken]) -> Number | None:
            held = held_by(arg)
            if held is not None and held.kind == "number" and _is_integral(held.value):
                assert isinstance(held.value, (int, float))
                return int(held.value)
            return None

        def string_of(arg: Sequence[VbaToken]) -> str | None:
            literal = _literal_string(arg)
            if literal is not None:
                return literal
            held = held_by(arg)
            return held.value if held is not None and held.kind == "string" and isinstance(held.value, str) else None

        def literal_or_known(arg: Sequence[VbaToken]) -> Number | None:
            literal = _integer_literal_value(arg)
            return literal if literal is not None else known_number(arg)

        # A document the procedure just added (XLIDE issue #497).
        if host == "Word":
            if not documents_at:
                documents_at.append(_new_documents_at(source, proc, activity))
            documents = documents_at[0].get(id(stmt))
            if documents:
                _check_new_document_uses(stmt.span, statement_tokens(source, stmt.span), documents, literal_or_known, push)
        # A sheet the code knows is not active (XLIDE issue #470).
        if host == "Excel":
            if not sheets_at:
                sheets_at.append(_active_sheets_at(source, proc, activity))
            sheet_state = sheets_at[0].get(id(stmt))
            if sheet_state is not None:
                _check_unqualified_corners(stmt.span, statement_tokens(source, stmt.span), sheet_state, push)
        # Intersect and Union of literal ranges, and a sheet the code just added
        # (XLIDE issue #472).
        if host == "Excel":
            if not facts_at:
                facts_at.append(_sheet_facts_at(source, proc, activity))
            _check_sheet_facts(stmt.span, statement_tokens(source, stmt.span), facts_at[0].get(id(stmt)), push)
        if not stmt_counters:
            _check_span(
                source, stmt.span, host, model, member_ctx, env, arrays, source_names, literal_or_known, push, string_of
            )
            if workbook is not None:
                _check_workbook_sheet_access(source, stmt.span, workbook, source_names, literal_or_known, push)
            return

        def each_pass(values: Mapping[str, Number], report: PushFn) -> None:
            def value_of(arg: Sequence[VbaToken]) -> Number | None:
                literal = _integer_literal_value(arg)
                if literal is not None or len(values) == 0:
                    return literal if literal is not None else known_number(arg)
                toks = _significant(arg)
                if len(toks) != 1:
                    return None
                counted = values.get(_lower_name(toks[0]) or "")
                return counted if counted is not None else known_number(arg)

            _check_span(
                source, stmt.span, host, model, member_ctx, env, arrays, source_names, value_of, report, string_of
            )
            if workbook is not None:
                _check_workbook_sheet_access(source, stmt.span, workbook, source_names, value_of, report)

        check_each_counter_pass(source, stmt.span, stmt_counters, lambda atom, counter: None, each_pass, push)

    return visitor


@dataclass(frozen=True, slots=True)
class _ActiveSheetState:
    """What a statement knows of the active sheet: the sheet variables known not
    to be it, and a With's subject."""

    not_active: frozenset[str]
    subject: str | None = None


_ACTIVATING_CALLS: frozenset[str] = frozenset({"activate", "select", "add", "copy", "move"})


def _active_sheets_at(
    source: str, proc: ProcedureNode, activity: ConditionalActivityTracker | None
) -> dict[int, _ActiveSheetState]:
    """The sheet variables each statement knows are not the active sheet (XLIDE
    issue #470, measured in Excel 16.0): `Set w2 = Worksheets.Add` activates the
    new sheet, so every sheet the code held before is not active, and `w2` is.
    `Set w1 = ActiveSheet` and `w1.Activate` make w1 the active one. Any other
    statement that may activate a sheet, a call, a label, or a block other than
    a With ends what is known. Keyed by the statement node's id()."""
    out: dict[int, _ActiveSheetState] = {}
    sheets: set[str] = set()
    not_active: set[str] = set()
    stack: list[tuple[Iterator[BodyNode], str | None]] = [(iter(proc.body), None)]
    while stack:
        nodes, subject = stack[-1]
        for node in nodes:
            if _is_inactive(activity, node):
                continue
            if not is_leaf_statement(node):
                # The body runs from the state the block is entered with; after
                # it, what it may have activated is not known.
                body = _block_body(node)
                if body is not None:
                    inner_subject = subject
                    if isinstance(node, WithBlockNode):
                        header = statement_tokens_after_leading_label(
                            source, block_header_line_span(source, node.span)
                        )
                        if len(header) == 2:
                            inner_subject = _lower_name(header[1])
                    not_active = set(not_active)
                    stack.append((iter(body), inner_subject))
                    break
                not_active = set()
                continue
            if jump_target_label_declaration(source, node.span):
                not_active = set()
            if len(not_active) > 0:
                out[id(node)] = _ActiveSheetState(frozenset(not_active), subject)
            toks = statement_tokens_after_leading_label(source, node.span)
            words = [tok.raw_text.lower() for tok in toks]
            is_set = _word(words, 0) == "set" and _word(words, 2) == "="
            if is_set:
                target = words[1]
                value = "".join(words[3:])
                if value == "activesheet":
                    sheets.add(target)
                    not_active.discard(target)
                    continue
                # `Worksheets.Add`, `Sheets.Add(...)`, qualified by a workbook or not.
                if _SHEETS_ADD_VALUE.search(value) is not None:
                    not_active = {sheet for sheet in sheets if sheet != target}
                    sheets.add(target)
                    continue
                if "(" not in value or _SHEETS_CALL_VALUE.search(value) is not None:
                    not_active.discard(target)
                    continue
            # A plain assignment with no call keeps what is known.
            bare = bare_assignment_target(source, node.span)
            calls = any(
                _raw_at(toks, i + 1) == "("
                and tok.kind is TokenKind.IDENTIFIER
                and _raw_at(toks, i - 1) == "."
                and token_text(tok) in _ACTIVATING_CALLS
                for i, tok in enumerate(toks)
            )
            if (
                (bare is not None or is_set)
                and not calls
                and not any(word == "activate" or word == "select" for word in words)
            ):
                continue
            # `w1.Activate`: w1 is active, and nothing else is known.
            not_active = set()
        else:
            stack.pop()
            if stack:
                not_active = set()
    return out


def _check_unqualified_corners(
    span: Span, toks: Sequence[VbaToken], state: _ActiveSheetState, push: PushFn
) -> None:
    """`w1.Range(Cells(1, 1), Cells(2, 2))` with w1 known not to be the active
    sheet: the unqualified Cells, Range, Rows or Columns are the active sheet's,
    so the Range raises 1004 (XLIDE issue #470, measured in Excel 16.0)."""
    i = 0
    while i + 3 < len(toks):
        dotted = toks[i].raw_text == "." and token_text(toks[i + 1]) == "range" and toks[i + 2].raw_text == "("
        if not dotted:
            i += 1
            continue
        before = _at(toks, i - 1)
        owner = _lower_name(before) if before is not None and before.raw_text != ")" else None
        sheet = owner
        if sheet is None:
            sheet = state.subject if i == 0 or (_raw_at(toks, i - 1) or "") in ("=", "(", ",") else None
        if not sheet or sheet not in state.not_active:
            i += 1
            continue
        close = match_paren_from(toks, i + 2)
        args = split_top_level_token_groups(toks, i + 3, ",", close) if close > i + 3 else []
        if len(args) != 2:
            i += 1
            continue
        for arg in args:
            part = _significant(arg)
            word = token_text(_at(part, 0))
            if (
                word in ("cells", "range", "rows", "columns")
                and _raw_at(part, 1) == "("
                and match_paren_from(part, 1) == len(part) - 1
            ):
                shown = "".join(tok.raw_text for tok in part)
                named = toks[i - 1].raw_text if owner else f".{toks[i + 1].raw_text}"
                push(
                    "hostArgumentOutOfRange",
                    f"{shown} here is the active sheet's, and '{named}' is on a sheet that is not active, so "
                    "Range cannot span them. This will raise Run-time error '1004': Method 'Range' of object "
                    "'_Worksheet' failed.",
                    Span(span.start + part[0].start, span.start + part[-1].end),
                )
                break
        i += 1


@dataclass(frozen=True, slots=True)
class _SheetFacts:
    """The sheets each statement knows: new and still empty, and those a new one differs from."""

    # Sheets from `Worksheets.Add` nothing has written to yet.
    empty: frozenset[str]
    # Pairs of sheet variables known to be different sheets, `a|b`.
    distinct: frozenset[str]


# Members that write to a sheet or its cells, read on the way to a value.
_SHEET_EDITS: frozenset[str] = frozenset(
    {
        "add", "insert", "paste", "pastespecial", "copy", "autofill", "fill", "filldown", "fillright",
        "formula", "formular1c1", "value", "value2", "text", "clear", "clearcontents", "delete", "sort",
        "autofilter", "texttocolumns", "removeduplicates",
    }
)


def _pair(a: str, b: str) -> str:
    """`[a, b].sort().join('|')`."""
    return "|".join(sorted((a, b)))


def _sheet_facts_at(
    source: str, proc: ProcedureNode, activity: ConditionalActivityTracker | None
) -> dict[int, _SheetFacts]:
    """What each statement knows of the sheets the code just added (XLIDE issue
    #472, measured in Excel 16.0): `Set w2 = Worksheets.Add` gives an empty
    sheet, different from every sheet the code held before. A statement that may
    write to it, a call, a label, or a block ends what is known. Keyed by the
    statement node's id()."""
    out: dict[int, _SheetFacts] = {}
    held: set[str] = set()
    empty: set[str] = set()
    distinct: set[str] = set()
    stack: list[Iterator[BodyNode]] = [iter(proc.body)]
    while stack:
        for node in stack[-1]:
            if _is_inactive(activity, node):
                continue
            if not is_leaf_statement(node):
                body = _block_body(node)
                if body is not None:
                    empty, distinct = set(empty), set(distinct)
                    stack.append(iter(body))
                    break
                empty, distinct = set(), set()
                continue
            if jump_target_label_declaration(source, node.span):
                empty, distinct = set(), set()
            if len(empty) > 0 or len(distinct) > 0:
                out[id(node)] = _SheetFacts(frozenset(empty), frozenset(distinct))
            toks = statement_tokens_after_leading_label(source, node.span)
            words = [tok.raw_text.lower() for tok in toks]
            if _word(words, 0) == "set" and _word(words, 2) == "=":
                target = words[1]
                value = "".join(words[3:])
                empty.discard(target)
                for pair in list(distinct):
                    if target in pair.split("|"):
                        distinct.discard(pair)
                if _SHEETS_ADD_VALUE.search(value) is not None:
                    for sheet in held:
                        if sheet != target:
                            distinct.add(_pair(sheet, target))
                    empty.add(target)
                if value == "activesheet" or _SHEETS_VALUE.search(value) is not None:
                    held.add(target)
                continue
            # A read through the sheet into a variable keeps it; anything else may write.
            bare = bare_assignment_target(source, node.span)
            edits = any(
                _raw_at(toks, i - 1) == "." and token_text(tok) in _SHEET_EDITS for i, tok in enumerate(toks)
            )
            reads_only = bare is not None and bare[0].lower() not in empty and not edits
            if not reads_only:
                for sheet in list(empty):
                    if sheet in words:
                        empty.discard(sheet)
        else:
            stack.pop()
            if stack:
                empty, distinct = set(), set()
    return out


# The SpecialCells types an empty sheet has no cells of.
_EMPTY_SPECIAL_CELLS: frozenset[str] = frozenset(
    {"xlcelltypeconstants", "xlcelltypeformulas", "xlcelltypeblanks", "xlcelltypecomments", "2", "-4123", "4", "-4144"}
)


def _check_sheet_facts(span: Span, toks: Sequence[VbaToken], facts: _SheetFacts | None, push: PushFn) -> None:
    """Errors the literals and a new sheet prove (XLIDE issue #472, measured in
    Excel 16.0): Intersect of literal ranges that do not meet is Nothing (91 at
    its member), Union across two sheets raises 1004, and on a sheet just added
    ShowAllData, SpecialCells of constants, formulas, blanks or comments,
    AutoFilter and TextToColumns raise 1004 and Find is Nothing (91)."""

    def at(first: int, last: int) -> Span:
        return Span(span.start + toks[first].start, span.start + toks[last].end)

    for i in range(len(toks) - 1):
        word = token_text(toks[i])
        owner = range_method_owner(toks, i)
        if (word == "intersect" or word == "union") and toks[i + 1].raw_text == "(" and owner:
            close = match_paren_from(toks, i + 1)
            args = split_top_level_token_groups(toks, i + 2, ",", close) if close > i + 2 else []
            areas = [_sheet_literal_range(arg) for arg in args]
            if len(args) != 2 or any(area is None for area in areas):
                continue
            a, b = areas[0], areas[1]
            assert a is not None and b is not None
            # Intersect raises it too (XLIDE issue #680, measured in Excel 16.0).
            if a[0] != b[0] and facts is not None and _pair(a[0], b[0]) in facts.distinct:
                method = "Union" if word == "union" else "Intersect"
                push(
                    "hostArgumentOutOfRange",
                    f"{method} takes ranges of one sheet, and '{a[0]}' and '{b[0]}' are different sheets. This "
                    f"will raise Run-time error '1004': Method '{method}' of object '{owner}' failed.",
                    at(i, close),
                )
            elif (
                word == "intersect"
                and a[0] == b[0]
                and _raw_at(toks, close + 1) == "."
                and not _overlaps(a[1], b[1])
            ):
                shown = "".join(tok.raw_text for tok in toks[i + 2 : close])
                member = _raw_at(toks, close + 2) or ""
                push(
                    "hostArgumentOutOfRange",
                    f"{shown} do not meet, so Intersect is Nothing and has no '.{member}'. This will raise "
                    "Run-time error '91': Object variable or With block variable not set.",
                    at(i, close),
                )
            continue
        sheet = _lower_name(toks[i])
        if (
            not sheet
            or facts is None
            or sheet not in facts.empty
            or _raw_at(toks, i - 1) == "."
            or toks[i + 1].raw_text != "."
        ):
            continue
        # The chain from the sheet: `w2.Cells.SpecialCells(...)`, `w2.ShowAllData`.
        k = i + 2
        while k < len(toks):
            member = token_text(toks[k])
            open_index = k + 1 if _raw_at(toks, k + 1) == "(" else -1
            close = match_paren_from(toks, open_index) if open_index > 0 else k
            # `w2.Range("A1:B5").AutoFilter Field:=1`: a call statement's arguments.
            bare_call = (
                open_index < 0
                and i == 0
                and k + 1 < len(toks)
                and toks[k + 1].raw_text != "."
                and toks[k + 1].raw_text != "="
            )
            if open_index > 0 and close > open_index + 1:
                args = split_top_level_token_groups(toks, open_index + 1, ",", close)
            elif bare_call:
                args = split_top_level_token_groups(toks, k + 1, ",", len(toks))
            else:
                args = []
            first = "".join(tok.raw_text.lower() for tok in _significant(args[0])) if len(args) > 0 else None
            problem: str | None = None
            shown = toks[i].raw_text
            if member == "showalldata":
                problem = (
                    f"'{shown}' is a sheet the code just added, with no filter to show. This will raise "
                    "Run-time error '1004': Method 'ShowAllData' of object '_Worksheet' failed"
                )
            elif member == "specialcells" and first is not None and first in _EMPTY_SPECIAL_CELLS:
                problem = (
                    f"'{shown}' is a sheet the code just added, which has no such cells. This will raise "
                    "Run-time error '1004': No cells were found."
                )
            elif (member == "autofilter" or member == "texttocolumns") and len(args) > 0:
                why = (
                    "This can't be applied to the selected range"
                    if member == "autofilter"
                    else "No data was selected to parse"
                )
                problem = (
                    f"'{shown}' is a sheet the code just added, and {toks[k].raw_text} has no data to act on. "
                    f"This will raise Run-time error '1004': {why}"
                )
            elif (
                member == "find"
                and _raw_at(toks, close + 1) == "."
                and len(args) > 0
                and len(args[0]) == 1
                and args[0][0].kind is TokenKind.STRING_LITERAL
                and args[0][0].raw_text != '""'
            ):
                after = _raw_at(toks, close + 2) or ""
                problem = (
                    f"'{shown}' is a sheet the code just added, so Find finds nothing and returns Nothing, "
                    f"which has no '.{after}'. This will raise Run-time error '91': Object variable or With "
                    "block variable not set"
                )
            if problem:
                push("hostArgumentOutOfRange", f"{problem}.", at(k, close))
                break
            if _raw_at(toks, close + 1) != ".":
                break
            k = close + 2


def range_method_owner(toks: Sequence[VbaToken], i: int) -> str | None:
    """The object an Intersect or Union at `i` is a method of, as Excel names it
    in its errors: `_Global` for the bare name, `_Application` after
    `Application.` or `Excel.Application.`. None after any other dot."""
    if _raw_at(toks, i - 1) != ".":
        return "_Global"
    if token_text(_at(toks, i - 2)) != "application":
        return None
    qualified = _raw_at(toks, i - 3) == "."
    if not qualified or (token_text(_at(toks, i - 4)) == "excel" and _raw_at(toks, i - 5) != "."):
        return "_Application"
    return None


def literal_intersect_is_nothing(value: Sequence[VbaToken]) -> bool:
    """Whether a Set's value is an Intersect of two literal ranges of one sheet
    that share no cell, which is Nothing: `Intersect(ws.Range("A1"),
    ws.Range("C3"))`, through Application too (XLIDE issue #680, measured in
    Excel 16.0)."""
    toks = [tok for tok in value if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE]
    i = next((k for k, tok in enumerate(toks) if token_text(tok) == "intersect"), -1)
    if (
        i < 0
        or _raw_at(toks, i + 1) != "("
        or range_method_owner(toks, i) is None
        or any(tok.raw_text != "." and token_text(tok) not in ("application", "excel") for tok in toks[:i])
        or match_paren_from(toks, i + 1) != len(toks) - 1
    ):
        return False
    args = split_top_level_token_groups(toks, i + 2, ",", len(toks) - 1)
    areas = [_sheet_literal_range(arg) for arg in args]
    if len(args) != 2 or any(area is None for area in areas):
        return False
    a, b = areas[0], areas[1]
    assert a is not None and b is not None
    return a[0] == b[0] and not _overlaps(a[1], b[1])


def _sheet_literal_range(arg: Sequence[VbaToken]) -> tuple[str, _A1Area] | None:
    """`w2.Range("A1:B2")` or `Range("A1")`: the sheet variable (or "" for none) and the area."""
    toks = _significant(arg)
    at = 2 if len(toks) == 6 and toks[1].raw_text == "." else 0 if len(toks) == 4 else -1
    if (
        at < 0
        or token_text(toks[at]) != "range"
        or _raw_at(toks, at + 1) != "("
        or _kind_at(toks, at + 2) is not TokenKind.STRING_LITERAL
        or _raw_at(toks, at + 3) != ")"
    ):
        return None
    area = _parse_a1_address(string_literal_value(toks[at + 2].raw_text))
    if area is None or not area.valid or area.row is None or area.column is None:
        return None
    return ((_lower_name(toks[0]) or "") if at == 2 else "", area)


def _overlaps(a: _A1Area, b: _A1Area) -> bool:
    """Whether two A1 areas share a cell."""

    def box(area: _A1Area) -> tuple[int, int, int, int]:
        assert area.row is not None and area.column is not None
        end_row = area.end_row if area.end_row is not None else area.row
        end_column = area.end_column if area.end_column is not None else area.column
        return (
            min(area.row, end_row),
            max(area.row, end_row),
            min(area.column, end_column),
            max(area.column, end_column),
        )

    ar1, ar2, ac1, ac2 = box(a)
    br1, br2, bc1, bc2 = box(b)
    return ar1 <= br2 and br1 <= ar2 and ac1 <= bc2 and bc1 <= ac2


# Range members a call statement edits cells with.
_CELL_EDITS: frozenset[str] = frozenset(
    {
        "clearcontents", "clear", "clearformats", "insert", "delete", "paste", "pastespecial", "autofill",
        "filldown", "fillright", "merge", "unmerge", "sort",
    }
)

# The members of a sheet that reach its cells.
_CELL_PATHS: frozenset[str] = frozenset({"range", "cells", "rows", "columns", "usedrange"})

# The cell properties a write to still raises under any AllowFormatting flag.
_VALUE_PROPERTIES: frozenset[str] = frozenset(
    {"value", "value2", "formula", "formular1c1", "formula2", "formula2r1c1", "formulaarray", "formulalocal", "formular1c1local"}
)

_ALLOW_FLAGS: tuple[str, ...] = (
    "allowinsertingrows", "allowinsertingcolumns", "allowformattingcells", "allowformattingcolumns", "allowformattingrows",
)


@dataclass(frozen=True, slots=True)
class _ProtectedSheet:
    """A sheet the code protected: its password ('' for none) and the Allow flags it set True."""

    password: str
    allows: frozenset[str]


def _allowed_edit(words: Sequence[str], eq: int, edit: int, allows: AbstractSet[str]) -> bool:
    """Which Allow flag lets a protected sheet's edit run (XLIDE issue #684,
    measured in Excel 16.0): AllowInsertingRows a row insert,
    AllowInsertingColumns a column insert, AllowFormattingColumns a ColumnWidth,
    AllowFormattingRows a RowHeight, and AllowFormattingCells any other format. A
    value write, ClearContents and a Delete raise whatever the flags:
    AllowDeletingRows deletes only unlocked rows."""
    if edit > 0 and words[edit] == "insert":
        rows = any(word == "rows" or word == "entirerow" for word in words[:edit])
        columns = any(word == "columns" or word == "entirecolumn" for word in words[:edit])
        return (rows and not columns and "allowinsertingrows" in allows) or (
            columns and not rows and "allowinsertingcolumns" in allows
        )
    if eq > 0:
        prop = words[eq - 1]
        if prop in _VALUE_PROPERTIES:
            return False
        if prop == "columnwidth":
            return "allowformattingcolumns" in allows
        if prop == "rowheight":
            return "allowformattingrows" in allows
        return "allowformattingcells" in allows
    return False


# What each change to a workbook's sheets raises while its structure is
# protected (XLIDE issue #684, measured in Excel 16.0).
_STRUCTURE_ERRORS: Mapping[str, str] = {
    "add": "Method 'Add' of object 'Sheets' failed",
    "name": "Method 'Name' of object '_Worksheet' failed",
    "delete": "Method 'Delete' of object '_Worksheet' failed",
    "visible": "Method 'Visible' of object '_Worksheet' failed",
    "copy": "Workbook is protected and cannot be changed",
}


@dataclass(frozen=True, slots=True)
class _ProtectArguments:
    args: list[list[VbaToken]]
    password_arg: list[VbaToken] | None
    # Its password, '' for none, None when not a literal.
    password: str | None

    def named(self, name: str) -> list[VbaToken] | None:
        for arg in self.args:
            if token_text(_at(arg, 0)) == name and _raw_at(arg, 1) == ":=":
                return arg[2:]
        return None


def _protect_arguments(toks: Sequence[VbaToken]) -> _ProtectArguments:
    """The Protect or Unprotect arguments at `toks[3]`."""
    if len(toks) > 3:
        paren = toks[3].raw_text == "("
        args = split_top_level_token_groups(
            toks, 4 if paren else 3, ",", match_paren_from(toks, 3) if paren else len(toks)
        )
    else:
        args = []
    partial = _ProtectArguments(args, None, None)
    password_arg = partial.named("password")
    if password_arg is None and len(args) > 0 and _raw_at(args[0], 1) != ":=":
        password_arg = args[0]
    if password_arg is None:
        password: str | None = ""
    elif len(password_arg) == 1 and password_arg[0].kind is TokenKind.STRING_LITERAL:
        password = string_literal_value(password_arg[0].raw_text)
    else:
        password = None
    return _ProtectArguments(args, password_arg, password)


def _check_protected_sheets(
    source: str, proc: ProcedureNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """The sheets the code just protected, with their password, and the faults
    that follow (XLIDE issue #471, measured in Excel 16.0): after `w2.Protect`, a
    write to its cells or a cell edit raises 1004, and `w2.Unprotect` with
    another password raises 1004. `Protect UserInterfaceOnly:=True` lets the code
    write, a right Unprotect ends it, and a cell's Locked set by the code, a
    call, a label or the end of a block ends what is known. A sheet protected
    with no password takes any at Unprotect, and the Allow flags let their edits
    run (XLIDE issue #684).

    The workbooks whose structure the code protected (XLIDE issue #684, measured
    in Excel 16.0): `wb.Protect "pw"`, Structure True unless given False. Adding
    a sheet to it raises 1004, and so does renaming, deleting, hiding or copying
    a sheet the code took from it, until Unprotect."""
    protected_sheets: dict[str, _ProtectedSheet] = {}
    protected_books: dict[str, str] = {}
    # The workbooks by name, and which one each sheet variable was taken from.
    books: set[str] = {"activeworkbook", "thisworkbook"}
    sheet_books: dict[str, str] = {}
    # Sheets the code unlocked a cell on: which cells stay writable is not followed.
    unlocked: set[str] = set()

    stack: list[Iterator[BodyNode]] = [iter(proc.body)]
    while stack:
        for node in stack[-1]:
            if _is_inactive(activity, node):
                continue
            if not is_leaf_statement(node):
                body = _block_body(node)
                if body is not None:
                    protected_sheets = dict(protected_sheets)
                    protected_books = dict(protected_books)
                    sheet_books = dict(sheet_books)
                    stack.append(iter(body))
                    break
                protected_sheets, protected_books, sheet_books = {}, {}, {}
                continue
            if jump_target_label_declaration(source, node.span):
                protected_sheets, protected_books, sheet_books = {}, {}, {}
            toks = statement_tokens_after_leading_label(source, node.span)
            words = [tok.raw_text.lower() for tok in toks]
            sheet = _word(words, 0)

            def at(first: int, last: int, node: BodyNode = node, toks: Sequence[VbaToken] = toks) -> Span:
                return Span(node.span.start + toks[first].start, node.span.start + toks[last].end)

            if "locked" in words and sheet is not None:
                unlocked.add(sheet)
            # Another workbook made active: what ActiveWorkbook was is not known.
            if "activate" in words or "workbooks" in words:
                protected_books.pop("activeworkbook", None)
                for name, held_book in list(sheet_books.items()):
                    if held_book == "activeworkbook":
                        sheet_books.pop(name, None)
            # Adding a sheet to a workbook whose structure is protected.
            for i in range(len(toks) - 2):
                if (words[i] == "worksheets" or words[i] == "sheets") and words[i + 1] == "." and words[i + 2] == "add":
                    added_to: str | None
                    if _word(words, i - 1) == ".":
                        before = _word(words, i - 2)
                        added_to = before if before is not None and before in books and _word(words, i - 3) != "." else None
                    else:
                        added_to = "activeworkbook"
                    if added_to and added_to in protected_books:
                        who = (
                            "The active workbook"
                            if added_to == "activeworkbook" and _word(words, i - 1) != "."
                            else f"'{toks[i - 2].raw_text}'"
                        )
                        push(
                            "hostArgumentOutOfRange",
                            f"{who} has its structure protected here, so no sheet can be added. This will raise "
                            f"Run-time error '1004': {_STRUCTURE_ERRORS['add']}.",
                            at(i - 2 if _word(words, i - 1) == "." else i, i + 2),
                        )
            if _word(words, 0) == "set" and _word(words, 2) == "=":
                target = words[1]
                value = words[3:]
                protected_sheets.pop(target, None)
                protected_books.pop(target, None)
                sheet_books.pop(target, None)
                if _word(value, 0) == "workbooks" or (
                    _word(value, 0) in ("activeworkbook", "thisworkbook") and len(value) == 1
                ):
                    books.add(target)
                # `Set ws = wb.Worksheets(1)`, `Worksheets.Add`, `ActiveSheet`.
                head = _word(value, 0)
                qualified = head is not None and head in books and _word(value, 1) == "."
                rest = value[2:] if qualified else value
                first = _word(rest, 0)
                takes = (
                    len(rest) == 1
                    if first == "activesheet"
                    else (first == "worksheets" or first == "sheets")
                    and (_word(rest, 1) == "(" or (_word(rest, 1) == "." and _word(rest, 2) == "add"))
                )
                if takes:
                    sheet_books[target] = value[0] if qualified else "activeworkbook"
                continue
            if (
                sheet is not None
                and sheet in books
                and _word(words, 1) == "."
                and (_word(words, 2) == "protect" or _word(words, 2) == "unprotect")
            ):
                parsed = _protect_arguments(toks)
                if words[2] == "protect":
                    structure_arg = parsed.named("structure")
                    if structure_arg is None and len(parsed.args) > 1 and _raw_at(parsed.args[1], 1) != ":=":
                        structure_arg = parsed.args[1]
                    structure = (
                        "true"
                        if structure_arg is None
                        else token_text(structure_arg[0])
                        if len(structure_arg) == 1
                        else None
                    )
                    if parsed.password is not None and structure == "true":
                        protected_books[sheet] = parsed.password
                    else:
                        protected_books.pop(sheet, None)
                else:
                    held_password = protected_books.get(sheet)
                    if (
                        held_password is not None
                        and held_password != ""
                        and parsed.password is not None
                        and parsed.password != held_password
                        and parsed.password_arg is not None
                    ):
                        push(
                            "hostArgumentOutOfRange",
                            f"'{toks[0].raw_text}' was protected with another password, which Unprotect must "
                            "match. This will raise Run-time error '1004': The password you supplied is not "
                            "correct.",
                            at(2, len(toks) - 1),
                        )
                        continue
                    protected_books.pop(sheet, None)
                continue
            # Renaming, deleting, hiding or copying a sheet of a protected workbook.
            book_of = sheet_books.get(sheet) if sheet is not None else None
            if book_of and book_of in protected_books and _word(words, 1) == ".":
                second = _word(words, 2)
                if (second == "name" or second == "visible") and _word(words, 3) == "=":
                    change: str | None = second
                elif second == "delete" and len(toks) == 3:
                    change = "delete"
                elif second == "copy":
                    change = "copy"
                else:
                    change = None
                if change:
                    push(
                        "hostArgumentOutOfRange",
                        f"'{toks[0].raw_text}' is a sheet of a workbook whose structure is protected here, so its "
                        f"sheets cannot be changed. This will raise Run-time error '1004': "
                        f"{_STRUCTURE_ERRORS[change]}.",
                        at(0, 2),
                    )
                    continue
            if _word(words, 1) == "." and (_word(words, 2) == "protect" or _word(words, 2) == "unprotect"):
                assert sheet is not None
                parsed = _protect_arguments(toks)
                if words[2] == "protect":
                    ui_only = parsed.named("userinterfaceonly")
                    if (
                        parsed.password is None
                        or sheet in unlocked
                        or (ui_only is not None and token_text(_at(ui_only, 0)) != "false")
                    ):
                        protected_sheets.pop(sheet, None)
                    else:
                        allows = frozenset(
                            flag
                            for flag in _ALLOW_FLAGS
                            if (value_toks := parsed.named(flag)) is not None and token_text(_at(value_toks, 0)) != "false"
                        )
                        protected_sheets[sheet] = _ProtectedSheet(parsed.password, allows)
                else:
                    protection_held = protected_sheets.get(sheet)
                    held_password = protection_held.password if protection_held is not None else None
                    # A sheet protected with no password takes any (XLIDE issue #684).
                    if (
                        held_password is not None
                        and held_password != ""
                        and parsed.password is not None
                        and parsed.password != held_password
                        and parsed.password_arg is not None
                    ):
                        push(
                            "hostArgumentOutOfRange",
                            f"'{toks[0].raw_text}' was protected with another password, which Unprotect must "
                            "match. This will raise Run-time error '1004': The password you supplied is not "
                            "correct.",
                            at(2, len(toks) - 1),
                        )
                        continue
                    protected_sheets.pop(sheet, None)
                continue
            protection = protected_sheets.get(sheet) if sheet is not None else None
            if protection is not None and _word(words, 1) == "." and _word(words, 2) in _CELL_PATHS:
                assert sheet is not None
                eq = next((i for i, tok in enumerate(toks) if tok.raw_text == "=" and i > 2), -1)
                edit = next(
                    (
                        i
                        for i, tok in enumerate(toks)
                        if i > 2 and toks[i - 1].raw_text == "." and token_text(tok) in _CELL_EDITS
                    ),
                    -1,
                )
                if "locked" in words:
                    protected_sheets.pop(sheet, None)
                    continue
                if (eq > 0 or edit > 0) and not _allowed_edit(words, eq, edit, protection.allows):
                    push(
                        "hostArgumentOutOfRange",
                        f"'{toks[0].raw_text}' is protected here, so its cells cannot be changed. This will raise "
                        "Run-time error '1004': The cell or chart you're trying to change is on a protected sheet.",
                        at(0, (eq if eq > 0 else edit + 1) - 1),
                    )
                continue
            # A read keeps what is known, and so does setting a property of
            # Application, `Application.DisplayAlerts = False`; a call or another
            # use of a sheet may unprotect it.
            bare: object = bare_assignment_target(source, node.span)
            if bare is None and _word(words, 0) == "application" and _word(words, 1) == "." and _word(words, 3) == "=":
                bare = _word(words, 2)
            if not bare or any(token_text(tok) == "unprotect" for tok in toks):
                protected_sheets = {}
                protected_books = {}
        else:
            stack.pop()
            if stack:
                protected_sheets, protected_books, sheet_books = {}, {}, {}


# Members that add to or edit a document, read on the way to a value.
_DOCUMENT_EDITS: frozenset[str] = frozenset(
    {
        "add", "addfield", "addpicture", "addtable", "insertafter", "insertbefore", "insertparagraph",
        "insertparagraphafter", "insertparagraphbefore", "insertbreak", "insertfile", "paste", "delete", "cut",
        "converttotable", "typetext",
    }
)


@dataclass(frozen=True, slots=True)
class _NewDocument:
    """A Word document the procedure just added, and the text it wrote into it, if any."""

    # The whole text the code set through Content.Text, "" for an untouched
    # document, None for text an expression gives that is not known.
    text: str | None


# The text constants a document's text is built with (XLIDE issue #694).
_TEXT_CONSTANTS: Mapping[str, str] = {
    "vbcr": "\r",
    "vblf": "\n",
    "vbcrlf": "\r\n",
    "vbnewline": "\r\n",
    "vbtab": "\t",
}


def _new_documents_at(
    source: str, proc: ProcedureNode, activity: ConditionalActivityTracker | None
) -> dict[int, dict[str, _NewDocument]]:
    """The documents each statement sees as new (XLIDE issue #497, measured in
    Word 16.0): `Set d = Documents.Add` gives an empty document, and
    `d.Content.Text = "..."` sets its whole text. A statement that names d other
    than to read through it, a label, or a block that names it ends what is
    known. Keyed by the statement node's id()."""
    out: dict[int, dict[str, _NewDocument]] = {}
    fold_context = StringFoldContext(
        name_value=lambda tok: _TEXT_CONSTANTS.get(token_text(tok)),
        integer_value=lambda toks: None,
    )
    # Each frame: the nodes left, the state they see, and the block whose body
    # they are (None for the procedure's own).
    stack: list[tuple[Iterator[BodyNode], dict[str, _NewDocument], BodyNode | None]] = [
        (iter(proc.body), {}, None)
    ]
    while stack:
        nodes, state, _owner = stack[-1]
        for node in nodes:
            if _is_inactive(activity, node):
                continue
            if not is_leaf_statement(node):
                body = _block_body(node)
                if body is not None:
                    stack.append((iter(body), dict(state), node))
                    break
                continue
            if jump_target_label_declaration(source, node.span):
                state.clear()
            if len(state) > 0:
                out[id(node)] = dict(state)
            toks = statement_tokens_after_leading_label(source, node.span)
            words = [tok.raw_text.lower() for tok in toks]
            # `Set d = Documents.Add` or `Set d = Documents.Add()`, plain.
            if _word(words, 0) == "set" and _word(words, 2) == "=":
                start = 5 if _word(words, 3) == "application" and _word(words, 4) == "." else 3
            else:
                start = -1
            if (
                start > 0
                and _word(words, start) == "documents"
                and _word(words, start + 1) == "."
                and _word(words, start + 2) == "add"
                and (
                    len(toks) == start + 3
                    or (len(toks) == start + 5 and words[start + 3] == "(" and words[start + 4] == ")")
                )
            ):
                state[words[1]] = _NewDocument("")
                continue
            # `d.Content.Text = "..."` writes the whole text, and so does an
            # expression that does not read the document, its text known where
            # every part is: `"One." & vbCr & "Two."` (XLIDE issue #694).
            target = _word(words, 0)
            if (
                target is not None
                and target in state
                and len(toks) > 6
                and words[1] == "."
                and words[2] == "content"
                and words[3] == "."
                and words[4] == "text"
                and words[5] == "="
                and not any(_lower_name(tok) == target for tok in toks[6:])
            ):
                state[target] = _NewDocument(fold_string_expression(toks[6:], fold_context))
                continue
            # Anything but a read through the document may change it, and so may
            # a method that adds or edits on the way: `x = d.Tables.Add(...)`.
            eq = next((i for i, tok in enumerate(toks) if tok.raw_text == "="), -1)
            edits = any(
                _raw_at(toks, i - 1) == "." and token_text(tok) in _DOCUMENT_EDITS for i, tok in enumerate(toks)
            )
            reads = not edits and bare_assignment_target(source, node.span) is not None and (
                target is None or target not in state
            )
            for lower in list(state.keys()):
                if lower in words and not (
                    reads
                    and all(
                        i <= eq or _lower_name(tok) != lower or _raw_at(toks, i + 1) == "."
                        for i, tok in enumerate(toks)
                    )
                ):
                    state.pop(lower, None)
        else:
            stack.pop()
            if stack and _owner is not None:
                parent_state = stack[-1][1]
                named = {_lower_name(tok) for tok in statement_tokens(source, _owner.span)}
                for lower in list(parent_state.keys()):
                    if lower in named:
                        parent_state.pop(lower, None)
    return out


def _new_document_count(document: _NewDocument, member: str) -> int | None:
    """What a new document holds of each collection, counted from its text."""
    # Text adds no table, field, bookmark, hyperlink, list, comment or section:
    # Chr(12) is a page break, and a URL stays text (XLIDE issue #694, measured
    # in Word 16.0).
    if member in ("tables", "fields", "inlineshapes", "bookmarks", "hyperlinks", "lists", "comments"):
        return 0
    if member == "sections":
        return 1
    text = document.text
    if text is None:
        return None
    # vbCr, vbLf and vbCrLf each end a paragraph; Chr(11) breaks a line.
    breaks = len(_LINE_BREAKS.findall(text))
    if member == "paragraphs":
        return breaks + 1
    if member == "sentences":
        # Each sentence ends at a stop or at the paragraph's end.
        return len(_SENTENCE_STOPS.findall(text)) + 1 if breaks == 0 else None
    if member == "words" or member == "characters":
        return None if "\n" in text else utf16_length(text) + 1  # the final paragraph mark
    return None


def _check_new_document_uses(
    span: Span,
    toks: Sequence[VbaToken],
    documents: Mapping[str, _NewDocument],
    value_of: ValueOf,
    push: PushFn,
) -> None:
    """Members of a new document past what it holds: `d.Tables(1)`, `d.Words(50)`, `d.Range(0, 99999)`."""
    for i in range(len(toks) - 4):
        name = _lower_name(toks[i])
        document = documents.get(name) if name else None
        if (
            document is None
            or _raw_at(toks, i - 1) == "."
            or toks[i + 1].raw_text != "."
            or toks[i + 3].raw_text != "("
        ):
            continue
        member = token_text(toks[i + 2])
        close = match_paren_from(toks, i + 3)
        args = split_top_level_token_groups(toks, i + 4, ",", close) if close > i + 4 else []
        at = Span(span.start + toks[i + 2].start, span.start + toks[close].end)
        if document.text is None:
            what = "whose text the code set"
        elif document.text:
            what = f"whose text the code set to {utf16_length(document.text)} character(s)"
        else:
            what = "which the code just added"
        if member == "range" and len(args) == 2:
            end = value_of(args[1])
            characters = (
                None if document.text is None or "\n" in document.text else utf16_length(document.text) + 1
            )
            if end is not None and characters is not None and end > characters:
                push(
                    "hostArgumentOutOfRange",
                    f"'{toks[i].raw_text}' is a new document {what}, so it ends at position {characters}, and "
                    f"Range ends at {_n(end)}. This will raise Run-time error '4608': Value out of range.",
                    at,
                )
            continue
        count = _new_document_count(document, member)
        if count is None or len(args) != 1:
            continue
        named = member == "bookmarks" and len(args[0]) == 1 and args[0][0].kind is TokenKind.STRING_LITERAL
        index = None if named else value_of(args[0])
        if named or (index is not None and index > count):
            shown = "".join(tok.raw_text for tok in toks[i + 2 : close + 1])
            push(
                "hostArgumentOutOfRange",
                f"'{toks[i].raw_text}' is a new document {what}, which has {count} {toks[i + 2].raw_text}, so "
                f"{shown} does not exist. This will raise Run-time error '5941': The requested member of the "
                "collection does not exist.",
                at,
            )


def _check_span(
    source: str,
    span: Span,
    host: str,
    model: HostObjectModel | None,
    member_ctx: MemberCompletionContext,
    env: Mapping[str, str],
    arrays: AbstractSet[str],
    source_names: AbstractSet[str],
    value_of: ValueOf,
    push: PushFn,
    string_of: StringOf | None = None,
) -> None:
    string_of = string_of if string_of is not None else _literal_string
    toks = statement_tokens(source, span)

    def at(first: int, last: int) -> Span:
        return Span(span.start + toks[first].start, span.start + toks[last].end)

    if host == "Excel":
        _check_sheet_name_assignment(source, span, toks, member_ctx, source_names, push)
        # `Range("A1:B2").Areas(2)`: an address with no comma is one area (XLIDE
        # issue #278, measured in Excel 16.0: 1004).
        for i in range(len(toks) - 8):
            if (
                token_text(toks[i]) != "range"
                or ("range" in source_names and _raw_at(toks, i - 1) != ".")
                or toks[i + 1].raw_text != "("
                or toks[i + 2].kind is not TokenKind.STRING_LITERAL
                or toks[i + 3].raw_text != ")"
                or toks[i + 4].raw_text != "."
                or token_text(toks[i + 5]) != "areas"
                or toks[i + 6].raw_text != "("
                or toks[i + 7].kind is not TokenKind.INTEGER_LITERAL
                or toks[i + 8].raw_text != ")"
            ):
                continue
            address = toks[i + 2].raw_text[1:-1]
            areas_index = js_number(_INTEGER_SUFFIX.sub("", toks[i + 7].raw_text))
            if _AREAS_ADDRESS.search(address) is not None and areas_index > 1:
                push(
                    "hostArgumentOutOfRange",
                    f"Range(\"{address}\") is one area, so Areas({_n(areas_index)}) names none. This will raise "
                    "Run-time error '1004': Application-defined or object-defined error.",
                    at(i + 7, i + 7),
                )
        _check_before_and_after(source, span, toks, member_ctx, push)
    chains: dict[int, _CellBlock | None] = {}
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
        elif callee.returns and _is_collection_type(callee.returns, model) and callee.open_index > 0:
            collection = callee.returns
        relative = host == "Excel" and callee.receiver == "Excel.Range" and lower in _RANGE_RELATIVE_MEMBERS
        if collection and len(callee.args) == 1 and lower not in _RANGE_COORDINATE_MEMBERS and not relative:
            index = value_of(callee.args[0])
            if index is not None and index < 1:
                error = _collection_index_error(host, collection, model)
                push(
                    "hostArgumentOutOfRange",
                    f"Index {_n(index)} is never an element: {host} collections start at 1. "
                    f"This will raise Run-time error '{error.number}': {error.text}.",
                    at(callee.open_index + 1, callee.close_index - 1),
                )
                continue
            # `Worksheets(Worksheets.Count + 1)`: past the last element (XLIDE
            # issue #309, measured in Excel and Word 16.0).
            chain = _chain_text(
                toks,
                _receiver_start(toks, callee.name_index),
                callee.name_index - 2 if lower == "item" else callee.name_index,
            )
            past = _count_offset(callee.args[0], chain) if chain else None
            if past is not None and past >= 1:
                error = _collection_index_error(host, collection, model)
                push(
                    "hostArgumentOutOfRange",
                    f"{chain}.Count + {_n(past)} is past the last element of {chain}. This will raise "
                    f"Run-time error '{error.number}': {error.text}.",
                    at(callee.open_index + 1, callee.close_index - 1),
                )
                continue
        _check_argument_limits(span, callee, value_of, push)
        _check_insertion_position(span, toks, callee, value_of, push)
        if host == "Word":
            _check_word_names(span, callee, string_of, push)
        if host == "Excel":
            _check_excel_method_arguments(span, toks, callee, string_of, push)
            _check_excel_callee(
                source, span, toks, callee, callee_span, env, arrays, source_names, value_of, push, string_of, chains
            )
        elif host == "Word":
            if lower == "range" and callee.receiver == "Word.Document" and callee.open_index > 0:
                start = value_of(callee.args[0]) if len(callee.args) > 0 else None
                end = value_of(callee.args[1]) if len(callee.args) > 1 else None
                bad = (
                    (start is not None and start < 0)
                    or (end is not None and end < 0)
                    or (start is not None and end is not None and end < start)
                )
                if bad:
                    push(
                        "hostArgumentOutOfRange",
                        "Document.Range takes character positions from 0, with End at or after Start. This "
                        "will raise Run-time error '4608': Value out of range.",
                        at(callee.open_index + 1, callee.close_index - 1),
                    )


@dataclass(frozen=True, slots=True)
class _Limit:
    """An argument a host method refuses, found by name or by position."""

    receivers: tuple[str, ...]
    member: str
    parameter: str
    position: int
    runs: str
    # Inclusive ranges the host refuses; an open end is unbounded.
    refused: tuple[tuple[int | None, int | None], ...]
    error: _RaisedError


# Counts a host method refuses (XLIDE issue #204, measured in Excel and Word
# 16.0): `Worksheets.Add Count:=0` raises 1004, and Word's Tables.Add takes 1 to
# 32767 rows and 1 to 63 columns (5148). Each argument is found by name or by
# position.
_APP_1004 = _RaisedError("1004", "Application-defined or object-defined error")
_PPT_VALUE_OUT_OF_RANGE = _RaisedError("-2147024809", "The specified value is out of range")
_PPT_INTEGER_OUT_OF_RANGE = _RaisedError("-2147188160", "Integer out of range")
_ARGUMENT_LIMITS: tuple[_Limit, ...] = (
    _Limit(
        ("Excel.Sheets", "Excel.Worksheets"), "add", "Count", 2, "1 or more",
        ((None, 0),), _RaisedError("1004", "Method 'Add' of object 'Sheets' failed"),
    ),
    # Enum arguments (XLIDE issue #244), swept from -100 to 999 in Excel 16.0 as
    # hostPropertyValues.ts's enum properties were.
    _Limit(
        ("Excel.Range",), "borders", "Index", 0, "1 to 12, or an xlBordersIndex constant",
        ((-100, 0), (13, 999)), _RaisedError("1004", "Unable to get the Item property of the Borders class"),
    ),
    _Limit(
        ("Excel.Range",), "end", "Direction", 0, "1 to 4, or an xlDirection constant",
        ((-100, 0), (5, 999)), _APP_1004,
    ),
    _Limit(
        ("Excel.Range",), "specialcells", "Type", 0, "an xlCellType constant",
        ((-100, 0), (13, 13), (17, 999)), _APP_1004,
    ),
    _Limit(
        ("Excel.Range",), "sort", "Order1", 1, "xlAscending (1) or xlDescending (2)",
        ((-100, 0), (3, 999)), _APP_1004,
    ),
    _Limit(
        ("Excel.Range",), "pastespecial", "Paste", 0, "an xlPasteType constant",
        ((-100, 0), (9, 10), (15, 999)), _APP_1004,
    ),
    _Limit(
        ("Excel.Range",), "insert", "Shift", 0, "1 to 4, or an xlInsertShiftDirection constant",
        ((-100, 0), (5, 999)), _APP_1004,
    ),
    _Limit(
        ("Excel.Range",), "delete", "Shift", 0, "1 to 4, or an xlDeleteShiftDirection constant",
        ((-100, 0), (5, 999)), _APP_1004,
    ),
    _Limit(
        ("Excel.Range",), "autofill", "Type", 1, "0 to 12, or an xlAutoFillType constant",
        ((-100, -1), (13, 999)), _APP_1004,
    ),
    # Word and PowerPoint (XLIDE issue #245), swept from -100 to 999 in 16.0.
    _Limit(
        ("Word.Selection",), "moveright", "Unit", 0, "1 to 3 or 16: wdCharacter, wdWord, wdSentence or wdCell",
        ((-100, 0), (4, 11), (13, 15), (17, 999)), _RaisedError("4120", "Bad parameter"),
    ),
    _Limit(
        ("Word.Selection",), "collapse", "Direction", 0, "wdCollapseEnd (0) or wdCollapseStart (1)",
        ((-100, -1), (2, 999)), _RaisedError("4120", "Bad parameter"),
    ),
    _Limit(
        ("Word.Selection",), "insertbreak", "Type", 0, "0 to 11, a wdBreakType constant",
        ((-100, -1), (12, 999)), _RaisedError("9118", "Parameter value was out of acceptable range"),
    ),
    _Limit(
        ("PowerPoint.Shapes",), "addshape", "Type", 0, "1 to 183, an msoAutoShapeType constant",
        ((-100, 0), (184, 999)), _PPT_VALUE_OUT_OF_RANGE,
    ),
    _Limit(
        ("PowerPoint.Slides",), "add", "Layout", 1, "a ppSlideLayout constant from 1",
        ((-100, 0), (37, 999)), _RaisedError("-2147024809", "Invalid enumeration value"),
    ),
    _Limit(
        ("PowerPoint.Shapes",), "addtable", "NumRows", 0, "1 to 75",
        ((-100, 0), (76, 999)), _PPT_INTEGER_OUT_OF_RANGE,
    ),
    _Limit(
        ("PowerPoint.Shapes",), "addtable", "NumColumns", 1, "1 to 75",
        ((-100, 0), (76, 999)), _PPT_INTEGER_OUT_OF_RANGE,
    ),
    _Limit(
        ("PowerPoint.Slides",), "range", "Index", 0, "1 or more",
        ((-100, 0),), _RaisedError("-2147188160", "Invalid request"),
    ),
    # XLIDE issue #311, measured in PowerPoint 16.0.
    _Limit(
        ("PowerPoint.Shapes",), "addtextbox", "Width", 3, "0 or more",
        ((None, -1),), _PPT_VALUE_OUT_OF_RANGE,
    ),
    _Limit(
        ("PowerPoint.Shapes",), "addtextbox", "Height", 4, "0 or more",
        ((None, -1),), _PPT_VALUE_OUT_OF_RANGE,
    ),
    _Limit(
        ("PowerPoint.Slide", "PowerPoint.SlideRange"), "moveto", "ToPos", 0, "1 or more",
        ((None, 0),), _PPT_INTEGER_OUT_OF_RANGE,
    ),
    # XLIDE issue #610, measured in PowerPoint 16.0: orientations 1 and 6 run, 0,
    # -2 (mixed), 7 and 9 do not; a width or height of 0 runs.
    _Limit(
        ("PowerPoint.Shapes",), "addtextbox", "Orientation", 0, "1 to 6, an msoTextOrientation constant",
        ((None, 0), (7, None)), _PPT_VALUE_OUT_OF_RANGE,
    ),
    _Limit(
        ("PowerPoint.Shapes",), "addshape", "Width", 3, "0 or more",
        ((None, -1),), _PPT_VALUE_OUT_OF_RANGE,
    ),
    _Limit(
        ("PowerPoint.Shapes",), "addshape", "Height", 4, "0 or more",
        ((None, -1),), _PPT_VALUE_OUT_OF_RANGE,
    ),
    _Limit(
        ("Word.Tables",), "add", "NumRows", 1, "1 to 32767",
        ((None, 0), (32768, None)), _RaisedError("5148", "The number must be between 1 and 32767"),
    ),
    _Limit(
        ("Word.Tables",), "add", "NumColumns", 2, "1 to 63",
        ((None, 0), (64, None)), _RaisedError("5148", "The number must be between 1 and 63"),
    ),
)


def _check_argument_limits(span: Span, callee: _HostCallee, value_of: ValueOf, push: PushFn) -> None:
    lower = callee.name.lower()
    for limit in _ARGUMENT_LIMITS:
        if limit.member != lower or callee.receiver not in limit.receivers:
            continue
        arg = _argument_by_name_or_position(callee.args, limit.parameter, limit.position)
        value = value_of(arg) if arg is not None and len(arg) > 0 else None
        if value is None or not any(
            (low is None or value >= low) and (high is None or value <= high) for low, high in limit.refused
        ):
            continue
        assert arg is not None
        receiver = callee.receiver[callee.receiver.find(".") + 1 :]
        push(
            "hostArgumentOutOfRange",
            f"{receiver}.{callee.name} takes {limit.parameter} {limit.runs}; {_n(value)} is outside that. This "
            f"will raise Run-time error '{limit.error.number}': {limit.error.text}.",
            _arg_span(span, arg),
        )


def _has_non_word_character(name: str) -> bool:
    """`/[^\\p{L}\\p{N}_]/u`: a character other than a letter, a number or `_`."""
    return any(ch != "_" and unicodedata.category(ch)[0] not in ("L", "N") for ch in name)


def _check_word_names(span: Span, callee: _HostCallee, string_of: StringOf, push: PushFn) -> None:
    """Names Word refuses (XLIDE issue #311, measured in Word 16.0): a bookmark
    name that is empty, starts with a digit, or holds anything but letters,
    digits and underscores (5828), and an empty style name (5167). A bookmark
    name past 40 letters is shortened, not refused. XLIDE issue #610 adds a style
    name of blanks (5167) and a built-in style's name in any case (5173), and an
    empty Variables name (-2147467259)."""
    if callee.name.lower() != "add":
        return
    arg = _argument_by_name_or_position(callee.args, "Name", 0)
    name = string_of(arg) if arg else None
    if name is None or arg is None:
        return
    if callee.receiver == "Word.Bookmarks":
        if name == "":
            why: str | None = "is empty"
        elif _LEADING_DIGIT.search(name) is not None:
            why = "starts with a digit"
        elif _has_non_word_character(name):
            why = "holds a character other than a letter, a digit or _"
        else:
            why = None
        if why:
            push(
                "hostArgumentOutOfRange",
                f"The bookmark name \"{name}\" {why}: a bookmark name starts with a letter and holds letters, "
                "digits and _. This will raise Run-time error '5828': Bad bookmark name.",
                _arg_span(span, arg),
            )
    elif callee.receiver == "Word.Styles" and js_trim(name) == "":
        push(
            "hostArgumentOutOfRange",
            f"A style needs a name, and \"{name}\" is none. This will raise Run-time error '5167': This is not a "
            "valid style name.",
            _arg_span(span, arg),
        )
    elif callee.receiver == "Word.Styles" and name.lower() in WORD_BUILTIN_STYLES:
        push(
            "hostArgumentOutOfRange",
            f"\"{name}\" is the name of a built-in style, which a style the code adds cannot take, in any case. "
            "This will raise Run-time error '5173': This style name already exists or is reserved for a "
            "built-in style.",
            _arg_span(span, arg),
        )
    elif callee.receiver == "Word.Variables" and name == "":
        push(
            "hostArgumentOutOfRange",
            "A document variable needs a name, and \"\" is none. This will raise Run-time error '-2147467259': "
            "Method 'Add' of object 'Variables' failed.",
            _arg_span(span, arg),
        )


def _check_excel_method_arguments(
    span: Span, toks: Sequence[VbaToken], callee: _HostCallee, string_of: StringOf, push: PushFn
) -> None:
    """Excel methods whose 1004 the arguments prove (XLIDE issue #308, measured
    in Excel 16.0): AutoFill into a range that does not hold its source, a Sort
    of several cells keyed on a column outside them, and Names.Add of a name
    Excel cannot hold."""
    lower = callee.name.lower()

    def given(name: str, position: int) -> list[VbaToken] | None:
        arg = _argument_by_name_or_position(callee.args, name, position)
        return arg if arg is not None and len(arg) > 0 else None

    if callee.receiver == "Excel.Range" and lower == "autofill":
        source_block = _literal_range_receiver(toks, callee.name_index - 1)
        destination = given("Destination", 0)
        target = _literal_range_argument(destination) if destination is not None else None
        if source_block is None or target is None or destination is None:
            return
        holds = (
            target.row <= source_block.row
            and target.column <= source_block.column
            and target.row + target.rows >= source_block.row + source_block.rows
            and target.column + target.width >= source_block.column + source_block.width
        )
        same = holds and target.rows == source_block.rows and target.width == source_block.width
        if not holds or same:
            push(
                "hostArgumentOutOfRange",
                f"AutoFill fills from {source_block.text} into {target.text}, which "
                f"{'is the source itself' if same else 'does not take it in'}: the destination must hold the "
                "source and reach past it. This will raise Run-time error '1004': AutoFill method of Range class "
                "failed.",
                _arg_span(span, destination),
            )
        return
    if callee.receiver == "Excel.Range" and lower == "sort" and given("Orientation", 10) is None:
        block = _literal_range_receiver(toks, callee.name_index - 1)
        if block is None or (block.rows == 1 and block.width == 1):
            return  # one cell sorts its current region, which the data decides
        for name, position in (("Key1", 0), ("Key2", 2), ("Key3", 5)):
            key = given(name, position)
            key_at = _literal_range_argument(key) if key is not None else None
            if (
                key is not None
                and key_at is not None
                and (key_at.column + key_at.width <= block.column or key_at.column >= block.column + block.width)
            ):
                push(
                    "hostArgumentOutOfRange",
                    f"{name} {key_at.text} lies outside the columns of {block.text}, which is what is sorted. This "
                    "will raise Run-time error '1004': The sort reference is not valid.",
                    _arg_span(span, key),
                )
                return
        return
    if callee.receiver == "Excel.Names" and lower == "add":
        arg = given("Name", 0)
        name_text = string_of(arg) if arg is not None else None
        why = None if name_text is None else _refused_name(name_text)
        if why and arg is not None:
            push(
                "hostArgumentOutOfRange",
                f"\"{name_text}\" {why}, so Excel holds no name of that spelling. This will raise Run-time error "
                "'1004': The syntax of this name isn't correct.",
                _arg_span(span, arg),
            )


def _refused_name(name: str) -> str | None:
    """Why Excel refuses a name, where it is plain (XLIDE issue #308, measured in
    Excel 16.0): a space, a digit first, or the spelling of a cell, A1 or R1C1.
    A0, XFE1 and A1048577 are names it accepts."""
    if _JS_SPACE.search(name) is not None:
        return "holds a space"
    if _LEADING_DIGIT.search(name) is not None:
        return "starts with a digit"
    cell = _A1_CELL.search(name)
    if cell is not None:
        row = _decimal_number(cell.group(2))
        if _column_number(cell.group(1)) <= _EXCEL_MAX_COLUMN and 1 <= row <= _EXCEL_MAX_ROW:
            return "is the address of a cell"
    if _R1C1_NAME.search(name) is not None:
        return "is an R1C1 address"
    return None


def _literal_range_argument(arg: Sequence[VbaToken]) -> _CellBlock | None:
    """`Range("B1:B3")` as a whole argument: its top-left cell and its size."""
    toks = _significant(arg)
    return _literal_range_at(toks, 0) if len(toks) == 4 else None


_BEFORE_AND_AFTER_TYPES: frozenset[str] = frozenset(
    {"Excel.Sheets", "Excel.Worksheets", "Excel.Charts", "Excel.Worksheet", "Excel.Chart"}
)


def _check_before_and_after(
    source: str, span: Span, toks: Sequence[VbaToken], member_ctx: MemberCompletionContext, push: PushFn
) -> None:
    """`Worksheets.Add Before:=..., After:=...` and Move or Copy of a sheet given
    both: the sheet goes before one or after one, never both (XLIDE issue #308,
    measured in Excel 16.0: 1004)."""
    for i in range(1, len(toks)):
        lower = token_text(toks[i])
        if toks[i - 1].raw_text != "." or (lower != "add" and lower != "move" and lower != "copy"):
            continue
        open_index = i + 1 if _raw_at(toks, i + 1) == "(" else -1
        close = match_paren_from(toks, open_index) if open_index > 0 else len(toks)
        if open_index < 0 and i != _first_executable_token_index_of_member_call(toks, i):
            continue
        first = open_index + 1 if open_index > 0 else i + 1
        args = _split_top_level(toks[first:close]) if close > first else []
        before = _argument_by_name_or_position(args, "Before", 0)
        after = _argument_by_name_or_position(args, "After", 1)
        if not before or not after:
            continue
        resolved = resolve_receiver_type_at(source, span.start + toks[i - 1].end, member_ctx) or ""
        parts = (resolved[len("union:") :] if resolved.startswith("union:") else resolved).split("|")
        if all(part in _BEFORE_AND_AFTER_TYPES for part in parts):
            last = close - 1 if close == len(toks) else close
            push(
                "hostArgumentOutOfRange",
                f"{toks[i].raw_text} takes Before or After, not both: a sheet goes before one sheet or after one. "
                f"This will raise Run-time error '1004': Method '{toks[i].raw_text}' failed.",
                Span(span.start + toks[i].start, span.start + toks[last].end),
            )


@dataclass(frozen=True, slots=True)
class _Insertion:
    receiver: str
    member: str
    parameter: str
    noun: str
    error: _RaisedError


# Methods that insert at a position from 1 to Count + 1 (XLIDE issue #309,
# measured in Excel and PowerPoint 16.0): 0 and below, and Count + 2 or more
# written against the collection's own Count, are refused.
_INSERTIONS: tuple[_Insertion, ...] = (
    _Insertion("PowerPoint.Slides", "add", "Index", "slide", _PPT_INTEGER_OUT_OF_RANGE),
    _Insertion("PowerPoint.Slides", "addslide", "Index", "slide", _PPT_INTEGER_OUT_OF_RANGE),
    _Insertion("Excel.ListRows", "add", "Position", "row", _RaisedError("9", "Subscript out of range")),
    _Insertion("Excel.ListColumns", "add", "Position", "column", _RaisedError("9", "Subscript out of range")),
)


def _check_insertion_position(
    span: Span, toks: Sequence[VbaToken], callee: _HostCallee, value_of: ValueOf, push: PushFn
) -> None:
    lower = callee.name.lower()
    insertion = next(
        (entry for entry in _INSERTIONS if entry.member == lower and entry.receiver == callee.receiver), None
    )
    arg = _argument_by_name_or_position(callee.args, insertion.parameter, 0) if insertion is not None else None
    if insertion is None or not arg:
        return
    value = value_of(arg)
    chain = _chain_text(toks, _receiver_start(toks, callee.name_index), callee.name_index - 2)
    past = _count_offset(arg, chain) if chain else None
    if value is not None and value < 1:
        shown: str | None = _n(value)
    elif past is not None and past >= 2:
        shown = f"{chain}.Count + {_n(past)}"
    else:
        shown = None
    if shown:
        type_name = callee.receiver[callee.receiver.find(".") + 1 :]
        push(
            "hostArgumentOutOfRange",
            f"{type_name}.{callee.name} puts the new {insertion.noun} at {insertion.parameter}, from 1 to Count + 1; "
            f"{shown} is outside that. This will raise Run-time error '{insertion.error.number}': "
            f"{insertion.error.text}.",
            _arg_span(span, arg),
        )


def _chain_text(toks: Sequence[VbaToken], start: int, end: int) -> str | None:
    """The text of `toks[start..end]` when it is names and dots only: `d.Paragraphs`."""
    part = _js_slice(toks, start, end + 1)
    if (
        len(part) > 0
        and all(
            (token_name(tok) is not None) if k % 2 == 0 else tok.raw_text == "." for k, tok in enumerate(part)
        )
        and len(part) % 2 == 1
    ):
        return "".join(tok.raw_text for tok in part)
    return None


def _js_slice(toks: Sequence[VbaToken], start: int, end: int) -> Sequence[VbaToken]:
    """`Array.prototype.slice(start, end)`: a negative bound counts from the end."""
    length = len(toks)
    first = max(length + start, 0) if start < 0 else min(start, length)
    last = max(length + end, 0) if end < 0 else min(end, length)
    return toks[first:last] if first < last else []


def _count_offset(arg: Sequence[VbaToken], chain: str) -> Number | None:
    """`k` for an argument written `<chain>.Count + k` (0 for `<chain>.Count`), or None."""
    toks = _significant(arg)
    count = next(
        (k for k, tok in enumerate(toks) if token_text(tok) == "count" and _raw_at(toks, k - 1) == "."), -1
    )
    if count < 2:
        return None
    text = _chain_text(toks, 0, count - 2)
    if text is None or text.lower() != chain.lower():
        return None
    rest = toks[count + 1 :]
    if len(rest) == 0:
        return 0
    if len(rest) != 2 or rest[1].kind is not TokenKind.INTEGER_LITERAL:
        return None
    k = js_number(rest[1].raw_text)
    if _is_integral(k):
        k = int(k)
    if rest[0].raw_text == "+":
        return k
    if rest[0].raw_text == "-":
        return -k
    return None


def _argument_by_name_or_position(
    args: Sequence[Sequence[VbaToken]], name: str, position: int
) -> list[VbaToken] | None:
    """An argument's value tokens, named (`Count:=0`) or at its position before any named one."""
    for index, arg in enumerate(args):
        toks = _significant(arg)
        if _raw_at(toks, 1) == ":=":
            if _lower_name(toks[0]) == name.lower():
                return toks[2:]
            continue
        if index == position:
            return toks
    return None


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
    qualified = _raw_at(toks, i - 1) == "."
    parenthesized = _raw_at(toks, i + 1) == "("
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
            return _HostCallee(name, global_member.get("returns"), "global", i, open_index, close_index, args)
        return None
    typed = _host_receiver_types(resolve_receiver_type_at(source, span.start + toks[i - 1].end, member_ctx), model)
    resolved = (
        typed
        if len(typed) > 0
        else _host_receiver_types(_range_item_type(source, span, toks, i - 1, member_ctx), model)
    )
    # Of a union, only a part that has the member can run the call: the member
    # is judged on the one part that has it. `ActiveSheet` is a Worksheet or a
    # Chart, and only a Worksheet has Cells (XLIDE issue #182).
    having = [part for part in resolved if resolve_host_member(part, name, model) is not None]
    # Parts that share the member and its type are judged as one:
    # `ActiveSheet.Shapes(0)` on a Worksheet or a Chart (XLIDE issue #276).
    returns: set[str | None] = set()
    for part in having:
        part_member = resolve_host_member(part, name, model)
        assert part_member is not None
        returns.add(part_member.get("returns"))
    if len(having) == 0 or (len(having) > 1 and (len(returns) != 1 or None in returns)):
        return None
    receiver = having[0]
    member = resolve_host_member(receiver, name, model)
    assert member is not None
    return _HostCallee(name, member.get("returns"), receiver, i, open_index, close_index, args)


def _range_item_type(
    source: str, span: Span, toks: Sequence[VbaToken], dot_index: int, member_ctx: MemberCompletionContext
) -> str | None:
    """`r.Item(1)` before the dot at `dot_index`, with r a Range: a Range too,
    though the model types Item as a Variant (XLIDE issue #559, measured in Excel
    16.0)."""
    if _raw_at(toks, dot_index) != "." or _raw_at(toks, dot_index - 1) != ")":
        return None
    close = dot_index - 1
    open_index = _open_paren_for(toks, close)
    if open_index < 3 or token_text(toks[open_index - 1]) != "item" or toks[open_index - 2].raw_text != ".":
        return None
    resolved = resolve_receiver_type_at(source, span.start + toks[open_index - 2].end, member_ctx)
    return "Excel.Range" if resolved == "Excel.Range" else None


def _host_receiver_types(resolved: str | None, model: HostObjectModel | None) -> list[str]:
    """The host types a resolved receiver may be. A one-part union - what a
    collection's Object-declared Item gives, `Worksheets(1)` (XLIDE issue #114) -
    is its part. A union with any part the model does not know is not judged."""
    if not resolved:
        return []
    parts = resolved[len("union:") :].split("|") if resolved.startswith("union:") else [resolved]
    return parts if all(get_host_type(part, model) is not None for part in parts) else []


def _first_executable_token_index_of_member_call(toks: Sequence[VbaToken], name_index: int) -> int:
    """The member call's name index when the statement is `a.b.Name args`, else -1."""
    # Back over each `.` to the name before it, stepping over a call's
    # parentheses: `Range("A1").Insert Shift:=9` (XLIDE issue #244).
    j = name_index
    while j >= 2 and toks[j - 1].raw_text == ".":
        k = j - 2
        if toks[k].raw_text == ")":
            open_index = _open_paren_for(toks, k)
            if open_index < 1:
                return -1
            k = open_index - 1
        j = k
    # Inside a With, `.Add "a b", ...` starts at its leading dot (XLIDE issue #610).
    first = first_executable_token_index(toks)
    return name_index if j == first or (_raw_at(toks, j - 1) == "." and j - 1 == first) else -1


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


def _collection_index_error(host: str, collection: str, model: HostObjectModel | None) -> _RaisedError:
    bare = collection[collection.find(".") + 1 :]
    if host == "PowerPoint":
        return _RaisedError("-2147188160", "Integer out of range")
    if bare == "Shapes":
        return _RaisedError("-2147024809", "The index into the specified collection is out of bounds")
    if host == "Word":
        return _RaisedError("5941", "The requested member of the collection does not exist")
    item = resolve_host_member(collection, "Item", model)
    if bare == "PivotTables" or bare == "Range" or (item is not None and item.get("returns") == "Excel.Range"):
        return _RaisedError("1004", "Application-defined or object-defined error")
    return _RaisedError("9", "Subscript out of range")


def _check_excel_callee(
    source: str,
    span: Span,
    toks: Sequence[VbaToken],
    callee: _HostCallee,
    callee_span: Span,
    env: Mapping[str, str],
    arrays: AbstractSet[str],
    source_names: AbstractSet[str],
    value_of: ValueOf,
    push: PushFn,
    string_of: StringOf | None = None,
    chains: dict[int, _CellBlock | None] | None = None,
) -> None:
    string_of = string_of if string_of is not None else _literal_string
    lower = callee.name.lower()
    if callee.open_index < 0:
        return
    args_span = (
        Span(span.start + toks[callee.open_index + 1].start, span.start + toks[callee.close_index - 1].end)
        if callee.close_index > callee.open_index + 1
        else callee_span
    )
    if callee.receiver == "Excel.WorksheetFunction":
        why = worksheet_function_refusal(lower, callee.args)
        if why:
            push(
                "hostArgumentOutOfRange",
                f"WorksheetFunction.{callee.name}: {why}. The worksheet error is raised as Run-time error '1004': "
                f"Unable to get the {callee.name} property of the WorksheetFunction class.",
                args_span,
            )
        return
    # A range counts its Cells, Rows and Columns from its own top-left cell, so
    # a literal `Range("B2")` receiver moves the far edge in. Any other receiver
    # starts at A1 or below, so the count alone past the edge is already off the
    # sheet (XLIDE issue #182).
    origin = _single_cell_receiver(toks, callee.name_index - 1)
    # A chain from a literal range, or a block of several cells: each step moves
    # or resizes the block, and one that leaves the sheet raises 1004 (XLIDE
    # issue #508, measured in Excel 16.0). One step from one literal cell is the
    # checks below.
    one_step = origin is not None and lower != "item"
    if callee.receiver == "Excel.Range" and lower in _CHAIN_MEMBERS and not one_step:
        block = _range_chain_receiver(
            toks, callee.name_index - 1, value_of, source_names, chains if chains is not None else {}
        )
        following = _chain_step(block, lower, callee.args, value_of) if block is not None else None
        if block is not None and following is not None and _off_sheet(following):
            rows = (
                f"row {_n(following.row)}"
                if following.rows == 1
                else f"rows {_n(following.row)} to {_n(following.row + following.rows - 1)}"
            )
            columns = (
                f"column {_n(following.column)}"
                if following.width == 1
                else f"columns {_n(following.column)} to {_n(following.column + following.width - 1)}"
            )
            shown = ", ".join("" if len(arg) == 0 else _shown_value(value_of(arg)) for arg in callee.args)
            push(
                "hostArgumentOutOfRange",
                f"{callee.name}({shown}) on {block.text} reaches {rows}, {columns}, off the sheet. This will raise "
                "Run-time error '1004': Application-defined or object-defined error.",
                args_span,
            )
            return
    # On a range, 0 and below reach above or left of it, and are judged only on
    # a range the code spells out (XLIDE issue #275).
    if (
        callee.receiver == "Excel.Range"
        and lower in _RANGE_RELATIVE_MEMBERS
        and _check_before_range(span, toks, callee, value_of, push)
    ):
        return
    from_row: Number = origin.row if origin is not None else 1
    from_column: Number = origin.column if origin is not None else 1
    from_text = f" from {origin.text}" if origin is not None else ""
    if lower == "cells" and callee.returns == "Excel.Range":
        # A column given by its letters: "AB" and "$a" run; "XFE", "AAAA", "" and
        # "A " raise 13; "A1" and "5" raise 1004 (XLIDE issue #243).
        column_arg = callee.args[1] if len(callee.args) > 1 else None
        column_text = string_of(column_arg) if column_arg is not None else None
        if column_arg is not None and len(column_arg) == 1 and column_text is not None:
            letters = _LETTERS_COLUMN.search(column_text)
            if letters is None or _column_number(letters.group(1)) > _EXCEL_MAX_COLUMN:
                digits = _ASCII_DIGIT.search(column_text) is not None
                push(
                    "hostArgumentOutOfRange",
                    f"Cells takes a column as a number or its letters, and \"{column_text}\" is neither. This will "
                    "raise Run-time error '1004': Application-defined or object-defined error."
                    if digits
                    else f"\"{column_text}\" names no column: the letters run A to XFD. This will raise Run-time "
                    "error '13': Type mismatch.",
                    _arg_span(span, column_arg),
                )
                return
        for arg in callee.args:
            value = value_of(arg)
            if value is not None and value < 1:
                push(
                    "hostArgumentOutOfRange",
                    f"Cells takes a row and a column of at least 1; {_n(value)} names no cell. This will raise "
                    "Run-time error '1004': Application-defined or object-defined error.",
                    _arg_span(span, arg),
                )
                return
        if len(callee.args) == 2:
            row = value_of(callee.args[0])
            column = value_of(callee.args[1])
            edge = _past_sheet_edge(
                None if row is None else from_row + row - 1, None if column is None else from_column + column - 1
            )
            if edge:
                push(
                    "hostArgumentOutOfRange",
                    f"Cells({_shown_value(row)}, {_shown_value(column)}){from_text} {edge}. This will raise Run-time "
                    "error '1004': Application-defined or object-defined error.",
                    args_span,
                )
        return
    if (lower == "rows" or lower == "columns") and callee.returns == "Excel.Range" and len(callee.args) == 1:
        # `Columns("XFE")`: a letter past XFD names no column (XLIDE issue #276,
        # measured in Excel 16.0: 13).
        letters_text = string_of(callee.args[0]) if lower == "columns" else None
        if (
            letters_text is not None
            and _LETTERS_ONLY.search(letters_text) is not None
            and from_column + _column_number(letters_text) - 1 > _EXCEL_MAX_COLUMN
        ):
            push(
                "hostArgumentOutOfRange",
                f"Columns(\"{letters_text}\"){from_text} names a column past XFD, the last. This will raise "
                "Run-time error '13': Type mismatch.",
                args_span,
            )
            return
        index = value_of(callee.args[0])
        if index is None:
            edge = None
        elif lower == "rows":
            edge = _past_sheet_edge(from_row + index - 1, None)
        else:
            edge = _past_sheet_edge(None, from_column + index - 1)
        if edge:
            push(
                "hostArgumentOutOfRange",
                f"{callee.name}({_shown_value(index)}){from_text} {edge}. This will raise Run-time error '1004': "
                "Application-defined or object-defined error.",
                args_span,
            )
        elif index is not None and index >= 1:
            # A whole row or column is many cells (XLIDE issue #454).
            _check_multi_cell_as_scalar(source, span, toks, callee, f"{callee.name}({_n(index)})", env, arrays, push)
        return
    if lower == "resize" and callee.receiver == "Excel.Range":
        for arg in callee.args:
            value = value_of(arg)
            if value is not None and value < 1:
                push(
                    "hostArgumentOutOfRange",
                    f"Resize needs at least one row and one column; {_n(value)} gives none. This will raise "
                    "Run-time error '1004': Application-defined or object-defined error.",
                    _arg_span(span, arg),
                )
                return
        rows_value = value_of(callee.args[0]) if len(callee.args) > 0 else None
        columns_value = value_of(callee.args[1]) if len(callee.args) > 1 else None
        edge = _past_sheet_edge(
            None if rows_value is None else from_row + rows_value - 1,
            None if columns_value is None else from_column + columns_value - 1,
        )
        if edge:
            shown = ", ".join(_shown_value(value_of(arg)) for arg in callee.args)
            push(
                "hostArgumentOutOfRange",
                f"Resize({shown}){from_text} {edge}. This will raise Run-time error '1004': Application-defined or "
                "object-defined error.",
                args_span,
            )
        return
    if lower == "offset" and callee.receiver == "Excel.Range":
        offset_origin = _single_cell_receiver(toks, callee.name_index - 1)
        if offset_origin is None:
            return
        row_offset = value_of(callee.args[0]) if len(callee.args) > 0 else 0
        column_offset = value_of(callee.args[1]) if len(callee.args) > 1 else 0
        if row_offset is None or column_offset is None:
            return
        row_at = _js_add(offset_origin.row, row_offset)
        column_at = _js_add(offset_origin.column, column_offset)
        if row_at < 1 or column_at < 1 or row_at > _EXCEL_MAX_ROW or column_at > _EXCEL_MAX_COLUMN:
            push(
                "hostArgumentOutOfRange",
                f"Offset({_n(row_offset)}, {_n(column_offset)}) from {offset_origin.text} lands at row {_n(row_at)}, "
                f"column {_n(column_at)}, off the sheet. This will raise Run-time error '1004': Application-defined "
                "or object-defined error.",
                args_span,
            )
        return
    if lower == "range" and callee.returns == "Excel.Range":
        areas: list[_A1Area | None] = []
        for arg in callee.args:
            text = string_of(arg)
            areas.append(None if text is None else _parse_a1_address(text))
        for k in range(len(callee.args)):
            area = areas[k]
            if area is not None and area.blank:
                push(
                    "hostArgumentOutOfRange",
                    f"Range takes an address or a name, and \"{area.text}\" is blank. This will raise Run-time error "
                    "'1004': Method 'Range' of object failed.",
                    _arg_span(span, callee.args[k]),
                )
                return
            if area is not None and not area.valid and not area.may_be_name:
                push(
                    "hostArgumentOutOfRange",
                    f"\"{area.text}\" is not a cell address Excel accepts: rows run 1 to {_EXCEL_MAX_ROW} and "
                    "columns A to XFD. This will raise Run-time error '1004': Method 'Range' of object failed.",
                    _arg_span(span, callee.args[k]),
                )
                return
        first_area = areas[0] if len(areas) > 0 else None
        if len(callee.args) == 1 and first_area is not None and first_area.valid and first_area.multi_cell:
            _check_multi_cell_as_scalar(
                source, span, toks, callee, f"Range(\"{first_area.text}\")", env, arrays, push
            )
        # `Range("A1", "B2")`: two different cells span a block (XLIDE issue #454).
        a = areas[0] if len(areas) > 0 else None
        b = areas[1] if len(areas) > 1 else None
        if (
            len(callee.args) == 2
            and a is not None
            and a.valid
            and b is not None
            and b.valid
            and not a.multi_cell
            and not b.multi_cell
            and a.row is not None
            and b.row is not None
            and (a.row != b.row or a.column != b.column)
        ):
            _check_multi_cell_as_scalar(
                source, span, toks, callee, f"Range(\"{a.text}\", \"{b.text}\")", env, arrays, push
            )
        # `Range(Cells(1, 1), Cells(2, 1))` the same way (XLIDE issue #492).
        if len(callee.args) == 2:
            from_cell = _cells_call(callee.args[0], value_of)
            to_cell = _cells_call(callee.args[1], value_of)
            if (
                from_cell is not None
                and to_cell is not None
                and (from_cell[0] != to_cell[0] or from_cell[1] != to_cell[1])
            ):
                shown = (
                    f"Range(Cells({_n(from_cell[0])}, {_n(from_cell[1])}), "
                    f"Cells({_n(to_cell[0])}, {_n(to_cell[1])}))"
                )
                _check_multi_cell_as_scalar(source, span, toks, callee, shown, env, arrays, push)


def _shown_value(value: Number | None) -> str:
    """`value ?? '...'` in a template literal."""
    return "..." if value is None else _n(value)


def _check_multi_cell_as_scalar(
    source: str,
    span: Span,
    toks: Sequence[VbaToken],
    callee: _HostCallee,
    display: str,
    env: Mapping[str, str],
    arrays: AbstractSet[str],
    push: PushFn,
) -> None:
    """`Range("A1:B2")` read as a value is a two-dimensional array (XLIDE issue
    #122): assigned to a scalar variable, or combined with a scalar operator, it
    is a type mismatch. `.Value` after it changes nothing."""
    end = callee.close_index
    if _raw_at(toks, end + 1) == "." and token_text(_at(toks, end + 2)) in ("value", "value2"):
        end += 2
    if _raw_at(toks, end + 1) == "." or _raw_at(toks, end + 1) == "(":
        return  # a member or an index: not the range read as a value
    start = callee.name_index if callee.receiver == "global" else _receiver_start(toks, callee.name_index)
    value_span = Span(span.start + toks[start].start, span.start + toks[end].end)

    def message(use: str) -> str:
        return (
            f"{display} read as a value is a two-dimensional array, {use}. "
            "This will raise Run-time error '13': Type mismatch."
        )

    # `Len(Range("A1:A2"))`, `InStr(1, Range("A1:A2"), "a")`: a whole argument of
    # a built-in that takes a single value (XLIDE issue #454).
    whole_argument = (_raw_at(toks, start - 1) or "") in ("(", ",") and (_raw_at(toks, end + 1) or "") in (")", ",")
    call = _enclosing_builtin(toks, start) if whole_argument else None
    if call is not None and call[0] in _SCALAR_ARGUMENT_BUILTINS:
        push("multiCellRangeAsScalar", message(f"which {call[1]} cannot take as one value"), value_span)
        return
    # `Select Case Range("A1:A2")` compares the array with each Case.
    head_index = first_executable_token_index(toks)
    if (
        token_text(_at(toks, head_index)) == "select"
        and token_text(_at(toks, head_index + 1)) == "case"
        and start == head_index + 2
        and end == len(toks) - 1
    ):
        push("multiCellRangeAsScalar", message("which Select Case cannot compare"), value_span)
        return
    bare = bare_assignment_target(source, span)
    if bare is not None:
        eq = _index_of_equals(toks)
        if eq == start - 1 and end == len(toks) - 1:
            declared = env.get(bare[0].lower())
            target = normalize_type(declared)
            if target and target in _SCALAR_TYPES:
                # Into an array of another element type it is its Variant
                # elements that do not fit (XLIDE issue #194).
                holder = (
                    f"whose Variant elements an array of {declared} cannot take"
                    if bare[0].lower() in arrays
                    else f"which a {declared} variable cannot hold"
                )
                push("multiCellRangeAsScalar", message(holder), value_span)
            return
        if eq < 0 or start <= eq:
            return
    else:
        head = token_text(_at(toks, first_executable_token_index(toks)))
        if head not in _CONDITION_HEADS:
            return
        # Only the condition of a one-line If is judged here: a range after Then
        # or Else belongs to that branch's own statement, `If r Is Nothing Then
        # Set r = ws.Range("A1:P36")` (XLIDE issue #140).
        then = next((k for k, tok in enumerate(toks) if token_text(tok) == "then"), -1) if head == "if" else -1
        if then > 0 and start > then:
            return
    # The whole condition: `If ws.Range("A1:A2") Then` reads the array as True
    # or False (XLIDE issue #492, measured in Excel 16.0).
    opener = token_text(_at(toks, start - 1))
    if (
        bare is None
        and opener in ("if", "elseif", "while", "until")
        and (end == len(toks) - 1 or token_text(_at(toks, end + 1)) == "then")
    ):
        keyword = {"elseif": "ElseIf", "if": "If", "while": "While"}.get(opener, "Until")
        push("multiCellRangeAsScalar", message(f"which {keyword} cannot read as True or False"), value_span)
        return
    # The operator on either side, never the assignment's own `=`.
    eq_index = _index_of_equals(toks) if bare is not None else -1
    before = None if start - 1 == eq_index else _at(toks, start - 1)
    after = _at(toks, end + 1)
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


# VBA built-ins that read each argument as one value, so a multi-cell Range
# given whole raises 13 (XLIDE issue #454, measured in Excel 16.0: CStr, Len,
# LenB, Val, CLng, CDbl, CBool, Trim, UCase, Left, InStr, Abs, Int, Format and
# Hex; the others here take the same kind of argument). IsEmpty, IsNumeric,
# IsArray, TypeName, VarType and UBound take the array and run.
_SCALAR_ARGUMENT_BUILTINS: frozenset[str] = frozenset(
    {
        "cstr", "len", "lenb", "val", "clng", "cint", "cdbl", "csng", "ccur", "cbyte", "cbool", "cdate",
        "trim", "ltrim", "rtrim", "ucase", "lcase", "left", "right", "mid", "instr", "abs", "int", "fix",
        "format", "hex", "oct",
    }
)


def _enclosing_builtin(toks: Sequence[VbaToken], index: int) -> tuple[str, str] | None:
    """The VBA built-in whose argument list holds the token at `index`, bare or
    VBA-qualified, `$` spellings included: its lowercased name and its display."""
    depth = 0
    for i in range(index - 1, -1, -1):
        raw = toks[i].raw_text
        if raw == ")":
            depth += 1
        elif raw == "(":
            if depth == 0:
                name_at = i - 2 if _raw_at(toks, i - 1) == "$" else i - 1
                name = token_name(_at(toks, name_at))
                qualified = _raw_at(toks, name_at - 1) == "."
                if not name or (qualified and token_text(_at(toks, name_at - 2)) != "vba"):
                    return None
                return name.lower(), "".join(tok.raw_text for tok in toks[name_at:i])
            depth -= 1
    return None


def _check_sheet_name_assignment(
    source: str,
    span: Span,
    toks: Sequence[VbaToken],
    member_ctx: MemberCompletionContext,
    source_names: AbstractSet[str],
    push: PushFn,
) -> None:
    """`Worksheets(1).Name = "a:b"`: the receiver's type and the literal decide."""
    n = len(toks)
    eq = _index_of_equals(toks)
    if eq < 2 or eq == n - 1 or token_text(toks[eq - 1]) != "name" or toks[eq - 2].raw_text != ".":
        return
    name = _spelled_out_text(toks[eq + 1 :], source_names)
    if name is None:
        return
    resolved = resolve_receiver_type_at(source, span.start + toks[eq - 2].end, member_ctx)
    # `Worksheets(1)` is a one-part union, `Sheets(1)` a Worksheet-or-Chart union: both are sheets.
    parts: list[str] = []
    if resolved:
        parts = resolved[len("union:") :].split("|") if resolved.startswith("union:") else [resolved]
    if len(parts) == 0 or not all(part == "Excel.Worksheet" or part == "Excel.Chart" for part in parts):
        return
    # JavaScript measures a string in UTF-16 code units, as Excel counts a name.
    length = utf16_length(name)
    problem: str | None = None
    error = "You typed an invalid name for a sheet or chart."
    if length == 0:
        problem = "a sheet name cannot be blank"
    elif length > _SHEET_NAME_MAX:
        problem = f"a sheet name has at most {_SHEET_NAME_MAX} characters, and this one has {length}"
    elif _SHEET_NAME_FORBIDDEN.search(name) is not None:
        problem = "a sheet name cannot contain any of : \\ / ? * [ ]"
    elif name.lower() == "history":
        problem = "Excel keeps History for itself, in any case"
        error = "History is a reserved name."
    elif name.startswith("'") or name.endswith("'"):
        problem = "a sheet name cannot start or end with an apostrophe"
    if problem:
        push(
            "sheetNameInvalid",
            f"Excel refuses this name: {problem}. This will raise Run-time error '1004': {error}",
            Span(span.start + toks[eq + 1].start, span.start + toks[n - 1].end),
        )


def _spelled_out_text(toks: Sequence[VbaToken], source_names: AbstractSet[str]) -> str | None:
    """The text an expression spells out from literals alone: a string literal,
    `String$(32, "a")`, `Space$(3)`, and `&` between them (XLIDE issue #276).
    None for anything else, or when the module declares its own String or Space."""
    parts: list[list[VbaToken]] = [[]]
    depth = 0
    for tok in toks:
        if tok.kind is TokenKind.COMMENT:
            continue
        depth += 1 if tok.raw_text == "(" else -1 if tok.raw_text == ")" else 0
        if depth == 0 and tok.raw_text == "&":
            parts.append([])
        else:
            parts[-1].append(tok)
    out: list[str] = []
    for written in parts:
        # `String$` lexes as String and a `$`.
        part = [written[0], *written[2:]] if _raw_at(written, 1) == "$" else written
        if len(part) == 1 and part[0].kind is TokenKind.STRING_LITERAL:
            out.append(string_literal_value(part[0].raw_text))
            continue
        fn = token_text(_at(part, 0))
        if fn.endswith("$"):
            fn = fn[:-1]
        count = (
            parse_vba_integer_literal(part[2].raw_text)
            if _kind_at(part, 2) is TokenKind.INTEGER_LITERAL
            else None
        )
        if fn in source_names or _raw_at(part, 1) != "(" or count is None or count < 0 or count > 1000:
            return None
        if fn == "space" and len(part) == 4 and part[3].raw_text == ")":
            out.append(" " * count)
        elif (
            fn == "string"
            and len(part) == 6
            and part[3].raw_text == ","
            and part[4].kind is TokenKind.STRING_LITERAL
            and part[5].raw_text == ")"
            and string_literal_value(part[4].raw_text) != ""
        ):
            out.append(string_literal_value(part[4].raw_text)[0] * count)
        else:
            return None
    return "".join(out)


def _past_sheet_edge(row: Number | None, column: Number | None) -> str | None:
    """Where a row or column past the bottom or right edge of the sheet lands, in words."""
    if row is not None and row > _EXCEL_MAX_ROW:
        return f"reaches row {_n(row)}, past the last row of the sheet, {_EXCEL_MAX_ROW}"
    if column is not None and column > _EXCEL_MAX_COLUMN:
        return f"reaches column {_n(column)}, past the last column of the sheet, {_EXCEL_MAX_COLUMN} (XFD)"
    return None


def _single_cell_receiver(toks: Sequence[VbaToken], dot_index: int) -> _CellBlock | None:
    """The single-cell literal `Range("B2")` before the dot at `dot_index`, when that is the receiver."""
    block = _literal_range_receiver(toks, dot_index)
    return block if block is not None and block.rows == 1 and block.width == 1 else None


def _cells_call(arg: Sequence[VbaToken], value_of: ValueOf) -> tuple[Number, Number] | None:
    """`Cells(1, 1)` or `ws.Cells(1, 1)` with known numbers, as an argument of Range: its row and column."""
    toks = _significant(arg)
    at = 2 if len(toks) > 2 and _raw_at(toks, 1) == "." else 0
    if (
        token_text(_at(toks, at)) != "cells"
        or _raw_at(toks, at + 1) != "("
        or match_paren_from(toks, at + 1) != len(toks) - 1
    ):
        return None
    args = split_top_level_token_groups(toks, at + 2, ",", len(toks) - 1)
    row = value_of(args[0]) if len(args) == 2 else None
    column = value_of(args[1]) if len(args) == 2 else None
    if row is not None and column is not None and row >= 1 and column >= 1:
        return row, column
    return None


# The Range members a chain follows (XLIDE issue #508).
_CHAIN_MEMBERS: frozenset[str] = frozenset(
    {"offset", "resize", "cells", "item", "rows", "columns", "entirerow", "entirecolumn"}
)


def _range_chain_receiver(
    toks: Sequence[VbaToken],
    dot_index: int,
    value_of: ValueOf,
    source_names: AbstractSet[str],
    memo: dict[int, _CellBlock | None],
) -> _CellBlock | None:
    """The block of cells the Range expression before the dot at `dot_index`
    names, followed through Offset, Resize, Cells, Item, Rows, Columns, EntireRow
    and EntireColumn from a literal `Range("B2:C3")` (XLIDE issue #508, measured
    in Excel 16.0). None where any step is not known, and where a step already
    lands off the sheet, which that step reports.

    Upstream recurses once per link; this walks back to the chain's start, then
    applies each link on the way out. `memo` holds the block before each dot of
    one statement, so the callees along one chain share the work."""
    # Each pending link: the dot it ends at, and what it does to the block before it.
    pending: list[tuple[int, Callable[[_CellBlock | None], _CellBlock | None]]] = []
    d = dot_index
    result: _CellBlock | None = None
    while True:
        if d in memo:
            result = memo[d]
            break
        if _raw_at(toks, d) != ".":
            result = None
            break
        last = _at(toks, d - 1)
        word = token_text(last)
        if (word == "entirerow" or word == "entirecolumn") and _raw_at(toks, d - 2) == ".":
            assert last is not None
            pending.append((d, _entire_link(word, last.raw_text)))
            d -= 2
            continue
        # `Cells`, `Rows` and `Columns` of a sheet, without an index: the whole
        # sheet (XLIDE issue #628, measured in Excel 16.0).
        if (word == "cells" or word == "rows" or word == "columns") and _of_sheet(toks, d - 1, source_names):
            assert last is not None
            mode = "rows" if word == "rows" else "columns" if word == "columns" else None
            result = _CellBlock(1, 1, _EXCEL_MAX_ROW, _EXCEL_MAX_COLUMN, last.raw_text, mode)
            break
        if last is None or last.raw_text != ")":
            result = None
            break
        close = d - 1
        open_index = _open_paren_for(toks, close)
        name = token_text(_at(toks, open_index - 1))
        if open_index < 1:
            result = None
            break
        if name == "range":
            # A Range on a range counts from that range: not followed.
            block = _literal_range_receiver(toks, d)
            if block is None:
                block = _whole_lines_at(toks, open_index - 1)
            if _raw_at(toks, open_index - 2) == ".":
                pending.append((d, _range_gate(block)))
                d = open_index - 2
                continue
            result = block
            break
        # `Cells(1, 2)`, `Rows(3)`, `Columns(2)` of the sheet: a cell, a whole row
        # or a whole column (XLIDE issues #308 and #628, measured in Excel 16.0).
        if _of_sheet(toks, open_index - 1, source_names) and name in ("cells", "rows", "columns"):
            result = _sheet_member_block(toks, open_index, close, name, value_of)
            break
        if name not in _CHAIN_MEMBERS or _raw_at(toks, open_index - 2) != ".":
            result = None
            break
        pending.append((d, _step_link(toks, open_index, close, name, value_of)))
        d = open_index - 2
    memo[d] = result
    for dot, link in reversed(pending):
        result = link(result)
        memo[dot] = result
    return result


def _entire_link(word: str, written: str) -> Callable[[_CellBlock | None], _CellBlock | None]:
    def apply(inner: _CellBlock | None) -> _CellBlock | None:
        if inner is None:
            return None
        text = f"{inner.text}.{written}"
        if word == "entirerow":
            return _CellBlock(inner.row, 1, inner.rows, _EXCEL_MAX_COLUMN, text, "rows")
        return _CellBlock(1, inner.column, _EXCEL_MAX_ROW, inner.width, text, "columns")

    return apply


def _range_gate(block: _CellBlock | None) -> Callable[[_CellBlock | None], _CellBlock | None]:
    def apply(inner: _CellBlock | None) -> _CellBlock | None:
        return None if inner is not None else block

    return apply


def _step_link(
    toks: Sequence[VbaToken], open_index: int, close: int, name: str, value_of: ValueOf
) -> Callable[[_CellBlock | None], _CellBlock | None]:
    def apply(inner: _CellBlock | None) -> _CellBlock | None:
        if inner is None:
            return None
        args = split_top_level_token_groups(toks, open_index + 1, ",", close) if close > open_index + 1 else []
        stepped = _chain_step(inner, name, args, value_of)
        if stepped is None or _off_sheet(stepped):
            return None
        written = "".join(tok.raw_text for tok in toks[open_index - 1 : close + 1])
        return dataclasses.replace(stepped, text=f"{inner.text}.{written}")

    return apply


def _sheet_member_block(
    toks: Sequence[VbaToken], open_index: int, close: int, name: str, value_of: ValueOf
) -> _CellBlock | None:
    args = split_top_level_token_groups(toks, open_index + 1, ",", close) if close > open_index + 1 else []
    first = value_of(args[0]) if len(args) > 0 else None
    text = "".join(tok.raw_text for tok in toks[open_index - 1 : close + 1])
    if name == "cells" and len(args) == 2:
        column = value_of(args[1])
        if (
            first is not None
            and column is not None
            and first >= 1
            and column >= 1
            and first <= _EXCEL_MAX_ROW
            and column <= _EXCEL_MAX_COLUMN
        ):
            return _CellBlock(first, column, 1, 1, text)
        return None
    if len(args) != 1 or first is None or first < 1:
        return None
    # Item on a row counts rows: `Rows(5).Item(2)` is row 6.
    if name == "rows":
        return _CellBlock(first, 1, 1, _EXCEL_MAX_COLUMN, text, "rows") if first <= _EXCEL_MAX_ROW else None
    if name == "columns" and first <= _EXCEL_MAX_COLUMN:
        return _CellBlock(1, first, _EXCEL_MAX_ROW, 1, text, "columns")
    return None


def _of_sheet(toks: Sequence[VbaToken], at: int, source_names: AbstractSet[str]) -> bool:
    """Whether the member at `toks[at]` is the sheet's own: unqualified, or after
    `ActiveSheet`, `Worksheets(...)` or `Sheets(...)` (XLIDE issue #628).
    Unqualified, a name the code declares is its own."""
    if _raw_at(toks, at - 1) != ".":
        return token_text(_at(toks, at)) not in source_names
    before = _at(toks, at - 2)
    if token_text(before) == "activesheet":
        return True
    if before is None or before.raw_text != ")":
        return False
    open_index = _open_paren_for(toks, at - 2)
    name = token_text(_at(toks, open_index - 1))
    return open_index >= 1 and (name == "worksheets" or name == "sheets")


def _whole_lines_at(toks: Sequence[VbaToken], at: int) -> _CellBlock | None:
    """`Range("5:6")` or `Range("C:D")` starting at `toks[at]`: whole rows or whole
    columns (XLIDE issue #628)."""
    if (
        token_text(_at(toks, at)) != "range"
        or _raw_at(toks, at + 1) != "("
        or _kind_at(toks, at + 2) is not TokenKind.STRING_LITERAL
        or _raw_at(toks, at + 3) != ")"
    ):
        return None
    text = string_literal_value(toks[at + 2].raw_text)
    rows = _WHOLE_ROWS.search(text)
    if rows is not None:
        top, bottom = sorted((_decimal_number(rows.group(1)), _decimal_number(rows.group(2))))
        if top >= 1 and bottom <= _EXCEL_MAX_ROW:
            return _CellBlock(top, 1, bottom - top + 1, _EXCEL_MAX_COLUMN, f"Range(\"{text}\")")
        return None
    columns = _WHOLE_COLUMNS.search(text)
    if columns is not None:
        left, right = sorted((_column_number(columns.group(1)), _column_number(columns.group(2))))
        if right <= _EXCEL_MAX_COLUMN:
            return _CellBlock(1, left, _EXCEL_MAX_ROW, right - left + 1, f"Range(\"{text}\")")
        return None
    return None


def _js_trunc_div(a: Number, b: Number) -> Number:
    """`Math.trunc(a / b)`."""
    if isinstance(a, int) and isinstance(b, int):
        quotient = abs(a) // abs(b)
        return quotient if (a < 0) == (b < 0) else -quotient
    return math.trunc(a / b)


def _js_rem(a: Number, b: Number) -> Number:
    """`a % b`: the remainder takes the sign of the dividend."""
    if isinstance(a, int) and isinstance(b, int):
        remainder = abs(a) % abs(b)
        return -remainder if a < 0 else remainder
    return math.fmod(a, b)


def _chain_step(
    block: _CellBlock, name: str, args: Sequence[Sequence[VbaToken]], value_of: ValueOf
) -> _CellBlock | None:
    """One member applied to a block: the block it names, or None where an argument is not known."""

    def value(k: int, missing: Number) -> Number | None:
        return missing if k >= len(args) or len(args[k]) == 0 else value_of(args[k])

    if name == "offset":
        rows = value(0, 0)
        columns = value(1, 0)
        if rows is None or columns is None:
            return None
        return dataclasses.replace(block, row=block.row + rows, column=block.column + columns)
    if name == "resize":
        rows = value(0, block.rows)
        width = value(1, block.width)
        if rows is None or width is None or rows < 1 or width < 1:
            return None
        return dataclasses.replace(block, rows=rows, width=width)
    if name == "cells" or name == "item":
        if len(args) == 2:
            row = value(0, 1)
            column = value(1, 1)
            if row is None or column is None:
                return None
            return _CellBlock(block.row + row - 1, block.column + column - 1, 1, 1)
        index = value(0, 1) if len(args) == 1 else None
        if index is None:
            return None
        if name == "item" and block.mode == "columns":
            return dataclasses.replace(block, column=block.column + index - 1, width=1)
        if name == "item" and block.mode == "rows":
            return dataclasses.replace(block, row=block.row + index - 1, rows=1)
        k = index - 1
        return _CellBlock(block.row + _js_trunc_div(k, block.width), block.column + _js_rem(k, block.width), 1, 1)
    if name == "rows":
        index = value(0, 1) if len(args) == 1 else None
        if index is None:
            return None
        return dataclasses.replace(block, row=block.row + index - 1, rows=1, mode="rows")
    if name == "columns":
        index = value(0, 1) if len(args) == 1 else None
        if index is None:
            return None
        return dataclasses.replace(block, column=block.column + index - 1, width=1, mode="columns")
    return None


def _off_sheet(block: _CellBlock) -> bool:
    """Whether any cell of the block is off the sheet."""
    return (
        block.row < 1
        or block.column < 1
        or block.row + block.rows - 1 > _EXCEL_MAX_ROW
        or block.column + block.width - 1 > _EXCEL_MAX_COLUMN
    )


def _literal_range_receiver(toks: Sequence[VbaToken], dot_index: int) -> _CellBlock | None:
    """The literal `Range("B2:C3")` before the dot at `dot_index`: its top-left cell and its size."""
    if _raw_at(toks, dot_index) != "." or _raw_at(toks, dot_index - 1) != ")":
        return None
    open_index = _open_paren_for(toks, dot_index - 1)
    return None if open_index < 1 else _literal_range_at(toks, open_index - 1)


def _literal_range_at(toks: Sequence[VbaToken], at: int) -> _CellBlock | None:
    """`Range("B2:C3")` starting at `toks[at]`: its top-left cell and its size."""
    if (
        token_text(_at(toks, at)) != "range"
        or _raw_at(toks, at + 1) != "("
        or _kind_at(toks, at + 2) is not TokenKind.STRING_LITERAL
        or _raw_at(toks, at + 3) != ")"
    ):
        return None
    area = _parse_a1_address(string_literal_value(toks[at + 2].raw_text))
    if area is None or not area.valid or area.row is None or area.column is None:
        return None
    end_row = area.end_row if area.end_row is not None else area.row
    end_column = area.end_column if area.end_column is not None else area.column
    return _CellBlock(
        min(area.row, end_row),
        min(area.column, end_column),
        abs(end_row - area.row) + 1,
        abs(end_column - area.column) + 1,
        f"Range(\"{area.text}\")",
    )


def _check_before_range(
    span: Span, toks: Sequence[VbaToken], callee: _HostCallee, value_of: ValueOf, push: PushFn
) -> bool:
    """Cells, Item, Rows or Columns on a range at 0 or below (XLIDE issue #275,
    measured in Excel 16.0). They count from the range's top-left cell, so they
    raise 1004 only where they land above row 1 or left of column A, which is
    known for a literal `Range("B2")` receiver alone. One index over a range w
    columns wide is Cells((i - 1) \\ w + 1, (i - 1) Mod w + 1), with VBA's
    truncating \\ and Mod. Returns True when an index is 0 or below, reported or
    not, so the caller's checks for a sheet's own Cells do not run."""
    values = [value_of(arg) for arg in callee.args]
    if not any(value is not None and value < 1 for value in values):
        return False
    block = _literal_range_receiver(toks, callee.name_index - 1)
    lower = callee.name.lower()
    if block is None or len(callee.args) > 2 or ((lower == "rows" or lower == "columns") and len(callee.args) != 1):
        return True
    row: Number | None = None
    column: Number | None = None
    first = values[0]
    if lower == "rows":
        assert first is not None
        row = block.row + first - 1
    elif lower == "columns":
        assert first is not None
        column = block.column + first - 1
    elif len(callee.args) == 2:
        second = values[1]
        row = None if first is None else block.row + first - 1
        column = None if second is None else block.column + second - 1
    else:
        assert first is not None
        k = first - 1
        row = block.row + _js_trunc_div(k, block.width)
        column = block.column + _js_rem(k, block.width)
    if (row is not None and row < 1) or (column is not None and column < 1):
        where = (
            f"row {_n(row)}, above row 1"
            if row is not None and row < 1
            else f"column {_shown_column(column)}, left of column A"
        )
        shown = ", ".join(_shown_value(value) for value in values)
        push(
            "hostArgumentOutOfRange",
            f"{callee.name}({shown}) counts from the top-left cell of {block.text} and lands at {where}. This will "
            "raise Run-time error '1004': Application-defined or object-defined error.",
            Span(span.start + toks[callee.open_index + 1].start, span.start + toks[callee.close_index - 1].end),
        )
    return True


def _shown_column(column: Number | None) -> str:
    """`${column}` in a template literal: `undefined` for a column not known."""
    return "undefined" if column is None else _n(column)


def _receiver_start(toks: Sequence[VbaToken], name_index: int) -> int:
    """Index of the first token of the receiver chain that ends at the dot before `name_index`."""
    j = name_index
    while j >= 2 and toks[j - 1].raw_text == ".":
        j -= 1
        if _raw_at(toks, j - 1) == ")":
            open_index = _open_paren_for(toks, j - 1)
            j = open_index if open_index >= 1 else j
        j -= 1
    return j


def _parse_a1_address(text: str) -> _A1Area | None:
    """Parses an A1-style address literal: `A1`, `$A$1`, `A1:B2`, `A:A`, `1:1`,
    with an optional `Sheet1!` or `'My Sheet'!` prefix. Anything else (a name,
    an R1C1 address, a union) is not judged."""
    if js_trim(text) == "":
        return _A1Area(text, False, False, blank=True)
    body = _SHEET_PREFIX.sub("", text, count=1)
    # "R1C1" is an R1C1-style reference, which Range does not read and no
    # workbook name may be (XLIDE issue #276, measured in Excel 16.0: 1004).
    if _R1C1_BODY.search(body) is not None:
        return _A1Area(text, False, False)
    parts = body.split(":")
    if len(parts) > 2:
        return None
    # "A1:", ":A1" and ":" leave a side empty (XLIDE issue #243).
    if len(parts) == 2 and any(part == "" for part in parts):
        return _A1Area(text, False, False)
    cell_matches = [_A1_CELL.match(part) for part in parts]
    cells = [match for match in cell_matches if match is not None]
    if len(cells) == len(cell_matches):
        rows = [_decimal_number(match.group(2)) for match in cells]
        columns = [_column_number(match.group(1)) for match in cells]
        in_range = [
            1 <= rows[k] <= _EXCEL_MAX_ROW and 1 <= columns[k] <= _EXCEL_MAX_COLUMN for k in range(len(parts))
        ]
        valid = all(in_range)
        may_be_name = all(in_range[k] or "$" not in part for k, part in enumerate(parts))
        multi_cell = len(cells) == 2 and (rows[0] != rows[1] or columns[0] != columns[1])
        return _A1Area(
            text,
            valid,
            multi_cell,
            may_be_name=may_be_name,
            row=rows[0],
            column=columns[0],
            end_row=rows[1] if len(rows) > 1 else None,
            end_column=columns[1] if len(columns) > 1 else None,
        )
    if len(parts) == 2:
        column_matches = [_A1_COLUMN.match(part) for part in parts]
        columns_only = [match for match in column_matches if match is not None]
        if len(columns_only) == len(column_matches):
            column_in_range = [_column_number(match.group(1)) <= _EXCEL_MAX_COLUMN for match in columns_only]
            valid = all(column_in_range)
            may_be_name = all(column_in_range[k] or "$" not in part for k, part in enumerate(parts))
            return _A1Area(text, valid, True, may_be_name=may_be_name)
        row_matches = [_A1_ROW.match(part) for part in parts]
        rows_only = [match for match in row_matches if match is not None]
        if len(rows_only) == len(row_matches):
            valid = all(1 <= _decimal_number(match.group(1)) <= _EXCEL_MAX_ROW for match in rows_only)
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


@dataclass(frozen=True, slots=True)
class WorkbookSheetsCheck:
    """The saved workbook's sheets, and what the project's code may do to them."""

    sheets: Sequence[WorkbookSheetInfo]
    changes: SheetChanges


def workbook_sheets_to_check(opts: AnalyzeModuleOptions) -> WorkbookSheetsCheck | None:
    """Both halves or nothing: sheets alone cannot say what code adds at run time."""
    if opts.workbook_sheets is not None and opts.project_sheet_changes is not None:
        return WorkbookSheetsCheck(opts.workbook_sheets, opts.project_sheet_changes)
    return None


# The sheet kinds each of ThisWorkbook's sheet collections holds; None is every kind.
_SHEET_COLLECTION_KINDS: Mapping[str, str | None] = {
    "sheets": None,
    "worksheets": "worksheet",
    "charts": "chartsheet",
}

_SHEET_KIND_WORDS: Mapping[str, str] = {
    "worksheet": "worksheet",
    "chartsheet": "chart sheet",
    "dialogsheet": "dialog sheet",
    "macrosheet": "macro sheet",
}

# Names Excel gives a sheet that code adds or copies: Sheet4, Chart2, Sheet1 (2).
_MADE_SHEET_NAME = re.compile(
    r"^(sheet|chart|dialog|macro)[0-9]+\Z|[" + _WS + r"]\([0-9]+\)\Z", re.IGNORECASE | re.ASCII
)
_NON_PRINTABLE_ASCII = re.compile(r"[^\x20-\x7e]")


def _check_workbook_sheet_access(
    source: str,
    span: Span,
    workbook: WorkbookSheetsCheck,
    source_names: AbstractSet[str],
    value_of: ValueOf,
    push: PushFn,
) -> None:
    """`ThisWorkbook.Sheets("Missing")` and `ThisWorkbook.Worksheets(9)` on a
    workbook without that sheet raise 9 (XLIDE issue #229). ThisWorkbook only: a
    bare `Sheets` is the active workbook's, which may be any workbook. A name or
    an index that code in the project could have made - by adding, copying or
    naming a sheet - is left alone."""
    if "thisworkbook" in source_names:
        return
    toks = statement_tokens(source, span)

    def lower(i: int) -> str:
        tok = _at(toks, i)
        return tok.raw_text.lower() if tok is not None else ""

    for i in range(len(toks)):
        if lower(i) != "thisworkbook" or lower(i + 1) != ".":
            continue
        # `Application.ThisWorkbook` is the same object; `x.ThisWorkbook` is not known.
        if lower(i - 1) == "." and not (lower(i - 2) == "application" and lower(i - 3) != "."):
            continue
        collection = lower(i + 2)
        if collection not in _SHEET_COLLECTION_KINDS:
            continue
        kind = _SHEET_COLLECTION_KINDS[collection]
        open_index = i + 3
        if lower(open_index) == "." and lower(open_index + 1) == "item":
            open_index += 2
        if lower(open_index) != "(":
            continue
        close = match_paren_from(toks, open_index)
        if close < 0:
            continue
        args = _split_top_level(toks[open_index + 1 : close])
        if len(args) != 1:
            continue
        held = [sheet for sheet in workbook.sheets if kind is None or sheet.kind == kind]
        arg_tokens = _significant(args[0])
        where = Span(span.start + toks[open_index + 1].start, span.start + toks[close - 1].end)
        what = "worksheet" if collection == "worksheets" else "chart sheet" if collection == "charts" else "sheet"
        if len(arg_tokens) == 1 and arg_tokens[0].kind is TokenKind.STRING_LITERAL:
            name = arg_tokens[0].raw_text[1:-1].replace('""', '"')
            # Excel matches names without regard to case; outside ASCII its rule is not known here.
            if _NON_PRINTABLE_ASCII.search(name) is not None or any(
                sheet.name.lower() == name.lower() for sheet in held
            ):
                continue
            changes = workbook.changes
            if (
                changes.assigns_computed_name
                or name.lower() in changes.names_assigned
                or (changes.adds_sheets and _MADE_SHEET_NAME.search(name) is not None)
            ):
                continue
            other = next((sheet for sheet in workbook.sheets if sheet.name.lower() == name.lower()), None)
            detail = (
                f"'{other.name}' is a {_SHEET_KIND_WORDS[other.kind]}, not a {what}"
                if other is not None
                else f"this workbook has no {what} named '{name}'"
            )
            push(
                "sheetNotInWorkbook",
                f"{_capitalize(detail)}. This will raise Run-time error '9': Subscript out of range.",
                where,
            )
            continue
        index = value_of(args[0])
        # Index 0 and below are host-argument-out-of-range's.
        if index is None or index < 1 or index <= len(held) or workbook.changes.adds_sheets:
            continue
        count = f"1 {what}" if len(held) == 1 else f"{len(held)} {what}s"
        push(
            "sheetNotInWorkbook",
            f"This workbook has {count}, so index {_n(index)} is past the last. This will raise Run-time error '9': "
            "Subscript out of range.",
            where,
        )


def _capitalize(text: str) -> str:
    return text[:1].upper() + text[1:]


def _literal_string(arg: Sequence[VbaToken]) -> str | None:
    """The text of an argument that is one string literal."""
    toks = _significant(arg)
    if len(toks) == 1 and toks[0].kind is TokenKind.STRING_LITERAL:
        return string_literal_value(toks[0].raw_text)
    return None


def _integer_literal_value(arg: Sequence[VbaToken]) -> int | None:
    """The whole-number value of an argument that is a literal, optionally negated."""
    toks = _significant(arg)
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


def _js_add(a: Number, b: Number) -> Number:
    """`a + b` as JavaScript computes it: in doubles, so a sum past 2**53 rounds
    the way the upstream message prints it."""
    if isinstance(a, int) and isinstance(b, int) and abs(a + b) < 2**53:
        return a + b
    total = float(a) + float(b)
    return int(total) if total.is_integer() else total
