"""Rule: an error value read where a number or text is needed (XLIDE issue #310).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/errorValues.ts.

Measured in Excel 16.0 (build 20430, 2026-10-02): a Variant holding an error
value, from CVErr or from Excel, raises 13 as an operand of an arithmetic, `&` or
comparison operator, after Not, as an If condition or a Select Case subject, in
Val, Abs or Len, and Let into a typed local. IsError, TypeName, CStr, CLng and
CInt, and a Let into another Variant, run.

Issue #607 adds, each measured: unary minus, And, Or and Like; IIf's and
Choose's first argument; a Do or For condition or bound; the arguments of Int,
Fix, Sgn, Round, Sqr, CDate, Str, Format, Trim, Left, UCase, InStr, Mid, Hex, Chr
and Space, all 13, and of CByte, 6 (2042 does not fit); an array's index; a ByVal
typed parameter of the module's procedure; an element of Array(...) given to
Join; and WorksheetFunction.Sum, 1004. CDbl, CLng, CVar, IsError and
Application.Sum run.

Where the value is known: `CVErr(...)` itself; a Variant local whose
straight-line assignment is one; `Evaluate("1/0")` and `Evaluate("NA()")`; an
element of `Array(...)` that is one; and a cell right after the procedure wrote
`=1/0` or `=NA()` into it, through Formula or Value, as `Range("A1")`,
`ActiveSheet.Range("A1")` or `Cells(1, 1)`. A Variant local given any of these
holds the error from then on, though the cell changes after (issue #607).
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Iterator, Mapping, Sequence
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import statement_label_declaration
from ...js_compat import JS_WHITESPACE, js_number, js_number_to_string
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ModuleNode,
    ProcKind,
    ProcedureNode,
    SelectBlockNode,
    Span,
    StatementNode,
    is_leaf_statement,
)
from ...symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbol, VbaSymbolKind
from ...types.type_inference import procedure_symbol_for
from ...types.type_names import is_known_scalar_type, normalize_type
from ..call_extraction import string_literal_value
from ..context import PushFn, statement_tokens
from ..straight_line_values import ReachingAssignments, straight_line_assignments
from ..walker import (
    active_module_members,
    block_header_line_span,
    statement_and_branch_spans,
    token_name,
    token_text,
)
from .arrays import module_option_base

_BINARY_OPERATORS = frozenset(
    {
        "+", "-", "*", "/", "\\", "^", "mod", "&", "=", "<>", "<", ">", "<=", ">=",
        "and", "or", "xor", "eqv", "imp", "like",
    }
)

_COMPARISONS = frozenset({"=", "<>", "<", ">", "<=", ">="})

# VBA functions that need a number or text of their argument, any of them (issue #607).
_VALUE_FUNCTIONS = frozenset(
    {
        "val", "abs", "len", "int", "fix", "sgn", "round", "sqr", "cdate", "cbyte", "str", "format",
        "trim", "left", "ucase", "instr", "mid", "hex", "chr", "space",
    }
)

# Functions whose first argument is read as a number or a Boolean (issue #607).
_FIRST_ARGUMENT_FUNCTIONS = frozenset({"iif", "choose"})

_JS_SPACE = "[" + JS_WHITESPACE + "]"
# A formula whose value is an error whatever the sheet holds. `[0-9]` and the
# JavaScript whitespace set: upstream's `\d` and `\s`.
_ERROR_FORMULA = re.compile(
    "^" + _JS_SPACE + r"*(?:[0-9]+(?:\.[0-9]+)?" + _JS_SPACE + r"*/" + _JS_SPACE + r"*0|na\("
    + _JS_SPACE + r"*\))" + _JS_SPACE + r"*\Z",
    re.IGNORECASE | re.ASCII,
)
_CELL_ADDRESS = re.compile(r"[a-z]{1,3}[0-9]+")
_GIVEN_WHAT = re.compile(r"^'?([^' ]+)'? ")

_MESSAGE_TAIL = "This will raise Run-time error '13': Type mismatch."

_CELL_CHANGERS = frozenset({"activate", "select", "calculate", "clear", "delete", "insert"})
_CELL_KEEPING_CALLS = frozenset({"range", "cverr", "val", "abs", "len", "iserror", "evaluate"})


@dataclass(frozen=True, slots=True)
class _Found:
    end: int
    what: str


@dataclass(frozen=True, slots=True)
class _Given:
    text: str
    what: str


@dataclass(frozen=True, slots=True)
class _Cell:
    address: str
    end: int


def _literal_case_item(all_toks: Sequence[VbaToken]) -> bool:
    """`Case 1`, `Case "a", x`, `Case 1 To 5`, `Case Is > 2`: a Case with an item of
    literals only."""
    toks = _without_comments(all_toks)
    if token_text(_at(toks, 0)) != "case" or token_text(_at(toks, 1)) == "else":
        return False
    items = split_top_level_token_groups(toks, 1, ",", len(toks)) if len(toks) > 1 else []
    return any(
        len(item) > 0
        and all(
            _is_literal(tok) or token_text(tok) in ("to", "is") or tok.raw_text in _COMPARISONS
            for tok in item
        )
        for item in items
    )


def _is_literal(tok: VbaToken | None) -> bool:
    """A number or text literal."""
    return tok is not None and tok.kind in (
        TokenKind.INTEGER_LITERAL,
        TokenKind.FLOAT_LITERAL,
        TokenKind.STRING_LITERAL,
    )


def check_error_values(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    option_base = module_option_base(mod, activity)
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_procedure(source, member, symbols, activity, option_base, push)


def _check_procedure(
    source: str,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    option_base: int,
    push: PushFn,
) -> None:
    proc_sym = procedure_symbol_for(symbols, member)
    children = (proc_sym.children if proc_sym is not None else None) or []
    # By name, so a procedure of many locals asks in constant time (issue #322).
    locals_by_name: dict[str, VbaSymbol] = {
        child.name.lower(): child
        for child in children
        if child.kind in (VbaSymbolKind.LOCAL_VARIABLE, VbaSymbolKind.PARAMETER)
    }

    def variant_local(lower: str) -> bool:
        child = locals_by_name.get(lower)
        type_ = normalize_type(child.as_type if child is not None else None)
        return (
            child is not None
            and child.kind is VbaSymbolKind.LOCAL_VARIABLE
            and child.visibility is not SymbolVisibility.STATIC
            and not child.is_array
            and (type_ is None or type_ == "variant")
        )

    def scalar_local(lower: str) -> str | None:
        child = locals_by_name.get(lower)
        type_ = normalize_type(child.as_type if child is not None else None)
        if child is not None and not child.is_array and type_ and type_ != "variant" and is_known_scalar_type(type_):
            return child.as_type
        returns = normalize_type(member.return_type)
        if (
            lower == member.name.lower()
            and member.proc_kind is ProcKind.FUNCTION
            and member.return_type
            and returns != "variant"
            and returns is not None
            and is_known_scalar_type(returns)
        ):
            return member.return_type
        return None

    def array_local(lower: str) -> bool:
        child = locals_by_name.get(lower)
        return child is not None and child.kind is VbaSymbolKind.LOCAL_VARIABLE and child.is_array is True

    # `TakeL(v)` with `ByVal p As Long` (issue #607).
    def by_val_scalar_param(lower: str, slot: int) -> str | None:
        proc = next(
            (
                sym
                for sym in (symbols.root.children or [])
                if sym.kind in (VbaSymbolKind.FUNCTION, VbaSymbolKind.SUB) and sym.name.lower() == lower
            ),
            None,
        )
        params = [child for child in ((proc.children if proc is not None else None) or []) if child.kind is VbaSymbolKind.PARAMETER]
        param = params[slot] if 0 <= slot < len(params) else None
        type_ = normalize_type(param.as_type if param is not None else None)
        if param is not None and param.by_val and type_ and type_ != "variant" and is_known_scalar_type(type_):
            return f"{param.name} As {param.as_type}"
        return None

    reaching = straight_line_assignments(source, member.body, activity)

    def held_at(node: BodyNode) -> ReachingAssignments | None:
        # Upstream's Map keyed by node: the port keys node maps by id().
        return reaching.get(id(node))

    # Variant locals given an error value, by name: the text of the value the walk
    # saw reach, which must still reach where it is read (issue #607). A cell's
    # error is taken when it is read.
    error_locals: dict[str, _Given] = {}
    # Cells a formula of the procedure made an error, by address, in one straight
    # run of statements.
    error_cells: dict[str, str] = {}

    # A local the walk still sees holding what was given it in error.
    def given_error(toks: Sequence[VbaToken], i: int, held: ReachingAssignments | None) -> _Found | None:
        name = (
            None
            if _raw_at(toks, i - 1) == "." or _raw_at(toks, i + 1) == "." or _raw_at(toks, i + 1) == "("
            else token_name(_at(toks, i))
        )
        lower = name.lower() if name else None
        given = error_locals.get(lower) if lower else None
        now: str | None = None
        if lower and held is not None:
            value = _held_value(held, lower)
            now = " ".join(tok.raw_text for tok in _without_comments(value)) if value is not None else None
        if given is not None and now == given.text:
            return _Found(i + 1, f"'{toks[i].raw_text}' {given.what}")
        return None

    def check(
        span: Span,
        toks: Sequence[VbaToken],
        held: ReachingAssignments | None,
        header: bool,
        select_has_literal_case: bool = False,
    ) -> None:
        def operand(i: int) -> _Found | None:
            found = _error_operand(toks, i, held, variant_local, error_cells, option_base)
            return found if found is not None else given_error(toks, i, held)

        head = token_text(_at(toks, 0))
        # The `=` of an assignment is no comparison.
        assign_at = -1 if header else _assignment_equals(toks)
        # An operator reported through its left operand is not reported again
        # through its right one.
        reported_operator = -1
        # The target of `v = ...` is written, not read.
        i = assign_at + 1
        while i < len(toks):
            found = operand(i)
            if found is None:
                i += 1
                continue
            if i - 1 == reported_operator:
                i = found.end
                continue
            end, what = found.end, found.what
            at = Span(span.start + toks[i].start, span.start + toks[end - 1].end)
            before = _at(toks, i - 1)
            after = _at(toks, end)
            before_text = token_text(before)
            binary_before = (
                before_text in _BINARY_OPERATORS
                and i - 1 != assign_at
                and i - 1 > 0
                and (_raw_at(toks, i - 2) or "") not in ("(", ",", "=")
                and token_text(_at(toks, i - 2)) not in _BINARY_OPERATORS
            )
            binary_after = token_text(after) in _BINARY_OPERATORS
            use: str | None = None
            error = "13"
            # Two error values compare: `v = CVErr(2042)` runs (issue #310,
            # measured), so a comparison is judged against a literal only.
            operator = after if binary_after else before
            other = _at(toks, end + 1) if binary_after else _at(toks, i - 2)
            if binary_after:
                beyond = _at(toks, end + 2)
                other_whole = beyond is None or (
                    beyond.raw_text not in ("(", ".", "!") and token_text(beyond) not in _BINARY_OPERATORS
                )
            else:
                other_whole = token_text(_at(toks, i - 3)) not in _BINARY_OPERATORS and _raw_at(toks, i - 3) != "."
            other_held: list[VbaToken] | None = None
            if other is not None and token_name(other) and variant_local(token_text(other)) and held is not None:
                value = _held_value(held, token_text(other))
                other_held = _without_comments(value) if value is not None else None
            other_value = (
                _is_literal(other)
                or token_text(other) == "empty"
                or (
                    other_held is not None
                    and len(other_held) == 1
                    and (_is_literal(other_held[0]) or token_text(other_held[0]) == "empty")
                )
            )
            compared_with_value = (
                (operator.raw_text if operator is not None else "") not in _COMPARISONS
                or (other_value and other_whole)
            )
            if (binary_before or binary_after) and compared_with_value and operator is not None:
                use = f"an operand of {operator.raw_text}"
            elif before_text == "not":
                use = "the operand of Not"
            elif before is not None and before.raw_text == "-" and (
                i - 1 == 0
                or i - 1 == assign_at + 1
                or (_raw_at(toks, i - 2) or "") in ("(", ",", "=")
                or token_text(_at(toks, i - 2)) in _BINARY_OPERATORS
            ):
                use = "the operand of -"
            elif (
                header
                and head == "do"
                and token_text(_at(toks, 1)) in ("while", "until")
                and i == 2
                and end == len(toks)
            ):
                use = "the Do condition"
            elif (
                header
                and head == "loop"
                and token_text(_at(toks, 1)) in ("while", "until")
                and i == 2
                and end == len(toks)
            ):
                use = "the Loop condition"
            elif (header and head == "for" and before_text in ("to", "step")) or (
                header
                and head == "for"
                and before is not None
                and before.raw_text == "="
                and token_text(after) == "to"
            ):
                use = "a bound of the For"
            elif (
                before is not None
                and before.raw_text in ("(", ",")
                and after is not None
                and after.raw_text in (")", ",")
            ):
                call = _call_around(toks, i)
                if call is not None:
                    call_name, call_slot = call
                    name = token_text(toks[call_name])
                    qualified = _raw_at(toks, call_name - 1) == "."
                    if not qualified and name in _VALUE_FUNCTIONS:
                        use = f"an argument of {toks[call_name].raw_text}"
                        error = "6" if name == "cbyte" else "13"
                    elif not qualified and name in _FIRST_ARGUMENT_FUNCTIONS and call_slot == 0:
                        use = f"the first argument of {toks[call_name].raw_text}"
                    elif (
                        qualified
                        and token_text(_at(toks, call_name - 2)) == "worksheetfunction"
                        and name == "sum"
                        and _raw_at(toks, call_name - 3) != "."
                    ):
                        use = "an argument of WorksheetFunction.Sum"
                        error = "1004"
                    elif not qualified and name == "array" and _array_given_to_join(toks, call_name):
                        use = "an element of the array Join is given"
                    elif not qualified and array_local(name):
                        use = f"an index of '{toks[call_name].raw_text}'"
                    elif not qualified:
                        param = by_val_scalar_param(name, call_slot)
                        use = f"the argument of {toks[call_name].raw_text}'s ByVal {param}" if param else None
            elif header and head in ("if", "elseif") and i == 1 and token_text(after) == "then":
                use = "the If condition"
            elif (
                header
                and head == "select"
                and token_text(_at(toks, 1)) == "case"
                and i == 2
                and end == len(toks)
                and select_has_literal_case
            ):
                use = "the Select Case subject, compared with each Case"
            elif (
                i == assign_at + 1
                and end == len(toks)
                and (assign_at == 1 or (assign_at == 2 and head == "let"))
            ):
                target = toks[assign_at - 1]
                type_name = scalar_local(token_text(target))
                use = f"Let into '{target.raw_text}', a {type_name}" if type_name else None
            if use:
                if error == "6":
                    tail = "This will raise Run-time error '6': Overflow."
                elif error == "1004":
                    tail = (
                        "This will raise Run-time error '1004': Unable to get the Sum property of the "
                        "WorksheetFunction class."
                    )
                else:
                    tail = _MESSAGE_TAIL
                byte_note = ", whose 2042 or so does not fit a Byte" if error == "6" else ""
                push(
                    "variantValueMisuse",
                    f"{what}, which is no number or text, and here it is {use}{byte_note}. {tail}",
                    at,
                )
                if binary_after:
                    reported_operator = end
            i = end

    # `v = Range("A1").Value` after the error was written there, `v = a(1)`.
    def note_given(toks: Sequence[VbaToken], held: ReachingAssignments | None) -> None:
        name = token_name(_at(toks, 0))
        lower = name.lower() if name else None
        if not lower or _raw_at(toks, 1) != "=" or not variant_local(lower):
            return
        value = list(toks[2:])
        found = _error_operand(value, 0, held, variant_local, error_cells, option_base)
        if found is None:
            found = _array_literal_element(value, option_base)
        if found is not None and found.end == len(value):
            shown = _GIVEN_WHAT.sub(r"\1 ", found.what, count=1)
            error_locals[lower] = _Given(
                " ".join(tok.raw_text for tok in value), f"holds what {shown}, an error value,"
            )
        else:
            error_locals.pop(lower, None)

    # Upstream's recursive visit, on an explicit stack: a block clears the known
    # cells as it is entered and again once its body is done.
    stack: list[Iterator[BodyNode]] = [iter(member.body)]
    while stack:
        entered = False
        for node in stack[-1]:
            if activity is not None and activity.is_inactive(node.span):
                continue
            if not is_leaf_statement(node):
                # A Dim runs nothing; a block may run anything.
                body = getattr(node, "body", None)
                if isinstance(body, list):
                    error_cells = {}
                    header_span = block_header_line_span(source, node.span)
                    # A Case of a number or text compares the subject with a value.
                    literal_case = isinstance(node, SelectBlockNode) and any(
                        is_leaf_statement(item) and _literal_case_item(statement_tokens(source, item.span))
                        for item in body
                    )
                    check(header_span, statement_tokens(source, header_span), held_at(node), True, literal_case)
                    stack.append(iter(body))
                    entered = True
                    break
                continue
            if statement_label_declaration(source, node.span) is not None:
                error_cells = {}
            # A single-line If is judged as its condition, up to Then, and then
            # each branch as a statement of its own.
            branches = isinstance(node, StatementNode) and node.single_line_if_branches is not None
            for span in statement_and_branch_spans(node):
                toks = statement_tokens(source, span)
                if branches and span is node.span:
                    then = next((k for k, tok in enumerate(toks) if token_text(tok) == "then"), -1)
                    check(span, toks[: then + 1], held_at(node), True)
                else:
                    check(span, toks, held_at(node), False)
            own = statement_tokens(source, node.span)
            note_given(own, held_at(node))
            error_cells = _next_error_cells(own, error_cells)
        if entered:
            continue
        stack.pop()
        if stack:
            error_cells = {}


def _held_value(held: ReachingAssignments, lower: str) -> Sequence[VbaToken] | None:
    return held.get(lower)


def _error_operand(
    toks: Sequence[VbaToken],
    i: int,
    held: ReachingAssignments | None,
    variant_local: Callable[[str], bool],
    error_cells: Mapping[str, str],
    option_base: int,
) -> _Found | None:
    """The error value starting at `toks[i]`, with the index after it and how it is
    shown, or None."""
    if _raw_at(toks, i - 1) == ".":
        return None
    word = token_text(_at(toks, i))
    close = match_paren_from(toks, i + 1) if _raw_at(toks, i + 1) == "(" else -1
    if word == "cverr" and close > 0:
        return _Found(close + 1, f"{_joined(toks[i : close + 1])} is an error value")
    formula = (
        _strip_equals(string_literal_value(toks[i + 2].raw_text))
        if word == "evaluate" and close == i + 3 and toks[i + 2].kind is TokenKind.STRING_LITERAL
        else None
    )
    if formula is not None and _ERROR_FORMULA.search(formula) is not None:
        return _Found(close + 1, f"{_joined(toks[i : close + 1])} gives an error value")
    # `Range("A1").Value` after the procedure wrote `=1/0` there, or
    # `ActiveSheet.Range("A1")`, `Cells(1, 1)` (issue #607).
    cell = _cell_at(toks, i)
    if cell is not None and len(error_cells) > 0:
        value_read = (
            cell.end + 2
            if _raw_at(toks, cell.end) == "." and token_text(_at(toks, cell.end + 1)) in ("value", "value2")
            else -1
        )
        read = value_read if value_read > 0 else (cell.end if _raw_at(toks, cell.end) != "." else -1)
        if cell.address in error_cells and read > 0 and _raw_at(toks, read) != "(":
            return _Found(read, f"{_joined(toks[i:read])} holds the error value of {error_cells[cell.address]}")
        return None
    if cell is not None:
        return None
    name = token_name(_at(toks, i))
    lower = name.lower() if name else None
    if not lower or held is None or not variant_local(lower) or _raw_at(toks, i + 1) == ".":
        return None
    raw_value = _held_value(held, lower)
    if raw_value is None:
        return None
    value = _without_comments(raw_value)
    # `a(0)` with `a = Array(CVErr(2007))`.
    if close > 0:
        index = (
            js_number(toks[i + 2].raw_text)
            if close == i + 3 and toks[i + 2].kind is TokenKind.INTEGER_LITERAL
            else None
        )
        if (
            index is None
            or token_text(_at(value, 0)) != "array"
            or _raw_at(value, 1) != "("
            or match_paren_from(value, 1) != len(value) - 1
        ):
            return None
        elements = split_top_level_token_groups(value, 2, ",", len(value) - 1) if len(value) > 3 else []
        element = _js_element(elements, index - option_base)
        if element and _is_error_source(element):
            return _Found(
                close + 1,
                f"'{toks[i].raw_text}({js_number_to_string(index)})' holds an error value from "
                f"{_joined(element)} here",
            )
        return None
    if _is_error_source(value):
        return _Found(i + 1, f"'{toks[i].raw_text}' holds an error value from {_joined(value)} here")
    return None


def _assignment_equals(toks: Sequence[VbaToken]) -> int:
    """The index of the `=` that makes the statement an assignment: after an
    optional Let, a chain of names, each maybe indexed, as in `v = `, `v(0) = ` or
    `t.x = `. -1 for any other statement."""
    j = 1 if token_text(_at(toks, 0)) == "let" else 0
    while True:
        if not token_name(_at(toks, j)):
            return -1
        j += 1
        if _raw_at(toks, j) == "(":
            close = match_paren_from(toks, j)
            if close < 0:
                return -1
            j = close + 1
        if _raw_at(toks, j) == "=":
            return j
        if _raw_at(toks, j) != ".":
            return -1
        j += 1


def _cell_at(toks: Sequence[VbaToken], i: int) -> _Cell | None:
    """A cell named by literals at `toks[i]`: `Range("A1")`, `Cells(1, 1)`, each
    maybe after `ActiveSheet.`, with its A1 address in lower case and the index
    after it."""
    at = i
    if token_text(_at(toks, at)) == "activesheet" and _raw_at(toks, at + 1) == ".":
        at += 2
    elif _raw_at(toks, at - 1) == ".":
        return None
    word = token_text(_at(toks, at))
    third = _at(toks, at + 2)
    if (
        word == "range"
        and _raw_at(toks, at + 1) == "("
        and third is not None
        and third.kind is TokenKind.STRING_LITERAL
        and _raw_at(toks, at + 3) == ")"
    ):
        return _Cell(string_literal_value(third.raw_text).replace("$", "").lower(), at + 4)
    fifth = _at(toks, at + 4)
    if (
        word == "cells"
        and _raw_at(toks, at + 1) == "("
        and third is not None
        and third.kind is TokenKind.INTEGER_LITERAL
        and _raw_at(toks, at + 3) == ","
        and fifth is not None
        and fifth.kind is TokenKind.INTEGER_LITERAL
        and _raw_at(toks, at + 5) == ")"
    ):
        row = js_number(third.raw_text)
        column = js_number(fifth.raw_text)
        letters = ""
        while column > 0:
            letters = chr(97 + int(math.fmod(column - 1, 26))) + letters
            column = math.floor((column - 1) / 26)
        if letters and row > 0:
            return _Cell(f"{letters}{js_number_to_string(row)}", at + 6)
        return None
    return None


def _array_literal_element(value: Sequence[VbaToken], option_base: int) -> _Found | None:
    """`Array(1, CVErr(2007))(1)`: an element of an array literal that is an error
    value."""
    if token_text(_at(value, 0)) != "array" or _raw_at(value, 1) != "(":
        return None
    close = match_paren_from(value, 1)
    index_tok = _at(value, close + 2)
    if (
        close < 0
        or _raw_at(value, close + 1) != "("
        or index_tok is None
        or index_tok.kind is not TokenKind.INTEGER_LITERAL
        or _raw_at(value, close + 3) != ")"
    ):
        return None
    elements = split_top_level_token_groups(value, 2, ",", close) if close > 2 else []
    element = _js_element(elements, js_number(index_tok.raw_text) - option_base)
    if element and _is_error_source(element):
        return _Found(close + 4, f"an element {_joined(element)} of an array")
    return None


def _call_around(toks: Sequence[VbaToken], i: int) -> tuple[int, int] | None:
    """The call whose argument list holds `toks[i]`: its name's index and the
    argument's slot."""
    depth = 0
    slot = 0
    for k in range(i - 1, -1, -1):
        raw = toks[k].raw_text
        if raw == ")":
            depth += 1
        elif raw == "(":
            if depth == 0:
                return (k - 1, slot) if token_name(_at(toks, k - 1)) else None
            depth -= 1
        elif raw == "," and depth == 0:
            slot += 1
    return None


def _array_given_to_join(toks: Sequence[VbaToken], at: int) -> bool:
    """Whether the Array call at `at` is the first argument of Join."""
    return (
        _raw_at(toks, at - 1) == "("
        and token_text(_at(toks, at - 2)) == "join"
        and _raw_at(toks, at - 3) != "."
    )


def _is_error_source(value: Sequence[VbaToken]) -> bool:
    """`CVErr(...)`, or `Evaluate` of a literal formula that is an error, whole."""
    word = token_text(_at(value, 0))
    if _raw_at(value, 1) != "(" or match_paren_from(value, 1) != len(value) - 1:
        return False
    if word == "cverr":
        return True
    return (
        word == "evaluate"
        and len(value) == 4
        and value[2].kind is TokenKind.STRING_LITERAL
        and _ERROR_FORMULA.search(_strip_equals(string_literal_value(value[2].raw_text))) is not None
    )


def _next_error_cells(toks: Sequence[VbaToken], cells: Mapping[str, str]) -> dict[str, str]:
    """The cells known to hold an error after a statement: `Range("A1").Formula =
    "=1/0"` adds A1; anything else that names Range, Cells or a sheet, or may run
    other code, ends what is known."""
    # `Range("A1").Formula = "=1/0"`; Value takes a formula too (issue #607).
    cell = _cell_at(toks, 0)
    if (
        cell is not None
        and _raw_at(toks, cell.end) == "."
        and token_text(_at(toks, cell.end + 1)) in ("formula", "value")
        and _raw_at(toks, cell.end + 2) == "="
        and len(toks) == cell.end + 4
        and toks[cell.end + 3].kind is TokenKind.STRING_LITERAL
    ):
        formula = string_literal_value(toks[cell.end + 3].raw_text)
        following = dict(cells)
        if (
            formula.startswith("=")
            and _ERROR_FORMULA.search(formula[1:]) is not None
            and _CELL_ADDRESS.fullmatch(cell.address) is not None
        ):
            following[cell.address] = formula
        else:
            following.clear()
        return following
    # A read leaves the cells as they are; a plain Let to a local, or into the
    # Function's result, does too.
    rest = toks[2:]
    plain_let = (
        token_name(_at(toks, 0)) is not None
        and _raw_at(toks, 1) == "="
        and not any(token_text(tok) in _CELL_CHANGERS for tok in rest)
    )
    calls_other = any(
        tok.kind is TokenKind.IDENTIFIER
        and _raw_at(toks, k + 3) == "("
        and token_text(tok) not in _CELL_KEEPING_CALLS
        for k, tok in enumerate(rest)
    )
    return dict(cells) if plain_let and len(cells) > 0 and not calls_other else {}


def _strip_equals(text: str) -> str:
    """`.replace(/^=/, '')`: one leading '='."""
    return text[1:] if text.startswith("=") else text


def _js_element(items: Sequence[list[VbaToken]], index: float) -> list[VbaToken] | None:
    """`items[index]` as JavaScript reads it: undefined for a negative, fractional or
    out-of-range index."""
    if not math.isfinite(index) or index < 0 or not float(index).is_integer() or index >= len(items):
        return None
    return items[int(index)]


def _joined(toks: Sequence[VbaToken]) -> str:
    return "".join(tok.raw_text for tok in toks)


def _without_comments(toks: Sequence[VbaToken]) -> list[VbaToken]:
    return [tok for tok in toks if tok.kind is not TokenKind.COMMENT]


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None
