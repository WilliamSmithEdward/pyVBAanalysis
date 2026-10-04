"""Rule family: Access SQL literals and DAO recordsets whose failure the code
itself shows (XLIDE issue #312).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/accessData.ts.

Measured in Access 16.0 (2026-10-02), each case creating its own table; each
compiles and raises every time.

runtime-argument-value, on SQL in a string literal:
 - Execute of a SELECT raises 3065, RunSQL of one 2342; a SELECT ... INTO
   makes a table and runs, and a TRANSFORM raises 3065 too (XLIDE issue #611).
 - Execute of "", or of text whose first word is no SQL verb but which reads
   as SQL (`DELET FROM T1`), raises 3078: Execute takes it for the name of a
   query, and none has that name. OpenRecordset does the same.
 - A quote left open raises 3075 through Execute or OpenRecordset, and 2342
   through RunSQL. A double-quoted string is a string too, and a ' inside it
   is text. An INSERT with a parenthesis left open, or with neither VALUES nor
   SELECT, raises 3134; any other SQL with one left open raises 3075.
 - CreateQueryDef of SQL whose first word is no SQL verb raises 3129.
 - DLookup whose criteria end in a comparison with nothing after raises 2342.
   The other domain functions raise 3075 there, and every one does for
   criteria ending in AND or OR, or leaving a quote open (#611).

host-argument-out-of-range, on a DAO.Recordset local followed in a straight
line from `Set rs = ....OpenRecordset(...)`:
 - writing a field, or Update, with no Edit or AddNew since raises 3020; a
   Move ends an Edit, and `!Nm = x` inside `With rs` is rs's;
 - Edit, AddNew or Delete on a snapshot (dbOpenSnapshot), and Edit or AddNew
   on a forward-only one, raise 3251; on one opened dbReadOnly, 3027;
 - any use after rs.Close raises 3420, through another name for it too.
"""

from __future__ import annotations

import dataclasses
import re
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import statement_label_declaration
from ...js_compat import JS_WHITESPACE, js_trim
from ...lexer.token_helpers import split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    WithBlockNode,
    is_leaf_statement,
)
from ...symbols.symbol_model import ModuleSymbols
from ...types.type_inference import type_environment_for
from ...types.type_names import normalize_type
from ..call_extraction import string_literal_value
from ..context import PushFn
from ..dataflow import BlockEnteringState, walk_entering_blocks
from ..walker import (
    active_module_members,
    block_header_line_span,
    for_each_statement,
    match_paren_from,
    set_assignment_target,
    statement_and_branch_spans,
    statement_tokens,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .shared import names_in

# The words an Access SQL statement starts with.
_SQL_VERBS: frozenset[str] = frozenset(
    {"select", "insert", "update", "delete", "create", "alter", "drop", "transform", "parameters", "procedure"}
)

# JavaScript's `\s` (its own whitespace set), `\b` and `\w` (ASCII), `$` (no
# match before a trailing newline), and `/i` (no letter outside ASCII folds onto
# an ASCII one).
_WS = JS_WHITESPACE
_FLAGS = re.IGNORECASE | re.ASCII

_COMPARISON_AT_END = re.compile(r"(?:=|<>|<|>|\blike|\bin)[" + _WS + r"]*\Z", _FLAGS)
_LOGIC_AT_END = re.compile(r"\b(?:and|or|not)[" + _WS + r"]*\Z", _FLAGS)
_LAST_WORD = re.compile(r"(\w+)[" + _WS + r"]*\Z", re.ASCII)
_FIRST_WORD = re.compile(r"^[" + _WS + r"(]*([A-Za-z]+)")
_READS_AS_SQL = re.compile(r"\b(?:from|into|set|values|where)\b", _FLAGS)
_INSERT_SHAPE = re.compile(
    r"^[" + _WS + r"]*insert[" + _WS + r"]+into[" + _WS + r"]+(?:\[[^\]]*\]|[^" + _WS + r"(]+)["
    + _WS
    + r"]*(?:\([^)]*\)["
    + _WS
    + r"]*)?(?:values|select)\b",
    _FLAGS,
)
_INTO = re.compile(r"\binto\b", _FLAGS)

# Access's domain aggregate functions, which read their criteria as SQL.
_DOMAIN_FUNCTIONS: frozenset[str] = frozenset(
    {"dlookup", "dcount", "dsum", "davg", "dmin", "dmax", "dfirst", "dlast", "dstdev", "dstdevp", "dvar", "dvarp"}
)


def _at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """`toks[index]`, or None out of range the way a JavaScript index reads undefined."""
    return toks[index] if 0 <= index < len(toks) else None


def _raw(toks: Sequence[VbaToken], index: int) -> str | None:
    tok = _at(toks, index)
    return tok.raw_text if tok is not None else None


def check_access_data(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    host: str | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    if host != "Access":
        return
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        env = type_environment_for(symbols, member)

        def is_database(lower: str, env: Mapping[str, str] = env) -> bool:
            return (normalize_type(env.get(lower)) or "") in ("dao.database", "database")

        def is_recordset(lower: str, env: Mapping[str, str] = env) -> bool:
            return (normalize_type(env.get(lower)) or "") in ("dao.recordset", "recordset")

        def visit(stmt: LeafStatementNode, is_database: Callable[[str], bool] = is_database) -> None:
            for span in statement_and_branch_spans(stmt):
                _check_sql_literals(span, statement_tokens(source, span), is_database, push)

        for_each_statement(member.body, visit, activity)
        _check_recordsets(source, member.body, is_database, is_recordset, activity, push)


def _database_at(toks: Sequence[VbaToken], end: int, is_database: Callable[[str], bool]) -> bool:
    """`CurrentDb`, `Application.CurrentDb`, or a local As DAO.Database, ending at `toks[end]`."""
    word = token_text(_at(toks, end))
    qualified = _raw(toks, end - 1) == "."
    if word == "currentdb" or word == "codedb":
        return not qualified or token_text(_at(toks, end - 2)) == "application"
    if word == ")" and token_text(_at(toks, end - 1)) == "(" and token_text(_at(toks, end - 2)) in ("currentdb", "codedb"):
        return _raw(toks, end - 3) != "." or token_text(_at(toks, end - 4)) == "application"
    return not qualified and token_name(_at(toks, end)) is not None and is_database(word)


def _call_arguments(toks: Sequence[VbaToken], i: int) -> list[list[VbaToken]] | None:
    """The arguments of the call whose name is at `i`, written with parentheses or as a statement."""
    if _raw(toks, i + 1) == "(":
        close = match_paren_from(toks, i + 1)
        if close > i + 2:
            return split_top_level_token_groups(toks, i + 2, ",", close)
        return [] if close == i + 2 else None
    return split_top_level_token_groups(toks, i + 1, ",", len(toks)) if i + 1 < len(toks) else []


def _check_sql_literals(
    span: Span, toks: Sequence[VbaToken], is_database: Callable[[str], bool], push: PushFn
) -> None:
    for i in range(2, len(toks)):
        word = token_text(toks[i])
        on_database = toks[i - 1].raw_text == "." and _database_at(toks, i - 2, is_database)
        do_cmd = (
            toks[i - 1].raw_text == "."
            and token_text(toks[i - 2]) == "docmd"
            and (_raw(toks, i - 3) != "." or token_text(_at(toks, i - 4)) == "application")
        )
        kind = (
            word
            if on_database and (word == "execute" or word == "openrecordset" or word == "createquerydef")
            else "runsql"
            if do_cmd and word == "runsql"
            else None
        )
        if kind is None:
            continue
        args = _call_arguments(toks, i)
        position = 1 if kind == "createquerydef" else 0
        arg = args[position] if args is not None and position < len(args) else None
        if arg is None or len(arg) != 1 or arg[0].kind is not TokenKind.STRING_LITERAL:
            continue
        problem = _sql_problem(kind, string_literal_value(arg[0].raw_text))
        if problem:
            push("runtimeArgumentValue", f"{problem}.", Span(span.start + arg[0].start, span.start + arg[0].end))
    # `DLookup("Nm", "T1", "ID = ")`: criteria that end in a comparison, in AND
    # or OR, or leave a quote open.
    for i in range(len(toks)):
        name = token_text(toks[i])
        if name not in _DOMAIN_FUNCTIONS or _raw(toks, i - 1) == "." or _raw(toks, i + 1) != "(":
            continue
        call_args = _call_arguments(toks, i)
        criteria = call_args[2] if call_args is not None and len(call_args) > 2 else None
        if criteria is None or len(criteria) != 1 or criteria[0].kind is not TokenKind.STRING_LITERAL:
            continue
        text = string_literal_value(criteria[0].raw_text)
        at = Span(span.start + criteria[0].start, span.start + criteria[0].end)
        shown = toks[i].raw_text
        if _COMPARISON_AT_END.search(text) is not None:
            push(
                "runtimeArgumentValue",
                "The criteria of DLookup end in a comparison with nothing to compare with. This will raise "
                "Run-time error '2342': A RunSQL action requires an argument consisting of an SQL statement."
                if name == "dlookup"
                else f"The criteria of {shown} end in a comparison with nothing to compare with. This will "
                "raise Run-time error '3075': Syntax error (missing operator) in query expression.",
                at,
            )
        elif _LOGIC_AT_END.search(text) is not None:
            last = _LAST_WORD.search(text)
            assert last is not None
            push(
                "runtimeArgumentValue",
                f"The criteria of {shown} end in '{last.group(1)}' with nothing after it. This will raise "
                "Run-time error '3075': Syntax error (missing operator) in query expression.",
                at,
            )
        elif _open_quote(text):
            push(
                "runtimeArgumentValue",
                f"The criteria of {shown} leave a quote open. This will raise Run-time error '3075': Syntax "
                "error in string in query expression.",
                at,
            )


def _sql_problem(kind: str, sql: str) -> str | None:
    """Why Access refuses this SQL text, with the error it raises, or None."""
    first_match = _FIRST_WORD.search(sql)
    first = first_match.group(1).lower() if first_match is not None else None
    reads_as_sql = _READS_AS_SQL.search(sql) is not None
    quotes = _open_quote(sql)
    if kind == "runsql":
        if first == "select":
            return (
                f"RunSQL runs an action query, and \"{sql}\" is a SELECT. This will raise Run-time error "
                "'2342': A RunSQL action requires an argument consisting of an SQL statement"
            )
        return (
            "The SQL leaves a quote open. This will raise Run-time error '2342': A RunSQL action requires an "
            "argument consisting of an SQL statement"
            if quotes
            else None
        )
    if kind == "createquerydef":
        return (
            f"\"{first}\" starts no SQL statement. This will raise Run-time error '3129': Invalid SQL "
            "statement; expected 'DELETE', 'INSERT', 'PROCEDURE', 'SELECT', or 'UPDATE'"
            if first and first not in _SQL_VERBS
            else None
        )
    if js_trim(sql) == "" and kind == "execute":
        return (
            "Execute has no SQL to run, and no query is named \"\". This will raise Run-time error '3078': "
            "The Microsoft Access database engine cannot find the input table or query"
        )
    if first and first not in _SQL_VERBS and reads_as_sql:
        return (
            f"\"{first}\" starts no SQL statement, so the text is taken for the name of a table or query, and "
            "none has that name. This will raise Run-time error '3078': The Microsoft Access database engine "
            "cannot find the input table or query"
        )
    # The SQL is parsed before Execute asks what kind it is: a SELECT with a
    # parenthesis left open raises 3075, not 3065 (XLIDE issue #611).
    if quotes:
        return "The SQL leaves a quote open. This will raise Run-time error '3075': Syntax error in string in query expression"
    if kind == "execute" and first == "insert" and (_open_parenthesis(sql) or _INSERT_SHAPE.search(sql) is None):
        what = "leaves a parenthesis open" if _open_parenthesis(sql) else "has neither VALUES nor a SELECT"
        return f"The INSERT {what}. This will raise Run-time error '3134': Syntax error in INSERT INTO statement"
    if _open_parenthesis(sql):
        return (
            "The SQL leaves a parenthesis open. This will raise Run-time error '3075': Missing ), ], or Item "
            "in query expression"
        )
    # SELECT ... INTO makes a table, an action query (XLIDE issue #611).
    if kind == "execute" and (
        (first == "select" and _INTO.search(_outside_strings(sql)) is None) or first == "transform"
    ):
        shown = "SELECT" if first == "select" else "TRANSFORM"
        return (
            f"Execute runs an action query, and this is a {shown}. This will raise Run-time error '3065': "
            "Cannot execute a select query"
        )
    return None


def _scan_strings(sql: str) -> tuple[str, bool]:
    """The SQL with its strings blanked, and whether one is left open. A string
    is single- or double-quoted, its quote doubled inside it (XLIDE issue #611:
    `"it's"` is one string)."""
    outside: list[str] = []
    quote: str | None = None
    i = 0
    while i < len(sql):
        ch = sql[i]
        if quote is not None:
            if ch == quote and i + 1 < len(sql) and sql[i + 1] == quote:
                i += 1
            elif ch == quote:
                quote = None
            outside.append(" ")
            i += 1
            continue
        if ch == "'" or ch == '"':
            quote = ch
            outside.append(" ")
            i += 1
            continue
        outside.append(ch)
        i += 1
    return "".join(outside), quote is not None


def _open_quote(sql: str) -> bool:
    """Whether a string of the SQL is left open."""
    return _scan_strings(sql)[1]


def _outside_strings(sql: str) -> str:
    """The SQL outside its strings."""
    return _scan_strings(sql)[0]


def _open_parenthesis(sql: str) -> bool:
    """Whether a parenthesis outside strings is left open."""
    depth = 0
    for ch in _outside_strings(sql):
        depth += 1 if ch == "(" else -1 if ch == ")" else 0
    return depth > 0


@dataclass(slots=True)
class _RecordsetState:
    """What a straight line of statements knows of a recordset local."""

    snapshot: bool
    # dbOpenForwardOnly: Edit and AddNew raise 3251.
    forward_only: bool
    # Opened dbReadOnly: Edit and AddNew raise 3027.
    read_only: bool
    editing: bool
    closed: bool


def _copy_states(states: Mapping[str, _RecordsetState]) -> dict[str, _RecordsetState]:
    """A copy of the states, names that share one recordset sharing one copy."""
    copies: dict[int, _RecordsetState] = {}
    out: dict[str, _RecordsetState] = {}
    for lower, held in states.items():
        copy = copies.get(id(held))
        if copy is None:
            copy = dataclasses.replace(held)
            copies[id(held)] = copy
        out[lower] = copy
    return out


# The methods that move a recordset's current record, ending an Edit (XLIDE issue #611).
_MOVES: frozenset[str] = frozenset(
    {
        "movefirst", "movelast", "movenext", "moveprevious", "move", "findfirst", "findlast", "findnext",
        "findprevious", "seek", "requery",
    }
)


def _ordered_names(source: str, span: Span) -> list[str]:
    """namesIn in the order upstream's Set iterates it: first mention first."""
    out: dict[str, None] = {}
    for tok in statement_tokens_after_leading_label(source, span):
        name = token_name(tok)
        if name:
            out[name.lower()] = None
    return list(out)


def _check_recordsets(
    source: str,
    body: Sequence[BodyNode],
    is_database: Callable[[str], bool],
    is_recordset: Callable[[str], bool],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    state: dict[str, _RecordsetState] = {}
    with_subjects: list[str | None] = []

    def forget(names: Iterable[str]) -> None:
        for lower in names:
            state.pop(lower, None)

    def visit(node: BodyNode) -> None:
        if not is_leaf_statement(node):
            return
        if statement_label_declaration(source, node.span):
            state.clear()
        if isinstance(node, StatementNode) and node.single_line_if_branches:
            forget(names_in(source, node.span))
            return
        own = statement_tokens_after_leading_label(source, node.span)
        # Inside `With rs`, `!Nm = x` and `.Edit` are rs's (XLIDE issue #611).
        subject = with_subjects[-1] if with_subjects else None
        toks: list[VbaToken] = (
            [
                dataclasses.replace(own[0], kind=TokenKind.IDENTIFIER, raw_text=subject, end=own[0].start),
                *own,
            ]
            if subject and (_raw(own, 0) == "!" or _raw(own, 0) == ".")
            else list(own)
        )

        def at(first: int, last: int) -> Span:
            return Span(node.span.start + toks[first].start, node.span.start + toks[last].end)

        set_target = set_assignment_target(source, node.span)
        if set_target is not None:
            lower = set_target[0].lower()
            value = [tok for tok in set_target[2] if tok.kind is not TokenKind.COMMENT]
            # `Set r2 = rs` is another name for the same recordset.
            alias_name = token_name(value[0]) if len(value) == 1 else None
            alias = alias_name.lower() if alias_name is not None else None
            shared = state.get(alias) if alias else None
            forget(names_in(source, node.span))
            open_at = next(
                (
                    k
                    for k, tok in enumerate(value)
                    if token_text(tok) == "openrecordset"
                    and _raw(value, k - 1) == "."
                    and _database_at(value, k - 2, is_database)
                ),
                -1,
            )
            if (
                is_recordset(lower)
                and open_at > 0
                and _raw(value, open_at + 1) == "("
                and match_paren_from(value, open_at + 1) == len(value) - 1
            ):
                args = _call_arguments(value, open_at)
                kind = (
                    "".join(token_text(tok) for tok in args[1]) if args is not None and len(args) > 1 else None
                )
                options = (
                    "".join(token_text(tok) for tok in args[2]) if args is not None and len(args) > 2 else None
                )
                state[lower] = _RecordsetState(
                    snapshot=kind == "dbopensnapshot" or kind == "4",
                    forward_only=kind == "dbopenforwardonly" or kind == "8",
                    read_only=options == "dbreadonly" or options == "4",
                    editing=False,
                    closed=False,
                )
            if is_recordset(lower) and shared is not None:
                assert alias is not None
                state[lower] = shared
                state[alias] = shared
            return
        first_name = token_name(_at(toks, 0))
        lower_name = first_name.lower() if first_name is not None else None
        # A recordset read past the statement's head: `Main = rs!Nm`, `x = rs.EOF`.
        for name in _ordered_names(source, node.span):
            other = None if name == lower_name else state.get(name)
            if other is None:
                continue
            uses = [k for k, tok in enumerate(toks) if token_text(tok) == name and _raw(toks, k - 1) != "."]
            if any((_raw(toks, k + 1) or "") not in (".", "!", "(") for k in uses):
                state.pop(name, None)  # passed whole, which may change it
            elif other.closed and len(uses) > 0:
                push(
                    "hostArgumentOutOfRange",
                    f"'{toks[uses[0]].raw_text}' was closed above, and a closed recordset has no members. "
                    "This will raise Run-time error '3420': Object invalid or no longer set.",
                    at(uses[0], uses[0]),
                )
                state.pop(name, None)
        held = state.get(lower_name) if lower_name else None
        if held is None or lower_name is None:
            forget([name for name in names_in(source, node.span) if name not in state])
            return
        shown = toks[0].raw_text
        # Any use after Close.
        if held.closed:
            push(
                "hostArgumentOutOfRange",
                f"'{shown}' was closed above, and a closed recordset has no members. This will raise "
                "Run-time error '3420': Object invalid or no longer set.",
                at(0, 0),
            )
            state.pop(lower_name, None)
            return
        member = token_text(_at(toks, 2)) if _raw(toks, 1) == "." else None
        eq = next((k for k, tok in enumerate(toks) if tok.raw_text == "="), -1)
        # `rs!Nm = x`, `rs("Nm") = x`, `rs.Fields("Nm") = x`, `rs.Fields("Nm").Value = x`.
        field_write = eq > 1 and (_raw(toks, 1) == "!" or _raw(toks, 1) == "(" or member == "fields")
        if field_write or (member == "update" and len(toks) == 3):
            if not held.editing:
                push(
                    "hostArgumentOutOfRange",
                    f"'{shown}' is not being edited: no Edit or AddNew came since it was opened or last "
                    "updated. This will raise Run-time error '3020': Update or CancelUpdate without AddNew "
                    "or Edit.",
                    at(0, eq - 1 if eq > 1 else 2),
                )
                state.pop(lower_name, None)
                return
            if member == "update":
                held.editing = False
            return
        if (member == "edit" or member == "addnew" or member == "delete") and len(toks) == 3:
            if held.snapshot or (held.forward_only and member != "delete"):
                what = "a snapshot" if held.snapshot else "forward-only"
                push(
                    "hostArgumentOutOfRange",
                    f"'{shown}' is {what}, which cannot be changed. This will raise Run-time error '3251': "
                    "Operation is not supported for this type of object.",
                    at(2, 2),
                )
                state.pop(lower_name, None)
                return
            if held.read_only and member != "delete":
                push(
                    "hostArgumentOutOfRange",
                    f"'{shown}' was opened dbReadOnly, which cannot be changed. This will raise Run-time error "
                    "'3027': Cannot update. Database or object is read-only.",
                    at(2, 2),
                )
                state.pop(lower_name, None)
                return
            if member != "delete":
                held.editing = True
            else:
                # Delete moves nothing, and what it leaves is not followed.
                state.pop(lower_name, None)
            return
        if member is not None and member in _MOVES:
            held.editing = False
            return
        if member == "close" and len(toks) == 3:
            held.closed = True
            return
        if member == "cancelupdate" and len(toks) == 3:
            held.editing = False
            return
        # A read leaves it as it is; anything else may change it.
        if not (1 < eq < len(toks)) or any(token_text(tok) == lower_name for tok in toks[:eq]):
            state.pop(lower_name, None)

    # One copy per recordset, so two names for it stay one.
    def snapshot() -> dict[str, _RecordsetState]:
        return _copy_states(state)

    def restore(saved: Mapping[str, _RecordsetState]) -> None:
        state.clear()
        for lower, held in _copy_states(saved).items():
            state[lower] = held

    # `With rs` reads rs, and a body line reaching it by `!` or `.` names it.
    def touches(stmt: LeafStatementNode) -> set[str]:
        toks = statement_tokens_after_leading_label(source, stmt.span)
        if token_text(_at(toks, 0)) == "with" and len(toks) == 2:
            return set()
        names = set(names_in(source, stmt.span))
        subject = with_subjects[-1] if with_subjects else None
        if subject and (_raw(toks, 0) == "!" or _raw(toks, 0) == "."):
            names.add(subject.lower())
        return names

    def enter(node: BodyNode) -> None:
        if not isinstance(node, WithBlockNode):
            return
        header = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
        name = token_name(header[1]) if len(header) == 2 else None
        with_subjects.append(name if name and name.lower() in state else None)

    def exit_block(node: BodyNode) -> None:
        if isinstance(node, WithBlockNode):
            with_subjects.pop()

    walk_entering_blocks(
        source,
        body,
        lambda node: activity is not None and activity.is_inactive(node.span),
        visit,
        BlockEnteringState(
            snapshot=snapshot,
            restore=restore,
            forget=forget,
            touches=touches,
            enter=enter,
            exit=exit_block,
            with_body_runs_through=True,
        ),
    )
