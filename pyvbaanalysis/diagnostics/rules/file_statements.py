"""Rule family: file statements whose failure the code proves (XLIDE issue #123).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/fileStatements.ts.

Each case was measured in Excel 16.0 (build 20326, 2026-09-26): it compiles
and raises every time it runs.

 - file-number-zero (52, Bad file name or number): `As #0`, `LOF(0)`, and
   any literal past the numbers Open takes, 1 to 512: `As #513`,
   `Close #-1` (XLIDE issue #262).
 - file-used-after-close (52): `Close #f` and then `Print #f, ...` on the
   same number with no Open between.
 - file-mode-mismatch (54, Bad file mode): `Print #f`/`Write #f` on a file
   opened For Input; `Input #f`/`Line Input #f` on one opened For Output or
   Append.
 - file-already-open (55, File already open): two Opens As the same number
   with no Close between.
 - file-record-zero (63, Bad record number): `Seek #f, 0`, `Get #f, 0, x`,
   `Put #f, 0, x` - records and Binary positions start at 1. Seek raises
   in any mode, and on a number nothing opened (XLIDE issue #262).
 - file-read-past-end (62, Input past end of file): reading a file this
   procedure created empty - opened For Output, closed with nothing
   written, and opened For Input from the same path - before anything
   checks EOF or LOF (XLIDE issue #262).
 - Open's `Len = 0` raises 5 in every mode, and a Len literal past 32767
   raises 6 (XLIDE issue #262); those report as runtime-argument-value and
   arithmetic-overflow.

A file number is a literal, hex or octal included, or a local that FreeFile
fills once. The rule follows the top-level statements of a procedure in order; a
block between two statements that could touch the number ends what is known
about it, and a number named inside a block is never followed. A label, a GoSub
or a call to a procedure ends everything known, and Reset or a Close with no
number closes every open file (XLIDE issue #146).
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass, replace
from typing import Literal, Union

from ...call.call_context import bare_call_statement_target
from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import parse_vba_integer_literal
from ...flow.procedure_labels import jump_target_label_declaration
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    IfBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    is_leaf_statement,
)
from ..call_extraction import string_literal_value
from ..context import PushFn
from ..dataflow import BlockEnteringState, tracked_locals_named_whole, walk_entering_blocks
from ..opened_file_numbers import (
    OpenedFileNumbers,
    merge_opened_file_numbers,
    opened_file_numbers_in,
)
from ..walker import (
    active_module_members,
    bare_assignment_target,
    block_header_line_span,
    for_each_statement,
    is_inactive_node,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

_FILE_MODES = frozenset({"input", "output", "append", "random", "binary"})


@dataclass(frozen=True, slots=True)
class _OpenFile:
    mode: str
    span: Span
    # The `path:` key of the path it was opened from, when that is a name or a literal.
    path: str | None = None
    # For Output: whether anything was written.
    written: bool | None = None
    # For Input: the file is known to be empty, and nothing has checked EOF or LOF yet.
    empty_unchecked: bool | None = None
    # The `path:` key it was opened from, in any mode, while that name still holds it.
    open_path: str | None = None


# The modes each statement works in; any other raises 54, Bad file mode
# (XLIDE issue #419, measured in Excel 16.0). Input and Line Input read a Binary
# file too; Get and Put need Binary or Random.
_STATEMENT_MODES: Mapping[str, tuple[str, ...]] = {
    "print": ("output", "append"),
    "write": ("output", "append"),
    "input": ("input", "binary"),
    "line input": ("input", "binary"),
    "get": ("binary", "random"),
    "put": ("binary", "random"),
}

# What a path is known to name (XLIDE issue #682): an empty file, a file,
# nothing, an empty folder, or a folder something was made in.
_PathFact = Literal["empty", "file", "absent", "folder", "filled"]

# What is known about each file number key as the statements run, and under
# `path:` keys what each path names.
_FileState = Union[_OpenFile, Literal["closed"], _PathFact]
_FileStates = dict[str, _FileState]


@dataclass(frozen=True, slots=True)
class _OpenLen:
    value: int
    token: VbaToken


@dataclass(frozen=True, slots=True)
class _ParsedOpen:
    mode: str
    key: str | None
    number_index: int
    # The `path:` key of a path that is one name or one string literal.
    path: str | None = None
    len: _OpenLen | None = None


@dataclass(frozen=True, slots=True)
class _SignedLiteral:
    value: int
    first: VbaToken
    last: VbaToken


_FILE_STATEMENTS = frozenset(
    {"print", "write", "input", "line", "get", "put", "seek", "close", "lock", "unlock", "width"}
)

# The highest file number Open takes: As #512 runs and As #513 raises 52
# (XLIDE issue #262).
MAX_FILE_NUMBER = 512

# The functions that read a file's state without reading from it.
_FILE_STATE_FUNCTIONS = frozenset({"lof", "eof", "loc", "fileattr", "seek"})

# VBA functions that only read a name passed to them.
_READ_ONLY_INTRINSICS = frozenset({"len", "lenb", "dir", "filelen", "filedatetime", "getattr"})

_DIGITS_RE = re.compile(r"\d+", re.ASCII)


def check_file_statements(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    project_opened: OpenedFileNumbers | None = None,
) -> None:
    if project_opened is not None:
        _check_unopened_numbers(
            source,
            mod,
            activity,
            push,
            merge_opened_file_numbers([project_opened, opened_file_numbers_in(source)]),
        )
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        _check_procedure(source, member, activity, push)


def _check_procedure(
    source: str,
    member: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    states: _FileStates = {}
    # Under On Error Resume Next a statement that fails goes on to the
    # next: nothing it would raise is reported (XLIDE issue #682).
    resume_next = False

    # Blocks are entered with the state they start with; a block may open,
    # close or reopen anything it names (XLIDE issue #237).
    def visit(node: BodyNode) -> None:
        nonlocal resume_next
        if not is_leaf_statement(node):
            return  # a Dim inside the body declares, and runs nothing
        toks = statement_tokens_after_leading_label(source, node.span)
        if len(toks) == 0:
            return
        if token_text(toks[0]) == "on" and token_text(_token_at(toks, 1)) == "error":
            resume_next = token_text(_token_at(toks, 2)) == "resume"
            return
        if isinstance(node, StatementNode) and node.single_line_if_branches is not None:
            # A single-line If runs its statement on one path only.
            for key in _file_number_keys_in(toks):
                states.pop(key, None)
            _forget_paths_named_in(states, toks)
            return
        # A label may be reached from anywhere, an error handler's included, so
        # nothing is known there; and a call to a procedure may open or close any
        # file (XLIDE issue #146).
        if (
            jump_target_label_declaration(source, node.span) is not None
            or token_text(toks[0]) == "gosub"
        ):
            states.clear()
        if not resume_next:
            _check_path_functions(node.span, toks, states, push)
        if _check_open_path_use(node.span, toks, states, push, resume_next):
            return
        if (
            not _is_file_statement_head(token_text(toks[0]))
            and bare_call_statement_target(source, node.span) is not None
        ):
            states.clear()
            return
        # `f = FreeFile` again names a new file: what was known about f ends.
        # The value may still name a file number: `Main = LOF(0)`.
        assigned = bare_assignment_target(source, node.span)
        if assigned is not None:
            states.pop(assigned[0].lower(), None)
            _forget_path(states, f"path:{assigned[0].lower()}")
        # A path passed whole to a procedure may come back changed; a file
        # statement only reads it.
        if not _is_file_statement_head(token_text(toks[0])):

            def tracked(name: str) -> bool:
                return any(
                    key.startswith("path:") and name in _path_key_names(key) for key in states
                ) or any(
                    isinstance(state, _OpenFile)
                    and state.open_path is not None
                    and name in _path_key_names(state.open_path)
                    for state in states.values()
                )

            for lower in list(
                tracked_locals_named_whole(toks, node.span.start, tracked, _READ_ONLY_INTRINSICS)
            ):
                _forget_path(states, f"path:{lower}")
        _check_statement(node.span, toks, states, push, resume_next)

    def snapshot() -> _FileStates:
        return dict(states)

    def restore(saved: _FileStates) -> None:
        states.clear()
        for key, state in saved.items():
            states[key] = state

    def forget(keys: AbstractSet[str]) -> None:
        if "*" in keys:
            states.clear()
        for key in keys:
            states.pop(key, None)
            _forget_path(states, f"path:{key}")

    def touches(stmt: LeafStatementNode) -> set[str]:
        return _file_keys_touched_by(source, stmt)

    def enter(node: BodyNode) -> None:
        # `Do Until EOF(f)` checks before its body reads.
        header = statement_tokens_after_leading_label(
            source, block_header_line_span(source, node.span)
        )
        _mark_checked(states, header)
        # A condition may test the path: `If Len(Dir(p)) > 0 Then Kill p`.
        _forget_paths_named_in(states, header)
        if isinstance(node, IfBlockNode):
            for branch in node.branches:
                _forget_paths_named_in(
                    states, statement_tokens_after_leading_label(source, branch.header_span)
                )

    walk_entering_blocks(
        source,
        member.body,
        lambda node: is_inactive_node(activity, node),
        visit,
        BlockEnteringState(
            snapshot=snapshot,
            restore=restore,
            forget=forget,
            touches=touches,
            enter=enter,
        ),
    )


# File statements that raise 52 on a number nothing opened; Close runs
# (XLIDE issue #419).
_NUMBERED_STATEMENTS = frozenset(
    {"print", "write", "input", "line", "get", "put", "seek", "lock", "unlock", "width"}
)


def _check_unopened_numbers(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    opened: OpenedFileNumbers,
) -> None:
    """A literal file number no Open in the project names, while none names a
    variable or FreeFile: `Print #1, "x"` and `EOF(1)` raise 52, "Bad file
    name or number", wherever they run (XLIDE issue #419, measured in Excel
    16.0). The project's Opens come from the index, and this module's from its
    text as it stands."""
    if opened.any:
        return

    def unopened(tok: VbaToken | None) -> int | None:
        value = (
            int(tok.raw_text)
            if tok is not None
            and tok.kind is TokenKind.INTEGER_LITERAL
            and _DIGITS_RE.fullmatch(tok.raw_text) is not None
            else None
        )
        return (
            value
            if value is not None
            and value >= 1
            and value <= MAX_FILE_NUMBER
            and value not in opened.numbers
            else None
        )

    def report(base: Span, tok: VbaToken, value: int) -> None:
        push(
            "fileNumberZero",
            f"File number {value} is opened by no Open statement in this project, so nothing "
            "can be open on it. This will raise Run-time error '52': Bad file name or number.",
            Span(base.start + tok.start, base.start + tok.end),
        )

    def visit(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            toks = [
                tok
                for tok in statement_tokens_after_leading_label(source, span)
                if token_text(tok) != "else"
            ]
            head = token_text(_token_at(toks, 0))
            if head == "line":
                number_at = 2 if token_text(_token_at(toks, 1)) == "input" else -1
            else:
                number_at = 1 if head in _NUMBERED_STATEMENTS else -1
            marker = _token_at(toks, number_at) if number_at > 0 else None
            if number_at > 0 and marker is not None and marker.raw_text == "#":
                number_value = unopened(_token_at(toks, number_at + 1))
                if number_value is not None:
                    report(span, toks[number_at + 1], number_value)
                    continue
            # `EOF(1)`, `LOF(1)`, `Input(1, #1)`.
            for i in range(len(toks) - 2):
                name = token_text(toks[i])
                if toks[i + 1].raw_text != "(" or (i > 0 and toks[i - 1].raw_text == "."):
                    continue
                number_index: int | None
                if name in _FILE_STATE_FUNCTIONS:
                    number_index = i + 2
                elif (
                    name == "input" and _raw_at(toks, i + 3) == "," and _raw_at(toks, i + 4) == "#"
                ):
                    number_index = i + 5
                else:
                    number_index = None
                number = _token_at(toks, number_index) if number_index is not None else None
                if number is None or number_index is None:
                    continue
                following = _raw_at(toks, number_index + 1)
                value = unopened(number) if following == ")" or following == "," else None
                if value is not None:
                    report(span, number, value)

    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        for_each_statement(member.body, visit, activity)


def _check_statement(
    base: Span,
    toks: Sequence[VbaToken],
    states: _FileStates,
    push: PushFn,
    resume_next: bool,
) -> None:
    def at(tok: VbaToken) -> Span:
        return Span(base.start + tok.start, base.start + tok.end)

    def span_of(first: VbaToken, last: VbaToken) -> Span:
        return Span(base.start + first.start, base.start + last.end)

    head = token_text(toks[0])
    # Function forms: LOF(0), EOF(600), Loc(-1), FileAttr(0, 1), Seek(0).
    for i in range(len(toks) - 2):
        name = token_text(toks[i])
        if (
            name not in _FILE_STATE_FUNCTIONS
            or toks[i + 1].raw_text != "("
            or (i > 0 and toks[i - 1].raw_text == ".")
        ):
            continue
        impossible = _impossible_file_number(toks, i + 2)
        after = _raw_at(toks, _index_of(toks, impossible.last) + 1) if impossible else None
        if impossible is not None and (after == ")" or after == ","):
            push(
                "fileNumberZero",
                f"File number {impossible.value} is never open: file numbers run from 1 to "
                f"{MAX_FILE_NUMBER}. This will raise Run-time error '52': Bad file name or number.",
                span_of(impossible.first, impossible.last),
            )
    _mark_checked(states, toks)
    _report_empty_input_function(toks, states, push, at)
    if head == "open":
        opened = _parse_open(toks)
        if opened is None:
            return
        impossible = _impossible_file_number(toks, opened.number_index)
        if impossible is not None:
            push(
                "fileNumberZero",
                f"File number {impossible.value} cannot be opened: file numbers run from 1 to "
                f"{MAX_FILE_NUMBER}. This will raise Run-time error '52': Bad file name or number.",
                span_of(impossible.first, impossible.last),
            )
            return
        if opened.len is not None:
            len_value = opened.len.value
            if len_value == 0:
                push(
                    "runtimeArgumentValue",
                    "Argument 'Len' of 'Open' is 0; this will raise Run-time error '5': "
                    "Invalid procedure call or argument.",
                    at(opened.len.token),
                )
                return
            if len_value > 32767:
                push(
                    "arithmeticOverflow",
                    f"Open's Len of {len_value} does not fit an Integer. "
                    "This will raise Run-time error '6': Overflow.",
                    at(opened.len.token),
                )
                return
        empty = opened.path is not None and states.get(opened.path) == "empty"
        # For Input finds no file the procedure deleted; the other modes make one
        # (XLIDE issue #682).
        if (
            opened.path is not None
            and opened.mode == "input"
            and states.get(opened.path) == "absent"
            and not resume_next
        ):
            for_at = next((k for k, tok in enumerate(toks) if token_text(tok) == "for"), -1)
            path_toks = list(toks[1:for_at])
            push(
                "runtimeArgumentValue",
                f"Open For Input finds no file at {' '.join(tok.raw_text for tok in path_toks)}, "
                f"which this procedure deleted or moved above. {_PATH_ERRORS['missing']}",
                span_of(toks[0], path_toks[-1] if path_toks else toks[0]),
            )
        if opened.path is not None and opened.mode != "input":
            # Output empties it, and the other modes may write to it.
            states.pop(opened.path, None)
            _made_in(states, opened.path)
        if opened.key is None:
            return
        previous = states.get(opened.key)
        if isinstance(previous, _OpenFile):
            push(
                "fileAlreadyOpen",
                f"File number {_describe_key(opened.key)} is still open from the Open "
                "statement above; opening it again raises Run-time error '55': File already "
                "open. Close it first.",
                at(toks[opened.number_index]),
            )
        states[opened.key] = _OpenFile(
            opened.mode,
            base,
            path=opened.path if opened.mode == "output" and opened.path is not None else None,
            written=False if opened.mode == "output" and opened.path is not None else None,
            empty_unchecked=True if opened.mode == "input" and empty else None,
            open_path=opened.path,
        )
        return
    if head == "close" or head == "reset":
        for index in _close_number_indexes(toks):
            impossible = _impossible_file_number(toks, index)
            if impossible is not None:
                push(
                    "fileNumberZero",
                    f"File number {impossible.value} is never open: file numbers run from 1 to "
                    f"{MAX_FILE_NUMBER}. This will raise Run-time error '52': Bad file name or "
                    "number.",
                    span_of(impossible.first, impossible.last),
                )
                return
        keys = [] if head == "reset" else _file_number_keys_in(toks[1:])
        if len(keys) == 0:
            # `Close` with no number, and `Reset`, close every open file: a later
            # `Print #1` raises 52 in Excel (XLIDE issue #146).
            for open_key, known in list(states.items()):
                if isinstance(known, _OpenFile):
                    _close_file(states, open_key)
            return
        for closed in keys:
            _close_file(states, closed)
        return
    if head not in _FILE_STATEMENTS:
        return
    # `Line Input #f, x`: the statement word is two tokens.
    if head == "line":
        number_index = 2 if token_text(_token_at(toks, 1)) == "input" else -1
    else:
        number_index = 1
    if number_index < 0:
        return
    number_start = number_index + 1 if _raw_at(toks, number_index) == "#" else number_index
    number_token = _token_at(toks, number_start)
    if number_token is None:
        return
    impossible = _impossible_file_number(toks, number_start)
    if impossible is not None:
        push(
            "fileNumberZero",
            f"File number {impossible.value} is never open: file numbers run from 1 to "
            f"{MAX_FILE_NUMBER}. This will raise Run-time error '52': Bad file name or number.",
            span_of(impossible.first, impossible.last),
        )
        return
    key = _file_number_key(number_token)
    if key is None:
        return
    state = states.get(key)
    if state == "closed":
        push(
            "fileUsedAfterClose",
            f"File number {_describe_key(key)} was closed above and not opened again. "
            "This will raise Run-time error '52': Bad file name or number.",
            at(number_token),
        )
        return
    statement = "line input" if head == "line" else head
    # `Seek #f, 0` raises 63 whatever f is open for, and when nothing opened it
    # (XLIDE issue #262); Get and Put only reach the record in Binary and Random.
    if statement == "seek" or (
        (statement == "get" or statement == "put")
        and isinstance(state, _OpenFile)
        and (state.mode == "binary" or state.mode == "random")
    ):
        comma = next(
            (
                index
                for index, tok in enumerate(toks)
                if index > number_start and tok.raw_text == ","
            ),
            -1,
        )
        record = _record_below_one(toks, comma + 1) if comma > 0 else None
        if record is not None and (statement == "seek" or record.value == 0):
            push(
                "fileRecordZero",
                f"{statement[:1].upper() + statement[1:]} with record number {record.value}: "
                "records and Binary positions start at 1. "
                "This will raise Run-time error '63': Bad record number.",
                span_of(record.first, record.last),
            )
            return
    if not isinstance(state, _OpenFile):
        return
    writes = statement == "print" or statement == "write"
    reads = statement == "input" or statement == "line input"
    modes = _STATEMENT_MODES.get(statement)
    if modes is not None and state.mode not in modes:
        word = "Line Input" if statement == "line input" else head[:1].upper() + head[1:]
        push(
            "fileModeMismatch",
            f"'{word} #' on a file opened For {_mode_word(state.mode)} raises "
            "Run-time error '54': Bad file mode.",
            at(toks[0]),
        )
        return
    if (
        (writes or statement == "put")
        and state.mode == "output"
        and not state.written
        and not _writes_nothing(toks, number_start)
    ):
        state = replace(state, written=True)
        states[key] = state
    if reads and state.empty_unchecked:
        word = "Line Input" if statement == "line input" else "Input"
        push(
            "fileReadPastEnd",
            f"'{word} #' reads file {_describe_key(key)}, which this procedure created empty "
            "and reopened For Input without checking EOF. "
            "This will raise Run-time error '62': Input past end of file.",
            at(toks[0]),
        )
        states[key] = replace(state, empty_unchecked=False)


def _report_empty_input_function(
    toks: Sequence[VbaToken],
    states: _FileStates,
    push: PushFn,
    at: Callable[[VbaToken], Span],
) -> None:
    """`Input(n, #f)` and `InputB$(n, f)` read the file too."""
    for i in range(len(toks) - 1):
        name = token_text(toks[i])
        if (name != "input" and name != "inputb") or i == 0 or toks[i - 1].raw_text == ".":
            continue
        open_at = i + 2 if toks[i + 1].raw_text == "$" else i + 1
        if _raw_at(toks, open_at) != "(":
            continue
        comma = next((k for k, tok in enumerate(toks) if k > open_at and tok.raw_text == ","), -1)
        if comma <= 0:
            continue
        number_tok = (
            _token_at(toks, comma + 2)
            if _raw_at(toks, comma + 1) == "#"
            else _token_at(toks, comma + 1)
        )
        key = _file_number_key(number_tok)
        state = states.get(key) if key else None
        # `Input(1, #1)` reads an Input or Binary file only (XLIDE issue #419).
        if (
            key
            and isinstance(state, _OpenFile)
            and state.mode != "input"
            and state.mode != "binary"
        ):
            push(
                "fileModeMismatch",
                f"'{toks[i].raw_text}' reads file {_describe_key(key)}, opened For "
                f"{_mode_word(state.mode)}. This will raise Run-time error '54': Bad file mode.",
                at(toks[i]),
            )
            continue
        if key and isinstance(state, _OpenFile) and state.empty_unchecked:
            push(
                "fileReadPastEnd",
                f"'{toks[i].raw_text}' reads file {_describe_key(key)}, which this procedure "
                "created empty and reopened For Input without checking EOF. "
                "This will raise Run-time error '62': Input past end of file.",
                at(toks[i]),
            )
            states[key] = replace(state, empty_unchecked=False)


def _mark_checked(states: _FileStates, toks: Sequence[VbaToken]) -> None:
    """EOF, LOF, Loc and Seek on a file: a read after them may be guarded."""
    for i in range(len(toks) - 2):
        if token_text(toks[i]) not in _FILE_STATE_FUNCTIONS or toks[i + 1].raw_text != "(":
            continue
        key = _file_number_key(
            _token_at(toks, i + 3) if toks[i + 2].raw_text == "#" else toks[i + 2]
        )
        state = states.get(key) if key else None
        if key and isinstance(state, _OpenFile) and state.empty_unchecked:
            states[key] = replace(state, empty_unchecked=False)
    if token_text(_token_at(toks, 0)) == "seek":
        number_tok = _token_at(toks, 2) if _raw_at(toks, 1) == "#" else _token_at(toks, 1)
        key = _file_number_key(number_tok)
        state = states.get(key) if key else None
        if key and isinstance(state, _OpenFile) and state.empty_unchecked:
            states[key] = replace(state, empty_unchecked=False)


def _close_file(states: _FileStates, key: str) -> None:
    """Closes a number; an Output file closed with nothing written leaves its path empty."""
    state = states.get(key)
    if isinstance(state, _OpenFile) and state.path is not None:
        states[state.path] = (
            "empty" if state.mode == "output" and state.written is False else "file"
        )
    elif isinstance(state, _OpenFile) and state.open_path is not None and state.mode != "input":
        # Append, Binary and Random made the file if it was not there.
        states[state.open_path] = "file"
    states[key] = "closed"


def _writes_nothing(toks: Sequence[VbaToken], number_start: int) -> bool:
    """`Print #f, "";` writes nothing (measured: a later read raises 62)."""
    rest = toks[number_start + 2 :]
    return (
        token_text(toks[0]) == "print"
        and len(rest) == 2
        and rest[0].raw_text == '""'
        and rest[1].raw_text == ";"
    )


def _forget_paths_named_in(states: _FileStates, toks: Sequence[VbaToken]) -> None:
    for tok in toks:
        name = token_name(tok)
        lower = name.lower() if name is not None else None
        if lower:
            _forget_path(states, f"path:{lower}")


_PATH_WILDCARD_RE = re.compile(r'["*?]')


def _path_key_of(toks: Sequence[VbaToken]) -> str | None:
    """A path the rule can follow, as a `path:` key: names and string literals
    joined by `&`, `p` or `d & "\\a.txt"`. Not a literal with a wildcard, which
    Kill takes as a pattern."""
    operands = [tok for tok in toks if tok.kind is not TokenKind.COMMENT]
    if len(operands) % 2 == 0:
        return None
    parts: list[str] = []
    for k, tok in enumerate(operands):
        if k % 2 == 1:
            if tok.raw_text != "&":
                return None
            continue
        name = token_name(tok) if tok.kind is TokenKind.IDENTIFIER else None
        literal = (
            string_literal_value(tok.raw_text) if tok.kind is TokenKind.STRING_LITERAL else None
        )
        if literal is not None and _PATH_WILDCARD_RE.search(literal) is None:
            parts.append(f'"{literal}"')
        elif name is not None:
            parts.append(name.lower())
        else:
            return None
    return f"path:{'&'.join(parts)}"


_PATH_KEY_PART_RE = re.compile(r'"[^"]*"|[^&"]+')


def _path_key_names(key: str) -> set[str]:
    """The names a `path:` key is built from."""
    return {
        part for part in _PATH_KEY_PART_RE.findall(key[len("path:") :]) if not part.startswith('"')
    }


# JS `.` stops at line terminators; a path key never holds one, as a string
# literal cannot span lines.
_JOINED_PARENT_RE = re.compile(r'(path:.+)&"\\([^"\\]+)"')
_LITERAL_PARENT_RE = re.compile(r'path:"(.+)\\[^"\\]+"')


def _parent_path_key(key: str) -> str | None:
    """The folder a path is directly in, when its last part is a literal `\\name`:
    `d` for `d & "\\a.txt"`."""
    joined = _JOINED_PARENT_RE.fullmatch(key)
    if joined is not None:
        return joined.group(1)
    literal = _LITERAL_PARENT_RE.fullmatch(key)
    return f'path:"{literal.group(1)}"' if literal is not None else None


_PLAIN_NAME_RE = re.compile(r'[^"&]+')


def _forget_path(states: _FileStates, path: str) -> None:
    """A path key whose name may now hold another path: neither an empty file nor
    an open one is known by it, nor by any path built from that name."""
    name = path[len("path:") :]
    plain = _PLAIN_NAME_RE.fullmatch(name) is not None

    def named(key: str) -> bool:
        return key == path or (plain and key.startswith("path:") and name in _path_key_names(key))

    for key, state in list(states.items()):
        if named(key):
            del states[key]
        elif (
            isinstance(state, _OpenFile) and state.open_path is not None and named(state.open_path)
        ):
            states[key] = replace(state, open_path=None)


def _made_in(states: _FileStates, path: str) -> None:
    """A file or folder made at a path: what is known of the folder it is in."""
    parent = _parent_path_key(path)
    if parent and (states.get(parent) == "folder" or states.get(parent) == "filled"):
        states[parent] = "filled"


def _gone_from(states: _FileStates, path: str) -> None:
    """A file or folder gone from a path: the folder it was in may now be empty."""
    parent = _parent_path_key(path)
    if parent:
        states.pop(parent, None)


# The errors a statement on a path the procedure deleted or made raises
# (XLIDE issue #682).
_PATH_ERRORS: Mapping[str, str] = {
    "missing": "This will raise Run-time error '53': File not found.",
    "exists": "This will raise Run-time error '58': File already exists.",
    "access": "This will raise Run-time error '75': Path/File access error.",
    "noFolder": "This will raise Run-time error '76': Path not found.",
}

_PATH_STATEMENT_WORDS: Mapping[str, str] = {
    "kill": "Kill",
    "filecopy": "FileCopy",
    "name": "Name",
    "mkdir": "MkDir",
    "rmdir": "RmDir",
}


def _check_open_path_use(
    base: Span,
    toks: Sequence[VbaToken],
    states: _FileStates,
    push: PushFn,
    resume_next: bool,
) -> bool:
    """`Kill p`, `FileCopy p, q` or `Name p As q` while p is open (FileCopy: open
    in any mode but Input): Run-time error 55, File already open (XLIDE issue
    #419, measured in Excel 16.0). And what these and MkDir and RmDir find where
    the procedure deleted or made something (XLIDE issue #682, measured in Excel
    16.0): Kill, FileCopy or Name of a file it deleted raises 53, Name onto a
    file it made 58, MkDir of a folder it made 75, RmDir of a folder it made
    something in 75 and of one it removed 76. Under On Error Resume Next nothing
    is reported, and a Kill or RmDir still leaves nothing there. True when the
    statement is one of these, which changes no file number."""
    head = token_text(toks[0])
    path_statement = (
        head == "kill" or head == "filecopy" or head == "name" or head == "mkdir" or head == "rmdir"
    )
    second = _raw_at(toks, 1)
    if (
        not path_statement
        or (head == "name" and not any(token_text(tok) == "as" for tok in toks))
        or (second or "") in ("=", ".", "(")
    ):
        return False
    end = next(
        (
            k
            for k, tok in enumerate(toks)
            if k > 0 and (tok.raw_text == "," or token_text(tok) == "as")
        ),
        -1,
    )
    path_toks = list(toks[1:]) if end < 0 else list(toks[1:end])
    target_toks = [] if end < 0 else list(toks[end + 1 :])
    path = _path_key_of(path_toks)
    target = _path_key_of(target_toks)

    def shown(part: Sequence[VbaToken]) -> str:
        return " ".join(tok.raw_text for tok in part)

    def span(part: Sequence[VbaToken]) -> Span:
        return Span(base.start + toks[0].start, base.start + part[-1].end)

    word = _PATH_STATEMENT_WORDS[head]
    # FileCopy reads a file open For Input, and is refused one open in any other mode.
    open_entry = (
        next(
            (
                (key, state)
                for key, state in states.items()
                if isinstance(state, _OpenFile)
                and state.open_path == path
                and (head != "filecopy" or state.mode != "input")
            ),
            None,
        )
        if path
        else None
    )
    if open_entry is not None and head != "mkdir" and head != "rmdir":
        push(
            "fileAlreadyOpen",
            f"'{shown(path_toks)}' is the path of file {_describe_key(open_entry[0])}, still "
            f"open from the Open statement above; {word} on an open file raises Run-time error "
            "'55': File already open. Close it first.",
            span(path_toks),
        )
    fact = states.get(path) if path else None
    target_fact = states.get(target) if target else None
    problem: str | None = None
    if (head == "kill" or head == "filecopy" or head == "name") and fact == "absent":
        problem = (
            f"{word} finds no file at {shown(path_toks)}, which this procedure deleted or "
            f"moved above. {_PATH_ERRORS['missing']}"
        )
    elif head == "name" and (target_fact == "file" or target_fact == "empty"):
        problem = (
            f"Name finds {shown(target_toks)} already there, a file this procedure made "
            f"above. {_PATH_ERRORS['exists']}"
        )
    elif head == "mkdir" and (fact == "folder" or fact == "filled"):
        problem = (
            f"MkDir finds the folder {shown(path_toks)} already there, made by this "
            f"procedure above. {_PATH_ERRORS['access']}"
        )
    elif head == "rmdir" and fact == "filled":
        problem = (
            f"{shown(path_toks)} holds what this procedure made in it above, and RmDir "
            f"removes only an empty folder. {_PATH_ERRORS['access']}"
        )
    elif head == "rmdir" and fact == "absent":
        problem = (
            f"RmDir finds no folder {shown(path_toks)}, which this procedure removed "
            f"above. {_PATH_ERRORS['noFolder']}"
        )
    if problem and open_entry is None and not resume_next:
        push(
            "runtimeArgumentValue",
            problem,
            span(target_toks if head == "name" and "'58'" in problem else path_toks),
        )
    # What the statement leaves, when it runs or Resume Next goes past it.
    if not path:
        if target:
            _forget_path(states, target)
        return True
    # A Kill of an open file, or an RmDir of a full folder, leaves it there.
    if (head == "kill" and open_entry is None) or (head == "rmdir" and fact != "filled"):
        _forget_path(states, path)
        states[path] = "absent"
        _gone_from(states, path)
    elif head == "mkdir":
        states[path] = "filled" if fact == "filled" else "folder"
        _made_in(states, path)
    elif head == "name" or head == "filecopy":
        # Under Resume Next it may not have run, so nothing is known of either path.
        ran = not resume_next and not problem and open_entry is None
        if head == "name":
            _forget_path(states, path)
            _gone_from(states, path)
        if target:
            _forget_path(states, target)
        if ran and target:
            if head == "name":
                states[path] = "absent"
            states[target] = "empty" if fact == "empty" else "file"
            _made_in(states, target)
    return True


def _check_path_functions(
    base: Span, toks: Sequence[VbaToken], states: _FileStates, push: PushFn
) -> None:
    """FileLen, GetAttr and FileDateTime of a file the procedure deleted raise 53
    (XLIDE issue #682)."""
    for i in range(len(toks) - 1):
        name = token_text(toks[i])
        if (
            (name != "filelen" and name != "getattr" and name != "filedatetime")
            or toks[i + 1].raw_text != "("
            or (i > 0 and toks[i - 1].raw_text == ".")
        ):
            continue
        close = match_paren_from(toks, i + 1)
        path_toks = list(toks[i + 2 : close]) if close >= 0 else []
        path = _path_key_of(path_toks) if close > i + 2 else None
        if path and states.get(path) == "absent":
            push(
                "runtimeArgumentValue",
                f"{toks[i].raw_text} finds no file at {' '.join(tok.raw_text for tok in path_toks)}, "
                f"which this procedure deleted or moved above. {_PATH_ERRORS['missing']}",
                Span(base.start + toks[i].start, base.start + toks[close].end),
            )


def _impossible_file_number(toks: Sequence[VbaToken], index: int) -> _SignedLiteral | None:
    """The literal file number at toks[index], a minus sign included, when no file
    can have it: 0, below 0, or past 512."""
    literal = _signed_integer_at(toks, index)
    return (
        literal
        if literal is not None and (literal.value < 1 or literal.value > MAX_FILE_NUMBER)
        else None
    )


def _record_below_one(toks: Sequence[VbaToken], index: int) -> _SignedLiteral | None:
    """A record or position literal below 1 at toks[index], standing alone in its slot."""
    literal = _signed_integer_at(toks, index)
    if literal is None or literal.value >= 1:
        return None
    after = _token_at(toks, _index_of(toks, literal.last) + 1)
    return literal if after is None or after.raw_text == "," else None


def _signed_integer_at(toks: Sequence[VbaToken], index: int) -> _SignedLiteral | None:
    negative = _raw_at(toks, index) == "-"
    tok = _token_at(toks, index + 1 if negative else index)
    if tok is None or tok.kind is not TokenKind.INTEGER_LITERAL:
        return None
    value = parse_vba_integer_literal(tok.raw_text)
    if value is None:
        return None
    first = toks[index]
    return _SignedLiteral(-value if negative else value, first, tok)


def _close_number_indexes(toks: Sequence[VbaToken]) -> list[int]:
    """Where each number a Close names starts: after `#`, or alone in its slot."""
    if token_text(_token_at(toks, 0)) != "close":
        return []
    out: list[int] = []
    for i in range(1, len(toks)):
        if toks[i].raw_text == "#":
            out.append(i + 1)
        elif i == 1 or toks[i - 1].raw_text == ",":
            out.append(i)
    return out


def _parse_open(toks: Sequence[VbaToken]) -> _ParsedOpen | None:
    mode = "random"
    number_index = -1
    path_end = -1
    depth = 0
    for i in range(1, len(toks)):
        raw = toks[i].raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
        if depth != 0:
            continue
        word = token_text(toks[i])
        if word == "for":
            path_end = i if path_end < 0 else path_end
            following = token_text(_token_at(toks, i + 1))
            if following in _FILE_MODES:
                mode = following
        elif word == "access" or word == "shared" or word == "lock":
            path_end = i if path_end < 0 else path_end
        elif word == "as":
            path_end = i if path_end < 0 else path_end
            number_index = i + 2 if _raw_at(toks, i + 1) == "#" else i + 1
            break
    if number_index < 0 or _token_at(toks, number_index) is None:
        return None
    path = _path_key_of(toks[1:path_end])
    # `Len = 0` after the number.
    len_at = next(
        (
            k
            for k, tok in enumerate(toks)
            if k > number_index and token_text(tok) == "len" and _raw_at(toks, k + 1) == "="
        ),
        -1,
    )
    len_tok = _token_at(toks, len_at + 2) if len_at > 0 else None
    len_value = (
        parse_vba_integer_literal(len_tok.raw_text)
        if len_tok is not None
        and len_tok.kind is TokenKind.INTEGER_LITERAL
        and _token_at(toks, len_at + 3) is None
        else None
    )
    return _ParsedOpen(
        mode,
        _file_number_key(toks[number_index]),
        number_index,
        path=path if path else None,
        len=_OpenLen(len_value, len_tok) if len_value is not None and len_tok is not None else None,
    )


def _file_number_key(tok: VbaToken | None) -> str | None:
    """The key a file number token identifies: its literal value, or the variable's name."""
    if tok is None:
        return None
    if tok.kind is TokenKind.INTEGER_LITERAL:
        # `#&H1` is file 1, not #NaN (XLIDE issue #146).
        value = parse_vba_integer_literal(tok.raw_text)
        return None if value is None else f"#{value}"
    name = token_name(tok)
    return name.lower() if name else None


def _is_file_statement_head(head: str) -> bool:
    """True for the statement words this rule follows: Open, Close, Reset and the
    file I/O statements."""
    return head in ("open", "close", "reset") or head in _FILE_STATEMENTS


def _describe_key(key: str) -> str:
    return key if key.startswith("#") else f"'{key}'"


def _mode_word(mode: str) -> str:
    return mode[:1].upper() + mode[1:]


def _token_at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """toks[index], or None past either end, as a JavaScript out-of-range read."""
    return toks[index] if 0 <= index < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], index: int) -> str | None:
    """toks[index]?.rawText."""
    tok = _token_at(toks, index)
    return tok.raw_text if tok is not None else None


def _index_of(toks: Sequence[VbaToken], tok: VbaToken) -> int:
    """toks.indexOf(tok): the first position holding this very token, or -1."""
    return next((k for k, candidate in enumerate(toks) if candidate is tok), -1)


def _file_number_keys_in(toks: Sequence[VbaToken]) -> list[str]:
    """Every file-number key a statement's tokens name after `#` or `As`."""
    out: list[str] = []
    for i, tok in enumerate(toks):
        if tok.raw_text == "#" and i + 1 < len(toks):
            key = _file_number_key(toks[i + 1])
            if key:
                out.append(key)
        elif i == 0 or toks[i - 1].raw_text == ",":
            key = _file_number_key(tok)
            if key and (i + 1 >= len(toks) or toks[i + 1].raw_text == ","):
                out.append(key)
    return out


_PATH_STATEMENT_HEADS = frozenset({"kill", "filecopy", "name", "mkdir", "rmdir"})


def _file_keys_touched_by(source: str, node: LeafStatementNode) -> set[str]:
    """The file number keys a statement may open, close or reopen, and the name it
    assigns; `*` when it may close anything: a procedure call, Reset, or a Close
    with no number."""
    out: set[str] = set()
    toks = statement_tokens_after_leading_label(source, node.span)
    head = token_text(_token_at(toks, 0))
    if (
        not _is_file_statement_head(head)
        and bare_call_statement_target(source, node.span) is not None
    ):
        out.add("*")
    if _is_file_statement_head(head):
        for key in _file_number_keys_in(toks):
            out.add(key)
        opened = _parse_open(toks) if head == "open" else None
        if opened is not None and opened.key:
            out.add(opened.key)
        if opened is not None and opened.path:
            out.add(opened.path[len("path:") :])
            parent = _parent_path_key(opened.path)
            if parent:
                out.add(parent[len("path:") :])
        if head == "reset" or (head == "close" and len(_file_number_keys_in(toks[1:])) == 0):
            out.add("*")
    # Kill, FileCopy, Name, MkDir and RmDir change what their paths and the
    # folders those are in name (XLIDE issue #682).
    if head in _PATH_STATEMENT_HEADS:
        parts: list[list[VbaToken]] = []
        for group in split_top_level_token_groups(toks, 1, ",", len(toks)):
            as_at = next((k for k, tok in enumerate(group) if token_text(tok) == "as"), -1)
            if as_at < 0:
                parts.append(group)
            else:
                parts.append(group[:as_at])
                parts.append(group[as_at + 1 :])
        for part in parts:
            path = _path_key_of(part)
            for path_key in [path, _parent_path_key(path)] if path else []:
                if path_key:
                    out.add(path_key[len("path:") :])
    target = bare_assignment_target(source, node.span)
    if target is not None:
        out.add(target[0].lower())
    return out
