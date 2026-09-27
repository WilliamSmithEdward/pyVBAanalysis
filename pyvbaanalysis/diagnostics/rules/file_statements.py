"""Rule family: file statements whose failure the code proves (XLIDE issue #123).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/fileStatements.ts.

Each case was measured in Excel 16.0 (build 20326, 2026-09-26): it compiles
and raises every time it runs.

 - file-number-zero (52, Bad file name or number): `As #0`, `LOF(0)` - a
   file number is 1 to 511, and 0 is never one.
 - file-used-after-close (52): `Close #f` and then `Print #f, ...` on the
   same number with no Open between.
 - file-mode-mismatch (54, Bad file mode): `Print #f`/`Write #f` on a file
   opened For Input; `Input #f`/`Line Input #f` on one opened For Output or
   Append.
 - file-already-open (55, File already open): two Opens As the same number
   with no Close between.
 - file-record-zero (63, Bad record number): `Seek #f, 0`, `Get #f, 0, x`,
   `Put #f, 0, x` - records and Binary positions start at 1.

A file number is a literal, hex or octal included, or a local that FreeFile
fills once. The rule follows the top-level statements of a procedure in order; a
block between two statements that could touch the number ends what is known
about it, and a number named inside a block is never followed. A label, a GoSub
or a call to a procedure ends everything known, and Reset or a Close with no
number closes every open file (XLIDE issue #146).
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Literal, Union

from ...call.call_context import bare_call_statement_target
from ...conditional import ConditionalActivityTracker, inactive_node_skip
from ...constants.integer_constant_expression import parse_vba_integer_literal
from ...flow.procedure_labels import statement_label_declaration
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    VariableGroupNode,
    is_leaf_statement,
    iter_body_nodes_in_context,
)
from ..context import PushFn
from ..walker import (
    active_module_members,
    bare_assignment_target,
    is_inactive_node,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

_FILE_MODES = frozenset({"input", "output", "append", "random", "binary"})


@dataclass(frozen=True, slots=True)
class _OpenFile:
    mode: str
    span: Span


# What is known about each file number key as the statements run.
_FileStates = dict[str, Union[_OpenFile, Literal["closed"]]]


@dataclass(frozen=True, slots=True)
class _OpenStatement:
    mode: str
    key: str | None
    number_token: VbaToken


_FILE_STATEMENTS = frozenset(
    {"print", "write", "input", "line", "get", "put", "seek", "close", "lock", "unlock", "width"}
)

_FILE_NUMBER_FUNCTIONS = frozenset({"lof", "eof", "loc", "fileattr", "seek"})

_ZERO_LITERAL_RE = re.compile(r"0+[%&^]?")


def check_file_statements(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        states: _FileStates = {}
        numbers_in_blocks = _file_numbers_named_in_blocks(source, member.body, activity)
        for node in member.body:
            if is_inactive_node(activity, node):
                continue
            if isinstance(node, VariableGroupNode):
                continue  # a Dim inside the body declares, and runs nothing
            if not is_leaf_statement(node):
                # A block may open, close or reopen anything it names.
                if "*" in numbers_in_blocks:
                    states.clear()
                for key in numbers_in_blocks:
                    states.pop(key, None)
                continue
            toks = statement_tokens_after_leading_label(source, node.span)
            if len(toks) == 0:
                continue
            if isinstance(node, StatementNode) and node.single_line_if_branches is not None:
                # A single-line If runs its statement on one path only.
                for key in _file_number_keys_in(toks):
                    states.pop(key, None)
                continue
            # A label may be reached from anywhere, an error handler's included, so
            # nothing is known there; and a call to a procedure may open or close any
            # file (XLIDE issue #146).
            if (
                statement_label_declaration(source, node.span) is not None
                or token_text(toks[0]) == "gosub"
            ):
                states.clear()
            if (
                not _is_file_statement_head(token_text(toks[0]))
                and bare_call_statement_target(source, node.span) is not None
            ):
                states.clear()
                continue
            # `f = FreeFile` again names a new file: what was known about f ends.
            # The value may still name a file number: `Main = LOF(0)`.
            assigned = bare_assignment_target(source, node.span)
            if assigned is not None:
                states.pop(assigned[0].lower(), None)
            _check_statement(node.span, toks, states, push)


def _check_statement(
    base: Span, toks: Sequence[VbaToken], states: _FileStates, push: PushFn
) -> None:
    def at(tok: VbaToken) -> Span:
        return Span(base.start + tok.start, base.start + tok.end)

    head = token_text(toks[0])
    # Function forms: LOF(0), EOF(0), Loc(0), FileAttr(0, 1), Seek(0).
    for i in range(len(toks) - 2):
        name = token_text(toks[i])
        after = _token_at(toks, i + 3)
        if (
            name in _FILE_NUMBER_FUNCTIONS
            and toks[i + 1].raw_text == "("
            and not (i > 0 and toks[i - 1].raw_text == ".")
            and _is_zero_literal(toks[i + 2])
            and after is not None
            and (after.raw_text == ")" or after.raw_text == ",")
        ):
            push(
                "fileNumberZero",
                "File number 0 is never open: file numbers run from 1 to 511. "
                "This will raise Run-time error '52': Bad file name or number.",
                at(toks[i + 2]),
            )
    if head == "open":
        opened = _parse_open(toks)
        if opened is None:
            return
        if _is_zero_literal(opened.number_token):
            push(
                "fileNumberZero",
                "File number 0 cannot be opened: file numbers run from 1 to 511. "
                "This will raise Run-time error '52': Bad file name or number.",
                at(opened.number_token),
            )
            return
        if opened.key is None:
            return
        previous = states.get(opened.key)
        if isinstance(previous, _OpenFile):
            push(
                "fileAlreadyOpen",
                f"File number {_describe_key(opened.key)} is still open from the Open "
                "statement above; opening it again raises Run-time error '55': File already "
                "open. Close it first.",
                at(opened.number_token),
            )
        states[opened.key] = _OpenFile(opened.mode, base)
        return
    if head == "close" or head == "reset":
        keys = [] if head == "reset" else _file_number_keys_in(toks[1:])
        if len(keys) == 0:
            # `Close` with no number, and `Reset`, close every open file: a later
            # `Print #1` raises 52 in Excel (XLIDE issue #146).
            for open_key, known in list(states.items()):
                if known != "closed":
                    states[open_key] = "closed"
            return
        for closed in keys:
            states[closed] = "closed"
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
    marker = _token_at(toks, number_index)
    has_marker = marker is not None and marker.raw_text == "#"
    number_start = number_index + 1 if has_marker else number_index
    number_token = _token_at(toks, number_start)
    if number_token is None:
        return
    if _is_zero_literal(number_token):
        push(
            "fileNumberZero",
            "File number 0 is never open: file numbers run from 1 to 511. "
            "This will raise Run-time error '52': Bad file name or number.",
            at(number_token),
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
    if not isinstance(state, _OpenFile):
        return
    statement = "line input" if head == "line" else head
    writes = statement == "print" or statement == "write"
    reads = statement == "input" or statement == "line input"
    if (writes and state.mode == "input") or (
        reads and (state.mode == "output" or state.mode == "append")
    ):
        word = "Line Input" if statement == "line input" else head[:1].upper() + head[1:]
        push(
            "fileModeMismatch",
            f"'{word} #' on a file opened For {_mode_word(state.mode)} raises "
            "Run-time error '54': Bad file mode.",
            at(toks[0]),
        )
        return
    if (statement == "seek" or statement == "get" or statement == "put") and (
        state.mode == "binary" or state.mode == "random"
    ):
        # `Seek #f, 0` / `Get #f, 0, x`: the record or position after the comma.
        comma = next(
            (
                index
                for index, tok in enumerate(toks)
                if index > number_start and tok.raw_text == ","
            ),
            -1,
        )
        record = _token_at(toks, comma + 1) if comma > 0 else None
        following = _token_at(toks, comma + 2)
        if (
            record is not None
            and _is_zero_literal(record)
            and (following is None or following.raw_text == ",")
        ):
            push(
                "fileRecordZero",
                f"{statement[:1].upper() + statement[1:]} with record number 0: records and "
                "Binary positions start at 1. "
                "This will raise Run-time error '63': Bad record number.",
                at(record),
            )


def _parse_open(toks: Sequence[VbaToken]) -> _OpenStatement | None:
    mode = "random"
    number_token: VbaToken | None = None
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
            following = token_text(_token_at(toks, i + 1))
            if following in _FILE_MODES:
                mode = following
        elif word == "as":
            marker = _token_at(toks, i + 1)
            number_token = (
                _token_at(toks, i + 2) if marker is not None and marker.raw_text == "#" else marker
            )
            break
    if number_token is None:
        return None
    return _OpenStatement(mode, _file_number_key(number_token), number_token)


def _file_number_key(tok: VbaToken) -> str | None:
    """The key a file number token identifies: its literal value, or the variable's name."""
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


def _is_zero_literal(tok: VbaToken | None) -> bool:
    return (
        tok is not None
        and tok.kind is TokenKind.INTEGER_LITERAL
        and _ZERO_LITERAL_RE.fullmatch(tok.raw_text) is not None
    )


def _mode_word(mode: str) -> str:
    return mode[:1].upper() + mode[1:]


def _token_at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """toks[index], or None past either end, as a JavaScript out-of-range read."""
    return toks[index] if 0 <= index < len(toks) else None


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


def _file_numbers_named_in_blocks(
    source: str,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
) -> set[str]:
    """File-number keys any nested block's statements name, plus every local a
    block assigns (`f = FreeFile` inside an If): a top-level Open or Close of
    such a number is followed only until the block, and a Close inside one is
    not seen at all.

    The context is whether a node sits inside a block: False for the
    procedure's own statements, True at every depth below them.
    """
    out: set[str] = set()
    for node, nested in iter_body_nodes_in_context(
        body, False, lambda _block, _outer: True, inactive_node_skip(activity)
    ):
        if not nested or not is_leaf_statement(node):
            continue
        toks = statement_tokens_after_leading_label(source, node.span)
        head = token_text(_token_at(toks, 0))
        if (
            not _is_file_statement_head(head)
            and bare_call_statement_target(source, node.span) is not None
        ):
            out.add("*")  # a procedure called inside the block may close anything
        if _is_file_statement_head(head):
            for key in _file_number_keys_in(toks):
                out.add(key)
            opened = _parse_open(toks) if head == "open" else None
            if opened is not None and opened.key:
                out.add(opened.key)
            if head == "reset" or (head == "close" and len(_file_number_keys_in(toks[1:])) == 0):
                out.add("*")
        target = bare_assignment_target(source, node.span)
        if target is not None:
            out.add(target[0].lower())
    return out
