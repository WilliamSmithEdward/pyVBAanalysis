"""Rule family: late-bound objects whose state the code's literals make plain
(XLIDE issue #477).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/lateBoundObjects.ts.
Measured in Excel 16.0 (build 20430, 2026-10-02); each compiles and raises when
it runs. A ProgID no class has, a RegExp member or pattern VBScript refuses, Null
into Test or Execute, and text into a Boolean flag are late_bound_members.py's;
this follows what the objects hold from one statement to the next.

 - A VBScript.RegExp's Execute(s)(n) past the matches, or SubMatches(k) past
   the groups, raises 5.
 - A Scripting.FileSystemObject's OpenTextFile with an IOMode other than 1, 2
   or 8 raises 5. A TextStream opened to read refuses a write, and one opened
   to write or append refuses a read, with 54; after its Close any member but
   Close raises 91. Through one path variable or literal, the procedure's own
   file operations are followed: a file it created and wrote nothing to raises
   62 when read, CreateTextFile with Overwrite False on it 58, and DeleteFile,
   GetFile or OpenTextFile to read or append after it deleted the file 53.
   `If fso.FileExists(p) Then fso.DeleteFile p` leaves no p.
 - An ADODB.Recordset never opened raises 3704 at MoveNext, MoveFirst, EOF,
   BOF, Close or RecordCount. State and Fields.Count run.
 - An MSXML2.DOMDocument with no element, new or loaded from malformed XML,
   raises 91 at DocumentElement's members, and SelectSingleNode of a path no
   element of literal XML is on gives Nothing, whose members raise 91. A path
   with a step and no node test raises -2147467259.

A local is followed from the Set that creates it; any other mention of it whole,
which may hand it to code that changes it, ends what is known.

Upstream runs Execute's pattern with JavaScript's RegExp. The port translates the
pattern to Python's `re` with JavaScript's meaning (`.`, `^`, `$`, `\\d`, `\\w`,
`\\s`, `\\b`, identity escapes, a `{` that is no quantifier); the few forms it
cannot translate (`\\D`, `\\W` or `\\S` inside a class) report nothing.
"""

from __future__ import annotations

import dataclasses
import json
import re
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass
from typing import Union

from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import statement_label_declaration
from ...js_compat import JS_WHITESPACE, js_trim
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    is_leaf_statement,
)
from ..call_extraction import string_literal_value
from ..context import PushFn
from ..dataflow import BlockEnteringState, walk_entering_blocks
from ..known_string_calls import StringFoldContext, fold_string_expression
from ..walker import (
    active_module_members,
    set_assignment_target,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .shared import names_in

# The text constants an Execute subject is built with (XLIDE issue #685).
_SUBJECT_CONSTANTS: dict[str, str] = {
    "vbcr": chr(13),
    "vblf": chr(10),
    "vbcrlf": chr(13) + chr(10),
    "vbnewline": chr(13) + chr(10),
    "vbtab": chr(9),
}


@dataclass(frozen=True, slots=True)
class XmlElement:
    name: str
    children: tuple[XmlElement, ...]
    # Each attribute's value as written; None where it holds an entity reference.
    attributes: Mapping[str, str | None] | None = None


@dataclass(slots=True)
class _RegExpState:
    pattern: str | None
    flags: dict[str, bool | None]


@dataclass(slots=True)
class _FsoState:
    # What the procedure has done to each path, by the variable or literal that names it:
    # "empty", "full" or "absent".
    files: dict[str, str]


@dataclass(slots=True)
class _TextStreamState:
    mode: str  # "read" | "write"
    closed: bool
    # Opened to read a file the procedure left empty.
    empty: bool | None = None
    # The path it writes.
    file: str | None = None


@dataclass(slots=True)
class _RecordsetState:
    open: bool | None = None
    closed: bool | None = None


@dataclass(slots=True)
class _DomDocState:
    # The document element, or None for a document with none: new, or loaded from
    # malformed XML.
    root: XmlElement | None


_LateObject = Union[_RegExpState, _FsoState, _TextStreamState, _RecordsetState, _DomDocState]

_REGEXP_MEMBERS: frozenset[str] = frozenset(
    {"pattern", "global", "ignorecase", "multiline", "test", "execute", "replace"}
)
_REGEXP_FLAGS: frozenset[str] = frozenset({"global", "ignorecase", "multiline"})
_READS: frozenset[str] = frozenset({"readline", "readall", "read", "skip", "skipline"})
_WRITES: frozenset[str] = frozenset({"write", "writeline", "writeblanklines"})
_CLOSED_RECORDSET: frozenset[str] = frozenset(
    {"movenext", "moveprevious", "movefirst", "movelast", "eof", "bof", "close", "recordcount"}
)


def _at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """toks[index], or None past either end (JavaScript's `toks[i]?.`)."""
    return toks[index] if 0 <= index < len(toks) else None


def _raw(toks: Sequence[VbaToken], index: int) -> str | None:
    tok = _at(toks, index)
    return tok.raw_text if tok is not None else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name else None


def check_late_bound_objects(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    project_callables: Iterable[str] = (),
) -> None:
    callables = {node.name.lower() for node in mod.members if isinstance(node, ProcedureNode)} | {name.lower() for name in project_callables}
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        _check_procedure(source, member, activity, push, callables)


def _check_procedure(
    source: str, member: ProcedureNode, activity: ConditionalActivityTracker | None, push: PushFn,
    callables: Iterable[str],
) -> None:
    if re.search(r"\bon\s+error\b", source[member.span.start:member.span.end], re.IGNORECASE):
        return
    from ..walker import for_each_variable_group
    from ...parser.nodes import VariableGroupNode
    locals_: set[str] = set()

    def collect(group: VariableGroupNode) -> None:
        if not group.is_const and group.modifier.lower() != "static":
            locals_.update(decl.name.lower() for decl in group.declarations)

    for_each_variable_group(member.body, collect, activity)
    states: dict[str, _LateObject] = {}

    def forget(names: Iterable[str]) -> None:
        for lower in names:
            states.pop(lower, None)

    def visit(node: BodyNode) -> None:
        if not is_leaf_statement(node):
            return
        toks = statement_tokens_after_leading_label(source, node.span)
        if any(token_name(tok) is not None and token_text(tok) in callables and not (token_text(tok) == member.name.lower() and _raw(toks, i + 1) == "=") for i, tok in enumerate(toks)):
            states.clear()
        file_owners = {token_text(tok) for i, tok in enumerate(toks) if _raw(toks, i + 1) == "." and isinstance(states.get(token_text(tok)), _FsoState)}
        if file_owners:
            for lower, held in states.items():
                if isinstance(held, _FsoState) and (lower not in file_owners or len(file_owners) > 1):
                    held.files = {}
                if isinstance(held, _TextStreamState):
                    held.empty = None
        if (
            statement_label_declaration(source, node.span) is not None
            or token_text(_at(toks, 0)) == "gosub"
        ):
            states.clear()
        # `p = ...` names another path with the same variable.
        head = 1 if token_text(_at(toks, 0)) == "let" else 0
        assigned = _lower_name(_at(toks, head)) if _raw(toks, head + 1) == "=" else None
        if assigned:
            for state in states.values():
                if isinstance(state, _FsoState):
                    state.files.pop(assigned, None)
        # A one-line If: its condition and each arm are checked on a copy, and what
        # an arm may change is no longer known after it.
        if isinstance(node, StatementNode) and node.single_line_if_branches:
            then = next((k for k, tok in enumerate(toks) if token_text(tok) == "then"), -1)
            _check_members(node.span, toks[1 : then if then > 0 else len(toks)], states, push)
            # JavaScript's `toks.slice(1, then)`, a negative `then` counting from the end.
            condition = toks[1:then]
            for branch in node.single_line_if_branches:
                if set_assignment_target(source, branch) is not None:
                    forget(names_in(source, branch))
                    continue
                arm_toks = [
                    tok
                    for tok in statement_tokens_after_leading_label(source, branch)
                    if token_text(tok) != "else"
                ]
                # `If fso.FileExists(p) Then fso.DeleteFile p` deletes only what is
                # there, and leaves no p either way.
                guarded = next(
                    (
                        (lower, state)
                        for lower, state in states.items()
                        if isinstance(state, _FsoState)
                        and _deletes_what_it_tests(condition, arm_toks, lower) is not None
                    ),
                    None,
                )
                if guarded is not None and isinstance(guarded[1], _FsoState):
                    deleted = _deletes_what_it_tests(condition, arm_toks, guarded[0])
                    if deleted is not None:
                        guarded[1].files[deleted] = "absent"
                    continue
                arm = {lower: _copy(state) for lower, state in states.items()}
                _check_members(branch, arm_toks, arm, push)
                for lower, state in list(states.items()):
                    after = arm.get(lower)
                    if isinstance(state, _FsoState) and isinstance(after, _FsoState):
                        # A path the arm may change is unknown after it.
                        for file in {*state.files, *after.files}:
                            if state.files.get(file) != after.files.get(file):
                                state.files.pop(file, None)
                    elif after != state:
                        states.pop(lower, None)
            return
        target = set_assignment_target(source, node.span)
        if target is not None:
            lower = target[0].lower()
            eq = next((k for k, tok in enumerate(toks) if tok.raw_text == "="), -1)
            value = list(toks[eq + 1 :])
            created = _created_object(value, states)
            _check_members(node.span, toks[eq + 1 :], states, push)
            # `Set ts = fso.CreateTextFile(p)` reads fso, which stays followed with
            # what the call did to its files (XLIDE issue #685).
            named = set(names_in(source, node.span))
            if isinstance(created, _TextStreamState):
                named.discard(_lower_name(_at(value, 0)) or "")
            forget(named)
            if created is not None and lower in locals_:
                states[lower] = created
            return
        _check_members(node.span, toks, states, push)

    def snapshot() -> dict[str, _LateObject]:
        return {lower: _copy(state) for lower, state in states.items()}

    def restore(saved: dict[str, _LateObject]) -> None:
        states.clear()
        for lower, state in saved.items():
            states[lower] = _copy(state)

    def touches(stmt: LeafStatementNode) -> set[str]:
        return names_in(source, stmt.span)

    walk_entering_blocks(
        source,
        member.body,
        lambda node: activity is not None and activity.is_inactive(node.span),
        visit,
        BlockEnteringState(snapshot=snapshot, restore=restore, forget=forget, touches=touches),
    )


def _copy(state: _LateObject) -> _LateObject:
    # A document's element tree is never changed in place, so it is shared.
    if isinstance(state, _RegExpState):
        return _RegExpState(state.pattern, dict(state.flags))
    if isinstance(state, _FsoState):
        return _FsoState(dict(state.files))
    return dataclasses.replace(state)


_DOMDOCUMENT_RE = re.compile(r"^msxml2\.domdocument(\.[0-9]+\.[0-9]+)?\Z")


def _created_object(value: Sequence[VbaToken], states: Mapping[str, _LateObject]) -> _LateObject | None:
    """What `CreateObject("...")` or `fso.OpenTextFile(...)` gives, to be followed
    under the Set's target."""
    prog_id = _create_object_prog_id(value)
    if prog_id is not None:
        id_ = prog_id.lower()
        if id_ == "vbscript.regexp":
            return _RegExpState("", {"global": False, "ignorecase": False, "multiline": False})
        if id_ == "scripting.filesystemobject":
            return _FsoState({})
        if id_ == "adodb.recordset":
            return _RecordsetState()
        if _DOMDOCUMENT_RE.match(id_) is not None or id_ == "microsoft.xmldom":
            return _DomDocState(None)
        return None
    # `fso.CreateTextFile(p)` writes; `fso.OpenTextFile(p, mode)` reads with mode 1
    # or none, and writes with 2 or 8.
    owner = _lower_name(_at(value, 0))
    method = token_text(_at(value, 2))
    fso = states.get(owner) if owner else None
    if (
        not isinstance(fso, _FsoState)
        or _raw(value, 1) != "."
        or _raw(value, 3) != "("
        or match_paren_from(value, 3) != len(value) - 1
    ):
        return None
    file = _path_key(split_top_level_token_groups(value, 4, ",", len(value) - 1)[0])
    # A stream to write leaves the file empty until something is written:
    # `Set ts = fso.CreateTextFile(p): ts.Close`, then reading p raises 62 (XLIDE
    # issue #685, measured in Excel 16.0). Append keeps what it held.
    if method == "createtextfile":
        if file is not None:
            fso.files[file] = "empty"
        return _TextStreamState("write", False, file=file)
    if method == "opentextfile":
        mode = _io_mode(value, 3)
        if mode == 2 and file is not None:
            fso.files[file] = "empty"
        if mode == "none" or mode == 1:
            return _TextStreamState(
                "read", False, empty=file is not None and fso.files.get(file) == "empty"
            )
        if mode == 2 or mode == 8:
            return _TextStreamState("write", False, file=file)
        return None
    return None


def _deletes_what_it_tests(
    condition: Sequence[VbaToken], arm: Sequence[VbaToken], fso: str
) -> str | None:
    """The path `fso.FileExists(p)` tests when the arm is `fso.DeleteFile p` on the same path."""
    tested = (
        _path_key(condition[4 : len(condition) - 1])
        if len(condition) >= 6
        and _lower_name(condition[0]) == fso
        and condition[1].raw_text == "."
        and token_text(condition[2]) == "fileexists"
        and condition[3].raw_text == "("
        and match_paren_from(condition, 3) == len(condition) - 1
        else None
    )
    deleted = (
        _path_key([arm[3]])
        if len(arm) == 4
        and _lower_name(arm[0]) == fso
        and arm[1].raw_text == "."
        and token_text(arm[2]) == "deletefile"
        else None
    )
    return tested if tested is not None and tested == deleted else None


def _path_key(arg: Sequence[VbaToken] | None) -> str | None:
    """The key a path argument is known by: a variable's name, or a literal's text."""
    if arg is None:
        return None
    toks = [tok for tok in arg if tok.kind is not TokenKind.COMMENT]
    if len(toks) != 1:
        return None
    return toks[0].raw_text if toks[0].kind is TokenKind.STRING_LITERAL else _lower_name(toks[0])


_DIGITS_RE = re.compile(r"^[0-9]+\Z")


def _io_mode(toks: Sequence[VbaToken], open_index: int) -> int | str | None:
    """The IOMode an `OpenTextFile(` at `open_index` passes: a literal, "none" when
    omitted, None otherwise."""
    args = split_top_level_token_groups(toks, open_index + 1, ",", match_paren_from(toks, open_index))
    arg = [tok for tok in args[1] if tok.kind is not TokenKind.COMMENT] if len(args) > 1 else None
    if not arg:
        return "none"
    if (
        len(arg) == 1
        and arg[0].kind is TokenKind.INTEGER_LITERAL
        and _DIGITS_RE.match(arg[0].raw_text) is not None
    ):
        return int(arg[0].raw_text)
    return None


def _create_object_prog_id(value: Sequence[VbaToken]) -> str | None:
    """The ProgID of a whole `CreateObject("...")` value."""
    at = 2 if token_text(_at(value, 0)) == "vba" and _raw(value, 1) == "." else 0
    literal = _at(value, at + 2)
    if (
        token_text(_at(value, at)) != "createobject"
        or _raw(value, at + 1) != "("
        or literal is None
        or literal.kind is not TokenKind.STRING_LITERAL
        or _raw(value, at + 3) not in (")", ",")
        or match_paren_from(value, at + 1) != len(value) - 1
    ):
        return None
    return string_literal_value(literal.raw_text)


def _check_members(
    base: Span, toks: Sequence[VbaToken], states: dict[str, _LateObject], push: PushFn
) -> None:
    """Each `x.Member` on a followed object; any other mention of x ends what is known of it."""

    def at(tok: VbaToken) -> Span:
        return Span(base.start + tok.start, base.start + tok.end)

    ended: dict[str, None] = {}
    for i in range(len(toks)):
        lower = _lower_name(toks[i])
        state = states.get(lower) if lower else None
        if state is None or lower is None or _raw(toks, i - 1) in (".", "!"):
            continue
        if _raw(toks, i + 1) != "." or not token_name(_at(toks, i + 2)):
            ended[lower] = None
            continue
        member_tok = toks[i + 2]
        member_name = token_text(member_tok)
        assigned = list(toks[4:]) if i == 0 and _raw(toks, 3) == "=" else None
        if isinstance(state, _TextStreamState):
            if state.closed and member_name != "close":
                push(
                    "objectVariableNotSet",
                    f"'{toks[i].raw_text}' was closed, so its {member_tok.raw_text} has no stream "
                    "to reach. This will raise Run-time error '91': Object variable or With block "
                    "variable not set.",
                    at(member_tok),
                )
                ended[lower] = None
            elif (state.mode == "write" and member_name in _READS) or (
                state.mode == "read" and member_name in _WRITES
            ):
                push(
                    "fileModeMismatch",
                    f"'{toks[i].raw_text}' was opened to {'read' if state.mode == 'read' else 'write'}, "
                    f"so {member_tok.raw_text} cannot use it. This will raise Run-time error '54': "
                    "Bad file mode.",
                    at(member_tok),
                )
            elif state.empty and member_name in _READS:
                push(
                    "fileModeMismatch",
                    f"'{toks[i].raw_text}' reads a file this procedure left empty, so "
                    f"{member_tok.raw_text} reads past its end. This will raise Run-time error "
                    "'62': Input past end of file.",
                    at(member_tok),
                )
            elif member_name in _WRITES and state.file is not None:
                for other in states.values():
                    if isinstance(other, _FsoState):
                        other.files[state.file] = "full"
            elif member_name == "close":
                state.closed = True
            continue
        if isinstance(state, _RecordsetState):
            # Open, then Close, leaves it closed again: a second Close raises 3704
            # too (XLIDE issue #685, measured in Excel 16.0).
            if member_name == "open":
                state.open = True
            elif state.open and member_name == "close":
                state.open = False
                state.closed = True
            elif not state.open and member_name in _CLOSED_RECORDSET:
                how = "was closed above" if state.closed else "was never opened"
                push(
                    "lateBoundObjectState",
                    f"'{toks[i].raw_text}' {how}, so {member_tok.raw_text} has no records to work "
                    "on. This will raise Run-time error '3704': Operation is not allowed when the "
                    "object is closed.",
                    at(member_tok),
                )
            continue
        if isinstance(state, _FsoState):
            _file_operation(toks, i, state, at, push)
            continue
        if isinstance(state, _DomDocState):

            def end_document(name: str = lower) -> None:
                ended[name] = None

            _check_document(toks, i, state, at, push, end_document)
            continue
        # A RegExp. Its members, patterns and flags are runtime-member-not-found's
        # and runtime-argument-value's; this follows what Execute finds.
        if member_name not in _REGEXP_MEMBERS:
            continue
        if assigned is not None:
            value = assigned[0] if len(assigned) == 1 else None
            if member_name == "pattern":
                state.pattern = (
                    string_literal_value(value.raw_text)
                    if value is not None and value.kind is TokenKind.STRING_LITERAL
                    else None
                )
            elif member_name in _REGEXP_FLAGS:
                text = (
                    js_trim(string_literal_value(value.raw_text)).lower()
                    if value is not None and value.kind is TokenKind.STRING_LITERAL
                    else None
                )
                word = token_text(value) if value is not None else ""
                state.flags[member_name] = (
                    True
                    if word == "true" or text == "true"
                    else False
                    if word == "false" or text == "false"
                    else None
                )
            break
        if member_name in ("test", "execute", "replace") and _raw(toks, i + 3) == "(":
            if state.pattern is None or vbscript_pattern_error(state.pattern) is not None:
                continue
            close = match_paren_from(toks, i + 3)
            args = split_top_level_token_groups(toks, i + 4, ",", close)
            # The subject as a literal, or literals and vbCr, vbLf, vbCrLf and vbTab
            # joined by &: `"x" & vbLf & "x"` (XLIDE issue #685).
            subject = (
                fold_string_expression(
                    args[0],
                    StringFoldContext(
                        name_value=lambda tok: _SUBJECT_CONSTANTS.get(token_text(tok)),
                        integer_value=lambda _toks: None,
                    ),
                )
                if member_name == "execute" and len(args) == 1
                else None
            )
            if (
                member_name == "execute"
                and subject is not None
                and not (state.flags.get("multiline") and "\r" in subject)
            ):
                hit = _match_index_fault(state.pattern, state.flags, subject, toks, close)
                if hit is not None:
                    hit_at, message = hit
                    push("collectionIndexOutOfRange", message, at(toks[hit_at]))
    for lower in ended:
        states.pop(lower, None)


def _match_index_fault(
    pattern: str,
    flags: Mapping[str, bool | None],
    subject: str,
    toks: Sequence[VbaToken],
    close: int,
) -> tuple[int, str] | None:
    """`Execute(s)(n)` past the matches, or `(n).SubMatches(k)` past the groups, from
    the token after Execute's `)`. Matched as JavaScript does, which VBScript's
    patterns share for what the checker lets through."""
    index = _literal_index_at(toks, close + 1)
    global_flag = flags.get("global")
    ignorecase = flags.get("ignorecase")
    multiline = flags.get("multiline")
    if index is None or global_flag is None or ignorecase is None or multiline is None:
        return None
    expression = _js_regexp(pattern, ignorecase, multiline)
    if expression is None:
        return None
    matches: list[re.Match[str]] = []
    pos = 0
    while True:
        found = expression.search(subject, pos)
        if found is None:
            break
        if found.group(0) == "":
            return None
        matches.append(found)
        if not global_flag:
            break
        pos = found.end()
    value, end = index
    if value >= len(matches):
        count = (
            "no match"
            if len(matches) == 0
            else f"{len(matches)} match{'' if len(matches) == 1 else 'es'}"
        )
        return (
            close + 2,
            f"The pattern finds {count} in {json.dumps(subject, ensure_ascii=False)}, so match "
            f"{value} is past them. This will raise Run-time error '5': Invalid procedure call or "
            "argument.",
        )
    after = end + 1
    if _raw(toks, after) == "." and token_text(_at(toks, after + 1)) == "submatches":
        group = _literal_index_at(toks, after + 2)
        groups = expression.groups
        if group is not None and group[0] >= groups:
            return (
                after + 3,
                f"The pattern has {groups} group{'' if groups == 1 else 's'}, so "
                f"SubMatches({group[0]}) is past them. This will raise Run-time error '5': "
                "Invalid procedure call or argument.",
            )
    return None


# FileSystemObject members that take a path first and change no file.
_FILE_READS: frozenset[str] = frozenset({
    "fileexists", "folderexists", "getfilename", "getbasename", "getextensionname",
    "getparentfoldername", "getabsolutepathname", "buildpath", "gettempname",
    "getspecialfolder", "drives",
})


def _file_operation(
    toks: Sequence[VbaToken],
    i: int,
    state: _FsoState,
    at: Callable[[VbaToken], Span],
    push: PushFn,
) -> None:
    """What a FileSystemObject call does to the path it names first (XLIDE issue
    #477, measured in Excel 16.0): CreateTextFile with Overwrite False on a file the
    procedure made raises 58; DeleteFile, GetFile and OpenTextFile to read or
    append without Create on a file it deleted raise 53. CreateTextFile, and
    OpenTextFile to write, leave the file empty; DeleteFile leaves none. Any other
    call may change any file."""
    member_tok = toks[i + 2]
    member = token_text(member_tok)
    if member in _FILE_READS:
        return
    if member == "opentextfile" and _raw(toks, i + 3) == "(":
        mode = _io_mode(toks, i + 3)
        if isinstance(mode, int) and mode not in (1, 2, 8):
            arg = split_top_level_token_groups(toks, i + 4, ",", match_paren_from(toks, i + 3))[1]
            push(
                "runtimeArgumentValue",
                f"OpenTextFile's IOMode is 1 to read, 2 to write or 8 to append, and {mode} is "
                "none of them. This will raise Run-time error '5': Invalid procedure call or "
                "argument.",
                at(arg[0]),
            )
    parens = _raw(toks, i + 3) == "("
    end = match_paren_from(toks, i + 3) if parens else len(toks)
    first = i + 4 if parens else i + 3
    args = split_top_level_token_groups(toks, first, ",", end) if end > first else []
    file = _path_key(args[0] if args else None)
    if file is None or member not in ("createtextfile", "opentextfile", "deletefile", "getfile"):
        state.files = {}
        return
    fact = state.files.get(file)
    shown = args[0][0].raw_text

    def literal(k: int) -> str | None:
        if k >= len(args):
            return None
        arg = [tok for tok in args[k] if tok.kind is not TokenKind.COMMENT]
        return (token_text(arg[0]) or arg[0].raw_text) if len(arg) == 1 else None

    if member == "createtextfile":
        if literal(1) == "false" and fact in ("empty", "full"):
            push(
                "runtimeArgumentValue",
                f"{member_tok.raw_text} with Overwrite False finds {shown}, which this procedure "
                "made. This will raise Run-time error '58': File already exists.",
                at(member_tok),
            )
        state.files[file] = "empty"
        return
    if member == "opentextfile":
        mode = _io_mode(toks, i + 3)
        create = literal(2) == "true"
        if fact == "absent" and (mode == "none" or mode == 1 or (mode == 8 and not create)):
            push(
                "runtimeArgumentValue",
                f"{member_tok.raw_text} finds no {shown}, which this procedure deleted. This will "
                "raise Run-time error '53': File not found.",
                at(member_tok),
            )
        if mode == 2 and (create or fact != "absent"):
            state.files[file] = "empty"
        elif mode == 8 and create and fact == "absent":
            state.files[file] = "empty"
        return
    if fact == "absent":
        push(
            "runtimeArgumentValue",
            f"{member_tok.raw_text} finds no {shown}, which this procedure deleted. This will "
            "raise Run-time error '53': File not found.",
            at(member_tok),
        )
    if member == "deletefile":
        state.files[file] = "absent"


# Members of a DOMDocument that read it and change nothing.
_DOCUMENT_READS: frozenset[str] = frozenset({
    "documentelement", "selectsinglenode", "selectnodes", "xml", "text", "parseerror",
    "getelementsbytagname",
})

_NO_NODE_TEST_RE = re.compile(r"(^|/)\[")


def _check_document(
    toks: Sequence[VbaToken],
    i: int,
    state: _DomDocState,
    at: Callable[[VbaToken], Span],
    push: PushFn,
    end: Callable[[], None],
) -> None:
    """`doc.LoadXML "<a>"` and what reads the document after it: a document with no
    element, new or loaded from malformed XML, has a DocumentElement of Nothing;
    SelectSingleNode of a path no element is on gives Nothing; and a path with a
    step and no node test raises -2147467259 (XLIDE issue #477, measured in Excel
    16.0). Anything else that may change the document ends what is known of it."""
    member_tok = toks[i + 2]
    member_name = token_text(member_tok)
    if member_name == "loadxml":
        arg: VbaToken | None
        if _raw(toks, i + 3) == "(" and _raw(toks, i + 5) == ")":
            arg = toks[i + 4]
        elif i == 0 and len(toks) == 4:
            arg = toks[3]
        else:
            arg = None
        parsed = (
            parse_xml(string_literal_value(arg.raw_text))
            if arg is not None and arg.kind is TokenKind.STRING_LITERAL
            else _NOT_JUDGED
        )
        if parsed is _NOT_JUDGED:
            end()
        else:
            state.root = parsed if isinstance(parsed, XmlElement) else None
        return
    if member_name not in _DOCUMENT_READS:
        end()
        return

    def reached(after: int) -> VbaToken | None:
        return toks[after + 1] if _raw(toks, after) == "." and token_name(_at(toks, after + 1)) else None

    if member_name == "documentelement" and state.root is None:
        following = reached(i + 3)
        if following is not None:
            push(
                "objectVariableNotSet",
                f"'{toks[i].raw_text}' holds no element, as it is new or was loaded from malformed "
                f"XML, so its DocumentElement is Nothing and '.{following.raw_text}' has no object "
                "to reach. This will raise Run-time error '91': Object variable or With block "
                "variable not set.",
                at(following),
            )
        return
    path_tok = _at(toks, i + 4)
    if (
        member_name in ("selectsinglenode", "selectnodes")
        and _raw(toks, i + 3) == "("
        and path_tok is not None
        and path_tok.kind is TokenKind.STRING_LITERAL
        and _raw(toks, i + 5) == ")"
    ):
        path = string_literal_value(path_tok.raw_text)
        if _NO_NODE_TEST_RE.search(path) is not None:
            push(
                "runtimeArgumentValue",
                f"The XPath '{path}' has a step with no node test before its '['. This will raise "
                "Run-time error '-2147467259': NodeTest expected here.",
                at(path_tok),
            )
            return
        following = reached(i + 6) if member_name == "selectsinglenode" else None
        if following is not None and (state.root is None or _path_finds(state.root, path) is False):
            push(
                "objectVariableNotSet",
                f"No element of '{toks[i].raw_text}' is on the path '{path}', so SelectSingleNode "
                f"gives Nothing and '.{following.raw_text}' has no object to reach. This will "
                "raise Run-time error '91': Object variable or With block variable not set.",
                at(following),
            )


_JS_SPACE_CHARS = re.escape(JS_WHITESPACE)
_XPATH_NAME = r"[A-Za-z_][A-Za-z0-9_.\-]*"
_XPATH_STEP = (
    rf"{_XPATH_NAME}(?:\[(?:[0-9]+|@{_XPATH_NAME}[{_JS_SPACE_CHARS}]*=[{_JS_SPACE_CHARS}]*"
    r"""(?:'[^']*'|"[^"]*"))\])?"""
)
_XPATH_DESCENDANT_RE = re.compile(rf"^//{_XPATH_STEP}\Z")
_XPATH_PATH_RE = re.compile(rf"^/?{_XPATH_STEP}(/{_XPATH_STEP})*\Z")
_XPATH_STEP_RE = re.compile(_XPATH_STEP)
_XPATH_PARTS_RE = re.compile(
    rf"^([^\[]+)(?:\[(?:([0-9]+)|@([^={_JS_SPACE_CHARS}]+)[{_JS_SPACE_CHARS}]*=[{_JS_SPACE_CHARS}]*"
    r"""(?:'([^']*)'|"([^"]*)"))\])?\Z"""
)


def _path_finds(root: XmlElement, path: str) -> bool | None:
    """Whether an element is on `//name`, `/a/b` or `a/b`, each step with an
    optional `[n]` or `[@attr='value']` predicate: `//b[3]`, the third b of some
    parent, and `//b[@id='2']` (XLIDE issue #685, measured in Excel 16.0 with
    MSXML 6). None for any other path. Names match as written."""
    descendant = _XPATH_DESCENDANT_RE.match(path) is not None
    if not descendant and _XPATH_PATH_RE.match(path) is None:
        return None

    # What a step keeps of one parent's children: None where an attribute it reads
    # is not known.
    def take(siblings: Sequence[XmlElement], text: str) -> list[XmlElement] | None:
        parts = _XPATH_PARTS_RE.match(text)
        if parts is None:
            return []
        named = [element for element in siblings if element.name == parts.group(1)]
        if parts.group(2) is not None:
            k = int(parts.group(2)) - 1
            return [named[k]] if 0 <= k < len(named) else []
        attr = parts.group(3)
        if attr is None:
            return named
        wanted = parts.group(4) if parts.group(4) is not None else parts.group(5)
        if any(
            element.attributes is None
            or (attr in element.attributes and element.attributes[attr] is None)
            for element in named
        ):
            return None
        return [
            element
            for element in named
            if element.attributes is not None and element.attributes.get(attr) == wanted
        ]

    if descendant:
        # Every parent's children, the document's own (the root) included.
        groups: list[Sequence[XmlElement]] = [(root,)]
        pending: list[XmlElement] = [root]
        while pending:
            element = pending.pop()
            groups.append(element.children)
            pending.extend(reversed(element.children))
        unknown = False
        for group in groups:
            kept = take(group, path[2:])
            if kept is None:
                unknown = True
            elif len(kept) > 0:
                return True
        return None if unknown else False
    steps = _XPATH_STEP_RE.findall(path[1:] if path.startswith("/") else path)
    level: list[Sequence[XmlElement]] = [(root,)]
    for text in steps:
        found: list[XmlElement] = []
        for group in level:
            kept = take(group, text)
            if kept is None:
                return None
            found.extend(kept)
        if len(found) == 0:
            return False
        level = [element.children for element in found]
    return True


class _NotJudged:
    """parse_xml's "this does not judge it" (upstream's undefined)."""


_NOT_JUDGED = _NotJudged()

_XML_NAME_RE = re.compile(r"[A-Za-z_:][A-Za-z0-9_.:\-]*")
_ENTITY_FAULT_RE = re.compile(r"&(?!(amp|lt|gt|quot|apos|#[0-9]+|#x[0-9A-Fa-f]+);)")


def parse_xml(text: str) -> XmlElement | None | _NotJudged:
    """The document element of `text` as XML: None where it is malformed, and
    _NOT_JUDGED where this does not judge it (a DOCTYPE). Read strictly, as MSXML's
    LoadXML does: quoted attributes, one root, entities only as `&amp;`, `&lt;`,
    `&gt;`, `&quot;`, `&apos;` or a character reference."""
    i = 0
    length = len(text)

    def char(k: int) -> str | None:
        return text[k] if 0 <= k < length else None

    def space() -> None:
        nonlocal i
        while i < length and text[i] in JS_WHITESPACE:
            i += 1

    def read_name() -> str | None:
        nonlocal i
        found = _XML_NAME_RE.match(text, i)
        if found is None:
            return None
        i += len(found.group(0))
        return found.group(0)

    def skip_misc() -> bool | None:
        nonlocal i
        while True:
            space()
            if text.startswith("<!--", i):
                close = text.find("-->", i + 4)
                if close < 0:
                    return False
                i = close + 3
            elif text.startswith("<?", i):
                close = text.find("?>", i + 2)
                if close < 0:
                    return False
                i = close + 2
            elif text.startswith("<!DOCTYPE", i):
                return None
            else:
                return True

    def element() -> XmlElement | None:
        nonlocal i
        if char(i) != "<":
            return None
        i += 1
        tag = read_name()
        if not tag:
            return None
        seen: set[str] = set()
        attributes: dict[str, str | None] = {}
        while True:
            before = i
            space()
            if text.startswith("/>", i):
                i += 2
                return XmlElement(tag, (), attributes)
            if char(i) == ">":
                i += 1
                break
            if i == before:
                return None
            attr = read_name()
            space()
            if not attr or attr in seen or char(i) != "=":
                return None
            seen.add(attr)
            i += 1
            space()
            quote = char(i)
            if quote != '"' and quote != "'":
                return None
            close = text.find(quote, i + 1)
            if close < 0 or "<" in text[i + 1 : close] or not _entities_ok(text[i + 1 : close]):
                return None
            raw = text[i + 1 : close]
            attributes[attr] = None if "&" in raw else raw
            i = close + 1
        children: list[XmlElement] = []
        while True:
            lt = text.find("<", i)
            if lt < 0 or not _entities_ok(text[i:lt]):
                return None
            i = lt
            if text.startswith("</", i):
                i += 2
                closing = read_name()
                space()
                if closing != tag or char(i) != ">":
                    return None
                i += 1
                return XmlElement(tag, tuple(children), attributes)
            if text.startswith("<!--", i):
                close = text.find("-->", i + 4)
                if close < 0:
                    return None
                i = close + 3
            elif text.startswith("<![CDATA[", i):
                close = text.find("]]>", i + 9)
                if close < 0:
                    return None
                i = close + 3
            elif text.startswith("<?", i):
                close = text.find("?>", i + 2)
                if close < 0:
                    return None
                i = close + 2
            else:
                child = element()
                if child is None:
                    return None
                children.append(child)

    before = skip_misc()
    if before is not True:
        return _NOT_JUDGED if before is None else None
    root = element()
    if root is None:
        return None
    after = skip_misc()
    if after is None:
        return _NOT_JUDGED
    return root if after and i == length else None


def _entities_ok(text: str) -> bool:
    """Whether every `&` in XML text starts an entity XML defines."""
    return _ENTITY_FAULT_RE.search(text) is None


def _literal_index_at(toks: Sequence[VbaToken], open_index: int) -> tuple[int, int] | None:
    """A `(n)` with n a whole literal at `open_index`: (n, the index of its `)`)."""
    literal = _at(toks, open_index + 1)
    if (
        _raw(toks, open_index) == "("
        and literal is not None
        and literal.kind is TokenKind.INTEGER_LITERAL
        and _DIGITS_RE.match(literal.raw_text) is not None
        and _raw(toks, open_index + 2) == ")"
    ):
        return (int(literal.raw_text), open_index + 2)
    return None


# Matched at the brace's index (upstream's `^` on the rest of the pattern).
_BRACES_RE = re.compile(r"\{([0-9]+)(,([0-9]*))?\}")


def vbscript_pattern_error(pattern: str) -> int | None:
    """The error VBScript's regular expressions raise for a pattern they cannot
    read, or None: 5020 for an open group, 5019 for an open class, 5018 for a
    quantifier with nothing to repeat, 5017 otherwise (XLIDE issue #477, measured
    in Excel 16.0)."""
    depth = 0
    # Whether what came last can take a quantifier, and whether it was one.
    atom = False
    quantified = False
    # A `?` right after a quantifier makes it lazy, once: `a??` runs, `a+??` does not.
    lazyable = False
    length = len(pattern)
    i = 0
    while i < length:
        c = pattern[i]
        if c == "?" and lazyable:
            lazyable = False
            i += 1
            continue
        lazyable = False
        if c == "\\":
            if i + 1 >= length:
                return 5017
            i += 1
            # `\b` and `\B` are anchors, which nothing repeats: `\b*` raises 5018.
            atom = pattern[i] != "b" and pattern[i] != "B"
            quantified = False
            i += 1
            continue
        if c == "[":
            j = i + 1
            if j < length and pattern[j] == "^":
                j += 1
            # A `]` first in the class is one of its characters: `[]a]` runs, and
            # `[^]` is left open (5019).
            if j < length and pattern[j] == "]":
                j += 1
            while j < length and pattern[j] != "]":
                if pattern[j] == "\\":
                    j += 1
                j += 1
            if j >= length:
                return 5019
            i = j
            atom = True
            quantified = False
            i += 1
            continue
        if c == "(":
            if i + 1 < length and pattern[i + 1] == "?":
                following = pattern[i + 2] if i + 2 < length else "x"
                if following not in ":=!":
                    return 5017
                i += 2
            depth += 1
            atom = False
            quantified = False
            i += 1
            continue
        if c == ")":
            if depth == 0:
                return 5017
            depth -= 1
            atom = True
            quantified = False
            i += 1
            continue
        if c == "|":
            atom = False
            quantified = False
            i += 1
            continue
        braces = _BRACES_RE.match(pattern, i) if c == "{" else None
        if c in ("*", "+", "?") or braces is not None:
            if not atom or quantified:
                return 5018
            if braces is not None:
                upper = braces.group(3)
                if upper is not None and upper != "" and int(upper) < int(braces.group(1)):
                    return 5017
                i += len(braces.group(0)) - 1
            quantified = True
            lazyable = True
            i += 1
            continue
        atom = c != "^" and c != "$"
        quantified = False
        i += 1
    return 5020 if depth > 0 else None


# --- JavaScript RegExp semantics in Python's re ---------------------------------

_JS_LINE_TERMINATORS = "\n\r\u2028\u2029"
_JS_WORD = "A-Za-z0-9_"
_JS_CLASS_ESCAPES: dict[str, str] = {"d": "0-9", "w": _JS_WORD, "s": _JS_SPACE_CHARS}
_JS_ATOM_ESCAPES: dict[str, str] = {
    "d": "[0-9]",
    "D": "[^0-9]",
    "w": f"[{_JS_WORD}]",
    "W": f"[^{_JS_WORD}]",
    "s": f"[{_JS_SPACE_CHARS}]",
    "S": f"[^{_JS_SPACE_CHARS}]",
    "b": f"(?:(?<=[{_JS_WORD}])(?![{_JS_WORD}])|(?<![{_JS_WORD}])(?=[{_JS_WORD}]))",
    "B": f"(?:(?<=[{_JS_WORD}])(?=[{_JS_WORD}])|(?<![{_JS_WORD}])(?![{_JS_WORD}]))",
}
_JS_CONTROL_ESCAPES: dict[str, str] = {"n": "\n", "r": "\r", "t": "\t", "f": "\f", "v": "\v"}
_HEX2_RE = re.compile(r"[0-9A-Fa-f]{2}")
_HEX4_RE = re.compile(r"[0-9A-Fa-f]{4}")
_JS_QUANTIFIER_RE = re.compile(r"\{[0-9]+(,[0-9]*)?\}")


def _js_escape_char(pattern: str, i: int, in_class: bool) -> tuple[str, int] | None:
    """The Python for the JavaScript escape whose letter is at pattern[i]: (text,
    index after it), as a literal character where it names one. None where the
    translation is not known."""
    c = pattern[i]
    if c in _JS_CONTROL_ESCAPES:
        return re.escape(_JS_CONTROL_ESCAPES[c]), i + 1
    if c == "0" and not (i + 1 < len(pattern) and pattern[i + 1].isdigit()):
        return re.escape("\0"), i + 1
    if c == "c":
        letter = pattern[i + 1] if i + 1 < len(pattern) else ""
        if letter.isascii() and letter.isalpha():
            return re.escape(chr(ord(letter) % 32)), i + 2
        return re.escape("\\c") if not in_class else re.escape("\\") + "c", i + 1
    if c == "x" and _HEX2_RE.match(pattern, i + 1):
        return re.escape(chr(int(pattern[i + 1 : i + 3], 16))), i + 3
    if c == "u" and _HEX4_RE.match(pattern, i + 1):
        return re.escape(chr(int(pattern[i + 1 : i + 5], 16))), i + 5
    if in_class and c == "b":
        return re.escape("\b"), i + 1
    if c.isdigit():
        return None
    return re.escape(c), i + 1


def _js_regexp(pattern: str, ignorecase: bool, multiline: bool) -> re.Pattern[str] | None:
    """`new RegExp(pattern, 'g' + flags)`, as a Python pattern with JavaScript's
    meaning; None where the pattern does not compile or is not translated."""
    out: list[str] = []
    length = len(pattern)
    i = 0
    while i < length:
        c = pattern[i]
        if c == "\\":
            if i + 1 >= length:
                return None
            n = pattern[i + 1]
            if n in _JS_ATOM_ESCAPES:
                out.append(_JS_ATOM_ESCAPES[n])
                i += 2
                continue
            if n.isdigit() and n != "0":
                # A back reference.
                j = i + 1
                while j < length and pattern[j].isdigit():
                    j += 1
                out.append("\\" + pattern[i + 1 : j])
                i = j
                continue
            escaped = _js_escape_char(pattern, i + 1, False)
            if escaped is None:
                return None
            out.append(escaped[0])
            i = escaped[1]
            continue
        if c == "[":
            j = i + 1
            negated = j < length and pattern[j] == "^"
            if negated:
                j += 1
            if j < length and pattern[j] == "]":
                # `[]` matches nothing and `[^]` anything.
                out.append("[\\s\\S]" if negated else "(?!)")
                i = j + 1
                continue
            body: list[str] = []
            while j < length and pattern[j] != "]":
                ch = pattern[j]
                if ch == "\\":
                    if j + 1 >= length:
                        return None
                    n = pattern[j + 1]
                    if n in _JS_CLASS_ESCAPES:
                        body.append(_JS_CLASS_ESCAPES[n])
                        j += 2
                        continue
                    if n in ("D", "W", "S"):
                        return None
                    escaped = _js_escape_char(pattern, j + 1, True)
                    if escaped is None:
                        return None
                    body.append(escaped[0])
                    j = escaped[1]
                    continue
                body.append(ch if ch == "-" else re.escape(ch))
                j += 1
            if j >= length:
                return None
            out.append(("[^" if negated else "[") + "".join(body) + "]")
            i = j + 1
            continue
        if c == ".":
            out.append(f"[^{_JS_LINE_TERMINATORS}]")
        elif c == "^":
            out.append(f"(?:\\A|(?<=[{_JS_LINE_TERMINATORS}]))" if multiline else "\\A")
        elif c == "$":
            out.append(f"(?=[{_JS_LINE_TERMINATORS}]|\\Z)" if multiline else "\\Z")
        elif c == "(" and pattern.startswith("(?", i):
            out.append(pattern[i : i + 3])
            i += 3
            continue
        elif c == "{":
            quantifier = _JS_QUANTIFIER_RE.match(pattern, i)
            if quantifier is not None:
                out.append(quantifier.group(0))
                i = quantifier.end()
                continue
            out.append("\\{")
        elif c == "}":
            out.append("\\}")
        else:
            out.append(c)
        i += 1
    try:
        return re.compile("".join(out), re.IGNORECASE if ignorecase else 0)
    except re.error:
        return None
