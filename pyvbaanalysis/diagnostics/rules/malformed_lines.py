"""Rule family: lines the VBE cannot parse, as typing leaves them (XLIDE issue #234).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/malformedLines.ts.
Measured in Excel 16.0 (build 20326, 2026-09-30) with a full compile.

- malformed-statement:
    `#asdf`, a directive that is not #If, #ElseIf, #Else, #End If or
    #Const -> "Syntax error" in a procedure and after one, "Expected: If
    or Else or ElseIf or End or EndIf or Const" above the first.
    `[asdf`, a bracket never closed -> "Syntax error", "Missing end
    bracket" in an Enum.
    `Sub` or `Function` with no name -> "Expected: identifier" above the
    first procedure, "Syntax error" after one.
    `Private Sub S(a b c)`, `S(x Property)`: a parameter followed by
    another word -> "Expected: list separator or )".
    An Enum line that is not `name [= value]`: `asdf qwer`, `.asdf` ->
    "Invalid inside Enum"; `asdf & qwer` -> "Expected: expression".
- reserved-keyword-in-expression: a statement keyword where a value
    goes: `Array(1, Const, 3)`, `Debug.Print RaiseEvent`, `Case Open`,
    `Dim a(Implements) As Long` -> "Syntax error"; `Const K = Dim` ->
    "Expected: expression".
- statement-outside-procedure: a lone `:` after a procedure -> "Only
    comments may appear after End Sub, End Function, or End Property". It
    compiles above the first procedure.

None of it is judged in an inactive #If branch, which the VBE does not
parse.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from typing import Literal, Protocol

from ...conditional import ConditionalActivityTracker
from ...js_compat import js_trim
from ...lexer.token_helpers import first_token_at_or_after
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import EnumNode, ModuleNode, ProcedureNode, Span, VariableGroupNode
from ..context import PushFn
from ..walker import (
    active_module_members,
    bare_assignment_target,
    for_each_statement,
    for_each_variable_group,
    match_paren_from,
    statement_tokens,
    token_text,
)

_DIRECTIVE_WORDS = frozenset({"if", "elseif", "else", "end", "endif", "const"})

# Statement keywords that can never stand where a value goes. `To`, `Is`,
# `Else`, `As` and `New` are left out: each has a place inside a value's
# statement (`Case 1 To 3`, `Case Is > 1`, `Case Else`, `Set x = New C`).
_STATEMENT_KEYWORDS = frozenset(
    {
        "const", "raiseevent", "open", "close", "implements", "dim", "redim", "static", "sub",
        "function", "declare", "enum", "event", "option", "call", "goto", "gosub", "exit",
        "resume", "do", "loop", "wend", "while", "until", "with", "select", "case", "next", "for",
        "each", "then", "elseif", "if", "end", "public", "private", "friend", "global", "let",
        "set", "stop", "return", "lock", "unlock", "preserve", "withevents", "type", "put", "get",
        "kill", "step",
    }
)  # fmt: skip

_PARAMETER_MODIFIERS = frozenset({"optional", "byval", "byref", "paramarray"})

_TYPE_CHARACTER_RE = re.compile(r"^[$%&!#@]$")

_Place = Literal["procedure", "top", "after"]


class _HasSpan(Protocol):
    @property
    def span(self) -> Span: ...


def check_malformed_lines(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    procedures = [m for m in mod.members if isinstance(m, ProcedureNode)]
    named_procedures = [p for p in procedures if p.name != ""]
    first_named = named_procedures[0] if named_procedures else None

    def place(offset: int) -> _Place:
        if _inside_member(named_procedures, offset):
            return "procedure"
        return "top" if first_named is None or offset < first_named.span.start else "after"

    enums = [m for m in mod.members if isinstance(m, EnumNode)]

    def in_enum(offset: int) -> bool:
        return _inside_member(enums, offset)

    def active(span: Span) -> bool:
        return not (activity is not None and activity.is_inactive(span))

    toks = tokenize_cached(source)

    for i, tok in enumerate(toks):
        span = Span(tok.start, tok.end)
        if tok.kind is TokenKind.DIRECTIVE and active(span):
            following = toks[i + 1] if i + 1 < len(toks) else None
            word = (
                following
                if following is not None and following.kind is not TokenKind.NEWLINE
                else None
            )
            if word is None or token_text(word) not in _DIRECTIVE_WORDS:
                error = (
                    "Expected: If or Else or ElseIf or End or EndIf or Const"
                    if place(tok.start) == "top"
                    else "Syntax error"
                )
                shown = f"#{word.raw_text}" if word is not None else "#"
                push(
                    "malformedStatement",
                    f"'{shown}' is no compiler directive: only #If, #ElseIf, #Else, #End If and "
                    f"#Const are. This is a VBE compile error: {error}.",
                    Span(tok.start, word.end) if word is not None else span,
                )
        elif (
            tok.kind is TokenKind.BRACKETED_IDENTIFIER
            and not tok.raw_text.endswith("]")
            and active(span)
        ):
            error = "Missing end bracket" if in_enum(tok.start) else "Syntax error"
            push(
                "malformedStatement",
                f"'{tok.raw_text}' opens a bracketed name that never closes. This is a VBE compile "
                f"error: {error}.",
                span,
            )
        elif (
            tok.kind is TokenKind.COLON
            and _line_holds_only_colons(toks, i)
            and place(tok.start) == "after"
            and active(span)
        ):
            push(
                "statementOutsideProcedure",
                "A ':' after a procedure separates nothing. This is a VBE compile error: Only "
                "comments may appear after End Sub, End Function, or End Property.",
                span,
            )

    _check_statement_forms(
        toks, lambda span: active(span) and place(span.start) == "procedure", push
    )

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_procedure_header(toks, member, place, push)

            def visit_statement(stmt: _HasSpan) -> None:
                _check_value_keywords(source, stmt.span, "statement", push)
                _check_keyword_qualifiers(source, stmt.span, push)
                _check_value_words(source, stmt.span, push)

            for_each_statement(member.body, visit_statement, activity)

            def visit_group(group: VariableGroupNode) -> None:
                _check_declaration_keywords(source, group, push)

            for_each_variable_group(member.body, visit_group, activity)
        elif isinstance(member, VariableGroupNode):
            _check_declaration_keywords(source, member, push)
        elif isinstance(member, EnumNode):
            _check_enum_lines(source, member, activity, push)


def _inside_member(members: Sequence[_HasSpan], offset: int) -> bool:
    """Members are source-ordered and non-overlapping; their boundary tokens are outside."""
    lo = 0
    hi = len(members)
    while lo < hi:
        mid = lo + (hi - lo) // 2
        if members[mid].span.start < offset:
            lo = mid + 1
        else:
            hi = mid
    return lo > 0 and offset < members[lo - 1].span.end


# The operators a VB.NET-style compound assignment puts before `=`: `x += 1`.
# The other operators are an operator run, invalid-expression-syntax's.
_COMPOUND_ASSIGNMENT_OPERATORS = frozenset({"+", "-"})


@dataclass(slots=True)
class _SelectState:
    saw_else: bool


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) outside the list."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    """`toks[i]?.rawText`."""
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


def _check_statement_forms(
    toks: Sequence[VbaToken],
    in_procedure: Callable[[Span], bool],
    push: PushFn,
) -> None:
    """Statement forms the VBE refuses (XLIDE issue #369, measured in Excel 16.0):
    `Else If b Then` on its own line, `x += 1`, `Exit While`,
    `Call Debug.Print(...)` and any Call of Debug, a Case after Case Else, a
    Loop with a condition after a Do with one, a Loop with both While and
    Until, and `Case Is` with To or Like. Each statement of each line is read
    in order; `in_procedure` says whether a span is active code inside a
    procedure."""
    selects: list[_SelectState] = []
    dos: list[bool] = []

    def at(first: VbaToken, last: VbaToken) -> Span:
        return Span(first.start, last.end)

    def report(first: VbaToken, last: VbaToken, what: str, error: str) -> None:
        push("malformedStatement", f"{what} This is a VBE compile error: {error}.", at(first, last))

    def check_case_items(items: Sequence[VbaToken]) -> None:
        depth = 0
        start = 0
        for k in range(len(items) + 1):
            tok = _at(items, k)
            if tok is not None and tok.raw_text == "(":
                depth += 1
            elif tok is not None and tok.raw_text == ")":
                depth -= 1
            if tok is not None and (depth > 0 or tok.raw_text != ","):
                continue
            item = items[start:k]
            start = k + 1
            if token_text(_at(item, 0)) != "is":
                continue
            if token_text(_at(item, 1)) == "like":
                report(
                    item[0],
                    item[1],
                    "Case Is takes a comparison operator, and Like is none.",
                    "Syntax error",
                )
            else:
                to = next((t for t in item if token_text(t) == "to"), None)
                if to is not None:
                    report(
                        item[0],
                        to,
                        "Case Is takes one value: a range with To is a Case of its own.",
                        "Syntax error",
                    )

    def check_statement(words: Sequence[VbaToken]) -> None:
        head = token_text(words[0])
        for k in range(len(words) - 1):
            if token_text(words[k]) == "exit" and token_text(words[k + 1]) == "while":
                report(
                    words[k],
                    words[k + 1],
                    "'Exit While' is no statement: a While loop has no exit of its own, so use Do "
                    "While ... Loop and Exit Do.",
                    "Syntax error",
                )
            if (
                token_text(words[k]) == "call"
                and token_text(words[k + 1]) == "debug"
                and _raw_at(words, k + 2) == "."
                and _at(words, k + 3) is not None
            ):
                report(
                    words[k],
                    words[k + 3],
                    f"Call cannot run Debug.{words[k + 3].raw_text}: write it without Call.",
                    "Syntax error",
                )
        # `x += 1`: an operator glued to the assignment's `=`.
        if words[0].kind is TokenKind.IDENTIFIER:
            eq = next((k for k, tok in enumerate(words) if tok.raw_text == "="), -1)
            if eq > 1:
                op = words[eq - 1]
                if (
                    op.end == words[eq].start
                    and op.raw_text in _COMPOUND_ASSIGNMENT_OPERATORS
                    and all(
                        tok.raw_text == "." or tok.kind is TokenKind.IDENTIFIER
                        for tok in words[1 : eq - 1]
                    )
                ):
                    target = "".join(tok.raw_text for tok in words[0 : eq - 1])
                    report(
                        op,
                        words[eq],
                        f"'{op.raw_text}=' is no VBA operator: write {target} = {target} "
                        f"{op.raw_text} ...",
                        "Syntax error",
                    )
        if head == "select" and token_text(_at(words, 1)) == "case":
            selects.append(_SelectState(saw_else=False))
        elif head == "end" and token_text(_at(words, 1)) == "select":
            if selects:
                selects.pop()
        elif head == "case":
            select = selects[-1] if selects else None
            if token_text(_at(words, 1)) == "else":
                if select is not None:
                    select.saw_else = True
            else:
                if select is not None and select.saw_else:
                    report(
                        words[0],
                        words[-1],
                        "A Case after Case Else in the same Select Case can never run.",
                        "Case without Select Case",
                    )
                check_case_items(words[1:])
        elif head == "do":
            dos.append(
                len(words) > 1
                and (token_text(words[1]) == "while" or token_text(words[1]) == "until")
            )
        elif head == "loop":
            conditioned = dos.pop() if dos else None
            kinds = [
                tok for tok in words[1:] if token_text(tok) == "while" or token_text(tok) == "until"
            ]
            if len(kinds) > 1:
                report(
                    kinds[0],
                    kinds[1],
                    "A Loop takes one condition, While or Until, not both.",
                    "Syntax error",
                )
            elif len(kinds) == 1 and conditioned:
                report(
                    words[0],
                    kinds[0],
                    "A Loop with a condition closes a Do with none; this Do has its own.",
                    "Loop without Do",
                )

    line: list[VbaToken] = []

    def flush_line() -> None:
        nonlocal line
        words = [tok for tok in line if tok.kind is not TokenKind.COMMENT]
        line = []
        if len(words) == 0 or not in_procedure(at(words[0], words[-1])):
            return
        # `Else If b Then` with nothing after Then: Else and a block If header.
        if (
            token_text(words[0]) == "else"
            and token_text(_at(words, 1)) == "if"
            and token_text(words[-1]) == "then"
        ):
            report(
                words[0],
                words[-1],
                "'Else If ... Then' starts a block If on the Else line; write ElseIf, one word.",
                "Syntax error",
            )
        for statement in _split_statements(words):
            check_statement(statement)

    for tok in toks:
        if tok.kind is TokenKind.NEWLINE:
            flush_line()
        else:
            line.append(tok)
    flush_line()


def _split_statements(words: Sequence[VbaToken]) -> list[list[VbaToken]]:
    """A line's statements, split at top-level colons; a label's colon ends the label."""
    out: list[list[VbaToken]] = [[]]
    for tok in words:
        if tok.kind is TokenKind.COLON:
            out.append([])
        else:
            out[-1].append(tok)
    return [statement for statement in out if len(statement) > 0]


def _line_holds_only_colons(toks: Sequence[VbaToken], index: int) -> bool:
    """True when the colon at `index` stands on a line with nothing else but colons and a comment."""
    j = index - 1
    while j >= 0 and toks[j].kind is not TokenKind.NEWLINE:
        if toks[j].kind is not TokenKind.COLON:
            return False
        j -= 1
    j = index + 1
    while j < len(toks) and toks[j].kind is not TokenKind.NEWLINE:
        if toks[j].kind is not TokenKind.COLON and toks[j].kind is not TokenKind.COMMENT:
            return False
        j += 1
    return True


def _check_procedure_header(
    toks: Sequence[VbaToken],
    proc: ProcedureNode,
    place: Callable[[int], _Place],
    push: PushFn,
) -> None:
    """`Sub` with no name, and a parameter followed by another word."""
    i = first_token_at_or_after(toks, proc.span.start)
    if i == len(toks):
        return
    header: list[VbaToken] = []
    while (
        i < len(toks)
        and toks[i].kind is not TokenKind.NEWLINE
        and toks[i].kind is not TokenKind.COMMENT
    ):
        header.append(toks[i])
        i += 1
    if proc.name == "":
        keyword = next(
            (tok for tok in header if token_text(tok) in ("sub", "function", "property")), None
        )
        if keyword is not None:
            error = "Expected: identifier" if place(proc.span.start) == "top" else "Syntax error"
            push(
                "malformedStatement",
                f"'{keyword.raw_text}' needs a name after it. This is a VBE compile error: "
                f"{error}.",
                Span(keyword.start, keyword.end),
            )
        return
    name_span = proc.name_span
    name_at = next(
        (
            k
            for k, tok in enumerate(header)
            if name_span is not None and tok.start == name_span.start
        ),
        -1,
    )
    open_ = name_at + 1 if name_at >= 0 and _raw_at(header, name_at + 1) == "(" else -1
    if open_ < 0:
        return
    close = match_paren_from(header, open_)
    if close < 0:
        return
    from_ = open_ + 1
    for k in range(open_ + 1, close + 1):
        if k < close and header[k].raw_text != ",":
            continue
        junk = _parameter_junk(header[from_:k])
        if junk is not None:
            push(
                "malformedStatement",
                f"'{junk.raw_text}' follows a parameter's name where only As, = or a comma can. "
                "This is a VBE compile error: Expected: list separator or ).",
                Span(junk.start, junk.end),
            )
            return
        from_ = k + 1


def _parameter_junk(param: Sequence[VbaToken]) -> VbaToken | None:
    """The token after a parameter's name that no parameter form allows there."""
    i = 0
    while i < len(param) and token_text(param[i]) in _PARAMETER_MODIFIERS:
        i += 1
    name = _at(param, i)
    if name is None:
        return None
    i += 1
    # A type character glued to the name, `s$`, and an array's `()`.
    after = _at(param, i)
    if after is not None and after.start == name.end and _TYPE_CHARACTER_RE.search(after.raw_text):
        i += 1
    if _raw_at(param, i) == "(" and _raw_at(param, i + 1) == ")":
        i += 2
    nxt = _at(param, i)
    if (
        nxt is None
        or nxt.kind is TokenKind.COMMENT
        or token_text(nxt) == "as"
        or nxt.raw_text == "="
    ):
        return None
    return nxt


def _check_value_keywords(
    source: str, span: Span, context: Literal["statement", "const"], push: PushFn
) -> None:
    """A statement keyword where a value goes: inside parentheses, after an
    assignment's `=`, after Print, and after Case."""
    toks = statement_tokens(source, span)
    print_ = next((k for k, tok in enumerate(toks) if token_text(tok) == "print"), -1)
    if context == "const":
        value_from = 0
    elif token_text(_at(toks, 0)) == "case":
        value_from = 1
    elif print_ >= 0:
        value_from = print_ + 1
    else:
        value_from = _value_start(source, span, toks)
    depth = 0
    for i, tok in enumerate(toks):
        raw = tok.raw_text
        if raw == "(":
            depth += 1
            continue
        if raw == ")":
            depth -= 1
            continue
        if tok.kind is not TokenKind.KEYWORD or token_text(tok) not in _STATEMENT_KEYWORDS:
            continue
        if depth == 0 and (value_from < 0 or i < value_from):
            continue
        # `.Open`, `!Close` and `Type:=1` name a member or an argument.
        prev = _raw_at(toks, i - 1)
        if prev == "." or prev == "!" or _raw_at(toks, i + 1) == ":=":
            continue
        error = "Expected: expression" if context == "const" else "Syntax error"
        push(
            "reservedKeywordInExpression",
            f"'{tok.raw_text}' is a statement keyword and cannot stand where a value goes. This "
            f"is a VBE compile error: {error}.",
            Span(span.start + tok.start, span.start + tok.end),
        )
        return


# Words with a statement's or a print list's meaning only, named where a value
# goes (XLIDE issue #318, measured in Excel 16.0): `TypeName(Tab)`, `Main = Print`,
# `1 + Shared` are a Syntax error, and so are `Tab(5)` and `Spc(5)` outside a
# Print list. Input, Len and Array take an argument list, and Seek
# needs its argument: bare, the first three are a Syntax error and
# Seek is "Argument not optional".
VALUE_WORD_ERRORS: Mapping[str, str] = {
    "tab": "Syntax error",
    "spc": "Syntax error",
    "print": "Syntax error",
    "write": "Syntax error",
    "shared": "Syntax error",
    "input": "Syntax error",
    "len": "Syntax error",
    "array": "Syntax error",
    "seek": "Argument not optional",
}


def _check_value_words(source: str, span: Span, push: PushFn) -> None:
    toks = statement_tokens(source, span)
    value_from = _value_start(source, span, toks)
    # A Print list, `Debug.Print "a"; Tab(5)` or `Print #1, Spc(2)`, takes Tab and Spc.
    print_ = next(
        (
            k
            for k, tok in enumerate(toks)
            if token_text(tok) == "print" and (k == 0 or toks[k - 1].raw_text == ".")
        ),
        -1,
    )
    depth = 0
    for i, tok in enumerate(toks):
        depth += 1 if tok.raw_text == "(" else -1 if tok.raw_text == ")" else 0
        word = token_text(tok)
        error = VALUE_WORD_ERRORS.get(word)
        nxt = _raw_at(toks, i + 1)
        if (
            not error
            or (depth == 0 and (value_from < 0 or i < value_from))
            or (_raw_at(toks, i - 1) or "") in (".", "!")
            or (nxt or "") in (".", "!", ":=")
        ):
            continue
        if (word == "tab" or word == "spc") and print_ >= 0 and i > print_:
            continue
        if word in ("input", "len", "seek", "array") and (nxt == "(" or nxt == "$"):
            continue
        what = (
            "is a function that needs its file number"
            if word == "seek"
            else "belongs to a Print list"
            if word == "tab" or word == "spc"
            else "is a reserved word"
        )
        push(
            "argumentCount" if word == "seek" else "reservedKeywordInExpression",
            f"'{tok.raw_text}' {what} and cannot stand here as a value. This is a VBE compile "
            f"error: {error}.",
            Span(span.start + tok.start, span.start + tok.end),
        )
        return


# Words that cannot qualify a member: `Main = Print.Hi()` is a Syntax error
# whether or not a module of that name exists (XLIDE issue #247, measured in Excel
# 16.0). Get, Put, Open and Stop are statement keywords, reported above.
_NO_QUALIFIER_WORDS = frozenset(
    {"circle", "pset", "scale", "print", "input", "tab", "spc", "array", "lbound", "date"}
)


def _check_keyword_qualifiers(source: str, span: Span, push: PushFn) -> None:
    toks = statement_tokens(source, span)
    for i in range(len(toks) - 1):
        word = token_text(toks[i])
        prev = _raw_at(toks, i - 1)
        # Opening a statement, `Date.Hi` compiles and raises 424 at run time.
        if (
            word not in _NO_QUALIFIER_WORDS
            or toks[i + 1].raw_text != "."
            or prev == "."
            or prev == "!"
            or (i == 0 and word == "date")
        ):
            continue
        error = (
            "Method not valid without suitable object"
            if i == 0 and word == "print"
            else "Syntax error"
        )
        push(
            "reservedKeywordInExpression",
            f"'{toks[i].raw_text}' is a reserved word and cannot qualify a member, even when a "
            f"module bears the name. This is a VBE compile error: {error}.",
            Span(span.start + toks[i].start, span.start + toks[i].end),
        )
        return


def _value_start(source: str, span: Span, toks: Sequence[VbaToken]) -> int:
    """Where an assignment's value starts, or -1 for any other statement."""
    bare = bare_assignment_target(source, span)
    if bare is None or len(bare[2]) == 0:
        return -1
    # Both are relative to the statement's start.
    first = bare[2][0]
    return next((k for k, tok in enumerate(toks) if tok.start == first.start), -1)


def _check_declaration_keywords(source: str, group: VariableGroupNode, push: PushFn) -> None:
    """Array bounds and a Const's value: `Dim a(Implements) As Long`, `Const K = Dim`."""
    for decl in group.declarations:
        toks = statement_tokens(source, decl.span)
        eq = (
            next((k for k, tok in enumerate(toks) if tok.raw_text == "="), -1)
            if group.is_const
            else -1
        )
        if eq >= 0:
            value = toks[eq + 1 :]
            if len(value) > 0:
                _check_value_keywords(
                    source,
                    Span(decl.span.start + value[0].start, decl.span.end),
                    "const",
                    push,
                )
            continue
        _check_value_keywords(source, decl.span, "statement", push)


def _check_enum_lines(
    source: str,
    member: EnumNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """An Enum line that is not `name [= value]`."""
    for line in member.members:
        if activity is not None and activity.is_inactive(line.span):
            continue
        toks = statement_tokens(source, line.span)
        name = _at(toks, 0)
        if name is None:
            continue
        # `.asdf`. A number, `123 = 2`, is invalid-identifier-start's.
        if name.kind in (TokenKind.PUNCTUATION, TokenKind.OPERATOR, TokenKind.UNKNOWN):
            text = js_trim(source[line.span.start : line.span.end])
            push(
                "malformedStatement",
                f"'{text}' is no Enum member: a member is a name, with = and a value if it has "
                "one. This is a VBE compile error: Invalid inside Enum.",
                line.span,
            )
            continue
        nxt = _at(toks, 1)
        if (
            nxt is None
            or nxt.raw_text == "="
            or (nxt.start == name.end and _TYPE_CHARACTER_RE.search(nxt.raw_text))
        ):
            continue
        operator = nxt.kind is TokenKind.OPERATOR
        push(
            "malformedStatement",
            f"'{nxt.raw_text}' follows Enum member '{name.raw_text}' where only = can. This is a "
            "VBE compile error: Expected: expression."
            if operator
            else f"'{nxt.raw_text}' follows Enum member '{name.raw_text}': a member is a name, "
            "with = and a value if it has one. This is a VBE compile error: Invalid inside Enum.",
            Span(line.span.start + nxt.start, line.span.start + nxt.end),
        )
