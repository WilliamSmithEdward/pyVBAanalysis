"""Read/write classification for identifier occurrences.

Ported from xlide_vscode/src/analyzer/references/referenceKinds.ts (XLIDE issue #55).

A reference's KIND is decided by the logical statement it sits in, read straight off
the cached token stream: the assignment family writes its target's terminal name,
declarations write the names they introduce, and everything else reads. The rules
are syntactic on purpose, since VBA spells every write it can prove as a statement
shape, and the one gray zone, passing a variable to a ByRef parameter, stays a read:
claiming a write there would need call-site signature resolution and would be wrong
for every ByVal.

`x = x + 1` is therefore TWO references with distinct kinds; `Mid(s, 1, 2) = t` and
`ReDim Preserve a(n)` are the honest readwrites (they modify in place).
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Sequence
from typing import Literal

from ..lexer.token_helpers import token_word
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize_cached

ReferenceKind = Literal["read", "write", "readwrite"]
_Mark = Callable[[VbaToken, ReferenceKind], None]

_DECL_HEADS = frozenset({"dim", "static", "const", "private", "public", "global", "friend"})
_SIGNATURE_HEADS = frozenset({"sub", "function", "property", "declare", "enum", "type", "event"})
_PARAM_MODIFIERS = frozenset({"byval", "byref", "optional", "paramarray"})


def _is_name(token: VbaToken | None) -> bool:
    return token is not None and token.kind in (TokenKind.IDENTIFIER, TokenKind.BRACKETED_IDENTIFIER)


def _at(seg: Sequence[VbaToken], index: int) -> VbaToken | None:
    return seg[index] if 0 <= index < len(seg) else None


def _raw(seg: Sequence[VbaToken], index: int) -> str | None:
    token = _at(seg, index)
    return token.raw_text if token is not None else None


def _past_parens(seg: Sequence[VbaToken], open_index: int) -> int:
    """Index just past a balanced paren group starting at `open_index`, else +1."""
    depth = 0
    for i in range(open_index, len(seg)):
        if seg[i].raw_text == "(":
            depth += 1
        elif seg[i].raw_text == ")":
            depth -= 1
            if depth == 0:
                return i + 1
    return open_index + 1


def _depth_zero_index_of(seg: Sequence[VbaToken], raw: str, start: int) -> int:
    """First index of `raw` at paren depth zero in [start, len(seg)), or -1."""
    depth = 0
    for i in range(start, len(seg)):
        text = seg[i].raw_text
        if text == "(":
            depth += 1
        elif text == ")":
            depth -= 1
        elif depth == 0 and text == raw:
            return i
    return -1


def _mark_assignment_target(seg: Sequence[VbaToken], start: int, eq: int, mark: _Mark) -> None:
    """The assignment-target rule: within [start, eq), the WRITE lands on the
    terminal name of the target chain, the name whose only suffix before the `=` is
    a run of balanced paren groups. `x =`, `x(i) =`, `a.b.c =`, `a(i).b =` and
    With's `.y =` all select exactly one name; every other name on the left
    (receivers, indexes) reads."""
    for k in range(start, eq):
        if not _is_name(seg[k]):
            continue
        j = k + 1
        while j < eq and seg[j].raw_text == "(":
            j = _past_parens(seg, j)
        if j == eq:
            mark(seg[k], "write")


def _classify_signature(seg: Sequence[VbaToken], mark: _Mark) -> None:
    """Sub/Function/Property/Event signatures: the name and its parameters."""
    i = 0
    while i < len(seg) and not _is_name(seg[i]):
        if seg[i].raw_text == "(":
            return
        i += 1
    # Names after the head keywords are the procedure name (Property carries
    # Get/Let/Set first, which are keywords).
    if _is_name(_at(seg, i)):
        mark(seg[i], "write")
        i += 1
    if _raw(seg, i) != "(":
        return
    depth = 0
    expect_param = True
    while i < len(seg):
        raw = seg[i].raw_text
        word = token_word(seg[i])
        i += 1
        if raw == "(":
            depth += 1
            if depth == 1:
                expect_param = True
            continue
        if raw == ")":
            depth -= 1
            if depth == 0:
                break
            continue
        if depth != 1:
            continue
        if raw == ",":
            expect_param = True
            continue
        if word in ("as", "new"):
            expect_param = False
            continue
        if word in _PARAM_MODIFIERS:
            continue
        if raw == "=":  # Optional defaults read
            expect_param = False
            continue
        if expect_param and _is_name(seg[i - 1]):
            mark(seg[i - 1], "write")
            expect_param = False


def _classify_segment(
    seg: Sequence[VbaToken], wanted: frozenset[int], out: dict[int, ReferenceKind]
) -> None:
    if not seg or not any(token.start in wanted for token in seg):
        return

    def mark(token: VbaToken, kind: ReferenceKind) -> None:
        if token.start in wanted:
            out[token.start] = kind

    head = 0
    head_word = token_word(seg[head])

    # Set / Let are assignment statements with a keyword prefix.
    if head_word in ("set", "let", "lset", "rset"):
        eq = _depth_zero_index_of(seg, "=", head + 1)
        if eq > 0:
            _mark_assignment_target(seg, head + 1, eq, mark)
        return

    if head_word == "for":
        if token_word(_at(seg, 1)) == "each":
            if _is_name(_at(seg, 2)):
                mark(seg[2], "write")
        elif _is_name(_at(seg, 1)):
            mark(seg[1], "write")
        return  # the bounds and the collection read

    if head_word == "redim":
        i = 1
        kind: ReferenceKind = "write"
        if token_word(_at(seg, i)) == "preserve":
            kind = "readwrite"
            i += 1
        expect_target = True
        depth = 0
        while i < len(seg):
            raw = seg[i].raw_text
            if raw == "(":
                depth += 1
            elif raw == ")":
                depth -= 1
            elif depth > 0:
                pass
            elif raw == ",":
                expect_target = True
            elif token_word(seg[i]) == "as":
                expect_target = False
            elif expect_target and _is_name(seg[i]):
                mark(seg[i], kind)
                expect_target = False
            i += 1
        return

    if head_word == "erase":
        depth = 0
        for i in range(1, len(seg)):
            raw = seg[i].raw_text
            if raw == "(":
                depth += 1
            elif raw == ")":
                depth -= 1
            elif depth == 0 and _is_name(seg[i]):
                mark(seg[i], "write")
        return

    # Mid/MidB statements modify their first argument in place. The lexer gives a
    # type character a token of its own, so `Mid$(` is `Mid`, `$`, `(` (XLIDE #80).
    if head_word in ("mid", "midb"):
        open_index = 2 if _raw(seg, 1) == "$" and seg[1].start == seg[0].end else 1
        if _raw(seg, open_index) == "(":
            close = _past_parens(seg, open_index)
            if _raw(seg, close) == "=":
                for i in range(open_index + 1, close):
                    if _is_name(seg[i]):
                        mark(seg[i], "readwrite")
                        break
            return

    # Input #f, a, b / Line Input #f, s / Get #f, pos, var fill their trailing
    # variables.
    is_line_input = head_word == "line" and token_word(_at(seg, 1)) == "input"
    if (head_word in ("input", "get") or is_line_input) and any(t.raw_text == "#" for t in seg):
        commas_needed = 2 if head_word == "get" else 1
        commas = 0
        depth = 0
        for i in range(1, len(seg)):
            raw = seg[i].raw_text
            if raw == "(":
                depth += 1
            elif raw == ")":
                depth -= 1
            elif depth == 0 and raw == ",":
                commas += 1
            elif (
                depth == 0
                and commas >= commas_needed
                and _is_name(seg[i])
                and _raw(seg, i - 1) != "."
            ):
                mark(seg[i], "write")
        return

    if head_word in _DECL_HEADS:
        i = head + 1
        if token_word(_at(seg, i)) == "const" or head_word == "const":
            # Const A = 1, B = 2: names write, initializers read.
            if head_word != "const":
                i += 1
            expect_name = True
            depth = 0
            while i < len(seg):
                raw = seg[i].raw_text
                if raw == "(":
                    depth += 1
                elif raw == ")":
                    depth -= 1
                elif depth == 0 and raw == ",":
                    expect_name = True
                elif depth == 0 and raw == "=":
                    expect_name = False
                elif depth == 0 and expect_name and _is_name(seg[i]):
                    mark(seg[i], "write")
                    expect_name = False
                i += 1
            return
        if token_word(_at(seg, i)) in _SIGNATURE_HEADS or head_word in _SIGNATURE_HEADS:
            _classify_signature(seg, mark)
            return
        # Dim x As Long, y(10) As String, WithEvents app As Excel.Application
        expect_name = True
        depth = 0
        while i < len(seg):
            word = token_word(seg[i])
            raw = seg[i].raw_text
            i += 1
            if raw == "(":
                depth += 1
                continue
            if raw == ")":
                depth -= 1
                continue
            if depth > 0:
                continue
            if raw == ",":
                expect_name = True
                continue
            if word in ("as", "new"):
                expect_name = False
                continue
            if word == "withevents":
                continue
            if expect_name and _is_name(seg[i - 1]):
                mark(seg[i - 1], "write")
                expect_name = False
        return

    if head_word in _SIGNATURE_HEADS:
        _classify_signature(seg, mark)
        return

    # Inline If carries real statements after Then and Else on the same logical
    # line: `If x = 1 Then y = 2 Else z = 3` reads its condition and writes both
    # targets. Each tail classifies as its own statement.
    if head_word in ("if", "elseif", "else"):
        if head_word == "else":
            start = head + 1
        else:
            start = len(seg)
            depth = 0
            for i in range(head + 1, len(seg)):
                raw = seg[i].raw_text
                if raw == "(":
                    depth += 1
                elif raw == ")":
                    depth -= 1
                elif depth == 0 and token_word(seg[i]) == "then":
                    start = i + 1
                    break
        if start < len(seg):
            tail_start = start
            depth = 0
            for i in range(start, len(seg) + 1):
                at_else = i < len(seg) and depth == 0 and token_word(seg[i]) == "else"
                if i == len(seg) or at_else:
                    if i > tail_start:
                        _classify_segment(seg[tail_start:i], wanted, out)
                    tail_start = i + 1
                    continue
                raw = seg[i].raw_text
                if raw == "(":
                    depth += 1
                elif raw == ")":
                    depth -= 1
        return

    # A plain assignment starts with an expression: a name, Me, or With's leading
    # dot. Anything keyword-led (If, While, Call, Debug...) reads.
    starts_expression = _is_name(seg[head]) or head_word == "me" or seg[head].raw_text == "."
    if starts_expression:
        eq = _depth_zero_index_of(seg, "=", head)
        if eq > 0 and seg[eq].kind is TokenKind.OPERATOR and seg[eq].raw_text == "=":
            _mark_assignment_target(seg, head, eq, mark)


def classify_reference_kinds(source: str, offsets: Iterable[int]) -> dict[int, ReferenceKind]:
    """Classify the identifier occurrences whose ABSOLUTE offsets are given.

    Unmatched offsets come back 'read', the honest default for every position that
    is not provably a write.
    """
    wanted = frozenset(offsets)
    out: dict[int, ReferenceKind] = {}
    if not wanted:
        return out
    seg: list[VbaToken] = []
    for token in tokenize_cached(source):
        if token.kind is TokenKind.NEWLINE or token.kind is TokenKind.COLON:
            _classify_segment(seg, wanted, out)
            seg = []
            continue
        if token.kind is TokenKind.COMMENT:
            continue
        seg.append(token)
    _classify_segment(seg, wanted, out)
    for offset in wanted:
        out.setdefault(offset, "read")
    return out
