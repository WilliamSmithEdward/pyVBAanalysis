"""The expressions inside statements that are not assignments or calls.

Ported from statementOperands.ts. The shared expression walk (exprwalk.py)
parses assignments, calls and If conditions. A Case item, a Print # or
Debug.Print item, a file number, a Seek position, a record number, a Lock
range or a ReDim bound is an expression too, and the VBE judges it as one: an array or a scalar
`Is` there is the same Type mismatch it is anywhere (XLIDE issue #1415,
measured in Excel, Word, PowerPoint and Access 16.0). This module finds those
expressions in a statement's own text, so every expression rule sees them
through the one walk. The arms of a single-line If are statements of their own
and read the same way; its condition is left to the rules that already read it.

Only operand positions are read. A Put's or an Input's variable may be an
array, and an Open's path and access words are left alone.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from typing import Generic, TypeVar

from ..lexer.token_helpers import token_word, tokens_without_leading_line_number
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize
from ..parser.nodes import ExprNode, StatementNode
from ..parser.parse_expression import parse_expression

_K = TypeVar("_K")
_V = TypeVar("_V")


class _IdentityTable(Generic[_K, _V]):
    """Values by the identity of a key object, where upstream keeps a WeakMap.
    Each entry keeps its key alive, so an id() is never reused while it stands,
    and past `capacity` the oldest entry goes."""

    __slots__ = ("_capacity", "_entries")

    def __init__(self, capacity: int) -> None:
        self._capacity = capacity
        self._entries: dict[int, tuple[_K, _V]] = {}

    def get(self, key: _K) -> _V | None:
        entry = self._entries.get(id(key))
        return entry[1] if entry is not None and entry[0] is key else None

    def set(self, key: _K, value: _V) -> None:
        self._entries.pop(id(key), None)
        self._entries[id(key)] = (key, value)
        if len(self._entries) > self._capacity:
            del self._entries[next(iter(self._entries))]


# A statement's text, its roots, and the Case items among them with their operators.
_ROOTS: _IdentityTable[
    StatementNode, tuple[str, tuple[ExprNode, ...], tuple[tuple[ExprNode, str], ...]]
] = _IdentityTable(4096)

# The Case items, each compared with the Select value as by `=` or its `Is` operator.
_CASE_COMPARANDS: _IdentityTable[ExprNode, str] = _IdentityTable(4096)


def case_comparison_operator(expr: ExprNode) -> str | None:
    """The operator a Case item is compared by, when `expr` is one: `=`, or `Is >`'s `>`."""
    return _CASE_COMPARANDS.get(expr)


# [from, to) of an operand, and for a Case item the operator it is compared by.
_OperandRange = tuple[int, int, "str | None"]


def statement_operand_roots(node: StatementNode) -> tuple[ExprNode, ...]:
    """The operand expressions of a statement, each a root for the expression walk."""
    cached = _ROOTS.get(node)
    if cached is not None and cached[0] == node.raw:
        # Upstream's WeakMap keeps a Case item's operator while the item lives;
        # the bounded table here may have dropped it, so it is set again.
        for expr, comparison in cached[2]:
            _CASE_COMPARANDS.set(expr, comparison)
        return cached[1]
    toks = [tok for tok in tokenize(node.raw) if tok.kind not in (TokenKind.COMMENT, TokenKind.NEWLINE)]
    # The tokens are this call's own, so they are moved to source offsets in place.
    for tok in toks:
        tok.start += node.span.start
        tok.end += node.span.start
    roots: list[ExprNode] = []
    comparands: list[tuple[ExprNode, str]] = []

    def add(offset: int) -> Callable[[_OperandRange], None]:
        def add_range(operand: _OperandRange) -> None:
            start, stop, comparison = operand
            expr = parse_expression(toks, offset + start, offset + stop).expr
            if expr is not None:
                roots.append(expr)
                if comparison is not None:
                    _CASE_COMPARANDS.set(expr, comparison)
                    comparands.append((expr, comparison))

        return add_range

    if node.single_line_if_branches:
        for arm in node.single_line_if_branches:
            in_arm = [tok for tok in toks if tok.start >= arm.start and tok.end <= arm.end]
            if in_arm:
                add_in_arm = add(toks.index(in_arm[0]))
                for operand in _arm_operand_ranges(in_arm):
                    add_in_arm(operand)
    else:
        body = tokens_without_leading_line_number(toks)
        add_in_body = add(len(toks) - len(body))
        for operand in _operand_ranges(body):
            add_in_body(operand)
    result = tuple(roots)
    _ROOTS.set(node, (node.raw, result, tuple(comparands)))
    return result


def _top_level(
    toks: Sequence[VbaToken], from_: int, to: int, stop: Callable[[VbaToken], bool]
) -> list[int]:
    """Indices of the top-level tokens in [from_, to) that `stop` matches."""
    out: list[int] = []
    depth = 0
    for i in range(from_, to):
        raw = toks[i].raw_text
        depth += 1 if raw == "(" else -1 if raw == ")" else 0
        if depth == 0 and stop(toks[i]):
            out.append(i)
    return out


def _pieces(
    toks: Sequence[VbaToken], from_: int, to: int, separators: frozenset[str]
) -> list[tuple[int, int]]:
    """[from, to) of the pieces between top-level separators."""
    out: list[tuple[int, int]] = []
    start = from_
    for at in _top_level(toks, from_, to, lambda tok: tok.raw_text in separators):
        out.append((start, at))
        start = at + 1
    out.append((start, to))
    return out


_COMMA = frozenset({","})
_PRINT_SEPARATORS = frozenset({",", ";"})
_RELATIONAL = frozenset({"=", "<>", "<", ">", "<=", ">="})


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    return toks[i].raw_text if 0 <= i < len(toks) else None


def _word_at(toks: Sequence[VbaToken], i: int) -> str:
    return token_word(toks[i]) if 0 <= i < len(toks) else ""


def _file_number(toks: Sequence[VbaToken], from_: int, to: int) -> _OperandRange:
    """`#n` or `n` at `from_`: the number's range, without its `#`."""
    return (from_ + 1, to, None) if _raw_at(toks, from_) == "#" else (from_, to, None)


def _arm_operand_ranges(toks: Sequence[VbaToken]) -> list[_OperandRange]:
    """The operand ranges of a single-line If arm. An arm can be an If of its own,
    `Else If s Is Nothing Then ...`, whose condition no other rule reads, so it
    is read here with its own arms (XLIDE issue #1417, measured in all four
    hosts). Upstream recurses once per nested If; an explicit stack of (arm
    start, arm tokens) does the same here, its ranges kept in upstream's order."""
    out: list[_OperandRange] = []
    # Each entry is an arm still to read, (its start, its tokens), or a range
    # already placed, (None, the range).
    pending: list[tuple[int, Sequence[VbaToken]] | tuple[None, _OperandRange]] = [(0, toks)]
    while pending:
        item = pending.pop()
        if item[0] is None:
            out.append(item[1])
            continue
        base, arm = item
        if _word_at(arm, 0) != "if":
            out.extend((start + base, stop + base, comparison) for start, stop, comparison in _operand_ranges(arm))
            continue
        then_at = _top_level(arm, 1, len(arm), lambda tok: token_word(tok) == "then")
        if not then_at:
            continue
        then = then_at[0]
        later: list[tuple[int, Sequence[VbaToken]] | tuple[None, _OperandRange]] = [
            (None, (base + 1, base + then, None))
        ]
        elses = _top_level(arm, then + 1, len(arm), lambda tok: token_word(tok) == "else")
        from_ = then + 1
        for to in [*elses, len(arm)]:
            if from_ < to:
                later.append((base + from_, arm[from_:to]))
            from_ = to + 1
        pending.extend(reversed(later))
    return [operand for operand in out if operand[0] < operand[1]]


def _operand_ranges(toks: Sequence[VbaToken]) -> list[_OperandRange]:
    """The operand ranges of one statement's tokens, by its leading words."""
    head = _word_at(toks, 0)
    end = len(toks)
    out: list[_OperandRange] = []

    def items(from_: int, separators: frozenset[str]) -> None:
        for start, stop in _pieces(toks, from_, end, separators):
            # `Tab(n)` and `Spc(n)` take their argument; a bare Tab is a column.
            word = _word_at(toks, start)
            if word in ("tab", "spc") and _raw_at(toks, start + 1) == "(" and _raw_at(toks, stop - 1) == ")":
                out.append((start + 2, stop - 1, None))
            elif start < stop:
                out.append((start, stop, None))

    if head == "case":
        if _word_at(toks, 1) != "else":
            for start, stop in _pieces(toks, 1, end, _COMMA):
                if _word_at(toks, start) == "is" and (_raw_at(toks, start + 1) or "") in _RELATIONAL:
                    out.append((start + 2, stop, toks[start + 1].raw_text))
                    continue
                to_at = _top_level(toks, start, stop, lambda tok: token_word(tok) == "to")
                if not to_at:
                    out.append((start, stop, "="))
                else:
                    out.append((start, to_at[0], ">="))
                    out.append((to_at[0] + 1, stop, "<="))
    elif head in ("print", "write"):
        if _raw_at(toks, 1) == "#":
            first, *rest = _pieces(toks, 1, end, _COMMA)
            out.append(_file_number(toks, first[0], first[1]))
            if rest:
                items(rest[0][0], _PRINT_SEPARATORS if head == "print" else _COMMA)
    elif head == "debug":
        if _raw_at(toks, 1) == "." and _word_at(toks, 2) == "print":
            items(3, _PRINT_SEPARATORS)
    elif head in ("seek", "close"):
        for start, stop in _pieces(toks, 1, end, _COMMA):
            out.append(_file_number(toks, start, stop))
    elif head in ("put", "get"):
        # The number and the record; the variable may be an array.
        parts = _pieces(toks, 1, end, _COMMA)
        out.append(_file_number(toks, parts[0][0], parts[0][1]))
        if len(parts) > 1:
            out.append((parts[1][0], parts[1][1], None))
    elif head in ("line", "input"):
        at = (2 if _word_at(toks, 1) == "input" else -1) if head == "line" else 1
        if at > 0 and _raw_at(toks, at) == "#":
            number = _pieces(toks, at, end, _COMMA)[0]
            out.append(_file_number(toks, number[0], number[1]))
    elif head in ("lock", "unlock"):
        parts = _pieces(toks, 1, end, _COMMA)
        out.append(_file_number(toks, parts[0][0], parts[0][1]))
        if len(parts) > 1:
            lo, hi = parts[1]
            to_at = _top_level(toks, lo, hi, lambda tok: token_word(tok) == "to")
            if not to_at:
                out.append((lo, hi, None))
            else:
                out.append((lo, to_at[0], None))
                out.append((to_at[0] + 1, hi, None))
    elif head == "open":
        as_at = _top_level(toks, 1, end, lambda tok: token_word(tok) == "as")
        if as_at:
            as_index = as_at[-1]
            len_at = _top_level(toks, as_index + 1, end, lambda tok: token_word(tok) == "len")
            out.append(_file_number(toks, as_index + 1, len_at[0] if len_at else end))
    elif head == "redim":
        out.extend(_redim_bound_ranges(toks))
    return [operand for operand in out if operand[0] < operand[1]]


def _redim_bound_ranges(toks: Sequence[VbaToken]) -> list[_OperandRange]:
    """The bounds of `ReDim [Preserve] a(1 To n, m) [As T], b(k)`: each lower and
    upper bound of each variable, an expression the VBE judges as one, so
    `ReDim a(n Is Nothing)` is the Type mismatch it is anywhere (XLIDE issue
    #1412, measured in Excel, Word, PowerPoint and Access 16.0)."""
    out: list[_OperandRange] = []
    from_ = 2 if _word_at(toks, 1) == "preserve" else 1
    for start, stop in _pieces(toks, from_, len(toks), _COMMA):
        open_ = next((i for i in range(start, stop) if toks[i].raw_text == "("), -1)
        if open_ < 0:
            continue
        depth = 0
        close = -1
        for i in range(open_, stop):
            raw = toks[i].raw_text
            depth += 1 if raw == "(" else -1 if raw == ")" else 0
            if depth == 0:
                close = i
                break
        if close < 0:
            continue
        for dim_start, dim_stop in _pieces(toks, open_ + 1, close, _COMMA):
            to_at = _top_level(toks, dim_start, dim_stop, lambda tok: token_word(tok) == "to")
            if not to_at:
                out.append((dim_start, dim_stop, None))
            else:
                out.append((dim_start, to_at[0], None))
                out.append((to_at[0] + 1, dim_stop, None))
    return [operand for operand in out if operand[0] < operand[1]]
