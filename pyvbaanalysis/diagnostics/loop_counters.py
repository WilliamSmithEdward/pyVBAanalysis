"""Loop counters and the values their passes take (XLIDE issue #200).

Ported from xlide_vscode/src/analyzer/diagnostics/loopCounters.ts.

The usual off-by-one bug is a counter one step past what it indexes:
`For i = 0 To Len(s) - 1` into Mid$, which starts at 1, or
`For i = LBound(a) To UBound(a) + 1` into a. The value rules check a statement
once as written; here each statement of a loop that runs on every pass learns
the counter's first and last values, so a rule can check the statement again
with the counter bound to each.

A counter is a For counter, or a Do/While counter the code steps itself:
`i = 1: Do While i <= 4 ... i = i + 1: Loop`. Its bounds are a whole number, or
Len, UBound, LBound or .Count of a name, plus or minus a whole number, so
`UBound(a) + 1` is known to pass a's last element whatever a holds.

A statement runs on every pass when it sits in the loop's own body, or in a With
there, before anything that may leave the pass: Exit, GoTo, Resume, Return, End,
Stop, a label (which a GoTo may reach), or a block holding any of those. A
statement inside an If or an inner loop may not run on the first or the last
pass, and is left alone. A loop whose body writes the counter has no counter
here, and a bound whose name the body writes (an assignment, ReDim, a ByRef
argument, Add or Remove) is not read: For reads its bounds once, before the
first pass.

Upstream's body walks recurse; here they run on explicit stacks.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Iterator, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass, field
from typing import Literal

from ..conditional import ConditionalActivityTracker
from ..constants.integer_constant_expression import parse_vba_integer_literal
from ..flow.procedure_labels import jump_target_label_declaration
from ..identity_cache import IdentityLru
from ..js_compat import js_number_to_string, utf16_length
from ..lexer.token_helpers import split_top_level_token_groups
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import (
    BodyNode,
    ConditionalDirectiveNode,
    DoBlockNode,
    ForBlockNode,
    LeafStatementNode,
    Span,
    StatementNode,
    VariableGroupNode,
    WhileBlockNode,
    WithBlockNode,
    is_leaf_statement,
    iter_body_nodes,
)
from .call_extraction import string_literal_value
from .context import PushFn
from .model import VbaDiagnosticData
from .walker import (
    bare_assignment_target,
    block_header_line_span,
    is_inactive_node,
    match_paren_from,
    set_assignment_target,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

CounterAtomKind = Literal["len", "ubound", "lbound", "count", "local"]


@dataclass(frozen=True, slots=True)
class CounterAtom:
    """A bound's named part: `Len(s)`, `UBound(a)`, `LBound(a, 2)`, `c.Count`."""

    kind: CounterAtomKind
    # The name it reads, lower-cased.
    name: str
    # As written, for messages.
    text: str
    # The dimension UBound or LBound asks for; 1 otherwise.
    dimension: int


@dataclass(frozen=True, slots=True)
class CounterValue:
    """A counter value: the atom plus the offset, or the offset alone."""

    offset: int
    atom: CounterAtom | None = None


@dataclass(frozen=True, slots=True, eq=False)
class LoopCounter:
    # As the header writes it.
    name: str
    first: CounterValue
    # None when the last pass cannot be told: a symbolic bound with a Step other
    # than 1 or -1.
    last: CounterValue | None
    step: int
    # "For", or "Do" for a counter the loop steps itself.
    loop: Literal["For", "Do"]
    # The loop, whose bounds read the values that hold as it starts.
    loop_node: BodyNode
    # A For loop's To bound, for working out the last pass where `last` is None.
    up_to: CounterValue | None = None


# The counter in force at a statement that runs on every pass, by lower-cased name.
CountersAt = Mapping[str, LoopCounter]

AtomValue = Callable[[CounterAtom, LoopCounter], "float | None"]


@dataclass(slots=True)
class StatementCounters:
    """Upstream's `ReadonlyMap<LeafStatementNode, CountersAt>`. Statement nodes are
    not hashable, so this keys them by identity and keeps them alive."""

    _by_id: dict[int, tuple[LeafStatementNode, CountersAt]] = field(default_factory=dict)

    def set(self, stmt: LeafStatementNode, counters: CountersAt) -> None:
        self._by_id[id(stmt)] = (stmt, counters)

    def get(self, stmt: LeafStatementNode) -> CountersAt | None:
        entry = self._by_id.get(id(stmt))
        return entry[1] if entry is not None else None

    def items(self) -> Iterator[tuple[LeafStatementNode, CountersAt]]:
        return iter(self._by_id.values())

    def __iter__(self) -> Iterator[tuple[LeafStatementNode, CountersAt]]:
        return self.items()

    def __len__(self) -> int:
        return len(self._by_id)


# Statement heads that write every name they mention.
_WRITING_HEADS = frozenset(
    ("set", "redim", "erase", "input", "get", "line", "lset", "rset", "mid", "mid$")
)

# Statement heads after which the rest of the pass may not run.
_LEAVING_HEADS = frozenset(("exit", "goto", "gosub", "resume", "return", "end", "stop"))

# Per body (and activity): the source the walk read, and its result.
_WALKS = IdentityLru(capacity=16)

_NO_COUNTERS = StatementCounters()


def loop_counters_at(
    source: str,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
) -> StatementCounters:
    """For each statement that runs on every pass of a loop with a counter, that
    counter."""
    cached: tuple[str, StatementCounters] | None = _WALKS.get(body, activity)
    if cached is not None and cached[0] == source:
        return cached[1]
    out = StatementCounters()
    zero_starts: list[Callable[[int, str], bool]] = []

    def zero_at_start_for(node: BodyNode) -> Callable[[str], bool]:
        def zero_at_start(lower: str) -> bool:
            if not zero_starts:
                zero_starts.append(_zero_start_lookup(source, body))
            return zero_starts[0](node.span.start, lower)

        return zero_at_start

    # Frames of (nodes, next index, in a loop): upstream recurses per block.
    stack: list[tuple[Sequence[BodyNode], list[int], bool]] = [(body, [0], False)]
    while stack:
        nodes, cursor, in_loop = stack[-1]
        k = cursor[0]
        if k >= len(nodes):
            stack.pop()
            continue
        cursor[0] = k + 1
        node = nodes[k]
        loop_body = getattr(node, "body", None)
        if is_inactive_node(activity, node) or not isinstance(loop_body, list):
            continue
        # A loop inside another may start on a later pass of the outer one.
        zero_at_start = None if in_loop else zero_at_start_for(node)
        found: tuple[_FoundCounter, Sequence[BodyNode]] | None
        if isinstance(node, ForBlockNode):
            found = _for_counter(
                source, node.span, node.each, node.control_variable, loop_body, activity
            )
        elif isinstance(node, (DoBlockNode, WhileBlockNode)):
            found = _stepped_counter(
                source,
                node.span,
                nodes[k - 1] if k > 0 else None,
                loop_body,
                activity,
                zero_at_start,
            )
        else:
            found = None
        if found is not None:
            spec, counted_body = found
            counter = LoopCounter(
                name=spec.name,
                first=spec.first,
                last=spec.last,
                step=spec.step,
                loop=spec.loop,
                loop_node=node,
                up_to=spec.up_to,
            )
            counters: CountersAt = {counter.name.lower(): counter}
            leaves: list[LeafStatementNode] = []
            _every_pass_leaves(source, counted_body, activity, leaves)
            for leaf in leaves:
                out.set(leaf, counters)
        stack.append(
            (
                loop_body,
                [0],
                in_loop or isinstance(node, (ForBlockNode, DoBlockNode, WhileBlockNode)),
            )
        )
    if len(out) == 0:
        # Most procedures have no counter, and the walk that says so costs less
        # than remembering it (XLIDE issue #200).
        return _NO_COUNTERS
    _WALKS.put((source, out), body, activity)
    return out


@dataclass(frozen=True, slots=True)
class _FoundCounter:
    name: str
    first: CounterValue
    last: CounterValue | None
    step: int
    loop: Literal["For", "Do"]
    up_to: CounterValue | None = None


def _for_counter(
    source: str,
    span: Span,
    each: bool,
    control_variable: str | None,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
) -> tuple[_FoundCounter, Sequence[BodyNode]] | None:
    """`For i = <bound> To <bound> [Step <whole number>]`."""
    if each or not control_variable:
        return None
    toks = statement_tokens_after_leading_label(source, block_header_line_span(source, span))
    eq = next((i for i, tok in enumerate(toks) if tok.raw_text == "="), -1)
    to = _top_level_word_index(toks, "to", eq + 1)
    if eq < 0 or to < 0:
        return None
    step_at = _top_level_word_index(toks, "step", to + 1)
    step = 1 if step_at < 0 else _whole_number(toks[step_at + 1 :])
    if step is None or step == 0:
        return None
    # Most bounds are no whole number and no atom, `.End(xlUp).Row`: read them
    # before walking the body.
    from_ = counter_value(toks[eq + 1 : to])
    bound = counter_value(toks[to + 1 : len(toks) if step_at < 0 else step_at]) if from_ else None
    if from_ is None or bound is None:
        return None
    written = _names_written_in(source, body, activity)
    # For reads its bounds once, so a local the body writes still bounds it with
    # what it held at the start, which a rule reads at the loop (XLIDE issue
    # #685). UBound and Len are left out: a symbolic check assumes the shape.
    first = _readable(from_, written, True)
    up_to = _readable(bound, written, True)
    if control_variable.lower() in written or first is None or up_to is None:
        return None
    last: CounterValue | None = up_to
    if abs(step) != 1:
        # The last pass is the last first + k*step not past the bound.
        last = (
            None
            if first.atom is not None or up_to.atom is not None
            else CounterValue(offset=first.offset + ((up_to.offset - first.offset) // step) * step)
        )
    return (
        _FoundCounter(name=control_variable, first=first, last=last, up_to=up_to, step=step, loop="For"),
        body,
    )


def _stepped_counter(
    source: str,
    span: Span,
    before: BodyNode | None,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    zero_at_start: Callable[[str], bool] | None = None,
) -> tuple[_FoundCounter, Sequence[BodyNode]] | None:
    """`i = <bound>` just before `Do While i <= <bound>` (or `<`, or `Do Until i >`
    or `>=`, or `While`), with `i = i + 1` the loop's last statement and no other
    write to i in it. Without that assignment, a numeric local nothing touches
    before the loop starts at 0 (XLIDE issue #350)."""
    header = statement_tokens_after_leading_label(source, block_header_line_span(source, span))
    first_word = token_text(_at(header, 0))
    i = 1 if first_word == "do" else 0 if first_word == "while" else -1
    test = token_text(_at(header, i))
    tested = token_name(_at(header, i + 1))
    if i < 0 or (test != "while" and test != "until") or not tested:
        return None
    lower = tested.lower()
    assigned = (
        bare_assignment_target(source, before.span)
        if before is not None and is_leaf_statement(before)
        else None
    )
    init_name: str
    init_first: CounterValue | None
    if assigned is not None and assigned[0].lower() == lower:
        init_name, init_first = assigned[0], counter_value(assigned[2])
    elif zero_at_start is not None and zero_at_start(lower):
        init_name, init_first = tested, CounterValue(offset=0)
    else:
        return None
    i += 2
    operator_tok = _at(header, i)
    operator = operator_tok.raw_text if operator_tok is not None else ""
    nxt = i + 1
    after_operator = _at(header, i + 1)
    if operator in ("<", ">") and after_operator is not None and after_operator.raw_text == "=":
        operator += "="
        nxt += 1
    # The last value the test lets in.
    shift: int | None
    if test == "while":
        shift = 0 if operator == "<=" else -1 if operator == "<" else None
    else:
        shift = 0 if operator == ">" else -1 if operator == ">=" else None
    if shift is None:
        return None
    passes = [
        node
        for node in body
        if not is_inactive_node(activity, node) and not isinstance(node, VariableGroupNode)
    ]
    increment = passes[-1] if passes else None
    if (
        increment is None
        or not is_leaf_statement(increment)
        or not _is_increment_of(source, increment.span, lower)
    ):
        return None
    rest = passes[:-1]
    written = _names_written_in(source, rest, activity)
    first = init_first
    limit = _readable(counter_value(header[nxt:]), written)
    if lower in written or first is None or first.atom is not None or limit is None:
        return None
    return (
        _FoundCounter(
            name=init_name,
            first=first,
            last=CounterValue(offset=limit.offset + shift, atom=limit.atom),
            step=1,
            loop="Do",
        ),
        rest,
    )


_ZERO_START_TYPES = frozenset(
    ("byte", "integer", "long", "longlong", "currency", "single", "double", "variant")
)
_ZERO_START_SUFFIXES: Mapping[str, str] = {
    "%": "integer",
    "&": "long",
    "^": "longlong",
    "@": "currency",
    "!": "single",
    "#": "double",
}

# JavaScript's /[A-Za-z0-9_$\u00C0-\uFFFF]/ reads an astral character as two
# code units in that range; here it is one code point past U+FFFF.
_IDENTIFIER_CHARACTERS = "A-Za-z0-9_$\u00c0-\U0010ffff"
_IDENTIFIER_RUN_RE = re.compile(f"[{_IDENTIFIER_CHARACTERS}]+")
_IDENTIFIER_WORD_RE = re.compile(f"^[{_IDENTIFIER_CHARACTERS}]+$")
_IDENTIFIER_CHARACTER_RE = re.compile(f"[{_IDENTIFIER_CHARACTERS}]")


def _never_zero(_loop_start: int, _lower: str) -> bool:
    return False


def _zero_start_lookup(source: str, body: Sequence[BodyNode]) -> Callable[[int, str], bool]:
    """Whether `lower` is a numeric local the procedure declares with Dim and no
    text before `loop_start` mentions but its declaration, in a body that holds
    no GoTo, GoSub or Resume to run the loop again: it is 0 as the loop starts
    (XLIDE issue #350)."""
    if len(body) == 0:
        return _never_zero
    start = body[0].span.start
    whole = source[start : body[-1].span.end]
    lowered = whole.lower()
    if _mentions(lowered, "goto") or _mentions(lowered, "resume") or _mentions(lowered, "gosub"):
        return _never_zero
    declared: dict[str, bool] = {}
    blanks: list[tuple[int, int]] = []
    for node in iter_body_nodes(body):
        if not isinstance(node, VariableGroupNode):
            continue
        # A group's first matching declaration wins; later groups replace it.
        seen: set[str] = set()
        for decl in node.declarations:
            lower = decl.name.lower()
            if lower in seen:
                continue
            seen.add(lower)
            if decl.as_type is not None:
                type_: str | None = decl.as_type.lower()
            elif decl.type_suffix:
                type_ = _ZERO_START_SUFFIXES.get(decl.type_suffix)
            else:
                type_ = "variant"
            declared[lower] = (
                not node.is_const
                and node.modifier.lower() == "dim"
                and not decl.is_array
                and type_ is not None
                and type_ in _ZERO_START_TYPES
            )
        blanks.append((node.span.start - start, node.span.end - start))
    parts: list[str] = []
    cursor = 0
    for from_, to in sorted(blanks, key=lambda blank: blank[0]):
        if to <= cursor:
            continue
        begin = max(cursor, from_)
        parts.append(whole[cursor:begin])
        parts.append(" " * (to - begin))
        cursor = to
    parts.append(whole[cursor:])
    masked = "".join(parts)
    folded = masked.lower()
    first_mentions: dict[str, int] = {}
    for match in _IDENTIFIER_RUN_RE.finditer(folded):
        word = match.group(0)
        if declared.get(word) and word not in first_mentions:
            first_mentions[word] = match.start()

    def lookup(loop_start: int, lower: str) -> bool:
        if not declared.get(lower):
            return False
        offset = loop_start - start
        # Expanded case folds and names outside the legacy word-character set
        # retain its exact prefix/boundary behavior instead of using text offsets.
        if len(folded) != len(masked) or not _IDENTIFIER_WORD_RE.match(lower):
            return not _mentions(masked[: max(0, offset)].lower(), lower)
        return first_mentions.get(lower, math.inf) >= offset

    return lookup


def _is_increment_of(source: str, span: Span, lower: str) -> bool:
    """`i = i + 1`."""
    target = bare_assignment_target(source, span)
    if target is None:
        return False
    value = [tok for tok in target[2] if tok.kind is not TokenKind.COMMENT]
    first_name = token_name(value[0]) if value else None
    return (
        target[0].lower() == lower
        and len(value) == 3
        and first_name is not None
        and first_name.lower() == lower
        and value[1].raw_text == "+"
        and value[2].kind is TokenKind.INTEGER_LITERAL
        and parse_vba_integer_literal(value[2].raw_text) == 1
    )


def _top_level_word_index(toks: Sequence[VbaToken], word: str, from_: int) -> int:
    depth = 0
    for i in range(max(0, from_), len(toks)):
        raw = toks[i].raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
        elif depth == 0 and token_text(toks[i]) == word:
            return i
    return -1


def _whole_number(toks: Sequence[VbaToken]) -> int | None:
    if len(toks) == 1 and toks[0].kind is TokenKind.INTEGER_LITERAL:
        return parse_vba_integer_literal(toks[0].raw_text)
    if (
        len(toks) == 2
        and toks[0].raw_text in ("-", "+")
        and toks[1].kind is TokenKind.INTEGER_LITERAL
    ):
        value = parse_vba_integer_literal(toks[1].raw_text)
        if value is None:
            return None
        return -value if toks[0].raw_text == "-" else value
    return None


def counter_value(tokens: Sequence[VbaToken]) -> CounterValue | None:
    """A bound: a whole number, or an atom plus or minus one."""
    toks = [tok for tok in tokens if tok.kind is not TokenKind.COMMENT]
    whole = _whole_number(toks)
    if whole is not None:
        return CounterValue(offset=whole)
    core: Sequence[VbaToken] = toks
    offset = 0
    sign_tok = _at(toks, len(toks) - 2)
    sign = sign_tok.raw_text if sign_tok is not None else None
    if len(toks) >= 3 and sign in ("+", "-") and toks[-1].kind is TokenKind.INTEGER_LITERAL:
        value = parse_vba_integer_literal(toks[-1].raw_text)
        if value is None:
            return None
        offset = -value if sign == "-" else value
        core = toks[:-2]
    # `s`, a local whose value a rule may know as the loop starts (XLIDE #346).
    if len(core) == 1 and core[0].kind is TokenKind.IDENTIFIER:
        return CounterValue(
            offset=offset,
            atom=CounterAtom(
                kind="local", name=core[0].raw_text.lower(), text=core[0].raw_text, dimension=1
            ),
        )
    # `c.Count`
    receiver = token_name(_at(core, 0))
    if (
        len(core) == 3
        and receiver
        and core[1].raw_text == "."
        and token_text(core[2]) == "count"
    ):
        return CounterValue(
            offset=offset,
            atom=CounterAtom(
                kind="count", name=receiver.lower(), text=f"{core[0].raw_text}.Count", dimension=1
            ),
        )
    callee = token_text(_at(core, 0))
    open_tok = _at(core, 1)
    if (
        callee not in ("len", "ubound", "lbound")
        or open_tok is None
        or open_tok.raw_text != "("
        or match_paren_from(core, 1) != len(core) - 1
    ):
        return None
    args = split_top_level_token_groups(core, 2, ",", len(core) - 1)
    if (
        callee == "len"
        and len(args) == 1
        and len(args[0]) == 1
        and args[0][0].kind is TokenKind.STRING_LITERAL
    ):
        return CounterValue(offset=utf16_length(string_literal_value(args[0][0].raw_text)) + offset)
    name = token_name(args[0][0]) if args and len(args[0]) == 1 else None
    dimension = _whole_number(args[1]) if len(args) == 2 else 1
    if not name or dimension is None or len(args) > (1 if callee == "len" else 2):
        return None
    text = f"{core[0].raw_text}({name}{f', {dimension}' if len(args) == 2 else ''})"
    kind: CounterAtomKind = "len" if callee == "len" else "ubound" if callee == "ubound" else "lbound"
    return CounterValue(
        offset=offset, atom=CounterAtom(kind=kind, name=name.lower(), text=text, dimension=dimension)
    )


def _readable(
    value: CounterValue | None, written: AbstractSet[str], read_once: bool = False
) -> CounterValue | None:
    """The value, unless the loop body writes the name its atom reads; a local of a
    For bound may be written."""
    if (
        value is not None
        and value.atom is not None
        and value.atom.name in written
        and not (read_once and value.atom.kind == "local")
    ):
        return None
    return value


def _names_written_in(
    source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> set[str]:
    """The names a body may write: assignment and Set targets, every name a writing
    statement mentions, bare arguments of a call (ByRef), receivers of Add, Remove
    and Clear, and the counters of loops inside it."""
    out: set[str] = set()
    skip = None if activity is None else (lambda node: activity.is_inactive(node.span))
    for node in iter_body_nodes(body, skip):
        if isinstance(node, ForBlockNode) and node.control_variable:
            out.add(node.control_variable.lower())
        if not is_leaf_statement(node):
            continue
        for span in statement_and_branch_spans(node):
            toks = statement_tokens_after_leading_label(source, span)
            target = bare_assignment_target(source, span) or set_assignment_target(source, span)
            if target is not None:
                out.add(target[0].lower())
            # `Mid(s, i, 1) = "x"` writes s alone: its start and length are read
            # (XLIDE issue #327).
            head = token_text(_at(toks, 0))
            second = _at(toks, 1)
            third = _at(toks, 2)
            if head in ("mid", "mid$", "midb", "midb$") and second is not None and (
                second.raw_text == "(" or (second.raw_text == "$" and third is not None and third.raw_text == "(")
            ):
                written = token_name(_at(toks, 2 if second.raw_text == "(" else 3))
                if written:
                    out.add(written.lower())
            elif head in _WRITING_HEADS:
                for tok in toks:
                    mentioned = token_name(tok)
                    if mentioned:
                        out.add(mentioned.lower())
            out.update(_call_statement_arguments(toks))
            for i in range(len(toks) - 2):
                member = token_text(toks[i + 2])
                if toks[i + 1].raw_text == "." and member in ("add", "remove", "clear"):
                    receiver = token_name(toks[i])
                    if receiver:
                        out.add(receiver.lower())
    return out


def _call_statement_arguments(toks: Sequence[VbaToken]) -> list[str]:
    """The names a call statement passes whole, which it may change ByRef: `Bump i`,
    `Call Bump(i)`, `obj.Move i`. A name inside an expression's parentheses,
    `a(i)` or `Mid$(s, i, 1)`, is read: a loop counter changed through a
    function's ByRef argument is rare enough to leave."""
    call = token_text(_at(toks, 0)) == "call"
    head = 1 if call else 0
    head_tok = _at(toks, head)
    if not token_name(head_tok) and (head_tok is None or head_tok.raw_text != "."):
        return []
    out: list[str] = []
    depth = 0
    for i in range(head + 1, len(toks)):
        raw = toks[i].raw_text
        if raw == "(":
            depth += 1
            continue
        if raw == ")":
            depth -= 1
            continue
        if depth == 0 and toks[i].kind is TokenKind.OPERATOR and raw == "=":
            return []  # an assignment, not a call
        name = token_name(toks[i])
        nxt = _at(toks, i + 1)
        whole = toks[i - 1].raw_text != "." and (nxt is None or nxt.raw_text not in ("(", "."))
        if name and whole and depth == (1 if call else 0):
            out.append(name.lower())
    return out


def _every_pass_leaves(
    source: str,
    body: Sequence[BodyNode],
    activity: ConditionalActivityTracker | None,
    out: list[LeafStatementNode],
) -> bool:
    """Collects the leaves of a loop body that run on every pass; False once
    something may end the pass early, so nothing after it does."""
    # Upstream recurses into each With; a stack of the bodies being walked.
    stack: list[Iterator[BodyNode]] = [iter(body)]
    while stack:
        for node in stack[-1]:
            if is_inactive_node(activity, node) or isinstance(
                node, (VariableGroupNode, ConditionalDirectiveNode)
            ):
                continue
            if is_leaf_statement(node):
                if jump_target_label_declaration(source, node.span) is not None or _may_leave(
                    source, node
                ):
                    return False
                if not isinstance(node, StatementNode) or node.single_line_if_branches is None:
                    out.append(node)
                continue
            if isinstance(node, WithBlockNode):
                stack.append(iter(node.body))
                break
            child = getattr(node, "body", None)
            if isinstance(child, list) and _block_may_leave(source, child):
                return False
        else:
            stack.pop()
    return True


def _may_leave(source: str, node: LeafStatementNode) -> bool:
    for span in statement_and_branch_spans(node):
        toks = statement_tokens_after_leading_label(source, span)
        head = token_text(_at(toks, 0))
        if head in _LEAVING_HEADS or (head == "err" and token_text(_at(toks, 2)) == "raise"):
            return True
    return False


def _block_may_leave(source: str, body: Sequence[BodyNode]) -> bool:
    return any(
        is_leaf_statement(node)
        and (
            jump_target_label_declaration(source, node.span) is not None
            or _may_leave(source, node)
        )
        for node in iter_body_nodes(body)
    )


def counter_number(
    value: CounterValue | None,
    counter: LoopCounter,
    atom_value: AtomValue,
) -> float | None:
    """The number a counter value is, given what each atom is; None when an atom
    is unknown."""
    if value is None:
        return None
    if value.atom is None:
        return value.offset
    base = atom_value(value.atom, counter)
    return None if base is None else base + value.offset


def counter_text(value: CounterValue) -> str:
    """`UBound(a) + 1`, `0`, `c.Count - 1`: a counter value as the code would write
    it."""
    if value.atom is None:
        return str(value.offset)
    if value.offset == 0:
        return value.atom.text
    return f"{value.atom.text} {'+' if value.offset > 0 else '-'} {abs(value.offset)}"


@dataclass(frozen=True, slots=True)
class CounterPass:
    pass_: Literal["first", "last"]
    counter: LoopCounter
    value: float


def numeric_counter_passes(counter: LoopCounter, atom_value: AtomValue) -> list[CounterPass]:
    """The passes whose value is a number, given the atoms a rule knows. None when
    the numbers show the loop never runs: `For i = 0 To Len("") - 1`."""
    first = counter_number(counter.first, counter, atom_value)
    last = counter_number(counter.last, counter, atom_value)
    # A symbolic bound with a Step: the last pass is the last first + k*step not
    # past it, once the numbers are known (XLIDE issue #685).
    if last is None and counter.up_to is not None and first is not None:
        bound = counter_number(counter.up_to, counter, atom_value)
        last = (
            None
            if bound is None
            else first + math.floor((bound - first) / counter.step) * counter.step
        )
    if (
        first is not None
        and last is not None
        and (first > last if counter.step > 0 else first < last)
    ):
        return []
    out: list[CounterPass] = []
    if first is not None:
        out.append(CounterPass(pass_="first", counter=counter, value=first))
    if last is not None and last != first:
        out.append(CounterPass(pass_="last", counter=counter, value=last))
    return out


def _number_text(value: float) -> str:
    """JavaScript's String(number)."""
    if isinstance(value, int):
        return str(value)
    return js_number_to_string(value)


_NO_VALUES: Mapping[str, float] = {}
_NO_PASSES: list[CounterPass] = []


def check_each_counter_pass(
    source: str,
    span: Span,
    counters: CountersAt | None,
    atom_value: AtomValue,
    check: Callable[[Mapping[str, float], PushFn], None],
    push: PushFn,
) -> None:
    """Runs a statement check as written, then once with the counter bound to each
    pass's value, and reports what only a pass makes fail, with the pass named.
    `check` reads the bound value for the counter's name in `values`."""
    # Most statements sit in no loop, or never read its counter: check them once,
    # as written, with nothing allocated.
    passes = (
        _named_counter_passes(source, span, counters, atom_value)
        if counters is not None
        else _NO_PASSES
    )
    if len(passes) == 0:
        check(_NO_VALUES, push)
        return
    seen: set[tuple[int, int]] = set()

    def as_written(
        rule: str, message: str, span: Span, data: VbaDiagnosticData | None = None
    ) -> None:
        seen.add((span.start, span.end))
        push(rule, message, span, data)

    check(_NO_VALUES, as_written)
    for counter_pass in passes:
        check(
            {counter_pass.counter.name.lower(): counter_pass.value},
            _pass_push(counter_pass, seen, push),
        )


def _pass_push(counter_pass: CounterPass, seen: set[tuple[int, int]], push: PushFn) -> PushFn:
    """What one pass reports: a finding the statement as written did not give, with
    the pass named."""
    counter = counter_pass.counter

    def on_pass(
        rule: str, message: str, span: Span, data: VbaDiagnosticData | None = None
    ) -> None:
        key = (span.start, span.end)
        if key in seen:
            return
        seen.add(key)
        push(
            rule,
            f"On the {counter_pass.pass_} pass of the {counter.loop} loop, where "
            f"'{counter.name}' is {_number_text(counter_pass.value)}: {message}",
            span,
            data,
        )

    return on_pass


def _named_counter_passes(
    source: str,
    span: Span,
    counters: CountersAt,
    atom_value: AtomValue,
) -> list[CounterPass]:
    """The numeric passes of the counters a statement names."""
    text = source[span.start : span.end].lower()
    out: list[CounterPass] = []
    for counter in counters.values():
        if _mentions(text, counter.name.lower()):
            out.extend(numeric_counter_passes(counter, atom_value))
    return out


def _mentions(text: str, name: str) -> bool:
    """Whether lower-cased text names a lower-cased identifier as a whole word."""
    at = text.find(name)
    while at >= 0:
        before = text[at - 1] if at > 0 else " "
        end = at + len(name)
        after = text[end] if end < len(text) else " "
        if not _IDENTIFIER_CHARACTER_RE.match(before) and not _IDENTIFIER_CHARACTER_RE.match(after):
            return True
        at = text.find(name, at + 1)
    return False


def _at(tokens: Sequence[VbaToken], i: int) -> VbaToken | None:
    return tokens[i] if 0 <= i < len(tokens) else None
