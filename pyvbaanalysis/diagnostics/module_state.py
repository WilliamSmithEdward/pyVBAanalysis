"""Module-level and class-level state the runtime rules can rely on (XLIDE issue
#241). A module variable that no statement ever writes keeps its initial value
everywhere: an object is Nothing, a dynamic array has no elements, a number is 0,
a String is "" and a Variant is Empty. Measured in Excel 16.0 (build 20326):
`Private m As Collection` then `m.Count` raises 91, and `Private z As Long` then
`10 / z` raises 11.

What counts as a write is read from the tokens alone and errs toward writing: an
assignment's target, any name in a Set, ReDim, Erase, For, Input, Get, Line
Input, LSet, RSet or Mid statement, and a whole name passed to anything but a VBA
library function, which may take it ByRef. A Private variable needs only its own
module; a Public one needs every module of the project, which the project index
supplies as `project_written_names`.

Ported from xlide_vscode/src/analyzer/diagnostics/moduleState.ts.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass, field

from ..identity_cache import IdentityLru
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize_cached
from ..parser.nodes import ProcedureNode
from ..symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbol, VbaSymbolKind
from .walker import token_name, token_text

# Statement heads that write every name they mention.
_WRITING_HEADS: frozenset[str] = frozenset(
    {"set", "let", "redim", "erase", "for", "input", "get", "line", "lset", "rset", "mid", "mid$"}
)

# Statement heads that declare rather than run.
_DECLARING_HEADS: frozenset[str] = frozenset(
    {
        "dim", "private", "public", "global", "friend", "static", "const", "sub", "function", "property",
        "declare", "type", "enum", "end", "option", "attribute", "implements", "event",
    }
)

# Heads after which an `=` compares rather than assigns.
_COMPARING_HEADS: frozenset[str] = frozenset(
    {"if", "elseif", "while", "do", "loop", "until", "case", "select", "debug", "print", "return", "call"}
)


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) out of range."""
    return toks[i] if 0 <= i < len(toks) else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name is not None else None


def written_names_in(source: str) -> AbstractSet[str]:
    """The lowercased names a module's code may write, by the rules at the top of
    this file. Over-reporting a write only keeps a rule quiet."""
    # A dict keeps the order names were first written, as upstream's Set does.
    written: dict[str, None] = {}
    procedures = _procedure_names_in(source)
    statement: list[VbaToken] = []

    def flush() -> None:
        nonlocal statement
        for segment in _segments_of(statement):
            _mark_writes(segment, procedures, written)
        statement = []

    for tok in tokenize_cached(source):
        if tok.kind is TokenKind.NEWLINE or tok.kind is TokenKind.COLON:
            flush()
        elif tok.kind is not TokenKind.COMMENT:
            statement.append(tok)
    flush()
    return written.keys()


def _procedure_names_in(source: str) -> AbstractSet[str]:
    """The procedures a module declares, which shadow the VBA library's functions."""
    out: set[str] = set()
    toks = [
        tok
        for tok in tokenize_cached(source)
        if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE
    ]
    for i in range(len(toks) - 1):
        word = token_text(toks[i])
        previous = token_text(_at(toks, i - 1))
        if (word == "sub" or word == "function") and previous != "end" and previous != "exit":
            name = _lower_name(toks[i + 1])
            if name:
                out.add(name)
    return out


def _segments_of(statement: Sequence[VbaToken]) -> list[list[VbaToken]]:
    """A statement, and each statement a single-line If runs after Then or Else."""
    out: list[list[VbaToken]] = []
    current: list[VbaToken] = []
    for tok in statement:
        word = token_text(tok)
        if word == "then" or word == "else":
            out.append(current)
            current = []
            continue
        current.append(tok)
    out.append(current)
    return out


def _mark_writes(segment: Sequence[VbaToken], procedures: AbstractSet[str], written: dict[str, None]) -> None:
    from ..lexer.token_helpers import top_level_equals_index
    from ..runtime.vba_runtime import resolve_runtime_function

    # A line number or label leads some statements.
    first = _at(segment, 0)
    start = 1 if first is not None and first.kind is TokenKind.INTEGER_LITERAL else 0
    toks = segment[start:]
    head = token_text(_at(toks, 0))
    if not head or head in _DECLARING_HEADS:
        return

    def mark_all(from_: int, to: int) -> None:
        for i in range(from_, to):
            name = _lower_name(toks[i])
            if name:
                written[name] = None

    if head in _WRITING_HEADS:
        mark_all(0, len(toks))
        return
    equals = top_level_equals_index(toks)
    assignment = equals > 0 and head not in _COMPARING_HEADS
    if assignment:
        mark_all(0, equals)
    # A whole name passed to a call may come back changed (ByRef).
    callee_at = _enclosing_callee_lookup(toks)
    for i in range(equals + 1 if assignment else 1, len(toks)):
        name = _lower_name(toks[i])
        before = _at(toks, i - 1)
        if not name or (before is not None and (before.raw_text == "." or before.raw_text == "!")):
            continue
        following = _at(toks, i + 1)
        nxt = following.raw_text if following is not None else None
        if nxt is not None and nxt != "," and nxt != ")":
            continue
        previous = before
        callee = callee_at(i)
        if callee is None:
            # At the top level: an argument of a call statement, `Fill s`.
            two_back = _at(toks, i - 2)
            call_statement = (
                not assignment
                and head not in _COMPARING_HEADS
                and (
                    (previous is not None and previous.raw_text == ",")
                    or (i == 1 and token_name(previous) is not None)
                    or (
                        previous is not None
                        and previous.kind is TokenKind.IDENTIFIER
                        and two_back is not None
                        and two_back.raw_text == "."
                    )
                )
            )
            if call_statement:
                written[name] = None
            continue
        if previous is None or (previous.raw_text != "(" and previous.raw_text != ","):
            continue
        runtime = resolve_runtime_function(callee) if callee not in procedures else None
        library = runtime is not None and runtime.kind == "function"
        if not library:
            written[name] = None


def _enclosing_callee_lookup(toks: Sequence[VbaToken]) -> Callable[[int], str | None]:
    """Enclosing callees for the increasing token offsets visited by _mark_writes."""
    stack: list[str] = []
    cursor = 0

    def lookup(at: int) -> str | None:
        nonlocal cursor
        # Scan only the prefix not already visited by a previous argument.
        while cursor < at:
            raw = toks[cursor].raw_text
            if raw == "(":
                stack.append(_lower_name(_at(toks, cursor - 1)) or "")
            elif raw == ")":
                if stack:
                    stack.pop()
            cursor += 1
        return stack[-1] if stack else None

    return lookup


@dataclass(slots=True)
class _ModuleState:
    source: str
    project_writes: AbstractSet[str] | None
    variables: Mapping[str, VbaSymbol]
    # Each procedure (by identity, held alive here) to its result.
    procedures: dict[int, tuple[ProcedureNode, Mapping[str, VbaSymbol]]] = field(default_factory=dict)


# Keyed by ModuleSymbols identity. A cleared entry is stored as an empty tuple:
# IdentityLru has no delete, and get() answers None for a key it never saw.
_PROJECT_WRITES = IdentityLru(capacity=16)
_MODULE_STATES = IdentityLru(capacity=16)
_CLEARED = ()


def remember_project_written_names(symbols: ModuleSymbols, names: AbstractSet[str] | None) -> None:
    """Records, for one analysis of a module, what the rest of the project may
    write. Without it a Public variable is never taken as unchanged."""
    _PROJECT_WRITES.put(names if names is not None else _CLEARED, symbols)
    _MODULE_STATES.put(_CLEARED, symbols)


def untouched_module_variables(source: str, symbols: ModuleSymbols) -> Mapping[str, VbaSymbol]:
    """The module's variables that nothing writes, by lowercased name. A Private
    or Dim variable needs only this module; a Public or Global one needs the whole
    project. `As New` and fixed-length Strings are left out."""
    return _module_state_for(source, symbols).variables


def _module_state_for(source: str, symbols: ModuleSymbols) -> _ModuleState:
    """Read-only state shared by consumers of this bound module and source."""
    remembered = _PROJECT_WRITES.get(symbols)
    project: AbstractSet[str] | None = remembered if remembered is not None and remembered is not _CLEARED else None
    cached = _MODULE_STATES.get(symbols)
    if (
        isinstance(cached, _ModuleState)
        and (cached.source is source or cached.source == source)
        and cached.project_writes is project
    ):
        return cached
    writes = written_names_in(source)
    out: dict[str, VbaSymbol] = {}
    for child in symbols.root.children or []:
        if (
            child.kind is not VbaSymbolKind.MODULE_VARIABLE
            or child.is_auto_instantiated
            or child.fixed_length is not None
        ):
            continue
        lower = child.name.lower()
        shared = child.visibility is SymbolVisibility.PUBLIC or child.visibility is SymbolVisibility.GLOBAL
        if lower in writes or (shared and (project is None or lower in project)):
            continue
        out[lower] = child
    state = _ModuleState(source=source, project_writes=project, variables=out)
    _MODULE_STATES.put(state, symbols)
    return state


def untouched_module_variables_in(
    source: str, symbols: ModuleSymbols, proc: ProcedureNode
) -> Mapping[str, VbaSymbol]:
    """untouched_module_variables less those a procedure's own local or parameter hides."""
    from ..types.type_inference import procedure_symbol_for

    state = _module_state_for(source, symbols)
    all_ = state.variables
    cached = state.procedures.get(id(proc))
    if cached is not None and cached[0] is proc:
        return cached[1]
    if len(all_) == 0:
        return all_
    proc_symbol = procedure_symbol_for(symbols, proc)
    hidden = {child.name.lower() for child in (proc_symbol.children if proc_symbol is not None else None) or []}
    for param in proc.params:
        hidden.add(param.name.lower())
    hidden.add(proc.name.lower())
    out: dict[str, VbaSymbol] | None = None
    for lower in hidden:
        if lower in all_:
            if out is None:
                out = dict(all_)
            del out[lower]
    result = out if out is not None else all_
    state.procedures[id(proc)] = (proc, result)
    return result
