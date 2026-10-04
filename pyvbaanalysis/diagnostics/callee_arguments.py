"""What a procedure of the module does with an argument passed to it whole (XLIDE
issue #449, each measured in Excel 16.0). A variable passed ByRef may come back
changed, so the state rules forget it at the call. Two kinds of callee cannot
change it: one whose parameter is ByVal, and one that never writes the parameter -
no assignment, Set, ReDim, Erase, Input, Get, LSet, RSet or Mid statement on it,
no For over it, and no call it is passed on to. `InitV c` with `ByVal c As
Collection`, and `Touch c` that only reads c, both leave c Nothing.

Ported from xlide_vscode/src/analyzer/diagnostics/calleeArguments.ts.
"""

from __future__ import annotations

import dataclasses
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field
from typing import Protocol

from ..identity_cache import IdentityLru
from ..lexer.token_helpers import split_top_level_token_groups
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import (
    ForBlockNode,
    ModuleNode,
    ProcedureNode,
    ProcKind,
    StatementNode,
    VariableGroupNode,
    is_leaf_statement,
    iter_body_nodes,
)
from .walker import (
    match_paren_from,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

class CalleeKeepsArgument(Protocol):
    """Whether the module's own procedure `callee` leaves an argument in slot
    `index`, or named `named`, as it was."""

    def __call__(self, callee: str, index: int, named: str | None = None) -> bool: ...

_WRITING_HEADS: frozenset[str] = frozenset(
    {"set", "let", "redim", "erase", "input", "line", "get", "lset", "rset", "mid", "mid$", "midb", "midb$", "for"}
)


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) out of range."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return None if tok is None else tok.raw_text


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name is not None else None


@dataclass(slots=True)
class _ModuleProcedures:
    """Upstream's per-module and per-procedure WeakMaps, held by one module entry:
    the module keeps its procedures alive, so their ids stay theirs."""

    # Each Sub or Function by lowercased name; None where the name is not one of a kind.
    by_name: dict[str, ProcedureNode | None]
    # Per procedure id: whether it keeps each parameter, by lowercased name.
    answers: dict[int, dict[str, bool]] = field(default_factory=dict)
    # Per procedure id: the replayed member calls on each parameter, None where unknown.
    member_calls: dict[int, dict[str, list[list[VbaToken]] | None]] = field(default_factory=dict)


_KEEPS_CACHE = IdentityLru()


def _module_procedures(source: str) -> _ModuleProcedures:
    from ..parser.parse_module import parse_module

    module: ModuleNode = parse_module(source)
    cached = _KEEPS_CACHE.get(module)
    if isinstance(cached, _ModuleProcedures):
        return cached
    by_name: dict[str, ProcedureNode | None] = {}
    for member in module.members:
        if not isinstance(member, ProcedureNode):
            continue
        key = member.name.lower()
        by_name[key] = (
            None
            if key in by_name or (member.proc_kind is not ProcKind.SUB and member.proc_kind is not ProcKind.FUNCTION)
            else member
        )
    entry = _ModuleProcedures(by_name=by_name)
    _KEEPS_CACHE.put(entry, module)
    return entry


def _procedure_named(source: str, lower: str) -> ProcedureNode | None:
    """The module's procedure of that name, when there is exactly one and it is a Sub or Function."""
    return _module_procedures(source).by_name.get(lower)


def _writes_parameter(source: str, proc: ProcedureNode, lower: str) -> bool:
    """Whether a procedure's body writes the parameter, or passes it on whole to a call.

    Upstream visits each block's body and, for an If block, each arm's body as
    well; an If block's flat body already holds every arm's statements, so one
    walk of every nested body reaches the same statements."""
    from ..runtime.vba_runtime import resolve_runtime_function

    def builtin(name: str) -> bool:
        return _procedure_named(source, name) is None and resolve_runtime_function(name) is not None

    for node in iter_body_nodes(proc.body):
        if isinstance(node, ForBlockNode) and node.control_variable is not None and node.control_variable.lower() == lower:
            return True
        if is_leaf_statement(node):
            for span in statement_and_branch_spans(node):
                if _statement_writes(statement_tokens_after_leading_label(source, span), lower, builtin):
                    return True
    return False


def _never_builtin(_lower: str) -> bool:
    return False


def _statement_writes(
    toks: Sequence[VbaToken], lower: str, builtin: Callable[[str], bool] = _never_builtin
) -> bool:
    """Whether one statement may write `lower`: as its target, under a writing
    statement, or passed on whole. A VBA function only reads what it is given:
    `Debug.Print TypeName(p)` leaves p alone (issue #685, measured in Excel 16.0)."""
    if not any(_lower_name(tok) == lower for tok in toks):
        return False
    head = token_text(_at(toks, 0))
    if head in _WRITING_HEADS:
        return True
    # `p = 5` and `p(1) = 5`: the parameter or an element of it.
    if _lower_name(_at(toks, 0)) == lower and (_raw(toks, 1) == "=" or _raw(toks, 1) == "("):
        return True
    # Passed on whole, `Other p`, `Other x, p` or `x = F(p)`: the next callee may
    # write it. A call statement opens with the callee's name.
    first = _at(toks, 0)
    call_statement = (
        first is not None
        and first.kind is TokenKind.IDENTIFIER
        and not any(tok.raw_text == "=" for tok in toks)
    )
    depth = 0
    for i in range(1, len(toks)):
        raw = toks[i].raw_text
        depth += 1 if raw == "(" else -1 if raw == ")" else 0
        if _lower_name(toks[i]) != lower:
            continue
        prev = _raw(toks, i - 1)
        nxt = _raw(toks, i + 1)
        if (
            prev == "."
            or prev == "!"
            or nxt == "."
            or nxt == "!"
            or nxt == "("
            or token_text(_at(toks, i + 1)) == "is"
        ):
            continue
        # After a comma inside a call's parentheses too: `F = G(a, p)` passes p
        # on (issue #665).
        if (
            prev == "("
            or prev == ":="
            or (prev == "," and depth > 0)
            or (call_statement and (prev == "," or i == 1))
        ):
            # The call whose parentheses hold it, when that is a VBA function.
            open_ = i - 1
            level = 0
            while open_ >= 0:
                text = toks[open_].raw_text
                if text == ")":
                    level += 1
                elif text == "(":
                    if level == 0:
                        break
                    level -= 1
                open_ -= 1
            callee = _lower_name(_at(toks, open_ - 1)) if open_ > 0 and _raw(toks, open_ - 2) != "." else None
            if depth > 0 and callee is not None and builtin(callee):
                continue
            return True
    return False


def callee_keeps_argument(source: str) -> CalleeKeepsArgument:
    """A keeps-argument test over the module's own procedures (issue #449)."""

    def keeps_argument(callee: str, index: int, named: str | None = None) -> bool:
        proc = _procedure_named(source, callee.lower())
        if proc is None:
            return False
        if named is not None:
            named_lower = named.lower()
            param = next((p for p in proc.params if p.name.lower() == named_lower), None)
        else:
            param = proc.params[index] if 0 <= index < len(proc.params) else None
        if param is None or param.param_array:
            return False
        if param.by_val:
            return True
        cache = _module_procedures(source)
        answers = cache.answers.setdefault(id(proc), {})
        key = param.name.lower()
        keeps = answers.get(key)
        if keeps is None:
            keeps = not _writes_parameter(source, proc, key)
            answers[key] = keeps
        return keeps

    return keeps_argument


# The member calls a call statement makes on each object passed to it whole, as
# statements on the caller's own name: `R1 c` with `R1(ByVal p)` doing `p.Remove
# 1` is `c.Remove 1` (issue #685, measured in Excel 16.0). A name is absent when
# the callee may do anything else with it.
CalleeMemberCalls = Callable[[Sequence[VbaToken]], Mapping[str, Sequence[Sequence[VbaToken]]]]

# Members whose calls a callee's replay keeps: the ones that add and remove.
_REPLAYED_MEMBERS: frozenset[str] = frozenset({"add", "remove", "removeall"})

# Statement heads after which a later statement may not run.
_LEAVING_HEADS: frozenset[str] = frozenset(
    {"exit", "goto", "gosub", "on", "resume", "end", "stop", "return", "error"}
)

# Statement heads that declare and run nothing.
_DECLARING_HEADS: frozenset[str] = frozenset({"dim", "const", "static"})

_ASCII_WORD = "[A-Za-z0-9_]"


def _whole_word_pattern(lower: str) -> re.Pattern[str]:
    """JavaScript's `\\b<lower>\\b` with the `i` flag: `\\b` sits between an ASCII
    word character and anything else, so a name that starts or ends with another
    letter needs an ASCII word character beside it, as it does there."""

    def boundary(ch: str, before: bool) -> str:
        word = re.fullmatch(_ASCII_WORD, ch) is not None
        if before:
            return f"(?<!{_ASCII_WORD})" if word else f"(?<={_ASCII_WORD})"
        return f"(?!{_ASCII_WORD})" if word else f"(?={_ASCII_WORD})"

    if not lower:
        return re.compile(r"(?:)")
    return re.compile(boundary(lower[0], True) + re.escape(lower) + boundary(lower[-1], False), re.IGNORECASE)


def _member_calls_on(source: str, proc: ProcedureNode, lower: str) -> list[list[VbaToken]] | None:
    """Every statement of a callee that names the parameter, when each is an Add or
    Remove on it with literal arguments, and the callee is one straight line with
    nothing that leaves early. None otherwise."""
    cache = _module_procedures(source).member_calls.setdefault(id(proc), {})
    if lower in cache:
        return cache[lower]
    calls: list[list[VbaToken]] = []
    known = True
    pattern: re.Pattern[str] | None = None
    for node in proc.body:
        if isinstance(node, VariableGroupNode):
            if pattern is None:
                pattern = _whole_word_pattern(lower)
            if pattern.search(source[node.span.start : node.span.end]) is None:
                continue  # a Dim or Const that runs nothing
        if not is_leaf_statement(node) or (
            isinstance(node, StatementNode) and node.single_line_if_branches is not None
        ):
            known = False
            break
        toks = statement_tokens_after_leading_label(source, node.span)
        if token_text(_at(toks, 0)) in _LEAVING_HEADS:
            known = False
            break
        # Anything else may reach the object another way, a module variable
        # holding it: only declarations are let pass.
        if not any(_lower_name(tok) == lower for tok in toks):
            if token_text(_at(toks, 0)) in _DECLARING_HEADS:
                continue
            known = False
            break
        at = 1 if token_text(_at(toks, 0)) == "call" else 0
        rest = toks[at + 3 :]
        literals_only = all(
            (tok.kind is not TokenKind.IDENTIFIER and tok.kind is not TokenKind.KEYWORD)
            or _raw(rest, k + 1) == ":="
            for k, tok in enumerate(rest)
        )
        if (
            _lower_name(_at(toks, at)) != lower
            or _raw(toks, at + 1) != "."
            or token_text(_at(toks, at + 2)) not in _REPLAYED_MEMBERS
            or not literals_only
        ):
            known = False
            break
        calls.append(list(toks[at:]))
    result = calls if known else None
    cache[lower] = result
    return result


def callee_member_calls(source: str) -> CalleeMemberCalls:
    """The member calls of the module's own callees (issue #685)."""

    def member_calls(toks: Sequence[VbaToken]) -> Mapping[str, Sequence[Sequence[VbaToken]]]:
        out: dict[str, Sequence[Sequence[VbaToken]]] = {}
        at = 1 if token_text(_at(toks, 0)) == "call" else 0
        name = token_name(_at(toks, at))
        proc = _procedure_named(source, name.lower()) if name and _raw(toks, at + 1) != "." else None
        if proc is None:
            return out
        args: list[list[VbaToken]]
        if at == 1:
            if _raw(toks, 2) != "(" or match_paren_from(toks, 2) != len(toks) - 1:
                return out
            args = split_top_level_token_groups(toks, 3, ",", len(toks) - 1)
        else:
            # `R1 (c)` passes c's value, not c.
            if _raw(toks, 1) == "(" or _raw(toks, 1) == "=":
                return out
            args = split_top_level_token_groups(toks, 1, ",")
        passed = [_lower_name(arg[0]) if len(arg) == 1 else None for arg in args]
        for k, arg in enumerate(args):
            param = proc.params[k] if k < len(proc.params) else None
            # One object passed twice may change through either parameter.
            if (
                len(arg) != 1
                or arg[0].kind is not TokenKind.IDENTIFIER
                or param is None
                or param.param_array
                or passed.count(passed[k]) != 1
            ):
                continue
            lower = param.name.lower()
            calls = _member_calls_on(source, proc, lower)
            if calls is not None:
                caller = arg[0].raw_text
                out[caller.lower()] = [
                    [
                        dataclasses.replace(tok, raw_text=caller) if _lower_name(tok) == lower else tok
                        for tok in stmt
                    ]
                    for stmt in calls
                ]
        return out

    return member_calls
